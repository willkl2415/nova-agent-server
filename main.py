"""
NOVA Agent Server v3.0 - AMPLIFIED OUTPUTS
FastAPI server for executing autonomous training agents with Claude AI
Generates professional .docx and .xlsx outputs meeting Output Amplification Specification

Endpoints:
- POST /api/execute - Start an agent task
- GET /api/status/{job_id} - Get task status
- GET /api/download/{job_id} - Download completed files
- GET /api/health - Health check
"""

import os
import json
import uuid
import asyncio
import re
from datetime import datetime
from typing import Optional, Dict, Any, List
from pathlib import Path

from fastapi import FastAPI, HTTPException, BackgroundTasks, Header
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import zipfile
import io
import anthropic

# Document generation
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Initialize FastAPI
app = FastAPI(
    title="NOVA Agent Server",
    description="Autonomous Training Agent Execution Server v3.0 - Amplified Outputs",
    version="3.0.0"
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory job storage
jobs: Dict[str, Dict[str, Any]] = {}

# Output directory
OUTPUT_DIR = Path("/tmp/nova-outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

# API Keys
API_SECRET = os.getenv("NOVA_API_SECRET", "")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

# Initialize Anthropic client
claude_client = None
if ANTHROPIC_API_KEY:
    claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


# ============================================================================
# REQUEST/RESPONSE MODELS
# ============================================================================

class TaskRequest(BaseModel):
    job_id: str
    agent: str
    parameters: Dict[str, Any] = {}
    callback_url: Optional[str] = None


class TaskResponse(BaseModel):
    job_id: str
    status: str
    message: str


class StatusResponse(BaseModel):
    job_id: str
    status: str
    progress: int = 0
    current_step: Optional[str] = None
    steps_completed: list = []
    error: Optional[str] = None
    created_at: Optional[str] = None
    completed_at: Optional[str] = None


# ============================================================================
# AUTHENTICATION
# ============================================================================

def verify_auth(authorization: Optional[str] = Header(None)):
    if API_SECRET and authorization != f"Bearer {API_SECRET}":
        raise HTTPException(status_code=401, detail="Unauthorized")


# ============================================================================
# API ENDPOINTS
# ============================================================================

@app.get("/api/health")
async def health_check():
    return {
        "status": "healthy",
        "service": "NOVA Agent Server",
        "version": "3.0.0",
        "claude_configured": claude_client is not None,
        "document_formats": ["docx", "xlsx"],
        "output_spec": "Amplified v1.0",
        "timestamp": datetime.utcnow().isoformat()
    }


@app.post("/api/execute", response_model=TaskResponse)
async def execute_task(
    request: TaskRequest,
    background_tasks: BackgroundTasks,
    authorization: Optional[str] = Header(None)
):
    print(f"[NOVA] Execute request received: agent={request.agent}, job_id={request.job_id}")
    verify_auth(authorization)
    
    job_id = request.job_id
    
    # Valid agents
    valid_agents = ['analysis', 'design', 'delivery', 'evaluation', 'full-package']
    
    # Support legacy names
    agent = request.agent
    if agent == 'tna':
        agent = 'analysis'
    elif agent == 'course-generator':
        agent = 'evaluation'
    elif agent == 'full-package':
        agent = 'evaluation'
    
    if agent not in valid_agents:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid agent: {request.agent}. Valid: {valid_agents}"
        )
    
    # Initialize job
    jobs[job_id] = {
        "job_id": job_id,
        "agent": agent,
        "parameters": request.parameters,
        "status": "queued",
        "progress": 0,
        "current_step": "Initializing...",
        "steps_completed": [],
        "error": None,
        "created_at": datetime.utcnow().isoformat(),
        "completed_at": None,
        "output_dir": str(OUTPUT_DIR / job_id)
    }
    
    # Create output directory
    job_output_dir = OUTPUT_DIR / job_id
    job_output_dir.mkdir(exist_ok=True)
    
    # Start agent execution
    print(f"[NOVA] Starting background task for job {job_id}")
    background_tasks.add_task(run_agent, job_id, agent, request.parameters)
    
    print(f"[NOVA] Returning queued response for job {job_id}")
    return TaskResponse(
        job_id=job_id,
        status="queued",
        message=f"Agent '{agent}' task queued for execution"
    )


@app.get("/api/status/{job_id}", response_model=StatusResponse)
async def get_status(job_id: str, authorization: Optional[str] = Header(None)):
    verify_auth(authorization)
    
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    return StatusResponse(
        job_id=job["job_id"],
        status=job["status"],
        progress=job["progress"],
        current_step=job["current_step"],
        steps_completed=job["steps_completed"],
        error=job["error"],
        created_at=job["created_at"],
        completed_at=job["completed_at"]
    )


@app.get("/api/download/{job_id}")
async def download_files(
    job_id: str,
    file: Optional[str] = None,
    authorization: Optional[str] = Header(None)
):
    verify_auth(authorization)
    
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    
    if job["status"] != "completed":
        raise HTTPException(
            status_code=400,
            detail=f"Job not completed. Current status: {job['status']}"
        )
    
    job_output_dir = Path(job["output_dir"])
    
    if not job_output_dir.exists():
        raise HTTPException(status_code=404, detail="Output files not found")
    
    # If specific file requested
    if file:
        file_path = job_output_dir / file
        if not file_path.exists():
            raise HTTPException(status_code=404, detail=f"File not found: {file}")
        
        media_type = "application/octet-stream"
        if file.endswith('.docx'):
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif file.endswith('.xlsx'):
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        
        return FileResponse(file_path, filename=file, media_type=media_type)
    
    # Return ZIP of all files
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in job_output_dir.rglob('*'):
            if file_path.is_file():
                arcname = file_path.relative_to(job_output_dir)
                zf.write(file_path, arcname)
    
    zip_buffer.seek(0)
    
    role_title = job["parameters"].get("role_title", "Package")
    safe_title = "".join(c for c in role_title if c.isalnum() or c in (' ', '-', '_')).strip()
    safe_title = safe_title.replace(' ', '_')
    filename = f"NOVA_{safe_title}_{job_id[:8]}.zip"
    
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ============================================================================
# DOCUMENT STYLING UTILITIES
# ============================================================================

def create_styled_document(title: str, subtitle: str = "", framework: str = "UK") -> Document:
    """Create a professionally styled Word document"""
    doc = Document()
    
    styles = doc.styles
    
    # Title style
    title_style = styles['Title']
    title_style.font.name = 'Arial'
    title_style.font.size = Pt(24)
    title_style.font.bold = True
    title_style.font.color.rgb = RGBColor(0, 51, 102)
    
    # Heading 1
    h1_style = styles['Heading 1']
    h1_style.font.name = 'Arial'
    h1_style.font.size = Pt(16)
    h1_style.font.bold = True
    h1_style.font.color.rgb = RGBColor(0, 51, 102)
    
    # Heading 2
    h2_style = styles['Heading 2']
    h2_style.font.name = 'Arial'
    h2_style.font.size = Pt(14)
    h2_style.font.bold = True
    h2_style.font.color.rgb = RGBColor(0, 51, 102)
    
    # Heading 3
    h3_style = styles['Heading 3']
    h3_style.font.name = 'Arial'
    h3_style.font.size = Pt(12)
    h3_style.font.bold = True
    h3_style.font.color.rgb = RGBColor(0, 51, 102)
    
    # Normal text
    normal_style = styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(11)
    
    # Set margins
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)
    
    # Add header
    header = doc.sections[0].header
    header_para = header.paragraphs[0]
    header_para.text = f"NOVA™ Training Documentation | {framework} Framework"
    header_para.style.font.size = Pt(9)
    header_para.style.font.color.rgb = RGBColor(128, 128, 128)
    
    # Add footer
    footer = doc.sections[0].footer
    footer_para = footer.paragraphs[0]
    footer_para.text = "Classification: OFFICIAL"
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    return doc


def add_title_page(doc: Document, title: str, subtitle: str, metadata: Dict[str, str]):
    """Add a professional title page"""
    for _ in range(3):
        doc.add_paragraph()
    
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run(title)
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = RGBColor(0, 51, 102)
    
    if subtitle:
        sub_para = doc.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = sub_para.add_run(subtitle)
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(80, 80, 80)
    
    for _ in range(4):
        doc.add_paragraph()
    
    table = doc.add_table(rows=len(metadata), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i, (key, value) in enumerate(metadata.items()):
        row = table.rows[i]
        row.cells[0].text = key
        row.cells[1].text = value
        row.cells[0].paragraphs[0].runs[0].bold = True
    
    doc.add_page_break()


def add_section_heading(doc: Document, text: str, level: int = 1):
    """Add a section heading"""
    doc.add_heading(text, level=level)


def add_table_from_data(doc: Document, headers: List[str], rows: List[List[str]], 
                         header_color: str = "003366"):
    """Add a formatted table"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    
    header_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = header
        para = cell.paragraphs[0]
        para.runs[0].bold = True
        para.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), header_color)
        cell._tc.get_or_add_tcPr().append(shading)
    
    for i, row_data in enumerate(rows):
        row = table.rows[i + 1]
        for j, cell_text in enumerate(row_data):
            row.cells[j].text = str(cell_text)
    
    doc.add_paragraph()


# ============================================================================
# CLAUDE API - INCREASED TOKEN LIMIT FOR AMPLIFIED OUTPUTS
# ============================================================================

async def call_claude(prompt: str, system_prompt: str = None, max_tokens: int = 16000) -> str:
    """Call Claude API to generate content"""
    if not claude_client:
        print("[NOVA] WARNING: Claude API not configured")
        return "[Claude API not configured - please set ANTHROPIC_API_KEY]"
    
    try:
        print(f"[NOVA] Calling Claude API (prompt length: {len(prompt)}, max_tokens: {max_tokens})")
        messages = [{"role": "user", "content": prompt}]
        
        kwargs = {
            "model": "claude-sonnet-4-20250514",
            "max_tokens": max_tokens,
            "messages": messages
        }
        
        if system_prompt:
            kwargs["system"] = system_prompt
        
        response = claude_client.messages.create(**kwargs)
        result = response.content[0].text
        print(f"[NOVA] Claude API response received (length: {len(result)})")
        return result
    except Exception as e:
        print(f"[NOVA] Claude API error: {str(e)}")
        return f"[Error calling Claude API: {str(e)}]"


# ============================================================================
# AMPLIFIED SYSTEM PROMPT
# ============================================================================

TRAINING_SYSTEM_PROMPT = """You are NOVA, an expert training analysis and design system generating AMPLIFIED, COMPREHENSIVE outputs that meet Defence procurement standards. You are proficient in:

- UK Defence Systems Approach to Training (DSAT) - JSP 822 V7.0 and DTSM 1-5 (2024 Edition)
- US Army Systems Approach to Training (SAT) - TRADOC 350-70 / ADDIE / ISD
- NATO Training - Bi-SC Directive 75-7
- ASD/AIA S6000T Training Analysis and Design Standard
- Corporate Learning & Development methodologies
- Competency-based training frameworks

CRITICAL OUTPUT REQUIREMENTS:
1. Generate COMPREHENSIVE, DETAILED outputs - never minimal or brief
2. All tables must be FULLY POPULATED with realistic, specific data
3. Provide RATIONALE and JUSTIFICATION for all decisions
4. Include QUANTIFIED METRICS (costs, percentages, timeframes) wherever possible
5. Reference appropriate doctrine/standards for compliance claims
6. Cross-reference between documents (RolePS tasks → Gap Analysis → TNR)
7. Use formal professional tone appropriate for official Defence/corporate documentation

When generating JSON:
- Populate ALL fields completely - no placeholders like "TBD" unless genuinely unknown
- Generate realistic, role-specific content
- Meet or exceed minimum quantity requirements specified in prompts
- Provide detailed narrative sections (300+ words where specified)"""


# ============================================================================
# AMPLIFIED CONTENT GENERATION FUNCTIONS
# ============================================================================

async def generate_scoping_content(role_title: str, framework: str, description: str = "") -> Dict:
    """Generate AMPLIFIED scoping report content per Output Amplification Specification"""
    framework_ref = get_framework_reference(framework, "scoping")
    current_date = datetime.utcnow().strftime('%d %B %Y')
    
    prompt = f"""Generate a COMPREHENSIVE Scoping Exercise Report for training analysis.

Role Title: {role_title}
Framework: {framework}
Additional Context: {description if description else 'General training requirement'}
Date: {current_date}

Generate the following as a JSON object. THIS MUST BE COMPREHENSIVE AND DETAILED:

{{
    "introduction": "3-4 paragraphs (~300 words) covering: purpose of this scoping exercise, methodology to be used, expected outcomes, and how this aligns with {framework_ref}",
    
    "background": {{
        "operational_context": "2 paragraphs (~200 words) on current operational environment, tempo, recent developments affecting this role",
        "strategic_context": "2 paragraphs (~200 words) on strategic drivers, organizational transformation, alignment to business objectives",
        "current_training_assessment": "2 paragraphs (~200 words) assessing existing training provision, known gaps, recent reviews"
    }},
    
    "scope_inclusions": [
        {{"area": "inclusion area", "rationale": "why included", "effort_percentage": 15, "key_questions": ["question 1", "question 2"]}}
    ],
    
    "scope_exclusions": [
        {{"area": "exclusion area", "rationale": "why excluded", "addressed_elsewhere": "where/how addressed", "risk_statement": "risk if boundary changed"}}
    ],
    
    "boundaries_matrix": [
        {{"boundary": "boundary name", "in_scope": "what is in", "out_of_scope": "what is out", "rationale": "why", "risk_if_changed": "consequence"}}
    ],
    
    "governance": {{
        "tra_responsibilities": ["list of 5-6 specific TRA responsibilities with approval gates"],
        "tda_responsibilities": ["list of 5-6 specific TDA responsibilities with resource commitments"],
        "escalation_procedures": "description of escalation thresholds and procedures"
    }},
    
    "stakeholders": [
        {{"stakeholder": "name/role", "role_in_project": "what they do", "interest_level": "H/M/L", "influence_level": "H/M/L", "engagement_strategy": "how to engage", "consultation_method": "meetings/workshops/etc", "frequency": "weekly/monthly/etc"}}
    ],
    
    "raci_matrix": [
        {{"activity": "TNA activity", "responsible": "role", "accountable": "role", "consulted": "roles", "informed": "roles"}}
    ],
    
    "assumptions": [
        {{"id": "A1", "assumption": "detailed assumption text", "impact_if_invalid": "H/M/L", "validation_method": "how to validate", "owner": "who owns this"}}
    ],
    
    "constraints": [
        {{"id": "C1", "constraint": "detailed constraint text", "type": "Resource/Time/Policy/Technical", "impact": "effect on project", "mitigation": "how to work within"}}
    ],
    
    "dependencies": [
        {{"id": "D1", "dependency": "what we depend on", "owner": "who provides", "required_by": "date/milestone", "risk_if_delayed": "impact", "mitigation": "alternative approach"}}
    ],
    
    "risks": [
        {{"id": "R1", "risk": "detailed risk description", "category": "Stakeholder/Resource/Data/Timeline/Technical/External", "likelihood": "1-5", "impact": "1-5", "risk_score": 0, "mitigation": "detailed mitigation strategy", "owner": "who manages", "status": "Open"}}
    ],
    
    "resource_estimate": {{
        "personnel": [
            {{"role": "role title", "grade": "grade/level", "fte": 0.5, "duration_weeks": 12, "cost_per_week": 1500, "total_cost": 9000, "justification": "why needed"}}
        ],
        "non_personnel": [
            {{"item": "item description", "quantity": 1, "unit_cost": 500, "total_cost": 500, "justification": "why needed"}}
        ],
        "contingency_rate": 15,
        "contingency_justification": "based on risk assessment",
        "base_cost": 0,
        "contingency_amount": 0,
        "grand_total": 0
    }},
    
    "timeline": {{
        "phases": [
            {{"phase": "phase name", "activities": ["activity 1", "activity 2"], "start_week": 1, "end_week": 4, "milestone": "milestone at end", "deliverables": ["deliverable 1"]}}
        ],
        "milestones": [
            {{"milestone": "milestone name", "date": "Week X / Date", "criteria": "what defines completion", "dependencies": "what must be done first"}}
        ],
        "critical_path": "description of critical path activities",
        "total_duration_weeks": 16
    }},
    
    "recommendations": [
        {{"recommendation": "specific recommendation", "rationale": "why recommended", "priority": "High/Medium/Low", "owner": "who should action"}}
    ]
}}

REQUIREMENTS:
- Generate MINIMUM 8-10 stakeholders with full Interest/Influence/Engagement details
- Generate MINIMUM 12 assumptions with validation methods
- Generate MINIMUM 10 constraints
- Generate MINIMUM 8 dependencies
- Generate MINIMUM 12 risks across ALL categories (Stakeholder, Resource, Data, Timeline, Technical, External)
- Generate 4-6 personnel resources and 4-6 non-personnel costs
- Generate 5-6 timeline phases with 20-25 total activities
- Generate MINIMUM 10 milestones
- Generate 6-8 scope inclusions and 4-6 exclusions
- Generate 6-8 boundaries in matrix
- Generate 10-12 RACI activities
- All narrative sections must be 200+ words
- Calculate all costs and totals

Be specific and realistic for a {role_title} role.
Return ONLY the JSON object, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT, max_tokens=16000)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except Exception as e:
        print(f"[NOVA] JSON parse error in scoping: {e}")
    
    # Return minimal fallback
    return {
        "introduction": f"This Scoping Exercise Report initiates comprehensive training analysis for the {role_title} role.",
        "background": {"operational_context": "Analysis required.", "strategic_context": "Strategic alignment required.", "current_training_assessment": "Current provision assessment required."},
        "scope_inclusions": [{"area": "Core role competencies", "rationale": "Primary focus", "effort_percentage": 100, "key_questions": ["What are the key tasks?"]}],
        "scope_exclusions": [],
        "stakeholders": [{"stakeholder": "Training Manager", "role_in_project": "Coordination", "interest_level": "H", "influence_level": "H", "engagement_strategy": "Regular meetings", "consultation_method": "Workshops", "frequency": "Weekly"}],
        "assumptions": [{"id": "A1", "assumption": "Current requirements are documented", "impact_if_invalid": "H", "validation_method": "Review documentation", "owner": "TRA"}],
        "constraints": [],
        "dependencies": [],
        "risks": [{"id": "R1", "risk": "Scope creep", "category": "Technical", "likelihood": "3", "impact": "3", "risk_score": 9, "mitigation": "Regular reviews", "owner": "PM", "status": "Open"}],
        "resource_estimate": {"personnel": [], "non_personnel": [], "grand_total": 0},
        "timeline": {"phases": [], "milestones": [], "total_duration_weeks": 16},
        "recommendations": [{"recommendation": "Proceed with analysis", "rationale": "Training need identified", "priority": "High", "owner": "TRA"}]
    }


async def generate_role_tasks(role_title: str, framework: str, description: str = "") -> Dict:
    """Generate AMPLIFIED Role Performance Statement with duties and sub-tasks"""
    framework_ref = get_framework_reference(framework, "roleps")
    
    prompt = f"""Generate a COMPREHENSIVE Role Performance Statement / Task Analysis for:

Role Title: {role_title}
Framework: {framework}
Reference: {framework_ref}
Additional Context: {description if description else 'General role requirements'}

Generate as a JSON object with FULL HEADER BLOCK and DUTY-BASED TASK STRUCTURE:

{{
    "header": {{
        "role_title": "{role_title}",
        "tdw_number": "TDW-2026-XXX",
        "duty_title": "Parent duty/job family",
        "duty_number": "D-XXX",
        "tra": "Training Requirements Authority name",
        "roleps_serial": "RPS-2026-001",
        "tda": "Training Delivery Authority name", 
        "issue_version": "1.0",
        "review_date": "January 2027",
        "security_classification": "OFFICIAL"
    }},
    
    "duties": [
        {{
            "duty_number": "1",
            "duty_title": "DUTY TITLE IN CAPS",
            "duty_description": "2-3 sentence description of this duty area",
            "tasks": [
                {{
                    "task_number": "1.1",
                    "performance": "Observable action verb + specific object + qualifier. Must be specific and measurable. Example: Configure network firewall rules to permit authorised traffic while blocking defined threat signatures",
                    "conditions": "Specific conditions including: environment (indoor/outdoor/simulated), equipment available (list specific systems/tools), references permitted (manuals/SOPs), supervision level (alone/supervised/as team lead), time constraints (under time pressure/within X hours)",
                    "standards": "Measurable criteria including: accuracy (100%/within 5% tolerance), time (within 30 minutes), quality (IAW SOP XYZ), frequency (on every occasion/95% of occasions), error tolerance (zero critical errors/max 2 minor deviations)",
                    "category": "FT",
                    "ksa": {{
                        "knowledge": ["List 2-3 specific knowledge requirements"],
                        "skills": ["List 2-3 specific skill requirements"],
                        "attitudes": ["List 1-2 attitude/behaviour requirements"]
                    }},
                    "criticality": "Safety-Critical/Mission-Critical/Important/Desirable",
                    "sub_tasks": [
                        {{
                            "sub_task_number": "1.1.1",
                            "performance": "Sub-task performance statement",
                            "conditions": "Sub-task specific conditions",
                            "standards": "Sub-task specific standards",
                            "category": "FT",
                            "ksa_type": "K/S/A"
                        }}
                    ]
                }}
            ]
        }}
    ],
    
    "summary": {{
        "total_duties": 0,
        "total_tasks": 0,
        "total_sub_tasks": 0,
        "category_breakdown": {{
            "FT": 0,
            "WPT": 0,
            "OJT": 0,
            "CBT": 0,
            "RTGS": 0
        }},
        "criticality_breakdown": {{
            "safety_critical": 0,
            "mission_critical": 0,
            "important": 0,
            "desirable": 0
        }}
    }}
}}

TRAINING CATEGORIES:
- FT = Formal Training (classroom/structured courses)
- WPT = Workplace Training (structured on-site training)
- OJT = On-the-Job Training (learning while working)
- CBT = Computer-Based Training (e-learning/digital)
- RTGS = Residual Training Gap Statement (not trained - risk accepted)

REQUIREMENTS:
- Generate MINIMUM 6-8 DUTY areas
- Generate MINIMUM 3-5 TASKS per duty
- Generate MINIMUM 2-4 SUB-TASKS per task
- TOTAL task + sub-task lines: MINIMUM 40-60
- Each Performance statement must begin with observable action verb (Bloom's taxonomy)
- Each Conditions statement must specify environment, equipment, references, supervision, time
- Each Standards statement must include accuracy, time, quality, frequency metrics
- Calculate all summary totals accurately

Be specific and realistic for a {role_title} role.
Return ONLY the JSON object, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT, max_tokens=16000)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except Exception as e:
        print(f"[NOVA] JSON parse error in role tasks: {e}")
    
    return {"header": {}, "duties": [], "summary": {}}


async def generate_gap_analysis(role_title: str, framework: str, tasks: Dict) -> Dict:
    """Generate AMPLIFIED Training Gap Analysis per Output Amplification Specification"""
    
    # Extract task summary from duties structure
    task_list = []
    if "duties" in tasks:
        for duty in tasks.get("duties", [])[:4]:
            for task in duty.get("tasks", [])[:3]:
                task_list.append(f"- {task.get('task_number', '')}: {task.get('performance', '')[:100]}")
    task_summary = "\n".join(task_list) if task_list else "Core role tasks"
    
    prompt = f"""Generate a COMPREHENSIVE Training Gap Analysis based on these role tasks.

Role Title: {role_title}
Framework: {framework}
Sample Tasks:
{task_summary}

Generate as a JSON object with FULL DETAIL for each gap:

{{
    "executive_summary": {{
        "total_gaps": 0,
        "critical_gaps": 0,
        "high_priority_gaps": 0,
        "medium_priority_gaps": 0,
        "low_priority_gaps": 0,
        "capability_risk_statement": "2-3 sentence statement on overall capability risk if gaps not addressed",
        "resource_implication_summary": "Summary of total resources needed to close all gaps",
        "timeline_for_closure": "Overall timeline estimate"
    }},
    
    "gaps": [
        {{
            "gap_id": "GAP-001",
            "gap_title": "Gap Title",
            "task_reference": "Task X.X: Task title from RolePS",
            "performance_requirement": "Full Performance-Conditions-Standards from RolePS task",
            "current_capability": {{
                "description": "Detailed description of current training/provision (50-75 words)",
                "provision_source": "Course name, provider, duration, last review date",
                "coverage_percentage": 40
            }},
            "required_capability": {{
                "description": "Full description of required competence (50-75 words)",
                "standard_reference": "Reference to doctrine/standard requiring this"
            }},
            "gap_description": "Precise nature of the shortfall - MUST be 75-100 words detailing exactly what is missing and why it matters",
            "gap_type": "Knowledge/Skill/Attitude/Experience/Equipment",
            "root_cause_analysis": "Why does this gap exist? Training not available/outdated/insufficient depth/changed requirements",
            "dif_rating": {{
                "difficulty": 3,
                "importance": 5,
                "frequency": 4,
                "overall_score": 60
            }},
            "criticality_assessment": {{
                "rating": "Mission-Critical",
                "rationale": "Why this criticality rating applies"
            }},
            "operational_impact": "Specific consequences if gap remains: mission failure modes, safety risks, efficiency losses (50-75 words)",
            "risk_if_unaddressed": {{
                "likelihood": 4,
                "impact": 5,
                "risk_score": 20,
                "risk_description": "What could go wrong"
            }},
            "recommended_intervention": {{
                "intervention_type": "Formal Course/E-Learning/OJT/Blended",
                "description": "Detailed description of recommended training solution (50-75 words)",
                "duration": "X days/hours",
                "method": "Primary delivery method",
                "assessment_approach": "How competence will be verified"
            }},
            "alternative_interventions": [
                {{"option": "Alternative 1", "pros": "advantages", "cons": "disadvantages"}},
                {{"option": "Alternative 2", "pros": "advantages", "cons": "disadvantages"}}
            ],
            "resource_estimate": {{
                "development_cost": 0,
                "delivery_cost_per_learner": 0,
                "total_estimated_cost": 0,
                "time_to_develop": "X weeks",
                "time_to_deliver": "X days"
            }},
            "priority": "Critical/High/Medium/Low",
            "priority_justification": "Why this priority rating",
            "suggested_timeline": "When gap should be addressed (immediate/Q1/Q2/Year 2)",
            "success_metrics": ["How we will know gap is closed - list 2-3 measurable criteria"]
        }}
    ],
    
    "prioritisation_matrix": {{
        "quick_wins": ["Gap IDs - low effort, high impact"],
        "major_projects": ["Gap IDs - high effort, high impact"],
        "fill_ins": ["Gap IDs - low effort, low impact"],
        "hard_slogs": ["Gap IDs - high effort, low impact"]
    }},
    
    "summary_by_type": {{
        "knowledge_gaps": 0,
        "skill_gaps": 0,
        "attitude_gaps": 0,
        "experience_gaps": 0,
        "equipment_gaps": 0
    }},
    
    "recommendations": [
        {{"recommendation": "specific recommendation", "priority": "High/Medium/Low", "owner": "who should action", "timeline": "when"}}
    ]
}}

REQUIREMENTS:
- Generate MINIMUM 12-15 gaps with FULL DETAIL for each
- Each gap_description must be 75-100 words
- Each operational_impact must be 50-75 words  
- Include gaps across ALL types (Knowledge, Skill, Attitude, Experience, Equipment)
- Calculate all scores and totals accurately
- Provide 2-3 alternative interventions per gap
- Generate 2-3 success metrics per gap
- Populate prioritisation_matrix with gap IDs

Be specific and realistic for a {role_title} role.
Return ONLY the JSON object, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT, max_tokens=16000)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except Exception as e:
        print(f"[NOVA] JSON parse error in gap analysis: {e}")
    
    return {"executive_summary": {}, "gaps": [], "recommendations": []}


async def generate_tnr_content(role_title: str, framework: str, tasks: Dict, gaps: Dict) -> Dict:
    """Generate AMPLIFIED Training Needs Report per Output Amplification Specification"""
    
    total_tasks = tasks.get("summary", {}).get("total_tasks", 0)
    total_gaps = gaps.get("executive_summary", {}).get("total_gaps", 0)
    critical_gaps = gaps.get("executive_summary", {}).get("critical_gaps", 0)
    
    prompt = f"""Generate a COMPREHENSIVE Training Needs Report based on completed analysis.

Role Title: {role_title}
Framework: {framework}
Tasks Identified: {total_tasks}
Gaps Identified: {total_gaps}
Critical Gaps: {critical_gaps}

Generate as a JSON object:

{{
    "executive_summary": {{
        "problem_statement": "1-2 sentences clearly stating the training need",
        "methodology_summary": "How the analysis was conducted (2-3 sentences)",
        "key_findings": [
            "Finding 1 with quantified impact",
            "Finding 2 with quantified impact",
            "Finding 3 with quantified impact",
            "Finding 4 with quantified impact",
            "Finding 5 with quantified impact",
            "Finding 6 with quantified impact",
            "Finding 7 with quantified impact"
        ],
        "principal_recommendation": "Single sentence stating the main recommendation",
        "resource_headline": "Total cost £X over Y months",
        "top_risks": [
            {{"risk": "Risk 1 if not approved", "impact": "consequence"}},
            {{"risk": "Risk 2 if not approved", "impact": "consequence"}},
            {{"risk": "Risk 3 if not approved", "impact": "consequence"}}
        ],
        "recommended_decision": "Approve/Modify/Reject with specific conditions"
    }},
    
    "background": "3-4 paragraphs (~400 words) providing full context for the training need, strategic drivers, and current situation",
    
    "analysis_findings": {{
        "role_analysis_summary": {{
            "complexity_assessment": "Low/Medium/High/Very High",
            "complexity_justification": "Why this complexity rating",
            "comparison_to_similar_roles": "How this compares to related roles",
            "career_pathway_context": "Where this role fits in career progression",
            "competency_framework_alignment": "How aligned to organizational competency framework"
        }},
        "task_analysis_summary": {{
            "total_tasks": 0,
            "ft_tasks": 0,
            "wpt_tasks": 0,
            "ojt_tasks": 0,
            "cbt_tasks": 0,
            "safety_critical_tasks": 0,
            "mission_critical_tasks": 0
        }},
        "ksa_analysis": {{
            "knowledge_requirements": [
                {{"area": "Knowledge area 1", "level": "Foundation/Intermediate/Advanced", "priority": "H/M/L"}}
            ],
            "skill_requirements": [
                {{"skill": "Skill 1", "type": "Technical/Interpersonal/Cognitive", "proficiency": "Novice/Competent/Expert", "priority": "H/M/L"}}
            ],
            "attitude_requirements": [
                {{"attitude": "Attitude 1", "importance": "H/M/L", "development_approach": "How to develop"}}
            ]
        }},
        "gap_summary": {{
            "total_gaps": 0,
            "by_priority": {{"critical": 0, "high": 0, "medium": 0, "low": 0}},
            "by_type": {{"knowledge": 0, "skill": 0, "attitude": 0}},
            "estimated_total_cost_to_close": 0
        }}
    }},
    
    "training_options": [
        {{
            "option_id": "A",
            "option_name": "Option Name",
            "description": "Comprehensive description of this option (200-300 words)",
            "delivery_methodology": {{
                "primary_method": "Residential/Distributed/Blended/Online",
                "secondary_methods": ["Supporting method 1", "Supporting method 2"],
                "technology_requirements": ["LMS", "Simulation", "Equipment"],
                "assessment_approach": "Formative and summative methods"
            }},
            "programme_structure": [
                {{"module": "Module 1", "duration": "X days", "topics": ["topic 1", "topic 2"], "delivery": "Classroom/Online"}}
            ],
            "coverage_analysis": {{
                "gaps_fully_addressed": ["GAP-001", "GAP-002"],
                "gaps_partially_addressed": ["GAP-003"],
                "gaps_not_addressed": ["GAP-004"],
                "coverage_percentage": 85
            }},
            "resource_requirements": {{
                "personnel": [{{"role": "Trainer", "quantity": 2, "duration": "X weeks", "cost": 0}}],
                "facilities": [{{"facility": "Training Room", "quantity": 1, "duration": "X days", "cost": 0}}],
                "equipment": [{{"item": "Equipment item", "quantity": 5, "cost": 0}}]
            }},
            "cost_benefit_analysis": {{
                "five_year_costs": {{
                    "year_1": {{"development": 0, "delivery": 0, "infrastructure": 0, "personnel": 0, "total": 0}},
                    "year_2": {{"development": 0, "delivery": 0, "infrastructure": 0, "personnel": 0, "total": 0}},
                    "year_3": {{"development": 0, "delivery": 0, "infrastructure": 0, "personnel": 0, "total": 0}},
                    "year_4": {{"development": 0, "delivery": 0, "infrastructure": 0, "personnel": 0, "total": 0}},
                    "year_5": {{"development": 0, "delivery": 0, "infrastructure": 0, "personnel": 0, "total": 0}},
                    "total_5_year": 0
                }},
                "benefits": [{{"benefit": "Benefit description", "value": "Quantified where possible", "timing": "When realised"}}],
                "roi_calculation": {{
                    "total_investment": 0,
                    "annual_benefit": 0,
                    "payback_period_years": 0,
                    "five_year_roi_percentage": 0
                }}
            }},
            "risk_assessment": [{{"risk": "Risk 1", "likelihood": "H/M/L", "impact": "H/M/L", "mitigation": "How to mitigate"}}],
            "swot_analysis": {{
                "strengths": ["Strength 1", "Strength 2", "Strength 3"],
                "weaknesses": ["Weakness 1", "Weakness 2", "Weakness 3"],
                "opportunities": ["Opportunity 1", "Opportunity 2"],
                "threats": ["Threat 1", "Threat 2"]
            }},
            "effectiveness_rating": {{"rating": "High/Medium/Low", "justification": "Why"}},
            "efficiency_rating": {{"rating": "High/Medium/Low", "justification": "Why"}},
            "feasibility_rating": {{"rating": "High/Medium/Low", "justification": "Why"}}
        }}
    ],
    
    "recommended_solution": {{
        "selected_option": "Option ID and name",
        "selection_rationale": "Why this option recommended (150-200 words)",
        "training_statement_preview": {{
            "tps_summary": "Training Performance Statement summary",
            "wps_summary": "Workplace Training Statement summary",
            "rtgs_summary": "Residual Training Gap Statement"
        }},
        "implementation_approach": [
            {{"phase": "Phase 1: Development", "duration": "X months", "activities": ["activity 1", "activity 2"], "milestone": "Milestone"}}
        ],
        "success_criteria": [{{"criterion": "Success criterion 1", "measure": "How measured", "target": "Target value"}}]
    }},
    
    "resource_implications": {{
        "total_development_cost": 0,
        "annual_delivery_cost": 0,
        "five_year_total_cost": 0,
        "funding_source": "How this will be funded",
        "budget_line": "Which budget this falls under"
    }},
    
    "risk_assessment": [
        {{"risk": "Risk if approved", "likelihood": "H/M/L", "impact": "H/M/L", "mitigation": "Strategy", "owner": "Who owns"}},
        {{"risk": "Risk if NOT approved", "likelihood": "H/M/L", "impact": "H/M/L", "consequence": "What happens"}}
    ],
    
    "recommendations": [
        {{"number": 1, "recommendation": "Specific recommendation", "rationale": "Why", "owner": "Who should action", "timeline": "When"}}
    ],
    
    "next_steps": [
        {{"step": 1, "action": "Action required", "owner": "Who", "deadline": "When"}}
    ],
    
    "approval_requirements": {{
        "approving_authority": "Who approves this TNR",
        "approval_date_required": "When approval needed",
        "conditions": "Any conditions for approval"
    }}
}}

REQUIREMENTS:
- Executive summary key_findings: MINIMUM 7 items with quantified impacts
- Background: MINIMUM 400 words
- Knowledge requirements: MINIMUM 12 areas
- Skill requirements: MINIMUM 12 skills
- Attitude requirements: MINIMUM 6 attitudes
- Training options: MINIMUM 3-4 complete options with full 5-year cost analysis
- Each option description: MINIMUM 200 words
- Full ROI calculation for each option
- SWOT analysis with 3+ items per category
- Implementation approach: MINIMUM 4 phases with 12+ months timeline
- Risk assessment: MINIMUM 8 risks (both approval and non-approval scenarios)
- Recommendations: MINIMUM 6 specific recommendations
- Calculate all costs and totals accurately

Return ONLY the JSON object, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT, max_tokens=16000)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except Exception as e:
        print(f"[NOVA] JSON parse error in TNR: {e}")
    
    return {"executive_summary": {}, "background": "", "training_options": [], "recommendations": []}


# ============================================================================
# FRAMEWORK UTILITIES
# ============================================================================

def get_framework_reference(framework: str, doc_type: str) -> str:
    """Get appropriate framework reference"""
    refs = {
        "UK": {
            "scoping": "DTSM 2 Section 1.2",
            "roleps": "JSP 822 / DTSM 2 Section 1.3",
            "gap": "DTSM 2 Section 1.4",
            "tnr": "DTSM 2 Section 1.7",
            "to": "DTSM 3 Section 2.1",
            "eo": "DTSM 3 Section 2.3"
        },
        "US": {
            "scoping": "TRADOC 350-70",
            "roleps": "Task Analysis (TRADOC)",
            "gap": "Training Needs Assessment",
            "tnr": "Training Requirements Analysis",
            "to": "Terminal Learning Objectives",
            "eo": "Enabling Learning Objectives"
        },
        "NATO": {
            "scoping": "Bi-SC 75-7",
            "roleps": "Duty Analysis",
            "gap": "Training Gap Assessment",
            "tnr": "Training Requirements Document",
            "to": "Training Outcomes",
            "eo": "Enabling Outcomes"
        },
        "ASD": {
            "scoping": "S6000T Clause 5",
            "roleps": "Task Specification",
            "gap": "Capability Gap Analysis",
            "tnr": "Training Specification",
            "to": "Training Requirements",
            "eo": "Sub-Task Requirements"
        },
        "Commercial": {
            "scoping": "Training Needs Analysis",
            "roleps": "Job Task Analysis",
            "gap": "Skills Gap Assessment",
            "tnr": "Training Requirements Document",
            "to": "Learning Objectives",
            "eo": "Enabling Objectives"
        }
    }
    return refs.get(framework, refs.get("Commercial", {})).get(doc_type, "Training Standards")


def get_framework_term(framework: str, term_type: str) -> str:
    """Get framework-specific terminology"""
    terms = {
        "UK": {"roleps": "Role Performance Statement", "tnr": "Training Needs Report",
               "to": "Training Objective", "eo": "Enabling Objective", "klp": "Key Learning Point"},
        "US": {"roleps": "Task List", "tnr": "Training Requirements Analysis",
               "to": "Terminal Learning Objective", "eo": "Enabling Learning Objective", "klp": "Learning Step"},
        "NATO": {"roleps": "Duty Analysis", "tnr": "Training Requirements Document",
                 "to": "Training Outcome", "eo": "Enabling Outcome", "klp": "Learning Point"},
        "ASD": {"roleps": "Task Specification", "tnr": "Training Specification",
                "to": "Training Requirement", "eo": "Sub-Task Requirement", "klp": "Task Element"},
        "Commercial": {"roleps": "Job Task Analysis", "tnr": "Training Requirements Document",
                       "to": "Learning Objective", "eo": "Enabling Objective", "klp": "Learning Point"}
    }
    return terms.get(framework, terms.get("Commercial", {})).get(term_type, term_type)


# ============================================================================
# AMPLIFIED DOCUMENT BUILDERS
# ============================================================================

def build_scoping_report(role_title: str, framework: str, content: Dict, 
                         output_path: Path) -> str:
    """Build AMPLIFIED Scoping Report document"""
    doc = create_styled_document("Scoping Exercise Report", role_title, framework)
    
    add_title_page(doc, "SCOPING EXERCISE REPORT", role_title, {
        "Document Type": f"Scoping Exercise Report ({get_framework_reference(framework, 'scoping')})",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # 1. Introduction
    add_section_heading(doc, "1. INTRODUCTION")
    doc.add_paragraph(content.get("introduction", ""))
    
    # 2. Background (Amplified)
    add_section_heading(doc, "2. BACKGROUND AND CONTEXT")
    background = content.get("background", {})
    if isinstance(background, dict):
        doc.add_heading("2.1 Operational Environment Analysis", level=2)
        doc.add_paragraph(background.get("operational_context", ""))
        
        doc.add_heading("2.2 Strategic Context", level=2)
        doc.add_paragraph(background.get("strategic_context", ""))
        
        doc.add_heading("2.3 Current Training Provision Assessment", level=2)
        doc.add_paragraph(background.get("current_training_assessment", ""))
    else:
        doc.add_paragraph(str(background))
    
    # 3. Scope (Amplified)
    add_section_heading(doc, "3. SCOPE")
    
    doc.add_heading("3.1 Inclusions", level=2)
    inclusions = content.get("scope_inclusions", [])
    if inclusions and isinstance(inclusions[0], dict):
        headers = ["Area", "Rationale", "Effort %", "Key Questions"]
        rows = [[
            i.get("area", ""),
            i.get("rationale", ""),
            str(i.get("effort_percentage", "")),
            "; ".join(i.get("key_questions", []))
        ] for i in inclusions]
        add_table_from_data(doc, headers, rows)
    else:
        for item in inclusions:
            doc.add_paragraph(f"• {item}", style='List Bullet')
    
    doc.add_heading("3.2 Exclusions", level=2)
    exclusions = content.get("scope_exclusions", [])
    if exclusions and isinstance(exclusions[0], dict):
        headers = ["Area", "Rationale", "Addressed Elsewhere", "Risk Statement"]
        rows = [[
            e.get("area", ""),
            e.get("rationale", ""),
            e.get("addressed_elsewhere", ""),
            e.get("risk_statement", "")
        ] for e in exclusions]
        add_table_from_data(doc, headers, rows)
    else:
        for item in exclusions:
            doc.add_paragraph(f"• {item}", style='List Bullet')
    
    doc.add_heading("3.3 Boundaries Matrix", level=2)
    boundaries = content.get("boundaries_matrix", [])
    if boundaries:
        headers = ["Boundary", "In Scope", "Out of Scope", "Rationale", "Risk if Changed"]
        rows = [[
            b.get("boundary", ""),
            b.get("in_scope", ""),
            b.get("out_of_scope", ""),
            b.get("rationale", ""),
            b.get("risk_if_changed", "")
        ] for b in boundaries]
        add_table_from_data(doc, headers, rows)
    
    # 4. Governance (Amplified)
    add_section_heading(doc, "4. GOVERNANCE")
    governance = content.get("governance", {})
    if governance:
        doc.add_heading("4.1 TRA Responsibilities", level=2)
        for resp in governance.get("tra_responsibilities", []):
            doc.add_paragraph(f"• {resp}", style='List Bullet')
        
        doc.add_heading("4.2 TDA Responsibilities", level=2)
        for resp in governance.get("tda_responsibilities", []):
            doc.add_paragraph(f"• {resp}", style='List Bullet')
        
        doc.add_heading("4.3 Escalation Procedures", level=2)
        doc.add_paragraph(governance.get("escalation_procedures", ""))
    
    # 5. Stakeholder Analysis (Amplified)
    add_section_heading(doc, "5. STAKEHOLDER ANALYSIS")
    stakeholders = content.get("stakeholders", [])
    if stakeholders:
        headers = ["Stakeholder", "Role", "Interest", "Influence", "Engagement Strategy", "Method", "Frequency"]
        rows = [[
            s.get("stakeholder", ""),
            s.get("role_in_project", s.get("role", "")),
            s.get("interest_level", ""),
            s.get("influence_level", ""),
            s.get("engagement_strategy", ""),
            s.get("consultation_method", ""),
            s.get("frequency", "")
        ] for s in stakeholders]
        add_table_from_data(doc, headers, rows)
    
    # 5.1 RACI Matrix
    raci = content.get("raci_matrix", [])
    if raci:
        doc.add_heading("5.1 RACI Matrix", level=2)
        headers = ["Activity", "Responsible", "Accountable", "Consulted", "Informed"]
        rows = [[
            r.get("activity", ""),
            r.get("responsible", ""),
            r.get("accountable", ""),
            r.get("consulted", ""),
            r.get("informed", "")
        ] for r in raci]
        add_table_from_data(doc, headers, rows)
    
    # 6. Assumptions Register (Amplified)
    add_section_heading(doc, "6. ASSUMPTIONS REGISTER")
    assumptions = content.get("assumptions", [])
    if assumptions and isinstance(assumptions[0], dict):
        headers = ["ID", "Assumption", "Impact if Invalid", "Validation Method", "Owner"]
        rows = [[
            a.get("id", ""),
            a.get("assumption", ""),
            a.get("impact_if_invalid", ""),
            a.get("validation_method", ""),
            a.get("owner", "")
        ] for a in assumptions]
        add_table_from_data(doc, headers, rows)
    else:
        for item in assumptions:
            doc.add_paragraph(f"• {item}", style='List Bullet')
    
    # 7. Constraints Register
    add_section_heading(doc, "7. CONSTRAINTS REGISTER")
    constraints = content.get("constraints", [])
    if constraints:
        headers = ["ID", "Constraint", "Type", "Impact", "Mitigation"]
        rows = [[
            c.get("id", ""),
            c.get("constraint", ""),
            c.get("type", ""),
            c.get("impact", ""),
            c.get("mitigation", "")
        ] for c in constraints]
        add_table_from_data(doc, headers, rows)
    
    # 8. Dependencies Register
    add_section_heading(doc, "8. DEPENDENCIES REGISTER")
    dependencies = content.get("dependencies", [])
    if dependencies:
        headers = ["ID", "Dependency", "Owner", "Required By", "Risk if Delayed", "Mitigation"]
        rows = [[
            d.get("id", ""),
            d.get("dependency", ""),
            d.get("owner", ""),
            d.get("required_by", ""),
            d.get("risk_if_delayed", ""),
            d.get("mitigation", "")
        ] for d in dependencies]
        add_table_from_data(doc, headers, rows)
    
    # 9. Risk Assessment
    add_section_heading(doc, "9. RISK ASSESSMENT")
    risks = content.get("risks", [])
    if risks:
        headers = ["ID", "Risk", "Category", "L", "I", "Score", "Mitigation", "Owner"]
        rows = [[
            r.get("id", ""),
            r.get("risk", ""),
            r.get("category", ""),
            str(r.get("likelihood", "")),
            str(r.get("impact", "")),
            str(r.get("risk_score", int(r.get("likelihood", 0) or 0) * int(r.get("impact", 0) or 0))),
            r.get("mitigation", ""),
            r.get("owner", "")
        ] for r in risks]
        add_table_from_data(doc, headers, rows)
    
    # 10. Resource Estimate
    add_section_heading(doc, "10. RESOURCE ESTIMATE")
    res = content.get("resource_estimate", {})
    
    doc.add_heading("10.1 Personnel Resources", level=2)
    personnel = res.get("personnel", [])
    if personnel:
        headers = ["Role", "Grade", "FTE", "Duration (Weeks)", "Cost/Week", "Total Cost", "Justification"]
        rows = [[
            p.get("role", ""),
            p.get("grade", ""),
            str(p.get("fte", "")),
            str(p.get("duration_weeks", "")),
            f"£{p.get('cost_per_week', 0):,}",
            f"£{p.get('total_cost', 0):,}",
            p.get("justification", "")
        ] for p in personnel]
        add_table_from_data(doc, headers, rows)
    
    doc.add_heading("10.2 Non-Personnel Costs", level=2)
    non_personnel = res.get("non_personnel", [])
    if non_personnel:
        headers = ["Item", "Quantity", "Unit Cost", "Total Cost", "Justification"]
        rows = [[
            n.get("item", ""),
            str(n.get("quantity", "")),
            f"£{n.get('unit_cost', 0):,}",
            f"£{n.get('total_cost', 0):,}",
            n.get("justification", "")
        ] for n in non_personnel]
        add_table_from_data(doc, headers, rows)
    
    doc.add_heading("10.3 Cost Summary", level=2)
    base_cost = res.get("base_cost", 0)
    contingency_rate = res.get("contingency_rate", 15)
    contingency_amount = res.get("contingency_amount", 0)
    grand_total = res.get("grand_total", 0)
    doc.add_paragraph(f"Base Cost: £{base_cost:,}")
    doc.add_paragraph(f"Contingency Rate: {contingency_rate}% ({res.get('contingency_justification', 'Based on risk assessment')})")
    doc.add_paragraph(f"Contingency Amount: £{contingency_amount:,}")
    p = doc.add_paragraph(f"Grand Total: £{grand_total:,}")
    p.runs[0].bold = True
    
    # 11. Timeline
    add_section_heading(doc, "11. TIMELINE")
    timeline = content.get("timeline", {})
    
    doc.add_heading("11.1 Phase Breakdown", level=2)
    phases = timeline.get("phases", [])
    if phases:
        headers = ["Phase", "Activities", "Start", "End", "Milestone", "Deliverables"]
        rows = [[
            p.get("phase", ""),
            "; ".join(p.get("activities", [])),
            f"Week {p.get('start_week', '')}",
            f"Week {p.get('end_week', '')}",
            p.get("milestone", ""),
            "; ".join(p.get("deliverables", []))
        ] for p in phases]
        add_table_from_data(doc, headers, rows)
    
    doc.add_heading("11.2 Milestones", level=2)
    milestones = timeline.get("milestones", [])
    if milestones:
        headers = ["Milestone", "Date", "Completion Criteria", "Dependencies"]
        rows = [[
            m.get("milestone", ""),
            m.get("date", ""),
            m.get("criteria", ""),
            m.get("dependencies", "")
        ] for m in milestones]
        add_table_from_data(doc, headers, rows)
    
    doc.add_heading("11.3 Critical Path", level=2)
    doc.add_paragraph(timeline.get("critical_path", "Critical path activities to be identified during detailed planning."))
    doc.add_paragraph(f"Total Estimated Duration: {timeline.get('total_duration_weeks', 16)} weeks")
    
    # 12. Recommendations
    add_section_heading(doc, "12. RECOMMENDATIONS")
    recommendations = content.get("recommendations", [])
    if recommendations and isinstance(recommendations[0], dict):
        for i, rec in enumerate(recommendations, 1):
            doc.add_paragraph(f"{i}. {rec.get('recommendation', '')}")
            doc.add_paragraph(f"   Rationale: {rec.get('rationale', '')}")
            doc.add_paragraph(f"   Priority: {rec.get('priority', '')} | Owner: {rec.get('owner', '')}")
    else:
        for i, item in enumerate(recommendations, 1):
            doc.add_paragraph(f"{i}. {item}")
    
    filename = "01_Scoping_Report.docx"
    doc.save(output_path / filename)
    return filename


def build_role_performance_statement(role_title: str, framework: str, tasks: Dict,
                                      output_path: Path) -> str:
    """Build AMPLIFIED Role Performance Statement document"""
    doc = create_styled_document("Role Performance Statement", role_title, framework)
    
    term = get_framework_term(framework, "roleps")
    
    header_data = tasks.get("header", {})
    add_title_page(doc, term.upper(), role_title, {
        "Document Type": f"{term} ({get_framework_reference(framework, 'roleps')})",
        "Role Title": header_data.get("role_title", role_title),
        "TDW Number": header_data.get("tdw_number", "TBD"),
        "RolePS Serial": header_data.get("roleps_serial", "TBD"),
        "TRA": header_data.get("tra", "TBD"),
        "TDA": header_data.get("tda", "TBD"),
        "Issue/Version": header_data.get("issue_version", "1.0"),
        "Review Date": header_data.get("review_date", "TBD"),
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Classification": header_data.get("security_classification", "OFFICIAL")
    })
    
    add_section_heading(doc, "1. INTRODUCTION")
    doc.add_paragraph(f"This {term} defines the tasks, performance standards, and training requirements for the {role_title} role in accordance with {get_framework_reference(framework, 'roleps')}.")
    
    add_section_heading(doc, "2. TASK ANALYSIS BY DUTY")
    
    duties = tasks.get("duties", [])
    for duty in duties:
        duty_num = duty.get("duty_number", "")
        duty_title = duty.get("duty_title", "")
        
        doc.add_heading(f"DUTY {duty_num}: {duty_title}", level=2)
        doc.add_paragraph(duty.get("duty_description", ""))
        
        duty_tasks = duty.get("tasks", [])
        if duty_tasks:
            headers = ["Task No.", "Performance", "Conditions", "Standards", "Cat.", "Criticality"]
            rows = []
            
            for task in duty_tasks:
                rows.append([
                    task.get("task_number", ""),
                    task.get("performance", ""),
                    task.get("conditions", ""),
                    task.get("standards", ""),
                    task.get("category", ""),
                    task.get("criticality", "")
                ])
                
                for sub in task.get("sub_tasks", []):
                    rows.append([
                        f"  {sub.get('sub_task_number', '')}",
                        sub.get("performance", ""),
                        sub.get("conditions", ""),
                        sub.get("standards", ""),
                        sub.get("category", ""),
                        sub.get("ksa_type", "")
                    ])
            
            add_table_from_data(doc, headers, rows)
        
        if duty_tasks and "ksa" in duty_tasks[0]:
            doc.add_heading(f"Knowledge, Skills & Attitudes - Duty {duty_num}", level=3)
            for task in duty_tasks:
                ksa = task.get("ksa", {})
                if ksa:
                    doc.add_paragraph(f"Task {task.get('task_number', '')}:")
                    if ksa.get("knowledge"):
                        doc.add_paragraph(f"  Knowledge: {'; '.join(ksa.get('knowledge', []))}")
                    if ksa.get("skills"):
                        doc.add_paragraph(f"  Skills: {'; '.join(ksa.get('skills', []))}")
                    if ksa.get("attitudes"):
                        doc.add_paragraph(f"  Attitudes: {'; '.join(ksa.get('attitudes', []))}")
        
        doc.add_paragraph()
    
    add_section_heading(doc, "3. SUMMARY")
    summary = tasks.get("summary", {})
    
    doc.add_paragraph(f"Total Duties: {summary.get('total_duties', len(duties))}")
    doc.add_paragraph(f"Total Tasks: {summary.get('total_tasks', 0)}")
    doc.add_paragraph(f"Total Sub-Tasks: {summary.get('total_sub_tasks', 0)}")
    
    doc.add_heading("3.1 Training Category Breakdown", level=2)
    cat_breakdown = summary.get("category_breakdown", {})
    for cat, count in cat_breakdown.items():
        doc.add_paragraph(f"• {cat}: {count} tasks", style='List Bullet')
    
    doc.add_heading("3.2 Criticality Breakdown", level=2)
    crit_breakdown = summary.get("criticality_breakdown", {})
    for crit, count in crit_breakdown.items():
        doc.add_paragraph(f"• {crit.replace('_', ' ').title()}: {count} tasks", style='List Bullet')
    
    filename = "02_Role_Performance_Statement.docx"
    doc.save(output_path / filename)
    return filename


def build_gap_analysis_report(role_title: str, framework: str, content: Dict,
                               output_path: Path) -> str:
    """Build AMPLIFIED Training Gap Analysis document"""
    doc = create_styled_document("Training Gap Analysis", role_title, framework)
    
    add_title_page(doc, "TRAINING GAP ANALYSIS", role_title, {
        "Document Type": f"Training Gap Analysis ({get_framework_reference(framework, 'gap')})",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    add_section_heading(doc, "1. EXECUTIVE SUMMARY")
    exec_summary = content.get("executive_summary", {})
    
    table = doc.add_table(rows=6, cols=2)
    table.style = 'Table Grid'
    summary_data = [
        ("Total Gaps Identified", str(exec_summary.get("total_gaps", 0))),
        ("Critical Priority", str(exec_summary.get("critical_gaps", 0))),
        ("High Priority", str(exec_summary.get("high_priority_gaps", 0))),
        ("Medium Priority", str(exec_summary.get("medium_priority_gaps", 0))),
        ("Low Priority", str(exec_summary.get("low_priority_gaps", 0))),
        ("Timeline for Closure", exec_summary.get("timeline_for_closure", "TBD"))
    ]
    for i, (label, value) in enumerate(summary_data):
        table.rows[i].cells[0].text = label
        table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
        table.rows[i].cells[1].text = value
    doc.add_paragraph()
    
    doc.add_heading("1.1 Capability Risk Statement", level=2)
    doc.add_paragraph(exec_summary.get("capability_risk_statement", ""))
    
    doc.add_heading("1.2 Resource Implications", level=2)
    doc.add_paragraph(exec_summary.get("resource_implication_summary", ""))
    
    add_section_heading(doc, "2. DETAILED GAP ANALYSIS")
    
    gaps = content.get("gaps", [])
    for gap in gaps:
        gap_id = gap.get("gap_id", "")
        gap_title = gap.get("gap_title", "")
        
        doc.add_heading(f"{gap_id}: {gap_title}", level=2)
        
        details = [
            ("Task Reference", gap.get("task_reference", "")),
            ("Performance Requirement", gap.get("performance_requirement", "")),
            ("Gap Type", gap.get("gap_type", "")),
            ("Priority", gap.get("priority", "")),
            ("Criticality", gap.get("criticality_assessment", {}).get("rating", "") if isinstance(gap.get("criticality_assessment"), dict) else "")
        ]
        
        table = doc.add_table(rows=len(details), cols=2)
        table.style = 'Table Grid'
        for i, (label, value) in enumerate(details):
            table.rows[i].cells[0].text = label
            table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
            table.rows[i].cells[1].text = str(value)
        doc.add_paragraph()
        
        doc.add_heading("Current Capability", level=3)
        current = gap.get("current_capability", {})
        if isinstance(current, dict):
            doc.add_paragraph(current.get("description", ""))
            doc.add_paragraph(f"Source: {current.get('provision_source', '')}")
            doc.add_paragraph(f"Coverage: {current.get('coverage_percentage', '')}%")
        else:
            doc.add_paragraph(str(current))
        
        doc.add_heading("Required Capability", level=3)
        required = gap.get("required_capability", {})
        if isinstance(required, dict):
            doc.add_paragraph(required.get("description", ""))
            doc.add_paragraph(f"Standard Reference: {required.get('standard_reference', '')}")
        else:
            doc.add_paragraph(str(required))
        
        doc.add_heading("Gap Description", level=3)
        doc.add_paragraph(gap.get("gap_description", ""))
        
        doc.add_heading("Root Cause Analysis", level=3)
        doc.add_paragraph(gap.get("root_cause_analysis", ""))
        
        doc.add_heading("Operational Impact", level=3)
        doc.add_paragraph(gap.get("operational_impact", ""))
        
        dif = gap.get("dif_rating", {})
        if dif:
            doc.add_heading("DIF Rating", level=3)
            doc.add_paragraph(f"Difficulty: {dif.get('difficulty', '')}/5 | Importance: {dif.get('importance', '')}/5 | Frequency: {dif.get('frequency', '')}/5 | Overall Score: {dif.get('overall_score', '')}")
        
        risk = gap.get("risk_if_unaddressed", {})
        if risk:
            doc.add_heading("Risk if Unaddressed", level=3)
            doc.add_paragraph(f"Likelihood: {risk.get('likelihood', '')}/5 | Impact: {risk.get('impact', '')}/5 | Risk Score: {risk.get('risk_score', '')}")
            doc.add_paragraph(risk.get("risk_description", ""))
        
        intervention = gap.get("recommended_intervention", {})
        if intervention:
            doc.add_heading("Recommended Intervention", level=3)
            doc.add_paragraph(f"Type: {intervention.get('intervention_type', '')}")
            doc.add_paragraph(intervention.get("description", ""))
            doc.add_paragraph(f"Duration: {intervention.get('duration', '')} | Method: {intervention.get('method', '')}")
            doc.add_paragraph(f"Assessment: {intervention.get('assessment_approach', '')}")
        
        alternatives = gap.get("alternative_interventions", [])
        if alternatives:
            doc.add_heading("Alternative Interventions", level=3)
            for alt in alternatives:
                doc.add_paragraph(f"• {alt.get('option', '')}: Pros - {alt.get('pros', '')}; Cons - {alt.get('cons', '')}", style='List Bullet')
        
        res_est = gap.get("resource_estimate", {})
        if res_est:
            doc.add_heading("Resource Estimate", level=3)
            doc.add_paragraph(f"Development Cost: £{res_est.get('development_cost', 0):,}")
            doc.add_paragraph(f"Delivery Cost per Learner: £{res_est.get('delivery_cost_per_learner', 0):,}")
            doc.add_paragraph(f"Total Estimated Cost: £{res_est.get('total_estimated_cost', 0):,}")
        
        metrics = gap.get("success_metrics", [])
        if metrics:
            doc.add_heading("Success Metrics", level=3)
            for m in metrics:
                doc.add_paragraph(f"• {m}", style='List Bullet')
        
        doc.add_paragraph()
    
    add_section_heading(doc, "3. PRIORITISATION MATRIX")
    matrix = content.get("prioritisation_matrix", {})
    if matrix:
        doc.add_heading("3.1 Quick Wins (Low Effort, High Impact)", level=2)
        doc.add_paragraph(", ".join(matrix.get("quick_wins", ["None identified"])))
        
        doc.add_heading("3.2 Major Projects (High Effort, High Impact)", level=2)
        doc.add_paragraph(", ".join(matrix.get("major_projects", ["None identified"])))
        
        doc.add_heading("3.3 Fill-ins (Low Effort, Low Impact)", level=2)
        doc.add_paragraph(", ".join(matrix.get("fill_ins", ["None identified"])))
        
        doc.add_heading("3.4 Hard Slogs (High Effort, Low Impact)", level=2)
        doc.add_paragraph(", ".join(matrix.get("hard_slogs", ["None identified"])))
    
    add_section_heading(doc, "4. SUMMARY BY GAP TYPE")
    by_type = content.get("summary_by_type", {})
    if by_type:
        for gap_type, count in by_type.items():
            doc.add_paragraph(f"• {gap_type.replace('_', ' ').title()}: {count}", style='List Bullet')
    
    add_section_heading(doc, "5. RECOMMENDATIONS")
    recommendations = content.get("recommendations", [])
    for rec in recommendations:
        if isinstance(rec, dict):
            doc.add_paragraph(f"• {rec.get('recommendation', '')} (Priority: {rec.get('priority', '')}, Owner: {rec.get('owner', '')}, Timeline: {rec.get('timeline', '')})", style='List Bullet')
        else:
            doc.add_paragraph(f"• {rec}", style='List Bullet')
    
    filename = "03_Training_Gap_Analysis.docx"
    doc.save(output_path / filename)
    return filename


def build_training_needs_report(role_title: str, framework: str, content: Dict,
                                 output_path: Path) -> str:
<<<<<<< HEAD
    """Build Training Needs Report document (JSP 822 / DTSM 2 compliant structure)"""
=======
    """Build AMPLIFIED Training Needs Report document"""
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    doc = create_styled_document("Training Needs Report", role_title, framework)
    
    term = get_framework_term(framework, "tnr")
    
    add_title_page(doc, term.upper(), role_title, {
        "Document Type": f"{term} ({get_framework_reference(framework, 'tnr')})",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL",
        "Status": "For Approval"
    })
    
<<<<<<< HEAD
    # ========================================
    # SECTION 1: EXECUTIVE SUMMARY
    # ========================================
=======
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    add_section_heading(doc, "1. EXECUTIVE SUMMARY")
    exec_summary = content.get("executive_summary", {})
    
    doc.add_heading("1.1 Problem Statement", level=2)
    doc.add_paragraph(exec_summary.get("problem_statement", ""))
    
<<<<<<< HEAD
    doc.add_heading("1.2 Key Findings", level=2)
    for finding in exec_summary.get("key_findings", []):
        doc.add_paragraph(f"• {finding}", style='List Bullet')
    
    doc.add_heading("1.3 Principal Recommendation", level=2)
    doc.add_paragraph(exec_summary.get("principal_recommendation", ""))
    
    doc.add_heading("1.4 Resource Headline", level=2)
    doc.add_paragraph(exec_summary.get("resource_headline", ""))
    
    doc.add_heading("1.5 Recommended Decision", level=2)
=======
    doc.add_heading("1.2 Methodology", level=2)
    doc.add_paragraph(exec_summary.get("methodology_summary", ""))
    
    doc.add_heading("1.3 Key Findings", level=2)
    for finding in exec_summary.get("key_findings", []):
        doc.add_paragraph(f"• {finding}", style='List Bullet')
    
    doc.add_heading("1.4 Principal Recommendation", level=2)
    doc.add_paragraph(exec_summary.get("principal_recommendation", ""))
    
    doc.add_heading("1.5 Resource Headline", level=2)
    doc.add_paragraph(exec_summary.get("resource_headline", ""))
    
    doc.add_heading("1.6 Key Risks if Not Approved", level=2)
    for risk in exec_summary.get("top_risks", []):
        if isinstance(risk, dict):
            doc.add_paragraph(f"• {risk.get('risk', '')}: {risk.get('impact', '')}", style='List Bullet')
        else:
            doc.add_paragraph(f"• {risk}", style='List Bullet')
    
    doc.add_heading("1.7 Recommended Decision", level=2)
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    rec_para = doc.add_paragraph(exec_summary.get("recommended_decision", ""))
    if rec_para.runs:
        rec_para.runs[0].bold = True
    
<<<<<<< HEAD
    # ========================================
    # SECTION 2: INTRODUCTION
    # ========================================
    add_section_heading(doc, "2. INTRODUCTION")
    intro = content.get("introduction", {})
    if isinstance(intro, str):
        doc.add_paragraph(intro)
    else:
        doc.add_paragraph(intro.get("overview", ""))
    
    # ========================================
    # SECTION 3: BACKGROUND
    # ========================================
    add_section_heading(doc, "3. BACKGROUND")
    doc.add_paragraph(content.get("background", ""))
    
    # ========================================
    # SECTION 4: AIM / PURPOSE
    # ========================================
    add_section_heading(doc, "4. AIM / PURPOSE")
    aim = content.get("aim_purpose", content.get("aim", ""))
    if isinstance(aim, str):
        doc.add_paragraph(aim)
    else:
        doc.add_paragraph(aim.get("aim", ""))
        if aim.get("objectives"):
            doc.add_heading("4.1 Objectives", level=2)
            for obj in aim.get("objectives", []):
                doc.add_paragraph(f"• {obj}", style='List Bullet')
    
    # ========================================
    # SECTION 5: TERMS OF REFERENCE
    # ========================================
    add_section_heading(doc, "5. TERMS OF REFERENCE")
    tor = content.get("terms_of_reference", {})
    if isinstance(tor, str):
        doc.add_paragraph(tor)
    else:
        doc.add_paragraph(tor.get("overview", ""))
        if tor.get("deliverables"):
            doc.add_heading("5.1 Deliverables", level=2)
            for d in tor.get("deliverables", []):
                doc.add_paragraph(f"• {d}", style='List Bullet')
    
    # ========================================
    # SECTION 6: SCOPE
    # ========================================
    add_section_heading(doc, "6. SCOPE")
    scope = content.get("scope", {})
    if isinstance(scope, str):
        doc.add_paragraph(scope)
    else:
        doc.add_heading("6.1 In Scope", level=2)
        for item in scope.get("in_scope", []):
            doc.add_paragraph(f"• {item}", style='List Bullet')
        doc.add_heading("6.2 Out of Scope", level=2)
        for item in scope.get("out_of_scope", []):
            doc.add_paragraph(f"• {item}", style='List Bullet')
    
    # ========================================
    # SECTION 7: STATEMENT OF REQUIREMENT
    # ========================================
    add_section_heading(doc, "7. STATEMENT OF REQUIREMENT")
    sor = content.get("statement_of_requirement", "")
    if isinstance(sor, str):
        doc.add_paragraph(sor)
    else:
        doc.add_paragraph(sor.get("overview", ""))
    
    # ========================================
    # SECTION 8: TRAINING NEEDS ANALYSIS SUPPORT GROUP (TNASG)
    # ========================================
    add_section_heading(doc, "8. TRAINING NEEDS ANALYSIS SUPPORT GROUP (TNASG)")
    tnasg = content.get("tnasg", content.get("governance", {}))
    if isinstance(tnasg, str):
        doc.add_paragraph(tnasg)
    else:
        if tnasg.get("members"):
            headers = ["Name", "Role", "Organisation"]
            rows = [[m.get("name", ""), m.get("role", ""), m.get("organisation", "")] for m in tnasg.get("members", [])]
            add_table_from_data(doc, headers, rows)
    
    # ========================================
    # SECTION 9: STAKEHOLDER ENGAGEMENT
    # ========================================
    add_section_heading(doc, "9. STAKEHOLDER ENGAGEMENT")
    stakeholders = content.get("stakeholder_engagement", content.get("stakeholders", {}))
    if isinstance(stakeholders, list):
        headers = ["Stakeholder", "Interest", "Engagement Method"]
        rows = [[s.get("name", ""), s.get("interest", ""), s.get("engagement", "")] for s in stakeholders]
        add_table_from_data(doc, headers, rows)
    elif isinstance(stakeholders, dict) and stakeholders.get("stakeholders"):
        headers = ["Stakeholder", "Interest", "Engagement Method"]
        rows = [[s.get("name", ""), s.get("interest", ""), s.get("engagement", "")] for s in stakeholders.get("stakeholders", [])]
        add_table_from_data(doc, headers, rows)
    
    # ========================================
    # SECTION 10: RISKS
    # ========================================
    add_section_heading(doc, "10. RISKS")
    risks = content.get("risks", content.get("risk_assessment", []))
    if risks:
        headers = ["Risk", "Likelihood", "Impact", "Mitigation", "Owner"]
        rows = [[
            r.get("risk", r.get("description", "")),
            r.get("likelihood", ""),
            r.get("impact", ""),
            r.get("mitigation", ""),
            r.get("owner", "")
        ] for r in risks]
        add_table_from_data(doc, headers, rows)
    
    # ========================================
    # SECTION 11: ASSUMPTIONS
    # ========================================
    add_section_heading(doc, "11. ASSUMPTIONS")
    assumptions = content.get("assumptions", [])
    if assumptions:
        for a in assumptions:
            if isinstance(a, dict):
                doc.add_paragraph(f"• {a.get('assumption', a.get('description', ''))}", style='List Bullet')
            else:
                doc.add_paragraph(f"• {a}", style='List Bullet')
    
    # ========================================
    # SECTION 12: DEPENDENCIES
    # ========================================
    add_section_heading(doc, "12. DEPENDENCIES")
    dependencies = content.get("dependencies", [])
    if dependencies:
        for d in dependencies:
            if isinstance(d, dict):
                doc.add_paragraph(f"• {d.get('dependency', d.get('description', ''))}", style='List Bullet')
            else:
                doc.add_paragraph(f"• {d}", style='List Bullet')
    
    # ========================================
    # SECTION 13: CONSTRAINTS
    # ========================================
    add_section_heading(doc, "13. CONSTRAINTS")
    constraints = content.get("constraints", [])
    if constraints:
        for c in constraints:
            if isinstance(c, dict):
                doc.add_paragraph(f"• {c.get('constraint', c.get('description', ''))}", style='List Bullet')
            else:
                doc.add_paragraph(f"• {c}", style='List Bullet')
    
    # ========================================
    # SECTION 14: RISK REGISTER (RAIDO)
    # ========================================
    add_section_heading(doc, "14. RISK REGISTER (RAIDO)")
    raido = content.get("raido", content.get("risk_register", {}))
    doc.add_paragraph("Full RAIDO register attached at Annex B.")
    
    # ========================================
    # SECTION 15: METHODOLOGY
    # ========================================
    add_section_heading(doc, "15. METHODOLOGY")
    methodology = content.get("methodology", exec_summary.get("methodology_summary", ""))
    if isinstance(methodology, str):
        doc.add_paragraph(methodology)
    else:
        doc.add_paragraph(methodology.get("overview", ""))
        if methodology.get("activities"):
            doc.add_heading("15.1 TNA Activities Conducted", level=2)
            for act in methodology.get("activities", []):
                doc.add_paragraph(f"• {act}", style='List Bullet')
    
    # ========================================
    # SECTION 16: PREVIOUS TRAINING STUDIES
    # ========================================
    add_section_heading(doc, "16. PREVIOUS TRAINING STUDIES")
    previous_studies = content.get("previous_training_studies", [])
    if previous_studies:
        for study in previous_studies:
            if isinstance(study, dict):
                doc.add_paragraph(f"• {study.get('title', '')}: {study.get('summary', '')}", style='List Bullet')
            else:
                doc.add_paragraph(f"• {study}", style='List Bullet')
    else:
        doc.add_paragraph("No previous training studies identified.")
    
    # ========================================
    # SECTION 17: STATEMENT OF TRAINED REQUIREMENT (SOTR)
    # ========================================
    add_section_heading(doc, "17. STATEMENT OF TRAINED REQUIREMENT (SOTR)")
    sotr = content.get("sotr", content.get("statement_of_trained_requirement", {}))
    if isinstance(sotr, str):
        doc.add_paragraph(sotr)
    else:
        doc.add_paragraph(sotr.get("overview", "Full SOTR attached at Annex C."))
    
    # ========================================
    # SECTION 18: STATEMENT OF TRAINED TASKS (SOTT)
    # ========================================
    add_section_heading(doc, "18. STATEMENT OF TRAINED TASKS (SOTT)")
    sott = content.get("sott", content.get("statement_of_trained_tasks", {}))
    if isinstance(sott, str):
        doc.add_paragraph(sott)
    else:
        doc.add_paragraph(sott.get("overview", "Full SOTT attached at Annex D."))
    
    # ========================================
    # SECTION 19: TRAINING GAP ANALYSIS
    # ========================================
    add_section_heading(doc, "19. TRAINING GAP ANALYSIS")
    findings = content.get("analysis_findings", {})
    gap_summary = findings.get("gap_summary", content.get("gap_summary", {}))
    if gap_summary:
        doc.add_paragraph(f"Total Gaps Identified: {gap_summary.get('total_gaps', 0)}")
        by_priority = gap_summary.get("by_priority", {})
        doc.add_paragraph(f"By Priority - Critical: {by_priority.get('critical', 0)}, High: {by_priority.get('high', 0)}, Medium: {by_priority.get('medium', 0)}, Low: {by_priority.get('low', 0)}")
    doc.add_paragraph("Full Training Gap Analysis attached at Annex E.")
    
    # ========================================
    # SECTION 20: STATEMENT OF TRAINING GAPS
    # ========================================
    add_section_heading(doc, "20. STATEMENT OF TRAINING GAPS")
    training_gaps = content.get("statement_of_training_gaps", content.get("training_gaps", []))
    if training_gaps:
        headers = ["Gap ID", "Description", "Priority", "Training Solution"]
        rows = [[
            g.get("gap_id", ""),
            g.get("description", ""),
            g.get("priority", ""),
            g.get("solution", "")
        ] for g in training_gaps[:10]]  # Show first 10
        add_table_from_data(doc, headers, rows)
        if len(training_gaps) > 10:
            doc.add_paragraph(f"Note: {len(training_gaps) - 10} additional gaps detailed in Annex E.")
    
    # ========================================
    # SECTION 21: FIDELITY ANALYSIS
    # ========================================
    add_section_heading(doc, "21. FIDELITY ANALYSIS")
    fidelity = content.get("fidelity_analysis", {})
    if isinstance(fidelity, str):
        doc.add_paragraph(fidelity)
    else:
        doc.add_paragraph(fidelity.get("overview", "Fidelity analysis determines the degree to which training must replicate operational conditions."))
    
    # ========================================
    # SECTION 22: DIF ANALYSIS
    # ========================================
    add_section_heading(doc, "22. DIF ANALYSIS")
    dif = content.get("dif_analysis", {})
    doc.add_paragraph("Difficulty-Importance-Frequency analysis conducted to prioritise training interventions.")
    doc.add_paragraph("Full DIF Analysis attached at Annex F.")
    
    # ========================================
    # SECTION 23: KSA ANALYSIS
    # ========================================
    add_section_heading(doc, "23. KSA ANALYSIS")
    ksa = findings.get("ksa_analysis", content.get("ksa_analysis", {}))
    
    doc.add_heading("23.1 Knowledge Requirements", level=2)
=======
    add_section_heading(doc, "2. BACKGROUND")
    doc.add_paragraph(content.get("background", ""))
    
    add_section_heading(doc, "3. ANALYSIS FINDINGS")
    findings = content.get("analysis_findings", {})
    
    doc.add_heading("3.1 Role Analysis Summary", level=2)
    role_analysis = findings.get("role_analysis_summary", {})
    if role_analysis:
        doc.add_paragraph(f"Complexity Assessment: {role_analysis.get('complexity_assessment', '')}")
        doc.add_paragraph(f"Justification: {role_analysis.get('complexity_justification', '')}")
        doc.add_paragraph(f"Comparison to Similar Roles: {role_analysis.get('comparison_to_similar_roles', '')}")
        doc.add_paragraph(f"Career Pathway Context: {role_analysis.get('career_pathway_context', '')}")
        doc.add_paragraph(f"Competency Framework Alignment: {role_analysis.get('competency_framework_alignment', '')}")
    
    doc.add_heading("3.2 Task Analysis Summary", level=2)
    task_summary = findings.get("task_analysis_summary", {})
    if task_summary:
        table = doc.add_table(rows=7, cols=2)
        table.style = 'Table Grid'
        task_data = [
            ("Total Tasks", str(task_summary.get("total_tasks", 0))),
            ("Formal Training (FT)", str(task_summary.get("ft_tasks", 0))),
            ("Workplace Training (WPT)", str(task_summary.get("wpt_tasks", 0))),
            ("On-the-Job (OJT)", str(task_summary.get("ojt_tasks", 0))),
            ("Computer-Based (CBT)", str(task_summary.get("cbt_tasks", 0))),
            ("Safety-Critical Tasks", str(task_summary.get("safety_critical_tasks", 0))),
            ("Mission-Critical Tasks", str(task_summary.get("mission_critical_tasks", 0)))
        ]
        for i, (label, value) in enumerate(task_data):
            table.rows[i].cells[0].text = label
            table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
            table.rows[i].cells[1].text = value
        doc.add_paragraph()
    
    doc.add_heading("3.3 Knowledge, Skills and Attitudes Analysis", level=2)
    ksa = findings.get("ksa_analysis", {})
    
    doc.add_heading("3.3.1 Knowledge Requirements", level=3)
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    knowledge = ksa.get("knowledge_requirements", [])
    if knowledge:
        headers = ["Knowledge Area", "Level", "Priority"]
        rows = [[k.get("area", ""), k.get("level", ""), k.get("priority", "")] for k in knowledge]
        add_table_from_data(doc, headers, rows)
    
<<<<<<< HEAD
    doc.add_heading("23.2 Skill Requirements", level=2)
=======
    doc.add_heading("3.3.2 Skill Requirements", level=3)
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    skills = ksa.get("skill_requirements", [])
    if skills:
        headers = ["Skill", "Type", "Proficiency", "Priority"]
        rows = [[s.get("skill", ""), s.get("type", ""), s.get("proficiency", ""), s.get("priority", "")] for s in skills]
        add_table_from_data(doc, headers, rows)
    
<<<<<<< HEAD
    doc.add_heading("23.3 Attitude/Behaviour Requirements", level=2)
=======
    doc.add_heading("3.3.3 Attitude/Behaviour Requirements", level=3)
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    attitudes = ksa.get("attitude_requirements", [])
    if attitudes:
        headers = ["Attitude", "Importance", "Development Approach"]
        rows = [[a.get("attitude", ""), a.get("importance", ""), a.get("development_approach", "")] for a in attitudes]
        add_table_from_data(doc, headers, rows)
    
<<<<<<< HEAD
    doc.add_paragraph("Full KSA Analysis attached at Annex G.")
    
    # ========================================
    # SECTION 24: METHODS & MEDIA ANALYSIS
    # ========================================
    add_section_heading(doc, "24. METHODS & MEDIA ANALYSIS")
    methods_media = content.get("methods_media_analysis", {})
    if isinstance(methods_media, str):
        doc.add_paragraph(methods_media)
    else:
        doc.add_paragraph(methods_media.get("overview", "Analysis of appropriate training methods and media for identified training requirements."))
    
    # ========================================
    # SECTION 25: TRAINING OPTIONS ANALYSIS
    # ========================================
    add_section_heading(doc, "25. TRAINING OPTIONS ANALYSIS")
=======
    doc.add_heading("3.4 Gap Summary", level=2)
    gap_summary = findings.get("gap_summary", {})
    if gap_summary:
        doc.add_paragraph(f"Total Gaps: {gap_summary.get('total_gaps', 0)}")
        by_priority = gap_summary.get("by_priority", {})
        doc.add_paragraph(f"By Priority - Critical: {by_priority.get('critical', 0)}, High: {by_priority.get('high', 0)}, Medium: {by_priority.get('medium', 0)}, Low: {by_priority.get('low', 0)}")
        doc.add_paragraph(f"Estimated Total Cost to Close: £{gap_summary.get('estimated_total_cost_to_close', 0):,}")
    
    add_section_heading(doc, "4. TRAINING OPTIONS ANALYSIS")
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    
    options = content.get("training_options", [])
    for opt in options:
        doc.add_heading(f"Option {opt.get('option_id', '')}: {opt.get('option_name', '')}", level=2)
        
        doc.add_heading("Description", level=3)
        doc.add_paragraph(opt.get("description", ""))
        
        methodology = opt.get("delivery_methodology", {})
        if methodology:
            doc.add_heading("Delivery Methodology", level=3)
            doc.add_paragraph(f"Primary Method: {methodology.get('primary_method', '')}")
            doc.add_paragraph(f"Secondary Methods: {', '.join(methodology.get('secondary_methods', []))}")
<<<<<<< HEAD
        
        swot = opt.get("swot_analysis", {})
        if swot:
            doc.add_heading("SWOT Analysis", level=3)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.rows[0].cells[0].text = "STRENGTHS\n" + "\n".join([f"• {s}" for s in swot.get("strengths", [])])
            table.rows[0].cells[1].text = "WEAKNESSES\n" + "\n".join([f"• {w}" for w in swot.get("weaknesses", [])])
            table.rows[1].cells[0].text = "OPPORTUNITIES\n" + "\n".join([f"• {o}" for o in swot.get("opportunities", [])])
            table.rows[1].cells[1].text = "THREATS\n" + "\n".join([f"• {t}" for t in swot.get("threats", [])])
            doc.add_paragraph()
    
    doc.add_paragraph("Full Training Options Analysis attached at Annex H.")
    
    # ========================================
    # SECTION 26: COST BENEFIT ANALYSIS (CBA)
    # ========================================
    add_section_heading(doc, "26. COST BENEFIT ANALYSIS (CBA)")
    
    for opt in options:
        cost_benefit = opt.get("cost_benefit_analysis", {})
        if cost_benefit:
            doc.add_heading(f"Option {opt.get('option_id', '')}: {opt.get('option_name', '')}", level=2)
=======
            doc.add_paragraph(f"Technology Requirements: {', '.join(methodology.get('technology_requirements', []))}")
            doc.add_paragraph(f"Assessment Approach: {methodology.get('assessment_approach', '')}")
        
        structure = opt.get("programme_structure", [])
        if structure:
            doc.add_heading("Programme Structure", level=3)
            headers = ["Module", "Duration", "Topics", "Delivery"]
            rows = [[
                m.get("module", ""),
                m.get("duration", ""),
                "; ".join(m.get("topics", [])),
                m.get("delivery", "")
            ] for m in structure]
            add_table_from_data(doc, headers, rows)
        
        coverage = opt.get("coverage_analysis", {})
        if coverage:
            doc.add_heading("Coverage Analysis", level=3)
            doc.add_paragraph(f"Coverage: {coverage.get('coverage_percentage', 0)}%")
            doc.add_paragraph(f"Fully Addressed: {', '.join(coverage.get('gaps_fully_addressed', []))}")
            doc.add_paragraph(f"Partially Addressed: {', '.join(coverage.get('gaps_partially_addressed', []))}")
            doc.add_paragraph(f"Not Addressed: {', '.join(coverage.get('gaps_not_addressed', []))}")
        
        cost_benefit = opt.get("cost_benefit_analysis", {})
        if cost_benefit:
            doc.add_heading("5-Year Cost Analysis", level=3)
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
            five_year = cost_benefit.get("five_year_costs", {})
            if five_year:
                headers = ["Year", "Development", "Delivery", "Infrastructure", "Personnel", "Total"]
                rows = []
                for year_num in range(1, 6):
                    year_key = f"year_{year_num}"
                    year_data = five_year.get(year_key, {})
                    rows.append([
                        f"Year {year_num}",
                        f"£{year_data.get('development', 0):,}",
                        f"£{year_data.get('delivery', 0):,}",
                        f"£{year_data.get('infrastructure', 0):,}",
                        f"£{year_data.get('personnel', 0):,}",
                        f"£{year_data.get('total', 0):,}"
                    ])
                rows.append(["5-Year Total", "", "", "", "", f"£{five_year.get('total_5_year', 0):,}"])
                add_table_from_data(doc, headers, rows)
            
            roi = cost_benefit.get("roi_calculation", {})
            if roi:
                doc.add_heading("Return on Investment", level=3)
                doc.add_paragraph(f"Total Investment: £{roi.get('total_investment', 0):,}")
                doc.add_paragraph(f"Annual Benefit: £{roi.get('annual_benefit', 0):,}")
                doc.add_paragraph(f"Payback Period: {roi.get('payback_period_years', 0)} years")
                doc.add_paragraph(f"5-Year ROI: {roi.get('five_year_roi_percentage', 0)}%")
<<<<<<< HEAD
    
    doc.add_paragraph("Full Cost Benefit Analysis attached at Annex I.")
    
    # ========================================
    # SECTION 27: TRAINING PLAN
    # ========================================
    add_section_heading(doc, "27. TRAINING PLAN")
    recommended = content.get("recommended_solution", {})
    
    doc.add_paragraph(f"Selected Option: {recommended.get('selected_option', '')}")
    doc.add_heading("27.1 Selection Rationale", level=2)
=======
        
        swot = opt.get("swot_analysis", {})
        if swot:
            doc.add_heading("SWOT Analysis", level=3)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.rows[0].cells[0].text = "STRENGTHS\n" + "\n".join([f"• {s}" for s in swot.get("strengths", [])])
            table.rows[0].cells[1].text = "WEAKNESSES\n" + "\n".join([f"• {w}" for w in swot.get("weaknesses", [])])
            table.rows[1].cells[0].text = "OPPORTUNITIES\n" + "\n".join([f"• {o}" for o in swot.get("opportunities", [])])
            table.rows[1].cells[1].text = "THREATS\n" + "\n".join([f"• {t}" for t in swot.get("threats", [])])
            doc.add_paragraph()
        
        doc.add_heading("Ratings", level=3)
        eff = opt.get("effectiveness_rating", {})
        doc.add_paragraph(f"Effectiveness: {eff.get('rating', '')} - {eff.get('justification', '')}")
        eff2 = opt.get("efficiency_rating", {})
        doc.add_paragraph(f"Efficiency: {eff2.get('rating', '')} - {eff2.get('justification', '')}")
        feas = opt.get("feasibility_rating", {})
        doc.add_paragraph(f"Feasibility: {feas.get('rating', '')} - {feas.get('justification', '')}")
        
        doc.add_page_break()
    
    add_section_heading(doc, "5. RECOMMENDED SOLUTION")
    recommended = content.get("recommended_solution", {})
    
    doc.add_paragraph(f"Selected Option: {recommended.get('selected_option', '')}")
    doc.add_heading("5.1 Selection Rationale", level=2)
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    doc.add_paragraph(recommended.get("selection_rationale", ""))
    
    tsp = recommended.get("training_statement_preview", {})
    if tsp:
<<<<<<< HEAD
        doc.add_heading("27.2 Training Statement Preview", level=2)
        doc.add_paragraph(f"Training Performance Statement (TPS): {tsp.get('tps_summary', '')}")
        doc.add_paragraph(f"Workplace Training Statement (WTS): {tsp.get('wps_summary', '')}")
=======
        doc.add_heading("5.2 Training Statement Preview", level=2)
        doc.add_paragraph(f"Training Performance Statement (TPS): {tsp.get('tps_summary', '')}")
        doc.add_paragraph(f"Workplace Training Statement (WPS): {tsp.get('wps_summary', '')}")
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
        doc.add_paragraph(f"Residual Training Gap Statement (RTGS): {tsp.get('rtgs_summary', '')}")
    
    impl = recommended.get("implementation_approach", [])
    if impl:
<<<<<<< HEAD
        doc.add_heading("27.3 Implementation Timeline", level=2)
=======
        doc.add_heading("5.3 Implementation Timeline", level=2)
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
        headers = ["Phase", "Duration", "Activities", "Milestone"]
        rows = [[
            p.get("phase", ""),
            p.get("duration", ""),
            "; ".join(p.get("activities", [])),
            p.get("milestone", "")
        ] for p in impl]
        add_table_from_data(doc, headers, rows)
    
<<<<<<< HEAD
    # ========================================
    # SECTION 28: THROUGH LIFE MANAGEMENT PLAN (TLMP)
    # ========================================
    add_section_heading(doc, "28. THROUGH LIFE MANAGEMENT PLAN (TLMP)")
    tlmp = content.get("tlmp", content.get("through_life_management", {}))
    if isinstance(tlmp, str):
        doc.add_paragraph(tlmp)
    else:
        doc.add_paragraph(tlmp.get("overview", "Through Life Management Plan to be developed during Design phase."))
    
    # ========================================
    # SECTION 29: CONCLUSIONS
    # ========================================
    add_section_heading(doc, "29. CONCLUSIONS")
    conclusions = content.get("conclusions", [])
    if isinstance(conclusions, str):
        doc.add_paragraph(conclusions)
    elif conclusions:
        for c in conclusions:
            if isinstance(c, dict):
                doc.add_paragraph(f"• {c.get('conclusion', c.get('finding', ''))}", style='List Bullet')
            else:
                doc.add_paragraph(f"• {c}", style='List Bullet')
    
    # ========================================
    # SECTION 30: RECOMMENDATIONS
    # ========================================
    add_section_heading(doc, "30. RECOMMENDATIONS")
=======
    success = recommended.get("success_criteria", [])
    if success:
        doc.add_heading("5.4 Success Criteria", level=2)
        headers = ["Criterion", "Measure", "Target"]
        rows = [[s.get("criterion", ""), s.get("measure", ""), s.get("target", "")] for s in success]
        add_table_from_data(doc, headers, rows)
    
    add_section_heading(doc, "6. RESOURCE IMPLICATIONS")
    resources = content.get("resource_implications", {})
    if resources:
        doc.add_paragraph(f"Total Development Cost: £{resources.get('total_development_cost', 0):,}")
        doc.add_paragraph(f"Annual Delivery Cost: £{resources.get('annual_delivery_cost', 0):,}")
        doc.add_paragraph(f"5-Year Total Cost: £{resources.get('five_year_total_cost', 0):,}")
        doc.add_paragraph(f"Funding Source: {resources.get('funding_source', '')}")
        doc.add_paragraph(f"Budget Line: {resources.get('budget_line', '')}")
    
    add_section_heading(doc, "7. RISK ASSESSMENT")
    risks = content.get("risk_assessment", [])
    if risks:
        headers = ["Risk", "L", "I", "Mitigation/Consequence", "Owner"]
        rows = [[
            r.get("risk", ""),
            r.get("likelihood", ""),
            r.get("impact", ""),
            r.get("mitigation", r.get("consequence", "")),
            r.get("owner", "")
        ] for r in risks]
        add_table_from_data(doc, headers, rows)
    
    add_section_heading(doc, "8. RECOMMENDATIONS")
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    recommendations = content.get("recommendations", [])
    for rec in recommendations:
        if isinstance(rec, dict):
            doc.add_paragraph(f"{rec.get('number', '')}. {rec.get('recommendation', '')}")
            doc.add_paragraph(f"   Rationale: {rec.get('rationale', '')}")
            doc.add_paragraph(f"   Owner: {rec.get('owner', '')} | Timeline: {rec.get('timeline', '')}")
        else:
            doc.add_paragraph(f"• {rec}", style='List Bullet')
    
<<<<<<< HEAD
    # ========================================
    # REFERENCES
    # ========================================
    add_section_heading(doc, "REFERENCES")
    references = content.get("references", [])
    if references:
        for ref in references:
            if isinstance(ref, dict):
                doc.add_paragraph(f"• {ref.get('reference', ref.get('title', ''))}", style='List Bullet')
            else:
                doc.add_paragraph(f"• {ref}", style='List Bullet')
    else:
        doc.add_paragraph("• JSP 822 V7.0 - Defence Individual Training Policy", style='List Bullet')
        doc.add_paragraph("• DTSM 2 (2024 Edition) - Analysis of Individual Training", style='List Bullet')
    
    # ========================================
    # ANNEXES
    # ========================================
    doc.add_page_break()
    add_section_heading(doc, "ANNEXES")
    doc.add_paragraph("Annex A: Stakeholder Engagement Record")
    doc.add_paragraph("Annex B: Risk Register (RAIDO)")
    doc.add_paragraph("Annex C: Statement of Trained Requirement (SOTR)")
    doc.add_paragraph("Annex D: Statement of Trained Tasks (SOTT)")
    doc.add_paragraph("Annex E: Training Gap Analysis")
    doc.add_paragraph("Annex F: DIF Analysis")
    doc.add_paragraph("Annex G: KSA Analysis")
    doc.add_paragraph("Annex H: Training Options Analysis")
    doc.add_paragraph("Annex I: Cost Benefit Analysis")
=======
    add_section_heading(doc, "9. NEXT STEPS")
    next_steps = content.get("next_steps", [])
    for step in next_steps:
        if isinstance(step, dict):
            doc.add_paragraph(f"{step.get('step', '')}. {step.get('action', '')} - Owner: {step.get('owner', '')} - Deadline: {step.get('deadline', '')}")
        else:
            doc.add_paragraph(f"• {step}", style='List Bullet')
    
    add_section_heading(doc, "10. APPROVAL")
    approval = content.get("approval_requirements", {})
    if approval:
        doc.add_paragraph(f"Approving Authority: {approval.get('approving_authority', '')}")
        doc.add_paragraph(f"Approval Required By: {approval.get('approval_date_required', '')}")
        doc.add_paragraph(f"Conditions: {approval.get('conditions', '')}")
>>>>>>> 7fa108fcb46c2e1a5b1e1c203e00f7526a489655
    
    filename = "04_Training_Needs_Report.docx"
    doc.save(output_path / filename)
    return filename


# ============================================================================
# AGENT RUNNERS
# ============================================================================

def update_progress(job_id: str, progress: int, step: str):
    """Update job progress"""
    if job_id in jobs:
        jobs[job_id]["progress"] = progress
        jobs[job_id]["current_step"] = step
        jobs[job_id]["steps_completed"].append(step)
        print(f"[NOVA] Job {job_id}: {progress}% - {step}")


async def run_agent(job_id: str, agent: str, parameters: Dict[str, Any]):
    """Run the specified agent"""
    print(f"[NOVA] Starting agent execution: {agent} for job {job_id}")
    
    try:
        jobs[job_id]["status"] = "running"
        
        if agent == "analysis":
            await run_analysis_agent(job_id, parameters)
        elif agent == "design":
            await run_design_agent(job_id, parameters)
        elif agent == "delivery":
            await run_delivery_agent(job_id, parameters)
        elif agent == "evaluation":
            await run_evaluation_agent(job_id, parameters)
        else:
            raise ValueError(f"Unknown agent: {agent}")
        
        jobs[job_id]["status"] = "completed"
        jobs[job_id]["completed_at"] = datetime.utcnow().isoformat()
        print(f"[NOVA] Job {job_id} completed successfully")
        
    except Exception as e:
        print(f"[NOVA] Error in job {job_id}: {str(e)}")
        import traceback
        traceback.print_exc()
        jobs[job_id]["status"] = "failed"
        jobs[job_id]["error"] = str(e)


async def run_analysis_agent(job_id: str, parameters: Dict[str, Any]):
    """Analysis Agent - AMPLIFIED OUTPUTS"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "Commercial")
    description = parameters.get("description", "")
    
    output_dir = Path(jobs[job_id]["output_dir"])
    analysis_dir = output_dir / "01_Analysis"
    analysis_dir.mkdir(exist_ok=True)
    
    files_generated = []
    
    update_progress(job_id, 10, "Generating Amplified Scoping Exercise Report")
    scoping_content = await generate_scoping_content(role_title, framework, description)
    
    update_progress(job_id, 20, "Building Scoping Report Document")
    filename = build_scoping_report(role_title, framework, scoping_content, analysis_dir)
    files_generated.append(filename)
    
    update_progress(job_id, 30, "Generating Amplified Role Performance Statement")
    tasks = await generate_role_tasks(role_title, framework, description)
    
    update_progress(job_id, 45, "Building Role Performance Statement Document")
    filename = build_role_performance_statement(role_title, framework, tasks, analysis_dir)
    files_generated.append(filename)
    
    update_progress(job_id, 55, "Generating Amplified Training Gap Analysis")
    gaps = await generate_gap_analysis(role_title, framework, tasks)
    
    update_progress(job_id, 70, "Building Gap Analysis Document")
    filename = build_gap_analysis_report(role_title, framework, gaps, analysis_dir)
    files_generated.append(filename)
    
    update_progress(job_id, 80, "Generating Amplified Training Needs Report")
    tnr_content = await generate_tnr_content(role_title, framework, tasks, gaps)
    
    update_progress(job_id, 90, "Building Training Needs Report Document")
    filename = build_training_needs_report(role_title, framework, tnr_content, analysis_dir)
    files_generated.append(filename)
    
    update_progress(job_id, 100, "Analysis Phase Complete")
    
    return files_generated


async def run_design_agent(job_id: str, parameters: Dict[str, Any]):
    """Design Agent - placeholder for future amplification"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "Commercial")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    update_progress(job_id, 50, "Design Agent - Implementation Pending")
    
    design_dir = output_dir / "02_Design"
    design_dir.mkdir(exist_ok=True)
    
    doc = create_styled_document("Learning Specification", role_title, framework)
    add_title_page(doc, "LEARNING SPECIFICATION", role_title, {
        "Document Type": "Learning Specification (DTSM 3)",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Status": "Pending Full Implementation"
    })
    doc.add_paragraph("Design Agent with Amplified Outputs - Implementation in progress.")
    
    doc.save(design_dir / "01_Learning_Specification.docx")
    
    update_progress(job_id, 100, "Design Phase Complete")


async def run_delivery_agent(job_id: str, parameters: Dict[str, Any]):
    """Delivery Agent - placeholder for future amplification"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "Commercial")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    update_progress(job_id, 50, "Delivery Agent - Implementation Pending")
    
    delivery_dir = output_dir / "03_Delivery"
    delivery_dir.mkdir(exist_ok=True)
    
    doc = create_styled_document("Lesson Plans", role_title, framework)
    add_title_page(doc, "LESSON PLANS", role_title, {
        "Document Type": "Lesson Plans (DTSM 4)",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Status": "Pending Full Implementation"
    })
    doc.add_paragraph("Delivery Agent with Amplified Outputs - Implementation in progress.")
    
    doc.save(delivery_dir / "01_Lesson_Plans.docx")
    
    update_progress(job_id, 100, "Delivery Phase Complete")


async def run_evaluation_agent(job_id: str, parameters: Dict[str, Any]):
    """Evaluation Agent - placeholder for future amplification"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "Commercial")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    update_progress(job_id, 50, "Evaluation Agent - Implementation Pending")
    
    eval_dir = output_dir / "04_Evaluation"
    eval_dir.mkdir(exist_ok=True)
    
    doc = create_styled_document("Evaluation Strategy", role_title, framework)
    add_title_page(doc, "EVALUATION STRATEGY", role_title, {
        "Document Type": "Evaluation Strategy (DTSM 5)",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Status": "Pending Full Implementation"
    })
    doc.add_paragraph("Evaluation Agent with Amplified Outputs - Implementation in progress.")
    
    doc.save(eval_dir / "01_Evaluation_Strategy.docx")
    
    update_progress(job_id, 100, "Evaluation Phase Complete")


# ============================================================================
# RUN SERVER
# ============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
