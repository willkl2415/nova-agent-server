"""
NOVA Agent Server v2.0
FastAPI server for executing autonomous training agents with Claude AI
Generates professional .docx and .xlsx outputs

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
    description="Autonomous Training Agent Execution Server v2.0",
    version="2.0.0"
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
        "version": "2.0.0",
        "claude_configured": claude_client is not None,
        "document_formats": ["docx", "xlsx"],
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
    
    # Valid agents - renamed TNA to Analysis, added evaluation
    valid_agents = ['analysis', 'design', 'delivery', 'evaluation', 'full-package']
    
    # Support legacy names and map full-package to evaluation
    agent = request.agent
    if agent == 'tna':
        agent = 'analysis'
    elif agent == 'course-generator':
        agent = 'evaluation'
    elif agent == 'full-package':
        agent = 'evaluation'  # Full Package now produces Evaluation outputs
    
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
    
    # Create output directory for this job
    job_output_dir = OUTPUT_DIR / job_id
    job_output_dir.mkdir(exist_ok=True)
    
    # Start agent execution in background
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
        
        # Determine media type
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
    
    # Set up styles
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
    
    # Add footer with page numbers
    footer = doc.sections[0].footer
    footer_para = footer.paragraphs[0]
    footer_para.text = "Classification: OFFICIAL"
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    return doc


def add_title_page(doc: Document, title: str, subtitle: str, metadata: Dict[str, str]):
    """Add a professional title page"""
    # Add spacing
    for _ in range(3):
        doc.add_paragraph()
    
    # Main title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run(title)
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = RGBColor(0, 51, 102)
    
    # Subtitle
    if subtitle:
        sub_para = doc.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = sub_para.add_run(subtitle)
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(80, 80, 80)
    
    # Spacing
    for _ in range(4):
        doc.add_paragraph()
    
    # Metadata table
    table = doc.add_table(rows=len(metadata), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i, (key, value) in enumerate(metadata.items()):
        row = table.rows[i]
        row.cells[0].text = key
        row.cells[1].text = value
        row.cells[0].paragraphs[0].runs[0].bold = True
    
    # Page break
    doc.add_page_break()


def add_section_heading(doc: Document, text: str, level: int = 1):
    """Add a section heading"""
    if level == 1:
        doc.add_heading(text, level=1)
    else:
        doc.add_heading(text, level=2)


def add_table_from_data(doc: Document, headers: List[str], rows: List[List[str]], 
                         header_color: str = "003366"):
    """Add a formatted table"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    
    # Header row
    header_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = header
        # Style header
        para = cell.paragraphs[0]
        para.runs[0].bold = True
        para.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        # Background color
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), header_color)
        cell._tc.get_or_add_tcPr().append(shading)
    
    # Data rows
    for i, row_data in enumerate(rows):
        row = table.rows[i + 1]
        for j, cell_text in enumerate(row_data):
            row.cells[j].text = str(cell_text)
    
    doc.add_paragraph()  # Spacing after table


# ============================================================================
# CLAUDE API
# ============================================================================

async def call_claude(prompt: str, system_prompt: str = None) -> str:
    """Call Claude API to generate content"""
    if not claude_client:
        print("[NOVA] WARNING: Claude API not configured")
        return "[Claude API not configured - please set ANTHROPIC_API_KEY]"
    
    try:
        print(f"[NOVA] Calling Claude API (prompt length: {len(prompt)})")
        messages = [{"role": "user", "content": prompt}]
        
        kwargs = {
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 8192,
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


# System prompt - Framework agnostic
TRAINING_SYSTEM_PROMPT = """You are NOVA, an expert training analysis and design system. You are proficient in:
- UK Defence Systems Approach to Training (DSAT) - JSP 822 and DTSM 1-5
- US Army Systems Approach to Training (SAT) - TRADOC 350-70 / ADDIE
- NATO Training - Bi-SC Directive 75-7
- ASD/AIA S6000T Training Analysis and Design Standard
- Corporate Learning & Development methodologies
- Competency-based training frameworks

You generate professional, methodology-compliant training documentation.
Adapt your terminology and references based on the specified framework.
Always use formal professional tone appropriate for official documentation.

When generating structured content:
- Use clear section headers
- Provide specific, actionable content
- Include realistic examples appropriate to the role
- Reference appropriate methodology standards"""


# ============================================================================
# CONTENT GENERATION FUNCTIONS
# ============================================================================

async def generate_scoping_content(role_title: str, framework: str, description: str = "") -> Dict:
    """Generate scoping report content"""
    framework_ref = get_framework_reference(framework, "scoping")
    
    prompt = f"""Generate content for a Scoping Exercise Report for training analysis.

Role Title: {role_title}
Framework: {framework}
Additional Context: {description if description else 'None provided'}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Generate the following sections as a JSON object with these exact keys:
{{
    "introduction": "2-3 paragraphs on purpose and scope",
    "background": "2-3 paragraphs on context and need",
    "scope_inclusions": ["list of 5-7 items in scope"],
    "scope_exclusions": ["list of 3-5 items out of scope"],
    "stakeholders": [
        {{"role": "stakeholder role", "responsibility": "what they do"}}
    ],
    "assumptions": ["list of 5-7 key assumptions"],
    "risks": [
        {{"risk": "description", "impact": "High/Medium/Low", "mitigation": "action"}}
    ],
    "resource_estimate": {{
        "duration": "estimated time",
        "personnel": "required team",
        "budget": "estimated cost range"
    }},
    "recommendations": ["list of 3-5 next steps"]
}}

Be specific and realistic for a {role_title} role.
Reference {framework_ref} where appropriate.
Return ONLY the JSON object, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        # Extract JSON from response
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    # Fallback structure
    return {
        "introduction": f"This Scoping Exercise Report initiates the training needs analysis for the {role_title} role.",
        "background": "Training analysis required to establish capability requirements.",
        "scope_inclusions": ["Role analysis", "Task identification", "Gap analysis"],
        "scope_exclusions": ["Equipment procurement", "Facility upgrades"],
        "stakeholders": [{"role": "Training Manager", "responsibility": "Overall coordination"}],
        "assumptions": ["Current role requirements are documented"],
        "risks": [{"risk": "Scope creep", "impact": "Medium", "mitigation": "Regular reviews"}],
        "resource_estimate": {"duration": "4-6 weeks", "personnel": "2-3 analysts", "budget": "TBD"},
        "recommendations": ["Proceed with role analysis", "Engage stakeholders"]
    }


async def generate_role_tasks(role_title: str, framework: str, description: str = "") -> List[Dict]:
    """Generate role performance tasks"""
    prompt = f"""Generate realistic tasks for a Role Performance Statement / Task Analysis.

Role Title: {role_title}
Framework: {framework}
Additional Context: {description if description else 'None provided'}

Generate 10-12 specific tasks as a JSON array. Each task should have:
{{
    "task_number": "1.0",
    "performance": "What the person must do (action verb + object)",
    "conditions": "Under what circumstances (equipment, environment, resources)",
    "standards": "To what measurable standard (time, accuracy, compliance)",
    "category": "FT/WPT/OJT/CBT",
    "ksa": "Knowledge, Skills, and Attitudes required"
}}

Categories:
- FT = Formal Training (classroom/structured)
- WPT = Workplace Training
- OJT = On-the-Job Training
- CBT = Computer-Based Training

Make tasks specific, measurable, and realistic for a {role_title}.
Return ONLY the JSON array, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\[[\s\S]*\]', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    # Fallback
    return [
        {
            "task_number": "1.0",
            "performance": f"Perform core {role_title} duties",
            "conditions": "Standard workplace environment",
            "standards": "In accordance with procedures",
            "category": "FT",
            "ksa": "Knowledge of role requirements"
        }
    ]


async def generate_gap_analysis(role_title: str, framework: str, tasks: List[Dict]) -> List[Dict]:
    """Generate training gap analysis"""
    task_summary = "\n".join([f"- {t.get('performance', 'Task')}" for t in tasks[:5]])
    
    prompt = f"""Generate a Training Gap Analysis based on these role tasks:

Role Title: {role_title}
Framework: {framework}
Sample Tasks:
{task_summary}

Generate gap analysis as a JSON array with 8-10 gaps:
{{
    "skill_area": "Specific skill or competency area",
    "current_provision": "What training currently exists",
    "required_standard": "What standard is needed",
    "gap_description": "Description of the gap",
    "risk_rating": "High/Medium/Low",
    "priority": 1-10 (1 = highest)
}}

Return ONLY the JSON array."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\[[\s\S]*\]', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    return [{"skill_area": "Core competencies", "current_provision": "Limited", 
             "required_standard": "Full proficiency", "gap_description": "Gap identified",
             "risk_rating": "Medium", "priority": 1}]


async def generate_training_objectives(role_title: str, framework: str, tasks: List[Dict]) -> List[Dict]:
    """Generate Training Objectives"""
    task_summary = "\n".join([f"- {t.get('performance', 'Task')}" for t in tasks[:6]])
    
    prompt = f"""Generate Training Objectives based on role tasks.

Role Title: {role_title}
Framework: {framework}
Tasks:
{task_summary}

Generate 8-10 Training Objectives as JSON array:
{{
    "to_number": "TO 1",
    "objective": "Full objective statement",
    "performance": "Observable action (verb + object)",
    "conditions": "Circumstances and resources",
    "standards": "Measurable criteria",
    "domain": "Cognitive/Psychomotor/Affective",
    "level": "Remember/Understand/Apply/Analyze/Evaluate/Create",
    "assessment_method": "How competence will be verified"
}}

Make objectives SMART: Specific, Measurable, Achievable, Relevant, Time-bound.
Return ONLY the JSON array."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\[[\s\S]*\]', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    return [{"to_number": "TO 1", "objective": f"Perform {role_title} duties",
             "performance": "Demonstrate competency", "conditions": "Standard environment",
             "standards": "To required standard", "domain": "Cognitive", 
             "level": "Apply", "assessment_method": "Practical assessment"}]


async def generate_enabling_objectives(role_title: str, framework: str, 
                                        training_objectives: List[Dict]) -> List[Dict]:
    """Generate Enabling Objectives and Key Learning Points"""
    to_summary = "\n".join([f"- {to.get('to_number', 'TO')}: {to.get('objective', '')[:80]}" 
                           for to in training_objectives[:4]])
    
    prompt = f"""Generate Enabling Objectives (EOs) and Key Learning Points (KLPs) for these Training Objectives.

Role Title: {role_title}
Training Objectives:
{to_summary}

Generate as JSON array - for each TO, create 2-4 EOs, and for each EO, 3-5 KLPs:
{{
    "to_number": "TO 1",
    "to_text": "Training objective text",
    "enabling_objectives": [
        {{
            "eo_number": "EO 1.1",
            "eo_text": "Enabling objective text",
            "klps": [
                {{"klp_number": "KLP 1.1.1", "klp_text": "Key learning point", "type": "K/S/A"}}
            ]
        }}
    ]
}}

K = Knowledge, S = Skill, A = Attitude
Return ONLY the JSON array."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\[[\s\S]*\]', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    return []


async def generate_lesson_plans(role_title: str, framework: str, 
                                 training_objectives: List[Dict]) -> List[Dict]:
    """Generate Lesson Plans"""
    to_summary = "\n".join([f"- {to.get('to_number', 'TO')}: {to.get('objective', '')[:60]}" 
                           for to in training_objectives[:3]])
    
    prompt = f"""Generate detailed Lesson Plans for training delivery.

Role Title: {role_title}
Training Objectives:
{to_summary}

Generate 3-4 lesson plans as JSON array:
{{
    "lesson_number": 1,
    "title": "Lesson title",
    "duration": "Duration in minutes",
    "objectives_covered": ["TO 1", "TO 2"],
    "prerequisites": ["What learners need to know"],
    "resources": ["Equipment, materials needed"],
    "introduction": {{
        "duration": "5 mins",
        "attention_getter": "Hook to engage learners",
        "learning_outcomes": "What they will achieve",
        "overview": "Lesson structure"
    }},
    "present": {{
        "duration": "20 mins",
        "content_points": ["Key points to cover"],
        "trainer_notes": "Guidance for instructor",
        "visual_aids": ["Slides, demonstrations"]
    }},
    "apply": {{
        "duration": "20 mins",
        "activities": ["Practical exercises"],
        "assessment": "Formative check"
    }},
    "review": {{
        "duration": "5 mins",
        "summary": "Key takeaways",
        "questions": ["Discussion questions"],
        "next_lesson": "Link to next topic"
    }},
    "safety_notes": "Any safety considerations",
    "common_errors": ["Typical mistakes to address"]
}}

Return ONLY the JSON array."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\[[\s\S]*\]', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    return []


async def generate_assessments(role_title: str, framework: str,
                                training_objectives: List[Dict]) -> Dict:
    """Generate Assessment Instruments"""
    to_list = [to.get('to_number', 'TO') for to in training_objectives[:5]]
    
    prompt = f"""Generate Assessment Instruments for training evaluation.

Role Title: {role_title}
Training Objectives: {', '.join(to_list)}

Generate as JSON object:
{{
    "assessment_strategy": {{
        "purpose": "Overall assessment approach",
        "pass_criteria": "What constitutes a pass",
        "remediation_policy": "Support for those who fail",
        "ai_policy": "Policy on AI tool usage in assessments"
    }},
    "practical_assessments": [
        {{
            "title": "Assessment title",
            "objectives_assessed": ["TO 1"],
            "description": "What learner must demonstrate",
            "criteria": ["Observable criteria"],
            "pass_standard": "Minimum acceptable performance",
            "time_limit": "Duration allowed",
            "assessor_guidance": "Notes for assessor"
        }}
    ],
    "theory_questions": [
        {{
            "question_number": 1,
            "question": "Question text",
            "options": ["A) Option", "B) Option", "C) Option", "D) Option"],
            "correct_answer": "A",
            "objective_assessed": "TO 1",
            "difficulty": "Easy/Medium/Hard"
        }}
    ],
    "marking_scheme": {{
        "practical_weighting": "60%",
        "theory_weighting": "40%",
        "overall_pass_mark": "70%"
    }}
}}

Generate at least 2 practical assessments and 10 theory questions.
Return ONLY the JSON object."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    return {
        "assessment_strategy": {"purpose": "Evaluate competency", "pass_criteria": "70%"},
        "practical_assessments": [],
        "theory_questions": [],
        "marking_scheme": {"overall_pass_mark": "70%"}
    }


# ============================================================================
# DOCUMENT BUILDERS
# ============================================================================

def build_scoping_report(role_title: str, framework: str, content: Dict, 
                         output_path: Path) -> str:
    """Build Scoping Report document"""
    doc = create_styled_document("Scoping Exercise Report", role_title, framework)
    
    # Title page
    add_title_page(doc, "SCOPING EXERCISE REPORT", role_title, {
        "Document Type": "Scoping Exercise Report",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Introduction
    add_section_heading(doc, "1. INTRODUCTION")
    doc.add_paragraph(content.get("introduction", ""))
    
    # Background
    add_section_heading(doc, "2. BACKGROUND")
    doc.add_paragraph(content.get("background", ""))
    
    # Scope
    add_section_heading(doc, "3. SCOPE")
    doc.add_heading("3.1 Inclusions", level=2)
    for item in content.get("scope_inclusions", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("3.2 Exclusions", level=2)
    for item in content.get("scope_exclusions", []):
        doc.add_paragraph(item, style='List Bullet')
    
    # Stakeholders
    add_section_heading(doc, "4. STAKEHOLDERS")
    stakeholders = content.get("stakeholders", [])
    if stakeholders:
        add_table_from_data(doc, 
            ["Role", "Responsibility"],
            [[s.get("role", ""), s.get("responsibility", "")] for s in stakeholders]
        )
    
    # Assumptions
    add_section_heading(doc, "5. ASSUMPTIONS")
    for item in content.get("assumptions", []):
        doc.add_paragraph(item, style='List Bullet')
    
    # Risks
    add_section_heading(doc, "6. RISKS")
    risks = content.get("risks", [])
    if risks:
        add_table_from_data(doc,
            ["Risk", "Impact", "Mitigation"],
            [[r.get("risk", ""), r.get("impact", ""), r.get("mitigation", "")] for r in risks]
        )
    
    # Resource Estimate
    add_section_heading(doc, "7. RESOURCE ESTIMATE")
    res = content.get("resource_estimate", {})
    doc.add_paragraph(f"Duration: {res.get('duration', 'TBD')}")
    doc.add_paragraph(f"Personnel: {res.get('personnel', 'TBD')}")
    doc.add_paragraph(f"Budget: {res.get('budget', 'TBD')}")
    
    # Recommendations
    add_section_heading(doc, "8. RECOMMENDATIONS")
    for i, item in enumerate(content.get("recommendations", []), 1):
        doc.add_paragraph(f"{i}. {item}")
    
    # Save
    filename = "01_Scoping_Report.docx"
    doc.save(output_path / filename)
    return filename


def build_role_performance_statement(role_title: str, framework: str, tasks: List[Dict],
                                      output_path: Path) -> str:
    """Build Role Performance Statement document"""
    doc = create_styled_document("Role Performance Statement", role_title, framework)
    
    # Title page
    term = get_framework_term(framework, "roleps")
    add_title_page(doc, term.upper(), role_title, {
        "Document Type": term,
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Introduction
    add_section_heading(doc, "1. INTRODUCTION")
    doc.add_paragraph(f"This {term} defines the tasks, performance standards, and training "
                     f"requirements for the {role_title} role.")
    
    # Task Analysis
    add_section_heading(doc, "2. TASK ANALYSIS")
    
    # Build task table
    headers = ["Task No.", "Performance", "Conditions", "Standards", "Category", "KSA"]
    rows = []
    for task in tasks:
        rows.append([
            task.get("task_number", ""),
            task.get("performance", ""),
            task.get("conditions", ""),
            task.get("standards", ""),
            task.get("category", ""),
            task.get("ksa", "")
        ])
    
    add_table_from_data(doc, headers, rows)
    
    # Summary
    add_section_heading(doc, "3. SUMMARY")
    doc.add_paragraph(f"Total Tasks Identified: {len(tasks)}")
    
    # Count by category
    categories = {}
    for task in tasks:
        cat = task.get("category", "Other")
        categories[cat] = categories.get(cat, 0) + 1
    
    for cat, count in categories.items():
        doc.add_paragraph(f"• {cat}: {count} tasks", style='List Bullet')
    
    filename = "02_Role_Performance_Statement.docx"
    doc.save(output_path / filename)
    return filename


def build_gap_analysis(role_title: str, framework: str, gaps: List[Dict],
                       output_path: Path) -> str:
    """Build Training Gap Analysis document"""
    doc = create_styled_document("Training Gap Analysis", role_title, framework)
    
    add_title_page(doc, "TRAINING GAP ANALYSIS", role_title, {
        "Document Type": "Training Gap Analysis",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Executive Summary
    add_section_heading(doc, "1. EXECUTIVE SUMMARY")
    high_risks = len([g for g in gaps if g.get("risk_rating") == "High"])
    doc.add_paragraph(f"This analysis identifies {len(gaps)} training gaps for the {role_title} role. "
                     f"Of these, {high_risks} are rated as high priority.")
    
    # Gap Analysis Table
    add_section_heading(doc, "2. GAP ANALYSIS")
    headers = ["Skill Area", "Current Provision", "Required Standard", "Gap", "Risk", "Priority"]
    rows = [[
        g.get("skill_area", ""),
        g.get("current_provision", ""),
        g.get("required_standard", ""),
        g.get("gap_description", ""),
        g.get("risk_rating", ""),
        str(g.get("priority", ""))
    ] for g in gaps]
    
    add_table_from_data(doc, headers, rows)
    
    # Recommendations
    add_section_heading(doc, "3. RECOMMENDATIONS")
    doc.add_paragraph("Based on the gap analysis, the following actions are recommended:")
    doc.add_paragraph("1. Address high-priority gaps through formal training intervention")
    doc.add_paragraph("2. Develop workplace training packages for medium-priority gaps")
    doc.add_paragraph("3. Implement continuous professional development for ongoing needs")
    
    filename = "03_Training_Gap_Analysis.docx"
    doc.save(output_path / filename)
    return filename


def build_training_needs_report(role_title: str, framework: str, 
                                 scoping: Dict, tasks: List[Dict], gaps: List[Dict],
                                 output_path: Path) -> str:
    """Build Training Needs Report (executive summary)"""
    doc = create_styled_document("Training Needs Report", role_title, framework)
    
    term = get_framework_term(framework, "tnr")
    add_title_page(doc, term.upper(), role_title, {
        "Document Type": term,
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL",
        "Status": "For Approval"
    })
    
    # Executive Summary
    add_section_heading(doc, "1. EXECUTIVE SUMMARY")
    high_gaps = len([g for g in gaps if g.get("risk_rating") == "High"])
    doc.add_paragraph(
        f"This Training Needs Report presents the findings of the training analysis conducted "
        f"for the {role_title} role. The analysis identified {len(tasks)} core tasks and "
        f"{len(gaps)} training gaps, of which {high_gaps} require priority attention."
    )
    
    # Background
    add_section_heading(doc, "2. BACKGROUND")
    doc.add_paragraph(scoping.get("background", "Training analysis required to establish requirements."))
    
    # Key Findings
    add_section_heading(doc, "3. KEY FINDINGS")
    doc.add_heading("3.1 Task Analysis", level=2)
    doc.add_paragraph(f"{len(tasks)} tasks identified for the role.")
    
    doc.add_heading("3.2 Gap Analysis", level=2)
    doc.add_paragraph(f"{len(gaps)} training gaps identified.")
    
    # Options
    add_section_heading(doc, "4. TRAINING OPTIONS")
    options = [
        ("Option A: Formal Course", "Structured residential training programme", "High", "£££"),
        ("Option B: Blended Learning", "Mix of online and face-to-face delivery", "Medium", "££"),
        ("Option C: Workplace Training", "On-the-job training with mentoring", "Low", "£"),
    ]
    add_table_from_data(doc, 
        ["Option", "Description", "Effectiveness", "Cost"],
        [list(o) for o in options]
    )
    
    # Recommendations
    add_section_heading(doc, "5. RECOMMENDATIONS")
    for item in scoping.get("recommendations", ["Proceed with training development"]):
        doc.add_paragraph(f"• {item}", style='List Bullet')
    
    # Next Steps
    add_section_heading(doc, "6. NEXT STEPS")
    doc.add_paragraph("1. Approval of this report by governance board")
    doc.add_paragraph("2. Proceed to Design phase")
    doc.add_paragraph("3. Develop detailed training specification")
    
    filename = "04_Training_Needs_Report.docx"
    doc.save(output_path / filename)
    return filename


def build_training_objectives_doc(role_title: str, framework: str, 
                                   objectives: List[Dict], output_path: Path) -> str:
    """Build Training Objectives document"""
    doc = create_styled_document("Training Objectives", role_title, framework)
    
    term = get_framework_term(framework, "to")
    add_title_page(doc, f"{term.upper()}S", role_title, {
        "Document Type": f"{term}s",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0"
    })
    
    # Introduction
    add_section_heading(doc, "1. INTRODUCTION")
    doc.add_paragraph(f"This document defines the {term}s for the {role_title} training programme.")
    
    # Objectives
    add_section_heading(doc, f"2. {term.upper()}S")
    
    for obj in objectives:
        doc.add_heading(f"{obj.get('to_number', 'TO')}: {obj.get('objective', '')[:60]}", level=2)
        
        # Details table
        table = doc.add_table(rows=6, cols=2)
        table.style = 'Table Grid'
        
        details = [
            ("Performance", obj.get("performance", "")),
            ("Conditions", obj.get("conditions", "")),
            ("Standards", obj.get("standards", "")),
            ("Domain", obj.get("domain", "")),
            ("Level", obj.get("level", "")),
            ("Assessment", obj.get("assessment_method", ""))
        ]
        
        for i, (label, value) in enumerate(details):
            table.rows[i].cells[0].text = label
            table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
            table.rows[i].cells[1].text = value
        
        doc.add_paragraph()
    
    filename = "05_Training_Objectives.docx"
    doc.save(output_path / filename)
    return filename


def build_enabling_objectives_doc(role_title: str, framework: str,
                                   eo_data: List[Dict], output_path: Path) -> str:
    """Build Enabling Objectives and KLPs document"""
    doc = create_styled_document("Enabling Objectives", role_title, framework)
    
    eo_term = get_framework_term(framework, "eo")
    klp_term = get_framework_term(framework, "klp")
    
    add_title_page(doc, f"{eo_term.upper()}S & {klp_term.upper()}S", role_title, {
        "Document Type": f"{eo_term}s and {klp_term}s",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0"
    })
    
    add_section_heading(doc, "1. OBJECTIVES HIERARCHY")
    
    for to_item in eo_data:
        # Training Objective
        doc.add_heading(f"{to_item.get('to_number', 'TO')}: {to_item.get('to_text', '')[:50]}", level=2)
        
        for eo in to_item.get("enabling_objectives", []):
            # Enabling Objective
            para = doc.add_paragraph()
            run = para.add_run(f"{eo.get('eo_number', 'EO')}: {eo.get('eo_text', '')}")
            run.bold = True
            
            # KLPs
            for klp in eo.get("klps", []):
                klp_para = doc.add_paragraph(style='List Bullet')
                klp_para.add_run(f"{klp.get('klp_number', 'KLP')}: ").bold = True
                klp_para.add_run(f"{klp.get('klp_text', '')} ")
                klp_para.add_run(f"[{klp.get('type', 'K')}]").italic = True
        
        doc.add_paragraph()
    
    filename = "06_Enabling_Objectives.docx"
    doc.save(output_path / filename)
    return filename


def build_lesson_plans_doc(role_title: str, framework: str,
                            lessons: List[Dict], output_path: Path) -> str:
    """Build Lesson Plans document"""
    doc = create_styled_document("Lesson Plans", role_title, framework)
    
    add_title_page(doc, "LESSON PLANS", role_title, {
        "Document Type": "Lesson Plans",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0"
    })
    
    for lesson in lessons:
        # Lesson header
        add_section_heading(doc, f"LESSON {lesson.get('lesson_number', '')}: {lesson.get('title', '')}")
        
        # Metadata table
        table = doc.add_table(rows=4, cols=2)
        table.style = 'Table Grid'
        meta = [
            ("Duration", lesson.get("duration", "")),
            ("Objectives", ", ".join(lesson.get("objectives_covered", []))),
            ("Prerequisites", ", ".join(lesson.get("prerequisites", []))),
            ("Resources", ", ".join(lesson.get("resources", [])))
        ]
        for i, (k, v) in enumerate(meta):
            table.rows[i].cells[0].text = k
            table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
            table.rows[i].cells[1].text = str(v)
        
        doc.add_paragraph()
        
        # Introduction
        intro = lesson.get("introduction", {})
        doc.add_heading("Introduction", level=2)
        doc.add_paragraph(f"Duration: {intro.get('duration', '5 mins')}")
        doc.add_paragraph(f"Attention Getter: {intro.get('attention_getter', '')}")
        doc.add_paragraph(f"Learning Outcomes: {intro.get('learning_outcomes', '')}")
        
        # Present
        present = lesson.get("present", {})
        doc.add_heading("Present (Main Content)", level=2)
        doc.add_paragraph(f"Duration: {present.get('duration', '20 mins')}")
        for point in present.get("content_points", []):
            doc.add_paragraph(f"• {point}", style='List Bullet')
        doc.add_paragraph(f"Trainer Notes: {present.get('trainer_notes', '')}")
        
        # Apply
        apply_section = lesson.get("apply", {})
        doc.add_heading("Apply (Practice)", level=2)
        doc.add_paragraph(f"Duration: {apply_section.get('duration', '20 mins')}")
        for activity in apply_section.get("activities", []):
            doc.add_paragraph(f"• {activity}", style='List Bullet')
        
        # Review
        review = lesson.get("review", {})
        doc.add_heading("Review", level=2)
        doc.add_paragraph(f"Duration: {review.get('duration', '5 mins')}")
        doc.add_paragraph(f"Summary: {review.get('summary', '')}")
        
        # Notes
        if lesson.get("safety_notes"):
            doc.add_heading("Safety Notes", level=2)
            doc.add_paragraph(lesson.get("safety_notes", ""))
        
        if lesson.get("common_errors"):
            doc.add_heading("Common Errors", level=2)
            for error in lesson.get("common_errors", []):
                doc.add_paragraph(f"• {error}", style='List Bullet')
        
        doc.add_page_break()
    
    filename = "07_Lesson_Plans.docx"
    doc.save(output_path / filename)
    return filename


def build_assessments_doc(role_title: str, framework: str,
                           assessments: Dict, output_path: Path) -> str:
    """Build Assessment Instruments document"""
    doc = create_styled_document("Assessment Instruments", role_title, framework)
    
    add_title_page(doc, "ASSESSMENT INSTRUMENTS", role_title, {
        "Document Type": "Assessment Instruments",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0"
    })
    
    # Assessment Strategy
    add_section_heading(doc, "1. ASSESSMENT STRATEGY")
    strategy = assessments.get("assessment_strategy", {})
    doc.add_paragraph(f"Purpose: {strategy.get('purpose', '')}")
    doc.add_paragraph(f"Pass Criteria: {strategy.get('pass_criteria', '')}")
    doc.add_paragraph(f"Remediation: {strategy.get('remediation_policy', '')}")
    doc.add_paragraph(f"AI Policy: {strategy.get('ai_policy', 'AI tools not permitted during assessment')}")
    
    # Practical Assessments
    add_section_heading(doc, "2. PRACTICAL ASSESSMENTS")
    for pa in assessments.get("practical_assessments", []):
        doc.add_heading(pa.get("title", "Practical Assessment"), level=2)
        doc.add_paragraph(f"Objectives Assessed: {', '.join(pa.get('objectives_assessed', []))}")
        doc.add_paragraph(f"Description: {pa.get('description', '')}")
        doc.add_paragraph(f"Time Limit: {pa.get('time_limit', '')}")
        doc.add_paragraph(f"Pass Standard: {pa.get('pass_standard', '')}")
        
        doc.add_heading("Assessment Criteria", level=3)
        for criterion in pa.get("criteria", []):
            doc.add_paragraph(f"☐ {criterion}", style='List Bullet')
        
        doc.add_paragraph(f"Assessor Guidance: {pa.get('assessor_guidance', '')}")
        doc.add_paragraph()
    
    # Theory Assessment
    add_section_heading(doc, "3. THEORY ASSESSMENT")
    questions = assessments.get("theory_questions", [])
    for q in questions:
        doc.add_paragraph(f"Q{q.get('question_number', '')}: {q.get('question', '')}")
        for option in q.get("options", []):
            doc.add_paragraph(f"    {option}")
        doc.add_paragraph(f"    Correct: {q.get('correct_answer', '')} | "
                         f"Objective: {q.get('objective_assessed', '')} | "
                         f"Difficulty: {q.get('difficulty', '')}")
        doc.add_paragraph()
    
    # Marking Scheme
    add_section_heading(doc, "4. MARKING SCHEME")
    marking = assessments.get("marking_scheme", {})
    for key, value in marking.items():
        doc.add_paragraph(f"• {key.replace('_', ' ').title()}: {value}", style='List Bullet')
    
    filename = "08_Assessment_Instruments.docx"
    doc.save(output_path / filename)
    return filename


def build_compliance_certificate(role_title: str, framework: str, job_id: str,
                                  files_generated: List[str], output_path: Path) -> str:
    """Build Compliance Certificate"""
    doc = create_styled_document("Compliance Certificate", role_title, framework)
    
    # Center everything
    for _ in range(4):
        doc.add_paragraph()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("NOVA™ TRAINING PACKAGE")
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = RGBColor(0, 51, 102)
    
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Compliance Certificate")
    run.font.size = Pt(18)
    
    for _ in range(2):
        doc.add_paragraph()
    
    # Certificate content
    content = doc.add_paragraph()
    content.alignment = WD_ALIGN_PARAGRAPH.CENTER
    content.add_run(f"Role: {role_title}\n\n").bold = True
    content.add_run(f"Framework: {framework}\n")
    content.add_run(f"Generated: {datetime.utcnow().strftime('%d %B %Y %H:%M UTC')}\n")
    content.add_run(f"Job ID: {job_id}\n\n")
    
    # Files list
    files_para = doc.add_paragraph()
    files_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    files_para.add_run("Documents Generated:\n").bold = True
    for f in files_generated:
        files_para.add_run(f"✓ {f}\n")
    
    # Status
    for _ in range(2):
        doc.add_paragraph()
    
    status = doc.add_paragraph()
    status.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = status.add_run("PACKAGE COMPLETE")
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0, 128, 0)
    
    # Footer
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("\n\nNOVA™ - Training Intelligence System")
    
    filename = "00_Compliance_Certificate.docx"
    doc.save(output_path / filename)
    return filename


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
        }
    }
    return refs.get(framework, refs["UK"]).get(doc_type, "")


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
                "to": "Training Requirement", "eo": "Sub-Task Requirement", "klp": "Task Element"}
    }
    return terms.get(framework, terms["UK"]).get(term_type, term_type)


# ============================================================================
# AGENT EXECUTION
# ============================================================================

def update_progress(job_id: str, progress: int, step: str):
    """Helper to update job progress"""
    if job_id in jobs:
        jobs[job_id]["progress"] = progress
        jobs[job_id]["current_step"] = step
        if step not in jobs[job_id]["steps_completed"]:
            jobs[job_id]["steps_completed"].append(step)


async def run_agent(job_id: str, agent: str, parameters: Dict[str, Any]):
    """Execute the specified agent"""
    print(f"[NOVA] Background task started for job {job_id}, agent={agent}")
    job = jobs[job_id]
    job["status"] = "running"
    
    try:
        if agent == "analysis":
            print(f"[NOVA] Running analysis agent for job {job_id}")
            await run_analysis_agent(job_id, parameters)
        elif agent == "design":
            print(f"[NOVA] Running design agent for job {job_id}")
            await run_design_agent(job_id, parameters)
        elif agent == "delivery":
            print(f"[NOVA] Running delivery agent for job {job_id}")
            await run_delivery_agent(job_id, parameters)
        elif agent == "evaluation":
            print(f"[NOVA] Running evaluation agent for job {job_id}")
            await run_evaluation_agent(job_id, parameters)
        elif agent == "full-package":
            print(f"[NOVA] Running evaluation agent (full-package) for job {job_id}")
            await run_evaluation_agent(job_id, parameters)
        
        job["status"] = "completed"
        job["progress"] = 100
        job["completed_at"] = datetime.utcnow().isoformat()
        print(f"[NOVA] Job {job_id} completed successfully")
        
    except Exception as e:
        print(f"[NOVA] Job {job_id} failed with error: {str(e)}")
        job["status"] = "failed"
        job["error"] = str(e)
        job["completed_at"] = datetime.utcnow().isoformat()


async def run_analysis_agent(job_id: str, parameters: Dict[str, Any]):
    """Analysis Agent (formerly TNA)"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "UK")
    description = parameters.get("role_description", "")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    print(f"[NOVA] Analysis agent starting: role={role_title}, framework={framework}")
    
    analysis_dir = output_dir / "01_Analysis"
    analysis_dir.mkdir(exist_ok=True)
    
    files_generated = []
    
    # Step 1: Generate Scoping Report
    print(f"[NOVA] Step 1: Generating Scoping Report")
    update_progress(job_id, 15, "Generating Scoping Report")
    scoping_content = await generate_scoping_content(role_title, framework, description)
    print(f"[NOVA] Scoping content generated, building document")
    filename = build_scoping_report(role_title, framework, scoping_content, analysis_dir)
    files_generated.append(filename)
    print(f"[NOVA] Scoping report saved: {filename}")
    
    # Step 2: Generate Role Tasks
    print(f"[NOVA] Step 2: Generating Role Performance Statement")
    update_progress(job_id, 35, "Generating Role Performance Statement")
    tasks = await generate_role_tasks(role_title, framework, description)
    print(f"[NOVA] Tasks generated: {len(tasks)} tasks")
    filename = build_role_performance_statement(role_title, framework, tasks, analysis_dir)
    files_generated.append(filename)
    print(f"[NOVA] RolePS saved: {filename}")
    
    # Step 3: Gap Analysis
    print(f"[NOVA] Step 3: Conducting Training Gap Analysis")
    update_progress(job_id, 55, "Conducting Training Gap Analysis")
    gaps = await generate_gap_analysis(role_title, framework, tasks)
    print(f"[NOVA] Gap analysis generated")
    filename = build_gap_analysis(role_title, framework, gaps, analysis_dir)
    files_generated.append(filename)
    print(f"[NOVA] Gap analysis saved: {filename}")
    
    # Step 4: Training Needs Report
    print(f"[NOVA] Step 4: Compiling Training Needs Report")
    update_progress(job_id, 75, "Compiling Training Needs Report")
    filename = build_training_needs_report(role_title, framework, scoping_content, tasks, gaps, analysis_dir)
    files_generated.append(filename)
    print(f"[NOVA] TNR saved: {filename}")
    
    # Store data for other agents
    jobs[job_id]["analysis_data"] = {
        "scoping": scoping_content,
        "tasks": tasks,
        "gaps": gaps
    }
    
    update_progress(job_id, 100, "Analysis Complete")


async def run_design_agent(job_id: str, parameters: Dict[str, Any]):
    """Design Agent"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "UK")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    design_dir = output_dir / "02_Design"
    design_dir.mkdir(exist_ok=True)
    
    # Get analysis data if available
    analysis_data = jobs[job_id].get("analysis_data", {})
    tasks = analysis_data.get("tasks", [])
    
    # If no tasks, generate them
    if not tasks:
        update_progress(job_id, 10, "Generating task data")
        tasks = await generate_role_tasks(role_title, framework)
    
    files_generated = []
    
    # Step 1: Training Objectives
    update_progress(job_id, 25, "Generating Training Objectives")
    objectives = await generate_training_objectives(role_title, framework, tasks)
    filename = build_training_objectives_doc(role_title, framework, objectives, design_dir)
    files_generated.append(filename)
    
    # Step 2: Enabling Objectives
    update_progress(job_id, 50, "Creating Enabling Objectives")
    eo_data = await generate_enabling_objectives(role_title, framework, objectives)
    filename = build_enabling_objectives_doc(role_title, framework, eo_data, design_dir)
    files_generated.append(filename)
    
    # Store for delivery
    jobs[job_id]["design_data"] = {
        "objectives": objectives,
        "enabling_objectives": eo_data
    }
    
    update_progress(job_id, 100, "Design Complete")


async def run_delivery_agent(job_id: str, parameters: Dict[str, Any]):
    """Delivery Agent"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "UK")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    delivery_dir = output_dir / "03_Delivery"
    delivery_dir.mkdir(exist_ok=True)
    
    # Get design data if available
    design_data = jobs[job_id].get("design_data", {})
    objectives = design_data.get("objectives", [])
    
    if not objectives:
        update_progress(job_id, 10, "Generating objectives data")
        tasks = await generate_role_tasks(role_title, framework)
        objectives = await generate_training_objectives(role_title, framework, tasks)
    
    files_generated = []
    
    # Step 1: Lesson Plans
    update_progress(job_id, 35, "Generating Lesson Plans")
    lessons = await generate_lesson_plans(role_title, framework, objectives)
    filename = build_lesson_plans_doc(role_title, framework, lessons, delivery_dir)
    files_generated.append(filename)
    
    # Step 2: Assessments
    update_progress(job_id, 70, "Creating Assessment Instruments")
    assessments = await generate_assessments(role_title, framework, objectives)
    filename = build_assessments_doc(role_title, framework, assessments, delivery_dir)
    files_generated.append(filename)
    
    update_progress(job_id, 100, "Delivery Complete")


async def run_evaluation_agent(job_id: str, parameters: Dict[str, Any]):
    """Evaluation and Assurance Agent - DSAT Element 4 (DTSM 5)"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "UK")
    description = parameters.get("role_description", "")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    print(f"[NOVA] Evaluation agent starting: role={role_title}, framework={framework}")
    
    eval_dir = output_dir / "04_Evaluation"
    eval_dir.mkdir(exist_ok=True)
    
    files_generated = []
    
    # Step 1: Generate Evaluation Strategy
    print(f"[NOVA] Step 1: Generating Evaluation Strategy")
    update_progress(job_id, 15, "Generating Evaluation Strategy")
    eval_strategy = await generate_evaluation_strategy(role_title, framework, description)
    filename = build_evaluation_strategy_doc(role_title, framework, eval_strategy, eval_dir)
    files_generated.append(filename)
    print(f"[NOVA] Evaluation Strategy saved: {filename}")
    
    # Step 2: Generate Internal Validation Plan
    print(f"[NOVA] Step 2: Generating Internal Validation Plan")
    update_progress(job_id, 35, "Generating Internal Validation Plan")
    inval_plan = await generate_internal_validation_plan(role_title, framework, eval_strategy)
    filename = build_internal_validation_doc(role_title, framework, inval_plan, eval_dir)
    files_generated.append(filename)
    print(f"[NOVA] InVal Plan saved: {filename}")
    
    # Step 3: Generate External Validation Plan
    print(f"[NOVA] Step 3: Generating External Validation Plan")
    update_progress(job_id, 55, "Generating External Validation Plan")
    exval_plan = await generate_external_validation_plan(role_title, framework, eval_strategy)
    filename = build_external_validation_doc(role_title, framework, exval_plan, eval_dir)
    files_generated.append(filename)
    print(f"[NOVA] ExVal Plan saved: {filename}")
    
    # Step 4: Generate Training Needs Evaluation
    print(f"[NOVA] Step 4: Generating Training Needs Evaluation")
    update_progress(job_id, 75, "Generating Training Needs Evaluation")
    tne = await generate_training_needs_evaluation(role_title, framework, eval_strategy)
    filename = build_training_needs_evaluation_doc(role_title, framework, tne, eval_dir)
    files_generated.append(filename)
    print(f"[NOVA] TNE saved: {filename}")
    
    # Step 5: Generate Quality Assurance Framework
    print(f"[NOVA] Step 5: Generating Quality Assurance Framework")
    update_progress(job_id, 90, "Generating Quality Assurance Framework")
    qa_framework = await generate_qa_framework(role_title, framework, eval_strategy, inval_plan)
    filename = build_qa_framework_doc(role_title, framework, qa_framework, eval_dir)
    files_generated.append(filename)
    print(f"[NOVA] QA Framework saved: {filename}")
    
    # Store evaluation data
    jobs[job_id]["evaluation_data"] = {
        "strategy": eval_strategy,
        "inval": inval_plan,
        "exval": exval_plan,
        "tne": tne,
        "qa": qa_framework
    }
    
    update_progress(job_id, 100, "Evaluation Complete")


# ============================================================================
# EVALUATION CONTENT GENERATION FUNCTIONS
# ============================================================================

async def generate_evaluation_strategy(role_title: str, framework: str, description: str = "") -> Dict:
    """Generate Evaluation Strategy content (DTSM 5 Section 4.2)"""
    framework_ref = get_framework_reference(framework, "evaluation")
    
    prompt = f"""Generate content for an Evaluation Strategy document for training evaluation.

Role Title: {role_title}
Framework: {framework}
Additional Context: {description if description else 'None provided'}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Generate the following sections as a JSON object with these exact keys:
{{
    "purpose": "2-3 paragraphs on purpose and scope of evaluation",
    "evaluation_approach": {{
        "philosophy": "Overall evaluation philosophy",
        "methodology": "Kirkpatrick or other framework used",
        "scope": "What will be evaluated"
    }},
    "kirkpatrick_levels": [
        {{
            "level": 1,
            "name": "Reaction",
            "focus": "What is measured at this level",
            "methods": ["List of data collection methods"],
            "timing": "When data is collected",
            "responsibility": "Who collects data",
            "success_criteria": "Target metrics"
        }}
    ],
    "data_collection_plan": [
        {{
            "data_type": "Type of data",
            "source": "Where collected from",
            "method": "How collected",
            "frequency": "How often",
            "owner": "Responsible party"
        }}
    ],
    "reporting_schedule": [
        {{
            "report_type": "Type of report",
            "frequency": "How often produced",
            "audience": "Who receives it",
            "content": "What it contains"
        }}
    ],
    "continuous_improvement": {{
        "feedback_loop": "How findings feed back into training",
        "review_cycle": "Frequency of strategy review",
        "governance": "Who approves changes"
    }},
    "resources_required": {{
        "personnel": "Evaluation team requirements",
        "tools": "Systems and instruments needed",
        "budget": "Estimated evaluation costs"
    }}
}}

Ensure Kirkpatrick levels 1-4 are all included (Reaction, Learning, Behaviour, Results).
Reference {framework_ref} where appropriate.
Return ONLY the JSON object, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    # Fallback structure
    return {
        "purpose": f"This Evaluation Strategy establishes the framework for measuring training effectiveness for the {role_title} role.",
        "evaluation_approach": {"philosophy": "Evidence-based evaluation", "methodology": "Kirkpatrick Four-Level Model"},
        "kirkpatrick_levels": [
            {"level": 1, "name": "Reaction", "focus": "Trainee satisfaction", "methods": ["End of course survey"], "timing": "Immediately post-training"},
            {"level": 2, "name": "Learning", "focus": "Knowledge/skill acquisition", "methods": ["Assessments"], "timing": "End of course"},
            {"level": 3, "name": "Behaviour", "focus": "Transfer to workplace", "methods": ["Supervisor assessment"], "timing": "3-6 months post"},
            {"level": 4, "name": "Results", "focus": "Operational impact", "methods": ["Performance metrics"], "timing": "6-12 months post"}
        ],
        "data_collection_plan": [],
        "reporting_schedule": [],
        "continuous_improvement": {"feedback_loop": "Quarterly review", "review_cycle": "Annual"},
        "resources_required": {"personnel": "Evaluation team", "tools": "Survey platform", "budget": "TBD"}
    }


async def generate_internal_validation_plan(role_title: str, framework: str, eval_strategy: Dict) -> Dict:
    """Generate Internal Validation (InVal) Plan (DTSM 5 Section 4.3)"""
    
    prompt = f"""Generate an Internal Validation (InVal) Plan for continuous quality assurance.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Internal Validation is continuous quality assurance conducted BY the Training Provider.

Generate as JSON object:
{{
    "purpose": "Purpose of internal validation",
    "scope": "What InVal covers",
    "responsibilities": {{
        "inval_team": "Who conducts InVal",
        "training_provider": "TP responsibilities",
        "instructors": "Trainer responsibilities"
    }},
    "validation_activities": [
        {{
            "activity": "Activity name",
            "description": "What it involves",
            "frequency": "How often",
            "method": "How conducted",
            "outputs": "What it produces",
            "criteria": "Success measures"
        }}
    ],
    "observation_protocol": {{
        "frequency": "How often lessons observed",
        "sample_size": "Coverage required",
        "criteria": ["Observation criteria list"],
        "feedback_process": "How feedback given to trainers",
        "documentation": "Records maintained"
    }},
    "trainee_feedback": {{
        "mechanisms": ["Feedback collection methods"],
        "timing": "When collected",
        "analysis": "How analysed",
        "action_triggers": "Thresholds for action"
    }},
    "assessment_monitoring": {{
        "aspects": ["What is monitored about assessments"],
        "sampling": "Assessment sampling approach",
        "moderation": "Moderation process"
    }},
    "reporting": {{
        "internal_reports": "Reports to Training Provider management",
        "ceb_reports": "Reports to Customer Executive Board",
        "frequency": "Reporting cycle"
    }},
    "improvement_process": {{
        "identification": "How issues identified",
        "prioritisation": "How issues prioritised",
        "implementation": "How changes implemented",
        "verification": "How improvements verified"
    }}
}}

Return ONLY the JSON object."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    return {
        "purpose": f"To provide continuous quality assurance for {role_title} training",
        "scope": "All training delivery and assessment",
        "validation_activities": [],
        "observation_protocol": {"frequency": "Monthly", "criteria": []},
        "trainee_feedback": {"mechanisms": ["Course evaluation forms"]},
        "reporting": {"frequency": "Quarterly"}
    }


async def generate_external_validation_plan(role_title: str, framework: str, eval_strategy: Dict) -> Dict:
    """Generate External Validation (ExVal) Plan (DTSM 5 Section 4.4)"""
    
    prompt = f"""Generate an External Validation (ExVal) Plan for independent training audit.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

External Validation is independent audit conducted BY/FOR the Training Requirements Authority (TRA).
ExVal measures whether training produces personnel who can perform their role in the workplace.

Generate as JSON object:
{{
    "purpose": "Purpose of external validation",
    "scope": "What ExVal covers",
    "governance": {{
        "tra_role": "TRA responsibilities in ExVal",
        "exval_team": "Who conducts ExVal",
        "independence": "How independence maintained"
    }},
    "validation_focus": [
        {{
            "area": "Validation focus area",
            "questions": ["Key questions to answer"],
            "evidence_sources": ["Where evidence gathered"],
            "criteria": "Success criteria"
        }}
    ],
    "workplace_assessment": {{
        "purpose": "Why assess in workplace",
        "timing": "When assessment occurs post-training",
        "sample_selection": "How trainees selected",
        "assessment_method": "How competence assessed",
        "assessor_requirements": "Who can assess",
        "documentation": "Records required"
    }},
    "data_collection": [
        {{
            "method": "Data collection method",
            "source": "Data source",
            "sample_size": "How much data",
            "timing": "When collected"
        }}
    ],
    "employer_consultation": {{
        "purpose": "Why consult employers",
        "stakeholders": ["Who to consult"],
        "questions": ["Key questions to ask"],
        "method": "Consultation approach"
    }},
    "exval_cycle": {{
        "frequency": "How often full ExVal conducted",
        "phases": ["Phase list"],
        "duration": "How long ExVal takes"
    }},
    "reporting": {{
        "report_structure": ["Report sections"],
        "recommendations_format": "How recommendations presented",
        "approval_process": "Report approval route",
        "distribution": "Who receives report"
    }},
    "action_planning": {{
        "response_requirements": "How TDA must respond",
        "timeframes": "Action completion timeframes",
        "follow_up": "How actions verified"
    }}
}}

Return ONLY the JSON object."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    return {
        "purpose": f"To independently validate that {role_title} training produces workplace-ready personnel",
        "scope": "Training effectiveness and transfer",
        "governance": {"tra_role": "Commission and approve ExVal"},
        "validation_focus": [],
        "workplace_assessment": {"timing": "3-6 months post-training"},
        "exval_cycle": {"frequency": "Annual"},
        "reporting": {"report_structure": ["Executive Summary", "Findings", "Recommendations"]}
    }


async def generate_training_needs_evaluation(role_title: str, framework: str, eval_strategy: Dict) -> Dict:
    """Generate Training Needs Evaluation (TNE) content (DTSM 5 Section 1.8)"""
    
    prompt = f"""Generate a Training Needs Evaluation (TNE) - post-TNA quality assessment.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

TNE assesses the quality and completeness of the Training Needs Analysis that was conducted.
It feeds back into the Analysis phase for continuous improvement.

Generate as JSON object:
{{
    "purpose": "Purpose of TNE",
    "scope": "What TNE assesses",
    "evaluation_criteria": [
        {{
            "criterion": "Evaluation criterion",
            "description": "What is assessed",
            "rating_scale": "How rated",
            "evidence_required": "Evidence to examine"
        }}
    ],
    "tna_quality_assessment": {{
        "scoping_review": {{
            "aspects": ["Aspects to review"],
            "questions": ["Review questions"],
            "rating": "Assessment outcome"
        }},
        "role_analysis_review": {{
            "aspects": ["Aspects to review"],
            "questions": ["Review questions"],
            "rating": "Assessment outcome"
        }},
        "gap_analysis_review": {{
            "aspects": ["Aspects to review"],
            "questions": ["Review questions"],
            "rating": "Assessment outcome"
        }},
        "stakeholder_engagement_review": {{
            "aspects": ["Aspects to review"],
            "questions": ["Review questions"],
            "rating": "Assessment outcome"
        }}
    }},
    "findings": [
        {{
            "area": "TNA area",
            "finding": "What was found",
            "impact": "Impact on training quality",
            "recommendation": "Improvement action"
        }}
    ],
    "overall_assessment": {{
        "tna_quality_rating": "Overall rating (Satisfactory/Requires Improvement/Unsatisfactory)",
        "confidence_level": "Confidence in TNA outputs",
        "key_strengths": ["TNA strengths"],
        "key_weaknesses": ["TNA weaknesses"],
        "critical_gaps": ["Any critical gaps in analysis"]
    }},
    "feedback_to_analysis": {{
        "immediate_actions": ["Actions needed now"],
        "future_improvements": ["Improvements for next TNA cycle"],
        "methodology_recommendations": ["Process improvements"]
    }},
    "lessons_learned": ["Key lessons from this TNA"]
}}

Return ONLY the JSON object."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    return {
        "purpose": f"To assess the quality and completeness of the Training Needs Analysis for {role_title}",
        "scope": "All TNA deliverables",
        "evaluation_criteria": [],
        "tna_quality_assessment": {},
        "findings": [],
        "overall_assessment": {"tna_quality_rating": "Satisfactory"},
        "feedback_to_analysis": {"immediate_actions": [], "future_improvements": []}
    }


async def generate_qa_framework(role_title: str, framework: str, eval_strategy: Dict, inval_plan: Dict) -> Dict:
    """Generate Quality Assurance Framework (TQM components)"""
    
    prompt = f"""Generate a Quality Assurance Framework document for training quality management.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

This document provides the Training Quality Manual (TQM) framework and quality assurance procedures.

Generate as JSON object:
{{
    "purpose": "Purpose of QA Framework",
    "scope": "QA scope and boundaries",
    "quality_policy": {{
        "policy_statement": "Overall quality commitment",
        "quality_objectives": ["Measurable objectives"],
        "management_commitment": "Leadership responsibilities"
    }},
    "quality_standards": [
        {{
            "standard": "Quality standard",
            "requirement": "What it requires",
            "measurement": "How measured",
            "target": "Target level"
        }}
    ],
    "quality_roles": [
        {{
            "role": "QA role title",
            "responsibilities": ["Role responsibilities"],
            "authority": "Decision authority",
            "reporting_line": "Reports to"
        }}
    ],
    "quality_processes": [
        {{
            "process": "Process name",
            "purpose": "What it achieves",
            "inputs": ["Process inputs"],
            "outputs": ["Process outputs"],
            "owner": "Process owner",
            "frequency": "How often run"
        }}
    ],
    "quality_records": {{
        "records_required": [
            {{
                "record_type": "Type of record",
                "content": "What it contains",
                "retention_period": "How long kept",
                "storage": "Where stored",
                "access": "Who can access"
            }}
        ],
        "management_system": "How records managed",
        "audit_requirements": "Audit accessibility requirements"
    }},
    "continuous_improvement": {{
        "improvement_cycle": "PDCA or other model",
        "sources_of_improvement": ["Where improvements come from"],
        "prioritisation": "How improvements prioritised",
        "implementation": "How changes implemented",
        "review_frequency": "QA framework review cycle"
    }},
    "compliance_monitoring": {{
        "internal_audits": "Internal audit approach",
        "external_audits": "External audit preparation",
        "non_conformance": "How non-conformances handled",
        "corrective_actions": "Corrective action process"
    }},
    "kpis": [
        {{
            "kpi": "KPI name",
            "definition": "What it measures",
            "target": "Target value",
            "frequency": "Measurement frequency",
            "owner": "Who monitors"
        }}
    ]
}}

Return ONLY the JSON object."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    return {
        "purpose": f"To establish quality assurance framework for {role_title} training",
        "scope": "All training activities",
        "quality_policy": {"policy_statement": "Commitment to training excellence"},
        "quality_standards": [],
        "quality_roles": [],
        "quality_processes": [],
        "quality_records": {"records_required": []},
        "continuous_improvement": {},
        "kpis": []
    }


# ============================================================================
# EVALUATION DOCUMENT BUILDERS
# ============================================================================

def build_evaluation_strategy_doc(role_title: str, framework: str, content: Dict, 
                                   output_path: Path) -> str:
    """Build Evaluation Strategy document"""
    doc = create_styled_document("Evaluation Strategy", role_title, framework)
    
    add_title_page(doc, "EVALUATION STRATEGY", role_title, {
        "Document Type": "Evaluation Strategy (DTSM 5 Section 4.2)",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Purpose
    add_section_heading(doc, "1. PURPOSE AND SCOPE")
    doc.add_paragraph(content.get("purpose", ""))
    
    approach = content.get("evaluation_approach", {})
    if approach:
        doc.add_heading("1.1 Evaluation Approach", level=2)
        doc.add_paragraph(f"Philosophy: {approach.get('philosophy', '')}")
        doc.add_paragraph(f"Methodology: {approach.get('methodology', '')}")
        doc.add_paragraph(f"Scope: {approach.get('scope', '')}")
    
    # Kirkpatrick Levels
    add_section_heading(doc, "2. EVALUATION FRAMEWORK (KIRKPATRICK MODEL)")
    levels = content.get("kirkpatrick_levels", [])
    if levels:
        headers = ["Level", "Name", "Focus", "Methods", "Timing", "Success Criteria"]
        rows = []
        for level in levels:
            methods = level.get("methods", [])
            methods_str = ", ".join(methods) if isinstance(methods, list) else str(methods)
            rows.append([
                str(level.get("level", "")),
                level.get("name", ""),
                level.get("focus", ""),
                methods_str,
                level.get("timing", ""),
                level.get("success_criteria", "")
            ])
        add_table_from_data(doc, headers, rows)
    
    # Data Collection Plan
    add_section_heading(doc, "3. DATA COLLECTION PLAN")
    data_plan = content.get("data_collection_plan", [])
    if data_plan:
        headers = ["Data Type", "Source", "Method", "Frequency", "Owner"]
        rows = [[
            d.get("data_type", ""),
            d.get("source", ""),
            d.get("method", ""),
            d.get("frequency", ""),
            d.get("owner", "")
        ] for d in data_plan]
        add_table_from_data(doc, headers, rows)
    
    # Reporting Schedule
    add_section_heading(doc, "4. REPORTING SCHEDULE")
    reports = content.get("reporting_schedule", [])
    if reports:
        headers = ["Report Type", "Frequency", "Audience", "Content"]
        rows = [[
            r.get("report_type", ""),
            r.get("frequency", ""),
            r.get("audience", ""),
            r.get("content", "")
        ] for r in reports]
        add_table_from_data(doc, headers, rows)
    
    # Continuous Improvement
    add_section_heading(doc, "5. CONTINUOUS IMPROVEMENT")
    ci = content.get("continuous_improvement", {})
    if ci:
        doc.add_paragraph(f"Feedback Loop: {ci.get('feedback_loop', '')}")
        doc.add_paragraph(f"Review Cycle: {ci.get('review_cycle', '')}")
        doc.add_paragraph(f"Governance: {ci.get('governance', '')}")
    
    # Resources
    add_section_heading(doc, "6. RESOURCES REQUIRED")
    res = content.get("resources_required", {})
    if res:
        doc.add_paragraph(f"Personnel: {res.get('personnel', '')}")
        doc.add_paragraph(f"Tools: {res.get('tools', '')}")
        doc.add_paragraph(f"Budget: {res.get('budget', '')}")
    
    filename = "01_Evaluation_Strategy.docx"
    doc.save(output_path / filename)
    return filename


def build_internal_validation_doc(role_title: str, framework: str, content: Dict,
                                   output_path: Path) -> str:
    """Build Internal Validation Plan document"""
    doc = create_styled_document("Internal Validation Plan", role_title, framework)
    
    add_title_page(doc, "INTERNAL VALIDATION (InVal) PLAN", role_title, {
        "Document Type": "Internal Validation Plan (DTSM 5 Section 4.3)",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Purpose
    add_section_heading(doc, "1. PURPOSE AND SCOPE")
    doc.add_paragraph(content.get("purpose", ""))
    doc.add_paragraph(f"Scope: {content.get('scope', '')}")
    
    # Responsibilities
    add_section_heading(doc, "2. RESPONSIBILITIES")
    resp = content.get("responsibilities", {})
    if resp:
        for key, value in resp.items():
            doc.add_paragraph(f"{key.replace('_', ' ').title()}: {value}")
    
    # Validation Activities
    add_section_heading(doc, "3. VALIDATION ACTIVITIES")
    activities = content.get("validation_activities", [])
    if activities:
        headers = ["Activity", "Description", "Frequency", "Method", "Outputs"]
        rows = [[
            a.get("activity", ""),
            a.get("description", ""),
            a.get("frequency", ""),
            a.get("method", ""),
            a.get("outputs", "")
        ] for a in activities]
        add_table_from_data(doc, headers, rows)
    
    # Observation Protocol
    add_section_heading(doc, "4. LESSON OBSERVATION PROTOCOL")
    obs = content.get("observation_protocol", {})
    if obs:
        doc.add_paragraph(f"Frequency: {obs.get('frequency', '')}")
        doc.add_paragraph(f"Sample Size: {obs.get('sample_size', '')}")
        criteria = obs.get("criteria", [])
        if criteria:
            doc.add_heading("4.1 Observation Criteria", level=2)
            for c in criteria:
                doc.add_paragraph(f"• {c}", style='List Bullet')
    
    # Trainee Feedback
    add_section_heading(doc, "5. TRAINEE FEEDBACK MECHANISMS")
    feedback = content.get("trainee_feedback", {})
    if feedback:
        mechanisms = feedback.get("mechanisms", [])
        for m in mechanisms:
            doc.add_paragraph(f"• {m}", style='List Bullet')
    
    # Reporting
    add_section_heading(doc, "6. REPORTING")
    reporting = content.get("reporting", {})
    if reporting:
        doc.add_paragraph(f"Internal Reports: {reporting.get('internal_reports', '')}")
        doc.add_paragraph(f"CEB Reports: {reporting.get('ceb_reports', '')}")
        doc.add_paragraph(f"Frequency: {reporting.get('frequency', '')}")
    
    filename = "02_Internal_Validation_Plan.docx"
    doc.save(output_path / filename)
    return filename


def build_external_validation_doc(role_title: str, framework: str, content: Dict,
                                   output_path: Path) -> str:
    """Build External Validation Plan document"""
    doc = create_styled_document("External Validation Plan", role_title, framework)
    
    add_title_page(doc, "EXTERNAL VALIDATION (ExVal) PLAN", role_title, {
        "Document Type": "External Validation Plan (DTSM 5 Section 4.4)",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Purpose
    add_section_heading(doc, "1. PURPOSE AND SCOPE")
    doc.add_paragraph(content.get("purpose", ""))
    doc.add_paragraph(f"Scope: {content.get('scope', '')}")
    
    # Governance
    add_section_heading(doc, "2. GOVERNANCE")
    gov = content.get("governance", {})
    if gov:
        doc.add_paragraph(f"TRA Role: {gov.get('tra_role', '')}")
        doc.add_paragraph(f"ExVal Team: {gov.get('exval_team', '')}")
        doc.add_paragraph(f"Independence: {gov.get('independence', '')}")
    
    # Validation Focus
    add_section_heading(doc, "3. VALIDATION FOCUS AREAS")
    focus = content.get("validation_focus", [])
    if focus:
        for f in focus:
            doc.add_heading(f"3.x {f.get('area', '')}", level=2)
            questions = f.get("questions", [])
            for q in questions:
                doc.add_paragraph(f"• {q}", style='List Bullet')
    
    # Workplace Assessment
    add_section_heading(doc, "4. WORKPLACE COMPETENCE ASSESSMENT")
    workplace = content.get("workplace_assessment", {})
    if workplace:
        doc.add_paragraph(f"Purpose: {workplace.get('purpose', '')}")
        doc.add_paragraph(f"Timing: {workplace.get('timing', '')}")
        doc.add_paragraph(f"Sample Selection: {workplace.get('sample_selection', '')}")
        doc.add_paragraph(f"Assessment Method: {workplace.get('assessment_method', '')}")
    
    # Employer Consultation
    add_section_heading(doc, "5. EMPLOYER CONSULTATION")
    employer = content.get("employer_consultation", {})
    if employer:
        doc.add_paragraph(f"Purpose: {employer.get('purpose', '')}")
        stakeholders = employer.get("stakeholders", [])
        if stakeholders:
            doc.add_paragraph("Stakeholders to consult:")
            for s in stakeholders:
                doc.add_paragraph(f"• {s}", style='List Bullet')
    
    # ExVal Cycle
    add_section_heading(doc, "6. EXVAL CYCLE")
    cycle = content.get("exval_cycle", {})
    if cycle:
        doc.add_paragraph(f"Frequency: {cycle.get('frequency', '')}")
        doc.add_paragraph(f"Duration: {cycle.get('duration', '')}")
        phases = cycle.get("phases", [])
        if phases:
            for p in phases:
                doc.add_paragraph(f"• {p}", style='List Bullet')
    
    # Reporting
    add_section_heading(doc, "7. REPORTING AND ACTION PLANNING")
    reporting = content.get("reporting", {})
    if reporting:
        doc.add_paragraph(f"Approval Process: {reporting.get('approval_process', '')}")
        doc.add_paragraph(f"Distribution: {reporting.get('distribution', '')}")
    
    action = content.get("action_planning", {})
    if action:
        doc.add_paragraph(f"Response Requirements: {action.get('response_requirements', '')}")
        doc.add_paragraph(f"Timeframes: {action.get('timeframes', '')}")
    
    filename = "03_External_Validation_Plan.docx"
    doc.save(output_path / filename)
    return filename


def build_training_needs_evaluation_doc(role_title: str, framework: str, content: Dict,
                                         output_path: Path) -> str:
    """Build Training Needs Evaluation document"""
    doc = create_styled_document("Training Needs Evaluation", role_title, framework)
    
    add_title_page(doc, "TRAINING NEEDS EVALUATION (TNE)", role_title, {
        "Document Type": "Training Needs Evaluation (DTSM 5 Section 1.8)",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Purpose
    add_section_heading(doc, "1. PURPOSE AND SCOPE")
    doc.add_paragraph(content.get("purpose", ""))
    doc.add_paragraph(f"Scope: {content.get('scope', '')}")
    
    # Evaluation Criteria
    add_section_heading(doc, "2. EVALUATION CRITERIA")
    criteria = content.get("evaluation_criteria", [])
    if criteria:
        headers = ["Criterion", "Description", "Rating Scale", "Evidence Required"]
        rows = [[
            c.get("criterion", ""),
            c.get("description", ""),
            c.get("rating_scale", ""),
            c.get("evidence_required", "")
        ] for c in criteria]
        add_table_from_data(doc, headers, rows)
    
    # TNA Quality Assessment
    add_section_heading(doc, "3. TNA QUALITY ASSESSMENT")
    tna_qa = content.get("tna_quality_assessment", {})
    
    for section_key, section_data in tna_qa.items():
        if isinstance(section_data, dict):
            section_name = section_key.replace("_", " ").title()
            doc.add_heading(f"3.x {section_name}", level=2)
            aspects = section_data.get("aspects", [])
            for a in aspects:
                doc.add_paragraph(f"• {a}", style='List Bullet')
    
    # Findings
    add_section_heading(doc, "4. FINDINGS")
    findings = content.get("findings", [])
    if findings:
        headers = ["Area", "Finding", "Impact", "Recommendation"]
        rows = [[
            f.get("area", ""),
            f.get("finding", ""),
            f.get("impact", ""),
            f.get("recommendation", "")
        ] for f in findings]
        add_table_from_data(doc, headers, rows)
    
    # Overall Assessment
    add_section_heading(doc, "5. OVERALL ASSESSMENT")
    overall = content.get("overall_assessment", {})
    if overall:
        doc.add_paragraph(f"TNA Quality Rating: {overall.get('tna_quality_rating', '')}")
        doc.add_paragraph(f"Confidence Level: {overall.get('confidence_level', '')}")
        
        strengths = overall.get("key_strengths", [])
        if strengths:
            doc.add_heading("5.1 Key Strengths", level=2)
            for s in strengths:
                doc.add_paragraph(f"• {s}", style='List Bullet')
        
        weaknesses = overall.get("key_weaknesses", [])
        if weaknesses:
            doc.add_heading("5.2 Key Weaknesses", level=2)
            for w in weaknesses:
                doc.add_paragraph(f"• {w}", style='List Bullet')
    
    # Feedback to Analysis
    add_section_heading(doc, "6. FEEDBACK TO ANALYSIS PHASE")
    feedback = content.get("feedback_to_analysis", {})
    if feedback:
        immediate = feedback.get("immediate_actions", [])
        if immediate:
            doc.add_heading("6.1 Immediate Actions Required", level=2)
            for a in immediate:
                doc.add_paragraph(f"• {a}", style='List Bullet')
        
        future = feedback.get("future_improvements", [])
        if future:
            doc.add_heading("6.2 Future Improvements", level=2)
            for f in future:
                doc.add_paragraph(f"• {f}", style='List Bullet')
    
    filename = "04_Training_Needs_Evaluation.docx"
    doc.save(output_path / filename)
    return filename


def build_qa_framework_doc(role_title: str, framework: str, content: Dict,
                           output_path: Path) -> str:
    """Build Quality Assurance Framework document"""
    doc = create_styled_document("Quality Assurance Framework", role_title, framework)
    
    add_title_page(doc, "QUALITY ASSURANCE FRAMEWORK", role_title, {
        "Document Type": "Quality Assurance Framework (TQM)",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Purpose
    add_section_heading(doc, "1. PURPOSE AND SCOPE")
    doc.add_paragraph(content.get("purpose", ""))
    doc.add_paragraph(f"Scope: {content.get('scope', '')}")
    
    # Quality Policy
    add_section_heading(doc, "2. QUALITY POLICY")
    policy = content.get("quality_policy", {})
    if policy:
        doc.add_paragraph(policy.get("policy_statement", ""))
        objectives = policy.get("quality_objectives", [])
        if objectives:
            doc.add_heading("2.1 Quality Objectives", level=2)
            for o in objectives:
                doc.add_paragraph(f"• {o}", style='List Bullet')
    
    # Quality Standards
    add_section_heading(doc, "3. QUALITY STANDARDS")
    standards = content.get("quality_standards", [])
    if standards:
        headers = ["Standard", "Requirement", "Measurement", "Target"]
        rows = [[
            s.get("standard", ""),
            s.get("requirement", ""),
            s.get("measurement", ""),
            s.get("target", "")
        ] for s in standards]
        add_table_from_data(doc, headers, rows)
    
    # Quality Roles
    add_section_heading(doc, "4. QUALITY ROLES AND RESPONSIBILITIES")
    roles = content.get("quality_roles", [])
    if roles:
        for r in roles:
            doc.add_heading(f"4.x {r.get('role', '')}", level=2)
            responsibilities = r.get("responsibilities", [])
            for resp in responsibilities:
                doc.add_paragraph(f"• {resp}", style='List Bullet')
    
    # Quality Processes
    add_section_heading(doc, "5. QUALITY PROCESSES")
    processes = content.get("quality_processes", [])
    if processes:
        headers = ["Process", "Purpose", "Owner", "Frequency"]
        rows = [[
            p.get("process", ""),
            p.get("purpose", ""),
            p.get("owner", ""),
            p.get("frequency", "")
        ] for p in processes]
        add_table_from_data(doc, headers, rows)
    
    # Quality Records
    add_section_heading(doc, "6. QUALITY RECORDS MANAGEMENT")
    records = content.get("quality_records", {})
    records_list = records.get("records_required", [])
    if records_list:
        headers = ["Record Type", "Content", "Retention", "Storage"]
        rows = [[
            r.get("record_type", ""),
            r.get("content", ""),
            r.get("retention_period", ""),
            r.get("storage", "")
        ] for r in records_list]
        add_table_from_data(doc, headers, rows)
    
    # KPIs
    add_section_heading(doc, "7. KEY PERFORMANCE INDICATORS")
    kpis = content.get("kpis", [])
    if kpis:
        headers = ["KPI", "Definition", "Target", "Frequency", "Owner"]
        rows = [[
            k.get("kpi", ""),
            k.get("definition", ""),
            k.get("target", ""),
            k.get("frequency", ""),
            k.get("owner", "")
        ] for k in kpis]
        add_table_from_data(doc, headers, rows)
    
    # Continuous Improvement
    add_section_heading(doc, "8. CONTINUOUS IMPROVEMENT")
    ci = content.get("continuous_improvement", {})
    if ci:
        doc.add_paragraph(f"Improvement Cycle: {ci.get('improvement_cycle', '')}")
        sources = ci.get("sources_of_improvement", [])
        if sources:
            for s in sources:
                doc.add_paragraph(f"• {s}", style='List Bullet')
    
    filename = "05_Quality_Assurance_Framework.docx"
    doc.save(output_path / filename)
    return filename


async def run_full_package_agent(job_id: str, parameters: Dict[str, Any]):
    """Full Package Agent - redirects to Evaluation Agent"""
    # Full Package now produces Evaluation outputs
    await run_evaluation_agent(job_id, parameters)


# ============================================================================
# RUN SERVER
# ============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
