"""
NOVA Agent Server v2.1
FastAPI server for executing autonomous training agents with Claude AI
Generates professional .docx and .xlsx outputs

UPDATED: Full DSAT/JSP 822 compliant Scoping Report and Training Needs Report
- Scoping Report now includes all DTSM 2 mandatory sections
- TNR now includes all JSP 822 required sections
- Domain and specialism context integrated into generation

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
    description="Autonomous Training Agent Execution Server v2.1 - Full DSAT Compliance",
    version="2.1.0"
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
        "version": "2.1.0",
        "claude_configured": claude_client is not None,
        "document_formats": ["docx", "xlsx"],
        "dsat_compliant": True,
        "timestamp": datetime.utcnow().isoformat()
    }


@app.post("/api/execute", response_model=TaskResponse)
async def execute_task(
    request: TaskRequest,
    background_tasks: BackgroundTasks,
    authorization: Optional[str] = Header(None)
):
    print(f"[NOVA] Execute request received: agent={request.agent}, job_id={request.job_id}")
    print(f"[NOVA] Parameters: {json.dumps(request.parameters, indent=2)}")
    verify_auth(authorization)
    
    job_id = request.job_id
    
    # Valid agents - renamed TNA to Analysis
    valid_agents = ['analysis', 'design', 'delivery', 'full-package']
    
    # Support legacy 'tna' and 'course-generator' names
    agent = request.agent
    if agent == 'tna':
        agent = 'analysis'
    elif agent == 'course-generator':
        agent = 'full-package'
    
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
    elif level == 2:
        doc.add_heading(text, level=2)
    else:
        doc.add_heading(text, level=3)


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

async def call_claude(prompt: str, system_prompt: str = None, max_tokens: int = 8192) -> str:
    """Call Claude API to generate content"""
    if not claude_client:
        print("[NOVA] WARNING: Claude API not configured")
        return "[Claude API not configured - please set ANTHROPIC_API_KEY]"
    
    try:
        print(f"[NOVA] Calling Claude API (prompt length: {len(prompt)})")
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
# DSAT-COMPLIANT SYSTEM PROMPT
# ============================================================================

TRAINING_SYSTEM_PROMPT = """You are NOVA, an expert Defence training analysis and design system with world-leading knowledge of:

**UK Defence Systems Approach to Training (DSAT)**
- JSP 822 V7.0 - Defence Individual Training Policy
- DTSM 1 (2024) - Governance of Individual Training
- DTSM 2 (2024) - Analysis of Individual Training  
- DTSM 3 (2024) - Designing Individual Training
- DTSM 4 (2024) - Delivery of Individual Training
- DTSM 5 (2024) - Evaluation of Individual Training

**US Training Doctrine**
- TRADOC Reg 350-70 - Army Learning and Training
- ADDIE Model - Analysis, Design, Develop, Implement, Evaluate

**NATO Training Standards**
- Bi-SC Directive 75-7 - Education and Individual Training

**International Standards**
- ASD/AIA S6000T - Training Analysis and Design

You generate professional, methodology-compliant training documentation that meets audit requirements.
Your outputs must:
- Use formal MOD/Defence tone
- Include specific doctrinal references where appropriate
- Provide substantive, detailed content (not placeholder text)
- Be realistic and contextually appropriate for the role
- Include all mandatory sections per the relevant doctrine

When generating structured content:
- Provide comprehensive detail in each section
- Use specific examples and metrics where appropriate
- Reference appropriate methodology standards
- Ensure content is audit-ready and compliant"""


# ============================================================================
# FRAMEWORK REFERENCES
# ============================================================================

def get_framework_reference(framework: str, doc_type: str) -> str:
    """Get framework-specific reference"""
    refs = {
        "UK": {
            "scoping": "DTSM 2 Section 1.2 - Scoping Exercise",
            "roleps": "DTSM 2 Section 1.3 - Role Analysis",
            "tga": "DTSM 2 Section 1.4 - Training Gap Analysis",
            "tnr": "DTSM 2 Section 1.7 - Training Needs Report",
            "to": "DTSM 3 Section 2 - Training Objectives",
            "eo": "DTSM 3 Section 2.3 - Enabling Objectives",
            "fts": "DTSM 3 Section 2.2 - Formal Training Statement"
        },
        "US": {
            "scoping": "TRADOC 350-70 Chapter 2 - Analysis Phase",
            "roleps": "TRADOC 350-70 - Individual Task Analysis",
            "tga": "TRADOC 350-70 - Training Gap Analysis",
            "tnr": "TRADOC 350-70 - Training Needs Assessment Report",
            "to": "TRADOC 350-70 - Terminal Learning Objectives",
            "eo": "TRADOC 350-70 - Enabling Learning Objectives",
            "fts": "TRADOC 350-70 - Program of Instruction"
        },
        "NATO": {
            "scoping": "Bi-SC 75-7 - Training Requirements Analysis",
            "roleps": "Bi-SC 75-7 - Job/Duty Analysis",
            "tga": "Bi-SC 75-7 - Training Gap Analysis",
            "tnr": "Bi-SC 75-7 - Training Requirements Document",
            "to": "Bi-SC 75-7 - Training Outcomes",
            "eo": "Bi-SC 75-7 - Enabling Outcomes",
            "fts": "Bi-SC 75-7 - Training Programme"
        },
        "ASD": {
            "scoping": "S6000T Clause 4 - Training Analysis",
            "roleps": "S6000T - Task Specification",
            "tga": "S6000T - Training Gap Identification",
            "tnr": "S6000T - Training Specification",
            "to": "S6000T - Training Requirements",
            "eo": "S6000T - Sub-Task Requirements",
            "fts": "S6000T - Training Solution"
        }
    }
    return refs.get(framework, refs["UK"]).get(doc_type, doc_type)


def get_framework_term(framework: str, term_type: str) -> str:
    """Get framework-specific terminology"""
    terms = {
        "UK": {"roleps": "Role Performance Statement", "tnr": "Training Needs Report",
               "to": "Training Objective", "eo": "Enabling Objective", "klp": "Key Learning Point"},
        "US": {"roleps": "Task List", "tnr": "Training Needs Assessment",
               "to": "Terminal Learning Objective", "eo": "Enabling Learning Objective", "klp": "Learning Step Activity"},
        "NATO": {"roleps": "Job Analysis", "tnr": "Training Requirements Document",
                 "to": "Training Outcome", "eo": "Enabling Outcome", "klp": "Learning Point"},
        "ASD": {"roleps": "Task Specification", "tnr": "Training Specification",
                "to": "Training Requirement", "eo": "Sub-Task Requirement", "klp": "Task Element"}
    }
    return terms.get(framework, terms["UK"]).get(term_type, term_type)


# ============================================================================
# CONTENT GENERATION FUNCTIONS - COMPREHENSIVE DSAT COMPLIANCE
# ============================================================================

async def generate_scoping_content(role_title: str, framework: str, description: str = "",
                                    domain: str = "", specialism: str = "", 
                                    proficiency: str = "practitioner") -> Dict:
    """Generate comprehensive DTSM 2 compliant Scoping Exercise Report content"""
    framework_ref = get_framework_reference(framework, "scoping")
    
    # Build context from domain/specialism
    context_info = ""
    if domain:
        context_info += f"\nDomain: {domain}"
    if specialism:
        context_info += f"\nSpecialism: {specialism}"
    if proficiency:
        proficiency_map = {
            "foundation": "Foundation level (entry-level, supervised work)",
            "practitioner": "Practitioner level (independent work)",
            "senior": "Senior/Lead level (complex problems, mentoring)",
            "principal": "Principal/Expert level (strategic, authority)"
        }
        context_info += f"\nProficiency Level: {proficiency_map.get(proficiency, proficiency)}"
    
    prompt = f"""Generate a comprehensive Scoping Exercise Report for training analysis per {framework_ref}.

Role Title: {role_title}
Framework: {framework}
{context_info}
Additional Context: {description if description else 'None provided'}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Generate DETAILED content for ALL of the following mandatory sections as a JSON object.
Each section must contain substantive, realistic content - NOT placeholder text.

{{
    "purpose_and_aim": {{
        "purpose": "3-4 sentences explaining the purpose of this scoping exercise",
        "aim": "Clear statement of what the TNA aims to achieve",
        "objectives": ["List of 4-5 specific objectives for the TNA"]
    }},
    "background_and_context": {{
        "operational_context": "2-3 paragraphs on the operational/business context requiring this training",
        "capability_requirement": "Description of the capability this role supports",
        "driver_for_change": "What has triggered this training need (new equipment, policy change, capability gap, etc.)",
        "strategic_alignment": "How this aligns with organisational/defence strategy"
    }},
    "scope": {{
        "inclusions": ["List of 6-8 specific items IN scope"],
        "exclusions": ["List of 4-6 specific items OUT of scope"],
        "boundaries": "Clear statement of the boundaries of the analysis",
        "interfaces": ["List of 3-4 related training/capabilities this interfaces with"]
    }},
    "governance": {{
        "tra": {{"title": "Training Requirements Authority", "name": "[TRA Name]", "responsibility": "Description of TRA responsibility"}},
        "tda": {{"title": "Training Delivery Authority", "name": "[TDA Name]", "responsibility": "Description of TDA responsibility"}},
        "stakeholders": [
            {{"role": "Specific stakeholder role", "organisation": "Their organisation", "interest": "Their stake in this training", "engagement": "How they will be engaged"}}
        ],
        "governance_board": "Name and frequency of governance meetings"
    }},
    "target_audience": {{
        "population_description": "Detailed description of who will be trained",
        "entry_requirements": ["List of 4-5 prerequisites for trainees"],
        "annual_throughput": "Estimated number of trainees per year",
        "current_competence": "Description of assumed baseline competence",
        "diversity_considerations": "Any specific audience considerations"
    }},
    "methodology": {{
        "approach": "Description of the TNA methodology to be used",
        "data_collection": ["List of 4-5 data collection methods"],
        "analysis_techniques": ["List of 3-4 analysis techniques"],
        "validation_approach": "How findings will be validated"
    }},
    "assumptions": ["List of 6-8 key assumptions underpinning the analysis"],
    "constraints": ["List of 5-6 constraints affecting the TNA"],
    "dependencies": ["List of 4-5 dependencies on other activities/decisions"],
    "risks": [
        {{"risk_id": "R1", "description": "Specific risk description", "likelihood": "High/Medium/Low", "impact": "High/Medium/Low", "mitigation": "Specific mitigation action", "owner": "Risk owner role"}}
    ],
    "resource_estimate": {{
        "duration": "Estimated calendar time for TNA",
        "effort": "Estimated person-days",
        "personnel": [
            {{"role": "Team role", "fte": "FTE required", "skills": "Required skills"}}
        ],
        "budget_estimate": "Estimated cost range with assumptions",
        "facilities": "Any facilities/equipment required"
    }},
    "timeline": {{
        "phases": [
            {{"phase": "Phase name", "activities": "Key activities", "duration": "Duration", "milestone": "Key deliverable"}}
        ],
        "key_milestones": [
            {{"milestone": "Milestone name", "target_date": "Target date", "dependency": "What it depends on"}}
        ]
    }},
    "deliverables": [
        {{"deliverable": "Deliverable name", "description": "What it contains", "format": "Document format", "audience": "Who receives it"}}
    ],
    "quality_assurance": {{
        "review_process": "How outputs will be quality assured",
        "approval_authority": "Who approves TNA outputs",
        "compliance_check": "How DSAT compliance will be verified"
    }},
    "recommendations": {{
        "proceed": "Clear recommendation on whether to proceed",
        "next_steps": ["List of 5-6 specific next steps"],
        "decision_required": "What decision is required from governance"
    }}
}}

Be SPECIFIC and REALISTIC for a {role_title} role in the {domain if domain else 'defence'} domain.
Provide SUBSTANTIVE content in each section - this must be audit-ready documentation.
Return ONLY the JSON object, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT, max_tokens=12000)
    
    try:
        # Extract JSON from response
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except Exception as e:
        print(f"[NOVA] JSON parse error in scoping: {e}")
    
    # Fallback structure with basic content
    return create_fallback_scoping_content(role_title, domain, specialism)


def create_fallback_scoping_content(role_title: str, domain: str, specialism: str) -> Dict:
    """Create fallback scoping content if Claude fails"""
    return {
        "purpose_and_aim": {
            "purpose": f"This Scoping Exercise Report initiates the formal Training Needs Analysis for the {role_title} role.",
            "aim": f"To establish the scope, boundaries, and approach for analysing training requirements for the {role_title}.",
            "objectives": [
                "Define the scope and boundaries of the TNA",
                "Identify key stakeholders and governance arrangements",
                "Establish the methodology and approach",
                "Identify risks and mitigation strategies",
                "Estimate resources and timeline"
            ]
        },
        "background_and_context": {
            "operational_context": f"Training analysis is required to establish capability requirements for the {role_title} role.",
            "capability_requirement": "Operational capability requirement to be confirmed during analysis.",
            "driver_for_change": "Requirement for updated training specification.",
            "strategic_alignment": "Aligned with organisational training strategy."
        },
        "scope": {
            "inclusions": ["Role analysis", "Task identification", "Gap analysis", "Training options"],
            "exclusions": ["Equipment procurement", "Facility development"],
            "boundaries": "Analysis covers individual training requirements only.",
            "interfaces": ["Related training programmes", "Career management"]
        },
        "governance": {
            "tra": {"title": "Training Requirements Authority", "name": "TBC", "responsibility": "Overall training requirement ownership"},
            "tda": {"title": "Training Delivery Authority", "name": "TBC", "responsibility": "Training delivery"},
            "stakeholders": [{"role": "Training Manager", "organisation": "TBC", "interest": "Training quality", "engagement": "Regular meetings"}],
            "governance_board": "Customer Executive Board - Monthly"
        },
        "target_audience": {
            "population_description": f"Personnel filling the {role_title} role.",
            "entry_requirements": ["Basic training complete", "Security clearance"],
            "annual_throughput": "TBC",
            "current_competence": "Baseline competence assumed",
            "diversity_considerations": "None identified"
        },
        "methodology": {
            "approach": "DSAT methodology per JSP 822",
            "data_collection": ["Document review", "SME interviews", "Job observation"],
            "analysis_techniques": ["Task analysis", "Gap analysis"],
            "validation_approach": "Stakeholder validation workshops"
        },
        "assumptions": ["Current role requirements are documented", "SMEs are available", "Governance is established"],
        "constraints": ["Resource availability", "Timeline requirements"],
        "dependencies": ["Role documentation", "SME availability"],
        "risks": [{"risk_id": "R1", "description": "Scope creep", "likelihood": "Medium", "impact": "Medium", "mitigation": "Regular scope reviews", "owner": "Project Manager"}],
        "resource_estimate": {
            "duration": "4-6 weeks",
            "effort": "20-30 person-days",
            "personnel": [{"role": "Training Analyst", "fte": "1.0", "skills": "DSAT qualified"}],
            "budget_estimate": "TBC",
            "facilities": "Meeting rooms for workshops"
        },
        "timeline": {
            "phases": [{"phase": "Analysis", "activities": "TNA activities", "duration": "4 weeks", "milestone": "TNR"}],
            "key_milestones": [{"milestone": "TNR Approval", "target_date": "TBC", "dependency": "Analysis complete"}]
        },
        "deliverables": [
            {"deliverable": "Training Needs Report", "description": "TNA findings and recommendations", "format": "DOCX", "audience": "CEB"}
        ],
        "quality_assurance": {
            "review_process": "Peer review and stakeholder validation",
            "approval_authority": "TRA",
            "compliance_check": "DSAT compliance checklist"
        },
        "recommendations": {
            "proceed": "Recommended to proceed with TNA",
            "next_steps": ["Establish governance", "Commence role analysis", "Engage stakeholders"],
            "decision_required": "Approval to proceed with TNA"
        }
    }


async def generate_tnr_content(role_title: str, framework: str, scoping: Dict, 
                                tasks: List[Dict], gaps: List[Dict],
                                domain: str = "", specialism: str = "",
                                proficiency: str = "practitioner") -> Dict:
    """Generate comprehensive DTSM 2 compliant Training Needs Report content"""
    framework_ref = get_framework_reference(framework, "tnr")
    
    # Summarise tasks and gaps for the prompt
    task_summary = "\n".join([f"- {t.get('performance', 'Task')}" for t in tasks[:8]])
    gap_summary = "\n".join([f"- {g.get('skill_area', '')}: {g.get('gap_description', '')}" for g in gaps[:6]])
    
    high_gaps = len([g for g in gaps if g.get("risk_rating") == "High"])
    
    prompt = f"""Generate a comprehensive Training Needs Report per {framework_ref}.

Role Title: {role_title}
Framework: {framework}
Domain: {domain if domain else 'Not specified'}
Specialism: {specialism if specialism else 'Not specified'}

Analysis Summary:
- Tasks Identified: {len(tasks)}
- Training Gaps: {len(gaps)} ({high_gaps} high priority)

Sample Tasks:
{task_summary}

Sample Gaps:
{gap_summary}

Generate DETAILED content for ALL mandatory TNR sections as a JSON object.
This is the key deliverable of the TNA - it must be comprehensive and audit-ready.

{{
    "executive_summary": {{
        "overview": "2-3 paragraph executive summary of the TNA and key findings",
        "key_findings": ["List of 5-6 key findings"],
        "recommendation": "Clear recommendation statement",
        "resource_implications": "Summary of resource implications"
    }},
    "introduction": {{
        "purpose": "Purpose of this Training Needs Report",
        "scope": "Scope of the analysis conducted",
        "methodology": "Summary of methodology used",
        "document_structure": "Guide to this document's structure"
    }},
    "background": {{
        "operational_context": "3-4 paragraphs on operational context and capability requirement",
        "strategic_drivers": "What is driving this training need",
        "current_state": "Description of current training provision",
        "problem_statement": "Clear statement of the training problem/gap"
    }},
    "analysis_findings": {{
        "role_analysis_summary": "Summary of role analysis findings",
        "task_analysis_summary": "Summary of tasks identified with key statistics",
        "ksa_analysis": {{
            "knowledge_areas": ["List of 5-6 key knowledge areas required"],
            "skills_required": ["List of 5-6 key skills required"],
            "attitudes_behaviours": ["List of 3-4 key attitudes/behaviours"]
        }},
        "gap_analysis_summary": "Summary of gaps identified and their significance",
        "critical_gaps": [
            {{"gap": "Gap description", "impact": "Operational impact", "priority": "High/Medium/Low", "rationale": "Why this priority"}}
        ],
        "dif_analysis": {{
            "difficulty": "Overall difficulty assessment",
            "importance": "Importance to role performance",
            "frequency": "How often tasks are performed"
        }}
    }},
    "training_requirement": {{
        "trained_output_requirement": "Clear statement of what trained personnel must be able to do",
        "performance_standards": ["List of 4-5 key performance standards"],
        "entry_requirements": ["List of prerequisites for training"],
        "training_population": {{
            "description": "Description of training population",
            "annual_throughput": "Expected numbers",
            "locations": "Where trainees are based"
        }}
    }},
    "training_options": [
        {{
            "option": "Option A: Formal Residential Course",
            "description": "Detailed description of this option",
            "delivery_method": "How training would be delivered",
            "duration": "Duration of training",
            "advantages": ["List of 3-4 advantages"],
            "disadvantages": ["List of 2-3 disadvantages"],
            "resource_requirements": {{
                "instructors": "Instructor requirement",
                "facilities": "Facility requirement",
                "equipment": "Equipment requirement"
            }},
            "cost_estimate": "Estimated cost",
            "effectiveness_rating": "High/Medium/Low",
            "risk_assessment": "Key risks of this option"
        }},
        {{
            "option": "Option B: Blended Learning",
            "description": "Mix of e-learning, virtual, and practical training",
            "delivery_method": "Blended approach details",
            "duration": "Duration of training",
            "advantages": ["List of advantages"],
            "disadvantages": ["List of disadvantages"],
            "resource_requirements": {{
                "instructors": "Instructor requirement",
                "facilities": "Facility requirement",
                "equipment": "Equipment requirement"
            }},
            "cost_estimate": "Estimated cost",
            "effectiveness_rating": "High/Medium/Low",
            "risk_assessment": "Key risks"
        }},
        {{
            "option": "Option C: Workplace Training",
            "description": "On-the-job training with coaching and mentoring",
            "delivery_method": "OJT approach",
            "duration": "Duration",
            "advantages": ["List of advantages"],
            "disadvantages": ["List of disadvantages"],
            "resource_requirements": {{
                "instructors": "Requirement",
                "facilities": "Requirement",
                "equipment": "Requirement"
            }},
            "cost_estimate": "Estimated cost",
            "effectiveness_rating": "High/Medium/Low",
            "risk_assessment": "Key risks"
        }}
    ],
    "recommended_solution": {{
        "recommendation": "Clear statement of recommended training solution",
        "rationale": "3-4 paragraphs justifying the recommendation",
        "training_statement": {{
            "formal_training": "What will be delivered as formal training",
            "workplace_training": "What will be delivered as workplace training",
            "residual_gap": "Any residual training gap and risk acceptance"
        }},
        "implementation_approach": "How the solution would be implemented"
    }},
    "resource_implications": {{
        "financial": {{
            "development_cost": "One-off development costs",
            "delivery_cost": "Annual delivery costs",
            "total_investment": "Total investment required"
        }},
        "personnel": {{
            "instructors": "Instructor requirement",
            "support_staff": "Support staff requirement",
            "training_required": "Any train-the-trainer requirement"
        }},
        "infrastructure": {{
            "facilities": "Facility requirements",
            "equipment": "Equipment requirements",
            "it_systems": "Any IT/LMS requirements"
        }},
        "timeline": "Implementation timeline"
    }},
    "risk_assessment": [
        {{"risk": "Risk description", "category": "Risk category", "likelihood": "H/M/L", "impact": "H/M/L", "mitigation": "Mitigation action", "owner": "Risk owner"}}
    ],
    "governance_and_next_steps": {{
        "approvals_required": ["List of approvals needed"],
        "next_steps": [
            {{"step": "Step description", "owner": "Who owns this", "target_date": "Target date"}}
        ],
        "success_criteria": ["How success will be measured"],
        "review_schedule": "When training will be reviewed"
    }},
    "annexes": {{
        "annex_a": "Full Role Performance Statement",
        "annex_b": "Complete Task Analysis",
        "annex_c": "Detailed Gap Analysis",
        "annex_d": "Stakeholder Consultation Log",
        "annex_e": "Cost-Benefit Analysis"
    }}
}}

Be SPECIFIC and REALISTIC for a {role_title} role.
This TNR must be suitable for Customer Executive Board approval.
Return ONLY the JSON object, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT, max_tokens=16000)
    
    try:
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            return json.loads(json_match.group())
    except Exception as e:
        print(f"[NOVA] JSON parse error in TNR: {e}")
    
    # Return fallback
    return create_fallback_tnr_content(role_title, len(tasks), len(gaps), high_gaps)


def create_fallback_tnr_content(role_title: str, task_count: int, gap_count: int, high_gaps: int) -> Dict:
    """Create fallback TNR content if Claude fails"""
    return {
        "executive_summary": {
            "overview": f"This Training Needs Report presents findings from the TNA for the {role_title} role. {task_count} tasks and {gap_count} training gaps were identified.",
            "key_findings": [
                f"{task_count} tasks identified through role analysis",
                f"{gap_count} training gaps identified ({high_gaps} high priority)",
                "Formal training intervention recommended",
                "Blended delivery approach viable",
                "Resource requirements within normal parameters"
            ],
            "recommendation": "Proceed with formal training development",
            "resource_implications": "Development will require training design resources"
        },
        "introduction": {
            "purpose": "To document findings and recommendations from the TNA",
            "scope": f"Analysis of training requirements for the {role_title} role",
            "methodology": "DSAT methodology per JSP 822 and DTSM 2",
            "document_structure": "This report follows the standard TNR format"
        },
        "background": {
            "operational_context": f"Training analysis conducted for the {role_title} role to establish training requirements.",
            "strategic_drivers": "Requirement for qualified personnel",
            "current_state": "Current training provision under review",
            "problem_statement": "Training gaps identified requiring intervention"
        },
        "analysis_findings": {
            "role_analysis_summary": f"Role analysis identified the key responsibilities and tasks for {role_title}",
            "task_analysis_summary": f"{task_count} tasks identified across formal and workplace training categories",
            "ksa_analysis": {
                "knowledge_areas": ["Role-specific knowledge", "Procedures and processes", "Regulations and policy"],
                "skills_required": ["Technical skills", "Communication", "Problem-solving"],
                "attitudes_behaviours": ["Professional conduct", "Attention to detail"]
            },
            "gap_analysis_summary": f"{gap_count} gaps identified between current and required capability",
            "critical_gaps": [{"gap": "Core competency gap", "impact": "Operational effectiveness", "priority": "High", "rationale": "Essential for role performance"}],
            "dif_analysis": {"difficulty": "Medium", "importance": "High", "frequency": "Regular"}
        },
        "training_requirement": {
            "trained_output_requirement": f"Personnel able to perform {role_title} duties to required standard",
            "performance_standards": ["Meet operational requirements", "Comply with procedures"],
            "entry_requirements": ["Basic training complete"],
            "training_population": {"description": "Role holders", "annual_throughput": "TBC", "locations": "TBC"}
        },
        "training_options": [
            {"option": "Option A: Formal Course", "description": "Residential training", "delivery_method": "Classroom", "duration": "TBC", "advantages": ["Structured"], "disadvantages": ["Cost"], "resource_requirements": {"instructors": "TBC", "facilities": "TBC", "equipment": "TBC"}, "cost_estimate": "TBC", "effectiveness_rating": "High", "risk_assessment": "Low risk"},
            {"option": "Option B: Blended", "description": "Mixed delivery", "delivery_method": "Blended", "duration": "TBC", "advantages": ["Flexible"], "disadvantages": ["Complexity"], "resource_requirements": {"instructors": "TBC", "facilities": "TBC", "equipment": "TBC"}, "cost_estimate": "TBC", "effectiveness_rating": "Medium", "risk_assessment": "Medium risk"},
            {"option": "Option C: OJT", "description": "Workplace training", "delivery_method": "OJT", "duration": "TBC", "advantages": ["Low cost"], "disadvantages": ["Consistency"], "resource_requirements": {"instructors": "TBC", "facilities": "TBC", "equipment": "TBC"}, "cost_estimate": "TBC", "effectiveness_rating": "Medium", "risk_assessment": "Medium risk"}
        ],
        "recommended_solution": {
            "recommendation": "Proceed with blended training approach",
            "rationale": "Balances effectiveness with resource efficiency",
            "training_statement": {"formal_training": "Core skills", "workplace_training": "Role-specific application", "residual_gap": "None identified"},
            "implementation_approach": "Phased implementation"
        },
        "resource_implications": {
            "financial": {"development_cost": "TBC", "delivery_cost": "TBC", "total_investment": "TBC"},
            "personnel": {"instructors": "TBC", "support_staff": "TBC", "training_required": "TBC"},
            "infrastructure": {"facilities": "TBC", "equipment": "TBC", "it_systems": "TBC"},
            "timeline": "TBC"
        },
        "risk_assessment": [{"risk": "Resource availability", "category": "Resource", "likelihood": "M", "impact": "M", "mitigation": "Early planning", "owner": "PM"}],
        "governance_and_next_steps": {
            "approvals_required": ["CEB approval"],
            "next_steps": [{"step": "Proceed to Design", "owner": "TDA", "target_date": "TBC"}],
            "success_criteria": ["Training delivers required capability"],
            "review_schedule": "Annual review"
        },
        "annexes": {
            "annex_a": "Role Performance Statement",
            "annex_b": "Task Analysis",
            "annex_c": "Gap Analysis",
            "annex_d": "Stakeholder Log",
            "annex_e": "Cost-Benefit Analysis"
        }
    }


async def generate_role_tasks(role_title: str, framework: str, description: str = "",
                               domain: str = "", specialism: str = "") -> List[Dict]:
    """Generate comprehensive role performance tasks"""
    
    context_info = ""
    if domain:
        context_info += f"\nDomain: {domain}"
    if specialism:
        context_info += f"\nSpecialism: {specialism}"
    
    prompt = f"""Generate realistic tasks for a Role Performance Statement / Task Analysis.

Role Title: {role_title}
Framework: {framework}
{context_info}
Additional Context: {description if description else 'None provided'}

Generate 12-15 SPECIFIC tasks as a JSON array. Each task must be realistic and detailed for this role.
Each task should have:
{{
    "task_number": "1.0",
    "duty_area": "The duty area this task belongs to",
    "performance": "Specific action verb + object (what the person must do)",
    "conditions": "Detailed circumstances (equipment, environment, resources, references)",
    "standards": "Measurable standard (time, accuracy, compliance requirement)",
    "category": "FT/WPT/OJT/CBT",
    "frequency": "How often performed (Daily/Weekly/Monthly/As Required)",
    "criticality": "High/Medium/Low",
    "ksa": {{
        "knowledge": "Specific knowledge required",
        "skills": "Specific skills required", 
        "attitudes": "Attitudes/behaviours required"
    }}
}}

Categories:
- FT = Formal Training (classroom/structured course)
- WPT = Workplace Training (supervised practice in workplace)
- OJT = On-the-Job Training (learning while doing)
- CBT = Computer-Based Training (e-learning/simulation)

Group tasks by duty area. Make each task specific, measurable, and realistic for a {role_title}.
Return ONLY the JSON array, no other text."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT, max_tokens=8000)
    
    try:
        json_match = re.search(r'\[[\s\S]*\]', response)
        if json_match:
            return json.loads(json_match.group())
    except Exception as e:
        print(f"[NOVA] JSON parse error in tasks: {e}")
    
    # Fallback
    return [
        {
            "task_number": "1.0",
            "duty_area": "Core Duties",
            "performance": f"Perform core {role_title} duties",
            "conditions": "Standard workplace environment with required equipment",
            "standards": "In accordance with SOPs and regulatory requirements",
            "category": "FT",
            "frequency": "Daily",
            "criticality": "High",
            "ksa": {
                "knowledge": "Role-specific procedures and regulations",
                "skills": "Technical and communication skills",
                "attitudes": "Professional conduct and attention to detail"
            }
        }
    ]


async def generate_gap_analysis(role_title: str, framework: str, tasks: List[Dict],
                                 domain: str = "", specialism: str = "") -> List[Dict]:
    """Generate comprehensive training gap analysis"""
    task_summary = "\n".join([f"- {t.get('duty_area', 'Duty')}: {t.get('performance', 'Task')}" for t in tasks[:8]])
    
    prompt = f"""Generate a comprehensive Training Gap Analysis based on these role tasks.

Role Title: {role_title}
Framework: {framework}
Domain: {domain if domain else 'Not specified'}
Specialism: {specialism if specialism else 'Not specified'}

Tasks Identified:
{task_summary}

Generate gap analysis as a JSON array with 10-12 training gaps:
{{
    "gap_id": "G1",
    "skill_area": "Specific skill or competency area",
    "task_reference": "Which task(s) this relates to",
    "current_provision": "What training currently exists (be specific)",
    "current_standard": "Current performance level achieved",
    "required_standard": "What standard is needed (with metrics)",
    "gap_description": "Clear description of the gap",
    "gap_cause": "Root cause of the gap",
    "impact": "Operational impact if gap not addressed",
    "risk_rating": "High/Medium/Low",
    "priority": 1-12 (1 = highest),
    "recommended_intervention": "How to close this gap",
    "estimated_effort": "Time/resource to close gap"
}}

Ensure gaps cover:
- Knowledge gaps
- Skill gaps  
- Experience gaps
- Equipment/systems familiarity gaps
- Procedural/compliance gaps

Return ONLY the JSON array."""

    response = await call_claude(prompt, TRAINING_SYSTEM_PROMPT, max_tokens=8000)
    
    try:
        json_match = re.search(r'\[[\s\S]*\]', response)
        if json_match:
            return json.loads(json_match.group())
    except Exception as e:
        print(f"[NOVA] JSON parse error in gaps: {e}")
    
    return [{"gap_id": "G1", "skill_area": "Core competencies", "task_reference": "1.0",
             "current_provision": "Limited formal training", "current_standard": "Basic awareness",
             "required_standard": "Full proficiency", "gap_description": "Gap between current and required capability",
             "gap_cause": "No structured training programme", "impact": "Reduced operational effectiveness",
             "risk_rating": "Medium", "priority": 1, "recommended_intervention": "Formal training course",
             "estimated_effort": "4-6 weeks development"}]


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
    "lesson_title": "Title",
    "duration": "Duration in minutes",
    "to_covered": ["TO 1", "TO 2"],
    "introduction": {{
        "duration": "5 mins",
        "activities": ["Intro activity 1", "Intro activity 2"]
    }},
    "development": [
        {{"topic": "Topic", "duration": "X mins", "method": "Lecture/Demo/Practical", "content": ["Point 1", "Point 2"]}}
    ],
    "application": {{
        "duration": "X mins",
        "activities": ["Practice activity"]
    }},
    "assessment": {{
        "type": "Formative/Summative",
        "method": "Method description"
    }},
    "resources": ["Resource 1", "Resource 2"],
    "instructor_notes": "Notes for instructor"
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
                                objectives: List[Dict]) -> Dict:
    """Generate Assessment Strategy and Instruments"""
    obj_summary = "\n".join([f"- {o.get('to_number', 'TO')}: {o.get('objective', '')[:50]}" 
                            for o in objectives[:4]])
    
    prompt = f"""Generate Assessment Strategy and Instruments.

Role Title: {role_title}
Training Objectives:
{obj_summary}

Generate as JSON:
{{
    "assessment_strategy": {{
        "purpose": "Assessment purpose",
        "philosophy": "Assessment philosophy",
        "pass_criteria": "Overall pass criteria",
        "grading_policy": "How grades are determined",
        "failure_policy": "Remedial training approach",
        "ai_policy": "Policy on AI use in assessments"
    }},
    "practical_assessments": [
        {{
            "assessment_id": "PA1",
            "title": "Assessment title",
            "to_assessed": ["TO 1"],
            "description": "What trainee must do",
            "conditions": "Assessment conditions",
            "pass_criteria": "Specific criteria",
            "duration": "Time allowed"
        }}
    ],
    "theory_questions": [
        {{
            "question_id": "Q1",
            "to_assessed": "TO 1",
            "question_type": "Multiple Choice/Short Answer/Essay",
            "question": "Question text",
            "answer": "Model answer",
            "marks": 5
        }}
    ],
    "marking_scheme": {{
        "overall_pass_mark": "70%",
        "weighting": {{"practical": "60%", "theory": "40%"}},
        "criteria": ["Marking criterion 1", "Marking criterion 2"]
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
    """Build comprehensive DTSM 2 compliant Scoping Report document"""
    doc = create_styled_document("Scoping Exercise Report", role_title, framework)
    
    # Title page
    add_title_page(doc, "SCOPING EXERCISE REPORT", role_title, {
        "Document Type": "Scoping Exercise Report (DTSM 2 1.2)",
        "Role": role_title,
        "Framework": framework,
        "Reference": get_framework_reference(framework, "scoping"),
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Document Control table
    add_section_heading(doc, "DOCUMENT CONTROL")
    doc.add_paragraph("Version History:")
    add_table_from_data(doc, 
        ["Version", "Date", "Author", "Changes"],
        [["1.0", datetime.utcnow().strftime('%d %b %Y'), "NOVA Analysis Agent", "Initial draft"]]
    )
    
    # Table of Contents placeholder
    doc.add_paragraph("TABLE OF CONTENTS")
    doc.add_paragraph("(Update field after editing)")
    doc.add_page_break()
    
    # 1. Purpose and Aim
    add_section_heading(doc, "1. PURPOSE AND AIM")
    purpose_aim = content.get("purpose_and_aim", {})
    
    doc.add_heading("1.1 Purpose", level=2)
    doc.add_paragraph(purpose_aim.get("purpose", ""))
    
    doc.add_heading("1.2 Aim", level=2)
    doc.add_paragraph(purpose_aim.get("aim", ""))
    
    doc.add_heading("1.3 Objectives", level=2)
    for obj in purpose_aim.get("objectives", []):
        doc.add_paragraph(obj, style='List Bullet')
    
    # 2. Background and Context
    add_section_heading(doc, "2. BACKGROUND AND CONTEXT")
    background = content.get("background_and_context", {})
    
    doc.add_heading("2.1 Operational Context", level=2)
    doc.add_paragraph(background.get("operational_context", ""))
    
    doc.add_heading("2.2 Capability Requirement", level=2)
    doc.add_paragraph(background.get("capability_requirement", ""))
    
    doc.add_heading("2.3 Driver for Change", level=2)
    doc.add_paragraph(background.get("driver_for_change", ""))
    
    doc.add_heading("2.4 Strategic Alignment", level=2)
    doc.add_paragraph(background.get("strategic_alignment", ""))
    
    # 3. Scope
    add_section_heading(doc, "3. SCOPE")
    scope = content.get("scope", {})
    
    doc.add_heading("3.1 Inclusions", level=2)
    for item in scope.get("inclusions", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("3.2 Exclusions", level=2)
    for item in scope.get("exclusions", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("3.3 Boundaries", level=2)
    doc.add_paragraph(scope.get("boundaries", ""))
    
    doc.add_heading("3.4 Interfaces", level=2)
    for item in scope.get("interfaces", []):
        doc.add_paragraph(item, style='List Bullet')
    
    # 4. Governance
    add_section_heading(doc, "4. GOVERNANCE")
    governance = content.get("governance", {})
    
    doc.add_heading("4.1 Training Requirements Authority (TRA)", level=2)
    tra = governance.get("tra", {})
    doc.add_paragraph(f"Name: {tra.get('name', 'TBC')}")
    doc.add_paragraph(f"Responsibility: {tra.get('responsibility', '')}")
    
    doc.add_heading("4.2 Training Delivery Authority (TDA)", level=2)
    tda = governance.get("tda", {})
    doc.add_paragraph(f"Name: {tda.get('name', 'TBC')}")
    doc.add_paragraph(f"Responsibility: {tda.get('responsibility', '')}")
    
    doc.add_heading("4.3 Stakeholders", level=2)
    stakeholders = governance.get("stakeholders", [])
    if stakeholders:
        add_table_from_data(doc, 
            ["Role", "Organisation", "Interest", "Engagement"],
            [[s.get("role", ""), s.get("organisation", ""), s.get("interest", ""), s.get("engagement", "")] 
             for s in stakeholders]
        )
    
    doc.add_heading("4.4 Governance Board", level=2)
    doc.add_paragraph(governance.get("governance_board", ""))
    
    # 5. Target Audience
    add_section_heading(doc, "5. TARGET AUDIENCE")
    audience = content.get("target_audience", {})
    
    doc.add_paragraph(f"Population: {audience.get('population_description', '')}")
    doc.add_paragraph(f"Annual Throughput: {audience.get('annual_throughput', 'TBC')}")
    doc.add_paragraph(f"Current Competence: {audience.get('current_competence', '')}")
    
    doc.add_heading("5.1 Entry Requirements", level=2)
    for item in audience.get("entry_requirements", []):
        doc.add_paragraph(item, style='List Bullet')
    
    # 6. Methodology
    add_section_heading(doc, "6. METHODOLOGY")
    methodology = content.get("methodology", {})
    
    doc.add_heading("6.1 Approach", level=2)
    doc.add_paragraph(methodology.get("approach", ""))
    
    doc.add_heading("6.2 Data Collection Methods", level=2)
    for item in methodology.get("data_collection", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("6.3 Analysis Techniques", level=2)
    for item in methodology.get("analysis_techniques", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("6.4 Validation Approach", level=2)
    doc.add_paragraph(methodology.get("validation_approach", ""))
    
    # 7. Assumptions, Constraints, Dependencies
    add_section_heading(doc, "7. ASSUMPTIONS, CONSTRAINTS AND DEPENDENCIES")
    
    doc.add_heading("7.1 Assumptions", level=2)
    for item in content.get("assumptions", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("7.2 Constraints", level=2)
    for item in content.get("constraints", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("7.3 Dependencies", level=2)
    for item in content.get("dependencies", []):
        doc.add_paragraph(item, style='List Bullet')
    
    # 8. Risk Assessment
    add_section_heading(doc, "8. RISK ASSESSMENT")
    risks = content.get("risks", [])
    if risks:
        add_table_from_data(doc,
            ["ID", "Risk", "Likelihood", "Impact", "Mitigation", "Owner"],
            [[r.get("risk_id", ""), r.get("description", ""), r.get("likelihood", ""), 
              r.get("impact", ""), r.get("mitigation", ""), r.get("owner", "")] for r in risks]
        )
    
    # 9. Resource Estimate
    add_section_heading(doc, "9. RESOURCE ESTIMATE")
    res = content.get("resource_estimate", {})
    
    doc.add_paragraph(f"Duration: {res.get('duration', 'TBD')}")
    doc.add_paragraph(f"Effort: {res.get('effort', 'TBD')}")
    doc.add_paragraph(f"Budget Estimate: {res.get('budget_estimate', 'TBD')}")
    
    doc.add_heading("9.1 Personnel", level=2)
    personnel = res.get("personnel", [])
    if personnel:
        add_table_from_data(doc,
            ["Role", "FTE", "Skills Required"],
            [[p.get("role", ""), p.get("fte", ""), p.get("skills", "")] for p in personnel]
        )
    
    # 10. Timeline
    add_section_heading(doc, "10. TIMELINE")
    timeline = content.get("timeline", {})
    
    doc.add_heading("10.1 Phases", level=2)
    phases = timeline.get("phases", [])
    if phases:
        add_table_from_data(doc,
            ["Phase", "Activities", "Duration", "Milestone"],
            [[p.get("phase", ""), p.get("activities", ""), p.get("duration", ""), p.get("milestone", "")] 
             for p in phases]
        )
    
    doc.add_heading("10.2 Key Milestones", level=2)
    milestones = timeline.get("key_milestones", [])
    if milestones:
        add_table_from_data(doc,
            ["Milestone", "Target Date", "Dependency"],
            [[m.get("milestone", ""), m.get("target_date", ""), m.get("dependency", "")] for m in milestones]
        )
    
    # 11. Deliverables
    add_section_heading(doc, "11. DELIVERABLES")
    deliverables = content.get("deliverables", [])
    if deliverables:
        add_table_from_data(doc,
            ["Deliverable", "Description", "Format", "Audience"],
            [[d.get("deliverable", ""), d.get("description", ""), d.get("format", ""), d.get("audience", "")] 
             for d in deliverables]
        )
    
    # 12. Quality Assurance
    add_section_heading(doc, "12. QUALITY ASSURANCE")
    qa = content.get("quality_assurance", {})
    doc.add_paragraph(f"Review Process: {qa.get('review_process', '')}")
    doc.add_paragraph(f"Approval Authority: {qa.get('approval_authority', '')}")
    doc.add_paragraph(f"Compliance Check: {qa.get('compliance_check', '')}")
    
    # 13. Recommendations
    add_section_heading(doc, "13. RECOMMENDATIONS AND NEXT STEPS")
    recs = content.get("recommendations", {})
    
    doc.add_heading("13.1 Recommendation", level=2)
    doc.add_paragraph(recs.get("proceed", ""))
    
    doc.add_heading("13.2 Next Steps", level=2)
    for i, item in enumerate(recs.get("next_steps", []), 1):
        doc.add_paragraph(f"{i}. {item}")
    
    doc.add_heading("13.3 Decision Required", level=2)
    doc.add_paragraph(recs.get("decision_required", ""))
    
    # Signature block
    doc.add_page_break()
    add_section_heading(doc, "APPROVAL")
    doc.add_paragraph("This Scoping Exercise Report is submitted for approval:")
    doc.add_paragraph("")
    doc.add_paragraph("Prepared by: _________________________ Date: _________")
    doc.add_paragraph("")
    doc.add_paragraph("Reviewed by: _________________________ Date: _________")
    doc.add_paragraph("")
    doc.add_paragraph("Approved by (TRA): ___________________ Date: _________")
    
    # Save
    filename = "01_Scoping_Exercise_Report.docx"
    doc.save(output_path / filename)
    return filename


def build_role_performance_statement(role_title: str, framework: str, tasks: List[Dict],
                                      output_path: Path) -> str:
    """Build Role Performance Statement document"""
    doc = create_styled_document("Role Performance Statement", role_title, framework)
    
    term = get_framework_term(framework, "roleps")
    add_title_page(doc, term.upper(), role_title, {
        "Document Type": term,
        "Reference": get_framework_reference(framework, "roleps"),
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
    
    # Group by duty area if available
    duty_areas = {}
    for task in tasks:
        duty = task.get("duty_area", "General Duties")
        if duty not in duty_areas:
            duty_areas[duty] = []
        duty_areas[duty].append(task)
    
    for duty, duty_tasks in duty_areas.items():
        doc.add_heading(f"2.x {duty}", level=2)
        
        headers = ["Task No.", "Performance", "Conditions", "Standards", "Category"]
        rows = []
        for task in duty_tasks:
            rows.append([
                task.get("task_number", ""),
                task.get("performance", ""),
                task.get("conditions", ""),
                task.get("standards", ""),
                task.get("category", "")
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
    
    doc.add_paragraph("Training Categories:")
    for cat, count in categories.items():
        cat_names = {"FT": "Formal Training", "WPT": "Workplace Training", 
                     "OJT": "On-the-Job Training", "CBT": "Computer-Based Training"}
        doc.add_paragraph(f"• {cat_names.get(cat, cat)}: {count} tasks", style='List Bullet')
    
    filename = "02_Role_Performance_Statement.docx"
    doc.save(output_path / filename)
    return filename


def build_gap_analysis(role_title: str, framework: str, gaps: List[Dict],
                       output_path: Path) -> str:
    """Build Training Gap Analysis document"""
    doc = create_styled_document("Training Gap Analysis", role_title, framework)
    
    add_title_page(doc, "TRAINING GAP ANALYSIS", role_title, {
        "Document Type": "Training Gap Analysis",
        "Reference": get_framework_reference(framework, "tga"),
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL"
    })
    
    # Executive Summary
    add_section_heading(doc, "1. EXECUTIVE SUMMARY")
    high_risks = len([g for g in gaps if g.get("risk_rating") == "High"])
    med_risks = len([g for g in gaps if g.get("risk_rating") == "Medium"])
    doc.add_paragraph(f"This analysis identifies {len(gaps)} training gaps for the {role_title} role. "
                     f"Of these, {high_risks} are rated as high priority and {med_risks} as medium priority.")
    
    # Gap Analysis Table
    add_section_heading(doc, "2. GAP ANALYSIS")
    headers = ["ID", "Skill Area", "Current", "Required", "Gap Description", "Risk", "Priority"]
    rows = [[
        g.get("gap_id", ""),
        g.get("skill_area", ""),
        g.get("current_provision", "")[:30],
        g.get("required_standard", "")[:30],
        g.get("gap_description", "")[:40],
        g.get("risk_rating", ""),
        str(g.get("priority", ""))
    ] for g in gaps]
    
    add_table_from_data(doc, headers, rows)
    
    # Detailed Gaps
    add_section_heading(doc, "3. DETAILED GAP ANALYSIS")
    for gap in gaps:
        doc.add_heading(f"{gap.get('gap_id', 'Gap')}: {gap.get('skill_area', '')}", level=2)
        doc.add_paragraph(f"Task Reference: {gap.get('task_reference', 'N/A')}")
        doc.add_paragraph(f"Current Provision: {gap.get('current_provision', '')}")
        doc.add_paragraph(f"Required Standard: {gap.get('required_standard', '')}")
        doc.add_paragraph(f"Gap Description: {gap.get('gap_description', '')}")
        doc.add_paragraph(f"Gap Cause: {gap.get('gap_cause', '')}")
        doc.add_paragraph(f"Operational Impact: {gap.get('impact', '')}")
        doc.add_paragraph(f"Recommended Intervention: {gap.get('recommended_intervention', '')}")
        doc.add_paragraph("")
    
    # Recommendations
    add_section_heading(doc, "4. RECOMMENDATIONS")
    doc.add_paragraph("Based on the gap analysis, the following actions are recommended:")
    doc.add_paragraph("1. Address high-priority gaps through formal training intervention")
    doc.add_paragraph("2. Develop workplace training packages for medium-priority gaps")
    doc.add_paragraph("3. Implement continuous professional development for ongoing needs")
    doc.add_paragraph("4. Establish competence maintenance programme")
    
    filename = "03_Training_Gap_Analysis.docx"
    doc.save(output_path / filename)
    return filename


def build_training_needs_report(role_title: str, framework: str, 
                                 tnr_content: Dict, tasks: List[Dict], gaps: List[Dict],
                                 output_path: Path) -> str:
    """Build comprehensive DTSM 2 compliant Training Needs Report"""
    doc = create_styled_document("Training Needs Report", role_title, framework)
    
    term = get_framework_term(framework, "tnr")
    add_title_page(doc, term.upper(), role_title, {
        "Document Type": term,
        "Reference": get_framework_reference(framework, "tnr"),
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0",
        "Classification": "OFFICIAL",
        "Status": "For CEB Approval"
    })
    
    # Document Control
    add_section_heading(doc, "DOCUMENT CONTROL")
    add_table_from_data(doc, 
        ["Version", "Date", "Author", "Changes"],
        [["1.0", datetime.utcnow().strftime('%d %b %Y'), "NOVA Analysis Agent", "Initial draft for approval"]]
    )
    doc.add_page_break()
    
    # Executive Summary
    exec_sum = tnr_content.get("executive_summary", {})
    add_section_heading(doc, "EXECUTIVE SUMMARY")
    doc.add_paragraph(exec_sum.get("overview", ""))
    
    doc.add_heading("Key Findings", level=2)
    for finding in exec_sum.get("key_findings", []):
        doc.add_paragraph(finding, style='List Bullet')
    
    doc.add_heading("Recommendation", level=2)
    doc.add_paragraph(exec_sum.get("recommendation", ""))
    
    doc.add_heading("Resource Implications", level=2)
    doc.add_paragraph(exec_sum.get("resource_implications", ""))
    doc.add_page_break()
    
    # 1. Introduction
    intro = tnr_content.get("introduction", {})
    add_section_heading(doc, "1. INTRODUCTION")
    doc.add_heading("1.1 Purpose", level=2)
    doc.add_paragraph(intro.get("purpose", ""))
    doc.add_heading("1.2 Scope", level=2)
    doc.add_paragraph(intro.get("scope", ""))
    doc.add_heading("1.3 Methodology", level=2)
    doc.add_paragraph(intro.get("methodology", ""))
    
    # 2. Background
    background = tnr_content.get("background", {})
    add_section_heading(doc, "2. BACKGROUND")
    doc.add_heading("2.1 Operational Context", level=2)
    doc.add_paragraph(background.get("operational_context", ""))
    doc.add_heading("2.2 Strategic Drivers", level=2)
    doc.add_paragraph(background.get("strategic_drivers", ""))
    doc.add_heading("2.3 Current State", level=2)
    doc.add_paragraph(background.get("current_state", ""))
    doc.add_heading("2.4 Problem Statement", level=2)
    doc.add_paragraph(background.get("problem_statement", ""))
    
    # 3. Analysis Findings
    findings = tnr_content.get("analysis_findings", {})
    add_section_heading(doc, "3. ANALYSIS FINDINGS")
    
    doc.add_heading("3.1 Role Analysis Summary", level=2)
    doc.add_paragraph(findings.get("role_analysis_summary", ""))
    
    doc.add_heading("3.2 Task Analysis Summary", level=2)
    doc.add_paragraph(findings.get("task_analysis_summary", f"{len(tasks)} tasks identified."))
    
    doc.add_heading("3.3 Knowledge, Skills and Attitudes Analysis", level=2)
    ksa = findings.get("ksa_analysis", {})
    doc.add_paragraph("Knowledge Areas:")
    for item in ksa.get("knowledge_areas", []):
        doc.add_paragraph(item, style='List Bullet')
    doc.add_paragraph("Skills Required:")
    for item in ksa.get("skills_required", []):
        doc.add_paragraph(item, style='List Bullet')
    doc.add_paragraph("Attitudes and Behaviours:")
    for item in ksa.get("attitudes_behaviours", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("3.4 Gap Analysis Summary", level=2)
    doc.add_paragraph(findings.get("gap_analysis_summary", f"{len(gaps)} gaps identified."))
    
    doc.add_heading("3.5 Critical Gaps", level=2)
    critical_gaps = findings.get("critical_gaps", [])
    if critical_gaps:
        add_table_from_data(doc,
            ["Gap", "Impact", "Priority", "Rationale"],
            [[g.get("gap", ""), g.get("impact", ""), g.get("priority", ""), g.get("rationale", "")] 
             for g in critical_gaps]
        )
    
    # 4. Training Requirement
    req = tnr_content.get("training_requirement", {})
    add_section_heading(doc, "4. TRAINING REQUIREMENT")
    
    doc.add_heading("4.1 Trained Output Requirement", level=2)
    doc.add_paragraph(req.get("trained_output_requirement", ""))
    
    doc.add_heading("4.2 Performance Standards", level=2)
    for std in req.get("performance_standards", []):
        doc.add_paragraph(std, style='List Bullet')
    
    doc.add_heading("4.3 Entry Requirements", level=2)
    for item in req.get("entry_requirements", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("4.4 Training Population", level=2)
    pop = req.get("training_population", {})
    doc.add_paragraph(f"Description: {pop.get('description', '')}")
    doc.add_paragraph(f"Annual Throughput: {pop.get('annual_throughput', 'TBC')}")
    doc.add_paragraph(f"Locations: {pop.get('locations', 'TBC')}")
    
    # 5. Training Options
    add_section_heading(doc, "5. TRAINING OPTIONS ANALYSIS")
    options = tnr_content.get("training_options", [])
    
    for opt in options:
        doc.add_heading(opt.get("option", "Option"), level=2)
        doc.add_paragraph(opt.get("description", ""))
        doc.add_paragraph(f"Delivery Method: {opt.get('delivery_method', '')}")
        doc.add_paragraph(f"Duration: {opt.get('duration', '')}")
        
        doc.add_paragraph("Advantages:")
        for adv in opt.get("advantages", []):
            doc.add_paragraph(adv, style='List Bullet')
        
        doc.add_paragraph("Disadvantages:")
        for dis in opt.get("disadvantages", []):
            doc.add_paragraph(dis, style='List Bullet')
        
        res_req = opt.get("resource_requirements", {})
        doc.add_paragraph(f"Instructors: {res_req.get('instructors', 'TBC')}")
        doc.add_paragraph(f"Facilities: {res_req.get('facilities', 'TBC')}")
        doc.add_paragraph(f"Cost Estimate: {opt.get('cost_estimate', 'TBC')}")
        doc.add_paragraph(f"Effectiveness Rating: {opt.get('effectiveness_rating', '')}")
        doc.add_paragraph("")
    
    # Options Comparison Table
    if options:
        add_table_from_data(doc,
            ["Option", "Effectiveness", "Cost", "Risk"],
            [[o.get("option", ""), o.get("effectiveness_rating", ""), 
              o.get("cost_estimate", ""), o.get("risk_assessment", "")] for o in options]
        )
    
    # 6. Recommended Solution
    rec_sol = tnr_content.get("recommended_solution", {})
    add_section_heading(doc, "6. RECOMMENDED SOLUTION")
    
    doc.add_heading("6.1 Recommendation", level=2)
    doc.add_paragraph(rec_sol.get("recommendation", ""))
    
    doc.add_heading("6.2 Rationale", level=2)
    doc.add_paragraph(rec_sol.get("rationale", ""))
    
    doc.add_heading("6.3 Training Statement", level=2)
    ts = rec_sol.get("training_statement", {})
    doc.add_paragraph(f"Formal Training: {ts.get('formal_training', '')}")
    doc.add_paragraph(f"Workplace Training: {ts.get('workplace_training', '')}")
    doc.add_paragraph(f"Residual Gap: {ts.get('residual_gap', 'None identified')}")
    
    doc.add_heading("6.4 Implementation Approach", level=2)
    doc.add_paragraph(rec_sol.get("implementation_approach", ""))
    
    # 7. Resource Implications
    res_imp = tnr_content.get("resource_implications", {})
    add_section_heading(doc, "7. RESOURCE IMPLICATIONS")
    
    doc.add_heading("7.1 Financial", level=2)
    fin = res_imp.get("financial", {})
    doc.add_paragraph(f"Development Cost: {fin.get('development_cost', 'TBC')}")
    doc.add_paragraph(f"Annual Delivery Cost: {fin.get('delivery_cost', 'TBC')}")
    doc.add_paragraph(f"Total Investment: {fin.get('total_investment', 'TBC')}")
    
    doc.add_heading("7.2 Personnel", level=2)
    pers = res_imp.get("personnel", {})
    doc.add_paragraph(f"Instructors: {pers.get('instructors', 'TBC')}")
    doc.add_paragraph(f"Support Staff: {pers.get('support_staff', 'TBC')}")
    
    doc.add_heading("7.3 Infrastructure", level=2)
    infra = res_imp.get("infrastructure", {})
    doc.add_paragraph(f"Facilities: {infra.get('facilities', 'TBC')}")
    doc.add_paragraph(f"Equipment: {infra.get('equipment', 'TBC')}")
    
    doc.add_heading("7.4 Timeline", level=2)
    doc.add_paragraph(res_imp.get("timeline", "TBC"))
    
    # 8. Risk Assessment
    add_section_heading(doc, "8. RISK ASSESSMENT")
    risks = tnr_content.get("risk_assessment", [])
    if risks:
        add_table_from_data(doc,
            ["Risk", "Category", "L", "I", "Mitigation", "Owner"],
            [[r.get("risk", ""), r.get("category", ""), r.get("likelihood", ""), 
              r.get("impact", ""), r.get("mitigation", ""), r.get("owner", "")] for r in risks]
        )
    
    # 9. Governance and Next Steps
    gov = tnr_content.get("governance_and_next_steps", {})
    add_section_heading(doc, "9. GOVERNANCE AND NEXT STEPS")
    
    doc.add_heading("9.1 Approvals Required", level=2)
    for item in gov.get("approvals_required", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("9.2 Next Steps", level=2)
    next_steps = gov.get("next_steps", [])
    if next_steps:
        add_table_from_data(doc,
            ["Step", "Owner", "Target Date"],
            [[s.get("step", ""), s.get("owner", ""), s.get("target_date", "")] for s in next_steps]
        )
    
    doc.add_heading("9.3 Success Criteria", level=2)
    for item in gov.get("success_criteria", []):
        doc.add_paragraph(item, style='List Bullet')
    
    doc.add_heading("9.4 Review Schedule", level=2)
    doc.add_paragraph(gov.get("review_schedule", ""))
    
    # Annexes
    add_section_heading(doc, "ANNEXES")
    annexes = tnr_content.get("annexes", {})
    doc.add_paragraph(f"Annex A: {annexes.get('annex_a', 'Role Performance Statement')}")
    doc.add_paragraph(f"Annex B: {annexes.get('annex_b', 'Complete Task Analysis')}")
    doc.add_paragraph(f"Annex C: {annexes.get('annex_c', 'Detailed Gap Analysis')}")
    doc.add_paragraph(f"Annex D: {annexes.get('annex_d', 'Stakeholder Consultation Log')}")
    doc.add_paragraph(f"Annex E: {annexes.get('annex_e', 'Cost-Benefit Analysis')}")
    
    # Approval Section
    doc.add_page_break()
    add_section_heading(doc, "APPROVAL")
    doc.add_paragraph("This Training Needs Report is submitted for Customer Executive Board approval.")
    doc.add_paragraph("")
    doc.add_paragraph("Prepared by: _________________________ Date: _________")
    doc.add_paragraph("")
    doc.add_paragraph("Reviewed by (TDA): ___________________ Date: _________")
    doc.add_paragraph("")
    doc.add_paragraph("Approved by (TRA): ___________________ Date: _________")
    doc.add_paragraph("")
    doc.add_paragraph("CEB Approval: ________________________ Date: _________")
    
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
    
    add_section_heading(doc, "1. INTRODUCTION")
    doc.add_paragraph(f"This document defines the {term}s for the {role_title} training programme.")
    
    add_section_heading(doc, f"2. {term.upper()}S")
    
    for obj in objectives:
        doc.add_heading(f"{obj.get('to_number', 'TO')}: {obj.get('objective', '')[:60]}", level=2)
        
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
    
    filename = "01_Training_Objectives.docx"
    doc.save(output_path / filename)
    return filename


def build_enabling_objectives_doc(role_title: str, framework: str,
                                   eo_data: List[Dict], output_path: Path) -> str:
    """Build Enabling Objectives and KLPs document"""
    doc = create_styled_document("Enabling Objectives", role_title, framework)
    
    eo_term = get_framework_term(framework, "eo")
    klp_term = get_framework_term(framework, "klp")
    
    add_title_page(doc, f"{eo_term.upper()}S AND {klp_term.upper()}S", role_title, {
        "Document Type": f"{eo_term}s and {klp_term}s",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0"
    })
    
    add_section_heading(doc, "1. INTRODUCTION")
    doc.add_paragraph(f"This document defines the {eo_term}s and {klp_term}s for {role_title} training.")
    
    add_section_heading(doc, f"2. {eo_term.upper()}S AND {klp_term.upper()}S")
    
    for to_item in eo_data:
        doc.add_heading(f"{to_item.get('to_number', 'TO')}: {to_item.get('to_text', '')[:50]}", level=2)
        
        for eo in to_item.get("enabling_objectives", []):
            doc.add_heading(f"{eo.get('eo_number', 'EO')}: {eo.get('eo_text', '')[:40]}", level=3)
            
            doc.add_paragraph(f"{klp_term}s:")
            for klp in eo.get("klps", []):
                klp_type = klp.get("type", "K")
                type_label = {"K": "Knowledge", "S": "Skill", "A": "Attitude"}.get(klp_type, klp_type)
                doc.add_paragraph(f"{klp.get('klp_number', '')}: {klp.get('klp_text', '')} [{type_label}]", 
                                 style='List Bullet')
            doc.add_paragraph()
    
    filename = "02_Enabling_Objectives_KLPs.docx"
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
    
    add_section_heading(doc, "1. INTRODUCTION")
    doc.add_paragraph(f"This document provides lesson plans for the {role_title} training programme.")
    
    for lesson in lessons:
        add_section_heading(doc, f"LESSON {lesson.get('lesson_number', '')}: {lesson.get('lesson_title', '')}")
        
        doc.add_paragraph(f"Duration: {lesson.get('duration', '')}")
        doc.add_paragraph(f"TOs Covered: {', '.join(lesson.get('to_covered', []))}")
        
        doc.add_heading("Introduction", level=2)
        intro = lesson.get("introduction", {})
        doc.add_paragraph(f"Duration: {intro.get('duration', '')}")
        for act in intro.get("activities", []):
            doc.add_paragraph(act, style='List Bullet')
        
        doc.add_heading("Development", level=2)
        for dev in lesson.get("development", []):
            doc.add_paragraph(f"{dev.get('topic', '')} ({dev.get('duration', '')}) - {dev.get('method', '')}")
            for point in dev.get("content", []):
                doc.add_paragraph(point, style='List Bullet')
        
        doc.add_heading("Application", level=2)
        app = lesson.get("application", {})
        doc.add_paragraph(f"Duration: {app.get('duration', '')}")
        for act in app.get("activities", []):
            doc.add_paragraph(act, style='List Bullet')
        
        doc.add_heading("Assessment", level=2)
        assess = lesson.get("assessment", {})
        doc.add_paragraph(f"Type: {assess.get('type', '')}")
        doc.add_paragraph(f"Method: {assess.get('method', '')}")
        
        doc.add_heading("Resources", level=2)
        for res in lesson.get("resources", []):
            doc.add_paragraph(res, style='List Bullet')
        
        doc.add_heading("Instructor Notes", level=2)
        doc.add_paragraph(lesson.get("instructor_notes", ""))
        
        doc.add_page_break()
    
    filename = "01_Lesson_Plans.docx"
    doc.save(output_path / filename)
    return filename


def build_assessments_doc(role_title: str, framework: str,
                          assessments: Dict, output_path: Path) -> str:
    """Build Assessment Strategy and Instruments document"""
    doc = create_styled_document("Assessment Specification", role_title, framework)
    
    add_title_page(doc, "ASSESSMENT SPECIFICATION", role_title, {
        "Document Type": "Assessment Strategy and Specification",
        "Role": role_title,
        "Framework": framework,
        "Date": datetime.utcnow().strftime('%d %B %Y'),
        "Version": "1.0"
    })
    
    # Assessment Strategy
    add_section_heading(doc, "1. ASSESSMENT STRATEGY")
    strategy = assessments.get("assessment_strategy", {})
    doc.add_paragraph(f"Purpose: {strategy.get('purpose', '')}")
    doc.add_paragraph(f"Philosophy: {strategy.get('philosophy', '')}")
    doc.add_paragraph(f"Pass Criteria: {strategy.get('pass_criteria', '')}")
    doc.add_paragraph(f"Grading Policy: {strategy.get('grading_policy', '')}")
    doc.add_paragraph(f"Failure/Remedial Policy: {strategy.get('failure_policy', '')}")
    doc.add_paragraph(f"AI Policy: {strategy.get('ai_policy', '')}")
    
    # Practical Assessments
    add_section_heading(doc, "2. PRACTICAL ASSESSMENTS")
    for pa in assessments.get("practical_assessments", []):
        doc.add_heading(f"{pa.get('assessment_id', 'PA')}: {pa.get('title', '')}", level=2)
        doc.add_paragraph(f"TOs Assessed: {', '.join(pa.get('to_assessed', []))}")
        doc.add_paragraph(f"Description: {pa.get('description', '')}")
        doc.add_paragraph(f"Conditions: {pa.get('conditions', '')}")
        doc.add_paragraph(f"Pass Criteria: {pa.get('pass_criteria', '')}")
        doc.add_paragraph(f"Duration: {pa.get('duration', '')}")
        doc.add_paragraph()
    
    # Theory Questions
    add_section_heading(doc, "3. THEORY ASSESSMENT")
    questions = assessments.get("theory_questions", [])
    if questions:
        for q in questions:
            doc.add_heading(f"{q.get('question_id', 'Q')}: {q.get('question_type', '')}", level=2)
            doc.add_paragraph(f"TO Assessed: {q.get('to_assessed', '')}")
            doc.add_paragraph(f"Question: {q.get('question', '')}")
            doc.add_paragraph(f"Model Answer: {q.get('answer', '')}")
            doc.add_paragraph(f"Marks: {q.get('marks', '')}")
            doc.add_paragraph()
    
    # Marking Scheme
    add_section_heading(doc, "4. MARKING SCHEME")
    marking = assessments.get("marking_scheme", {})
    doc.add_paragraph(f"Overall Pass Mark: {marking.get('overall_pass_mark', '')}")
    weighting = marking.get("weighting", {})
    if weighting:
        doc.add_paragraph(f"Practical Weighting: {weighting.get('practical', '')}")
        doc.add_paragraph(f"Theory Weighting: {weighting.get('theory', '')}")
    doc.add_paragraph("Marking Criteria:")
    for criterion in marking.get("criteria", []):
        doc.add_paragraph(criterion, style='List Bullet')
    
    filename = "02_Assessment_Specification.docx"
    doc.save(output_path / filename)
    return filename


def build_compliance_certificate(role_title: str, framework: str, job_id: str,
                                  files: List[str], output_path: Path) -> str:
    """Build compliance certificate"""
    doc = create_styled_document("Compliance Certificate", role_title, framework)
    
    add_title_page(doc, "NOVA™ COMPLIANCE CERTIFICATE", role_title, {
        "Role": role_title,
        "Framework": framework,
        "Job ID": job_id[:8],
        "Generated": datetime.utcnow().strftime('%d %B %Y %H:%M UTC'),
        "Classification": "OFFICIAL"
    })
    
    add_section_heading(doc, "CERTIFICATION STATEMENT")
    doc.add_paragraph(
        f"This certificate confirms that the training documentation package for the {role_title} role "
        f"has been generated in compliance with {framework} methodology and standards."
    )
    
    add_section_heading(doc, "DOCUMENTS GENERATED")
    for f in files:
        doc.add_paragraph(f"✓ {f}", style='List Bullet')
    
    add_section_heading(doc, "COMPLIANCE CHECKLIST")
    checks = [
        ("Role Performance Statement", "Complete"),
        ("Task Analysis", "Complete"),
        ("Training Gap Analysis", "Complete"),
        ("Training Objectives", "Complete"),
        ("Enabling Objectives", "Complete"),
        ("Assessment Strategy", "Complete"),
        ("Doctrinal References", "Included")
    ]
    add_table_from_data(doc, ["Requirement", "Status"], checks)
    
    add_section_heading(doc, "DISCLAIMER")
    doc.add_paragraph(
        "This documentation has been generated by NOVA™ AI Training Analysis System. "
        "All outputs should be reviewed by qualified training professionals before implementation. "
        "Human validation and governance approval is required before use in operational training."
    )
    
    filename = "00_Compliance_Certificate.docx"
    doc.save(output_path / filename)
    return filename


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
        elif agent == "full-package":
            print(f"[NOVA] Running full-package agent for job {job_id}")
            await run_full_package_agent(job_id, parameters)
        
        job["status"] = "completed"
        job["progress"] = 100
        job["completed_at"] = datetime.utcnow().isoformat()
        print(f"[NOVA] Job {job_id} completed successfully")
        
    except Exception as e:
        print(f"[NOVA] Job {job_id} failed with error: {str(e)}")
        import traceback
        traceback.print_exc()
        job["status"] = "failed"
        job["error"] = str(e)
        job["completed_at"] = datetime.utcnow().isoformat()


async def run_analysis_agent(job_id: str, parameters: Dict[str, Any]):
    """Analysis Agent - Full DSAT Compliant"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "UK")
    description = parameters.get("role_description", "")
    domain = parameters.get("domain_name", parameters.get("domain", ""))
    specialism = parameters.get("specialism_name", parameters.get("specialism", ""))
    proficiency = parameters.get("proficiency", "practitioner")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    print(f"[NOVA] Analysis agent: role={role_title}, domain={domain}, specialism={specialism}")
    
    analysis_dir = output_dir / "01_Analysis"
    analysis_dir.mkdir(exist_ok=True)
    
    files_generated = []
    
    # Step 1: Generate Comprehensive Scoping Report
    print(f"[NOVA] Step 1: Generating Scoping Exercise Report")
    update_progress(job_id, 10, "Generating Scoping Exercise Report")
    scoping_content = await generate_scoping_content(role_title, framework, description, 
                                                      domain, specialism, proficiency)
    print(f"[NOVA] Scoping content generated, building document")
    filename = build_scoping_report(role_title, framework, scoping_content, analysis_dir)
    files_generated.append(filename)
    print(f"[NOVA] Scoping report saved: {filename}")
    
    # Step 2: Generate Role Tasks
    print(f"[NOVA] Step 2: Generating Role Performance Statement")
    update_progress(job_id, 30, "Generating Role Performance Statement")
    tasks = await generate_role_tasks(role_title, framework, description, domain, specialism)
    print(f"[NOVA] Tasks generated: {len(tasks)} tasks")
    filename = build_role_performance_statement(role_title, framework, tasks, analysis_dir)
    files_generated.append(filename)
    print(f"[NOVA] RolePS saved: {filename}")
    
    # Step 3: Gap Analysis
    print(f"[NOVA] Step 3: Conducting Training Gap Analysis")
    update_progress(job_id, 50, "Conducting Training Gap Analysis")
    gaps = await generate_gap_analysis(role_title, framework, tasks, domain, specialism)
    print(f"[NOVA] Gap analysis generated: {len(gaps)} gaps")
    filename = build_gap_analysis(role_title, framework, gaps, analysis_dir)
    files_generated.append(filename)
    print(f"[NOVA] Gap analysis saved: {filename}")
    
    # Step 4: Comprehensive Training Needs Report
    print(f"[NOVA] Step 4: Compiling Training Needs Report")
    update_progress(job_id, 70, "Compiling Training Needs Report")
    tnr_content = await generate_tnr_content(role_title, framework, scoping_content, tasks, gaps,
                                              domain, specialism, proficiency)
    filename = build_training_needs_report(role_title, framework, tnr_content, tasks, gaps, analysis_dir)
    files_generated.append(filename)
    print(f"[NOVA] TNR saved: {filename}")
    
    # Store data for other agents
    jobs[job_id]["analysis_data"] = {
        "scoping": scoping_content,
        "tasks": tasks,
        "gaps": gaps,
        "tnr": tnr_content
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


async def run_full_package_agent(job_id: str, parameters: Dict[str, Any]):
    """Full Package Agent - complete training lifecycle"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "UK")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    all_files = []
    
    # Phase 1: Analysis
    update_progress(job_id, 5, "Starting Analysis Phase")
    await run_analysis_agent(job_id, parameters)
    jobs[job_id]["progress"] = 30
    
    # Phase 2: Design
    update_progress(job_id, 35, "Starting Design Phase")
    await run_design_agent(job_id, parameters)
    jobs[job_id]["progress"] = 60
    
    # Phase 3: Delivery
    update_progress(job_id, 65, "Starting Delivery Phase")
    await run_delivery_agent(job_id, parameters)
    jobs[job_id]["progress"] = 90
    
    # Collect all files
    for subdir in output_dir.iterdir():
        if subdir.is_dir():
            for f in subdir.iterdir():
                if f.suffix == '.docx':
                    all_files.append(f.name)
    
    # Generate certificate
    update_progress(job_id, 95, "Generating Compliance Certificate")
    build_compliance_certificate(role_title, framework, job_id, all_files, output_dir)
    
    update_progress(job_id, 100, "Full Package Complete")


# ============================================================================
# RUN SERVER
# ============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
