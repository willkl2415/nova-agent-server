"""
NOVA Agent Server v4.1
Autonomous Training Agent Execution Server

Changes from v4.0:
- Removed Scoping Report and Gap Analysis (Analysis Agent now produces 2 docs)
- RolePS in landscape with proper template layout
- TNR with proper styling (Roboto, 6pt spacing, 1.5 line height)
- Professional table styling with colored headers
- Fixed progress tracking

Endpoints:
- POST /api/execute - Start agent task
- GET /api/status/{job_id} - Get task status
- GET /api/download/{job_id} - Download ZIP
- GET /api/health - Health check
"""

import os
import json
import asyncio
import re
from datetime import datetime
from typing import Optional, Dict, Any, List
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import zipfile
import io
import anthropic

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

# ============================================================================
# APP SETUP
# ============================================================================

app = FastAPI(title="NOVA Agent Server", version="4.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Storage - use a simple class to ensure updates are visible
class JobStore:
    def __init__(self):
        self._jobs: Dict[str, Dict] = {}
    
    def get(self, job_id: str) -> Optional[Dict]:
        return self._jobs.get(job_id)
    
    def set(self, job_id: str, data: Dict):
        self._jobs[job_id] = data
    
    def update(self, job_id: str, **kwargs):
        if job_id in self._jobs:
            self._jobs[job_id].update(kwargs)
    
    def exists(self, job_id: str) -> bool:
        return job_id in self._jobs

jobs = JobStore()

OUTPUT_DIR = Path("/tmp/nova-outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

# Claude client
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY) if ANTHROPIC_API_KEY else None

print(f"[NOVA v4.1] Started. Claude configured: {claude_client is not None}")


# ============================================================================
# MODELS
# ============================================================================

class TaskRequest(BaseModel):
    job_id: str
    agent: str
    parameters: Dict[str, Any] = {}

class TaskResponse(BaseModel):
    job_id: str
    status: str
    message: str

class StatusResponse(BaseModel):
    job_id: str
    status: str
    progress: int
    current_step: str
    steps_completed: List[str]
    error: Optional[str]
    created_at: str
    completed_at: Optional[str]


# ============================================================================
# ENDPOINTS
# ============================================================================

@app.get("/api/health")
async def health():
    return {"status": "healthy", "version": "4.1.0", "claude": claude_client is not None}


@app.post("/api/execute", response_model=TaskResponse)
async def execute_task(request: TaskRequest, background_tasks: BackgroundTasks):
    """Start an agent task"""
    job_id = request.job_id
    agent = request.agent
    
    # Normalize agent name
    if agent in ['tna', 'analysis']:
        agent = 'analysis'
    
    print(f"[NOVA] Execute: job={job_id}, agent={agent}")
    
    # Create job directory
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(exist_ok=True)
    
    # Initialize job
    jobs.set(job_id, {
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
        "output_dir": str(job_dir)
    })
    
    # Run in background
    background_tasks.add_task(run_agent, job_id, agent, request.parameters)
    
    return TaskResponse(job_id=job_id, status="queued", message=f"Agent '{agent}' queued")


@app.get("/api/status/{job_id}", response_model=StatusResponse)
async def get_status(job_id: str):
    """Get job status"""
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    
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
async def download_files(job_id: str):
    """Download job outputs as ZIP"""
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    
    job_dir = Path(job["output_dir"])
    if not job_dir.exists():
        raise HTTPException(404, "Output directory not found")
    
    # Create ZIP in memory
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in job_dir.rglob('*'):
            if file_path.is_file():
                arc_name = file_path.relative_to(job_dir)
                zf.write(file_path, arc_name)
    
    zip_buffer.seek(0)
    
    from fastapi.responses import StreamingResponse
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=NOVA_{job_id}.zip"}
    )


# ============================================================================
# PROGRESS TRACKING
# ============================================================================

def update_job(job_id: str, progress: int, step: str):
    """Update job progress - called synchronously"""
    job = jobs.get(job_id)
    if job:
        job["progress"] = progress
        job["current_step"] = step
        if step not in job["steps_completed"]:
            job["steps_completed"].append(step)
        print(f"[NOVA] {job_id}: {progress}% - {step}")


# ============================================================================
# AGENT EXECUTION
# ============================================================================

async def run_agent(job_id: str, agent: str, parameters: Dict):
    """Run the specified agent"""
    try:
        job = jobs.get(job_id)
        job["status"] = "running"
        
        if agent == "analysis":
            await run_analysis_agent(job_id, parameters)
        else:
            raise ValueError(f"Agent '{agent}' not implemented")
        
        job["status"] = "completed"
        job["progress"] = 100
        job["current_step"] = "Complete"
        job["completed_at"] = datetime.utcnow().isoformat()
        print(f"[NOVA] Job {job_id} completed")
        
    except Exception as e:
        print(f"[NOVA] Job {job_id} failed: {e}")
        import traceback
        traceback.print_exc()
        job = jobs.get(job_id)
        if job:
            job["status"] = "failed"
            job["error"] = str(e)


# ============================================================================
# CLAUDE API
# ============================================================================

async def call_claude(prompt: str, max_tokens: int = 4000) -> str:
    """Call Claude API and return response text"""
    if not claude_client:
        raise Exception("Claude API not configured")
    
    print(f"[NOVA] Calling Claude ({len(prompt)} chars)...")
    
    # Run synchronous API call in thread pool to not block
    loop = asyncio.get_event_loop()
    response = await loop.run_in_executor(
        None,
        lambda: claude_client.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}]
        )
    )
    
    result = response.content[0].text
    print(f"[NOVA] Claude response: {len(result)} chars")
    return result


def parse_json(text: str) -> Dict:
    """Parse JSON from Claude response"""
    # Strategy 1: Look for ```json block
    match = re.search(r'```json\s*([\s\S]*?)\s*```', text)
    if match:
        try:
            return json.loads(match.group(1))
        except:
            pass
    
    # Strategy 2: Find first { to last }
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end > start:
        try:
            return json.loads(text[start:end+1])
        except:
            # Try fixing common issues
            json_str = text[start:end+1]
            json_str = re.sub(r',\s*}', '}', json_str)
            json_str = re.sub(r',\s*]', ']', json_str)
            try:
                return json.loads(json_str)
            except:
                pass
    
    raise Exception(f"Could not parse JSON from response: {text[:500]}...")


# ============================================================================
# ANALYSIS AGENT - Only RolePS and TNR
# ============================================================================

async def run_analysis_agent(job_id: str, parameters: Dict):
    """Generate 2 Analysis documents: RolePS and TNR"""
    role_title = parameters.get("role_title", "Training Specialist")
    framework = parameters.get("framework", "UK-DSAT")
    role_desc = parameters.get("role_description", "")
    
    output_dir = Path(jobs.get(job_id)["output_dir"]) / "01_Analysis"
    output_dir.mkdir(exist_ok=True)
    
    update_job(job_id, 5, "Starting Analysis Agent...")
    
    # Generate RolePS
    update_job(job_id, 10, "Generating Role Performance Statement...")
    roleps = await generate_roleps(role_title, framework, role_desc)
    update_job(job_id, 40, "Building RolePS document...")
    build_roleps_doc(roleps, role_title, framework, output_dir / "01_Role_Performance_Statement.docx")
    update_job(job_id, 50, "✓ Role Performance Statement complete")
    
    # Generate TNR
    update_job(job_id, 55, "Generating Training Needs Report...")
    tnr = await generate_tnr(role_title, framework, roleps)
    update_job(job_id, 85, "Building TNR document...")
    build_tnr_doc(tnr, role_title, output_dir / "02_Training_Needs_Report.docx")
    update_job(job_id, 95, "✓ Training Needs Report complete")
    
    update_job(job_id, 100, "Analysis Phase Complete")


# ============================================================================
# CONTENT GENERATION
# ============================================================================

async def generate_roleps(role_title: str, framework: str, description: str) -> Dict:
    """Generate Role Performance Statement content matching template format"""
    prompt = f"""Generate a Role Performance Statement (RolePS) for:

Role: {role_title}
Framework: {framework}
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "security_classification": "OFFICIAL",
        "role_title": "{role_title}",
        "role_number": "2025/001",
        "duty_title": "Primary duty description for this role",
        "duty_number": "1",
        "tra": "Training Requirements Authority name",
        "tda": "Training Delivery Authority name",
        "roleps_reference": "RPS-2025-001",
        "issue_status": "Draft v1.0"
    }},
    "tasks": [
        {{
            "task_number": "1.1",
            "performance": "Clear task performance statement",
            "conditions": [
                "Environment: Office/Field/Deployed",
                "Equipment: Specific equipment used",
                "Situation: Working context"
            ],
            "standards": [
                "Standard 1 to be met",
                "Standard 2 to be met"
            ],
            "training_category": "3",
            "notes": [
                "Knowledge: What they need to know",
                "Skill: What they need to do",
                "Attitude: Behavioral requirements"
            ]
        }}
    ]
}}

Generate 8-12 tasks with:
- Clear performance statements
- 3-5 conditions per task
- 2-4 standards per task  
- Training category (1-4)
- 3-5 notes (Knowledge/Skill/Attitude items)

Make content specific and realistic for a {role_title} role.
Return ONLY the JSON, no other text."""

    response = await call_claude(prompt, max_tokens=6000)
    return parse_json(response)


async def generate_tnr(role_title: str, framework: str, roleps: Dict) -> Dict:
    """Generate Training Needs Report content"""
    num_tasks = len(roleps.get("tasks", []))
    
    prompt = f"""Generate a Training Needs Report for:

Role: {role_title}
Framework: {framework}
Number of tasks analysed: {num_tasks}

Return a JSON object with:
{{
    "executive_summary": "3-4 paragraphs summarizing the training needs analysis, key findings, and recommendations. Make this comprehensive and specific to the role.",
    "introduction": {{
        "purpose": "Purpose of this Training Needs Report",
        "scope": "Scope of the analysis conducted",
        "methodology": "Methodology used for the analysis"
    }},
    "key_findings": [
        "Finding 1 - detailed finding about training needs",
        "Finding 2 - another key finding",
        "Finding 3 - additional finding"
    ],
    "training_requirements": [
        {{
            "id": "TR1",
            "requirement": "Specific training requirement",
            "priority": "Critical",
            "delivery_method": "Blended"
        }}
    ],
    "recommendations": [
        {{
            "id": "R1",
            "recommendation": "Specific recommendation",
            "rationale": "Why this is recommended",
            "priority": "High",
            "timeline": "Short-term"
        }}
    ],
    "resource_requirements": {{
        "budget_estimate": 75000,
        "timeline_months": 6,
        "personnel_required": "2 trainers, 1 training manager"
    }},
    "conclusion": "2-3 paragraphs with conclusions and next steps"
}}

Generate 5 key findings, 6 training requirements, and 5 recommendations.
Make all content specific and realistic for a {role_title} role.
Return ONLY the JSON, no other text."""

    response = await call_claude(prompt, max_tokens=5000)
    return parse_json(response)


# ============================================================================
# DOCUMENT STYLING HELPERS
# ============================================================================

# NOVA brand colors
NOVA_DARK_RED = RGBColor(139, 69, 69)  # Dark red for headers
NOVA_LIGHT_GRAY = RGBColor(245, 245, 245)  # Light gray for alternating rows
NOVA_DARK_BLUE = RGBColor(31, 56, 100)  # Dark blue for headings

def set_cell_shading(cell, color_hex: str):
    """Set cell background color"""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    cell._tc.get_or_add_tcPr().append(shading)

def set_paragraph_spacing(paragraph, before_pt=6, after_pt=6, line_spacing=1.5):
    """Set paragraph spacing: before/after in points, line spacing as multiplier"""
    pPr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(int(before_pt * 20)))  # Convert to twips
    spacing.set(qn('w:after'), str(int(after_pt * 20)))
    spacing.set(qn('w:line'), str(int(line_spacing * 240)))  # 240 twips = single line
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)

def create_styled_paragraph(doc, text: str, font_name: str = "Roboto", font_size: int = 11, 
                           bold: bool = False, color: RGBColor = None):
    """Create a paragraph with standard styling"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    set_paragraph_spacing(p, before_pt=6, after_pt=6, line_spacing=1.5)
    return p

def create_styled_heading(doc, text: str, level: int = 1):
    """Create a styled heading"""
    heading = doc.add_heading(text, level)
    for run in heading.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    return heading

def create_styled_table(doc, headers: List[str], rows: List[List[str]], col_widths: List[float] = None):
    """Create a professionally styled table"""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    # Header row with dark red background
    header_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = header
        # Style header cell
        set_cell_shading(cell, "8B4545")  # Dark red
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.name = "Roboto"
                run.font.size = Pt(10)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)  # White text
    
    # Data rows with alternating colors
    for row_idx, row_data in enumerate(rows):
        row = table.add_row()
        for i, cell_text in enumerate(row_data):
            cell = row.cells[i]
            cell.text = str(cell_text) if cell_text else ""
            # Alternating row colors
            if row_idx % 2 == 1:
                set_cell_shading(cell, "F5F5F5")  # Light gray
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(10)
    
    # Set column widths if provided
    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(width)
    
    return table


# ============================================================================
# ROLEPS DOCUMENT BUILDER - Landscape with template layout
# ============================================================================

def build_roleps_doc(data: Dict, role_title: str, framework: str, filepath: Path):
    """Build Role Performance Statement in landscape format matching template"""
    doc = Document()
    
    # Set to landscape
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    # Swap width and height for landscape
    new_width = section.page_height
    new_height = section.page_width
    section.page_width = new_width
    section.page_height = new_height
    
    # Set margins
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    
    header = data.get("header", {})
    
    # Security Classification
    p = doc.add_paragraph()
    run = p.add_run(header.get("security_classification", "OFFICIAL"))
    run.font.name = "Roboto"
    run.font.size = Pt(12)
    run.font.bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Header information table (2 columns layout)
    header_table = doc.add_table(rows=4, cols=4)
    header_table.style = 'Table Grid'
    
    # Row 1: Role Title | Role Number
    header_table.rows[0].cells[0].text = "ROLE TITLE(S):"
    header_table.rows[0].cells[1].text = header.get("role_title", role_title)
    header_table.rows[0].cells[2].text = "ROLE NUMBER(S):"
    header_table.rows[0].cells[3].text = header.get("role_number", "")
    
    # Row 2: Duty Title | Duty Number
    header_table.rows[1].cells[0].text = "DUTY TITLE(S):"
    header_table.rows[1].cells[1].text = header.get("duty_title", "")
    header_table.rows[1].cells[2].text = "DUTY NUMBER(S):"
    header_table.rows[1].cells[3].text = header.get("duty_number", "")
    
    # Row 3: TRA | RolePS Reference
    header_table.rows[2].cells[0].text = "TRA:"
    header_table.rows[2].cells[1].text = header.get("tra", "")
    header_table.rows[2].cells[2].text = "ROLE PS REFERENCE:"
    header_table.rows[2].cells[3].text = header.get("roleps_reference", "")
    
    # Row 4: TDA | Issue Status
    header_table.rows[3].cells[0].text = "TDA:"
    header_table.rows[3].cells[1].text = header.get("tda", "")
    header_table.rows[3].cells[2].text = "ISSUE STATUS:"
    header_table.rows[3].cells[3].text = header.get("issue_status", "")
    
    # Style header table
    for row in header_table.rows:
        for i, cell in enumerate(row.cells):
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(10)
                    if i % 2 == 0:  # Label columns
                        run.font.bold = True
    
    doc.add_paragraph()  # Spacing
    
    # Task table with 6 columns
    task_headers = ["Task/Sub Task Number", "Performance", "Conditions", "Standards", "Training Category", "Notes"]
    
    # Create task table
    task_table = doc.add_table(rows=1, cols=6)
    task_table.style = 'Table Grid'
    
    # Header row
    header_row = task_table.rows[0]
    for i, h in enumerate(task_headers):
        cell = header_row.cells[i]
        cell.text = h
        set_cell_shading(cell, "8B4545")  # Dark red
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.name = "Roboto"
                run.font.size = Pt(10)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Add task rows
    tasks = data.get("tasks", [])
    for task_idx, task in enumerate(tasks):
        row = task_table.add_row()
        
        # Task Number
        row.cells[0].text = str(task.get("task_number", ""))
        
        # Performance
        row.cells[1].text = task.get("performance", "")
        
        # Conditions (join with newlines)
        conditions = task.get("conditions", [])
        row.cells[2].text = "\n".join(conditions) if isinstance(conditions, list) else str(conditions)
        
        # Standards (join with newlines)
        standards = task.get("standards", [])
        row.cells[3].text = "\n".join(standards) if isinstance(standards, list) else str(standards)
        
        # Training Category
        row.cells[4].text = str(task.get("training_category", ""))
        
        # Notes (join with newlines)
        notes = task.get("notes", [])
        row.cells[5].text = "\n".join(notes) if isinstance(notes, list) else str(notes)
        
        # Alternating row colors
        if task_idx % 2 == 1:
            for cell in row.cells:
                set_cell_shading(cell, "F5F5F5")
        
        # Style all cells
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(9)
    
    # Set column widths (landscape gives us ~10 inches usable)
    col_widths = [1.0, 2.0, 2.0, 1.5, 0.8, 2.5]
    for row in task_table.rows:
        for i, width in enumerate(col_widths):
            row.cells[i].width = Inches(width)
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


# ============================================================================
# TNR DOCUMENT BUILDER - Professional styling
# ============================================================================

def build_tnr_doc(data: Dict, role_title: str, filepath: Path):
    """Build Training Needs Report with professional styling"""
    doc = Document()
    
    # Set margins
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    
    # Title
    title = doc.add_heading("TRAINING NEEDS REPORT", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.size = Pt(24)
        run.font.color.rgb = NOVA_DARK_BLUE
    
    # Subtitle
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(16)
    run.font.color.rgb = NOVA_DARK_BLUE
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Horizontal line
    doc.add_paragraph("_" * 80)
    
    # 1. Executive Summary
    create_styled_heading(doc, "1. EXECUTIVE SUMMARY", 1)
    exec_summary = data.get("executive_summary", "")
    # Split into paragraphs if it's one long string
    paragraphs = exec_summary.split('\n\n') if '\n\n' in exec_summary else [exec_summary]
    for para_text in paragraphs:
        if para_text.strip():
            create_styled_paragraph(doc, para_text.strip())
    
    # 2. Introduction
    create_styled_heading(doc, "2. INTRODUCTION", 1)
    intro = data.get("introduction", {})
    
    create_styled_paragraph(doc, f"Purpose: {intro.get('purpose', '')}")
    create_styled_paragraph(doc, f"Scope: {intro.get('scope', '')}")
    create_styled_paragraph(doc, f"Methodology: {intro.get('methodology', '')}")
    
    # 3. Key Findings
    create_styled_heading(doc, "3. KEY FINDINGS", 1)
    findings = data.get("key_findings", [])
    for i, finding in enumerate(findings, 1):
        create_styled_paragraph(doc, f"{i}. {finding}")
    
    # 4. Training Requirements - styled table
    create_styled_heading(doc, "4. TRAINING REQUIREMENTS", 1)
    requirements = data.get("training_requirements", [])
    if requirements:
        rows = [[r.get("id", ""), r.get("requirement", ""), r.get("priority", ""), r.get("delivery_method", "")] for r in requirements]
        create_styled_table(doc, ["ID", "Requirement", "Priority", "Delivery Method"], rows, [0.6, 4.0, 0.9, 1.1])
    
    doc.add_paragraph()  # Spacing
    
    # 5. Recommendations - styled table
    create_styled_heading(doc, "5. RECOMMENDATIONS", 1)
    recommendations = data.get("recommendations", [])
    if recommendations:
        rows = [[r.get("id", ""), r.get("recommendation", ""), r.get("priority", ""), r.get("timeline", "")] for r in recommendations]
        create_styled_table(doc, ["ID", "Recommendation", "Priority", "Timeline"], rows, [0.6, 4.0, 0.8, 1.2])
    
    doc.add_paragraph()  # Spacing
    
    # 6. Resource Requirements
    create_styled_heading(doc, "6. RESOURCE REQUIREMENTS", 1)
    resources = data.get("resource_requirements", {})
    create_styled_paragraph(doc, f"Estimated Budget: £{resources.get('budget_estimate', 0):,}")
    create_styled_paragraph(doc, f"Timeline: {resources.get('timeline_months', 0)} months")
    create_styled_paragraph(doc, f"Personnel: {resources.get('personnel_required', '')}")
    
    # 7. Conclusion
    create_styled_heading(doc, "7. CONCLUSION", 1)
    conclusion = data.get("conclusion", "")
    paragraphs = conclusion.split('\n\n') if '\n\n' in conclusion else [conclusion]
    for para_text in paragraphs:
        if para_text.strip():
            create_styled_paragraph(doc, para_text.strip())
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
