"""
NOVA Agent Server v4.0
Autonomous Training Agent Execution Server

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

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import zipfile
import io
import anthropic

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============================================================================
# APP SETUP
# ============================================================================

app = FastAPI(title="NOVA Agent Server", version="4.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Storage
jobs: Dict[str, Dict] = {}
OUTPUT_DIR = Path("/tmp/nova-outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

# Claude client
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY) if ANTHROPIC_API_KEY else None

print(f"[NOVA v4.0] Started. Claude configured: {claude_client is not None}")


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
    return {"status": "healthy", "version": "4.0.0", "claude": claude_client is not None}


@app.post("/api/execute", response_model=TaskResponse)
async def execute_task(request: TaskRequest, background_tasks: BackgroundTasks):
    """Start an agent task"""
    job_id = request.job_id
    agent = request.agent
    
    # Normalize agent name
    if agent in ['tna', 'analysis']:
        agent = 'analysis'
    
    print(f"[NOVA] Execute: job={job_id}, agent={agent}")
    
    # Create job
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(exist_ok=True)
    
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
        "output_dir": str(job_dir)
    }
    
    # Run in background
    background_tasks.add_task(run_agent, job_id, agent, request.parameters)
    
    return TaskResponse(job_id=job_id, status="queued", message=f"Agent '{agent}' queued")


@app.get("/api/status/{job_id}", response_model=StatusResponse)
async def get_status(job_id: str):
    """Get job status"""
    if job_id not in jobs:
        raise HTTPException(404, "Job not found")
    
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
async def download_files(job_id: str):
    """Download job outputs as ZIP"""
    if job_id not in jobs:
        raise HTTPException(404, "Job not found")
    
    job_dir = Path(jobs[job_id]["output_dir"])
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
# AGENT EXECUTION
# ============================================================================

def update_job(job_id: str, progress: int, step: str):
    """Update job progress"""
    if job_id in jobs:
        jobs[job_id]["progress"] = progress
        jobs[job_id]["current_step"] = step
        if step not in jobs[job_id]["steps_completed"]:
            jobs[job_id]["steps_completed"].append(step)
        print(f"[NOVA] {job_id}: {progress}% - {step}")


async def run_agent(job_id: str, agent: str, parameters: Dict):
    """Run the specified agent"""
    try:
        jobs[job_id]["status"] = "running"
        
        if agent == "analysis":
            await run_analysis_agent(job_id, parameters)
        else:
            raise ValueError(f"Agent '{agent}' not implemented")
        
        jobs[job_id]["status"] = "completed"
        jobs[job_id]["completed_at"] = datetime.utcnow().isoformat()
        print(f"[NOVA] Job {job_id} completed")
        
    except Exception as e:
        print(f"[NOVA] Job {job_id} failed: {e}")
        import traceback
        traceback.print_exc()
        jobs[job_id]["status"] = "failed"
        jobs[job_id]["error"] = str(e)


# ============================================================================
# CLAUDE API
# ============================================================================

async def call_claude(prompt: str, max_tokens: int = 4000) -> str:
    """Call Claude API and return response text"""
    if not claude_client:
        raise Exception("Claude API not configured")
    
    print(f"[NOVA] Calling Claude ({len(prompt)} chars)...")
    
    response = claude_client.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}]
    )
    
    result = response.content[0].text
    print(f"[NOVA] Claude response: {len(result)} chars")
    return result


def parse_json(text: str) -> Dict:
    """Parse JSON from Claude response"""
    # Try to find JSON in the response
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
# ANALYSIS AGENT
# ============================================================================

async def run_analysis_agent(job_id: str, parameters: Dict):
    """Generate 4 Analysis documents"""
    role_title = parameters.get("role_title", "Training Specialist")
    framework = parameters.get("framework", "UK-DSAT")
    role_desc = parameters.get("role_description", "")
    
    output_dir = Path(jobs[job_id]["output_dir"]) / "01_Analysis"
    output_dir.mkdir(exist_ok=True)
    
    update_job(job_id, 5, "Starting Analysis Agent...")
    
    # Generate all 4 documents
    update_job(job_id, 10, "Generating Scoping Report...")
    scoping = await generate_scoping(role_title, framework, role_desc)
    build_scoping_doc(scoping, role_title, output_dir / "01_Scoping_Report.docx")
    update_job(job_id, 30, "✓ Scoping Report complete")
    
    update_job(job_id, 35, "Generating Role Performance Statement...")
    roleps = await generate_roleps(role_title, framework, role_desc)
    build_roleps_doc(roleps, role_title, output_dir / "02_Role_Performance_Statement.docx")
    update_job(job_id, 55, "✓ Role Performance Statement complete")
    
    update_job(job_id, 60, "Generating Gap Analysis...")
    gaps = await generate_gaps(role_title, framework, roleps)
    build_gaps_doc(gaps, role_title, output_dir / "03_Training_Gap_Analysis.docx")
    update_job(job_id, 75, "✓ Gap Analysis complete")
    
    update_job(job_id, 80, "Generating Training Needs Report...")
    tnr = await generate_tnr(role_title, framework, scoping, roleps, gaps)
    build_tnr_doc(tnr, role_title, output_dir / "04_Training_Needs_Report.docx")
    update_job(job_id, 95, "✓ Training Needs Report complete")
    
    update_job(job_id, 100, "Analysis Phase Complete")


# ============================================================================
# CONTENT GENERATION
# ============================================================================

async def generate_scoping(role_title: str, framework: str, description: str) -> Dict:
    """Generate Scoping Report content"""
    prompt = f"""Generate a Scoping Exercise Report for training analysis.

Role: {role_title}
Framework: {framework}
Context: {description or 'Standard training requirement'}

Return a JSON object with:
{{
    "introduction": "2-3 paragraphs about purpose and methodology",
    "background": "2-3 paragraphs on operational context and current training",
    "stakeholders": [
        {{"name": "Stakeholder name", "role": "Their role", "interest": "High/Medium/Low"}}
    ],
    "assumptions": [
        {{"id": "A1", "text": "Assumption description", "impact": "Impact if invalid"}}
    ],
    "constraints": [
        {{"id": "C1", "text": "Constraint description", "mitigation": "How to mitigate"}}
    ],
    "risks": [
        {{"id": "R1", "text": "Risk description", "likelihood": "High/Medium/Low", "impact": "High/Medium/Low", "mitigation": "Mitigation action"}}
    ],
    "resources": {{
        "personnel": [{{"role": "Role name", "fte": 0.5, "weeks": 12, "cost": 15000}}],
        "total_cost": 50000
    }},
    "recommendations": [
        {{"text": "Recommendation", "priority": "High/Medium/Low"}}
    ]
}}

Generate 5 stakeholders, 5 assumptions, 4 constraints, 5 risks, 3 personnel resources, and 4 recommendations.
Make content specific and realistic for a {role_title} role.
Return ONLY the JSON, no other text."""

    response = await call_claude(prompt)
    return parse_json(response)


async def generate_roleps(role_title: str, framework: str, description: str) -> Dict:
    """Generate Role Performance Statement content"""
    prompt = f"""Generate a Role Performance Statement (Job Task Analysis) for:

Role: {role_title}
Framework: {framework}
Context: {description or 'Standard role requirements'}

Return a JSON object with:
{{
    "role_info": {{
        "title": "{role_title}",
        "purpose": "Brief description of role purpose",
        "reporting_to": "Manager/supervisor title"
    }},
    "duties": [
        {{
            "id": "D1",
            "title": "Duty title",
            "description": "Duty description",
            "tasks": [
                {{
                    "id": "T1.1",
                    "task": "Task description",
                    "performance": "Performance standard",
                    "conditions": "Conditions under which task performed",
                    "standard": "Measurable standard",
                    "frequency": "Daily/Weekly/Monthly/As Required",
                    "criticality": "High/Medium/Low"
                }}
            ]
        }}
    ]
}}

Generate 5 duties with 3-4 tasks each (15-20 total tasks).
Make tasks specific and realistic for a {role_title} role.
Return ONLY the JSON, no other text."""

    response = await call_claude(prompt)
    return parse_json(response)


async def generate_gaps(role_title: str, framework: str, roleps: Dict) -> Dict:
    """Generate Gap Analysis content"""
    # Extract task summaries for context
    task_list = []
    for duty in roleps.get("duties", []):
        for task in duty.get("tasks", []):
            task_list.append(task.get("task", ""))
    
    prompt = f"""Generate a Training Gap Analysis for:

Role: {role_title}
Framework: {framework}
Key Tasks: {', '.join(task_list[:10])}

Return a JSON object with:
{{
    "executive_summary": "2-3 paragraphs summarizing key gaps and recommendations",
    "gaps": [
        {{
            "id": "G1",
            "gap_title": "Gap title",
            "description": "Description of the training gap",
            "current_state": "Current capability level",
            "required_state": "Required capability level",
            "impact": "Impact of gap on operations",
            "priority": "Critical/High/Medium/Low",
            "recommended_solution": "How to address the gap"
        }}
    ],
    "summary": {{
        "critical_gaps": 2,
        "high_gaps": 3,
        "medium_gaps": 2,
        "total_gaps": 7
    }}
}}

Generate 7-10 training gaps.
Make gaps specific and realistic for a {role_title} role.
Return ONLY the JSON, no other text."""

    response = await call_claude(prompt)
    return parse_json(response)


async def generate_tnr(role_title: str, framework: str, scoping: Dict, roleps: Dict, gaps: Dict) -> Dict:
    """Generate Training Needs Report content"""
    prompt = f"""Generate a Training Needs Report for:

Role: {role_title}
Framework: {framework}
Number of gaps identified: {len(gaps.get('gaps', []))}

Return a JSON object with:
{{
    "executive_summary": "3-4 paragraphs summarizing findings and recommendations",
    "introduction": {{
        "purpose": "Purpose of this report",
        "scope": "Scope of the analysis",
        "methodology": "Methodology used"
    }},
    "findings": {{
        "key_findings": ["Finding 1", "Finding 2", "Finding 3"],
        "training_requirements": [
            {{
                "id": "TR1",
                "requirement": "Training requirement description",
                "rationale": "Why this training is needed",
                "priority": "Critical/High/Medium/Low",
                "delivery_method": "Classroom/E-learning/OJT/Blended"
            }}
        ]
    }},
    "recommendations": [
        {{
            "id": "R1",
            "recommendation": "Recommendation text",
            "rationale": "Why this is recommended",
            "priority": "High/Medium/Low",
            "timeline": "Immediate/Short-term/Medium-term"
        }}
    ],
    "resource_requirements": {{
        "budget_estimate": 75000,
        "timeline_months": 6,
        "personnel_required": "2 trainers, 1 training manager"
    }},
    "conclusion": "2-3 paragraphs with final conclusions and next steps"
}}

Generate 5 key findings, 6 training requirements, and 5 recommendations.
Make content specific and realistic for a {role_title} role.
Return ONLY the JSON, no other text."""

    response = await call_claude(prompt)
    return parse_json(response)


# ============================================================================
# DOCUMENT BUILDERS
# ============================================================================

def create_doc(title: str) -> Document:
    """Create a new document with standard formatting"""
    doc = Document()
    
    # Title
    heading = doc.add_heading(title, 0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    return doc


def add_section(doc: Document, title: str, level: int = 1):
    """Add a section heading"""
    doc.add_heading(title, level)


def add_para(doc: Document, text: str):
    """Add a paragraph"""
    if text:
        doc.add_paragraph(str(text))


def add_table(doc: Document, headers: List[str], rows: List[List[str]]):
    """Add a table"""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    # Header row
    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header
    
    # Data rows
    for row_data in rows:
        row = table.add_row()
        for i, cell_data in enumerate(row_data):
            if i < len(row.cells):
                row.cells[i].text = str(cell_data) if cell_data else ""


def build_scoping_doc(data: Dict, role_title: str, filepath: Path):
    """Build Scoping Report document"""
    doc = create_doc(f"SCOPING EXERCISE REPORT\n{role_title}")
    
    add_section(doc, "1. INTRODUCTION")
    add_para(doc, data.get("introduction", ""))
    
    add_section(doc, "2. BACKGROUND")
    add_para(doc, data.get("background", ""))
    
    add_section(doc, "3. STAKEHOLDER ANALYSIS")
    stakeholders = data.get("stakeholders", [])
    if stakeholders:
        add_table(doc, 
            ["Stakeholder", "Role", "Interest Level"],
            [[s.get("name", ""), s.get("role", ""), s.get("interest", "")] for s in stakeholders]
        )
    
    add_section(doc, "4. ASSUMPTIONS")
    assumptions = data.get("assumptions", [])
    if assumptions:
        add_table(doc,
            ["ID", "Assumption", "Impact if Invalid"],
            [[a.get("id", ""), a.get("text", ""), a.get("impact", "")] for a in assumptions]
        )
    
    add_section(doc, "5. CONSTRAINTS")
    constraints = data.get("constraints", [])
    if constraints:
        add_table(doc,
            ["ID", "Constraint", "Mitigation"],
            [[c.get("id", ""), c.get("text", ""), c.get("mitigation", "")] for c in constraints]
        )
    
    add_section(doc, "6. RISK REGISTER")
    risks = data.get("risks", [])
    if risks:
        add_table(doc,
            ["ID", "Risk", "Likelihood", "Impact", "Mitigation"],
            [[r.get("id", ""), r.get("text", ""), r.get("likelihood", ""), r.get("impact", ""), r.get("mitigation", "")] for r in risks]
        )
    
    add_section(doc, "7. RESOURCE ESTIMATE")
    resources = data.get("resources", {})
    personnel = resources.get("personnel", [])
    if personnel:
        add_table(doc,
            ["Role", "FTE", "Duration (weeks)", "Cost (£)"],
            [[p.get("role", ""), str(p.get("fte", "")), str(p.get("weeks", "")), f"£{p.get('cost', 0):,}"] for p in personnel]
        )
    add_para(doc, f"Total Estimated Cost: £{resources.get('total_cost', 0):,}")
    
    add_section(doc, "8. RECOMMENDATIONS")
    recommendations = data.get("recommendations", [])
    for i, rec in enumerate(recommendations, 1):
        add_para(doc, f"{i}. [{rec.get('priority', 'Medium')}] {rec.get('text', '')}")
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_roleps_doc(data: Dict, role_title: str, filepath: Path):
    """Build Role Performance Statement document"""
    doc = create_doc(f"ROLE PERFORMANCE STATEMENT\n{role_title}")
    
    role_info = data.get("role_info", {})
    add_section(doc, "1. ROLE OVERVIEW")
    add_para(doc, f"Role Title: {role_info.get('title', role_title)}")
    add_para(doc, f"Purpose: {role_info.get('purpose', '')}")
    add_para(doc, f"Reports To: {role_info.get('reporting_to', '')}")
    
    add_section(doc, "2. DUTIES AND TASKS")
    
    duties = data.get("duties", [])
    for duty in duties:
        add_section(doc, f"{duty.get('id', '')} - {duty.get('title', '')}", 2)
        add_para(doc, duty.get("description", ""))
        
        tasks = duty.get("tasks", [])
        if tasks:
            add_table(doc,
                ["ID", "Task", "Performance Standard", "Conditions", "Criticality"],
                [[t.get("id", ""), t.get("task", ""), t.get("performance", ""), t.get("conditions", ""), t.get("criticality", "")] for t in tasks]
            )
    
    add_section(doc, "3. SUMMARY")
    total_tasks = sum(len(d.get("tasks", [])) for d in duties)
    add_para(doc, f"Total Duties: {len(duties)}")
    add_para(doc, f"Total Tasks: {total_tasks}")
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_gaps_doc(data: Dict, role_title: str, filepath: Path):
    """Build Gap Analysis document"""
    doc = create_doc(f"TRAINING GAP ANALYSIS\n{role_title}")
    
    add_section(doc, "1. EXECUTIVE SUMMARY")
    add_para(doc, data.get("executive_summary", ""))
    
    add_section(doc, "2. IDENTIFIED GAPS")
    gaps = data.get("gaps", [])
    
    if gaps:
        add_table(doc,
            ["ID", "Gap", "Current State", "Required State", "Priority", "Solution"],
            [[g.get("id", ""), g.get("gap_title", ""), g.get("current_state", ""), g.get("required_state", ""), g.get("priority", ""), g.get("recommended_solution", "")] for g in gaps]
        )
    
    add_section(doc, "3. GAP DETAILS")
    for gap in gaps:
        add_section(doc, f"{gap.get('id', '')} - {gap.get('gap_title', '')}", 2)
        add_para(doc, f"Description: {gap.get('description', '')}")
        add_para(doc, f"Impact: {gap.get('impact', '')}")
        add_para(doc, f"Recommended Solution: {gap.get('recommended_solution', '')}")
    
    add_section(doc, "4. SUMMARY")
    summary = data.get("summary", {})
    add_para(doc, f"Critical Gaps: {summary.get('critical_gaps', 0)}")
    add_para(doc, f"High Priority Gaps: {summary.get('high_gaps', 0)}")
    add_para(doc, f"Medium Priority Gaps: {summary.get('medium_gaps', 0)}")
    add_para(doc, f"Total Gaps: {summary.get('total_gaps', len(gaps))}")
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_tnr_doc(data: Dict, role_title: str, filepath: Path):
    """Build Training Needs Report document"""
    doc = create_doc(f"TRAINING NEEDS REPORT\n{role_title}")
    
    add_section(doc, "1. EXECUTIVE SUMMARY")
    add_para(doc, data.get("executive_summary", ""))
    
    add_section(doc, "2. INTRODUCTION")
    intro = data.get("introduction", {})
    add_para(doc, f"Purpose: {intro.get('purpose', '')}")
    add_para(doc, f"Scope: {intro.get('scope', '')}")
    add_para(doc, f"Methodology: {intro.get('methodology', '')}")
    
    add_section(doc, "3. KEY FINDINGS")
    findings = data.get("findings", {})
    key_findings = findings.get("key_findings", [])
    for i, finding in enumerate(key_findings, 1):
        add_para(doc, f"{i}. {finding}")
    
    add_section(doc, "4. TRAINING REQUIREMENTS")
    requirements = findings.get("training_requirements", [])
    if requirements:
        add_table(doc,
            ["ID", "Requirement", "Priority", "Delivery Method"],
            [[r.get("id", ""), r.get("requirement", ""), r.get("priority", ""), r.get("delivery_method", "")] for r in requirements]
        )
    
    add_section(doc, "5. RECOMMENDATIONS")
    recommendations = data.get("recommendations", [])
    if recommendations:
        add_table(doc,
            ["ID", "Recommendation", "Priority", "Timeline"],
            [[r.get("id", ""), r.get("recommendation", ""), r.get("priority", ""), r.get("timeline", "")] for r in recommendations]
        )
    
    add_section(doc, "6. RESOURCE REQUIREMENTS")
    resources = data.get("resource_requirements", {})
    add_para(doc, f"Estimated Budget: £{resources.get('budget_estimate', 0):,}")
    add_para(doc, f"Timeline: {resources.get('timeline_months', 0)} months")
    add_para(doc, f"Personnel: {resources.get('personnel_required', '')}")
    
    add_section(doc, "7. CONCLUSION")
    add_para(doc, data.get("conclusion", ""))
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
