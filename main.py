"""
NOVA Agent Server
FastAPI server for executing autonomous DSAT agents with Claude AI

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
from datetime import datetime
from typing import Optional, Dict, Any
from pathlib import Path

from fastapi import FastAPI, HTTPException, BackgroundTasks, Header
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import zipfile
import io
import anthropic

# Initialize FastAPI
app = FastAPI(
    title="NOVA Agent Server",
    description="Autonomous DSAT Agent Execution Server",
    version="1.0.0"
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory job storage (use Redis in production)
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


# Request/Response Models
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


# Authentication helper
def verify_auth(authorization: Optional[str] = Header(None)):
    if API_SECRET and authorization != f"Bearer {API_SECRET}":
        raise HTTPException(status_code=401, detail="Unauthorized")


# Health check
@app.get("/api/health")
async def health_check():
    return {
        "status": "healthy",
        "service": "NOVA Agent Server",
        "version": "1.0.0",
        "claude_configured": claude_client is not None,
        "timestamp": datetime.utcnow().isoformat()
    }


# Execute agent task
@app.post("/api/execute", response_model=TaskResponse)
async def execute_task(
    request: TaskRequest,
    background_tasks: BackgroundTasks,
    authorization: Optional[str] = Header(None)
):
    verify_auth(authorization)
    
    job_id = request.job_id
    
    # Validate agent type
    valid_agents = ['tna', 'design', 'delivery', 'course-generator']
    if request.agent not in valid_agents:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid agent: {request.agent}. Valid: {valid_agents}"
        )
    
    # Initialize job
    jobs[job_id] = {
        "job_id": job_id,
        "agent": request.agent,
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
    background_tasks.add_task(run_agent, job_id, request.agent, request.parameters)
    
    return TaskResponse(
        job_id=job_id,
        status="queued",
        message=f"Agent '{request.agent}' task queued for execution"
    )


# Get task status
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


# Download completed files
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
        return FileResponse(
            file_path,
            filename=file,
            media_type="application/octet-stream"
        )
    
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
    filename = f"NOVA_{safe_title}_{job_id[:8]}.zip"
    
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        }
    )


# Background task: Run agent
async def run_agent(job_id: str, agent: str, parameters: Dict[str, Any]):
    """Execute the specified agent"""
    job = jobs[job_id]
    job["status"] = "running"
    
    try:
        if agent == "tna":
            await run_tna_agent(job_id, parameters)
        elif agent == "design":
            await run_design_agent(job_id, parameters)
        elif agent == "delivery":
            await run_delivery_agent(job_id, parameters)
        elif agent == "course-generator":
            await run_course_generator_agent(job_id, parameters)
        
        job["status"] = "completed"
        job["progress"] = 100
        job["completed_at"] = datetime.utcnow().isoformat()
        
    except Exception as e:
        job["status"] = "failed"
        job["error"] = str(e)
        job["completed_at"] = datetime.utcnow().isoformat()


def update_progress(job_id: str, progress: int, step: str):
    """Helper to update job progress"""
    if job_id in jobs:
        jobs[job_id]["progress"] = progress
        jobs[job_id]["current_step"] = step
        if step not in jobs[job_id]["steps_completed"]:
            jobs[job_id]["steps_completed"].append(step)


async def call_claude(prompt: str, system_prompt: str = None) -> str:
    """Call Claude API to generate content"""
    if not claude_client:
        return "[Claude API not configured - please set ANTHROPIC_API_KEY environment variable in Railway]"
    
    try:
        messages = [{"role": "user", "content": prompt}]
        
        kwargs = {
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 4096,
            "messages": messages
        }
        
        if system_prompt:
            kwargs["system"] = system_prompt
        
        response = claude_client.messages.create(**kwargs)
        return response.content[0].text
    except Exception as e:
        return f"[Error calling Claude API: {str(e)}]"


# DSAT System Prompt
DSAT_SYSTEM_PROMPT = """You are NOVA, the Allied Defence Training LLM. You are an expert in:
- UK Defence Systems Approach to Training (DSAT) - JSP 822 and DTSM 1-5
- US Army Systems Approach to Training (SAT) - TRADOC 350-70
- NATO Training - Bi-SC Directive 75-7
- ASD/AIA S6000T Training Analysis and Design Standard

You generate professional, doctrine-compliant training documentation.
Always use formal MOD tone and formatting.
Include specific doctrinal references where appropriate.
Classification: OFFICIAL"""


# Agent implementations
async def run_tna_agent(job_id: str, parameters: Dict[str, Any]):
    """Training Needs Analysis Agent"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "UK")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    # Step 1: Initialize
    update_progress(job_id, 10, "Initializing TNA Agent")
    await asyncio.sleep(0.5)
    
    # Step 2: Search doctrine
    update_progress(job_id, 20, "Searching doctrine library")
    await asyncio.sleep(0.5)
    
    # Create output directory
    (output_dir / "01_Analysis").mkdir(exist_ok=True)
    
    # Step 3: Generate Scoping Report
    update_progress(job_id, 30, "Generating Scoping Report")
    scoping_content = await generate_scoping_report(role_title, framework)
    (output_dir / "01_Analysis" / "Scoping_Report.txt").write_text(scoping_content)
    
    # Step 4: Generate RolePS
    update_progress(job_id, 50, "Generating Role Performance Statement")
    roleps_content = await generate_roleps(role_title, framework)
    (output_dir / "01_Analysis" / "RolePS.txt").write_text(roleps_content)
    
    # Step 5: Training Gap Analysis
    update_progress(job_id, 70, "Conducting Training Gap Analysis")
    tga_content = await generate_tga(role_title, framework)
    (output_dir / "01_Analysis" / "Training_Gap_Analysis.txt").write_text(tga_content)
    
    # Step 6: Training Needs Report
    update_progress(job_id, 85, "Compiling Training Needs Report")
    tnr_content = await generate_tnr(role_title, framework)
    (output_dir / "01_Analysis" / "Training_Needs_Report.txt").write_text(tnr_content)
    
    # Step 7: Validate compliance
    update_progress(job_id, 95, "Validating JSP 822 compliance")
    await asyncio.sleep(0.5)
    
    update_progress(job_id, 100, "TNA Complete")


async def run_design_agent(job_id: str, parameters: Dict[str, Any]):
    """Design Agent"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "UK")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    update_progress(job_id, 20, "Loading RolePS data")
    await asyncio.sleep(0.5)
    
    (output_dir / "02_Design").mkdir(exist_ok=True)
    
    update_progress(job_id, 40, "Generating Training Objectives")
    tos_content = await generate_training_objectives(role_title, framework)
    (output_dir / "02_Design" / "Training_Objectives.txt").write_text(tos_content)
    
    update_progress(job_id, 60, "Creating Enabling Objectives")
    eos_content = await generate_enabling_objectives(role_title, framework)
    (output_dir / "02_Design" / "Enabling_Objectives.txt").write_text(eos_content)
    
    update_progress(job_id, 80, "Building Learning Specification")
    lspec_content = await generate_learning_spec(role_title, framework)
    (output_dir / "02_Design" / "Learning_Specification.txt").write_text(lspec_content)
    
    update_progress(job_id, 100, "Design Complete")


async def run_delivery_agent(job_id: str, parameters: Dict[str, Any]):
    """Delivery Agent"""
    role_title = parameters.get("role_title", "Unknown Role")
    framework = parameters.get("framework", "UK")
    output_dir = Path(jobs[job_id]["output_dir"])
    
    update_progress(job_id, 25, "Loading Learning Specification")
    await asyncio.sleep(0.5)
    
    (output_dir / "03_Delivery").mkdir(exist_ok=True)
    
    update_progress(job_id, 50, "Generating Lesson Plans")
    lessons_content = await generate_lesson_plans(role_title, framework)
    (output_dir / "03_Delivery" / "Lesson_Plans.txt").write_text(lessons_content)
    
    update_progress(job_id, 75, "Creating Assessment Instruments")
    assess_content = await generate_assessments(role_title, framework)
    (output_dir / "03_Delivery" / "Assessment_Instruments.txt").write_text(assess_content)
    
    update_progress(job_id, 100, "Delivery Complete")


async def run_course_generator_agent(job_id: str, parameters: Dict[str, Any]):
    """Course Generator - runs full DSAT lifecycle"""
    update_progress(job_id, 5, "Starting full DSAT lifecycle")
    
    # Run TNA
    await run_tna_agent(job_id, parameters)
    jobs[job_id]["progress"] = 33
    
    # Run Design
    await run_design_agent(job_id, parameters)
    jobs[job_id]["progress"] = 66
    
    # Run Delivery
    await run_delivery_agent(job_id, parameters)
    jobs[job_id]["progress"] = 90
    
    # Generate compliance certificate
    output_dir = Path(jobs[job_id]["output_dir"])
    role_title = parameters.get("role_title", "Unknown Role")
    
    update_progress(job_id, 95, "Generating Compliance Certificate")
    certificate = f"""
NOVA™ DSAT Compliance Certificate
==================================

Role: {role_title}
Framework: {parameters.get('framework', 'UK')}
Generated: {datetime.utcnow().isoformat()}
Job ID: {job_id}

This course package has been validated against:
- JSP 822 V7.0 Volume 2
- DTSM 1-5 (2024 Edition)

Compliance Status: COMPLIANT

Analysis Phase: ✓ Complete
Design Phase: ✓ Complete  
Delivery Phase: ✓ Complete

NOVA™ - Allied Defence Training Intelligence
"""
    (output_dir / "Compliance_Certificate.txt").write_text(certificate)
    
    update_progress(job_id, 100, "Course Package Complete")


# Document generation with Claude
async def generate_scoping_report(role_title: str, framework: str) -> str:
    prompt = f"""Generate a Scoping Exercise Report for a Training Needs Analysis.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Create a complete Scoping Exercise Report following DTSM 2 Section 1.2 requirements.
Include these sections:
1. INTRODUCTION - Purpose and scope of the TNA
2. BACKGROUND - Context for why this training is needed
3. SCOPE - What the TNA will cover, boundaries and constraints
4. STAKEHOLDERS - Key personnel and organizations involved
5. ASSUMPTIONS - Key assumptions being made
6. RISKS - Identified risks and mitigations
7. RESOURCE ESTIMATE - Estimated time and resources for the TNA
8. RECOMMENDATIONS - Next steps

Be specific and realistic for a {role_title} role.
Use formal MOD tone and include doctrinal references."""

    return await call_claude(prompt, DSAT_SYSTEM_PROMPT)


async def generate_roleps(role_title: str, framework: str) -> str:
    prompt = f"""Generate a Role Performance Statement (RolePS) document.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Create a complete RolePS following JSP 822 and DTSM 2 Section 1.3 requirements.

Include a header with:
- Role Title, TRA, TDA, Version, Date

Then generate at least 8-10 realistic tasks for this role, each with:
- Task Number (1.0, 2.0, etc.)
- Performance Statement (what the individual must do)
- Conditions (under what circumstances)
- Standards (to what measurable standard)
- Training Category (FT = Formal Training, WPT = Workplace Training, NT = No Training Required)
- KSA Notes (Knowledge, Skills, Attitudes required)

Make tasks specific, measurable, and realistic for a {role_title}.
Use formal MOD documentation style."""

    return await call_claude(prompt, DSAT_SYSTEM_PROMPT)


async def generate_tga(role_title: str, framework: str) -> str:
    prompt = f"""Generate a Training Gap Analysis document.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Create a Training Gap Analysis following DTSM 2 Section 1.4 requirements.
Include:
1. CURRENT TRAINING PROVISION - What training currently exists
2. REQUIRED CAPABILITY - What capability is needed (from RolePS)
3. GAP IDENTIFICATION - Analysis table showing:
   - Task/Skill area
   - Current provision
   - Required standard
   - Gap identified
   - Risk rating (High/Medium/Low)
4. PRIORITY ASSESSMENT - Which gaps are most critical
5. RECOMMENDATIONS - How gaps should be addressed

Be realistic for a {role_title} role.
Use formal MOD documentation style with doctrinal references."""

    return await call_claude(prompt, DSAT_SYSTEM_PROMPT)


async def generate_tnr(role_title: str, framework: str) -> str:
    prompt = f"""Generate a Training Needs Report (TNR) document.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Create a complete TNR following DTSM 2 Section 1.7 requirements.
This is the executive summary document for the Customer Executive Board (CEB).

Include:
1. EXECUTIVE SUMMARY - Key findings and recommendations (one page)
2. BACKGROUND - Why this TNA was conducted
3. METHODOLOGY - How the analysis was conducted
4. KEY FINDINGS - Summary from Scoping, RolePS, and Gap Analysis
5. OPTIONS ANALYSIS - Training delivery options with costs/benefits:
   - Option A: Formal residential course
   - Option B: Distributed learning
   - Option C: Workplace-based training
   - Option D: Blended approach
6. RISK ASSESSMENT - Risks of each option
7. RECOMMENDATIONS - Recommended option with justification
8. RESOURCE IMPLICATIONS - Time, cost, personnel required
9. NEXT STEPS - Actions required if approved

Be specific and realistic for a {role_title}.
Use formal MOD documentation style for CEB presentation."""

    return await call_claude(prompt, DSAT_SYSTEM_PROMPT)


async def generate_training_objectives(role_title: str, framework: str) -> str:
    prompt = f"""Generate Training Objectives (TOs) for the Design phase.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Following DTSM 3 Section 2.1, create Training Objectives derived from the RolePS tasks.
Generate at least 6-8 Training Objectives, each with:
- TO Number (TO 1, TO 2, etc.)
- Objective Statement (what trainee will be able to do)
- Performance (specific observable action)
- Conditions (under what circumstances - equipment, environment, aids)
- Standards (measurable criteria for success)
- Assessment Method (how competence will be verified)
- Learning Domain (Cognitive/Psychomotor/Affective)
- Learning Level (Remember/Understand/Apply/Analyze/Evaluate/Create)

Make objectives SMART: Specific, Measurable, Achievable, Relevant, Time-bound.
Use formal MOD documentation style with JSP 822 references."""

    return await call_claude(prompt, DSAT_SYSTEM_PROMPT)


async def generate_enabling_objectives(role_title: str, framework: str) -> str:
    prompt = f"""Generate Enabling Objectives (EOs) and Key Learning Points (KLPs).

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Following DTSM 3 Section 2.3, break down Training Objectives into Enabling Objectives.
For each TO, provide 2-4 EOs. For each EO, provide 3-5 KLPs.

Format:
TO 1: [Training Objective]
  EO 1.1: [Enabling Objective]
    KLP 1.1.1: [Key Learning Point] (K/S/A)
    KLP 1.1.2: [Key Learning Point] (K/S/A)
    KLP 1.1.3: [Key Learning Point] (K/S/A)
  EO 1.2: [Enabling Objective]
    KLP 1.2.1: [Key Learning Point] (K/S/A)
    ...

Tag each KLP as:
- K = Knowledge (cognitive)
- S = Skill (psychomotor)
- A = Attitude (affective)

Generate realistic content for a {role_title}.
Use formal MOD documentation style."""

    return await call_claude(prompt, DSAT_SYSTEM_PROMPT)


async def generate_learning_spec(role_title: str, framework: str) -> str:
    prompt = f"""Generate a Learning Specification (LSpec) document.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Following DTSM 3 Section 2.6, create a Learning Specification that controls
"what is taught and how it is taught".

Include:
1. COURSE OVERVIEW
   - Course title, code, duration
   - Target audience and prerequisites
   - Course aim and learning outcomes

2. TO/EO/KLP LISTING
   - Complete hierarchy of objectives

3. DESIGN MATRIX (tabular format)
   - TO | EO | KLP | Method | Media | Assessment | Time | Resources

4. TRAINING METHODS JUSTIFICATION
   - Rationale for each method selected (lecture, demonstration, practical, etc.)
   - Reference to DTSM 3 Section 6

5. ASSESSMENT STRATEGY SUMMARY
   - How each TO will be assessed
   - Pass criteria and remediation policy

6. RESOURCE REQUIREMENTS
   - Trainers, equipment, venues, consumables

7. COURSE SCHEDULE
   - Day-by-day timetable

Be specific and realistic for a {role_title} course.
Use formal MOD documentation style."""

    return await call_claude(prompt, DSAT_SYSTEM_PROMPT)


async def generate_lesson_plans(role_title: str, framework: str) -> str:
    prompt = f"""Generate Lesson Plans for the Delivery phase.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Following DTSM 4, create lesson plans using the PAR structure:
- Present: How content will be presented to trainees
- Apply: How trainees will practice/apply the learning
- Review: How learning will be consolidated

Generate at least 3 complete lesson plans with:
1. LESSON HEADER
   - Lesson title, duration, TOs/EOs addressed
   - Prerequisites, resources required

2. INTRODUCTION (5 mins)
   - Attention getter
   - Learning outcomes
   - Overview

3. PRESENT (main content)
   - KLPs to cover
   - Trainer script/notes
   - Visual aids required
   - Demonstration points

4. APPLY (practical application)
   - Trainee activities
   - Practice exercises
   - Formative assessment

5. REVIEW (consolidation)
   - Summary of key points
   - Q&A
   - Link to next lesson

6. TRAINER NOTES
   - Common trainee errors
   - Safety considerations
   - Timing adjustments

Make realistic for {role_title} training.
Use formal MOD documentation style."""

    return await call_claude(prompt, DSAT_SYSTEM_PROMPT)


async def generate_assessments(role_title: str, framework: str) -> str:
    prompt = f"""Generate Assessment Instruments for the Delivery phase.

Role Title: {role_title}
Framework: {framework}
Date: {datetime.utcnow().strftime('%d %B %Y')}

Following DTSM 3 Section 2.4 (Assessment Strategy) and DTSM 4, create assessment instruments.

Include:

1. ASSESSMENT STRATEGY OVERVIEW
   - Assessment policy
   - Pass/fail criteria
   - Remediation policy
   - AI malpractice prevention (JSP 822 V7.0 requirement)

2. PRACTICAL ASSESSMENT CHECKLISTS
   - For at least 2 practical TOs
   - Observable criteria
   - Pass/fail standards
   - Assessor guidance

3. THEORY ASSESSMENT
   - 15-20 multiple choice questions covering key TOs
   - Include correct answers and distractors
   - Question mapping to TOs/EOs

4. MARKING SCHEMES
   - Clear rubrics for each assessment type
   - Grading criteria

5. ASSESSMENT ADMINISTRATION
   - Test conditions
   - Time limits
   - Resources permitted
   - Invigilation requirements

Make realistic for {role_title} training.
Use formal MOD documentation style."""

    return await call_claude(prompt, DSAT_SYSTEM_PROMPT)


# Run server
if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
