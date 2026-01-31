"""
NOVA Agent Server v5.1
Autonomous Training Agent Execution Server - Multi-Framework Support

Changes from v5.0:
- NEW: Research-based Analysis Agent with 10-step methodology and web search
- NEW: 18-section Analysis Report with full citations
- NEW: 6 input parameters (domain, specialism, role_title, proficiency_level, framework, role_description)
- NEW: Job/Task List with source references
- All outputs 100% factual with zero fabrication when using research mode
- Backward compatible with existing v5.0 Analysis Agent for simple queries

Previous features (v5.0):
- Framework template library integration
- Support for 11 frameworks (UK DSAT, US TRADOC, NATO Bi-SC, Australian SADL, 
  S6000T, ADDIE, SAM, ISO 29990, ATD/Kirkpatrick, Action Mapping)
- All 4 agents implemented (Analysis, Design, Delivery, Evaluation)
- Framework-specific outputs with correct terminology and structure
- Cross-doctrinal fusion support

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
import httpx
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

app = FastAPI(title="NOVA Agent Server", version="5.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Storage
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

# Template library URL
TEMPLATE_LIBRARY_URL = "https://nova-agent.pages.dev/data/NOVA_Framework_Templates_Library.json"

# Cached template library
_template_library: Optional[Dict] = None

print(f"[NOVA v5.1] Started. Claude configured: {claude_client is not None}")


# ============================================================================
# FRAMEWORK TEMPLATE LIBRARY
# ============================================================================

async def get_template_library() -> Dict:
    """Load framework template library (cached)"""
    global _template_library
    
    if _template_library is not None:
        return _template_library
    
    try:
        async with httpx.AsyncClient() as client:
            response = await client.get(TEMPLATE_LIBRARY_URL, timeout=10.0)
            if response.status_code == 200:
                _template_library = response.json()
                print(f"[NOVA] Template library loaded: {len(_template_library.get('metadata', {}).get('frameworks', []))} frameworks")
                return _template_library
    except Exception as e:
        print(f"[NOVA] Warning: Could not fetch template library: {e}")
    
    # Fallback to embedded minimal templates
    _template_library = get_fallback_templates()
    return _template_library


def get_fallback_templates() -> Dict:
    """Fallback minimal template library if fetch fails"""
    return {
        "metadata": {"version": "fallback", "frameworks": ["UK_DSAT"]},
        "UK_DSAT": {
            "framework_name": "UK Defence Systems Approach to Training",
            "citation_format": "[JSP 822 V7.0]"
        }
    }


def normalize_framework(framework: str) -> str:
    """Normalize framework name to template library key"""
    mapping = {
        # Allied Defence
        "uk-dsat": "UK_DSAT", "uk_dsat": "UK_DSAT", "dsat": "UK_DSAT", "uk": "UK_DSAT",
        "us-tradoc": "US_TRADOC", "us_tradoc": "US_TRADOC", "tradoc": "US_TRADOC", "us": "US_TRADOC",
        "nato": "NATO_BISC", "nato-bisc": "NATO_BISC", "nato_bisc": "NATO_BISC", "bisc": "NATO_BISC",
        "australian-sadl": "AUSTRALIAN_SADL", "australian_sadl": "AUSTRALIAN_SADL", "sadl": "AUSTRALIAN_SADL", "australia": "AUSTRALIAN_SADL",
        "s6000t": "S6000T", "asd-s6000t": "S6000T", "asd": "S6000T",
        # Industry Standards
        "addie": "ADDIE",
        "sam": "SAM", "agile": "SAM", "agile-framework": "SAM", "agile-learning-design": "SAM",
        "isd": "ADDIE",  # ISD maps to ADDIE (they're essentially the same methodology)
        "iso": "ISO_29990", "iso-29990": "ISO_29990", "iso_29990": "ISO_29990",
        # Commercial/Corporate
        "kirkpatrick": "KIRKPATRICK", "atd": "KIRKPATRICK",
        "action-mapping": "ACTION_MAPPING", "action_mapping": "ACTION_MAPPING", "cathy-moore": "ACTION_MAPPING",
        # Legacy dashboard.html values - map to appropriate frameworks
        "commercial": "KIRKPATRICK",  # Commercial/Corporate → Kirkpatrick (business-focused)
        "commercial-/-corporate": "KIRKPATRICK",
        "academic": "ADDIE",  # Academic → ADDIE (standard in higher ed)
        "custom": "ADDIE"  # Custom → ADDIE (most flexible base)
    }
    key = framework.lower().replace(" ", "-").replace("_", "-")
    return mapping.get(key, "UK_DSAT")


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
    return {"status": "healthy", "version": "5.1.0", "claude": claude_client is not None}


@app.post("/api/execute", response_model=TaskResponse)
async def execute_task(request: TaskRequest, background_tasks: BackgroundTasks):
    """Start an agent task"""
    job_id = request.job_id
    agent = request.agent.lower()
    
    # Normalize agent name
    agent_map = {
        'tna': 'analysis', 'analysis': 'analysis',
        'design': 'design',
        'delivery': 'delivery',
        'evaluation': 'evaluation', 'assurance': 'evaluation', 'full-package': 'evaluation'
    }
    agent = agent_map.get(agent, agent)
    
    print(f"[NOVA] Execute: job={job_id}, agent={agent}, framework={request.parameters.get('framework', 'UK-DSAT')}")
    
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
    """Update job progress"""
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
        
        # Load template library
        templates = await get_template_library()
        
        # Get framework
        framework = normalize_framework(parameters.get("framework", "UK-DSAT"))
        framework_templates = templates.get(framework, templates.get("UK_DSAT", {}))
        
        if agent == "analysis":
            await run_analysis_agent(job_id, parameters, framework, framework_templates)
        elif agent == "design":
            await run_design_agent(job_id, parameters, framework, framework_templates)
        elif agent == "delivery":
            await run_delivery_agent(job_id, parameters, framework, framework_templates)
        elif agent == "evaluation":
            await run_evaluation_agent(job_id, parameters, framework, framework_templates)
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
    match = re.search(r'```json\s*([\s\S]*?)\s*```', text)
    if match:
        try:
            return json.loads(match.group(1))
        except:
            pass
    
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end > start:
        try:
            return json.loads(text[start:end+1])
        except:
            json_str = text[start:end+1]
            json_str = re.sub(r',\s*}', '}', json_str)
            json_str = re.sub(r',\s*]', ']', json_str)
            try:
                return json.loads(json_str)
            except:
                pass
    
    raise Exception(f"Could not parse JSON from response: {text[:500]}...")


# ============================================================================
# CLAUDE API WITH WEB SEARCH (v5.1)
# ============================================================================

async def call_claude_with_search(system_prompt: str, user_prompt: str, max_tokens: int = 8000) -> Dict:
    """Call Claude API with web search tool enabled for research-based analysis"""
    if not claude_client:
        raise Exception("Claude API not configured")
    
    print(f"[NOVA] Calling Claude with web search ({len(user_prompt)} chars)...")
    
    # Define web search tool
    tools = [
        {
            "type": "web_search_20250305",
            "name": "web_search",
            "max_uses": 20
        }
    ]
    
    loop = asyncio.get_event_loop()
    
    try:
        response = await loop.run_in_executor(
            None,
            lambda: claude_client.messages.create(
                model="claude-sonnet-4-5-20250929",
                max_tokens=max_tokens,
                system=system_prompt,
                tools=tools,
                messages=[{"role": "user", "content": user_prompt}]
            )
        )
        
        # Extract text content and search info
        result = {
            "text": "",
            "citations": [],
            "searches_performed": []
        }
        
        for block in response.content:
            if hasattr(block, 'text'):
                result["text"] += block.text
            elif block.type == "tool_use" and block.name == "web_search":
                result["searches_performed"].append(block.input.get("query", ""))
        
        print(f"[NOVA] Claude response: {len(result['text'])} chars, {len(result['searches_performed'])} searches")
        return result
        
    except Exception as e:
        print(f"[NOVA] Claude web search API error: {e}, falling back to standard call")
        # Fallback to non-search call
        return await call_claude_fallback(system_prompt, user_prompt, max_tokens)


async def call_claude_fallback(system_prompt: str, user_prompt: str, max_tokens: int = 8000) -> Dict:
    """Fallback Claude call without web search"""
    if not claude_client:
        raise Exception("Claude API not configured")
    
    loop = asyncio.get_event_loop()
    response = await loop.run_in_executor(
        None,
        lambda: claude_client.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=max_tokens,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}]
        )
    )
    
    return {
        "text": response.content[0].text,
        "citations": [],
        "searches_performed": []
    }


# ============================================================================
# FRAMEWORK-SPECIFIC TERMINOLOGY
# ============================================================================

def get_terminology(framework: str) -> Dict[str, str]:
    """Get framework-specific terminology"""
    terms = {
        "UK_DSAT": {
            "task_list": "Role Performance Statement (RolePS)",
            "task_list_short": "RolePS",
            "top_objective": "Training Objective (TO)",
            "top_objective_short": "TO",
            "enabling_objective": "Enabling Objective (EO)",
            "enabling_objective_short": "EO",
            "learning_point": "Key Learning Point (KLP)",
            "learning_point_short": "KLP",
            "needs_report": "Training Needs Report (TNR)",
            "course_design": "Learning Specification (LSpec)",
            "lesson_plan": "Lesson Plan (PAR Structure)",
            "internal_eval": "Internal Validation (InVal)",
            "external_eval": "External Validation (ExVal)",
            "objective_format": "Performance-Conditions-Standards",
            "citation_prefix": "[JSP 822",
            "authority": "Training Requirements Authority (TRA)"
        },
        "US_TRADOC": {
            "task_list": "Individual Critical Task List (ICTL)",
            "task_list_short": "ICTL",
            "top_objective": "Terminal Learning Objective (TLO)",
            "top_objective_short": "TLO",
            "enabling_objective": "Enabling Learning Objective (ELO)",
            "enabling_objective_short": "ELO",
            "learning_point": "Learning Step Activity (LSA)",
            "learning_point_short": "LSA",
            "needs_report": "Training Requirements Document (TRD)",
            "course_design": "Program of Instruction (POI)",
            "lesson_plan": "Lesson Plan (5-Section Format)",
            "internal_eval": "After Action Review (AAR)",
            "external_eval": "Course Design Review (CDR)",
            "objective_format": "Action-Condition-Standard",
            "citation_prefix": "[TP 350-70",
            "authority": "Proponent"
        },
        "NATO_BISC": {
            "task_list": "Skills, Tasks, Proficiency (STP) Analysis",
            "task_list_short": "STP Analysis",
            "top_objective": "Performance Objective (PO)",
            "top_objective_short": "PO",
            "enabling_objective": "Enabling Learning Objective (ELO)",
            "enabling_objective_short": "ELO",
            "learning_point": "Learning Points",
            "learning_point_short": "LP",
            "needs_report": "Training Requirements Analysis (TRA) Report",
            "course_design": "Course Control Document II (CCD II)",
            "lesson_plan": "Course Control Document III (CCD III)",
            "internal_eval": "Student/Graduate Feedback Analysis",
            "external_eval": "Annual Quality Assurance Report (AQAR)",
            "objective_format": "Performance-Standard",
            "citation_prefix": "[Bi-SCD 075-007",
            "authority": "Requirements Authority (RA)"
        },
        "ADDIE": {
            "task_list": "Job/Task Analysis Report",
            "task_list_short": "Task Analysis",
            "top_objective": "Learning Objective",
            "top_objective_short": "LO",
            "enabling_objective": "Supporting Objective",
            "enabling_objective_short": "SO",
            "learning_point": "Content Item",
            "learning_point_short": "CI",
            "needs_report": "Analysis Phase Documentation",
            "course_design": "Design Document",
            "lesson_plan": "Lesson Plan",
            "internal_eval": "Formative Evaluation",
            "external_eval": "Summative Evaluation",
            "objective_format": "ABCD (Audience-Behavior-Condition-Degree)",
            "citation_prefix": "[ADDIE",
            "authority": "Project Sponsor"
        },
        "SAM": {
            "task_list": "Savvy Start Analysis Summary",
            "task_list_short": "Savvy Start",
            "top_objective": "Performance Objective",
            "top_objective_short": "PO",
            "enabling_objective": "Supporting Objective",
            "enabling_objective_short": "SO",
            "learning_point": "Practice Activity",
            "learning_point_short": "PA",
            "needs_report": "Preparation Phase Summary",
            "course_design": "Design Proof",
            "lesson_plan": "Iterative Lesson Module",
            "internal_eval": "Iteration Evaluation",
            "external_eval": "Gold Version Validation",
            "objective_format": "Performance-Assessment-Prototype",
            "citation_prefix": "[SAM",
            "authority": "Stakeholder Group"
        },
        "KIRKPATRICK": {
            "task_list": "Level 4 Results Definition",
            "task_list_short": "L4 Definition",
            "top_objective": "Business Outcome",
            "top_objective_short": "BO",
            "enabling_objective": "Behavior Indicator",
            "enabling_objective_short": "BI",
            "learning_point": "Learning Indicator",
            "learning_point_short": "LI",
            "needs_report": "ROE Statement",
            "course_design": "Evaluation Strategy",
            "lesson_plan": "Training Activity",
            "internal_eval": "Levels 1-2 Assessment",
            "external_eval": "Chain of Evidence Report",
            "objective_format": "Results-Behavior-Learning-Reaction",
            "citation_prefix": "[Kirkpatrick",
            "authority": "Sponsor"
        },
        "ACTION_MAPPING": {
            "task_list": "Business Goal Specification",
            "task_list_short": "Goal Spec",
            "top_objective": "Action Statement",
            "top_objective_short": "AS",
            "enabling_objective": "Sub-Action",
            "enabling_objective_short": "SA",
            "learning_point": "Decision Point",
            "learning_point_short": "DP",
            "needs_report": "Root Cause Analysis",
            "course_design": "Activity Design",
            "lesson_plan": "Scenario-Based Activity",
            "internal_eval": "Behavior Observation",
            "external_eval": "Business Impact Assessment",
            "objective_format": "Action-Scenario-Feedback",
            "citation_prefix": "[Action Mapping",
            "authority": "Business Sponsor"
        },
        "AUSTRALIAN_SADL": {
            "task_list": "Task Analysis",
            "task_list_short": "Task Analysis",
            "top_objective": "Learning Outcome (LO)",
            "top_objective_short": "LO",
            "enabling_objective": "Supporting Learning Outcome (SLO)",
            "enabling_objective_short": "SLO",
            "learning_point": "Learning Element",
            "learning_point_short": "LE",
            "needs_report": "Analysis Phase Report",
            "course_design": "Curriculum Design Document",
            "lesson_plan": "Lesson Plan",
            "internal_eval": "Formative Evaluation Report",
            "external_eval": "Summative Evaluation (Level 4)",
            "objective_format": "Conditions-Standards-Delivery-Assessment",
            "citation_prefix": "[DLM",
            "authority": "Training Authority"
        },
        "ISO_29990": {
            "task_list": "Learning Needs Determination",
            "task_list_short": "Needs Determination",
            "top_objective": "Learning Outcome",
            "top_objective_short": "LO",
            "enabling_objective": "Supporting Outcome",
            "enabling_objective_short": "SO",
            "learning_point": "Learning Element",
            "learning_point_short": "LE",
            "needs_report": "Learning Needs Analysis Record",
            "course_design": "Learning Service Design",
            "lesson_plan": "Learning Session Plan",
            "internal_eval": "Learning Service Evaluation",
            "external_eval": "Quality Management Review",
            "objective_format": "Outcome-Evidence-Assessment",
            "citation_prefix": "[ISO 29990",
            "authority": "Learning Service Provider (LSP)"
        },
        "S6000T": {
            "task_list": "Task Specification",
            "task_list_short": "Task Spec",
            "top_objective": "Training Requirement",
            "top_objective_short": "TR",
            "enabling_objective": "Sub-Task Requirement",
            "enabling_objective_short": "STR",
            "learning_point": "Task Element",
            "learning_point_short": "TE",
            "needs_report": "Training Analysis Report",
            "course_design": "Training Specification",
            "lesson_plan": "Training Module",
            "internal_eval": "Training Effectiveness Assessment",
            "external_eval": "Capability Validation",
            "objective_format": "Task-Condition-Standard",
            "citation_prefix": "[S6000T",
            "authority": "Training Authority (TA)"
        }
    }
    return terms.get(framework, terms["UK_DSAT"])


# ============================================================================
# ANALYSIS AGENT
# ============================================================================

async def run_analysis_agent(job_id: str, parameters: Dict, framework: str, templates: Dict):
    """Generate Analysis phase documents based on framework
    
    v5.1: If domain, specialism, and proficiency_level are provided, uses 
    research-based analysis with web search for factual, cited outputs.
    Otherwise, uses the original v5.0 generation approach.
    """
    # Check if research-based mode should be used (v5.1)
    domain = parameters.get("domain", "")
    specialism = parameters.get("specialism", "")
    proficiency_level = parameters.get("proficiency_level", "")
    
    if domain and specialism and proficiency_level:
        # Use new research-based analysis (v5.1)
        print(f"[NOVA] Using research-based Analysis Agent (v5.1)")
        await run_research_analysis_agent(job_id, parameters, framework)
        return
    
    # Original v5.0 analysis flow
    print(f"[NOVA] Using standard Analysis Agent (v5.0)")
    role_title = parameters.get("role_title", "Training Specialist")
    role_desc = parameters.get("role_description", "")
    terms = get_terminology(framework)
    
    output_dir = Path(jobs.get(job_id)["output_dir"]) / "01_Analysis"
    output_dir.mkdir(exist_ok=True)
    
    update_job(job_id, 5, f"Starting Analysis Agent ({framework})...")
    
    # Generate Task List / RolePS equivalent
    update_job(job_id, 10, f"Generating {terms['task_list_short']}...")
    task_list = await generate_task_list(role_title, framework, role_desc, terms)
    update_job(job_id, 40, f"Building {terms['task_list_short']} document...")
    
    filename = sanitize_filename(terms['task_list_short'])
    build_task_list_doc(task_list, role_title, framework, terms, output_dir / f"01_{filename}.docx")
    update_job(job_id, 50, f"✓ {terms['task_list_short']} complete")
    
    # Generate Training Needs Report equivalent
    update_job(job_id, 55, f"Generating {terms['needs_report']}...")
    needs_report = await generate_needs_report(role_title, framework, task_list, terms)
    update_job(job_id, 85, f"Building {terms['needs_report']} document...")
    
    filename = sanitize_filename(terms['needs_report'].split('(')[0].strip())
    build_needs_report_doc(needs_report, role_title, framework, terms, output_dir / f"02_{filename}.docx")
    update_job(job_id, 95, f"✓ {terms['needs_report']} complete")
    
    update_job(job_id, 100, "Analysis Phase Complete")


def sanitize_filename(name: str) -> str:
    """Sanitize string for use as filename"""
    return re.sub(r'[^\w\s-]', '', name).strip().replace(' ', '_')


async def generate_task_list(role_title: str, framework: str, description: str, terms: Dict) -> Dict:
    """Generate framework-specific task list"""
    
    if framework == "US_TRADOC":
        prompt = f"""Generate an Individual Critical Task List (ICTL) for:

Role: {role_title}
Framework: US TRADOC (TP 350-70-1)
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "security_classification": "UNCLASSIFIED",
        "mos": "MOS code for this role",
        "skill_level": "10/20/30/40",
        "proponent": "US Army Training Command",
        "effective_date": "{datetime.now().strftime('%Y-%m-%d')}",
        "supersedes": "N/A"
    }},
    "tasks": [
        {{
            "task_number": "xxx-xxx-xxxx",
            "task_title": "Clear task title",
            "training_domain": "INST",
            "sustainment_frequency": "AN",
            "sustainment_skill_level": "10",
            "subject_area": "Subject area name",
            "conditions": "Task conditions",
            "standards": "Task standards"
        }}
    ]
}}

Training domains: INST (Institutional), OP (Operational), S-D (Self-Development)
Sustainment: AN (Annual), SA (Semi-Annual), QT (Quarterly)

Generate 8-12 tasks with realistic task numbers (xxx-xxx-xxxx format).
Return ONLY the JSON, no other text."""

    elif framework == "NATO_BISC":
        prompt = f"""Generate a Skills, Tasks, Proficiency (STP) Analysis for:

Role: {role_title}
Framework: NATO Bi-SCD 075-007
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "security_classification": "NATO UNCLASSIFIED",
        "discipline": "Discipline name",
        "ra": "Requirements Authority name",
        "dh": "Department Head name",
        "etf": "Education and Training Facility"
    }},
    "tasks": [
        {{
            "task_id": "STP-001",
            "task_description": "Clear task description",
            "proficiency_level": "Knowledge",
            "collective_task": "Associated collective task",
            "gap_indicator": "Training gap indicator"
        }}
    ]
}}

Proficiency levels: Awareness, Knowledge, Skill, Mastery

Generate 8-12 tasks.
Return ONLY the JSON, no other text."""

    elif framework == "ADDIE":
        prompt = f"""Generate a Job/Task Analysis Report for:

Role: {role_title}
Framework: ADDIE Model
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "job_title": "{role_title}",
        "sme_sources": "Subject Matter Expert sources",
        "date_conducted": "{datetime.now().strftime('%Y-%m-%d')}",
        "analyst": "Training Analyst"
    }},
    "job_description": "Comprehensive job description",
    "major_duties": ["Duty 1", "Duty 2", "Duty 3"],
    "tasks": [
        {{
            "task_id": "T-001",
            "task_description": "Clear task description",
            "knowledge": ["Knowledge item 1", "Knowledge item 2"],
            "skills": ["Skill item 1", "Skill item 2"],
            "attitudes": ["Attitude item 1"],
            "criticality": "High",
            "frequency": "Daily"
        }}
    ]
}}

Criticality: High, Medium, Low
Frequency: Daily, Weekly, Monthly, As Required

Generate 8-12 tasks with full KSA breakdown.
Return ONLY the JSON, no other text."""

    elif framework == "SAM":
        prompt = f"""Generate a Savvy Start Analysis Summary for:

Role: {role_title}
Framework: SAM (Successive Approximation Model)
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "project_name": "{role_title} Training",
        "stakeholders": "Key stakeholder names",
        "savvy_start_date": "{datetime.now().strftime('%Y-%m-%d')}",
        "facilitator": "Session Facilitator"
    }},
    "background": "Background information about the training need",
    "target_audience": {{
        "description": "Target audience profile",
        "size": "Estimated audience size",
        "prior_knowledge": "Current knowledge level"
    }},
    "performance_gaps": [
        {{
            "gap_id": "PG-001",
            "current_state": "What they currently do",
            "desired_state": "What they should do",
            "impact": "Business impact of gap"
        }}
    ],
    "constraints": ["Constraint 1", "Constraint 2"],
    "prototype_concepts": [
        {{
            "concept_id": "PC-001",
            "description": "Prototype concept description",
            "approach": "Proposed approach"
        }}
    ]
}}

Generate 5-8 performance gaps and 3-5 prototype concepts.
Return ONLY the JSON, no other text."""

    elif framework == "KIRKPATRICK":
        prompt = f"""Generate a Level 4 Results Definition for:

Role: {role_title}
Framework: Kirkpatrick Four Levels
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "program_name": "{role_title} Training Program",
        "sponsor": "Training Sponsor",
        "date": "{datetime.now().strftime('%Y-%m-%d')}"
    }},
    "business_outcomes": [
        {{
            "outcome_id": "BO-001",
            "outcome": "Measurable business outcome",
            "current_metric": "Current baseline value",
            "target_metric": "Target value",
            "measurement_method": "How it will be measured"
        }}
    ],
    "leading_indicators": [
        {{
            "indicator_id": "LI-001",
            "indicator": "Leading indicator description",
            "data_source": "Where data comes from"
        }}
    ],
    "roe_statement": "Return on Expectations statement - what success looks like for the sponsor"
}}

Generate 4-6 business outcomes and 5-8 leading indicators.
Start with Level 4 Results - define 'Why' and purpose first.
Return ONLY the JSON, no other text."""

    elif framework == "ACTION_MAPPING":
        prompt = f"""Generate a Business Goal Specification for:

Role: {role_title}
Framework: Action Mapping (Cathy Moore)
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "sponsor": "Business Sponsor",
        "stakeholders": "Key stakeholders",
        "date": "{datetime.now().strftime('%Y-%m-%d')}"
    }},
    "measurable_goal": {{
        "statement": "By [date], [metric] will [change] from [current] to [target]",
        "metric": "The specific metric",
        "current_value": "Current baseline",
        "target_value": "Target value",
        "target_date": "Target achievement date"
    }},
    "actions": [
        {{
            "action_id": "A-001",
            "action": "What people need to DO (observable behavior)",
            "why_not_doing": "Why they aren't currently doing it",
            "is_training_solution": true,
            "alternative_solution": "Non-training solution if applicable"
        }}
    ],
    "minimum_viable_training": "Description of the minimum training needed to change behavior"
}}

Focus on actions, not information. Training is NOT always the solution.
Generate 6-10 actions with root cause analysis.
Return ONLY the JSON, no other text."""

    elif framework == "AUSTRALIAN_SADL":
        prompt = f"""Generate a Task Analysis for:

Role: {role_title}
Framework: Australian SADL (Defence Learning Manual)
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "role": "{role_title}",
        "unit": "Unit name",
        "service_branch": "Army/Navy/Air Force",
        "date_conducted": "{datetime.now().strftime('%Y-%m-%d')}",
        "analyst": "Training Analyst"
    }},
    "target_population": {{
        "description": "Target learner profile",
        "learner_subgroups": ["Subgroup 1", "Subgroup 2"],
        "motivations": "Key motivations",
        "cultural_norms": "Relevant cultural considerations",
        "prior_knowledge": "Expected prior knowledge",
        "learning_preferences": "Preferred learning styles"
    }},
    "tasks": [
        {{
            "task_id": "T-001",
            "task_description": "Clear task description",
            "conditions": "Task conditions",
            "standards": "Performance standards",
            "criticality": "Critical",
            "training_domain": "Institutional"
        }}
    ]
}}

Criticality: Critical, Essential, Enabling
Training Domain: Institutional, Workplace, Self-Directed

Generate 8-12 tasks with full context.
Return ONLY the JSON, no other text."""

    elif framework == "ISO_29990":
        prompt = f"""Generate a Learning Needs Determination for:

Role: {role_title}
Framework: ISO 29990 (Learning Services)
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "client": "Client organization",
        "lsp_reference": "LSP-2025-001",
        "date": "{datetime.now().strftime('%Y-%m-%d')}",
        "conducted_by": "Learning Consultant"
    }},
    "needs_identification": {{
        "context": "Organizational context",
        "trigger": "What triggered this learning need",
        "stakeholders": ["Stakeholder 1", "Stakeholder 2"]
    }},
    "learner_profile": {{
        "description": "Target learner description",
        "current_competence": "Current competence level",
        "desired_competence": "Desired competence level",
        "constraints": ["Time constraints", "Resource constraints"]
    }},
    "learning_needs": [
        {{
            "need_id": "LN-001",
            "need_description": "Specific learning need",
            "gap_description": "Current vs desired gap",
            "priority": "High",
            "proposed_solution": "Recommended learning service"
        }}
    ],
    "quality_requirements": {{
        "assessment_method": "How learning will be assessed",
        "success_criteria": "Criteria for successful learning",
        "evaluation_timeline": "When evaluation will occur"
    }}
}}

Priority: High, Medium, Low

Generate 6-10 learning needs.
Return ONLY the JSON, no other text."""

    elif framework == "S6000T":
        prompt = f"""Generate a Task Specification for:

Role: {role_title}
Framework: ASD/AIA S6000T (Training Analysis and Design)
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "system_name": "System/Equipment name",
        "task_spec_id": "TS-2025-001",
        "date": "{datetime.now().strftime('%Y-%m-%d')}",
        "ils_integration": "Integrated Logistic Support reference"
    }},
    "capability_context": {{
        "capability_requirement": "Operational capability requirement",
        "capability_gap": "Identified capability gap",
        "training_contribution": "How training addresses the gap"
    }},
    "tasks": [
        {{
            "task_id": "TSK-001",
            "task_description": "Task description derived from capability requirement",
            "conditions": "Operational conditions",
            "standards": "Performance standards",
            "task_type": "Operator",
            "frequency": "Routine",
            "ils_element": "Training"
        }}
    ]
}}

Task Type: Operator, Maintainer, Support
Frequency: Routine, Periodic, Contingency
ILS Elements: Training, Technical Publications, Support Equipment

Generate 8-12 tasks aligned with capability requirements.
Return ONLY the JSON, no other text."""

    else:  # UK_DSAT default
        prompt = f"""Generate a Role Performance Statement (RolePS) for:

Role: {role_title}
Framework: UK DSAT (JSP 822, DTSM 2)
Context: {description or 'Standard role requirements'}

Return a JSON object with this EXACT structure:
{{
    "header": {{
        "security_classification": "OFFICIAL",
        "role_title": "{role_title}",
        "role_number": "2025/001",
        "duty_title": "Primary duty description",
        "duty_number": "1",
        "tra": "Training Requirements Authority",
        "tda": "Training Delivery Authority",
        "roleps_reference": "RPS-2025-001",
        "issue_status": "Draft v1.0"
    }},
    "tasks": [
        {{
            "task_number": "1.1",
            "performance": "Clear task performance statement",
            "conditions": ["Environment condition", "Equipment condition", "Situation condition"],
            "standards": ["Standard 1", "Standard 2"],
            "training_category": "3",
            "notes": ["Knowledge: Required knowledge", "Skill: Required skill", "Attitude: Required attitude"]
        }}
    ]
}}

Training categories: 1 (Pre-joining), 2 (Phase 1), 3 (Phase 2), 4 (Workplace)
Generate 8-12 tasks with realistic content.
Return ONLY the JSON, no other text."""

    response = await call_claude(prompt, max_tokens=6000)
    return parse_json(response)


async def generate_needs_report(role_title: str, framework: str, task_list: Dict, terms: Dict) -> Dict:
    """Generate framework-specific training needs report"""
    num_tasks = len(task_list.get("tasks", []))
    
    if framework == "US_TRADOC":
        prompt = f"""Generate a Training Requirements Document (TRD) for:

Role: {role_title}
Framework: US TRADOC (TP 350-70-1)
Tasks analysed: {num_tasks}

Return a JSON object with:
{{
    "executive_summary": "3-4 paragraphs summarizing requirements",
    "mission_analysis": {{
        "mission_statement": "Mission statement",
        "critical_tasks": ["Critical task 1", "Critical task 2"],
        "enabling_tasks": ["Enabling task 1", "Enabling task 2"]
    }},
    "ctssb_results": "Critical Task and Site Selection Board results summary",
    "ictl_summary": "Summary of Individual Critical Task List analysis",
    "training_site_recommendations": [
        {{"site": "Site name", "rationale": "Why this site"}}
    ],
    "resource_requirements": {{
        "personnel": "Personnel requirements",
        "equipment": "Equipment requirements",
        "facilities": "Facility requirements",
        "budget_estimate": 75000
    }},
    "dotmlpfp_impact": {{
        "doctrine": "Doctrine impact",
        "organization": "Organization impact",
        "training": "Training impact",
        "materiel": "Materiel impact",
        "leadership": "Leadership impact",
        "personnel": "Personnel impact",
        "facilities": "Facilities impact",
        "policy": "Policy impact"
    }}
}}

Return ONLY the JSON, no other text."""

    elif framework == "KIRKPATRICK":
        prompt = f"""Generate an ROE (Return on Expectations) Statement for:

Role: {role_title}
Framework: Kirkpatrick Four Levels
Tasks analysed: {num_tasks}

Return a JSON object with:
{{
    "executive_summary": "3-4 paragraphs on expected return",
    "sponsor_expectations": [
        {{"expectation": "What sponsor expects", "success_indicator": "How we'll know it's achieved"}}
    ],
    "level_4_metrics": [
        {{"metric": "Business metric", "baseline": "Current value", "target": "Target value", "timeline": "When"}}
    ],
    "level_3_behaviors": [
        {{"behavior": "Observable behavior", "measurement": "How measured"}}
    ],
    "level_2_learning": [
        {{"learning": "What they'll learn", "assessment": "How assessed"}}
    ],
    "level_1_engagement": "How we'll ensure engagement and relevance",
    "chain_of_evidence": "How we'll link all four levels to demonstrate value",
    "recommendations": [
        {{"id": "R1", "recommendation": "Specific recommendation", "priority": "High"}}
    ]
}}

Remember: Start with Level 4 Results, work backward.
Return ONLY the JSON, no other text."""

    elif framework == "AUSTRALIAN_SADL":
        prompt = f"""Generate an Analysis Phase Report for:

Role: {role_title}
Framework: Australian SADL (Defence Learning Manual)
Tasks analysed: {num_tasks}

Return a JSON object with:
{{
    "executive_summary": "3-4 paragraphs summarizing analysis findings",
    "target_population_profile": {{
        "learner_subgroups": ["Subgroup descriptions"],
        "motivations": "Key learner motivations",
        "cultural_norms": "Relevant cultural considerations",
        "prior_knowledge": "Expected prior knowledge levels",
        "learning_preferences": "Preferred learning approaches"
    }},
    "performance_gap_analysis": {{
        "current_performance": "Description of current performance",
        "desired_performance": "Description of desired performance",
        "gap_description": "Analysis of the performance gap",
        "root_causes": ["Cause 1", "Cause 2", "Cause 3"]
    }},
    "task_analysis_summary": {{
        "total_tasks": {num_tasks},
        "critical_tasks": "Number of critical tasks",
        "training_domains": {{"institutional": 0, "workplace": 0, "self_directed": 0}}
    }},
    "constraints": ["Constraint 1", "Constraint 2", "Constraint 3"],
    "recommendations": [
        {{"id": "R1", "recommendation": "Specific recommendation", "priority": "High", "rationale": "Why"}}
    ],
    "adele_integration": {{
        "modules_required": "Estimated ADELE modules",
        "integration_notes": "Notes on ADELE LMS integration"
    }}
}}

IMPORTANT: Australian SADL requires summative assessment at Level 4.
Return ONLY the JSON, no other text."""

    elif framework == "ISO_29990":
        prompt = f"""Generate a Learning Needs Analysis Record for:

Role: {role_title}
Framework: ISO 29990 (Learning Services)
Learning needs identified: {num_tasks}

Return a JSON object with:
{{
    "executive_summary": "3-4 paragraphs summarizing learning needs",
    "client_requirements": {{
        "organization": "Client organization",
        "business_context": "Business context for learning",
        "success_criteria": "How client will measure success"
    }},
    "learner_needs": {{
        "current_competence": "Current competence assessment",
        "desired_outcomes": ["Outcome 1", "Outcome 2"],
        "constraints": ["Time", "Budget", "Access"]
    }},
    "gap_analysis": {{
        "competence_gaps": ["Gap 1", "Gap 2", "Gap 3"],
        "priority_gaps": ["Highest priority gaps"],
        "gap_causes": ["Root cause 1", "Root cause 2"]
    }},
    "proposed_learning_services": [
        {{"service_id": "LS1", "service_type": "Workshop", "description": "Service description", "duration": "2 days", "delivery_mode": "Face-to-face"}}
    ],
    "quality_requirements": {{
        "assessment_approach": "How learning will be assessed",
        "evaluation_method": "How service will be evaluated",
        "continuous_improvement": "How feedback will improve service"
    }},
    "recommendations": [
        {{"id": "R1", "recommendation": "Specific recommendation", "priority": "High"}}
    ]
}}

Return ONLY the JSON, no other text."""

    elif framework == "S6000T":
        prompt = f"""Generate a Training Analysis Report for:

Role: {role_title}
Framework: ASD/AIA S6000T (Training Analysis and Design)
Tasks analysed: {num_tasks}

Return a JSON object with:
{{
    "executive_summary": "3-4 paragraphs summarizing training analysis",
    "capability_analysis": {{
        "capability_requirement": "Operational capability being addressed",
        "capability_gap": "Identified gap in capability",
        "training_contribution": "How training addresses the gap"
    }},
    "task_analysis_summary": {{
        "total_tasks": {num_tasks},
        "operator_tasks": 0,
        "maintainer_tasks": 0,
        "support_tasks": 0
    }},
    "ils_integration": {{
        "training_ils_element": "Training requirements",
        "technical_publications": "Related documentation needs",
        "support_equipment": "Training equipment requirements",
        "manpower": "Personnel requirements"
    }},
    "training_requirements": [
        {{"req_id": "TR1", "requirement": "Training requirement", "task_reference": "TSK-001", "ils_link": "Training"}}
    ],
    "resource_implications": {{
        "facilities": "Facility requirements",
        "equipment": "Equipment requirements",
        "personnel": "Personnel requirements",
        "budget_estimate": 75000
    }},
    "recommendations": [
        {{"id": "R1", "recommendation": "Specific recommendation", "priority": "High", "ils_impact": "Training"}}
    ]
}}

S6000T requires traceability to capability requirements and ILS integration.
Return ONLY the JSON, no other text."""

    else:  # UK_DSAT and others
        prompt = f"""Generate a {terms['needs_report']} for:

Role: {role_title}
Framework: {framework}
Tasks analysed: {num_tasks}

Return a JSON object with:
{{
    "executive_summary": "3-4 paragraphs summarizing the training needs analysis",
    "introduction": {{
        "purpose": "Purpose of this report",
        "scope": "Scope of analysis",
        "methodology": "Methodology used"
    }},
    "key_findings": ["Finding 1", "Finding 2", "Finding 3", "Finding 4", "Finding 5"],
    "training_requirements": [
        {{"id": "TR1", "requirement": "Specific requirement", "priority": "Critical", "delivery_method": "Blended"}}
    ],
    "recommendations": [
        {{"id": "R1", "recommendation": "Specific recommendation", "rationale": "Why", "priority": "High", "timeline": "Short-term"}}
    ],
    "resource_requirements": {{
        "budget_estimate": 75000,
        "timeline_months": 6,
        "personnel_required": "2 trainers, 1 training manager"
    }},
    "conclusion": "2-3 paragraphs with conclusions and next steps"
}}

Generate 5 findings, 6 requirements, and 5 recommendations.
Return ONLY the JSON, no other text."""

    response = await call_claude(prompt, max_tokens=5000)
    return parse_json(response)


# ============================================================================
# RESEARCH-BASED ANALYSIS AGENT (v5.1)
# ============================================================================

def get_research_system_prompt() -> str:
    """System prompt for research-based analysis with strict anti-fabrication rules"""
    return """You are NOVA, a professional training analyst conducting research-based job/task analysis.

CRITICAL RULES - ZERO TOLERANCE FOR FABRICATION:
1. Every factual claim MUST be based on information found through web search
2. NEVER invent statistics, percentages, or quantified claims
3. NEVER fabricate methodology (no fake interviews, surveys, or focus groups)
4. NEVER hallucinate professional standards, regulations, or requirements
5. If information is not found through research, state: "Information not found through research"
6. All sources MUST be cited with URL and access date

OUTPUT FORMAT:
Return a JSON object. All text fields should contain factual, researched information with inline citations in format [Source Name, URL].

RESEARCH METHODOLOGY:
Use web search to find authoritative sources for:
- Professional body standards and requirements
- Competency frameworks (NOS, SFIA, NHS KSF, etc.)
- Legal and regulatory requirements
- Qualification requirements
- Industry standards and best practices

CITATION FORMAT:
Every factual statement must include source reference: "[Source: Organization Name - URL]"
"""


async def run_research_analysis_agent(job_id: str, parameters: Dict, framework: str):
    """
    Research-based Analysis Agent (v5.1)
    
    Uses web search to gather factual information about roles, producing
    outputs with full citations. Zero fabrication tolerance.
    
    Generates:
    - 01_Job_Task_Analysis.docx - Framework-compliant task list with sources
    - 02_Analysis_Report.docx - 18-section report with full citations
    - analysis_data.json - Raw JSON for reference
    """
    domain = parameters.get("domain", "General")
    specialism = parameters.get("specialism", "")
    role_title = parameters.get("role_title", "Training Specialist")
    proficiency_level = parameters.get("proficiency_level", "Mid-Level")
    role_description = parameters.get("role_description", "")
    
    terms = get_terminology(framework)
    
    output_dir = Path(jobs.get(job_id)["output_dir"]) / "01_Analysis"
    output_dir.mkdir(exist_ok=True)
    
    update_job(job_id, 2, f"Starting Research-Based Analysis ({framework})...")
    
    # Build the research prompt with all 10 research steps
    research_prompt = build_research_prompt(
        domain=domain,
        specialism=specialism,
        role_title=role_title,
        proficiency_level=proficiency_level,
        framework=framework,
        role_description=role_description,
        terms=terms
    )
    
    # Execute research with web search
    update_job(job_id, 5, "Step 1/10: Researching framework requirements...")
    
    try:
        research_result = await call_claude_with_search(
            system_prompt=get_research_system_prompt(),
            user_prompt=research_prompt,
            max_tokens=12000
        )
        
        update_job(job_id, 40, f"Research complete. {len(research_result.get('searches_performed', []))} searches performed.")
        
        # Parse the research output
        analysis_data = parse_research_output(research_result.get("text", ""))
        
        # Add metadata
        analysis_data["metadata"] = {
            "domain": domain,
            "specialism": specialism,
            "role_title": role_title,
            "proficiency_level": proficiency_level,
            "framework": framework,
            "framework_display": terms.get("framework_name", framework),
            "generated_date": datetime.now().isoformat(),
            "searches_performed": research_result.get("searches_performed", []),
            "nova_version": "5.1.0"
        }
        
    except Exception as e:
        print(f"[NOVA] Research failed: {e}")
        update_job(job_id, 40, f"Research encountered issues, generating with available data...")
        analysis_data = create_fallback_analysis(parameters, framework, terms)
    
    # Build documents
    update_job(job_id, 50, f"Building {terms['task_list']} document...")
    build_research_task_list_doc(
        analysis_data, 
        role_title, 
        framework, 
        terms, 
        output_dir / "01_Job_Task_Analysis.docx"
    )
    update_job(job_id, 70, f"✓ {terms['task_list']} complete")
    
    update_job(job_id, 75, "Building Analysis Report document...")
    build_research_analysis_report_doc(
        analysis_data,
        role_title,
        framework,
        terms,
        output_dir / "02_Analysis_Report.docx"
    )
    update_job(job_id, 90, "✓ Analysis Report complete")
    
    # Save raw JSON
    update_job(job_id, 95, "Saving analysis data...")
    with open(output_dir / "analysis_data.json", "w") as f:
        json.dump(analysis_data, f, indent=2, default=str)
    
    update_job(job_id, 100, "Analysis Phase Complete")


def build_research_prompt(domain: str, specialism: str, role_title: str, 
                          proficiency_level: str, framework: str, 
                          role_description: str, terms: Dict) -> str:
    """Build the comprehensive research prompt for the 10-step methodology"""
    
    framework_guidance = get_framework_research_guidance(framework, terms)
    
    return f"""Conduct comprehensive research-based analysis for the following role:

ANALYSIS PARAMETERS:
- Domain: {domain}
- Specialism: {specialism}
- Role Title: {role_title}
- Proficiency Level: {proficiency_level}
- Framework: {framework} ({terms.get('framework_name', framework)})
- Additional Context: {role_description or 'None provided'}

RESEARCH METHODOLOGY - Complete all 10 steps:

STEP 1: FRAMEWORK ISOLATION
Research the specific requirements of {framework}:
{framework_guidance}

STEP 2: DOMAIN RESEARCH
Search for "{domain}" industry standards, regulations, and training requirements.
Find: Industry bodies, regulatory requirements, common certifications.

STEP 3: SPECIALISM RESEARCH  
Search for "{specialism}" competency frameworks and professional standards.
Find: Specialist qualifications, technical standards, professional bodies.

STEP 4: ROLE TITLE RESEARCH
Search for "{role_title}" job descriptions, responsibilities, and competencies.
Find: Standard role definitions, typical duties, required competencies.

STEP 5: PROFICIENCY LEVEL MAPPING
Map "{proficiency_level}" to established frameworks:
- SFIA levels (if IT/Digital)
- NVQ/RQF levels (if vocational)
- NHS AfC bands (if healthcare)
- Military ranks/grades (if defence)
- Professional body grades

STEP 6: PROFESSIONAL BODY RESEARCH
Search for professional bodies and regulators for "{role_title}" in "{domain}".
Find: Registration requirements, CPD requirements, codes of conduct.

STEP 7: COMPETENCY FRAMEWORK MAPPING
Search for National Occupational Standards (NOS) or competency frameworks for "{specialism}".
Find: Specific competency units, performance criteria, knowledge requirements.

STEP 8: LEGAL/COMPLIANCE RESEARCH
Search for legal requirements for "{role_title}" roles.
Find: Statutory training, mandatory qualifications, compliance requirements.

STEP 9: PHYSICAL/MEDICAL/SECURITY REQUIREMENTS
Search for any physical, medical, or security requirements for "{role_title}".
Find: Fitness standards, health requirements, security clearance levels.

STEP 10: CPD/RECERTIFICATION RESEARCH
Search for continuing professional development requirements for "{role_title}".
Find: Revalidation periods, CPD hours, recertification requirements.

OUTPUT FORMAT - Return a JSON object with this exact structure:

{{
    "executive_summary": "2-3 paragraphs summarizing key findings with citations",
    
    "framework_analysis": {{
        "framework_name": "{terms.get('framework_name', framework)}",
        "key_requirements": ["Requirement 1 [Source]", "Requirement 2 [Source]"],
        "terminology_used": "{terms['task_list']} / {terms['top_objective']} / etc."
    }},
    
    "geographic_context": {{
        "jurisdiction": "UK/US/International",
        "regulatory_body": "Name of regulator [Source: URL]",
        "applicable_legislation": ["Act 1 [Source]", "Act 2 [Source]"]
    }},
    
    "professional_body": {{
        "name": "Professional body name",
        "registration_required": true/false,
        "registration_url": "URL",
        "code_of_conduct_url": "URL"
    }},
    
    "competency_framework": {{
        "framework_name": "NOS/SFIA/NHS KSF/etc.",
        "framework_url": "URL",
        "relevant_units": [
            {{"unit_code": "Code", "unit_title": "Title", "source": "URL"}}
        ]
    }},
    
    "role_profile": {{
        "standard_definition": "Standard role definition [Source]",
        "typical_reporting_line": "Reports to...",
        "typical_team_size": "X direct reports"
    }},
    
    "qualifications": {{
        "mandatory": ["Qualification 1 [Source]"],
        "desirable": ["Qualification 2 [Source]"],
        "professional_registration": "Required/Desirable/Not required [Source]"
    }},
    
    "experience": {{
        "minimum_years": "X years [Source]",
        "required_experience": ["Experience area 1 [Source]"],
        "desirable_experience": ["Experience area 2 [Source]"]
    }},
    
    "technical_skills": [
        {{"skill": "Skill name", "proficiency": "Level", "source": "URL"}}
    ],
    
    "soft_skills": [
        {{"skill": "Skill name", "importance": "Critical/Important/Desirable", "source": "URL"}}
    ],
    
    "behaviours": [
        {{"behaviour": "Behaviour description", "source": "URL"}}
    ],
    
    "physical_medical_security": {{
        "physical_requirements": ["Requirement [Source]"],
        "medical_requirements": ["Requirement [Source]"],
        "security_clearance": "Level required [Source]"
    }},
    
    "cpd_requirements": {{
        "annual_hours": "X hours [Source]",
        "revalidation_period": "X years [Source]",
        "activities": ["Activity type [Source]"]
    }},
    
    "career_progression": {{
        "typical_next_role": "Role title",
        "progression_requirements": ["Requirement [Source]"]
    }},
    
    "legal_compliance": [
        {{"requirement": "Legal requirement", "legislation": "Act name", "source": "URL"}}
    ],
    
    "professional_standards": [
        {{"standard": "Standard description", "source": "URL"}}
    ],
    
    "tasks": [
        {{
            "task_id": "TSK-001",
            "task_description": "Task description based on research [Source]",
            "knowledge_required": ["Knowledge item [Source]"],
            "skills_required": ["Skill item [Source]"],
            "criticality": "High/Medium/Low",
            "frequency": "Daily/Weekly/Monthly/As required",
            "source": "URL where this task was found"
        }}
    ],
    
    "citations": [
        {{
            "id": "1",
            "source_name": "Organization Name",
            "url": "https://...",
            "access_date": "{datetime.now().strftime('%Y-%m-%d')}",
            "description": "What information was obtained"
        }}
    ]
}}

CRITICAL REMINDERS:
- Every factual claim must have a citation
- Use web search for each research step
- If information cannot be found, state "Not found through research"
- Do not invent or fabricate any information
- Include full URLs for all sources"""


def get_framework_research_guidance(framework: str, terms: Dict) -> str:
    """Get framework-specific research guidance"""
    
    guidance = {
        "UK_DSAT": """Search for:
- JSP 822 Defence Individual Training policy requirements
- DTSM 2 Analysis of Individual Training requirements
- Role Performance Statement (RolePS) format and structure
- Training Needs Analysis (TNA) requirements
- Knowledge, Skills, Attitudes (KSA) categorisation""",

        "US_TRADOC": """Search for:
- TRADOC Regulation 350-70 training development requirements
- Individual Critical Task List (ICTL) format
- Task analysis methodology (CAR: Condition-Action-Standard)
- Training domain categorisation
- Skill level progression requirements""",

        "NATO_BISC": """Search for:
- NATO Bi-SCD 075-007 Education and Training requirements
- Standard Training Plan (STP) format
- Task analysis requirements
- Proficiency level definitions
- Interoperability training requirements""",

        "ADDIE": """Search for:
- ADDIE model Analysis phase requirements
- Job Task Analysis methodology
- Task inventory development
- Performance gap analysis
- Training needs assessment best practices""",

        "KIRKPATRICK": """Search for:
- Kirkpatrick Four-Level Evaluation Model
- Training needs assessment alignment with evaluation
- Performance-based training analysis
- Measurable learning outcomes development""",

        "ACTION_MAPPING": """Search for:
- Cathy Moore Action Mapping methodology
- Performance-focused analysis approach
- Business goal identification
- Action-based task analysis
- Minimum information principle""",

        "S6000T": """Search for:
- ASD/AIA S6000T Training Analysis and Design specification
- Task analysis requirements in S-Series ILS
- Training requirements traceability
- Capability-based training analysis""",

        "SAM": """Search for:
- Successive Approximation Model (SAM) requirements
- Rapid prototyping analysis phase
- Iterative needs assessment
- Stakeholder collaboration requirements""",

        "ISO_29990": """Search for:
- ISO 29990 Learning services requirements
- Needs analysis requirements
- Learner needs determination
- Learning service design inputs""",

        "AUSTRALIAN_SADL": """Search for:
- Australian Defence Force SADL requirements
- Systematic approach to training analysis
- Training needs analysis requirements
- Competency-based training analysis"""
    }
    
    return guidance.get(framework, """Search for:
- Framework-specific analysis requirements
- Task analysis methodology
- Competency identification requirements
- Training needs assessment standards""")


def parse_research_output(text: str) -> Dict:
    """Parse research output, extracting JSON from response"""
    try:
        # Try to find JSON in the response
        json_match = re.search(r'\{[\s\S]*\}', text)
        if json_match:
            return json.loads(json_match.group())
    except json.JSONDecodeError as e:
        print(f"[NOVA] JSON parse error: {e}")
    
    # Return minimal structure if parsing fails
    return {
        "executive_summary": text[:2000] if text else "Analysis could not be completed.",
        "tasks": [],
        "citations": [],
        "parse_error": True
    }


def create_fallback_analysis(parameters: Dict, framework: str, terms: Dict) -> Dict:
    """Create fallback analysis structure when research fails"""
    return {
        "executive_summary": f"Analysis for {parameters.get('role_title', 'Role')} in {parameters.get('domain', 'Domain')}. Research-based analysis was not completed. This document requires manual completion with verified sources.",
        "framework_analysis": {
            "framework_name": terms.get("framework_name", framework),
            "key_requirements": ["Manual research required"],
            "terminology_used": f"{terms['task_list']} / {terms['top_objective']}"
        },
        "tasks": [],
        "citations": [],
        "fallback_mode": True,
        "metadata": {
            "domain": parameters.get("domain", ""),
            "specialism": parameters.get("specialism", ""),
            "role_title": parameters.get("role_title", ""),
            "proficiency_level": parameters.get("proficiency_level", ""),
            "framework": framework,
            "generated_date": datetime.now().isoformat(),
            "nova_version": "5.1.0",
            "note": "Fallback mode - research could not be completed"
        }
    }


def build_research_task_list_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build research-based task list document with citations"""
    doc = Document()
    
    # Set to landscape for task lists
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width = section.page_height
    new_height = section.page_width
    section.page_width = new_width
    section.page_height = new_height
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    
    metadata = data.get("metadata", {})
    
    # Security Classification
    p = doc.add_paragraph()
    run = p.add_run("OFFICIAL")
    run.font.name = "Roboto"
    run.font.size = Pt(12)
    run.font.bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Title
    title = doc.add_heading(f"{terms['task_list']} - RESEARCH BASED", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    
    # Subtitle
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(14)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Framework badge
    badge = doc.add_paragraph()
    run = badge.add_run(f"Framework: {terms.get('framework_name', framework)} | Generated: {metadata.get('generated_date', '')[:10]}")
    run.font.name = "Roboto"
    run.font.size = Pt(10)
    run.font.italic = True
    badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Analysis Parameters Table
    create_styled_heading(doc, "Analysis Parameters", 1)
    param_table = doc.add_table(rows=6, cols=2)
    param_table.style = 'Table Grid'
    
    params = [
        ("Domain", metadata.get("domain", "")),
        ("Specialism", metadata.get("specialism", "")),
        ("Role Title", metadata.get("role_title", "")),
        ("Proficiency Level", metadata.get("proficiency_level", "")),
        ("Framework", metadata.get("framework_display", "")),
        ("Research Date", metadata.get("generated_date", "")[:10])
    ]
    
    for i, (label, value) in enumerate(params):
        param_table.rows[i].cells[0].text = label
        param_table.rows[i].cells[1].text = str(value)
        for cell in param_table.rows[i].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Task Table
    create_styled_heading(doc, "Job/Task Inventory", 1)
    
    tasks = data.get("tasks", [])
    
    if tasks:
        task_table = doc.add_table(rows=1, cols=6)
        task_table.style = 'Table Grid'
        
        headers = ["Task ID", "Task Description", "Knowledge Required", "Skills Required", "Criticality", "Frequency"]
        header_row = task_table.rows[0]
        set_repeat_table_header(header_row)
        
        for i, h in enumerate(headers):
            cell = header_row.cells[i]
            cell.text = h
            set_cell_shading(cell, "8B4545")
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(10)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
        
        for task_idx, task in enumerate(tasks):
            row = task_table.add_row()
            row.cells[0].text = str(task.get("task_id", f"TSK-{task_idx+1:03d}"))
            row.cells[1].text = task.get("task_description", "")
            
            knowledge = task.get("knowledge_required", [])
            row.cells[2].text = "\n".join(knowledge) if isinstance(knowledge, list) else str(knowledge)
            
            skills = task.get("skills_required", [])
            row.cells[3].text = "\n".join(skills) if isinstance(skills, list) else str(skills)
            
            row.cells[4].text = task.get("criticality", "Medium")
            row.cells[5].text = task.get("frequency", "As required")
            
            if task_idx % 2 == 1:
                for cell in row.cells:
                    set_cell_shading(cell, "F5F5F5")
            
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Roboto"
                        run.font.size = Pt(9)
    else:
        p = doc.add_paragraph()
        run = p.add_run("No tasks identified through research. Manual task analysis required.")
        run.font.name = "Roboto"
        run.font.italic = True
    
    doc.add_paragraph()
    
    # Research Sources
    create_styled_heading(doc, "Research Sources", 1)
    citations = data.get("citations", [])
    
    if citations:
        for citation in citations:
            p = doc.add_paragraph()
            run = p.add_run(f"[{citation.get('id', '')}] {citation.get('source_name', '')}")
            run.font.name = "Roboto"
            run.font.bold = True
            run.font.size = Pt(10)
            
            p2 = doc.add_paragraph()
            run2 = p2.add_run(f"    URL: {citation.get('url', 'N/A')}")
            run2.font.name = "Roboto"
            run2.font.size = Pt(9)
            
            p3 = doc.add_paragraph()
            run3 = p3.add_run(f"    Accessed: {citation.get('access_date', 'N/A')}")
            run3.font.name = "Roboto"
            run3.font.size = Pt(9)
    else:
        p = doc.add_paragraph()
        run = p.add_run("No citations recorded. Manual verification required.")
        run.font.name = "Roboto"
        run.font.italic = True
    
    # Disclaimer
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("DISCLAIMER: This document was generated using AI-assisted research. All information should be independently verified before use in formal training documentation.")
    run.font.name = "Roboto"
    run.font.size = Pt(8)
    run.font.italic = True
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_research_analysis_report_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build comprehensive 18-section analysis report with full citations"""
    doc = Document()
    
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    
    metadata = data.get("metadata", {})
    
    # Title Page
    doc.add_paragraph()
    doc.add_paragraph()
    
    title = doc.add_heading("ANALYSIS REPORT", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.size = Pt(28)
        run.font.color.rgb = NOVA_DARK_BLUE
    
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(18)
    run.font.color.rgb = NOVA_DARK_BLUE
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    info = doc.add_paragraph()
    run = info.add_run(f"Framework: {terms.get('framework_name', framework)}\n")
    run.font.name = "Roboto"
    run = info.add_run(f"Domain: {metadata.get('domain', 'N/A')}\n")
    run.font.name = "Roboto"
    run = info.add_run(f"Specialism: {metadata.get('specialism', 'N/A')}\n")
    run.font.name = "Roboto"
    run = info.add_run(f"Proficiency Level: {metadata.get('proficiency_level', 'N/A')}\n")
    run.font.name = "Roboto"
    run = info.add_run(f"Generated: {metadata.get('generated_date', '')[:10]}\n")
    run.font.name = "Roboto"
    run = info.add_run(f"NOVA Version: {metadata.get('nova_version', '5.1.0')}")
    run.font.name = "Roboto"
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_page_break()
    
    # Section 1: Executive Summary
    create_styled_heading(doc, "1. EXECUTIVE SUMMARY", 1)
    exec_summary = data.get("executive_summary", "No executive summary available.")
    sentences = split_into_sentences(exec_summary)
    for sentence in sentences:
        if sentence.strip():
            create_styled_paragraph(doc, sentence.strip())
    
    # Section 2: Framework Identification
    create_styled_heading(doc, "2. FRAMEWORK IDENTIFICATION", 1)
    framework_analysis = data.get("framework_analysis", {})
    create_styled_paragraph(doc, f"Framework: {framework_analysis.get('framework_name', framework)}", bold=True)
    
    key_reqs = framework_analysis.get("key_requirements", [])
    if key_reqs:
        create_styled_paragraph(doc, "Key Requirements:", bold=True)
        for req in key_reqs:
            create_styled_paragraph(doc, f"  • {req}")
    
    create_styled_paragraph(doc, f"Terminology: {framework_analysis.get('terminology_used', 'Standard terminology')}")
    
    # Section 3: Geographic/Jurisdictional Context
    create_styled_heading(doc, "3. GEOGRAPHIC/JURISDICTIONAL CONTEXT", 1)
    geo = data.get("geographic_context", {})
    create_styled_paragraph(doc, f"Jurisdiction: {geo.get('jurisdiction', 'Not specified')}")
    create_styled_paragraph(doc, f"Regulatory Body: {geo.get('regulatory_body', 'Not identified')}")
    
    legislation = geo.get("applicable_legislation", [])
    if legislation:
        create_styled_paragraph(doc, "Applicable Legislation:", bold=True)
        for leg in legislation:
            create_styled_paragraph(doc, f"  • {leg}")
    
    # Section 4: Professional Body/Regulator
    create_styled_heading(doc, "4. PROFESSIONAL BODY/REGULATOR", 1)
    prof_body = data.get("professional_body", {})
    create_styled_paragraph(doc, f"Professional Body: {prof_body.get('name', 'Not identified')}")
    create_styled_paragraph(doc, f"Registration Required: {'Yes' if prof_body.get('registration_required') else 'No'}")
    if prof_body.get("registration_url"):
        create_styled_paragraph(doc, f"Registration URL: {prof_body.get('registration_url')}")
    if prof_body.get("code_of_conduct_url"):
        create_styled_paragraph(doc, f"Code of Conduct: {prof_body.get('code_of_conduct_url')}")
    
    # Section 5: Competency Framework Mapping
    create_styled_heading(doc, "5. COMPETENCY FRAMEWORK MAPPING", 1)
    comp_framework = data.get("competency_framework", {})
    create_styled_paragraph(doc, f"Framework: {comp_framework.get('framework_name', 'Not identified')}")
    if comp_framework.get("framework_url"):
        create_styled_paragraph(doc, f"URL: {comp_framework.get('framework_url')}")
    
    units = comp_framework.get("relevant_units", [])
    if units:
        create_styled_paragraph(doc, "Relevant Competency Units:", bold=True)
        for unit in units:
            create_styled_paragraph(doc, f"  • {unit.get('unit_code', 'N/A')}: {unit.get('unit_title', 'N/A')}")
    
    # Section 6: Role Description
    create_styled_heading(doc, "6. ROLE DESCRIPTION", 1)
    role_profile = data.get("role_profile", {})
    create_styled_paragraph(doc, role_profile.get("standard_definition", "No standard definition found through research."))
    create_styled_paragraph(doc, f"Typical Reporting Line: {role_profile.get('typical_reporting_line', 'Not specified')}")
    create_styled_paragraph(doc, f"Typical Team Size: {role_profile.get('typical_team_size', 'Not specified')}")
    
    # Section 7: Qualifications
    create_styled_heading(doc, "7. QUALIFICATIONS", 1)
    quals = data.get("qualifications", {})
    
    mandatory = quals.get("mandatory", [])
    if mandatory:
        create_styled_paragraph(doc, "Mandatory Qualifications:", bold=True)
        for qual in mandatory:
            create_styled_paragraph(doc, f"  • {qual}")
    
    desirable = quals.get("desirable", [])
    if desirable:
        create_styled_paragraph(doc, "Desirable Qualifications:", bold=True)
        for qual in desirable:
            create_styled_paragraph(doc, f"  • {qual}")
    
    create_styled_paragraph(doc, f"Professional Registration: {quals.get('professional_registration', 'Not specified')}")
    
    # Section 8: Experience
    create_styled_heading(doc, "8. EXPERIENCE", 1)
    exp = data.get("experience", {})
    create_styled_paragraph(doc, f"Minimum Experience: {exp.get('minimum_years', 'Not specified')}")
    
    req_exp = exp.get("required_experience", [])
    if req_exp:
        create_styled_paragraph(doc, "Required Experience:", bold=True)
        for e in req_exp:
            create_styled_paragraph(doc, f"  • {e}")
    
    des_exp = exp.get("desirable_experience", [])
    if des_exp:
        create_styled_paragraph(doc, "Desirable Experience:", bold=True)
        for e in des_exp:
            create_styled_paragraph(doc, f"  • {e}")
    
    # Section 9: Technical Skills
    create_styled_heading(doc, "9. TECHNICAL SKILLS", 1)
    tech_skills = data.get("technical_skills", [])
    if tech_skills:
        rows = [[s.get("skill", ""), s.get("proficiency", ""), s.get("source", "")] for s in tech_skills]
        create_styled_table(doc, ["Skill", "Proficiency Level", "Source"], rows, [2.5, 1.5, 2.5])
    else:
        create_styled_paragraph(doc, "No technical skills identified through research.")
    
    doc.add_paragraph()
    
    # Section 10: Soft Skills
    create_styled_heading(doc, "10. SOFT SKILLS", 1)
    soft_skills = data.get("soft_skills", [])
    if soft_skills:
        rows = [[s.get("skill", ""), s.get("importance", ""), s.get("source", "")] for s in soft_skills]
        create_styled_table(doc, ["Skill", "Importance", "Source"], rows, [2.5, 1.5, 2.5])
    else:
        create_styled_paragraph(doc, "No soft skills identified through research.")
    
    doc.add_paragraph()
    
    # Section 11: Personal Traits/Behaviours
    create_styled_heading(doc, "11. PERSONAL TRAITS AND BEHAVIOURS", 1)
    behaviours = data.get("behaviours", [])
    if behaviours:
        for b in behaviours:
            create_styled_paragraph(doc, f"  • {b.get('behaviour', '')} [{b.get('source', 'No source')}]")
    else:
        create_styled_paragraph(doc, "No specific behaviours identified through research.")
    
    # Section 12: Physical/Medical/Security Requirements
    create_styled_heading(doc, "12. PHYSICAL, MEDICAL AND SECURITY REQUIREMENTS", 1)
    pms = data.get("physical_medical_security", {})
    
    physical = pms.get("physical_requirements", [])
    if physical:
        create_styled_paragraph(doc, "Physical Requirements:", bold=True)
        for req in physical:
            create_styled_paragraph(doc, f"  • {req}")
    
    medical = pms.get("medical_requirements", [])
    if medical:
        create_styled_paragraph(doc, "Medical Requirements:", bold=True)
        for req in medical:
            create_styled_paragraph(doc, f"  • {req}")
    
    create_styled_paragraph(doc, f"Security Clearance: {pms.get('security_clearance', 'Not specified')}")
    
    # Section 13: CPD/Recertification Requirements
    create_styled_heading(doc, "13. CPD AND RECERTIFICATION REQUIREMENTS", 1)
    cpd = data.get("cpd_requirements", {})
    create_styled_paragraph(doc, f"Annual CPD Hours: {cpd.get('annual_hours', 'Not specified')}")
    create_styled_paragraph(doc, f"Revalidation Period: {cpd.get('revalidation_period', 'Not specified')}")
    
    activities = cpd.get("activities", [])
    if activities:
        create_styled_paragraph(doc, "CPD Activities:", bold=True)
        for act in activities:
            create_styled_paragraph(doc, f"  • {act}")
    
    # Section 14: Career Progression Context
    create_styled_heading(doc, "14. CAREER PROGRESSION CONTEXT", 1)
    career = data.get("career_progression", {})
    create_styled_paragraph(doc, f"Typical Next Role: {career.get('typical_next_role', 'Not specified')}")
    
    prog_reqs = career.get("progression_requirements", [])
    if prog_reqs:
        create_styled_paragraph(doc, "Progression Requirements:", bold=True)
        for req in prog_reqs:
            create_styled_paragraph(doc, f"  • {req}")
    
    # Section 15: Legal Compliance
    create_styled_heading(doc, "15. LEGAL COMPLIANCE", 1)
    legal = data.get("legal_compliance", [])
    if legal:
        rows = [[l.get("requirement", ""), l.get("legislation", ""), l.get("source", "")] for l in legal]
        create_styled_table(doc, ["Requirement", "Legislation", "Source"], rows, [2.5, 2.0, 2.0])
    else:
        create_styled_paragraph(doc, "No specific legal compliance requirements identified through research.")
    
    doc.add_paragraph()
    
    # Section 16: Professional Standards
    create_styled_heading(doc, "16. PROFESSIONAL STANDARDS", 1)
    standards = data.get("professional_standards", [])
    if standards:
        for std in standards:
            create_styled_paragraph(doc, f"  • {std.get('standard', '')} [{std.get('source', 'No source')}]")
    else:
        create_styled_paragraph(doc, "No specific professional standards identified through research.")
    
    # Section 17: No Bias Statement
    create_styled_heading(doc, "17. EQUALITY AND DIVERSITY STATEMENT", 1)
    create_styled_paragraph(doc, "This analysis has been conducted in accordance with the Equality Act 2010 and does not discriminate on the basis of age, disability, gender reassignment, marriage and civil partnership, pregnancy and maternity, race, religion or belief, sex, or sexual orientation.")
    create_styled_paragraph(doc, "All requirements listed are genuine occupational requirements based on the inherent nature of the role and have been identified through objective research of authoritative sources.")
    
    # Section 18: Citations and Sources
    create_styled_heading(doc, "18. CITATIONS AND SOURCES", 1)
    citations = data.get("citations", [])
    
    if citations:
        rows = []
        for c in citations:
            rows.append([
                c.get("id", ""),
                c.get("source_name", ""),
                c.get("url", ""),
                c.get("access_date", "")
            ])
        create_styled_table(doc, ["#", "Source", "URL", "Accessed"], rows, [0.3, 2.0, 3.0, 1.0])
    else:
        create_styled_paragraph(doc, "No citations recorded. Manual verification required.")
    
    # Searches performed
    searches = metadata.get("searches_performed", [])
    if searches:
        doc.add_paragraph()
        create_styled_paragraph(doc, "Research Queries Executed:", bold=True)
        for i, search in enumerate(searches[:20], 1):  # Limit to first 20
            create_styled_paragraph(doc, f"  {i}. {search}")
    
    # Final Disclaimer
    doc.add_paragraph()
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("DISCLAIMER")
    run.font.name = "Roboto"
    run.font.bold = True
    run.font.size = Pt(10)
    
    p2 = doc.add_paragraph()
    run2 = p2.add_run("This document was generated using AI-assisted research conducted on the date shown above. Information may have changed since the research was conducted. All information should be independently verified against current authoritative sources before use in formal training documentation or decision-making. NOVA and its operators accept no liability for decisions made based on this analysis.")
    run2.font.name = "Roboto"
    run2.font.size = Pt(9)
    run2.font.italic = True
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


# ============================================================================
# DESIGN AGENT
# ============================================================================

async def run_design_agent(job_id: str, parameters: Dict, framework: str, templates: Dict):
    """Generate Design phase documents based on framework"""
    role_title = parameters.get("role_title", "Training Specialist")
    role_desc = parameters.get("role_description", "")
    terms = get_terminology(framework)
    
    output_dir = Path(jobs.get(job_id)["output_dir"]) / "02_Design"
    output_dir.mkdir(exist_ok=True)
    
    update_job(job_id, 5, f"Starting Design Agent ({framework})...")
    
    # Generate Learning Objectives
    update_job(job_id, 10, f"Generating {terms['top_objective_short']}s...")
    objectives = await generate_objectives(role_title, framework, role_desc, terms)
    update_job(job_id, 40, f"Building Objectives document...")
    build_objectives_doc(objectives, role_title, framework, terms, output_dir / f"01_{terms['top_objective_short']}_Hierarchy.docx")
    update_job(job_id, 50, f"✓ {terms['top_objective_short']} Hierarchy complete")
    
    # Generate Course Design Document
    update_job(job_id, 55, f"Generating {terms['course_design']}...")
    course_design = await generate_course_design(role_title, framework, objectives, terms)
    update_job(job_id, 85, f"Building {terms['course_design']} document...")
    filename = sanitize_filename(terms['course_design'].split('(')[0].strip())
    build_course_design_doc(course_design, role_title, framework, terms, output_dir / f"02_{filename}.docx")
    update_job(job_id, 95, f"✓ {terms['course_design']} complete")
    
    update_job(job_id, 100, "Design Phase Complete")


async def generate_objectives(role_title: str, framework: str, description: str, terms: Dict) -> Dict:
    """Generate framework-specific learning objectives"""
    
    if framework == "US_TRADOC":
        prompt = f"""Generate Terminal Learning Objectives (TLOs) with Enabling Learning Objectives (ELOs) for:

Role: {role_title}
Framework: US TRADOC (TP 350-70-1)
Format: Action-Condition-Standard

Return a JSON object:
{{
    "objectives": [
        {{
            "tlo_number": "TLO 1",
            "action": "Action verb + object",
            "condition": "Given [equipment, references, situation]",
            "standard": "To standard of [criteria]",
            "elos": [
                {{
                    "elo_number": "ELO 1.1",
                    "action": "Supporting action",
                    "condition": "Given [context]",
                    "standard": "Criteria"
                }}
            ],
            "lsas": [
                {{
                    "lsa_number": "LSA 1.1.1",
                    "activity": "Learning step activity description",
                    "method": "Lecture",
                    "duration_minutes": 30
                }}
            ]
        }}
    ]
}}

Generate 5-6 TLOs, each with 2-4 ELOs, each ELO with 2-3 LSAs.
Use measurable action verbs.
Return ONLY the JSON."""

    elif framework == "NATO_BISC":
        prompt = f"""Generate Performance Objectives (POs) with Enabling Learning Objectives (ELOs) for:

Role: {role_title}
Framework: NATO Bi-SCD 075-007
Format: Performance-Standard (outcome-based, NOT task-based)

Return a JSON object:
{{
    "objectives": [
        {{
            "po_number": "PO 1",
            "performance": "Outcome statement (what the learner will achieve)",
            "standard": {{
                "measure": "Basis for describing performance levels",
                "criterion": "Acceptable performance level"
            }},
            "elos": [
                {{
                    "elo_number": "ELO 1.1",
                    "statement": "Supporting learning objective",
                    "level": "Module"
                }}
            ]
        }}
    ]
}}

Levels: Chapter, Module, Lesson, Entry
Generate 5-6 POs, each with 2-4 ELOs.
Return ONLY the JSON."""

    elif framework == "ADDIE":
        prompt = f"""Generate Learning Objectives in ABCD format for:

Role: {role_title}
Framework: ADDIE Model
Format: Audience-Behavior-Condition-Degree

Return a JSON object:
{{
    "objectives": [
        {{
            "lo_number": "LO 1",
            "audience": "The learner",
            "behavior": "Will [action verb] [object]",
            "condition": "Given [circumstances]",
            "degree": "To [criteria for success]",
            "supporting_objectives": [
                {{
                    "so_number": "SO 1.1",
                    "statement": "Supporting objective statement"
                }}
            ]
        }}
    ]
}}

Generate 5-6 Learning Objectives with 2-4 supporting objectives each.
Return ONLY the JSON."""

    elif framework == "ACTION_MAPPING":
        prompt = f"""Generate Action Statements for:

Role: {role_title}
Framework: Action Mapping (Cathy Moore)
Focus: What people need to DO, not what they need to KNOW

Return a JSON object:
{{
    "business_goal": "Measurable business goal this training supports",
    "actions": [
        {{
            "action_number": "A1",
            "action_statement": "People will [observable action contributing to business goal]",
            "practice_activity": {{
                "scenario": "Realistic scenario description",
                "decision_point": "What decision they need to make",
                "correct_choice": "The right action",
                "consequences": "What happens based on their choice"
            }},
            "minimum_information": "Only the info needed to make the decision"
        }}
    ]
}}

Focus on actions, not information. Training changes what people DO.
Generate 6-8 actions with scenario-based practice activities.
Return ONLY the JSON."""

    else:  # UK_DSAT default
        prompt = f"""Generate Training Objectives (TOs) with Enabling Objectives (EOs) and Key Learning Points (KLPs) for:

Role: {role_title}
Framework: UK DSAT (JSP 822, DTSM 3)
Format: Performance-Conditions-Standards

Return a JSON object:
{{
    "objectives": [
        {{
            "to_number": "TO 1",
            "performance": "The trainee will [action verb] [object]",
            "conditions": "Given [environment, equipment, constraints]",
            "standards": "To the standard of [measurable criteria]",
            "eos": [
                {{
                    "eo_number": "EO 1.1",
                    "performance": "The trainee will [supporting action]",
                    "conditions": "Given [context]",
                    "standards": "To the standard of [criteria]",
                    "ksa_tag": "S"
                }}
            ],
            "klps": [
                {{
                    "klp_number": "KLP 1.1.1",
                    "statement": "Declarative statement of knowledge/skill",
                    "domain": "Knowledge"
                }}
            ]
        }}
    ]
}}

KSA tags: K (Knowledge), S (Skill), A (Attitude)
Domains: Knowledge, Skill, Attitude
Generate 5-6 TOs, each with 2-4 EOs, each EO with 3-5 KLPs.
Return ONLY the JSON."""

    response = await call_claude(prompt, max_tokens=6000)
    return parse_json(response)


async def generate_course_design(role_title: str, framework: str, objectives: Dict, terms: Dict) -> Dict:
    """Generate framework-specific course design document"""
    num_objectives = len(objectives.get("objectives", []))
    
    prompt = f"""Generate a {terms['course_design']} for:

Role: {role_title}
Framework: {framework}
Number of {terms['top_objective_short']}s: {num_objectives}

Return a JSON object with:
{{
    "course_overview": {{
        "course_title": "{role_title} Training Course",
        "duration": "X days/weeks",
        "target_audience": "Description of target audience",
        "prerequisites": ["Prerequisite 1", "Prerequisite 2"],
        "course_aim": "Overall course aim"
    }},
    "design_matrix": [
        {{
            "objective": "{terms['top_objective_short']} 1",
            "method": "Blended",
            "media": "Presentation, Practical",
            "assessment": "Practical assessment",
            "resources": "Equipment needed",
            "duration": "2 hours"
        }}
    ],
    "assessment_strategy": {{
        "formative": "Description of formative assessment approach",
        "summative": "Description of summative assessment approach",
        "pass_criteria": "Pass/fail criteria",
        "remediation": "Remediation approach"
    }},
    "methods_media_justification": "Evidence-based rationale for methods and media selection",
    "resource_requirements": {{
        "trainers": "Number and qualifications",
        "equipment": "Equipment list",
        "facilities": "Facility requirements",
        "materials": "Training materials"
    }},
    "schedule_outline": [
        {{"day": 1, "session": "Morning", "topic": "Topic name", "duration": "3 hours"}}
    ]
}}

Generate a comprehensive course design with 6-8 design matrix entries.
Return ONLY the JSON."""

    response = await call_claude(prompt, max_tokens=5000)
    return parse_json(response)


# ============================================================================
# DELIVERY AGENT
# ============================================================================

async def run_delivery_agent(job_id: str, parameters: Dict, framework: str, templates: Dict):
    """Generate Delivery phase documents based on framework"""
    role_title = parameters.get("role_title", "Training Specialist")
    role_desc = parameters.get("role_description", "")
    terms = get_terminology(framework)
    
    output_dir = Path(jobs.get(job_id)["output_dir"]) / "03_Delivery"
    output_dir.mkdir(exist_ok=True)
    
    update_job(job_id, 5, f"Starting Delivery Agent ({framework})...")
    
    # Generate Lesson Plans
    update_job(job_id, 10, f"Generating {terms['lesson_plan']}...")
    lesson_plans = await generate_lesson_plans(role_title, framework, role_desc, terms)
    update_job(job_id, 50, f"Building Lesson Plans...")
    build_lesson_plans_doc(lesson_plans, role_title, framework, terms, output_dir / "01_Lesson_Plans.docx")
    update_job(job_id, 60, f"✓ Lesson Plans complete")
    
    # Generate Assessment Instruments
    update_job(job_id, 65, f"Generating Assessment Instruments...")
    assessments = await generate_assessments(role_title, framework, terms)
    update_job(job_id, 90, f"Building Assessment Instruments...")
    build_assessments_doc(assessments, role_title, framework, terms, output_dir / "02_Assessment_Instruments.docx")
    update_job(job_id, 95, f"✓ Assessment Instruments complete")
    
    update_job(job_id, 100, "Delivery Phase Complete")


async def generate_lesson_plans(role_title: str, framework: str, description: str, terms: Dict) -> Dict:
    """Generate framework-specific lesson plans"""
    
    if framework == "US_TRADOC":
        prompt = f"""Generate Lesson Plans in 5-Section Format for:

Role: {role_title}
Framework: US TRADOC (TP 350-70-1)

Return a JSON object:
{{
    "lessons": [
        {{
            "lesson_number": "Lesson 1",
            "lesson_title": "Lesson title",
            "section_1_admin": {{
                "course_number": "Course number",
                "hours": "2.0",
                "method_of_instruction": "Lecture/Practical",
                "references": ["Reference 1", "Reference 2"]
            }},
            "section_2_intro": {{
                "motivator": "Why this is important",
                "tlo_statement": "Terminal Learning Objective",
                "safety": "Safety considerations",
                "risk_assessment": "Low/Medium/High",
                "evaluation_method": "Performance evaluation"
            }},
            "section_3_presentation": {{
                "elos": [
                    {{
                        "elo": "ELO statement",
                        "lsas": ["Learning step 1", "Learning step 2"],
                        "check_on_learning": "Question to verify understanding"
                    }}
                ]
            }},
            "section_4_summary": {{
                "review": "Summary of key points",
                "check_on_learning": "Final questions",
                "transition": "Connection to next lesson"
            }},
            "section_5_assessment": {{
                "testing_procedures": "How assessment conducted",
                "go_nogo_criteria": "Pass/fail criteria",
                "remediation": "What happens if NO-GO"
            }}
        }}
    ]
}}

Generate 3-4 lesson plans.
Return ONLY the JSON."""

    elif framework == "ACTION_MAPPING":
        prompt = f"""Generate Scenario-Based Activities for:

Role: {role_title}
Framework: Action Mapping (Cathy Moore)

Return a JSON object:
{{
    "activities": [
        {{
            "activity_number": "Activity 1",
            "business_goal_link": "Which business goal this supports",
            "scenario_setup": {{
                "context": "Realistic work situation",
                "character": "Role the learner plays",
                "challenge": "Problem they face"
            }},
            "decision_point": {{
                "question": "What should you do?",
                "options": [
                    {{"option": "Option A", "is_correct": false, "consequence": "What happens"}},
                    {{"option": "Option B", "is_correct": true, "consequence": "What happens"}},
                    {{"option": "Option C", "is_correct": false, "consequence": "What happens"}}
                ]
            }},
            "feedback": {{
                "correct": "Consequence-based feedback for correct choice",
                "incorrect": "Consequence-based feedback showing impact"
            }},
            "minimum_information": "Only what's needed to make the decision"
        }}
    ]
}}

Generate 4-5 scenario-based activities.
Focus on realistic decisions, not information recall.
Return ONLY the JSON."""

    else:  # UK_DSAT PAR format
        prompt = f"""Generate Lesson Plans in PAR (Present-Apply-Review) format for:

Role: {role_title}
Framework: UK DSAT (DTSM 4)

Return a JSON object:
{{
    "lessons": [
        {{
            "lesson_number": "Lesson 1",
            "lesson_title": "Lesson title",
            "duration": "2 hours",
            "tos_addressed": ["TO 1"],
            "eos_addressed": ["EO 1.1", "EO 1.2"],
            "prerequisites": "What learners need before this lesson",
            "resources": ["Resource 1", "Resource 2"],
            "present": {{
                "time_allocation": "40%",
                "klps": ["KLP 1", "KLP 2", "KLP 3"],
                "trainer_script": "Key teaching points for instructor",
                "visual_aids": ["Slide deck", "Demonstration video"]
            }},
            "apply": {{
                "time_allocation": "40%",
                "trainee_activities": ["Activity description"],
                "practical_exercises": ["Exercise 1 description"],
                "formative_assessment": "How to check understanding during practice"
            }},
            "review": {{
                "time_allocation": "20%",
                "summary": "Key points to reinforce",
                "questions": ["Review question 1", "Review question 2"],
                "consolidation": "How to consolidate learning"
            }}
        }}
    ]
}}

Generate 3-4 lesson plans with full PAR structure.
Time split: 40% Present, 40% Apply, 20% Review.
Return ONLY the JSON."""

    response = await call_claude(prompt, max_tokens=6000)
    return parse_json(response)


async def generate_assessments(role_title: str, framework: str, terms: Dict) -> Dict:
    """Generate framework-specific assessment instruments"""
    
    if framework == "KIRKPATRICK":
        prompt = f"""Generate Kirkpatrick Four-Level Assessment Instruments for:

Role: {role_title}
Framework: Kirkpatrick Model

Return a JSON object:
{{
    "level_1_reaction": {{
        "survey_title": "End-of-Course Reaction Survey",
        "items": [
            {{"item": "Survey question about engagement", "scale": "1-5"}},
            {{"item": "Survey question about relevance", "scale": "1-5"}}
        ],
        "open_questions": ["What was most valuable?", "What could be improved?"]
    }},
    "level_2_learning": {{
        "pre_test": [
            {{"question": "Pre-test question", "type": "multiple_choice", "answer": "Correct answer"}}
        ],
        "post_test": [
            {{"question": "Post-test question", "type": "multiple_choice", "answer": "Correct answer"}}
        ],
        "skills_checklist": [
            {{"skill": "Skill to demonstrate", "criteria": "What good looks like"}}
        ]
    }},
    "level_3_behavior": {{
        "observation_checklist": [
            {{"behavior": "Observable behavior", "frequency": "How often to observe"}}
        ],
        "manager_survey": [
            {{"question": "Question for manager about behavior change"}}
        ],
        "self_assessment": [
            {{"question": "Self-assessment question"}}
        ]
    }},
    "level_4_results": {{
        "metrics": [
            {{"metric": "Business metric", "baseline": "Before", "target": "After", "data_source": "Where from"}}
        ],
        "roi_calculation": "Method for calculating ROI"
    }}
}}

Generate comprehensive assessments for all four levels.
Return ONLY the JSON."""

    else:
        prompt = f"""Generate Assessment Instruments for:

Role: {role_title}
Framework: {framework}

Return a JSON object:
{{
    "assessment_overview": {{
        "strategy": "Overall assessment strategy",
        "formative_approach": "How formative assessment works",
        "summative_approach": "How summative assessment works",
        "pass_criteria": "Overall pass/fail criteria"
    }},
    "knowledge_tests": [
        {{
            "test_id": "KT-001",
            "objective_covered": "{terms['top_objective_short']} 1",
            "questions": [
                {{
                    "question": "Test question",
                    "type": "multiple_choice",
                    "options": ["A", "B", "C", "D"],
                    "correct_answer": "B",
                    "rationale": "Why this is correct"
                }}
            ]
        }}
    ],
    "practical_assessments": [
        {{
            "assessment_id": "PA-001",
            "objective_covered": "{terms['top_objective_short']} 1",
            "task": "Practical task to perform",
            "conditions": "Assessment conditions",
            "checklist": [
                {{"step": "Performance step", "criteria": "What good looks like"}}
            ],
            "pass_standard": "What constitutes a pass"
        }}
    ],
    "marking_scheme": {{
        "grading_scale": "Pass/Fail or percentage",
        "weighting": "How different components weighted",
        "moderation": "Quality assurance process"
    }}
}}

Generate 3-4 knowledge tests and 2-3 practical assessments.
Return ONLY the JSON."""

    response = await call_claude(prompt, max_tokens=5000)
    return parse_json(response)


# ============================================================================
# EVALUATION AGENT
# ============================================================================

async def run_evaluation_agent(job_id: str, parameters: Dict, framework: str, templates: Dict):
    """Generate Evaluation phase documents based on framework"""
    role_title = parameters.get("role_title", "Training Specialist")
    role_desc = parameters.get("role_description", "")
    terms = get_terminology(framework)
    
    output_dir = Path(jobs.get(job_id)["output_dir"]) / "04_Evaluation"
    output_dir.mkdir(exist_ok=True)
    
    update_job(job_id, 5, f"Starting Evaluation Agent ({framework})...")
    
    # Generate Internal Evaluation
    update_job(job_id, 10, f"Generating {terms['internal_eval']}...")
    internal_eval = await generate_internal_eval(role_title, framework, terms)
    update_job(job_id, 45, f"Building {terms['internal_eval']} document...")
    filename = sanitize_filename(terms['internal_eval'].split('(')[0].strip())
    build_evaluation_doc(internal_eval, role_title, framework, terms, output_dir / f"01_{filename}.docx", "internal")
    update_job(job_id, 55, f"✓ {terms['internal_eval']} complete")
    
    # Generate External Evaluation
    update_job(job_id, 60, f"Generating {terms['external_eval']}...")
    external_eval = await generate_external_eval(role_title, framework, terms)
    update_job(job_id, 90, f"Building {terms['external_eval']} document...")
    filename = sanitize_filename(terms['external_eval'].split('(')[0].strip())
    build_evaluation_doc(external_eval, role_title, framework, terms, output_dir / f"02_{filename}.docx", "external")
    update_job(job_id, 95, f"✓ {terms['external_eval']} complete")
    
    update_job(job_id, 100, "Evaluation Phase Complete")


async def generate_internal_eval(role_title: str, framework: str, terms: Dict) -> Dict:
    """Generate framework-specific internal evaluation"""
    
    if framework == "KIRKPATRICK":
        prompt = f"""Generate a Levels 1-2 Assessment Report for:

Role: {role_title}
Framework: Kirkpatrick Model

Return a JSON object:
{{
    "executive_summary": "Summary of Level 1 and 2 findings",
    "level_1_results": {{
        "response_rate": "XX%",
        "engagement_score": "X.X/5.0",
        "relevance_score": "X.X/5.0",
        "satisfaction_score": "X.X/5.0",
        "key_themes": ["Theme 1", "Theme 2"],
        "improvement_suggestions": ["Suggestion 1", "Suggestion 2"]
    }},
    "level_2_results": {{
        "pre_test_average": "XX%",
        "post_test_average": "XX%",
        "knowledge_gain": "XX percentage points",
        "skills_pass_rate": "XX%",
        "confidence_improvement": "Description of confidence changes"
    }},
    "recommendations": [
        {{"id": "R1", "recommendation": "Recommendation", "level": "L1/L2", "priority": "High"}}
    ]
}}

Return ONLY the JSON."""

    else:
        prompt = f"""Generate a {terms['internal_eval']} for:

Role: {role_title}
Framework: {framework}

Return a JSON object:
{{
    "executive_summary": "Summary of internal evaluation findings",
    "pilot_analysis": {{
        "course_dates": "When pilot ran",
        "participants": "Number and profile",
        "completion_rate": "XX%",
        "overall_assessment": "Summary assessment"
    }},
    "trainee_feedback": {{
        "satisfaction_score": "X.X/5.0",
        "relevance_score": "X.X/5.0",
        "key_positives": ["Positive 1", "Positive 2"],
        "key_concerns": ["Concern 1", "Concern 2"]
    }},
    "trainer_observations": [
        {{"observation": "Observation", "impact": "Impact on learning"}}
    ],
    "pass_fail_analysis": {{
        "overall_pass_rate": "XX%",
        "by_objective": [
            {{"{terms['top_objective_short']}": "1", "pass_rate": "XX%"}}
        ],
        "common_failure_points": ["Failure point 1", "Failure point 2"]
    }},
    "recommendations": [
        {{"id": "R1", "recommendation": "Recommendation", "priority": "High", "owner": "Who"}}
    ],
    "ceb_actions": [
        {{"action": "Action for CEB", "decision_required": "What decision needed"}}
    ]
}}

Return ONLY the JSON."""

    response = await call_claude(prompt, max_tokens=4000)
    return parse_json(response)


async def generate_external_eval(role_title: str, framework: str, terms: Dict) -> Dict:
    """Generate framework-specific external evaluation"""
    
    if framework == "KIRKPATRICK":
        prompt = f"""Generate a Chain of Evidence Report for:

Role: {role_title}
Framework: Kirkpatrick Model

Return a JSON object:
{{
    "executive_summary": "Executive narrative linking all four levels",
    "level_1_evidence": {{
        "data_summary": "Summary of reaction data",
        "key_finding": "Main finding"
    }},
    "level_2_evidence": {{
        "data_summary": "Summary of learning data",
        "key_finding": "Main finding"
    }},
    "level_3_evidence": {{
        "data_summary": "Summary of behavior data",
        "observation_period": "How long observed",
        "behavior_changes": ["Change 1", "Change 2"],
        "barriers_identified": ["Barrier 1", "Barrier 2"],
        "key_finding": "Main finding"
    }},
    "level_4_evidence": {{
        "metrics_achieved": [
            {{"metric": "Metric name", "baseline": "Before", "current": "After", "target": "Goal"}}
        ],
        "business_impact": "Description of business impact",
        "key_finding": "Main finding"
    }},
    "roi_calculation": {{
        "benefits": "Total benefits in currency",
        "costs": "Total costs in currency",
        "roi_percentage": "XX%",
        "calculation_method": "How calculated"
    }},
    "roe_assessment": "Return on Expectations - did we meet sponsor expectations?",
    "recommendations": [
        {{"recommendation": "Recommendation", "rationale": "Why"}}
    ]
}}

Return ONLY the JSON."""

    else:
        prompt = f"""Generate a {terms['external_eval']} for:

Role: {role_title}
Framework: {framework}

Return a JSON object:
{{
    "executive_summary": "Summary of external evaluation findings",
    "sampling_strategy": {{
        "population": "Total trained population",
        "sample_size": "Number in sample",
        "selection_method": "How selected",
        "observation_period": "How long after training"
    }},
    "workplace_performance": {{
        "methodology": "How workplace performance measured",
        "findings": [
            {{"area": "Performance area", "rating": "Rating", "evidence": "Evidence"}}
        ],
        "overall_assessment": "Summary of workplace performance"
    }},
    "transfer_assessment": {{
        "transfer_rate": "XX% applying learning on job",
        "barriers": ["Barrier 1", "Barrier 2"],
        "enablers": ["Enabler 1", "Enabler 2"]
    }},
    "line_manager_feedback": {{
        "response_rate": "XX%",
        "satisfaction_with_graduates": "X.X/5.0",
        "key_comments": ["Comment 1", "Comment 2"]
    }},
    "system_improvements": [
        {{"improvement": "Suggested improvement", "phase": "Analysis/Design/Delivery", "priority": "High"}}
    ],
    "recommendations": [
        {{"id": "R1", "recommendation": "Recommendation for TRA", "owner": "{terms['authority']}", "timeline": "Timeline"}}
    ]
}}

Return ONLY the JSON."""

    response = await call_claude(prompt, max_tokens=4000)
    return parse_json(response)


# ============================================================================
# DOCUMENT STYLING HELPERS
# ============================================================================

NOVA_DARK_RED = RGBColor(139, 69, 69)
NOVA_LIGHT_GRAY = RGBColor(245, 245, 245)
NOVA_DARK_BLUE = RGBColor(31, 56, 100)
NOVA_GREEN = RGBColor(34, 139, 34)

def split_into_sentences(text: str) -> List[str]:
    """Split text into sentences"""
    if not text:
        return []
    text = text.replace('e.g.', 'e_g_').replace('i.e.', 'i_e_').replace('etc.', 'etc_')
    sentences = re.split(r'(?<=[.!?])\s+', text)
    result = []
    for s in sentences:
        s = s.replace('e_g_', 'e.g.').replace('i_e_', 'i.e.').replace('etc_', 'etc.')
        if s.strip():
            result.append(s.strip())
    return result

def set_cell_shading(cell, color_hex: str):
    """Set cell background color"""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    cell._tc.get_or_add_tcPr().append(shading)

def set_repeat_table_header(row):
    """Set table row to repeat as header"""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    trPr.append(tblHeader)

def set_paragraph_spacing(paragraph, before_pt=6, after_pt=6, line_spacing=1.5):
    """Set paragraph spacing"""
    pPr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(int(before_pt * 20)))
    spacing.set(qn('w:after'), str(int(after_pt * 20)))
    spacing.set(qn('w:line'), str(int(line_spacing * 240)))
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)

def create_styled_paragraph(doc, text: str, font_name: str = "Roboto", font_size: int = 11, 
                           bold: bool = False, color: RGBColor = None):
    """Create styled paragraph"""
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
    """Create styled heading"""
    heading = doc.add_heading(text, level)
    for run in heading.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    return heading

def create_styled_table(doc, headers: List[str], rows: List[List[str]], col_widths: List[float] = None, header_color: str = "8B4545"):
    """Create professionally styled table"""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    header_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = header
        set_cell_shading(cell, header_color)
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.name = "Roboto"
                run.font.size = Pt(10)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
    
    for row_idx, row_data in enumerate(rows):
        row = table.add_row()
        for i, cell_text in enumerate(row_data):
            cell = row.cells[i]
            cell.text = str(cell_text) if cell_text else ""
            if row_idx % 2 == 1:
                set_cell_shading(cell, "F5F5F5")
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(10)
    
    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(width)
    
    return table


# ============================================================================
# DOCUMENT BUILDERS
# ============================================================================

def build_task_list_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build task list document (RolePS/ICTL/STP etc.)"""
    doc = Document()
    
    # Set to landscape for task lists
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width = section.page_height
    new_height = section.page_width
    section.page_width = new_width
    section.page_height = new_height
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
    
    # Title
    title = doc.add_heading(terms['task_list'], 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    
    # Subtitle
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(14)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Framework badge
    badge = doc.add_paragraph()
    run = badge.add_run(f"Framework: {framework}")
    run.font.name = "Roboto"
    run.font.size = Pt(10)
    run.font.italic = True
    badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Header information table
    header_table = doc.add_table(rows=4, cols=4)
    header_table.style = 'Table Grid'
    
    # Populate based on framework
    if framework == "US_TRADOC":
        header_table.rows[0].cells[0].text = "MOS:"
        header_table.rows[0].cells[1].text = header.get("mos", "")
        header_table.rows[0].cells[2].text = "SKILL LEVEL:"
        header_table.rows[0].cells[3].text = header.get("skill_level", "")
        header_table.rows[1].cells[0].text = "PROPONENT:"
        header_table.rows[1].cells[1].text = header.get("proponent", "")
        header_table.rows[1].cells[2].text = "EFFECTIVE DATE:"
        header_table.rows[1].cells[3].text = header.get("effective_date", "")
    elif framework == "NATO_BISC":
        header_table.rows[0].cells[0].text = "DISCIPLINE:"
        header_table.rows[0].cells[1].text = header.get("discipline", "")
        header_table.rows[0].cells[2].text = "RA:"
        header_table.rows[0].cells[3].text = header.get("ra", "")
        header_table.rows[1].cells[0].text = "DH:"
        header_table.rows[1].cells[1].text = header.get("dh", "")
        header_table.rows[1].cells[2].text = "ETF:"
        header_table.rows[1].cells[3].text = header.get("etf", "")
    else:  # UK DSAT default
        header_table.rows[0].cells[0].text = "ROLE TITLE(S):"
        header_table.rows[0].cells[1].text = header.get("role_title", role_title)
        header_table.rows[0].cells[2].text = "ROLE NUMBER(S):"
        header_table.rows[0].cells[3].text = header.get("role_number", "")
        header_table.rows[1].cells[0].text = "DUTY TITLE(S):"
        header_table.rows[1].cells[1].text = header.get("duty_title", "")
        header_table.rows[1].cells[2].text = "DUTY NUMBER(S):"
        header_table.rows[1].cells[3].text = header.get("duty_number", "")
        header_table.rows[2].cells[0].text = "TRA:"
        header_table.rows[2].cells[1].text = header.get("tra", "")
        header_table.rows[2].cells[2].text = "ROLEPS REFERENCE:"
        header_table.rows[2].cells[3].text = header.get("roleps_reference", "")
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
                    if i % 2 == 0:
                        run.font.bold = True
    
    doc.add_paragraph()
    
    # Task table - headers vary by framework
    if framework == "US_TRADOC":
        task_headers = ["Task Number", "Task Title", "Training Domain", "Sustainment Freq", "Skill Level", "Subject Area"]
    elif framework == "NATO_BISC":
        task_headers = ["Task ID", "Task Description", "Proficiency Level", "Collective Task", "Gap Indicator"]
    elif framework == "ADDIE":
        task_headers = ["Task ID", "Task Description", "Knowledge", "Skills", "Criticality", "Frequency"]
    else:
        task_headers = ["Task/Sub Task No.", "Performance", "Conditions", "Standards", "Training Cat.", "Notes"]
    
    task_table = doc.add_table(rows=1, cols=len(task_headers))
    task_table.style = 'Table Grid'
    
    header_row = task_table.rows[0]
    set_repeat_table_header(header_row)
    for i, h in enumerate(task_headers):
        cell = header_row.cells[i]
        cell.text = h
        set_cell_shading(cell, "8B4545")
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
        
        if framework == "US_TRADOC":
            row.cells[0].text = str(task.get("task_number", ""))
            row.cells[1].text = task.get("task_title", "")
            row.cells[2].text = task.get("training_domain", "")
            row.cells[3].text = task.get("sustainment_frequency", "")
            row.cells[4].text = str(task.get("sustainment_skill_level", ""))
            row.cells[5].text = task.get("subject_area", "")
        elif framework == "NATO_BISC":
            row.cells[0].text = str(task.get("task_id", ""))
            row.cells[1].text = task.get("task_description", "")
            row.cells[2].text = task.get("proficiency_level", "")
            row.cells[3].text = task.get("collective_task", "")
            row.cells[4].text = task.get("gap_indicator", "")
        elif framework == "ADDIE":
            row.cells[0].text = str(task.get("task_id", ""))
            row.cells[1].text = task.get("task_description", "")
            knowledge = task.get("knowledge", [])
            row.cells[2].text = "\n".join(knowledge) if isinstance(knowledge, list) else str(knowledge)
            skills = task.get("skills", [])
            row.cells[3].text = "\n".join(skills) if isinstance(skills, list) else str(skills)
            row.cells[4].text = task.get("criticality", "")
            row.cells[5].text = task.get("frequency", "")
        else:
            row.cells[0].text = str(task.get("task_number", ""))
            row.cells[1].text = task.get("performance", "")
            conditions = task.get("conditions", [])
            row.cells[2].text = "\n".join(conditions) if isinstance(conditions, list) else str(conditions)
            standards = task.get("standards", [])
            row.cells[3].text = "\n".join(standards) if isinstance(standards, list) else str(standards)
            row.cells[4].text = str(task.get("training_category", ""))
            notes = task.get("notes", [])
            row.cells[5].text = "\n".join(notes) if isinstance(notes, list) else str(notes)
        
        if task_idx % 2 == 1:
            for cell in row.cells:
                set_cell_shading(cell, "F5F5F5")
        
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(9)
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_needs_report_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build training needs report document"""
    doc = Document()
    
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    
    # Title
    title = doc.add_heading(terms['needs_report'].upper(), 0)
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
    
    # Framework badge
    badge = doc.add_paragraph()
    run = badge.add_run(f"Framework: {framework}")
    run.font.name = "Roboto"
    run.font.size = Pt(10)
    run.font.italic = True
    badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("_" * 80)
    
    # Executive Summary
    create_styled_heading(doc, "1. EXECUTIVE SUMMARY", 1)
    exec_summary = data.get("executive_summary", "")
    sentences = split_into_sentences(exec_summary)
    for sentence in sentences:
        if sentence.strip():
            create_styled_paragraph(doc, sentence.strip())
    
    # Introduction
    create_styled_heading(doc, "2. INTRODUCTION", 1)
    intro = data.get("introduction", {})
    if intro:
        create_styled_paragraph(doc, f"Purpose: {intro.get('purpose', '')}")
        create_styled_paragraph(doc, f"Scope: {intro.get('scope', '')}")
        create_styled_paragraph(doc, f"Methodology: {intro.get('methodology', '')}")
    
    # Key Findings
    create_styled_heading(doc, "3. KEY FINDINGS", 1)
    findings = data.get("key_findings", [])
    for i, finding in enumerate(findings, 1):
        create_styled_paragraph(doc, f"{i}. {finding}")
    
    # Training Requirements
    create_styled_heading(doc, "4. TRAINING REQUIREMENTS", 1)
    requirements = data.get("training_requirements", [])
    if requirements:
        rows = [[r.get("id", ""), r.get("requirement", ""), r.get("priority", ""), r.get("delivery_method", "")] for r in requirements]
        create_styled_table(doc, ["ID", "Requirement", "Priority", "Delivery Method"], rows, [0.6, 4.0, 0.9, 1.1])
    
    doc.add_paragraph()
    
    # Recommendations
    create_styled_heading(doc, "5. RECOMMENDATIONS", 1)
    recommendations = data.get("recommendations", [])
    if recommendations:
        rows = [[r.get("id", ""), r.get("recommendation", ""), r.get("priority", ""), r.get("timeline", "")] for r in recommendations]
        create_styled_table(doc, ["ID", "Recommendation", "Priority", "Timeline"], rows, [0.6, 4.0, 0.8, 1.2])
    
    doc.add_paragraph()
    
    # Resource Requirements
    create_styled_heading(doc, "6. RESOURCE REQUIREMENTS", 1)
    resources = data.get("resource_requirements", {})
    create_styled_paragraph(doc, f"Estimated Budget: £{resources.get('budget_estimate', 0):,}")
    create_styled_paragraph(doc, f"Timeline: {resources.get('timeline_months', 0)} months")
    create_styled_paragraph(doc, f"Personnel: {resources.get('personnel_required', '')}")
    
    # Conclusion
    create_styled_heading(doc, "7. CONCLUSION", 1)
    conclusion = data.get("conclusion", "")
    sentences = split_into_sentences(conclusion)
    for sentence in sentences:
        if sentence.strip():
            create_styled_paragraph(doc, sentence.strip())
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_objectives_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build learning objectives document"""
    doc = Document()
    
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Title
    title = doc.add_heading(f"{terms['top_objective']} HIERARCHY", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    
    # Subtitle
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(14)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    badge = doc.add_paragraph()
    run = badge.add_run(f"Framework: {framework} | Format: {terms['objective_format']}")
    run.font.name = "Roboto"
    run.font.size = Pt(10)
    run.font.italic = True
    badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("_" * 80)
    
    objectives = data.get("objectives", data.get("actions", []))
    
    for obj in objectives:
        # Top level objective
        obj_num = obj.get("to_number") or obj.get("tlo_number") or obj.get("po_number") or obj.get("lo_number") or obj.get("action_number", "")
        
        create_styled_heading(doc, obj_num, 2)
        
        # Format varies by framework
        if framework == "US_TRADOC":
            create_styled_paragraph(doc, f"Action: {obj.get('action', '')}", bold=True)
            create_styled_paragraph(doc, f"Condition: {obj.get('condition', '')}")
            create_styled_paragraph(doc, f"Standard: {obj.get('standard', '')}")
        elif framework == "NATO_BISC":
            create_styled_paragraph(doc, f"Performance: {obj.get('performance', '')}", bold=True)
            standard = obj.get('standard', {})
            if isinstance(standard, dict):
                create_styled_paragraph(doc, f"Measure: {standard.get('measure', '')}")
                create_styled_paragraph(doc, f"Criterion: {standard.get('criterion', '')}")
            else:
                create_styled_paragraph(doc, f"Standard: {standard}")
        elif framework == "ADDIE":
            create_styled_paragraph(doc, f"Audience: {obj.get('audience', '')}")
            create_styled_paragraph(doc, f"Behavior: {obj.get('behavior', '')}", bold=True)
            create_styled_paragraph(doc, f"Condition: {obj.get('condition', '')}")
            create_styled_paragraph(doc, f"Degree: {obj.get('degree', '')}")
        elif framework == "ACTION_MAPPING":
            create_styled_paragraph(doc, f"Action: {obj.get('action_statement', '')}", bold=True)
            practice = obj.get('practice_activity', {})
            if practice:
                create_styled_paragraph(doc, f"Scenario: {practice.get('scenario', '')}")
                create_styled_paragraph(doc, f"Decision Point: {practice.get('decision_point', '')}")
        else:  # UK DSAT
            create_styled_paragraph(doc, f"Performance: {obj.get('performance', '')}", bold=True)
            create_styled_paragraph(doc, f"Conditions: {obj.get('conditions', '')}")
            create_styled_paragraph(doc, f"Standards: {obj.get('standards', '')}")
        
        # Enabling objectives
        eos = obj.get("eos") or obj.get("elos") or obj.get("supporting_objectives", [])
        if eos:
            create_styled_heading(doc, f"{terms['enabling_objective_short']}s", 3)
            for eo in eos:
                eo_num = eo.get("eo_number") or eo.get("elo_number") or eo.get("so_number", "")
                p = doc.add_paragraph()
                run = p.add_run(f"{eo_num}: ")
                run.font.bold = True
                run.font.name = "Roboto"
                
                if framework == "UK_DSAT":
                    run2 = p.add_run(f"{eo.get('performance', eo.get('statement', ''))} [{eo.get('ksa_tag', 'S')}]")
                else:
                    run2 = p.add_run(eo.get('action', eo.get('statement', '')))
                run2.font.name = "Roboto"
        
        # KLPs/LSAs
        klps = obj.get("klps") or obj.get("lsas", [])
        if klps:
            create_styled_heading(doc, f"{terms['learning_point_short']}s", 3)
            for klp in klps:
                klp_num = klp.get("klp_number") or klp.get("lsa_number", "")
                p = doc.add_paragraph()
                run = p.add_run(f"{klp_num}: ")
                run.font.bold = True
                run.font.name = "Roboto"
                run2 = p.add_run(klp.get('statement', klp.get('activity', '')))
                run2.font.name = "Roboto"
                if klp.get('domain'):
                    run3 = p.add_run(f" [{klp.get('domain')}]")
                    run3.font.name = "Roboto"
                    run3.font.italic = True
        
        doc.add_paragraph()
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_course_design_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build course design document"""
    doc = Document()
    
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Title
    title = doc.add_heading(terms['course_design'].upper(), 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(14)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("_" * 80)
    
    # Course Overview
    create_styled_heading(doc, "1. COURSE OVERVIEW", 1)
    overview = data.get("course_overview", {})
    create_styled_paragraph(doc, f"Course Title: {overview.get('course_title', '')}", bold=True)
    create_styled_paragraph(doc, f"Duration: {overview.get('duration', '')}")
    create_styled_paragraph(doc, f"Target Audience: {overview.get('target_audience', '')}")
    create_styled_paragraph(doc, f"Course Aim: {overview.get('course_aim', '')}")
    
    prereqs = overview.get('prerequisites', [])
    if prereqs:
        create_styled_paragraph(doc, "Prerequisites:", bold=True)
        for prereq in prereqs:
            create_styled_paragraph(doc, f"  • {prereq}")
    
    # Design Matrix
    create_styled_heading(doc, "2. DESIGN MATRIX", 1)
    matrix = data.get("design_matrix", [])
    if matrix:
        rows = [[m.get("objective", ""), m.get("method", ""), m.get("media", ""), 
                 m.get("assessment", ""), m.get("duration", "")] for m in matrix]
        create_styled_table(doc, [terms['top_objective_short'], "Method", "Media", "Assessment", "Duration"], 
                          rows, [1.2, 1.2, 1.5, 1.5, 0.8])
    
    doc.add_paragraph()
    
    # Assessment Strategy
    create_styled_heading(doc, "3. ASSESSMENT STRATEGY", 1)
    assessment = data.get("assessment_strategy", {})
    create_styled_paragraph(doc, f"Formative: {assessment.get('formative', '')}")
    create_styled_paragraph(doc, f"Summative: {assessment.get('summative', '')}")
    create_styled_paragraph(doc, f"Pass Criteria: {assessment.get('pass_criteria', '')}")
    create_styled_paragraph(doc, f"Remediation: {assessment.get('remediation', '')}")
    
    # Methods & Media Justification
    create_styled_heading(doc, "4. METHODS & MEDIA JUSTIFICATION", 1)
    create_styled_paragraph(doc, data.get("methods_media_justification", ""))
    
    # Resource Requirements
    create_styled_heading(doc, "5. RESOURCE REQUIREMENTS", 1)
    resources = data.get("resource_requirements", {})
    create_styled_paragraph(doc, f"Trainers: {resources.get('trainers', '')}")
    create_styled_paragraph(doc, f"Equipment: {resources.get('equipment', '')}")
    create_styled_paragraph(doc, f"Facilities: {resources.get('facilities', '')}")
    create_styled_paragraph(doc, f"Materials: {resources.get('materials', '')}")
    
    # Schedule Outline
    create_styled_heading(doc, "6. SCHEDULE OUTLINE", 1)
    schedule = data.get("schedule_outline", [])
    if schedule:
        rows = [[str(s.get("day", "")), s.get("session", ""), s.get("topic", ""), s.get("duration", "")] for s in schedule]
        create_styled_table(doc, ["Day", "Session", "Topic", "Duration"], rows, [0.6, 1.0, 3.5, 1.0])
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_lesson_plans_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build lesson plans document"""
    doc = Document()
    
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Title
    title = doc.add_heading(f"{terms['lesson_plan'].upper()}S", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(14)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("_" * 80)
    
    lessons = data.get("lessons", data.get("activities", []))
    
    for lesson in lessons:
        lesson_num = lesson.get("lesson_number") or lesson.get("activity_number", "")
        lesson_title = lesson.get("lesson_title", "")
        
        create_styled_heading(doc, f"{lesson_num}: {lesson_title}", 2)
        
        if framework == "US_TRADOC":
            # 5-Section Format
            create_styled_heading(doc, "Section 1: Administrative Data", 3)
            admin = lesson.get("section_1_admin", {})
            create_styled_paragraph(doc, f"Hours: {admin.get('hours', '')}")
            create_styled_paragraph(doc, f"Method: {admin.get('method_of_instruction', '')}")
            
            create_styled_heading(doc, "Section 2: Introduction", 3)
            intro = lesson.get("section_2_intro", {})
            create_styled_paragraph(doc, f"Motivator: {intro.get('motivator', '')}")
            create_styled_paragraph(doc, f"TLO: {intro.get('tlo_statement', '')}")
            
            create_styled_heading(doc, "Section 3: Presentation", 3)
            pres = lesson.get("section_3_presentation", {})
            for elo in pres.get("elos", []):
                create_styled_paragraph(doc, f"ELO: {elo.get('elo', '')}", bold=True)
                for lsa in elo.get("lsas", []):
                    create_styled_paragraph(doc, f"  • {lsa}")
            
            create_styled_heading(doc, "Section 4: Summary/Review", 3)
            summary = lesson.get("section_4_summary", {})
            create_styled_paragraph(doc, summary.get("review", ""))
            
            create_styled_heading(doc, "Section 5: Assessment", 3)
            assess = lesson.get("section_5_assessment", {})
            create_styled_paragraph(doc, f"GO/NO-GO Criteria: {assess.get('go_nogo_criteria', '')}")
            
        elif framework == "ACTION_MAPPING":
            # Scenario-based
            scenario = lesson.get("scenario_setup", {})
            create_styled_paragraph(doc, f"Context: {scenario.get('context', '')}", bold=True)
            create_styled_paragraph(doc, f"Challenge: {scenario.get('challenge', '')}")
            
            decision = lesson.get("decision_point", {})
            create_styled_heading(doc, "Decision Point", 3)
            create_styled_paragraph(doc, decision.get("question", ""))
            
            for opt in decision.get("options", []):
                marker = "✓" if opt.get("is_correct") else "✗"
                create_styled_paragraph(doc, f"  {marker} {opt.get('option', '')}: {opt.get('consequence', '')}")
            
        else:
            # UK DSAT PAR Format
            create_styled_paragraph(doc, f"Duration: {lesson.get('duration', '')}")
            create_styled_paragraph(doc, f"TOs Addressed: {', '.join(lesson.get('tos_addressed', []))}")
            
            create_styled_heading(doc, "PRESENT (40%)", 3)
            present = lesson.get("present", {})
            klps = present.get("klps", [])
            for klp in klps:
                create_styled_paragraph(doc, f"  • {klp}")
            create_styled_paragraph(doc, f"Trainer Script: {present.get('trainer_script', '')}")
            
            create_styled_heading(doc, "APPLY (40%)", 3)
            apply_section = lesson.get("apply", {})
            for activity in apply_section.get("trainee_activities", []):
                create_styled_paragraph(doc, f"  • {activity}")
            
            create_styled_heading(doc, "REVIEW (20%)", 3)
            review = lesson.get("review", {})
            create_styled_paragraph(doc, f"Summary: {review.get('summary', '')}")
            for q in review.get("questions", []):
                create_styled_paragraph(doc, f"  Q: {q}")
        
        doc.add_paragraph()
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_assessments_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build assessment instruments document"""
    doc = Document()
    
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Title
    title = doc.add_heading("ASSESSMENT INSTRUMENTS", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(14)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("_" * 80)
    
    if framework == "KIRKPATRICK":
        # Kirkpatrick four levels
        create_styled_heading(doc, "LEVEL 1: REACTION", 1)
        l1 = data.get("level_1_reaction", {})
        create_styled_paragraph(doc, l1.get("survey_title", ""), bold=True)
        for item in l1.get("items", []):
            create_styled_paragraph(doc, f"  • {item.get('item', '')} (Scale: {item.get('scale', '')})")
        
        create_styled_heading(doc, "LEVEL 2: LEARNING", 1)
        l2 = data.get("level_2_learning", {})
        create_styled_heading(doc, "Pre-Test Questions", 3)
        for q in l2.get("pre_test", []):
            create_styled_paragraph(doc, f"  Q: {q.get('question', '')}")
        
        create_styled_heading(doc, "LEVEL 3: BEHAVIOR", 1)
        l3 = data.get("level_3_behavior", {})
        for obs in l3.get("observation_checklist", []):
            create_styled_paragraph(doc, f"  • {obs.get('behavior', '')}")
        
        create_styled_heading(doc, "LEVEL 4: RESULTS", 1)
        l4 = data.get("level_4_results", {})
        metrics = l4.get("metrics", [])
        if metrics:
            rows = [[m.get("metric", ""), m.get("baseline", ""), m.get("target", ""), m.get("data_source", "")] for m in metrics]
            create_styled_table(doc, ["Metric", "Baseline", "Target", "Data Source"], rows, [2.0, 1.5, 1.5, 1.5])
    else:
        # General assessment format
        create_styled_heading(doc, "1. ASSESSMENT OVERVIEW", 1)
        overview = data.get("assessment_overview", {})
        create_styled_paragraph(doc, f"Strategy: {overview.get('strategy', '')}")
        create_styled_paragraph(doc, f"Pass Criteria: {overview.get('pass_criteria', '')}")
        
        create_styled_heading(doc, "2. KNOWLEDGE TESTS", 1)
        for test in data.get("knowledge_tests", []):
            create_styled_paragraph(doc, f"{test.get('test_id', '')}: {test.get('objective_covered', '')}", bold=True)
            for q in test.get("questions", []):
                create_styled_paragraph(doc, f"  Q: {q.get('question', '')}")
                create_styled_paragraph(doc, f"     Answer: {q.get('correct_answer', '')}")
        
        create_styled_heading(doc, "3. PRACTICAL ASSESSMENTS", 1)
        for assess in data.get("practical_assessments", []):
            create_styled_paragraph(doc, f"{assess.get('assessment_id', '')}: {assess.get('task', '')}", bold=True)
            create_styled_paragraph(doc, f"Conditions: {assess.get('conditions', '')}")
            create_styled_paragraph(doc, f"Pass Standard: {assess.get('pass_standard', '')}")
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_evaluation_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path, eval_type: str):
    """Build evaluation document"""
    doc = Document()
    
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Title
    eval_name = terms['internal_eval'] if eval_type == "internal" else terms['external_eval']
    title = doc.add_heading(eval_name.upper(), 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(14)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("_" * 80)
    
    # Executive Summary
    create_styled_heading(doc, "1. EXECUTIVE SUMMARY", 1)
    exec_summary = data.get("executive_summary", "")
    sentences = split_into_sentences(exec_summary)
    for sentence in sentences:
        if sentence.strip():
            create_styled_paragraph(doc, sentence.strip())
    
    if framework == "KIRKPATRICK" and eval_type == "external":
        # Chain of Evidence format
        for level in ["level_1_evidence", "level_2_evidence", "level_3_evidence", "level_4_evidence"]:
            level_data = data.get(level, {})
            if level_data:
                level_num = level.split("_")[1]
                create_styled_heading(doc, f"LEVEL {level_num} EVIDENCE", 1)
                create_styled_paragraph(doc, f"Data Summary: {level_data.get('data_summary', '')}")
                create_styled_paragraph(doc, f"Key Finding: {level_data.get('key_finding', '')}")
        
        # ROI
        roi = data.get("roi_calculation", {})
        if roi:
            create_styled_heading(doc, "ROI CALCULATION", 1)
            create_styled_paragraph(doc, f"Benefits: {roi.get('benefits', '')}")
            create_styled_paragraph(doc, f"Costs: {roi.get('costs', '')}")
            create_styled_paragraph(doc, f"ROI: {roi.get('roi_percentage', '')}")
        
        # ROE
        create_styled_heading(doc, "RETURN ON EXPECTATIONS", 1)
        create_styled_paragraph(doc, data.get("roe_assessment", ""))
    
    elif eval_type == "internal":
        # Internal evaluation sections
        create_styled_heading(doc, "2. PILOT ANALYSIS", 1)
        pilot = data.get("pilot_analysis", {})
        create_styled_paragraph(doc, f"Course Dates: {pilot.get('course_dates', '')}")
        create_styled_paragraph(doc, f"Participants: {pilot.get('participants', '')}")
        create_styled_paragraph(doc, f"Completion Rate: {pilot.get('completion_rate', '')}")
        
        create_styled_heading(doc, "3. TRAINEE FEEDBACK", 1)
        feedback = data.get("trainee_feedback", {})
        create_styled_paragraph(doc, f"Satisfaction: {feedback.get('satisfaction_score', '')}")
        create_styled_paragraph(doc, f"Relevance: {feedback.get('relevance_score', '')}")
        
        create_styled_heading(doc, "4. PASS/FAIL ANALYSIS", 1)
        pf = data.get("pass_fail_analysis", {})
        create_styled_paragraph(doc, f"Overall Pass Rate: {pf.get('overall_pass_rate', '')}")
        
    else:
        # External evaluation sections
        create_styled_heading(doc, "2. SAMPLING STRATEGY", 1)
        sampling = data.get("sampling_strategy", {})
        create_styled_paragraph(doc, f"Population: {sampling.get('population', '')}")
        create_styled_paragraph(doc, f"Sample Size: {sampling.get('sample_size', '')}")
        
        create_styled_heading(doc, "3. WORKPLACE PERFORMANCE", 1)
        wp = data.get("workplace_performance", {})
        create_styled_paragraph(doc, f"Methodology: {wp.get('methodology', '')}")
        create_styled_paragraph(doc, f"Overall Assessment: {wp.get('overall_assessment', '')}")
        
        create_styled_heading(doc, "4. TRANSFER ASSESSMENT", 1)
        transfer = data.get("transfer_assessment", {})
        create_styled_paragraph(doc, f"Transfer Rate: {transfer.get('transfer_rate', '')}")
    
    # Recommendations
    create_styled_heading(doc, "RECOMMENDATIONS", 1)
    recommendations = data.get("recommendations", [])
    if recommendations:
        rows = [[r.get("id", ""), r.get("recommendation", ""), r.get("priority", r.get("owner", ""))] for r in recommendations]
        create_styled_table(doc, ["ID", "Recommendation", "Priority/Owner"], rows, [0.6, 4.5, 1.2])
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)


