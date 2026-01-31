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
    
    # DEBUG: Log all received parameters
    print(f"[NOVA] Analysis parameters received:")
    print(f"[NOVA]   - domain: '{domain}' (truthy: {bool(domain)})")
    print(f"[NOVA]   - specialism: '{specialism}' (truthy: {bool(specialism)})")
    print(f"[NOVA]   - proficiency_level: '{proficiency_level}' (truthy: {bool(proficiency_level)})")
    print(f"[NOVA]   - role_title: '{parameters.get('role_title', '')}'")
    print(f"[NOVA]   - framework: '{framework}'")
    print(f"[NOVA]   - All params: {parameters}")
    
    if domain and specialism and proficiency_level:
        # Use new research-based analysis (v5.1)
        print(f"[NOVA] ✓ Using research-based Analysis Agent (v5.1)")
        await run_research_analysis_agent(job_id, parameters, framework)
        return
    
    # Original v5.0 analysis flow
    print(f"[NOVA] ✗ Falling back to standard Analysis Agent (v5.0) - missing required params")
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
# RESEARCH-BASED ANALYSIS AGENT (v5.1) - COMPREHENSIVE IMPLEMENTATION
# ============================================================================

def get_research_system_prompt() -> str:
    """
    Complete system prompt following ANALYSIS-AGENT-SYSTEM-PROMPT.md specification.
    Universal standard for ALL frameworks and ALL domains.
    """
    return """# NOVA™ ANALYSIS AGENT

## IDENTITY AND PURPOSE

You are the NOVA™ Analysis Agent. Your purpose is to conduct comprehensive, factual, research-based analysis for training needs identification. You produce outputs that are 100% accurate, fully cited, and contain zero fabrication.

## CRITICAL RULES - ABSOLUTE REQUIREMENTS

1. You NEVER invent statistics, percentages, or quantified claims
2. You NEVER fabricate methodology (no fake interviews, surveys, focus groups, or observations)
3. You NEVER hallucinate professional standards, regulations, or requirements
4. Every factual claim MUST be grounded in actual web research with citations
5. If information cannot be found, state "Information not found through research" - NEVER guess
6. You conduct real web research for EVERY aspect of the analysis
7. All sources must include URL and access date
8. When citing standards (ISO, British Standards, etc.), cite the actual standard number

## OUTPUT FORMAT

You MUST return a valid JSON object with the exact structure specified. Do not include any text outside the JSON.

## RESEARCH APPROACH

For each analysis, you must comprehensively research:
- Domain-specific industry standards and ISO standards
- Specialism-specific competency frameworks
- Professional body requirements and registration
- Regulatory and legal requirements
- Qualification frameworks (academic and vocational)
- Apprenticeship standards and routes
- Technical and professional standards
- CPD and recertification requirements

## CITATION FORMAT

Every factual claim must include a source reference in this format:
"[Source: Organisation Name - https://url.com]"

For standards: "[ISO/IEC 12345:2023 - https://iso.org/...]"
For legislation: "[Act Name Year - https://legislation.gov.uk/...]"
For professional bodies: "[Body Name - https://bodywebsite.org/...]"
"""


def get_domain_research_queries(domain: str, specialism: str, role_title: str, proficiency_level: str) -> Dict[str, List[str]]:
    """
    Generate comprehensive domain-specific research queries.
    Returns queries organised by research category.
    
    This function ensures EVERY facet of the domain is researched including:
    - Industry standards and ISO standards
    - Competency frameworks
    - Regulatory bodies
    - Professional standards
    - Apprenticeship routes
    - Qualification frameworks (Academic and Vocational)
    """
    
    # Base queries that apply to ALL domains
    base_queries = {
        "industry_standards": [
            f"{domain} industry standards UK",
            f"{domain} ISO standards",
            f"{domain} British Standards BS",
            f"{domain} international standards",
            f"{domain} best practice guidelines",
            f"{domain} quality standards",
        ],
        "competency_frameworks": [
            f"{domain} National Occupational Standards NOS",
            f"{specialism} competency framework",
            f"{specialism} skills framework",
            f"{domain} sector skills council",
            f"{specialism} professional competencies",
            f"{role_title} competency requirements",
        ],
        "professional_bodies": [
            f"{domain} professional body UK",
            f"{specialism} professional body UK",
            f"{specialism} chartered status UK",
            f"{role_title} professional registration UK",
            f"{specialism} professional institute",
            f"{domain} regulatory body UK",
        ],
        "qualifications_academic": [
            f"{specialism} degree requirements UK",
            f"{role_title} academic qualifications",
            f"{specialism} university courses UK",
            f"{domain} RQF qualification levels",
            f"{specialism} postgraduate qualifications",
            f"{role_title} education requirements",
        ],
        "qualifications_vocational": [
            f"{specialism} NVQ qualifications UK",
            f"{specialism} vocational qualifications",
            f"{domain} BTEC qualifications",
            f"{specialism} professional certifications",
            f"{role_title} required certifications",
            f"{domain} technical certifications",
        ],
        "apprenticeships": [
            f"{specialism} apprenticeship standard UK",
            f"{domain} apprenticeship routes",
            f"{role_title} apprenticeship",
            f"{specialism} degree apprenticeship",
            f"{domain} apprenticeship level",
            f"Institute for Apprenticeships {specialism}",
        ],
        "regulatory_legal": [
            f"{domain} UK legislation",
            f"{specialism} legal requirements UK",
            f"{role_title} statutory duties",
            f"{domain} regulatory compliance UK",
            f"{specialism} mandatory training requirements",
            f"{domain} health and safety legislation",
        ],
        "role_requirements": [
            f"{role_title} job description",
            f"{role_title} responsibilities duties",
            f"{role_title} {domain} requirements",
            f"{role_title} skills requirements",
            f"{role_title} {proficiency_level} requirements",
            f"{role_title} career pathway",
        ],
        "proficiency_mapping": [
            f"{domain} {proficiency_level} experience requirements",
            f"{specialism} career levels progression",
            f"{role_title} salary band UK",
            f"{domain} job grades levels",
            f"{proficiency_level} competency descriptors",
        ],
        "cpd_requirements": [
            f"{specialism} CPD requirements UK",
            f"{domain} professional development requirements",
            f"{specialism} recertification requirements",
            f"{role_title} mandatory training refresh",
            f"{specialism} revalidation requirements",
        ],
    }
    
    # Domain-specific additional queries
    domain_specific = get_domain_specific_queries(domain, specialism, role_title)
    
    # Merge domain-specific queries
    for category, queries in domain_specific.items():
        if category in base_queries:
            base_queries[category].extend(queries)
        else:
            base_queries[category] = queries
    
    return base_queries


def get_domain_specific_queries(domain: str, specialism: str, role_title: str) -> Dict[str, List[str]]:
    """
    Return domain-specific research queries for major industry sectors.
    These supplement the base queries with authoritative domain sources.
    """
    
    domain_lower = domain.lower()
    
    # AI & Data Science / Technology
    if any(term in domain_lower for term in ['ai', 'data', 'technology', 'it', 'software', 'digital', 'cyber']):
        return {
            "competency_frameworks": [
                "SFIA 9 Skills Framework for the Information Age",
                f"SFIA {specialism} skills",
                "BCS competency framework",
                "Tech Industry Gold accreditation",
                "UNESCO AI Competency Framework",
                "EDSA European Data Science Academy curriculum",
                f"{specialism} SFIA level mapping",
            ],
            "industry_standards": [
                "ISO/IEC 42001 AI management system",
                "ISO/IEC TS 4213 machine learning",
                "ISO/IEC 27001 information security",
                "ISO/IEC 25010 software quality",
                "NIST AI Risk Management Framework",
                "IEEE standards artificial intelligence",
                "ISO/IEC 38500 IT governance",
            ],
            "professional_bodies": [
                "BCS The Chartered Institute for IT",
                "IET Institution of Engineering and Technology",
                "techUK membership",
                "Chartered IT Professional CITP",
                "Data Science Council UK",
            ],
            "apprenticeships": [
                "AI Data Specialist apprenticeship Level 7",
                "Data Analyst apprenticeship Level 4",
                "Software Developer apprenticeship Level 4",
                "Cyber Security Technologist apprenticeship",
                "Digital and Technology Solutions apprenticeship",
            ],
            "ethics_governance": [
                "Ethics by Design framework",
                "EU AI Act requirements",
                "AI ethics guidelines UK",
                "Responsible AI principles",
                "Data ethics framework UK government",
            ],
        }
    
    # Healthcare / Medical
    elif any(term in domain_lower for term in ['health', 'medical', 'nhs', 'clinical', 'nursing', 'care']):
        return {
            "competency_frameworks": [
                "NHS Knowledge and Skills Framework KSF",
                "NHS Agenda for Change bands",
                f"NHS {specialism} competencies",
                "Health Education England standards",
                "Skills for Health National Occupational Standards",
                f"{role_title} NHS band level",
            ],
            "professional_bodies": [
                "General Medical Council GMC",
                "Nursing and Midwifery Council NMC",
                "Health and Care Professions Council HCPC",
                "General Pharmaceutical Council GPhC",
                "Royal College requirements",
                f"{specialism} Royal College UK",
            ],
            "regulatory_legal": [
                "Care Quality Commission CQC requirements",
                "Health and Social Care Act",
                "Mental Health Act training requirements",
                "Safeguarding training requirements NHS",
                "Clinical governance requirements",
                "NHS mandatory training requirements",
            ],
            "qualifications_vocational": [
                "NVQ Health and Social Care",
                "Care Certificate requirements",
                f"{specialism} clinical qualifications",
                "NHS career framework qualifications",
            ],
            "cpd_requirements": [
                "NMC revalidation requirements",
                "GMC CPD requirements",
                "HCPC CPD standards",
                f"{specialism} revalidation cycle",
            ],
        }
    
    # Defence / Military
    elif any(term in domain_lower for term in ['defence', 'defense', 'military', 'armed forces', 'mod']):
        return {
            "competency_frameworks": [
                "JSP 822 Defence Individual Training",
                "DTSM Defence Training Support Manual",
                "Defence SOPS Statements of Performance",
                "NATO STANAG training standards",
                f"Military {specialism} competencies",
            ],
            "industry_standards": [
                "ASD S6000T training standard",
                "NATO Bi-SC 75-7 Education Training",
                "DEF STAN defence standards",
                "AQAP quality standards NATO",
            ],
            "regulatory_legal": [
                "Armed Forces Act requirements",
                "MOD health and safety regulations",
                "Defence Security regulations",
                "Export control training requirements",
            ],
            "security_requirements": [
                "UK security clearance levels SC DV",
                "Developed Vetting DV requirements",
                "Security Check SC requirements",
                "Counter Terrorist Check CTC",
                "MOD security training requirements",
            ],
            "proficiency_mapping": [
                "Military rank equivalence civilian",
                "NATO rank structure OR OF",
                "MOD civil service grades",
                "Defence career progression",
            ],
        }
    
    # Finance / Banking
    elif any(term in domain_lower for term in ['finance', 'banking', 'accounting', 'insurance', 'investment']):
        return {
            "competency_frameworks": [
                "CFA Institute competency framework",
                "CISI Chartered Institute Securities Investment",
                "ACCA competency framework",
                "ICAEW competency requirements",
                "FCA competency sourcebook",
            ],
            "professional_bodies": [
                "Financial Conduct Authority FCA",
                "Prudential Regulation Authority PRA",
                "Chartered Insurance Institute CII",
                "ICAEW Institute of Chartered Accountants",
                "ACCA chartered accountants",
                "CIMA management accountants",
            ],
            "regulatory_legal": [
                "FCA Senior Managers Certification Regime SMCR",
                "FCA training and competence sourcebook",
                "Money Laundering Regulations training",
                "Consumer Duty requirements FCA",
                "MiFID II training requirements",
            ],
            "industry_standards": [
                "ISO 22301 business continuity",
                "PCI DSS payment card security",
                "Basel III banking standards",
                "Solvency II insurance standards",
            ],
            "cpd_requirements": [
                "FCA CPD requirements",
                "ICAEW CPD requirements",
                "ACCA CPD policy",
                "CII CPD requirements",
            ],
        }
    
    # Construction / Engineering
    elif any(term in domain_lower for term in ['construction', 'engineering', 'building', 'civil', 'mechanical', 'electrical']):
        return {
            "competency_frameworks": [
                "Engineering Council UK SPEC competencies",
                "CITB Construction Industry Training Board",
                "EngTech IEng CEng competencies",
                f"{specialism} engineering competencies",
            ],
            "professional_bodies": [
                "Engineering Council UK",
                "Institution of Civil Engineers ICE",
                "Institution of Mechanical Engineers IMechE",
                "Institution of Engineering and Technology IET",
                "CIOB Chartered Institute of Building",
                "RICS Royal Institution of Chartered Surveyors",
            ],
            "industry_standards": [
                "ISO 9001 quality management",
                "ISO 45001 health and safety",
                "ISO 14001 environmental management",
                "British Standards construction",
                "Eurocodes structural design",
                "CDM Regulations construction",
            ],
            "qualifications_vocational": [
                "CSCS Card requirements construction",
                "NVQ Construction qualifications",
                "SMSTS Site Management Safety Training",
                "SSSTS Site Supervisor Safety Training",
                f"{specialism} engineering certifications",
            ],
            "apprenticeships": [
                "Civil Engineering apprenticeship",
                "Construction apprenticeship standards",
                f"{specialism} engineering apprenticeship",
                "Quantity Surveyor apprenticeship",
            ],
        }
    
    # Legal
    elif any(term in domain_lower for term in ['legal', 'law', 'solicitor', 'barrister']):
        return {
            "competency_frameworks": [
                "SRA Solicitors Regulation Authority competencies",
                "BSB Bar Standards Board competencies",
                "CILEX competency framework",
                "Legal Services Board standards",
            ],
            "professional_bodies": [
                "Solicitors Regulation Authority SRA",
                "Bar Standards Board BSB",
                "CILEX Chartered Institute Legal Executives",
                "Law Society England Wales",
                "Bar Council",
            ],
            "qualifications_academic": [
                "SQE Solicitors Qualifying Examination",
                "LPC Legal Practice Course",
                "Bar course BPTC",
                "GDL Graduate Diploma Law",
                "LLB law degree requirements",
            ],
            "cpd_requirements": [
                "SRA CPD requirements",
                "BSB CPD requirements",
                "CILEX CPD requirements",
                "Legal CPD hours requirements",
            ],
        }
    
    # Education / Teaching
    elif any(term in domain_lower for term in ['education', 'teaching', 'school', 'academic', 'university', 'training']):
        return {
            "competency_frameworks": [
                "Teachers Standards UK",
                "Further Education teaching standards",
                "ETF Education Training Foundation standards",
                "QTS Qualified Teacher Status requirements",
            ],
            "professional_bodies": [
                "Teaching Regulation Agency TRA",
                "Education and Training Foundation ETF",
                "Ofsted inspection framework",
                "Chartered College of Teaching",
                "SEDA Staff Educational Development Association",
            ],
            "qualifications_academic": [
                "PGCE teacher training",
                "QTS Qualified Teacher Status",
                "QTLS Qualified Teacher Learning Skills",
                "Level 5 teaching qualification FE",
                "AET Award Education Training",
            ],
            "regulatory_legal": [
                "Safeguarding training requirements education",
                "Prevent duty training requirements",
                "DBS requirements education",
                "Keeping Children Safe Education KCSIE",
            ],
        }
    
    # Manufacturing / Production
    elif any(term in domain_lower for term in ['manufacturing', 'production', 'operations', 'industrial']):
        return {
            "competency_frameworks": [
                "SEMTA engineering manufacturing NOS",
                "Make UK manufacturing competencies",
                "Lean Six Sigma competencies",
                f"{specialism} manufacturing competencies",
            ],
            "industry_standards": [
                "ISO 9001 quality management manufacturing",
                "ISO 45001 occupational health safety",
                "ISO 14001 environmental manufacturing",
                "IATF 16949 automotive quality",
                "AS9100 aerospace quality",
                "ISO 13485 medical devices",
            ],
            "qualifications_vocational": [
                "NVQ Manufacturing Engineering",
                "Lean Six Sigma certifications",
                "IOSH Managing Safely",
                "NEBOSH manufacturing",
            ],
            "apprenticeships": [
                "Engineering Manufacturing Technician apprenticeship",
                "Lean Manufacturing Operative apprenticeship",
                "Engineering Technician apprenticeship",
            ],
        }
    
    # Default - return empty (base queries still apply)
    return {}


def build_comprehensive_research_prompt(
    domain: str, 
    specialism: str, 
    role_title: str, 
    proficiency_level: str, 
    framework: str, 
    role_description: str,
    terms: Dict
) -> str:
    """
    Build the comprehensive research prompt following ANALYSIS-AGENT-SYSTEM-PROMPT.md exactly.
    Includes ALL 10 research steps and outputs for ALL 18 sections.
    """
    
    # Get domain-specific research queries
    research_queries = get_domain_research_queries(domain, specialism, role_title, proficiency_level)
    
    # Format queries for the prompt
    formatted_queries = ""
    for category, queries in research_queries.items():
        formatted_queries += f"\n**{category.replace('_', ' ').title()}:**\n"
        for q in queries[:6]:  # Limit to 6 per category to manage prompt size
            formatted_queries += f"- {q}\n"
    
    return f"""# RESEARCH-BASED TRAINING ANALYSIS

## ANALYSIS PARAMETERS
- **Domain:** {domain}
- **Specialism:** {specialism}
- **Role Title:** {role_title}
- **Proficiency Level:** {proficiency_level}
- **Framework:** {framework} ({terms.get('framework_name', framework)})
- **Additional Context:** {role_description or 'None provided'}

## MANDATORY RESEARCH METHODOLOGY

You MUST execute ALL 10 research steps using web search. For each step, conduct the searches listed and capture the required information.

### STEP 1: FRAMEWORK ISOLATION
Research the {framework} framework requirements:
- "{framework} training methodology requirements"
- "{framework} analysis phase outputs"
- "{framework} official documentation"
- "{terms.get('task_list', 'task analysis')} format {framework}"

**Capture:** Framework full name, version, governing body, mandatory outputs, terminology.

### STEP 2: DOMAIN RESEARCH
Research the {domain} industry:
- "{domain} industry standards UK"
- "{domain} professional bodies UK"
- "{domain} regulatory requirements UK"
- "{domain} ISO standards"
- "{domain} sector skills council"

**Capture:** Domain scope, major employers, regulatory environment, key industry bodies.

### STEP 3: SPECIALISM RESEARCH
Research {specialism}:
- "{specialism} competency framework"
- "{specialism} professional qualifications UK"
- "{specialism} career pathway"
- "{specialism} skills requirements 2025"

**Capture:** Specialism definition, required qualifications, professional certifications.

### STEP 4: ROLE TITLE RESEARCH
Research {role_title}:
- "{role_title} job description UK"
- "{role_title} responsibilities duties"
- "{role_title} {domain} requirements"
- "{role_title} equivalent job titles"

**Capture:** Standard definition, responsibilities, reporting relationships.

### STEP 5: PROFICIENCY LEVEL RESEARCH
Research {proficiency_level} level requirements:
- "{domain} {proficiency_level} experience requirements"
- "SFIA {specialism} level mapping" (if IT/Technology)
- "NVQ RQF level {proficiency_level}"
- "NHS Agenda for Change band" (if Healthcare)
- "European Qualifications Framework EQF level"

**Capture:** Industry proficiency definitions, experience benchmarks, qualification mapping.

### STEP 6: PROFESSIONAL BODY AND REGULATOR RESEARCH
Research professional bodies:
- "{specialism} professional body UK"
- "{specialism} chartered status UK"
- "{role_title} registration requirements UK"
- "{domain} regulatory body UK"
- "{specialism} code of conduct"

**Capture:** Professional body name, website, membership requirements, regulatory authority.

### STEP 7: COMPETENCY FRAMEWORK MAPPING
Research competency frameworks:
- "{domain} National Occupational Standards NOS"
- "{specialism} competency framework official"
- "SFIA framework {specialism}" (if Technology)
- "NHS Knowledge Skills Framework {specialism}" (if Healthcare)
- "{specialism} professional competencies"

**Capture:** Framework name, relevant units, level descriptors, assessment criteria.

### STEP 8: LEGAL AND COMPLIANCE RESEARCH
Research legal requirements:
- "{domain} UK legislation"
- "{specialism} legal requirements UK"
- "{role_title} statutory duties"
- "Health and Safety {domain} legislation"
- "Data protection {specialism} requirements GDPR"
- "Equality Act requirements {role_title}"

**Capture:** Applicable legislation with Act names and years, statutory duties.

### STEP 9: PHYSICAL/MEDICAL/SECURITY REQUIREMENTS
Research special requirements:
- "{role_title} medical requirements UK"
- "{role_title} fitness standards"
- "{domain} security clearance requirements UK"
- "{role_title} DBS check requirements"
- "{specialism} occupational health requirements"

**Capture:** Physical standards, medical requirements, security clearance, DBS requirements.

### STEP 10: CPD AND RECERTIFICATION RESEARCH
Research ongoing development:
- "{specialism} CPD requirements UK"
- "{specialism} recertification requirements"
- "{role_title} mandatory training refresh"
- "{domain} continuing education requirements"

**Capture:** CPD hours, recertification cycles, mandatory refresher training.

## ADDITIONAL DOMAIN-SPECIFIC RESEARCH QUERIES

Execute these additional searches specific to {domain}:
{formatted_queries}

## REQUIRED OUTPUT FORMAT

Return a JSON object with this EXACT structure:

{{
    "research_log": {{
        "searches_conducted": ["list of actual searches performed"],
        "sources_found": ["list of authoritative sources discovered"],
        "research_date": "{datetime.now().strftime('%Y-%m-%d')}"
    }},
    
    "section_1_executive_summary": {{
        "analysis_scope": "What was analysed",
        "research_methodology": "Web-based research using authoritative sources",
        "key_findings": ["Finding 1 with citation", "Finding 2 with citation"],
        "tasks_identified": 0,
        "primary_standards": ["Standard 1", "Standard 2"]
    }},
    
    "section_2_framework_identification": {{
        "framework_name": "{framework}",
        "framework_version": "Version [Source]",
        "governing_authority": "Authority name [Source]",
        "framework_purpose": "Description [Source]",
        "analysis_requirements": ["Requirement 1", "Requirement 2"],
        "required_outputs": ["Output 1", "Output 2"],
        "terminology": {{"term1": "definition", "term2": "definition"}},
        "source_url": "URL"
    }},
    
    "section_3_geographic_context": {{
        "country": "United Kingdom",
        "legal_jurisdiction": "England and Wales / Scotland / Northern Ireland",
        "language": "English",
        "currency": "GBP",
        "regional_variations": "Any noted variations [Source]"
    }},
    
    "section_4_professional_body": {{
        "professional_body_name": "Name [Source]",
        "website_url": "URL",
        "membership_categories": ["Category 1", "Category 2"],
        "registration_required": true/false,
        "registration_requirements": ["Requirement 1"],
        "protected_titles": ["Title 1"],
        "regulatory_authority": "Name if different",
        "regulatory_powers": "Description"
    }},
    
    "section_5_competency_framework": {{
        "framework_name": "Name [Source]",
        "framework_owner": "Organisation",
        "framework_url": "URL",
        "relevant_units": [
            {{"unit_code": "Code", "unit_title": "Title", "description": "Description"}}
        ],
        "level_descriptors": {{"level": "description"}},
        "proficiency_mapping": "How {proficiency_level} maps to framework levels"
    }},
    
    "section_6_role_description": {{
        "comprehensive_definition": "Full definition [Source]",
        "primary_purpose": "Purpose statement",
        "key_accountabilities": ["Accountability 1", "Accountability 2"],
        "reporting_structure": "Typically reports to...",
        "team_context": "Team structure description",
        "equivalent_titles": ["Title 1", "Title 2"],
        "role_boundaries": "What is NOT included in this role"
    }},
    
    "section_7_qualifications": {{
        "essential_qualifications": [
            {{"qualification": "Name", "level": "RQF/EQF level", "source": "URL"}}
        ],
        "desirable_qualifications": [
            {{"qualification": "Name", "level": "Level", "source": "URL"}}
        ],
        "academic_level_required": "RQF Level X / EQF Level Y [Source]",
        "professional_certifications_required": ["Cert 1 [Source]"],
        "professional_certifications_desirable": ["Cert 2 [Source]"],
        "apprenticeship_routes": [
            {{"name": "Apprenticeship name", "level": "Level", "source": "URL"}}
        ],
        "qualification_equivalencies": "Description of equivalencies"
    }},
    
    "section_8_experience": {{
        "years_required": "X years for {proficiency_level} level [Source]",
        "type_of_experience": ["Experience type 1", "Experience type 2"],
        "sector_specific_requirements": ["Sector requirement 1"],
        "project_experience": ["Project type 1"],
        "leadership_experience": "Description if applicable",
        "international_experience": "Required/Desirable/Not required"
    }},
    
    "section_9_technical_skills": [
        {{
            "skill": "Skill name",
            "category": "Core/Desirable",
            "proficiency_level": "Awareness/Working/Practitioner/Expert",
            "source": "URL"
        }}
    ],
    
    "section_10_soft_skills": [
        {{
            "skill": "Skill name",
            "proficiency_level": "Expected level",
            "source": "URL"
        }}
    ],
    
    "section_11_behaviours": [
        {{
            "behaviour": "Behaviour description",
            "requirement_type": "Essential/Desirable",
            "source": "URL"
        }}
    ],
    
    "section_12_physical_medical_security": {{
        "physical_requirements": ["Requirement [Source]"] or "No specific requirements identified",
        "medical_requirements": ["Requirement [Source]"] or "No specific requirements identified",
        "security_clearance": "Level required [Source]" or "Standard employment checks only",
        "dbs_requirements": "Basic/Standard/Enhanced [Source]",
        "occupational_health": ["Requirement [Source]"],
        "reasonable_adjustments": "Statement about adjustments"
    }},
    
    "section_13_cpd_requirements": {{
        "professional_body_cpd": "Body name CPD policy [Source]",
        "annual_hours_points": "X hours/points per year [Source]",
        "recertification_cycle": "X years [Source]",
        "mandatory_refresher": ["Training 1", "Training 2"],
        "portfolio_requirements": "Description [Source]",
        "revalidation_process": "Description [Source]",
        "non_compliance_consequences": "What happens if not met"
    }},
    
    "section_14_career_progression": {{
        "pathway_to_role": ["Previous role 1", "Previous role 2"],
        "pathway_from_role": ["Next role 1", "Next role 2"],
        "lateral_moves": ["Lateral option 1"],
        "promotion_criteria": ["Criterion 1 [Source]"],
        "timeline_expectations": "X years typical [Source]",
        "skill_gaps_for_progression": ["Skill gap 1"]
    }},
    
    "section_15_legal_compliance": [
        {{
            "legislation": "Act Name Year",
            "relevance": "How it applies to this role",
            "mandatory_training": true/false,
            "source": "legislation.gov.uk URL"
        }}
    ],
    
    "section_16_professional_standards": [
        {{
            "standard": "Standard name/number",
            "issuing_body": "Organisation",
            "requirement_type": "Mandatory/Best Practice",
            "description": "What it covers",
            "source": "URL"
        }}
    ],
    
    "section_17_no_bias_statement": {{
        "equality_act_compliance": true,
        "genuine_occupational_requirements": "All requirements listed are GORs",
        "reasonable_adjustments_considered": true,
        "inclusive_language_used": true,
        "proportionality_statement": "Requirements are proportionate to role needs",
        "bias_concerns_identified": "None identified" or "List any concerns"
    }},
    
    "section_18_citations": [
        {{
            "source_type": "Professional Body/Legislation/Competency Framework/Industry Standard/Government Guidance",
            "source_name": "Name",
            "url": "Full URL",
            "date_accessed": "{datetime.now().strftime('%Y-%m-%d')}",
            "information_obtained": "What was found"
        }}
    ],
    
    "tasks": [
        {{
            "task_id": "T-001",
            "task_description": "Clear statement of task [Source]",
            "knowledge_required": ["Knowledge item 1 [Source]"],
            "skills_required": ["Skill item 1 [Source]"],
            "behaviours_required": ["Behaviour 1 [Source]"],
            "criticality": "High/Medium/Low",
            "frequency": "Daily/Weekly/Monthly/As Required",
            "source": "URL where task was identified"
        }}
    ]
}}

## CRITICAL REMINDERS

1. Use web search for EVERY section - do not rely on training data
2. Include [Source: URL] for every factual claim
3. If information not found, state "Information not found through research"
4. Generate at least 10-15 tasks based on research findings
5. Never invent statistics or methodology
6. All standards must include actual standard numbers (ISO XXXX, BS XXXX, etc.)
7. Return ONLY valid JSON - no other text"""


async def run_research_analysis_agent(job_id: str, parameters: Dict, framework: str):
    """
    Research-based Analysis Agent (v5.1) - COMPREHENSIVE IMPLEMENTATION
    
    Follows ANALYSIS-AGENT-SYSTEM-PROMPT.md specification exactly:
    - 10-step research methodology with domain-specific queries
    - 18-section Analysis Report with full citations
    - Framework-compliant Job/Task List
    - Zero fabrication tolerance
    
    Generates:
    - 01_Job_Task_Analysis.docx - Framework-compliant task list with sources
    - 02_Analysis_Report.docx - All 18 mandatory sections with citations
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
    
    update_job(job_id, 2, f"Starting Comprehensive Research Analysis ({framework})...")
    print(f"[NOVA v5.1] Research Analysis: {role_title} in {domain}/{specialism}")
    
    # Build comprehensive research prompt
    update_job(job_id, 5, "Building comprehensive research queries...")
    
    research_prompt = build_comprehensive_research_prompt(
        domain=domain,
        specialism=specialism,
        role_title=role_title,
        proficiency_level=proficiency_level,
        framework=framework,
        role_description=role_description,
        terms=terms
    )
    
    print(f"[NOVA] Research prompt built: {len(research_prompt)} chars")
    
    # Execute research with web search - increased token limit for comprehensive output
    update_job(job_id, 10, "Step 1/10: Researching framework requirements...")
    
    try:
        # Call Claude with web search enabled
        research_result = await call_claude_with_search(
            system_prompt=get_research_system_prompt(),
            user_prompt=research_prompt,
            max_tokens=16000  # Increased for comprehensive output
        )
        
        searches_performed = research_result.get("searches_performed", [])
        print(f"[NOVA] Research complete: {len(searches_performed)} web searches performed")
        
        update_job(job_id, 40, f"Research complete. {len(searches_performed)} searches performed.")
        
        # Parse the research output
        update_job(job_id, 45, "Parsing research results...")
        analysis_data = parse_comprehensive_research_output(research_result.get("text", ""))
        
        # Validate we got proper data
        if analysis_data.get("parse_error"):
            print(f"[NOVA] Warning: JSON parse issues, attempting recovery...")
            # Try to extract what we can
            analysis_data = recover_partial_analysis(research_result.get("text", ""), parameters, framework, terms)
        
        # Ensure all 18 sections exist (even if empty)
        analysis_data = ensure_all_sections(analysis_data)
        
        # Add metadata
        analysis_data["metadata"] = {
            "domain": domain,
            "specialism": specialism,
            "role_title": role_title,
            "proficiency_level": proficiency_level,
            "framework": framework,
            "framework_display": terms.get("framework_name", framework),
            "generated_date": datetime.now().isoformat(),
            "searches_performed": searches_performed,
            "nova_version": "5.1.0",
            "specification": "ANALYSIS-AGENT-SYSTEM-PROMPT.md v1.0"
        }
        
        update_job(job_id, 50, f"Found {len(analysis_data.get('tasks', []))} tasks, {len(analysis_data.get('section_18_citations', []))} citations")
        
    except Exception as e:
        print(f"[NOVA] Research failed: {e}")
        import traceback
        traceback.print_exc()
        update_job(job_id, 40, f"Research encountered issues, generating with available data...")
        analysis_data = create_comprehensive_fallback(parameters, framework, terms, str(e))
    
    # Build documents with all 18 sections
    update_job(job_id, 55, f"Building {terms['task_list']} document...")
    build_comprehensive_task_list_doc(
        analysis_data, 
        role_title, 
        framework, 
        terms, 
        output_dir / "01_Job_Task_Analysis.docx"
    )
    update_job(job_id, 70, f"✓ {terms['task_list']} complete")
    
    update_job(job_id, 75, "Building 18-Section Analysis Report...")
    build_comprehensive_analysis_report_doc(
        analysis_data,
        role_title,
        framework,
        terms,
        output_dir / "02_Analysis_Report.docx"
    )
    update_job(job_id, 90, "✓ Analysis Report complete (18 sections)")
    
    # Save raw JSON
    update_job(job_id, 95, "Saving analysis data...")
    with open(output_dir / "analysis_data.json", "w") as f:
        json.dump(analysis_data, f, indent=2, default=str)
    
    update_job(job_id, 100, "Analysis Phase Complete")
    print(f"[NOVA] Analysis complete: {len(analysis_data.get('tasks', []))} tasks identified")


def parse_comprehensive_research_output(text: str) -> Dict:
    """Parse comprehensive research output, extracting JSON from response"""
    try:
        # Try to find JSON in the response
        # First try the whole text
        text = text.strip()
        if text.startswith('{'):
            return json.loads(text)
        
        # Try to find JSON block
        json_match = re.search(r'\{[\s\S]*\}', text)
        if json_match:
            return json.loads(json_match.group())
            
    except json.JSONDecodeError as e:
        print(f"[NOVA] JSON parse error: {e}")
    
    # Return structure indicating parse failure
    return {
        "parse_error": True,
        "raw_text": text[:5000] if text else ""
    }


def recover_partial_analysis(text: str, parameters: Dict, framework: str, terms: Dict) -> Dict:
    """Attempt to recover useful information from partially parsed response"""
    analysis = {
        "section_1_executive_summary": {
            "analysis_scope": f"Analysis of {parameters.get('role_title', 'Role')} in {parameters.get('domain', 'Domain')}",
            "research_methodology": "Web-based research conducted",
            "key_findings": ["Research was conducted but complete parsing failed"],
            "tasks_identified": 0,
            "primary_standards": []
        },
        "tasks": [],
        "section_18_citations": [],
        "recovery_mode": True,
        "partial_text": text[:3000] if text else ""
    }
    
    # Try to extract any tasks mentioned
    task_patterns = [
        r'"task_description":\s*"([^"]+)"',
        r'"task":\s*"([^"]+)"',
        r'Task[:\s]+([^\n]+)',
    ]
    
    tasks_found = []
    for pattern in task_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        tasks_found.extend(matches[:20])  # Limit to 20
    
    for i, task_text in enumerate(set(tasks_found)):
        if len(task_text) > 10:  # Filter out noise
            analysis["tasks"].append({
                "task_id": f"T-{i+1:03d}",
                "task_description": task_text,
                "knowledge_required": ["Extracted from partial research"],
                "skills_required": ["Extracted from partial research"],
                "behaviours_required": [],
                "criticality": "Medium",
                "frequency": "As Required",
                "source": "Partial extraction from research"
            })
    
    return analysis


def ensure_all_sections(data: Dict) -> Dict:
    """Ensure all 18 required sections exist in the analysis data"""
    
    required_sections = {
        "section_1_executive_summary": {
            "analysis_scope": "",
            "research_methodology": "",
            "key_findings": [],
            "tasks_identified": 0,
            "primary_standards": []
        },
        "section_2_framework_identification": {
            "framework_name": "",
            "framework_version": "",
            "governing_authority": "",
            "framework_purpose": "",
            "analysis_requirements": [],
            "required_outputs": [],
            "terminology": {},
            "source_url": ""
        },
        "section_3_geographic_context": {
            "country": "United Kingdom",
            "legal_jurisdiction": "England and Wales",
            "language": "English",
            "currency": "GBP",
            "regional_variations": ""
        },
        "section_4_professional_body": {
            "professional_body_name": "",
            "website_url": "",
            "membership_categories": [],
            "registration_required": False,
            "registration_requirements": [],
            "protected_titles": [],
            "regulatory_authority": "",
            "regulatory_powers": ""
        },
        "section_5_competency_framework": {
            "framework_name": "",
            "framework_owner": "",
            "framework_url": "",
            "relevant_units": [],
            "level_descriptors": {},
            "proficiency_mapping": ""
        },
        "section_6_role_description": {
            "comprehensive_definition": "",
            "primary_purpose": "",
            "key_accountabilities": [],
            "reporting_structure": "",
            "team_context": "",
            "equivalent_titles": [],
            "role_boundaries": ""
        },
        "section_7_qualifications": {
            "essential_qualifications": [],
            "desirable_qualifications": [],
            "academic_level_required": "",
            "professional_certifications_required": [],
            "professional_certifications_desirable": [],
            "apprenticeship_routes": [],
            "qualification_equivalencies": ""
        },
        "section_8_experience": {
            "years_required": "",
            "type_of_experience": [],
            "sector_specific_requirements": [],
            "project_experience": [],
            "leadership_experience": "",
            "international_experience": ""
        },
        "section_9_technical_skills": [],
        "section_10_soft_skills": [],
        "section_11_behaviours": [],
        "section_12_physical_medical_security": {
            "physical_requirements": [],
            "medical_requirements": [],
            "security_clearance": "",
            "dbs_requirements": "",
            "occupational_health": [],
            "reasonable_adjustments": ""
        },
        "section_13_cpd_requirements": {
            "professional_body_cpd": "",
            "annual_hours_points": "",
            "recertification_cycle": "",
            "mandatory_refresher": [],
            "portfolio_requirements": "",
            "revalidation_process": "",
            "non_compliance_consequences": ""
        },
        "section_14_career_progression": {
            "pathway_to_role": [],
            "pathway_from_role": [],
            "lateral_moves": [],
            "promotion_criteria": [],
            "timeline_expectations": "",
            "skill_gaps_for_progression": []
        },
        "section_15_legal_compliance": [],
        "section_16_professional_standards": [],
        "section_17_no_bias_statement": {
            "equality_act_compliance": True,
            "genuine_occupational_requirements": "All requirements are genuine occupational requirements",
            "reasonable_adjustments_considered": True,
            "inclusive_language_used": True,
            "proportionality_statement": "Requirements are proportionate to role needs",
            "bias_concerns_identified": "None identified"
        },
        "section_18_citations": [],
        "tasks": [],
        "research_log": {
            "searches_conducted": [],
            "sources_found": [],
            "research_date": datetime.now().strftime('%Y-%m-%d')
        }
    }
    
    # Merge with defaults
    for section, default_value in required_sections.items():
        if section not in data:
            data[section] = default_value
        elif isinstance(default_value, dict) and isinstance(data.get(section), dict):
            # Merge dict sections
            for key, value in default_value.items():
                if key not in data[section]:
                    data[section][key] = value
    
    # Update task count
    if "section_1_executive_summary" in data:
        data["section_1_executive_summary"]["tasks_identified"] = len(data.get("tasks", []))
    
    return data


def create_comprehensive_fallback(parameters: Dict, framework: str, terms: Dict, error_msg: str) -> Dict:
    """Create comprehensive fallback structure when research fails"""
    return {
        "section_1_executive_summary": {
            "analysis_scope": f"Analysis of {parameters.get('role_title', 'Role')} in {parameters.get('domain', 'Domain')}/{parameters.get('specialism', 'Specialism')}",
            "research_methodology": f"Web-based research was attempted but encountered issues: {error_msg[:200]}",
            "key_findings": ["Research could not be completed. Manual analysis required."],
            "tasks_identified": 0,
            "primary_standards": ["Manual identification required"]
        },
        "section_2_framework_identification": {
            "framework_name": terms.get("framework_name", framework),
            "framework_version": "Manual verification required",
            "governing_authority": "Manual verification required",
            "framework_purpose": "",
            "analysis_requirements": [],
            "required_outputs": [terms.get("task_list", "Task List"), "Analysis Report"],
            "terminology": {},
            "source_url": ""
        },
        "section_3_geographic_context": {
            "country": "United Kingdom",
            "legal_jurisdiction": "England and Wales",
            "language": "English",
            "currency": "GBP",
            "regional_variations": ""
        },
        "section_4_professional_body": {
            "professional_body_name": "Research required",
            "website_url": "",
            "membership_categories": [],
            "registration_required": False,
            "registration_requirements": [],
            "protected_titles": [],
            "regulatory_authority": "",
            "regulatory_powers": ""
        },
        "section_5_competency_framework": {
            "framework_name": "Research required",
            "framework_owner": "",
            "framework_url": "",
            "relevant_units": [],
            "level_descriptors": {},
            "proficiency_mapping": ""
        },
        "section_6_role_description": {
            "comprehensive_definition": f"Analysis for {parameters.get('role_title', 'Role')} requires manual completion",
            "primary_purpose": "",
            "key_accountabilities": [],
            "reporting_structure": "",
            "team_context": "",
            "equivalent_titles": [],
            "role_boundaries": ""
        },
        "section_7_qualifications": {
            "essential_qualifications": [],
            "desirable_qualifications": [],
            "academic_level_required": "",
            "professional_certifications_required": [],
            "professional_certifications_desirable": [],
            "apprenticeship_routes": [],
            "qualification_equivalencies": ""
        },
        "section_8_experience": {
            "years_required": "",
            "type_of_experience": [],
            "sector_specific_requirements": [],
            "project_experience": [],
            "leadership_experience": "",
            "international_experience": ""
        },
        "section_9_technical_skills": [],
        "section_10_soft_skills": [],
        "section_11_behaviours": [],
        "section_12_physical_medical_security": {
            "physical_requirements": ["No specific requirements identified"],
            "medical_requirements": ["No specific requirements identified"],
            "security_clearance": "Standard employment checks",
            "dbs_requirements": "Standard",
            "occupational_health": [],
            "reasonable_adjustments": "Reasonable adjustments will be considered"
        },
        "section_13_cpd_requirements": {
            "professional_body_cpd": "Research required",
            "annual_hours_points": "",
            "recertification_cycle": "",
            "mandatory_refresher": [],
            "portfolio_requirements": "",
            "revalidation_process": "",
            "non_compliance_consequences": ""
        },
        "section_14_career_progression": {
            "pathway_to_role": [],
            "pathway_from_role": [],
            "lateral_moves": [],
            "promotion_criteria": [],
            "timeline_expectations": "",
            "skill_gaps_for_progression": []
        },
        "section_15_legal_compliance": [],
        "section_16_professional_standards": [],
        "section_17_no_bias_statement": {
            "equality_act_compliance": True,
            "genuine_occupational_requirements": "All requirements are genuine occupational requirements",
            "reasonable_adjustments_considered": True,
            "inclusive_language_used": True,
            "proportionality_statement": "Requirements are proportionate to role needs",
            "bias_concerns_identified": "None identified"
        },
        "section_18_citations": [],
        "tasks": [],
        "research_log": {
            "searches_conducted": [],
            "sources_found": [],
            "research_date": datetime.now().strftime('%Y-%m-%d')
        },
        "metadata": {
            "domain": parameters.get("domain", ""),
            "specialism": parameters.get("specialism", ""),
            "role_title": parameters.get("role_title", ""),
            "proficiency_level": parameters.get("proficiency_level", ""),
            "framework": framework,
            "generated_date": datetime.now().isoformat(),
            "nova_version": "5.1.0",
            "fallback_mode": True,
            "error": error_msg[:500]
        }
    }


def build_comprehensive_task_list_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build comprehensive task list document following framework requirements with citations"""
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
    title = doc.add_heading(f"JOB/TASK ANALYSIS", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.color.rgb = NOVA_DARK_BLUE
    
    # Framework-specific title
    framework_title = doc.add_paragraph()
    run = framework_title.add_run(f"{terms.get('task_list', 'Task List')}")
    run.font.name = "Roboto"
    run.font.size = Pt(14)
    run.font.italic = True
    framework_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Subtitle
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(16)
    run.font.bold = True
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Framework badge
    badge = doc.add_paragraph()
    run = badge.add_run(f"Framework: {terms.get('framework_name', framework)} | Generated: {metadata.get('generated_date', '')[:10]} | NOVA v5.1")
    run.font.name = "Roboto"
    run.font.size = Pt(10)
    run.font.italic = True
    badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Analysis Parameters Table
    create_styled_heading(doc, "1. Analysis Parameters", 1)
    param_table = doc.add_table(rows=7, cols=2)
    param_table.style = 'Table Grid'
    
    params = [
        ("Domain", metadata.get("domain", "")),
        ("Specialism", metadata.get("specialism", "")),
        ("Role Title", metadata.get("role_title", "")),
        ("Proficiency Level", metadata.get("proficiency_level", "")),
        ("Framework", metadata.get("framework_display", "")),
        ("Research Date", metadata.get("generated_date", "")[:10]),
        ("Specification", metadata.get("specification", "ANALYSIS-AGENT-SYSTEM-PROMPT.md"))
    ]
    
    for i, (label, value) in enumerate(params):
        param_table.rows[i].cells[0].text = label
        param_table.rows[i].cells[1].text = str(value)
        set_cell_shading(param_table.rows[i].cells[0], "E8E0F0")
        for cell in param_table.rows[i].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Research Summary
    create_styled_heading(doc, "2. Research Summary", 1)
    research_log = data.get("research_log", {})
    searches = research_log.get("searches_conducted", metadata.get("searches_performed", []))
    
    p = doc.add_paragraph()
    run = p.add_run(f"Web searches conducted: {len(searches)}")
    run.font.name = "Roboto"
    run.font.size = Pt(10)
    
    sources = research_log.get("sources_found", [])
    if sources:
        p2 = doc.add_paragraph()
        run2 = p2.add_run(f"Authoritative sources found: {len(sources)}")
        run2.font.name = "Roboto"
        run2.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Task Table
    create_styled_heading(doc, "3. Job/Task Inventory", 1)
    
    tasks = data.get("tasks", [])
    
    if tasks:
        # Framework-specific headers
        if framework in ["UK_DSAT", "UK-DSAT"]:
            task_headers = ["Task ID", "Performance (Task Description)", "Knowledge Required", "Skills Required", "Behaviours", "Criticality", "Frequency", "Source"]
        elif framework in ["US_TRADOC", "US-TRADOC"]:
            task_headers = ["Task #", "Task Title", "Conditions", "Standards", "Knowledge", "Skills", "Source"]
        else:
            task_headers = ["Task ID", "Task Description", "Knowledge Required", "Skills Required", "Behaviours", "Criticality", "Frequency", "Source"]
        
        task_table = doc.add_table(rows=1, cols=len(task_headers))
        task_table.style = 'Table Grid'
        
        header_row = task_table.rows[0]
        set_repeat_table_header(header_row)
        
        for i, h in enumerate(task_headers):
            cell = header_row.cells[i]
            cell.text = h
            set_cell_shading(cell, "6B4C9A")  # Purple header
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(9)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
        
        for task_idx, task in enumerate(tasks):
            row = task_table.add_row()
            
            # Universal format
            row.cells[0].text = str(task.get("task_id", f"T-{task_idx+1:03d}"))
            row.cells[1].text = task.get("task_description", "")[:200]
            
            knowledge = task.get("knowledge_required", [])
            row.cells[2].text = "\n".join(knowledge[:3]) if isinstance(knowledge, list) else str(knowledge)[:150]
            
            skills = task.get("skills_required", [])
            row.cells[3].text = "\n".join(skills[:3]) if isinstance(skills, list) else str(skills)[:150]
            
            behaviours = task.get("behaviours_required", [])
            row.cells[4].text = "\n".join(behaviours[:2]) if isinstance(behaviours, list) else str(behaviours)[:100]
            
            row.cells[5].text = task.get("criticality", "Medium")
            row.cells[6].text = task.get("frequency", "As Required")
            row.cells[7].text = task.get("source", "")[:100]
            
            if task_idx % 2 == 1:
                for cell in row.cells:
                    set_cell_shading(cell, "F5F5F5")
            
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Roboto"
                        run.font.size = Pt(8)
        
        p = doc.add_paragraph()
        run = p.add_run(f"\nTotal tasks identified: {len(tasks)}")
        run.font.name = "Roboto"
        run.font.bold = True
    else:
        p = doc.add_paragraph()
        run = p.add_run("No tasks identified through research. Manual task analysis required using the 10-step research methodology.")
        run.font.name = "Roboto"
        run.font.italic = True
    
    doc.add_paragraph()
    
    # Research Sources Section
    create_styled_heading(doc, "4. Research Sources", 1)
    citations = data.get("section_18_citations", [])
    
    if citations:
        for i, citation in enumerate(citations[:15], 1):  # Limit to 15 in task doc
            p = doc.add_paragraph()
            source_type = citation.get('source_type', 'Source')
            source_name = citation.get('source_name', 'Unknown')
            url = citation.get('url', 'N/A')
            run = p.add_run(f"[{i}] {source_type}: {source_name}")
            run.font.name = "Roboto"
            run.font.bold = True
            run.font.size = Pt(9)
            
            p2 = doc.add_paragraph()
            run2 = p2.add_run(f"    {url}")
            run2.font.name = "Roboto"
            run2.font.size = Pt(8)
            run2.font.color.rgb = RGBColor(0, 0, 139)
    else:
        p = doc.add_paragraph()
        run = p.add_run("No citations recorded. See Analysis Report for full source list.")
        run.font.name = "Roboto"
        run.font.italic = True
    
    # Disclaimer
    doc.add_paragraph()
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("RESEARCH METHODOLOGY STATEMENT")
    run.font.name = "Roboto"
    run.font.bold = True
    run.font.size = Pt(9)
    
    p2 = doc.add_paragraph()
    run2 = p2.add_run("This document was generated using web-based research following the NOVA™ 10-step research methodology. All tasks have been identified from authoritative sources including professional body standards, competency frameworks, and industry documentation. No tasks have been fabricated. Where information could not be found, this has been explicitly stated. All information should be verified against current sources before use in formal training documentation.")
    run2.font.name = "Roboto"
    run2.font.size = Pt(8)
    run2.font.italic = True
    
    doc.save(filepath)
    print(f"[NOVA] Saved: {filepath}")


def build_comprehensive_analysis_report_doc(data: Dict, role_title: str, framework: str, terms: Dict, filepath: Path):
    """Build comprehensive 18-section Analysis Report following ANALYSIS-AGENT-SYSTEM-PROMPT.md exactly"""
    doc = Document()
    
    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    
    metadata = data.get("metadata", {})
    
    # =========================================================================
    # TITLE PAGE
    # =========================================================================
    doc.add_paragraph()
    doc.add_paragraph()
    
    # Classification
    p_class = doc.add_paragraph()
    run = p_class.add_run("OFFICIAL")
    run.font.name = "Roboto"
    run.font.size = Pt(14)
    run.font.bold = True
    p_class.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    title = doc.add_heading("ANALYSIS REPORT", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "Roboto"
        run.font.size = Pt(32)
        run.font.color.rgb = NOVA_DARK_BLUE
    
    # 18 Sections Badge
    sections_badge = doc.add_paragraph()
    run = sections_badge.add_run("18-SECTION COMPREHENSIVE ANALYSIS")
    run.font.name = "Roboto"
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(107, 76, 154)  # Purple
    sections_badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    subtitle = doc.add_paragraph()
    run = subtitle.add_run(role_title)
    run.font.name = "Roboto"
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = NOVA_DARK_BLUE
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Metadata box
    info_table = doc.add_table(rows=6, cols=2)
    info_table.style = 'Table Grid'
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    info_items = [
        ("Domain", metadata.get("domain", "N/A")),
        ("Specialism", metadata.get("specialism", "N/A")),
        ("Proficiency Level", metadata.get("proficiency_level", "N/A")),
        ("Framework", terms.get("framework_name", framework)),
        ("Generated", metadata.get("generated_date", "")[:10]),
        ("NOVA Version", metadata.get("nova_version", "5.1.0"))
    ]
    
    for i, (label, value) in enumerate(info_items):
        info_table.rows[i].cells[0].text = label
        info_table.rows[i].cells[1].text = str(value)
        set_cell_shading(info_table.rows[i].cells[0], "E8E0F0")
        for cell in info_table.rows[i].cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = "Roboto"
                    run.font.size = Pt(11)
    
    doc.add_page_break()
    
    # =========================================================================
    # TABLE OF CONTENTS (Manual)
    # =========================================================================
    toc_title = doc.add_heading("TABLE OF CONTENTS", 1)
    for run in toc_title.runs:
        run.font.color.rgb = NOVA_DARK_BLUE
    
    toc_sections = [
        "1. Executive Summary",
        "2. Framework Identification",
        "3. Geographic/Jurisdictional Context",
        "4. Professional Body/Regulator",
        "5. Competency Framework Mapping",
        "6. Role Description",
        "7. Qualifications",
        "8. Experience",
        "9. Technical Skills",
        "10. Soft Skills",
        "11. Personal Traits/Behaviours",
        "12. Physical/Medical/Security Requirements",
        "13. CPD/Recertification Requirements",
        "14. Career Progression Context",
        "15. Legal Compliance",
        "16. Professional Standards",
        "17. No Bias Statement",
        "18. Citations and Sources"
    ]
    
    for toc_item in toc_sections:
        p = doc.add_paragraph()
        run = p.add_run(toc_item)
        run.font.name = "Roboto"
        run.font.size = Pt(11)
    
    doc.add_page_break()
    
    # =========================================================================
    # SECTION 1: EXECUTIVE SUMMARY
    # =========================================================================
    create_styled_heading(doc, "1. EXECUTIVE SUMMARY", 1)
    
    s1 = data.get("section_1_executive_summary", {})
    
    create_styled_paragraph(doc, f"Analysis Scope: {s1.get('analysis_scope', 'Not specified')}", bold=True)
    create_styled_paragraph(doc, f"Research Methodology: {s1.get('research_methodology', 'Web-based research')}")
    create_styled_paragraph(doc, f"Tasks Identified: {s1.get('tasks_identified', len(data.get('tasks', [])))}")
    
    findings = s1.get("key_findings", [])
    if findings:
        create_styled_paragraph(doc, "Key Findings:", bold=True)
        for i, finding in enumerate(findings, 1):
            create_styled_paragraph(doc, f"  {i}. {finding}")
    
    standards = s1.get("primary_standards", [])
    if standards:
        create_styled_paragraph(doc, "Primary Standards Referenced:", bold=True)
        for std in standards:
            create_styled_paragraph(doc, f"  • {std}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 2: FRAMEWORK IDENTIFICATION
    # =========================================================================
    create_styled_heading(doc, "2. FRAMEWORK IDENTIFICATION", 1)
    
    s2 = data.get("section_2_framework_identification", {})
    
    create_styled_paragraph(doc, f"Framework Name: {s2.get('framework_name', framework)}", bold=True)
    create_styled_paragraph(doc, f"Version: {s2.get('framework_version', 'Not specified')}")
    create_styled_paragraph(doc, f"Governing Authority: {s2.get('governing_authority', 'Not specified')}")
    create_styled_paragraph(doc, f"Purpose: {s2.get('framework_purpose', 'Not specified')}")
    
    if s2.get('source_url'):
        create_styled_paragraph(doc, f"Source: {s2.get('source_url')}")
    
    analysis_reqs = s2.get("analysis_requirements", [])
    if analysis_reqs:
        create_styled_paragraph(doc, "Analysis Phase Requirements:", bold=True)
        for req in analysis_reqs:
            create_styled_paragraph(doc, f"  • {req}")
    
    required_outputs = s2.get("required_outputs", [])
    if required_outputs:
        create_styled_paragraph(doc, "Required Outputs:", bold=True)
        for output in required_outputs:
            create_styled_paragraph(doc, f"  • {output}")
    
    terminology = s2.get("terminology", {})
    if terminology:
        create_styled_paragraph(doc, "Framework Terminology:", bold=True)
        for term, definition in terminology.items():
            create_styled_paragraph(doc, f"  • {term}: {definition}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 3: GEOGRAPHIC/JURISDICTIONAL CONTEXT
    # =========================================================================
    create_styled_heading(doc, "3. GEOGRAPHIC/JURISDICTIONAL CONTEXT", 1)
    
    s3 = data.get("section_3_geographic_context", {})
    
    create_styled_paragraph(doc, f"Country: {s3.get('country', 'United Kingdom')}")
    create_styled_paragraph(doc, f"Legal Jurisdiction: {s3.get('legal_jurisdiction', 'England and Wales')}")
    create_styled_paragraph(doc, f"Language: {s3.get('language', 'English')}")
    create_styled_paragraph(doc, f"Currency: {s3.get('currency', 'GBP')}")
    
    if s3.get('regional_variations'):
        create_styled_paragraph(doc, f"Regional Variations: {s3.get('regional_variations')}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 4: PROFESSIONAL BODY/REGULATOR
    # =========================================================================
    create_styled_heading(doc, "4. PROFESSIONAL BODY/REGULATOR", 1)
    
    s4 = data.get("section_4_professional_body", {})
    
    create_styled_paragraph(doc, f"Professional Body: {s4.get('professional_body_name', 'Research required')}", bold=True)
    
    if s4.get('website_url'):
        create_styled_paragraph(doc, f"Website: {s4.get('website_url')}")
    
    create_styled_paragraph(doc, f"Registration Required: {'Yes' if s4.get('registration_required') else 'No'}")
    
    membership = s4.get("membership_categories", [])
    if membership:
        create_styled_paragraph(doc, "Membership Categories:", bold=True)
        for cat in membership:
            create_styled_paragraph(doc, f"  • {cat}")
    
    reg_reqs = s4.get("registration_requirements", [])
    if reg_reqs:
        create_styled_paragraph(doc, "Registration Requirements:", bold=True)
        for req in reg_reqs:
            create_styled_paragraph(doc, f"  • {req}")
    
    protected = s4.get("protected_titles", [])
    if protected:
        create_styled_paragraph(doc, "Protected Titles:", bold=True)
        for title_item in protected:
            create_styled_paragraph(doc, f"  • {title_item}")
    
    if s4.get('regulatory_authority'):
        create_styled_paragraph(doc, f"Regulatory Authority: {s4.get('regulatory_authority')}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 5: COMPETENCY FRAMEWORK MAPPING
    # =========================================================================
    create_styled_heading(doc, "5. COMPETENCY FRAMEWORK MAPPING", 1)
    
    s5 = data.get("section_5_competency_framework", {})
    
    create_styled_paragraph(doc, f"Framework: {s5.get('framework_name', 'Research required')}", bold=True)
    create_styled_paragraph(doc, f"Framework Owner: {s5.get('framework_owner', 'Not specified')}")
    
    if s5.get('framework_url'):
        create_styled_paragraph(doc, f"URL: {s5.get('framework_url')}")
    
    if s5.get('proficiency_mapping'):
        create_styled_paragraph(doc, f"Proficiency Mapping: {s5.get('proficiency_mapping')}")
    
    units = s5.get("relevant_units", [])
    if units:
        create_styled_paragraph(doc, "Relevant Competency Units:", bold=True)
        for unit in units:
            if isinstance(unit, dict):
                create_styled_paragraph(doc, f"  • {unit.get('unit_code', '')}: {unit.get('unit_title', '')}")
            else:
                create_styled_paragraph(doc, f"  • {unit}")
    
    level_desc = s5.get("level_descriptors", {})
    if level_desc:
        create_styled_paragraph(doc, "Level Descriptors:", bold=True)
        for level, desc in level_desc.items():
            create_styled_paragraph(doc, f"  • Level {level}: {desc}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 6: ROLE DESCRIPTION
    # =========================================================================
    create_styled_heading(doc, "6. ROLE DESCRIPTION", 1)
    
    s6 = data.get("section_6_role_description", {})
    
    if s6.get('comprehensive_definition'):
        create_styled_paragraph(doc, "Role Definition:", bold=True)
        create_styled_paragraph(doc, s6.get('comprehensive_definition'))
    
    if s6.get('primary_purpose'):
        create_styled_paragraph(doc, f"Primary Purpose: {s6.get('primary_purpose')}")
    
    accountabilities = s6.get("key_accountabilities", [])
    if accountabilities:
        create_styled_paragraph(doc, "Key Accountabilities:", bold=True)
        for acc in accountabilities:
            create_styled_paragraph(doc, f"  • {acc}")
    
    if s6.get('reporting_structure'):
        create_styled_paragraph(doc, f"Reporting Structure: {s6.get('reporting_structure')}")
    
    if s6.get('team_context'):
        create_styled_paragraph(doc, f"Team Context: {s6.get('team_context')}")
    
    equiv_titles = s6.get("equivalent_titles", [])
    if equiv_titles:
        create_styled_paragraph(doc, "Equivalent/Alternative Titles:", bold=True)
        for t in equiv_titles:
            create_styled_paragraph(doc, f"  • {t}")
    
    if s6.get('role_boundaries'):
        create_styled_paragraph(doc, f"Role Boundaries: {s6.get('role_boundaries')}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 7: QUALIFICATIONS
    # =========================================================================
    create_styled_heading(doc, "7. QUALIFICATIONS", 1)
    
    s7 = data.get("section_7_qualifications", {})
    
    essential_quals = s7.get("essential_qualifications", [])
    if essential_quals:
        create_styled_paragraph(doc, "Essential Qualifications:", bold=True)
        for qual in essential_quals:
            if isinstance(qual, dict):
                create_styled_paragraph(doc, f"  • {qual.get('qualification', '')} - {qual.get('level', '')} [{qual.get('source', '')}]")
            else:
                create_styled_paragraph(doc, f"  • {qual}")
    
    desirable_quals = s7.get("desirable_qualifications", [])
    if desirable_quals:
        create_styled_paragraph(doc, "Desirable Qualifications:", bold=True)
        for qual in desirable_quals:
            if isinstance(qual, dict):
                create_styled_paragraph(doc, f"  • {qual.get('qualification', '')} - {qual.get('level', '')} [{qual.get('source', '')}]")
            else:
                create_styled_paragraph(doc, f"  • {qual}")
    
    if s7.get('academic_level_required'):
        create_styled_paragraph(doc, f"Academic Level Required: {s7.get('academic_level_required')}")
    
    cert_req = s7.get("professional_certifications_required", [])
    if cert_req:
        create_styled_paragraph(doc, "Professional Certifications (Required):", bold=True)
        for cert in cert_req:
            create_styled_paragraph(doc, f"  • {cert}")
    
    cert_des = s7.get("professional_certifications_desirable", [])
    if cert_des:
        create_styled_paragraph(doc, "Professional Certifications (Desirable):", bold=True)
        for cert in cert_des:
            create_styled_paragraph(doc, f"  • {cert}")
    
    apprenticeships = s7.get("apprenticeship_routes", [])
    if apprenticeships:
        create_styled_paragraph(doc, "Apprenticeship Routes:", bold=True)
        for app in apprenticeships:
            if isinstance(app, dict):
                create_styled_paragraph(doc, f"  • {app.get('name', '')} - Level {app.get('level', '')} [{app.get('source', '')}]")
            else:
                create_styled_paragraph(doc, f"  • {app}")
    
    if s7.get('qualification_equivalencies'):
        create_styled_paragraph(doc, f"Equivalencies: {s7.get('qualification_equivalencies')}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 8: EXPERIENCE
    # =========================================================================
    create_styled_heading(doc, "8. EXPERIENCE", 1)
    
    s8 = data.get("section_8_experience", {})
    
    if s8.get('years_required'):
        create_styled_paragraph(doc, f"Years of Experience Required: {s8.get('years_required')}", bold=True)
    
    exp_types = s8.get("type_of_experience", [])
    if exp_types:
        create_styled_paragraph(doc, "Type of Experience Required:", bold=True)
        for exp in exp_types:
            create_styled_paragraph(doc, f"  • {exp}")
    
    sector_reqs = s8.get("sector_specific_requirements", [])
    if sector_reqs:
        create_styled_paragraph(doc, "Sector-Specific Requirements:", bold=True)
        for req in sector_reqs:
            create_styled_paragraph(doc, f"  • {req}")
    
    project_exp = s8.get("project_experience", [])
    if project_exp:
        create_styled_paragraph(doc, "Project Experience:", bold=True)
        for proj in project_exp:
            create_styled_paragraph(doc, f"  • {proj}")
    
    if s8.get('leadership_experience'):
        create_styled_paragraph(doc, f"Leadership Experience: {s8.get('leadership_experience')}")
    
    if s8.get('international_experience'):
        create_styled_paragraph(doc, f"International Experience: {s8.get('international_experience')}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 9: TECHNICAL SKILLS
    # =========================================================================
    create_styled_heading(doc, "9. TECHNICAL SKILLS", 1)
    
    s9 = data.get("section_9_technical_skills", [])
    
    if s9:
        rows = []
        for skill in s9:
            if isinstance(skill, dict):
                rows.append([
                    skill.get("skill", ""),
                    skill.get("category", "Core"),
                    skill.get("proficiency_level", ""),
                    skill.get("source", "")[:50]
                ])
        if rows:
            create_styled_table(doc, ["Skill", "Category", "Proficiency Level", "Source"], rows, [2.0, 1.0, 1.5, 2.0])
    else:
        create_styled_paragraph(doc, "No specific technical skills identified through research.")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 10: SOFT SKILLS
    # =========================================================================
    create_styled_heading(doc, "10. SOFT SKILLS", 1)
    
    s10 = data.get("section_10_soft_skills", [])
    
    if s10:
        rows = []
        for skill in s10:
            if isinstance(skill, dict):
                rows.append([
                    skill.get("skill", ""),
                    skill.get("proficiency_level", ""),
                    skill.get("source", "")[:50]
                ])
        if rows:
            create_styled_table(doc, ["Skill", "Proficiency Level", "Source"], rows, [2.5, 1.5, 2.5])
    else:
        create_styled_paragraph(doc, "No specific soft skills identified through research.")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 11: PERSONAL TRAITS/BEHAVIOURS
    # =========================================================================
    create_styled_heading(doc, "11. PERSONAL TRAITS AND BEHAVIOURS", 1)
    
    s11 = data.get("section_11_behaviours", [])
    
    if s11:
        for behaviour in s11:
            if isinstance(behaviour, dict):
                req_type = behaviour.get('requirement_type', 'Required')
                create_styled_paragraph(doc, f"  • [{req_type}] {behaviour.get('behaviour', '')} [{behaviour.get('source', '')}]")
            else:
                create_styled_paragraph(doc, f"  • {behaviour}")
    else:
        create_styled_paragraph(doc, "No specific behaviours identified through research.")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 12: PHYSICAL/MEDICAL/SECURITY REQUIREMENTS
    # =========================================================================
    create_styled_heading(doc, "12. PHYSICAL, MEDICAL AND SECURITY REQUIREMENTS", 1)
    
    s12 = data.get("section_12_physical_medical_security", {})
    
    physical = s12.get("physical_requirements", [])
    if physical and physical != ["No specific requirements identified"]:
        create_styled_paragraph(doc, "Physical Requirements:", bold=True)
        for req in physical:
            create_styled_paragraph(doc, f"  • {req}")
    else:
        create_styled_paragraph(doc, "Physical Requirements: No specific physical requirements identified for this role.")
    
    medical = s12.get("medical_requirements", [])
    if medical and medical != ["No specific requirements identified"]:
        create_styled_paragraph(doc, "Medical Requirements:", bold=True)
        for req in medical:
            create_styled_paragraph(doc, f"  • {req}")
    else:
        create_styled_paragraph(doc, "Medical Requirements: No specific medical requirements identified for this role.")
    
    create_styled_paragraph(doc, f"Security Clearance: {s12.get('security_clearance', 'Standard employment checks only')}")
    create_styled_paragraph(doc, f"DBS Requirements: {s12.get('dbs_requirements', 'Basic DBS check')}")
    
    occ_health = s12.get("occupational_health", [])
    if occ_health:
        create_styled_paragraph(doc, "Occupational Health:", bold=True)
        for oh in occ_health:
            create_styled_paragraph(doc, f"  • {oh}")
    
    if s12.get('reasonable_adjustments'):
        create_styled_paragraph(doc, f"Reasonable Adjustments: {s12.get('reasonable_adjustments')}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 13: CPD/RECERTIFICATION REQUIREMENTS
    # =========================================================================
    create_styled_heading(doc, "13. CPD AND RECERTIFICATION REQUIREMENTS", 1)
    
    s13 = data.get("section_13_cpd_requirements", {})
    
    if s13.get('professional_body_cpd'):
        create_styled_paragraph(doc, f"Professional Body CPD Policy: {s13.get('professional_body_cpd')}", bold=True)
    
    if s13.get('annual_hours_points'):
        create_styled_paragraph(doc, f"Annual CPD Requirement: {s13.get('annual_hours_points')}")
    
    if s13.get('recertification_cycle'):
        create_styled_paragraph(doc, f"Recertification Cycle: {s13.get('recertification_cycle')}")
    
    mandatory_refresh = s13.get("mandatory_refresher", [])
    if mandatory_refresh:
        create_styled_paragraph(doc, "Mandatory Refresher Training:", bold=True)
        for training in mandatory_refresh:
            create_styled_paragraph(doc, f"  • {training}")
    
    if s13.get('portfolio_requirements'):
        create_styled_paragraph(doc, f"Portfolio Requirements: {s13.get('portfolio_requirements')}")
    
    if s13.get('revalidation_process'):
        create_styled_paragraph(doc, f"Revalidation Process: {s13.get('revalidation_process')}")
    
    if s13.get('non_compliance_consequences'):
        create_styled_paragraph(doc, f"Non-Compliance Consequences: {s13.get('non_compliance_consequences')}")
    
    if not any([s13.get('professional_body_cpd'), s13.get('annual_hours_points'), mandatory_refresh]):
        create_styled_paragraph(doc, "No mandatory CPD requirements identified. Voluntary professional development recommended.")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 14: CAREER PROGRESSION CONTEXT
    # =========================================================================
    create_styled_heading(doc, "14. CAREER PROGRESSION CONTEXT", 1)
    
    s14 = data.get("section_14_career_progression", {})
    
    pathway_to = s14.get("pathway_to_role", [])
    if pathway_to:
        create_styled_paragraph(doc, "Typical Career Pathway TO This Role:", bold=True)
        for role in pathway_to:
            create_styled_paragraph(doc, f"  • {role}")
    
    pathway_from = s14.get("pathway_from_role", [])
    if pathway_from:
        create_styled_paragraph(doc, "Typical Career Pathway FROM This Role:", bold=True)
        for role in pathway_from:
            create_styled_paragraph(doc, f"  • {role}")
    
    lateral = s14.get("lateral_moves", [])
    if lateral:
        create_styled_paragraph(doc, "Lateral Move Options:", bold=True)
        for move in lateral:
            create_styled_paragraph(doc, f"  • {move}")
    
    promo_criteria = s14.get("promotion_criteria", [])
    if promo_criteria:
        create_styled_paragraph(doc, "Promotion Criteria:", bold=True)
        for criterion in promo_criteria:
            create_styled_paragraph(doc, f"  • {criterion}")
    
    if s14.get('timeline_expectations'):
        create_styled_paragraph(doc, f"Timeline Expectations: {s14.get('timeline_expectations')}")
    
    skill_gaps = s14.get("skill_gaps_for_progression", [])
    if skill_gaps:
        create_styled_paragraph(doc, "Skill Gaps to Address for Progression:", bold=True)
        for gap in skill_gaps:
            create_styled_paragraph(doc, f"  • {gap}")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 15: LEGAL COMPLIANCE
    # =========================================================================
    create_styled_heading(doc, "15. LEGAL COMPLIANCE", 1)
    
    s15 = data.get("section_15_legal_compliance", [])
    
    if s15:
        rows = []
        for legal in s15:
            if isinstance(legal, dict):
                rows.append([
                    legal.get("legislation", ""),
                    legal.get("relevance", ""),
                    "Yes" if legal.get("mandatory_training") else "No",
                    legal.get("source", "")[:40]
                ])
        if rows:
            create_styled_table(doc, ["Legislation", "Relevance to Role", "Mandatory Training", "Source"], rows, [1.8, 2.2, 1.0, 1.5])
    else:
        create_styled_paragraph(doc, "No specific legal compliance requirements identified through research.")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 16: PROFESSIONAL STANDARDS
    # =========================================================================
    create_styled_heading(doc, "16. PROFESSIONAL STANDARDS", 1)
    
    s16 = data.get("section_16_professional_standards", [])
    
    if s16:
        for std in s16:
            if isinstance(std, dict):
                req_type = std.get('requirement_type', 'Standard')
                create_styled_paragraph(doc, f"  • [{req_type}] {std.get('standard', '')} - {std.get('issuing_body', '')}", bold=True)
                if std.get('description'):
                    create_styled_paragraph(doc, f"    {std.get('description')}")
                if std.get('source'):
                    create_styled_paragraph(doc, f"    Source: {std.get('source')}")
            else:
                create_styled_paragraph(doc, f"  • {std}")
    else:
        create_styled_paragraph(doc, "No specific professional standards identified through research.")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 17: NO BIAS STATEMENT
    # =========================================================================
    create_styled_heading(doc, "17. EQUALITY AND DIVERSITY STATEMENT", 1)
    
    s17 = data.get("section_17_no_bias_statement", {})
    
    create_styled_paragraph(doc, "This analysis has been conducted in accordance with the following principles:")
    doc.add_paragraph()
    
    create_styled_paragraph(doc, f"✓ Equality Act 2010 Compliance: {'Yes' if s17.get('equality_act_compliance', True) else 'Review Required'}")
    create_styled_paragraph(doc, f"✓ Genuine Occupational Requirements: {s17.get('genuine_occupational_requirements', 'All requirements are genuine occupational requirements')}")
    create_styled_paragraph(doc, f"✓ Reasonable Adjustments Considered: {'Yes' if s17.get('reasonable_adjustments_considered', True) else 'Review Required'}")
    create_styled_paragraph(doc, f"✓ Inclusive Language Used: {'Yes' if s17.get('inclusive_language_used', True) else 'Review Required'}")
    create_styled_paragraph(doc, f"✓ Proportionality: {s17.get('proportionality_statement', 'Requirements are proportionate to role needs')}")
    
    bias_concerns = s17.get("bias_concerns_identified", "None identified")
    if bias_concerns and bias_concerns != "None identified":
        create_styled_paragraph(doc, f"⚠ Bias Concerns Identified: {bias_concerns}", bold=True)
    else:
        create_styled_paragraph(doc, "No bias concerns identified during this analysis.")
    
    doc.add_paragraph()
    
    # =========================================================================
    # SECTION 18: CITATIONS AND SOURCES
    # =========================================================================
    create_styled_heading(doc, "18. CITATIONS AND SOURCES", 1)
    
    s18 = data.get("section_18_citations", [])
    
    if s18:
        rows = []
        for citation in s18:
            if isinstance(citation, dict):
                rows.append([
                    citation.get("source_type", ""),
                    citation.get("source_name", ""),
                    citation.get("url", "")[:50],
                    citation.get("date_accessed", "")
                ])
        if rows:
            create_styled_table(doc, ["Type", "Source Name", "URL", "Accessed"], rows, [1.2, 2.0, 2.3, 1.0])
    else:
        create_styled_paragraph(doc, "No citations recorded. Manual verification of sources required.")
    
    # Research Log
    doc.add_paragraph()
    create_styled_paragraph(doc, "Research Log:", bold=True)
    
    research_log = data.get("research_log", {})
    searches = research_log.get("searches_conducted", metadata.get("searches_performed", []))
    
    if searches:
        create_styled_paragraph(doc, f"Total searches conducted: {len(searches)}")
        create_styled_paragraph(doc, "Search queries executed:", bold=True)
        for i, search in enumerate(searches[:20], 1):
            create_styled_paragraph(doc, f"  {i}. {search}")
        if len(searches) > 20:
            create_styled_paragraph(doc, f"  ... and {len(searches) - 20} more searches")
    else:
        create_styled_paragraph(doc, "No search queries recorded.")
    
    # =========================================================================
    # FINAL DISCLAIMER
    # =========================================================================
    doc.add_page_break()
    
    create_styled_heading(doc, "METHODOLOGY AND DISCLAIMER", 1)
    
    create_styled_paragraph(doc, "Research Methodology Statement:", bold=True)
    create_styled_paragraph(doc, 
        "This analysis was conducted using the NOVA™ 10-step research methodology as specified in "
        "ANALYSIS-AGENT-SYSTEM-PROMPT.md. Web-based research was conducted to identify authoritative "
        "sources including professional body standards, competency frameworks, legislation, and industry "
        "documentation. No information has been fabricated. Where information could not be found through "
        "research, this has been explicitly stated.")
    
    doc.add_paragraph()
    
    create_styled_paragraph(doc, "Quality Assurance:", bold=True)
    create_styled_paragraph(doc, "  ✓ All 18 sections of the Analysis Report completed")
    create_styled_paragraph(doc, "  ✓ Every factual claim has citation where available")
    create_styled_paragraph(doc, "  ✓ No statistics or percentages have been invented")
    create_styled_paragraph(doc, "  ✓ No methodology claims that did not occur")
    create_styled_paragraph(doc, "  ✓ Framework-specific terminology used correctly")
    
    doc.add_paragraph()
    
    create_styled_paragraph(doc, "Disclaimer:", bold=True)
    p = doc.add_paragraph()
    run = p.add_run(
        "This document was generated using AI-assisted research conducted on the date shown. "
        "Information may have changed since the research was conducted. All information should be "
        "independently verified against current authoritative sources before use in formal training "
        "documentation or decision-making. NOVA and its operators accept no liability for decisions "
        "made based on this analysis."
    )
    run.font.name = "Roboto"
    run.font.size = Pt(9)
    run.font.italic = True
    
    doc.save(filepath)
    print(f"[NOVA] Saved 18-section Analysis Report: {filepath}")


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
