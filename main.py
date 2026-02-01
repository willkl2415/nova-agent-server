"""
NOVA Agent Server v7.0 - TRUE MULTI-AGENT ARCHITECTURE
=======================================================

This version implements genuine autonomous agents that collaborate
to produce comprehensive role analysis. Each agent has a specific
mission and passes structured data to the next.

AGENT ARCHITECTURE:
1. Role Intelligence Agent - Market context, role definition, employers
2. Skills Architect Agent - Technical skills, soft skills, behaviours, frameworks
3. Compliance Agent - Legal, regulatory, professional body requirements
4. Quality Validator Agent - Completeness check, gap identification
5. Document Builder - Single consolidated report

OUTPUT: [Role Title] Skills Report.docx

Author: Claude AI for NOVA Project
Date: 1 February 2026
Version: 7.0.0
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
from fastapi.responses import FileResponse
from pydantic import BaseModel
import zipfile
import anthropic

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================================
# APP SETUP
# ============================================================================

app = FastAPI(title="NOVA Agent Server", version="7.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Job Storage
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
CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-5-20250929")
claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY) if ANTHROPIC_API_KEY else None

print(f"[NOVA v7.0] Multi-Agent Architecture Started. Claude: {claude_client is not None}")


# ============================================================================
# API MODELS
# ============================================================================

class ExecuteRequest(BaseModel):
    job_id: Optional[str] = None
    agent: Optional[str] = None
    agent_type: Optional[str] = None
    parameters: Dict[str, Any] = {}
    framework: str = "INDUSTRY_STANDARDS"


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
# FRAMEWORK TERMINOLOGY
# ============================================================================

def normalize_framework(framework: str) -> str:
    """Normalize framework name to standard key"""
    if not framework:
        return "INDUSTRY_STANDARDS"
    
    framework_lower = framework.lower().replace(" ", "_").replace("-", "_")
    
    mapping = {
        "uk_dsat": "UK_DSAT", "uk": "UK_DSAT", "dsat": "UK_DSAT", "jsp822": "UK_DSAT",
        "us_tradoc": "US_TRADOC", "us": "US_TRADOC", "tradoc": "US_TRADOC",
        "nato_bisc": "NATO_BISC", "nato": "NATO_BISC", "bisc": "NATO_BISC",
        "australian_sadl": "AUSTRALIAN_SADL", "sadl": "AUSTRALIAN_SADL",
        "s6000t": "S6000T", "asd_s6000t": "S6000T",
        "industry_standards": "INDUSTRY_STANDARDS", "industry": "INDUSTRY_STANDARDS",
        "commercial": "INDUSTRY_STANDARDS", "addie": "INDUSTRY_STANDARDS"
    }
    
    return mapping.get(framework_lower, "INDUSTRY_STANDARDS")


def get_terminology(framework: str) -> Dict[str, str]:
    """Get framework-specific terminology"""
    framework = normalize_framework(framework)
    
    TERMINOLOGY = {
        "UK_DSAT": {
            "framework_name": "UK Defence Systems Approach to Training (DSAT)",
            "report_type": "Role Performance Statement",
            "citation_prefix": "[JSP 822 V7.0]",
            "authority": "UK Ministry of Defence"
        },
        "US_TRADOC": {
            "framework_name": "US Army TRADOC",
            "report_type": "Individual Critical Task List",
            "citation_prefix": "[TRADOC Reg 350-70]",
            "authority": "US Army TRADOC"
        },
        "NATO_BISC": {
            "framework_name": "NATO Bi-SC Directive 075-007",
            "report_type": "Skills, Tasks, Proficiency Analysis",
            "citation_prefix": "[Bi-SCD 075-007]",
            "authority": "NATO Allied Command Transformation"
        },
        "INDUSTRY_STANDARDS": {
            "framework_name": "Industry Standards",
            "report_type": "Skills Report",
            "citation_prefix": "[Industry Standard]",
            "authority": "Professional Bodies"
        }
    }
    
    return TERMINOLOGY.get(framework, TERMINOLOGY["INDUSTRY_STANDARDS"])


def is_defence_framework(framework: str) -> bool:
    """Check if framework is defence-related"""
    return normalize_framework(framework) in ["UK_DSAT", "US_TRADOC", "NATO_BISC", "S6000T"]


# ============================================================================
# API ENDPOINTS
# ============================================================================

@app.get("/api/health")
async def health_check():
    return {
        "status": "healthy",
        "version": "7.0.0",
        "architecture": "multi-agent",
        "claude_configured": claude_client is not None,
        "timestamp": datetime.now().isoformat()
    }


@app.post("/api/execute", response_model=TaskResponse)
async def execute_agent(request: ExecuteRequest, background_tasks: BackgroundTasks):
    """Start an agent task"""
    try:
        if not claude_client:
            raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not configured")
        
        agent_type = request.agent_type or request.agent or "analysis"
        job_id = request.job_id or f"nova-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{os.urandom(4).hex()}"
        
        job_output_dir = OUTPUT_DIR / job_id
        job_output_dir.mkdir(parents=True, exist_ok=True)
        
        framework = normalize_framework(request.parameters.get("framework") or request.framework)
        
        jobs.set(job_id, {
            "job_id": job_id,
            "status": "running",
            "progress": 0,
            "message": "Initializing multi-agent system...",
            "current_step": "Initializing...",
            "steps_completed": [],
            "error": None,
            "agent_type": agent_type,
            "framework": framework,
            "parameters": request.parameters,
            "output_dir": str(job_output_dir),
            "created_at": datetime.now().isoformat(),
            "completed_at": None
        })
        
        print(f"[NOVA v7.0] Job {job_id}: {agent_type} with {framework}")
        
        background_tasks.add_task(run_agent_task, job_id, agent_type, request.parameters, framework)
        
        return TaskResponse(job_id=job_id, status="started", message=f"Multi-agent {agent_type} started")
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"[NOVA] Execute error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/status/{job_id}", response_model=StatusResponse)
async def get_status(job_id: str):
    """Get job status"""
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    
    return StatusResponse(
        job_id=job_id,
        status=job.get("status", "unknown"),
        progress=job.get("progress", 0),
        current_step=job.get("message", "Processing..."),
        steps_completed=job.get("steps_completed", []),
        error=job.get("error"),
        created_at=job.get("created_at", datetime.now().isoformat()),
        completed_at=job.get("completed_at")
    )


@app.get("/api/download/{job_id}")
async def download_results(job_id: str):
    """Download job results as ZIP"""
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    
    if job.get("status") != "completed":
        raise HTTPException(status_code=400, detail="Job not complete")
    
    output_dir = Path(job["output_dir"])
    zip_path = output_dir.parent / f"{job_id}.zip"
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in output_dir.rglob("*"):
            if file_path.is_file():
                arcname = file_path.relative_to(output_dir)
                zf.write(file_path, arcname)
    
    return FileResponse(path=str(zip_path), filename=f"NOVA_{job_id}.zip", media_type="application/zip")


def update_job(job_id: str, progress: int, message: str, status: str = None):
    """Update job progress"""
    update_data = {"progress": progress, "message": message, "current_step": message}
    if status:
        update_data["status"] = status
        if status == "completed":
            update_data["completed_at"] = datetime.now().isoformat()
    
    job = jobs.get(job_id)
    if job:
        steps = job.get("steps_completed", [])
        if message and "Agent" in message:
            steps.append(f"{progress}%: {message}")
        update_data["steps_completed"] = steps
    
    jobs.update(job_id, **update_data)
    print(f"[NOVA] {job_id}: {progress}% - {message}")


# ============================================================================
# CLAUDE API CALL
# ============================================================================

async def call_agent(agent_name: str, system_prompt: str, user_prompt: str, max_tokens: int = 8000) -> str:
    """Call Claude for a specific agent task with timeout"""
    if not claude_client:
        raise Exception("Claude API not configured")
    
    print(f"[NOVA] {agent_name}: Calling Claude ({len(user_prompt)} chars)...")
    
    loop = asyncio.get_event_loop()
    
    try:
        response = await asyncio.wait_for(
            loop.run_in_executor(
                None,
                lambda: claude_client.messages.create(
                    model=CLAUDE_MODEL,
                    max_tokens=max_tokens,
                    system=system_prompt,
                    messages=[{"role": "user", "content": user_prompt}]
                )
            ),
            timeout=90.0
        )
        
        result = response.content[0].text if response.content else ""
        print(f"[NOVA] {agent_name}: Response received ({len(result)} chars)")
        return result
        
    except asyncio.TimeoutError:
        raise Exception(f"{agent_name} timeout after 90 seconds")
    except Exception as e:
        raise Exception(f"{agent_name} error: {str(e)}")


def parse_json_response(text: str) -> Dict:
    """Parse JSON from Claude response"""
    if not text:
        return {}
    
    text = text.strip()
    
    # Try direct parse
    try:
        if text.startswith('{'):
            return json.loads(text)
    except:
        pass
    
    # Try to extract JSON from markdown
    json_match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', text)
    if json_match:
        try:
            return json.loads(json_match.group(1))
        except:
            pass
    
    # Try to find JSON object
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end != -1:
        try:
            return json.loads(text[start:end+1])
        except:
            pass
    
    return {"raw_response": text}


# ============================================================================
# AGENT TASK ROUTER
# ============================================================================

async def run_agent_task(job_id: str, agent_type: str, parameters: Dict, framework: str):
    """Route to appropriate agent"""
    try:
        terms = get_terminology(framework)
        
        if agent_type == "analysis":
            await run_multi_agent_analysis(job_id, parameters, framework, terms)
        elif agent_type == "design":
            await run_design_agent(job_id, parameters, framework, terms)
        elif agent_type == "delivery":
            await run_delivery_agent(job_id, parameters, framework, terms)
        elif agent_type == "evaluation":
            await run_evaluation_agent(job_id, parameters, framework, terms)
        else:
            raise ValueError(f"Unknown agent type: {agent_type}")
        
        jobs.update(job_id, status="completed", progress=100, message="Complete", completed_at=datetime.now().isoformat())
        
    except Exception as e:
        print(f"[NOVA] Error: {e}")
        import traceback
        traceback.print_exc()
        jobs.update(job_id, status="failed", message=str(e), error=str(e))


# ============================================================================
# MULTI-AGENT ANALYSIS SYSTEM
# ============================================================================

async def run_multi_agent_analysis(job_id: str, parameters: Dict, framework: str, terms: Dict):
    """
    MULTI-AGENT ANALYSIS ORCHESTRATOR
    
    Coordinates 4 specialized agents:
    1. Role Intelligence Agent - Role definition, market context
    2. Skills Architect Agent - Skills mapping, competency frameworks
    3. Compliance Agent - Legal, regulatory, professional requirements
    4. Quality Validator Agent - Validation and gap analysis
    
    Then builds single consolidated document.
    """
    
    # Extract parameters
    domain = parameters.get("domain", "Technology")
    specialism = parameters.get("specialism", domain)
    role_title = parameters.get("role_title", "Specialist")
    proficiency_level = parameters.get("proficiency_level", "Mid-Level")
    role_description = parameters.get("role_description", "")
    
    output_dir = Path(jobs.get(job_id)["output_dir"])
    output_dir.mkdir(exist_ok=True)
    
    print(f"[NOVA v7.0] Multi-Agent Analysis: {role_title}")
    print(f"[NOVA v7.0] Domain: {domain}, Specialism: {specialism}, Level: {proficiency_level}")
    
    # Initialize combined data
    combined_data = {
        "metadata": {
            "role_title": role_title,
            "domain": domain,
            "specialism": specialism,
            "proficiency_level": proficiency_level,
            "framework": framework,
            "generated_date": datetime.now().isoformat(),
            "nova_version": "7.0.0"
        }
    }
    
    # ========================================================================
    # AGENT 1: ROLE INTELLIGENCE AGENT
    # ========================================================================
    update_job(job_id, 5, "Agent 1: Role Intelligence Agent starting...")
    
    role_intel = await run_role_intelligence_agent(
        role_title=role_title,
        domain=domain,
        specialism=specialism,
        proficiency_level=proficiency_level,
        role_description=role_description
    )
    
    combined_data["role_intelligence"] = role_intel
    update_job(job_id, 25, "Agent 1: Role Intelligence Agent complete")
    
    # ========================================================================
    # AGENT 2: SKILLS ARCHITECT AGENT
    # ========================================================================
    update_job(job_id, 30, "Agent 2: Skills Architect Agent starting...")
    
    skills_data = await run_skills_architect_agent(
        role_title=role_title,
        domain=domain,
        specialism=specialism,
        proficiency_level=proficiency_level,
        role_intel=role_intel
    )
    
    combined_data["skills_architecture"] = skills_data
    update_job(job_id, 50, "Agent 2: Skills Architect Agent complete")
    
    # ========================================================================
    # AGENT 3: COMPLIANCE AGENT
    # ========================================================================
    update_job(job_id, 55, "Agent 3: Compliance Agent starting...")
    
    compliance_data = await run_compliance_agent(
        role_title=role_title,
        domain=domain,
        specialism=specialism
    )
    
    combined_data["compliance"] = compliance_data
    update_job(job_id, 75, "Agent 3: Compliance Agent complete")
    
    # ========================================================================
    # AGENT 4: QUALITY VALIDATOR AGENT
    # ========================================================================
    update_job(job_id, 80, "Agent 4: Quality Validator Agent starting...")
    
    validation = await run_quality_validator_agent(combined_data)
    
    combined_data["validation"] = validation
    update_job(job_id, 90, "Agent 4: Quality Validator Agent complete")
    
    # ========================================================================
    # BUILD DOCUMENT
    # ========================================================================
    update_job(job_id, 92, "Building Skills Report document...")
    
    # Sanitize role title for filename
    safe_title = re.sub(r'[^\w\s-]', '', role_title).strip().replace(' ', '_')
    doc_filename = f"{safe_title}_Skills_Report.docx"
    
    build_skills_report(combined_data, output_dir / doc_filename)
    
    # Save JSON
    with open(output_dir / "analysis_data.json", "w") as f:
        json.dump(combined_data, f, indent=2, default=str)
    
    update_job(job_id, 100, "Skills Report complete")


# ============================================================================
# AGENT 1: ROLE INTELLIGENCE AGENT
# ============================================================================

async def run_role_intelligence_agent(
    role_title: str,
    domain: str,
    specialism: str,
    proficiency_level: str,
    role_description: str
) -> Dict:
    """
    ROLE INTELLIGENCE AGENT
    
    Mission: Understand the role in its market context.
    
    Outputs:
    - Role identification (title, family, level, employment types)
    - Role purpose and business outcomes
    - Typical employers and sectors
    - Reporting structures
    - Related/equivalent job titles
    """
    
    system_prompt = """You are a Role Intelligence Agent specializing in job market analysis.

Your mission is to provide comprehensive role intelligence for workforce planning.

You have deep knowledge of:
- UK and international job markets
- Role hierarchies and career structures
- Industry sectors and employer types
- Job family classifications

Respond ONLY with valid JSON. No explanatory text."""
    
    user_prompt = f"""Analyse this role and provide comprehensive role intelligence:

ROLE: {role_title}
DOMAIN: {domain}
SPECIALISM: {specialism}
LEVEL: {proficiency_level}
CONTEXT: {role_description or 'Standard role in this field'}

Provide a JSON object with this structure:

{{
    "role_identification": {{
        "role_title": "{role_title}",
        "role_family": "The job family this belongs to (e.g., Technology / Engineering / Data)",
        "job_level": "The level (e.g., Senior Individual Contributor, Team Lead, Manager)",
        "employment_types": ["Permanent", "Contract", "Other relevant types"],
        "work_arrangements": ["Office", "Hybrid", "Remote"],
        "typical_team_size": "If they manage people, typical team size"
    }},
    
    "role_purpose": {{
        "summary": "A clear 1-2 sentence description of what this role does and why it exists",
        "business_outcomes": [
            "Key business outcome 1",
            "Key business outcome 2",
            "Key business outcome 3",
            "Key business outcome 4"
        ]
    }},
    
    "organisational_context": {{
        "typical_employers": [
            "Type of employer 1",
            "Type of employer 2",
            "Type of employer 3"
        ],
        "sectors": ["Sector 1", "Sector 2", "Sector 3"],
        "reporting_lines": [
            "Typical manager title 1",
            "Typical manager title 2"
        ],
        "direct_reports": ["Role that might report to this position"],
        "key_stakeholders": ["Internal or external stakeholders they work with"]
    }},
    
    "role_scope": {{
        "primary_focus_areas": [
            "Main area of responsibility 1",
            "Main area of responsibility 2",
            "Main area of responsibility 3"
        ],
        "decision_authority": "Description of what decisions they can make",
        "budget_responsibility": "Typical budget scope if any",
        "geographic_scope": "Local, Regional, National, or International"
    }},
    
    "equivalent_titles": [
        "Alternative job title 1",
        "Alternative job title 2",
        "Alternative job title 3"
    ],
    
    "career_context": {{
        "typical_entry_routes": [
            "Previous role 1",
            "Previous role 2"
        ],
        "progression_paths": {{
            "technical": ["Next technical role 1", "Next technical role 2"],
            "leadership": ["Next management role 1", "Next management role 2"]
        }},
        "typical_tenure": "How long people typically stay in this role"
    }}
}}

Be specific to {domain} and {specialism}. Use real job titles and realistic information."""

    response = await call_agent("Role Intelligence Agent", system_prompt, user_prompt)
    return parse_json_response(response)


# ============================================================================
# AGENT 2: SKILLS ARCHITECT AGENT
# ============================================================================

async def run_skills_architect_agent(
    role_title: str,
    domain: str,
    specialism: str,
    proficiency_level: str,
    role_intel: Dict
) -> Dict:
    """
    SKILLS ARCHITECT AGENT
    
    Mission: Map comprehensive skill requirements to frameworks.
    
    Outputs:
    - Technical skills with proficiency levels
    - Soft skills with proficiency levels
    - Behaviours and attributes
    - Framework mappings (SFIA, etc.)
    - Competency clusters
    """
    
    # Get context from role intelligence
    role_purpose = role_intel.get("role_purpose", {}).get("summary", "")
    focus_areas = role_intel.get("role_scope", {}).get("primary_focus_areas", [])
    
    system_prompt = """You are a Skills Architect Agent specializing in competency frameworks.

Your mission is to create comprehensive skill architectures for roles.

You have expert knowledge of:
- SFIA (Skills Framework for the Information Age) levels and skills
- Technical skill taxonomies
- Soft skill frameworks
- Behavioural competency models
- Industry-specific skill requirements

Respond ONLY with valid JSON. No explanatory text."""

    user_prompt = f"""Create a comprehensive skills architecture for this role:

ROLE: {role_title}
DOMAIN: {domain}
SPECIALISM: {specialism}
LEVEL: {proficiency_level}
PURPOSE: {role_purpose}
FOCUS AREAS: {', '.join(focus_areas) if focus_areas else 'General'}

Provide a JSON object with this structure:

{{
    "framework_mapping": {{
        "primary_framework": "SFIA",
        "sfia_level": "The SFIA level (1-7) appropriate for {proficiency_level}",
        "sfia_level_description": "Description of what that SFIA level means",
        "relevant_sfia_skills": [
            {{"code": "SFIA code", "name": "Skill name", "level": "Level for this role"}}
        ]
    }},
    
    "technical_skills": {{
        "programming_languages": [
            {{"skill": "Language name", "proficiency": "Expert/Advanced/Intermediate/Basic", "priority": "Essential/Desirable"}}
        ],
        "tools_and_platforms": [
            {{"skill": "Tool/Platform name", "proficiency": "Level", "priority": "Essential/Desirable"}}
        ],
        "technologies": [
            {{"skill": "Technology name", "proficiency": "Level", "priority": "Essential/Desirable"}}
        ],
        "methodologies": [
            {{"skill": "Methodology name", "proficiency": "Level", "priority": "Essential/Desirable"}}
        ],
        "domain_specific": [
            {{"skill": "Domain skill", "proficiency": "Level", "priority": "Essential/Desirable"}}
        ]
    }},
    
    "soft_skills": [
        {{
            "skill": "Soft skill name",
            "description": "What this means in the context of this role",
            "proficiency": "Advanced/Intermediate/Developing",
            "priority": "Essential/Desirable"
        }}
    ],
    
    "behaviours": [
        {{
            "behaviour": "Behaviour name",
            "description": "Observable behaviour description",
            "importance": "Critical/Important/Beneficial"
        }}
    ],
    
    "competency_clusters": [
        {{
            "cluster_name": "Name of competency cluster",
            "description": "What this cluster covers",
            "competencies": ["Competency 1", "Competency 2", "Competency 3"]
        }}
    ]
}}

Be specific and realistic for {role_title} at {proficiency_level} level in {domain}/{specialism}.
Include at least:
- 4-6 programming languages/tools
- 8-12 tools and platforms
- 6-10 technologies
- 4-6 methodologies
- 6-8 soft skills
- 5-7 behaviours"""

    response = await call_agent("Skills Architect Agent", system_prompt, user_prompt)
    return parse_json_response(response)


# ============================================================================
# AGENT 3: COMPLIANCE AGENT
# ============================================================================

async def run_compliance_agent(
    role_title: str,
    domain: str,
    specialism: str
) -> Dict:
    """
    COMPLIANCE AGENT
    
    Mission: Identify all legal, regulatory, and professional requirements.
    
    Outputs:
    - Professional bodies and membership requirements
    - Legal and regulatory requirements
    - Mandatory qualifications and certifications
    - Industry standards and codes of practice
    - Equality and diversity considerations
    """
    
    system_prompt = """You are a Compliance Agent specializing in professional and regulatory requirements.

Your mission is to identify all compliance requirements for roles.

You have expert knowledge of:
- UK professional bodies and their requirements
- UK employment law and regulations
- Data protection (UK GDPR, DPA 2018)
- Industry-specific regulations
- Professional certifications and qualifications
- Equality Act 2010 requirements

Respond ONLY with valid JSON. No explanatory text."""

    user_prompt = f"""Identify all compliance requirements for this role:

ROLE: {role_title}
DOMAIN: {domain}
SPECIALISM: {specialism}
JURISDICTION: United Kingdom

Provide a JSON object with this structure:

{{
    "professional_bodies": [
        {{
            "name": "Full name of professional body",
            "abbreviation": "Acronym",
            "role": "What they do for this profession",
            "membership_required": true/false,
            "membership_levels": ["Level 1", "Level 2"],
            "benefits": ["Benefit 1", "Benefit 2"]
        }}
    ],
    
    "regulatory_position": {{
        "statutory_regulation": true/false,
        "regulatory_body": "Name if applicable, or null",
        "protected_title": true/false,
        "summary": "Brief description of regulatory status"
    }},
    
    "legal_requirements": [
        {{
            "legislation": "Act/Regulation name and year",
            "relevance": "How it applies to this role",
            "key_obligations": ["Obligation 1", "Obligation 2"],
            "mandatory_training": true/false
        }}
    ],
    
    "qualifications": {{
        "essential": [
            {{
                "qualification": "Qualification name",
                "level": "RQF/EQF level or equivalent",
                "alternatives": ["Alternative qualification"]
            }}
        ],
        "desirable": [
            {{
                "qualification": "Qualification name",
                "level": "Level",
                "value_add": "Why it's desirable"
            }}
        ],
        "professional_certifications": [
            {{
                "certification": "Certification name",
                "issuing_body": "Who issues it",
                "validity_period": "How long it's valid",
                "priority": "Essential/Highly Desirable/Desirable"
            }}
        ]
    }},
    
    "experience_requirements": {{
        "minimum_years": "X-Y years",
        "experience_types": [
            {{
                "type": "Type of experience",
                "description": "What this involves",
                "priority": "Essential/Desirable"
            }}
        ]
    }},
    
    "cpd_requirements": {{
        "professional_body_cpd": "CPD policy of main professional body",
        "recommended_hours": "Annual hours",
        "mandatory_topics": ["Topic 1", "Topic 2"],
        "recertification_cycle": "Frequency"
    }},
    
    "security_and_vetting": {{
        "dbs_required": true/false,
        "dbs_level": "Basic/Standard/Enhanced",
        "security_clearance": "Level if applicable",
        "other_checks": ["Check 1", "Check 2"]
    }},
    
    "equality_and_inclusion": {{
        "legislation": "Equality Act 2010",
        "considerations": [
            "Consideration for inclusive recruitment and role design"
        ],
        "reasonable_adjustments": "Statement about workplace adjustments",
        "bias_considerations": ["Area where bias should be monitored"]
    }}
}}

Be specific to {role_title} in {domain}/{specialism} in the UK context.
Name REAL professional bodies, REAL legislation, REAL certifications."""

    response = await call_agent("Compliance Agent", system_prompt, user_prompt)
    return parse_json_response(response)


# ============================================================================
# AGENT 4: QUALITY VALIDATOR AGENT
# ============================================================================

async def run_quality_validator_agent(combined_data: Dict) -> Dict:
    """
    QUALITY VALIDATOR AGENT
    
    Mission: Validate completeness and identify gaps.
    
    Outputs:
    - Completeness score
    - Identified gaps
    - Recommendations
    - Quality assessment
    """
    
    system_prompt = """You are a Quality Validator Agent.

Your mission is to assess the quality and completeness of role analysis data.

You evaluate:
- Completeness of each section
- Consistency across sections
- Gaps that need attention
- Overall quality score

Respond ONLY with valid JSON. No explanatory text."""

    # Summarize what we have
    role_title = combined_data.get("metadata", {}).get("role_title", "Unknown")
    has_role_intel = bool(combined_data.get("role_intelligence", {}).get("role_identification"))
    has_skills = bool(combined_data.get("skills_architecture", {}).get("technical_skills"))
    has_compliance = bool(combined_data.get("compliance", {}).get("professional_bodies"))
    
    user_prompt = f"""Validate this role analysis data:

ROLE: {role_title}

DATA SUMMARY:
- Role Intelligence: {'Complete' if has_role_intel else 'Missing'}
- Skills Architecture: {'Complete' if has_skills else 'Missing'}
- Compliance Data: {'Complete' if has_compliance else 'Missing'}

ROLE INTELLIGENCE PREVIEW:
{json.dumps(combined_data.get('role_intelligence', {}), indent=2)[:1000]}

SKILLS PREVIEW:
{json.dumps(combined_data.get('skills_architecture', {}), indent=2)[:1000]}

COMPLIANCE PREVIEW:
{json.dumps(combined_data.get('compliance', {}), indent=2)[:1000]}

Provide validation results:

{{
    "validation_summary": {{
        "overall_score": <0-100>,
        "completeness_score": <0-100>,
        "consistency_score": <0-100>,
        "quality_rating": "Excellent/Good/Acceptable/Needs Improvement"
    }},
    
    "section_scores": {{
        "role_intelligence": <0-100>,
        "skills_architecture": <0-100>,
        "compliance": <0-100>
    }},
    
    "identified_gaps": [
        {{
            "section": "Section name",
            "gap": "Description of what's missing",
            "severity": "High/Medium/Low",
            "recommendation": "How to address"
        }}
    ],
    
    "strengths": [
        "Strength 1",
        "Strength 2"
    ],
    
    "recommendations": [
        "Recommendation 1",
        "Recommendation 2"
    ],
    
    "certification": {{
        "is_valid": true/false,
        "validation_date": "{datetime.now().isoformat()}",
        "validator": "NOVA Quality Validator Agent v7.0"
    }}
}}"""

    response = await call_agent("Quality Validator Agent", system_prompt, user_prompt)
    return parse_json_response(response)


# ============================================================================
# DOCUMENT BUILDER
# ============================================================================

def build_skills_report(data: Dict, output_path: Path):
    """
    Build consolidated Skills Report document.
    
    Single document containing all analysis from all agents.
    """
    
    doc = Document()
    
    # ========================================================================
    # HELPER FUNCTION FOR TABLE HEADER FORMATTING
    # ========================================================================
    
    def format_table_header(table, header_color_hex="365F91"):
        """
        Format table header row with:
        - Background color
        - White bold text
        - Repeat header row on page break
        """
        if len(table.rows) == 0:
            return
        
        header_row = table.rows[0]
        
        # Set repeat header row for tables spanning pages
        tr = header_row._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        trPr.append(tblHeader)
        
        # Format each cell in header row
        for cell in header_row.cells:
            # Set background color
            shading = OxmlElement('w:shd')
            shading.set(qn('w:fill'), header_color_hex)
            cell._tc.get_or_add_tcPr().append(shading)
            
            # Set text to white and bold
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    run.font.bold = True
    
    # ========================================================================
    # CONFIGURE STYLES
    # ========================================================================
    
    # Page margins
    for section in doc.sections:
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
    
    # Normal style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Roboto'
    normal_style.font.size = Pt(11)
    normal_style.paragraph_format.space_before = Pt(6)
    normal_style.paragraph_format.space_after = Pt(6)
    normal_style.paragraph_format.line_spacing = 1.5
    
    # Heading 1 style
    h1_style = doc.styles['Heading 1']
    h1_style.font.name = 'Roboto'
    h1_style.font.size = Pt(18)
    h1_style.font.bold = True
    h1_style.font.color.rgb = RGBColor(0x36, 0x5F, 0x91)  # #365F91
    h1_style.paragraph_format.space_before = Pt(12)
    h1_style.paragraph_format.space_after = Pt(6)
    
    # Heading 2 style
    h2_style = doc.styles['Heading 2']
    h2_style.font.name = 'Roboto'
    h2_style.font.size = Pt(14)
    h2_style.font.bold = True
    h2_style.font.color.rgb = RGBColor(0x00, 0x20, 0x60)  # #002060
    h2_style.paragraph_format.space_before = Pt(10)
    h2_style.paragraph_format.space_after = Pt(6)
    
    # List Bullet style
    list_style = doc.styles['List Bullet']
    list_style.paragraph_format.space_before = Pt(6)
    list_style.paragraph_format.space_after = Pt(6)
    list_style.paragraph_format.line_spacing = 1.5
    
    # Title style
    title_style = doc.styles['Title']
    title_style.font.name = 'Roboto'
    title_style.font.size = Pt(26)
    title_style.font.color.rgb = RGBColor(0x17, 0x36, 0x5D)  # #17365D
    title_style.paragraph_format.space_after = Pt(15)
    
    metadata = data.get("metadata", {})
    role_title = metadata.get("role_title", "Role")
    domain = metadata.get("domain", "")
    specialism = metadata.get("specialism", "")
    proficiency_level = metadata.get("proficiency_level", "")
    
    # ========================================================================
    # TITLE PAGE
    # ========================================================================
    
    title = doc.add_heading(role_title, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    subtitle = doc.add_paragraph("Skills Report")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.size = Pt(18)
    subtitle.runs[0].font.color.rgb = RGBColor(0x36, 0x5F, 0x91)  # Match Heading 1 color
    subtitle.paragraph_format.space_before = Pt(6)
    subtitle.paragraph_format.space_after = Pt(6)
    subtitle.paragraph_format.line_spacing = 1.5
    
    doc.add_paragraph()
    
    # Metadata table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.style = 'Table Grid'
    
    meta_rows = [
        ("Domain", domain),
        ("Specialism", specialism),
        ("Proficiency Level", proficiency_level),
        ("Report Date", datetime.now().strftime("%d %B %Y")),
        ("NOVA Version", "7.0.0")
    ]
    
    for i, (label, value) in enumerate(meta_rows):
        meta_table.rows[i].cells[0].text = label
        meta_table.rows[i].cells[1].text = str(value)
        meta_table.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
    
    doc.add_page_break()
    
    # ========================================================================
    # TABLE OF CONTENTS
    # ========================================================================
    
    doc.add_heading("Contents", level=1)
    
    toc_items = [
        "1. Executive Summary",
        "2. Role Identification",
        "3. Role Purpose and Scope",
        "4. Organisational Context",
        "5. Career Progression",
        "6. Skills Framework Mapping",
        "7. Technical Skills",
        "8. Soft Skills",
        "9. Behaviours",
        "10. Professional Bodies",
        "11. Legal and Regulatory Requirements",
        "12. Qualifications",
        "13. Experience Requirements",
        "14. CPD Requirements",
        "15. Security and Vetting",
        "16. Equality and Inclusion",
        "17. Quality Validation"
    ]
    
    for item in toc_items:
        doc.add_paragraph(item)
    
    doc.add_page_break()
    
    # ========================================================================
    # 1. EXECUTIVE SUMMARY
    # ========================================================================
    
    doc.add_heading("1. Executive Summary", level=1)
    
    role_intel = data.get("role_intelligence", {})
    role_purpose = role_intel.get("role_purpose", {})
    validation = data.get("validation", {})
    
    doc.add_heading("Role Overview", level=2)
    doc.add_paragraph(role_purpose.get("summary", f"Skills analysis for {role_title} at {proficiency_level} level."))
    
    doc.add_heading("Business Outcomes", level=2)
    outcomes = role_purpose.get("business_outcomes", [])
    for outcome in outcomes:
        doc.add_paragraph(outcome, style='List Bullet')
    
    # Quality score
    val_summary = validation.get("validation_summary", {})
    if val_summary.get("overall_score"):
        doc.add_heading("Quality Assessment", level=2)
        quality_table = doc.add_table(rows=2, cols=2)
        quality_table.style = 'Table Grid'
        quality_table.rows[0].cells[0].text = "Overall Score"
        quality_table.rows[0].cells[1].text = f"{val_summary.get('overall_score') or 'N/A'}%"
        quality_table.rows[1].cells[0].text = "Quality Rating"
        quality_table.rows[1].cells[1].text = val_summary.get('quality_rating') or 'N/A'
    
    doc.add_page_break()
    
    # ========================================================================
    # 2. ROLE IDENTIFICATION
    # ========================================================================
    
    doc.add_heading("2. Role Identification", level=1)
    
    role_id = role_intel.get("role_identification") or {}
    
    id_table = doc.add_table(rows=6, cols=2)
    id_table.style = 'Table Grid'
    
    id_rows = [
        ("Role Title", role_id.get("role_title") or role_title),
        ("Role Family", role_id.get("role_family") or "Not specified"),
        ("Job Level", role_id.get("job_level") or proficiency_level),
        ("Employment Types", ", ".join(role_id.get("employment_types") or ["Permanent"])),
        ("Work Arrangements", ", ".join(role_id.get("work_arrangements") or ["Hybrid"])),
        ("Team Size", role_id.get("typical_team_size") or "N/A")
    ]
    
    for i, (label, value) in enumerate(id_rows):
        id_table.rows[i].cells[0].text = label
        id_table.rows[i].cells[1].text = str(value)
        id_table.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
    
    # Equivalent titles
    equiv_titles = role_intel.get("equivalent_titles") or []
    if equiv_titles:
        doc.add_heading("Equivalent Job Titles", level=2)
        for title_item in equiv_titles:
            doc.add_paragraph(title_item, style='List Bullet')
    
    doc.add_page_break()
    
    # ========================================================================
    # 3. ROLE PURPOSE AND SCOPE
    # ========================================================================
    
    doc.add_heading("3. Role Purpose and Scope", level=1)
    
    doc.add_heading("Purpose", level=2)
    doc.add_paragraph(role_purpose.get("summary") or "Not specified")
    
    role_scope = role_intel.get("role_scope") or {}
    
    doc.add_heading("Primary Focus Areas", level=2)
    for area in role_scope.get("primary_focus_areas") or []:
        doc.add_paragraph(area, style='List Bullet')
    
    if role_scope.get("decision_authority"):
        doc.add_heading("Decision Authority", level=2)
        doc.add_paragraph(role_scope.get("decision_authority"))
    
    if role_scope.get("geographic_scope"):
        doc.add_heading("Geographic Scope", level=2)
        doc.add_paragraph(role_scope.get("geographic_scope"))
    
    doc.add_page_break()
    
    # ========================================================================
    # 4. ORGANISATIONAL CONTEXT
    # ========================================================================
    
    doc.add_heading("4. Organisational Context", level=1)
    
    org_context = role_intel.get("organisational_context") or {}
    
    doc.add_heading("Typical Employers", level=2)
    for employer in org_context.get("typical_employers") or []:
        doc.add_paragraph(employer, style='List Bullet')
    
    doc.add_heading("Sectors", level=2)
    for sector in org_context.get("sectors") or []:
        doc.add_paragraph(sector, style='List Bullet')
    
    doc.add_heading("Reporting Lines", level=2)
    for line in org_context.get("reporting_lines") or []:
        doc.add_paragraph(f"Reports to: {line}", style='List Bullet')
    
    doc.add_heading("Key Stakeholders", level=2)
    for stakeholder in org_context.get("key_stakeholders") or []:
        doc.add_paragraph(stakeholder, style='List Bullet')
    
    doc.add_page_break()
    
    # ========================================================================
    # 5. CAREER PROGRESSION
    # ========================================================================
    
    doc.add_heading("5. Career Progression", level=1)
    
    career = role_intel.get("career_context") or {}
    
    doc.add_heading("Entry Routes", level=2)
    for route in career.get("typical_entry_routes") or []:
        doc.add_paragraph(route, style='List Bullet')
    
    progression = career.get("progression_paths") or {}
    
    doc.add_heading("Technical Progression Path", level=2)
    for role_next in progression.get("technical") or []:
        doc.add_paragraph(role_next, style='List Bullet')
    
    doc.add_heading("Leadership Progression Path", level=2)
    for role_next in progression.get("leadership") or []:
        doc.add_paragraph(role_next, style='List Bullet')
    
    if career.get("typical_tenure"):
        doc.add_heading("Typical Tenure", level=2)
        doc.add_paragraph(career.get("typical_tenure"))
    
    doc.add_page_break()
    
    # ========================================================================
    # 6. SKILLS FRAMEWORK MAPPING
    # ========================================================================
    
    doc.add_heading("6. Skills Framework Mapping", level=1)
    
    skills_arch = data.get("skills_architecture") or {}
    framework_map = skills_arch.get("framework_mapping") or {}
    
    doc.add_heading("SFIA Mapping", level=2)
    
    sfia_table = doc.add_table(rows=3, cols=2)
    sfia_table.style = 'Table Grid'
    sfia_table.rows[0].cells[0].text = "SFIA Level"
    sfia_table.rows[0].cells[1].text = str(framework_map.get("sfia_level") or "5")
    sfia_table.rows[1].cells[0].text = "Level Description"
    sfia_table.rows[1].cells[1].text = framework_map.get("sfia_level_description") or "N/A"
    sfia_table.rows[2].cells[0].text = "Primary Framework"
    sfia_table.rows[2].cells[1].text = framework_map.get("primary_framework") or "SFIA"
    
    for row in sfia_table.rows:
        row.cells[0].paragraphs[0].runs[0].font.bold = True
    
    sfia_skills = framework_map.get("relevant_sfia_skills") or []
    if sfia_skills:
        doc.add_heading("Relevant SFIA Skills", level=2)
        skill_table = doc.add_table(rows=len(sfia_skills)+1, cols=3)
        skill_table.style = 'Table Grid'
        skill_table.rows[0].cells[0].text = "Code"
        skill_table.rows[0].cells[1].text = "Skill"
        skill_table.rows[0].cells[2].text = "Level"
        format_table_header(skill_table)
        
        for i, skill in enumerate(sfia_skills, 1):
            if i < len(skill_table.rows):
                skill_table.rows[i].cells[0].text = skill.get("code") or ""
                skill_table.rows[i].cells[1].text = skill.get("name") or ""
                skill_table.rows[i].cells[2].text = str(skill.get("level") or "")
    
    doc.add_page_break()
    
    # ========================================================================
    # 7. TECHNICAL SKILLS
    # ========================================================================
    
    doc.add_heading("7. Technical Skills", level=1)
    
    tech_skills = skills_arch.get("technical_skills") or {}
    
    # Helper function to add skills table
    def add_skills_section(title: str, skills: List[Dict]):
        if skills:
            doc.add_heading(title, level=2)
            table = doc.add_table(rows=len(skills)+1, cols=3)
            table.style = 'Table Grid'
            table.rows[0].cells[0].text = "Skill"
            table.rows[0].cells[1].text = "Proficiency"
            table.rows[0].cells[2].text = "Priority"
            format_table_header(table)
            
            for i, skill in enumerate(skills, 1):
                if i < len(table.rows):
                    table.rows[i].cells[0].text = skill.get("skill") or ""
                    table.rows[i].cells[1].text = skill.get("proficiency") or ""
                    table.rows[i].cells[2].text = skill.get("priority") or ""
    
    add_skills_section("Programming Languages", tech_skills.get("programming_languages") or [])
    add_skills_section("Tools and Platforms", tech_skills.get("tools_and_platforms") or [])
    add_skills_section("Technologies", tech_skills.get("technologies") or [])
    add_skills_section("Methodologies", tech_skills.get("methodologies") or [])
    add_skills_section("Domain-Specific Skills", tech_skills.get("domain_specific") or [])
    
    doc.add_page_break()
    
    # ========================================================================
    # 8. SOFT SKILLS
    # ========================================================================
    
    doc.add_heading("8. Soft Skills", level=1)
    
    soft_skills = skills_arch.get("soft_skills") or []
    
    if soft_skills:
        table = doc.add_table(rows=len(soft_skills)+1, cols=4)
        table.style = 'Table Grid'
        headers = ["Skill", "Description", "Proficiency", "Priority"]
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
        format_table_header(table)
        
        for i, skill in enumerate(soft_skills, 1):
            if i < len(table.rows):
                table.rows[i].cells[0].text = skill.get("skill") or ""
                table.rows[i].cells[1].text = skill.get("description") or ""
                table.rows[i].cells[2].text = skill.get("proficiency") or ""
                table.rows[i].cells[3].text = skill.get("priority") or ""
    
    doc.add_page_break()
    
    # ========================================================================
    # 9. BEHAVIOURS
    # ========================================================================
    
    doc.add_heading("9. Behaviours", level=1)
    
    behaviours = skills_arch.get("behaviours") or []
    
    if behaviours:
        table = doc.add_table(rows=len(behaviours)+1, cols=3)
        table.style = 'Table Grid'
        headers = ["Behaviour", "Description", "Importance"]
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
        format_table_header(table)
        
        for i, behav in enumerate(behaviours, 1):
            if i < len(table.rows):
                table.rows[i].cells[0].text = behav.get("behaviour") or ""
                table.rows[i].cells[1].text = behav.get("description") or ""
                table.rows[i].cells[2].text = behav.get("importance") or ""
    
    doc.add_page_break()
    
    # ========================================================================
    # 10. PROFESSIONAL BODIES
    # ========================================================================
    
    doc.add_heading("10. Professional Bodies", level=1)
    
    compliance = data.get("compliance", {})
    prof_bodies = compliance.get("professional_bodies") or []
    
    for body in prof_bodies:
        doc.add_heading(body.get("name") or "Professional Body", level=2)
        
        body_table = doc.add_table(rows=4, cols=2)
        body_table.style = 'Table Grid'
        body_table.rows[0].cells[0].text = "Abbreviation"
        body_table.rows[0].cells[1].text = body.get("abbreviation") or "N/A"
        body_table.rows[1].cells[0].text = "Role"
        body_table.rows[1].cells[1].text = body.get("role") or "N/A"
        body_table.rows[2].cells[0].text = "Membership Required"
        body_table.rows[2].cells[1].text = "Yes" if body.get("membership_required") else "No"
        body_table.rows[3].cells[0].text = "Membership Levels"
        body_table.rows[3].cells[1].text = ", ".join(body.get("membership_levels") or []) or "N/A"
        
        for row in body_table.rows:
            row.cells[0].paragraphs[0].runs[0].font.bold = True
        
        doc.add_paragraph()
    
    # Regulatory position
    reg_pos = compliance.get("regulatory_position") or {}
    if reg_pos:
        doc.add_heading("Regulatory Position", level=2)
        doc.add_paragraph(reg_pos.get("summary") or "Not specified")
    
    doc.add_page_break()
    
    # ========================================================================
    # 11. LEGAL AND REGULATORY REQUIREMENTS
    # ========================================================================
    
    doc.add_heading("11. Legal and Regulatory Requirements", level=1)
    
    legal_reqs = compliance.get("legal_requirements") or []
    
    for req in legal_reqs:
        doc.add_heading(req.get("legislation") or "Legislation", level=2)
        doc.add_paragraph(f"Relevance: {req.get('relevance') or 'N/A'}")
        
        obligations = req.get("key_obligations") or []
        if obligations:
            doc.add_paragraph("Key Obligations:")
            for ob in obligations:
                doc.add_paragraph(ob, style='List Bullet')
        
        if req.get("mandatory_training"):
            doc.add_paragraph("Mandatory training required: Yes")
        
        doc.add_paragraph()
    
    doc.add_page_break()
    
    # ========================================================================
    # 12. QUALIFICATIONS
    # ========================================================================
    
    doc.add_heading("12. Qualifications", level=1)
    
    quals = compliance.get("qualifications") or {}
    
    doc.add_heading("Essential Qualifications", level=2)
    for qual in quals.get("essential") or []:
        doc.add_paragraph(f"{qual.get('qualification') or 'N/A'} - {qual.get('level') or 'N/A'}", style='List Bullet')
    
    doc.add_heading("Desirable Qualifications", level=2)
    for qual in quals.get("desirable") or []:
        doc.add_paragraph(f"{qual.get('qualification') or 'N/A'} - {qual.get('value_add') or 'N/A'}", style='List Bullet')
    
    doc.add_heading("Professional Certifications", level=2)
    certs = quals.get("professional_certifications") or []
    if certs:
        cert_table = doc.add_table(rows=len(certs)+1, cols=4)
        cert_table.style = 'Table Grid'
        headers = ["Certification", "Issuing Body", "Validity", "Priority"]
        for i, header in enumerate(headers):
            cert_table.rows[0].cells[i].text = header
        format_table_header(cert_table)
        
        for i, cert in enumerate(certs, 1):
            if i < len(cert_table.rows):
                cert_table.rows[i].cells[0].text = cert.get("certification") or ""
                cert_table.rows[i].cells[1].text = cert.get("issuing_body") or ""
                cert_table.rows[i].cells[2].text = cert.get("validity_period") or ""
                cert_table.rows[i].cells[3].text = cert.get("priority") or ""
    
    doc.add_page_break()
    
    # ========================================================================
    # 13. EXPERIENCE REQUIREMENTS
    # ========================================================================
    
    doc.add_heading("13. Experience Requirements", level=1)
    
    exp_reqs = compliance.get("experience_requirements") or {}
    
    doc.add_heading("Minimum Experience", level=2)
    doc.add_paragraph(exp_reqs.get("minimum_years") or "Not specified")
    
    doc.add_heading("Experience Types", level=2)
    for exp in exp_reqs.get("experience_types") or []:
        doc.add_paragraph(f"{exp.get('type') or 'N/A'} ({exp.get('priority') or 'N/A'})", style='List Bullet')
        if exp.get("description"):
            doc.add_paragraph(f"  {exp.get('description')}")
    
    doc.add_page_break()
    
    # ========================================================================
    # 14. CPD REQUIREMENTS
    # ========================================================================
    
    doc.add_heading("14. CPD Requirements", level=1)
    
    cpd = compliance.get("cpd_requirements", {})
    
    cpd_table = doc.add_table(rows=4, cols=2)
    cpd_table.style = 'Table Grid'
    cpd_table.rows[0].cells[0].text = "Professional Body CPD"
    cpd_table.rows[0].cells[1].text = cpd.get("professional_body_cpd") or "N/A"
    cpd_table.rows[1].cells[0].text = "Recommended Hours"
    cpd_table.rows[1].cells[1].text = cpd.get("recommended_hours") or "N/A"
    cpd_table.rows[2].cells[0].text = "Recertification Cycle"
    cpd_table.rows[2].cells[1].text = cpd.get("recertification_cycle") or "N/A"
    cpd_table.rows[3].cells[0].text = "Mandatory Topics"
    cpd_table.rows[3].cells[1].text = ", ".join(cpd.get("mandatory_topics") or []) or "N/A"
    
    for row in cpd_table.rows:
        row.cells[0].paragraphs[0].runs[0].font.bold = True
    
    doc.add_page_break()
    
    # ========================================================================
    # 15. SECURITY AND VETTING
    # ========================================================================
    
    doc.add_heading("15. Security and Vetting", level=1)
    
    security = compliance.get("security_and_vetting", {})
    
    sec_table = doc.add_table(rows=4, cols=2)
    sec_table.style = 'Table Grid'
    sec_table.rows[0].cells[0].text = "DBS Required"
    sec_table.rows[0].cells[1].text = "Yes" if security.get("dbs_required") else "No"
    sec_table.rows[1].cells[0].text = "DBS Level"
    sec_table.rows[1].cells[1].text = security.get("dbs_level") or "N/A"
    sec_table.rows[2].cells[0].text = "Security Clearance"
    sec_table.rows[2].cells[1].text = security.get("security_clearance") or "N/A"
    sec_table.rows[3].cells[0].text = "Other Checks"
    sec_table.rows[3].cells[1].text = ", ".join(security.get("other_checks") or []) or "N/A"
    
    for row in sec_table.rows:
        row.cells[0].paragraphs[0].runs[0].font.bold = True
    
    doc.add_page_break()
    
    # ========================================================================
    # 16. EQUALITY AND INCLUSION
    # ========================================================================
    
    doc.add_heading("16. Equality and Inclusion", level=1)
    
    equality = compliance.get("equality_and_inclusion", {})
    
    doc.add_heading("Legislative Framework", level=2)
    doc.add_paragraph(equality.get("legislation") or "Equality Act 2010")
    
    doc.add_heading("Key Considerations", level=2)
    for consideration in equality.get("considerations") or []:
        doc.add_paragraph(consideration, style='List Bullet')
    
    doc.add_heading("Reasonable Adjustments", level=2)
    doc.add_paragraph(equality.get("reasonable_adjustments") or "Not specified")
    
    doc.add_page_break()
    
    # ========================================================================
    # 17. QUALITY VALIDATION
    # ========================================================================
    
    doc.add_heading("17. Quality Validation", level=1)
    
    val_summary = validation.get("validation_summary") or {}
    
    doc.add_heading("Validation Scores", level=2)
    
    val_table = doc.add_table(rows=4, cols=2)
    val_table.style = 'Table Grid'
    val_table.rows[0].cells[0].text = "Overall Score"
    val_table.rows[0].cells[1].text = f"{val_summary.get('overall_score') or 'N/A'}%"
    val_table.rows[1].cells[0].text = "Completeness Score"
    val_table.rows[1].cells[1].text = f"{val_summary.get('completeness_score') or 'N/A'}%"
    val_table.rows[2].cells[0].text = "Consistency Score"
    val_table.rows[2].cells[1].text = f"{val_summary.get('consistency_score') or 'N/A'}%"
    val_table.rows[3].cells[0].text = "Quality Rating"
    val_table.rows[3].cells[1].text = val_summary.get('quality_rating') or 'N/A'
    
    for row in val_table.rows:
        row.cells[0].paragraphs[0].runs[0].font.bold = True
    
    # Strengths
    strengths = validation.get("strengths") or []
    if strengths:
        doc.add_heading("Strengths", level=2)
        for strength in strengths:
            doc.add_paragraph(strength, style='List Bullet')
    
    # Recommendations
    recommendations = validation.get("recommendations") or []
    if recommendations:
        doc.add_heading("Recommendations", level=2)
        for rec in recommendations:
            doc.add_paragraph(rec, style='List Bullet')
    
    # Certification
    cert = validation.get("certification") or {}
    doc.add_heading("Certification", level=2)
    doc.add_paragraph(f"Validated: {cert.get('validation_date') or datetime.now().isoformat()}")
    doc.add_paragraph(f"Validator: {cert.get('validator') or 'NOVA Quality Validator Agent v7.0'}")
    
    # ========================================================================
    # SAVE DOCUMENT
    # ========================================================================
    
    doc.save(str(output_path))
    print(f"[NOVA] Skills Report saved: {output_path}")


# ============================================================================
# PLACEHOLDER AGENTS (Design, Delivery, Evaluation)
# ============================================================================

async def run_design_agent(job_id: str, parameters: Dict, framework: str, terms: Dict):
    """Design Agent - placeholder for future implementation"""
    update_job(job_id, 50, "Design Agent running...")
    output_dir = Path(jobs.get(job_id)["output_dir"])
    output_dir.mkdir(exist_ok=True)
    
    doc = Document()
    doc.add_heading("Design Phase - Coming Soon", level=0)
    doc.add_paragraph("This agent will be implemented in a future release.")
    doc.save(str(output_dir / "design_placeholder.docx"))
    
    update_job(job_id, 100, "Design Agent complete")


async def run_delivery_agent(job_id: str, parameters: Dict, framework: str, terms: Dict):
    """Delivery Agent - placeholder for future implementation"""
    update_job(job_id, 50, "Delivery Agent running...")
    output_dir = Path(jobs.get(job_id)["output_dir"])
    output_dir.mkdir(exist_ok=True)
    
    doc = Document()
    doc.add_heading("Delivery Phase - Coming Soon", level=0)
    doc.add_paragraph("This agent will be implemented in a future release.")
    doc.save(str(output_dir / "delivery_placeholder.docx"))
    
    update_job(job_id, 100, "Delivery Agent complete")


async def run_evaluation_agent(job_id: str, parameters: Dict, framework: str, terms: Dict):
    """Evaluation Agent - placeholder for future implementation"""
    update_job(job_id, 50, "Evaluation Agent running...")
    output_dir = Path(jobs.get(job_id)["output_dir"])
    output_dir.mkdir(exist_ok=True)
    
    doc = Document()
    doc.add_heading("Evaluation Phase - Coming Soon", level=0)
    doc.add_paragraph("This agent will be implemented in a future release.")
    doc.save(str(output_dir / "evaluation_placeholder.docx"))
    
    update_job(job_id, 100, "Evaluation Agent complete")


# ============================================================================
# STARTUP
# ============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
