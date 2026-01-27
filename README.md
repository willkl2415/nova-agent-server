# NOVA Agent Server

Autonomous DSAT Agent Execution Server for NOVA™ Allied Defence Training LLM.

## Overview

This FastAPI server executes autonomous agents that generate complete DSAT documentation packages:

- **TNA Agent**: Training Needs Analysis (Scoping Report, RolePS, TNR)
- **Design Agent**: Training Design (Learning Specification, Design Matrix)
- **Delivery Agent**: Training Delivery (Lesson Plans, Assessments)
- **Course Generator**: Complete DSAT lifecycle in one operation

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/execute` | Start an agent task |
| GET | `/api/status/{job_id}` | Get task status |
| GET | `/api/download/{job_id}` | Download completed files |
| GET | `/api/health` | Health check |

## Deploy to Railway

### 1. Create GitHub Repository

Create a new private repository called `nova-agent-server` and upload these files.

### 2. Deploy from Railway

1. Go to https://railway.app
2. Click "New Project" → "Deploy from GitHub repo"
3. Select your `nova-agent-server` repository

### 3. Configure Environment Variables

In Railway Dashboard → Variables:

```
NOVA_API_SECRET=your-secure-secret
OPENAI_API_KEY=sk-proj-xxx
ANTHROPIC_API_KEY=sk-ant-xxx
PINECONE_API_KEY=pcsk_xxx
```

### 4. Get Public URL

Railway Dashboard → Settings → Domains → Generate Domain

### 5. Update Cloudflare

Add to Cloudflare Pages environment variables:
```
NOVA_AGENT_SERVER_URL=https://your-railway-url.up.railway.app
```

## Local Development

```bash
pip install -r requirements.txt
cp .env.example .env
# Edit .env with your API keys
uvicorn main:app --reload --port 8000
```

## Test API

```bash
# Health check
curl http://localhost:8000/api/health

# Submit task
curl -X POST http://localhost:8000/api/execute \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer your-secret" \
  -d '{"job_id":"test-123","agent":"tna","parameters":{"role_title":"Test Role"}}'

# Check status
curl http://localhost:8000/api/status/test-123 \
  -H "Authorization: Bearer your-secret"
```

## File Structure

```
nova-agent-server/
├── main.py              # FastAPI server
├── requirements.txt     # Python dependencies
├── Dockerfile          # Container configuration
├── railway.json        # Railway configuration
├── Procfile           # Process configuration
├── .env.example       # Environment template
└── README.md          # This file
```

---

**NOVA™ - Allied Defence Training Intelligence**
