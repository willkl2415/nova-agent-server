# NOVA Agent Server v4.0

Autonomous Training Agent Execution Server

## Endpoints

- `POST /api/execute` - Start agent task
- `GET /api/status/{job_id}` - Get task status
- `GET /api/download/{job_id}` - Download ZIP
- `GET /api/health` - Health check

## Environment Variables

- `ANTHROPIC_API_KEY` - Claude API key (required)
- `PORT` - Server port (default: 8000)

## Deploy to Railway

1. Push to GitHub
2. Connect to Railway
3. Set ANTHROPIC_API_KEY environment variable
4. Deploy
