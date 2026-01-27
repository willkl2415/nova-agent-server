# NOVA Agent Server v2.0

Autonomous Training Agent Execution Server with professional document generation.

## What's New in v2.0

- **Professional .docx outputs** - All documents now generated as properly formatted Word documents
- **Framework-agnostic** - Works with UK DSAT, US TRADOC, NATO, and ASD/AIA S6000T
- **Renamed agents** - TNA → Analysis, Course Generator → Full Package
- **Enhanced formatting** - Title pages, tables, headers/footers, styling

## Agents

| Agent | Description | Outputs |
|-------|-------------|---------|
| `analysis` | Training Needs Analysis | Scoping Report, RolePS, Gap Analysis, TNR |
| `design` | Training Design | Training Objectives, Enabling Objectives |
| `delivery` | Training Delivery | Lesson Plans, Assessment Instruments |
| `full-package` | Complete lifecycle | All of the above + Compliance Certificate |

## API Endpoints

- `POST /api/execute` - Start an agent task
- `GET /api/status/{job_id}` - Get task status  
- `GET /api/download/{job_id}` - Download completed files
- `GET /api/health` - Health check

## Environment Variables

```
ANTHROPIC_API_KEY=your_key_here
NOVA_API_SECRET=optional_auth_secret
PORT=8000
```

## Deploy to Railway

1. Push this code to GitHub
2. Connect to Railway
3. Add ANTHROPIC_API_KEY environment variable
4. Deploy

## Local Development

```bash
pip install -r requirements.txt
python main.py
```

## Document Output Structure

```
/job_id/
├── 00_Compliance_Certificate.docx
├── 01_Analysis/
│   ├── 01_Scoping_Report.docx
│   ├── 02_Role_Performance_Statement.docx
│   ├── 03_Training_Gap_Analysis.docx
│   └── 04_Training_Needs_Report.docx
├── 02_Design/
│   ├── 05_Training_Objectives.docx
│   └── 06_Enabling_Objectives.docx
└── 03_Delivery/
    ├── 07_Lesson_Plans.docx
    └── 08_Assessment_Instruments.docx
```

## Backward Compatibility

Legacy agent names still work:
- `tna` → maps to `analysis`
- `course-generator` → maps to `full-package`
