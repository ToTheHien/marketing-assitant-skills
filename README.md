# Marketing Assistant Skills

> A collection of Claude Code skills that automate repetitive marketing operations for industrial real estate — email triage, weekly/monthly report generation, social media content review, and more

---

## Why I Built This

The rapid rise of AI Agents brings both challenges and remarkable opportunities. I wanted to share what I have learned with my beloved wife — to help her optimize the way she works.

These skills are built specifically around her day-to-day responsibilities as the **Chief Marketing Officer of an industrial park developer**. Her work involves tracking investor pipelines, managing investment promotion activities, reviewing marketing content, and staying on top of a constant stream of email, etc...

Hope that these tools become a reliable right hand for her, freeing up more time to focus on the high-judgment work that today's AI cannot yet handle on its own.

---

## Getting Started

### Prerequisites

- [Node.js](https://nodejs.org) v18+
- [Claude Code CLI](https://claude.ai/code) installed and authenticated
- Git

### 1. Clone the repository

```bash
git clone https://github.com/<your-username>/marketing_assitant_skills.git
cd marketing_assitant_skills
```

### 2. Choose your installation scope

<details>
<summary><strong>Option A — Project-level (applies to current project only)</strong></summary>

Create `.claude` folder inside project directory then copy the skills to your project directory

**Mac / Linux:**
```bash
mkdir -p .claude/
cp -r ~/marketing_assitant_skills/.claude/* ~/.claude/
```

**Windows (PowerShell):**
```powershell
mkdir -p .claude\
New-Item -ItemType Directory -Force ".claude\"
Copy-Item "marketing_assitant_skills\.claude\*" ".claude\"
```


</details>

<details>
<summary><strong>Option B — Global installation (available in every Claude Code session)</strong></summary>

Copy the skills to your global Claude config directory:

**Mac / Linux:**
```bash
mkdir -p ~/.claude/
cp -r ~/marketing_assitant_skills/.claude/* ~/.claude/
```

**Windows (PowerShell):**
```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.claude\"
Copy-Item "marketing_assitant_skills\.claude\*" "$env:USERPROFILE\.claude\"
```

After copying, open any folder in Claude Code and the skills will be available.

</details>

---

### 3. Configure Permissions

To avoid permission prompts on every run, add the following to your project's `.claude/settings.json`:

<details>
<summary><strong>Recommended permissions configuration</strong></summary>

```json
{
  "permissions": {
    "allow": [
      "Bash(date:*)",
      "Bash(echo:*)",
      "Bash(cat:*)",
      "Bash(ls:*)",
      "Bash(mkdir:*)",
      "Bash(wc:*)",
      "Bash(head:*)",
      "Bash(tail:*)",
      "Bash(sort:*)",
      "Bash(grep:*)",
      "Bash(tr:*)",
      "Bash(git add:*)",
      "Bash(git commit:*)",
      "Bash(git status:*)",
      "Bash(git log:*)",
      "Bash(git diff:*)",
      "Bash(git tag:*)"
    ],
    "deny": [
      "Read(.env)",
      "Read(.env.*)",
      "Read(**/secrets/*)",
      "Read(**/*credential*)",
      "Read(**/*.pem)",
      "Read(**/*.key)"
    ]
  },
  "enabledMcpjsonServers": [
    "memory",
    "sequential-thinking",
    "context7",
    "playwright",
    "google-chrome",
    "canva"
  ]
}
```

The `deny` rules prevent Claude from accidentally reading sensitive credential files. Adjust `enabledMcpjsonServers` to match the MCP servers you have configured.

</details>

---

### 4. Connect Google Workspace

Several skills read from Gmail and Google Drive. Follow the setup guide to authorize Claude Code with your Google account:

[**Google Workspace Connection Guide →**](./guideline/connect-to-ggl-workspace-guideline.md)

The guide walks you through:
- Creating a Google Cloud project and OAuth credentials
- Enabling the Gmail, Drive, Calendar, Docs, and Sheets APIs
- Running `gws auth login` to authenticate

---

## How It Works

Each skill is a slash command you run directly in the Claude Code chat. Claude reads your arguments, calls the relevant APIs or reads local files, and produces structured output saved to a timestamped folder inside the project.

### Skills Overview

| Skill | Command | What it does |
|---|---|---|
| Email Summary | `/email-summary` | Reads your Gmail inbox and produces a structured triage report |
| Investment Progress | `/investment-progress-summary` | Generates a `.docx` invesment progress report |
| Promote Activity | `/promote-activity-summary` | Generates a `.docx` report on investment promotion activities |
| Social Media Review | `/social-media-review` | Reviews a social post from Google Drive for quality and fit |
| Doc to Markdown | `/doc-to-markdown` | Converts a `.docx` file to clean Markdown |

---

### `/email-summary [days] [category]`

Connects to Gmail, fetches recent emails, and produces a prioritized summary organized into action-required, informational, and newsletter buckets.

**Usage:**
```
/email-summary              # Last 7 days, all categories (default)
/email-summary 14           # Last 14 days
/email-summary 7 invoice    # Last 7 days, filter by "invoice" category
```

**Output:** saved to `email_summary/YYYY-MM-DD/summary.md` (same-day reruns produce `summary(2).md`, `summary(3).md`, etc.)

---

### `/summary-investment-progress <mode> [references...] [--context <text>]`

Reads investor data from Google Drive links or local files and compiles a professional `.docx` report covering project progress, disbursement rates, and next-period plans.

**Usage:**
```
/summary-investment-progress weekly
/summary-investment-progress monthly https://drive.google.com/file/abc123
/summary-investment-progress weekly ./data/tuan18.xlsx --context 3 new investors signed this week
```

**Output:** saved to `investment_progress/weekly/YYYY-MM-DD/` or `investment_progress/monthly/YYYY-MM-DD/` (same-day reruns append `(2)`, `(3)`, etc. before the file extension)

---

### `/summary-promote-activity <mode> [references...] [--context <text>]`

Reads promotion activity data and generates a `.docx` report covering investor meetings, trade association engagements, and events — with pipeline status and follow-up actions.

**Usage:**
```
/summary-promote-activity weekly
/summary-promote-activity monthly https://drive.google.com/file/abc https://drive.google.com/file/xyz
/summary-promote-activity weekly ./data/tuan18.xlsx --context Site visit from KCCI delegation on Wednesday
```

**Output:** saved to `promote-activity/weekly/YYYY-MM-DD/` or `promote-activity/monthly/YYYY-MM-DD/` (same-day reruns append `(2)`, `(3)`, etc. before the file extension)

---

### `/social-media-review <google_drive_link> <platform> <objective>`

Downloads a draft post from Google Drive and runs a structured review against platform best practices and your stated campaign objective. Covers tone, clarity, call-to-action, and brand compliance.

**Supported platforms:** `Facebook`, `LinkedIn`, `Website`

**Usage:**
```
/social-media-review https://drive.google.com/file/abc LinkedIn "attract FDI investors"
/social-media-review https://drive.google.com/file/xyz Facebook "promote industrial park expansion"
```

**Output:** saved to `social_media/YYYY-MM-DD/review.md` (same-day reruns produce `review(2).md`, `review(3).md`, etc.)

---

### `/doc-to-markdown <input> <output>`

Converts a Word document (local path or URL) to clean Markdown, extracting embedded images into a companion folder.

**Usage:**
```
/doc-to-markdown ./documents/report.docx ./documents/output/
/doc-to-markdown https://example.com/file.docx ./output/
```

**Output:** `<filename>.md` and optional `<filename>_images/` folder in the output directory

---

## AI Agents

The project also includes specialized sub-agents that Claude uses automatically when relevant:

| Agent | Role |
|---|---|
| `brand-reviewer` | Reviews marketing content for brand voice and quality before publishing |
| `compliance-checker` | Checks materials against commercial real estate regulatory requirements |
| `data-analyst` | Analyzes pipeline and campaign performance data |
| `market-analyst` | Researches industrial real estate market trends and investment opportunities |

---

## Contributing

Feel free to open issues, suggest new skills, or submit pull requests. Improvements around new report formats, additional platforms, or better Vietnamese language handling are especially welcome.

**If this project saves you time — star the repo. It helps others find it.**

---

## License

MIT - Use freely, modify as needed, contribute back if you can.
