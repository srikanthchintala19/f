---
inclusion: always
---

# Resume Tailoring Workflow — Srikanth Chintala

This is the persistent playbook to follow **every time the user pastes a Job Description (JD)**.

## User-Confirmed Settings

| Setting | Value |
|---------|-------|
| **Canonical resume** | `Srikanth Chintala_Resume.pdf` (root of repo) |
| **Output formats** | Both `.docx` AND `.pdf` |
| **Target geography** | India |
| **Target seniority/titles** | PM, Senior PM, Lead PM, Program Manager, PMO Analyst, PM Analyst, and similar variants |
| **LinkedIn URL** | https://www.linkedin.com/in/chintalasrikanth97 |
| **Cover letter** | Yes — generate every time |
| **Change log style** | `Before → After + Why` (which JD requirement it addresses) |

## The 3-Step Playbook (Execute in Order)

### Step 1 — Senior Recruiter Evaluation
Act as a senior recruiter (10+ years India tech/consulting hiring experience). Evaluate the resume against the JD with rigor.

**Must include:**
- **Match Score** (out of 100) with breakdown: Skills, Experience, Keywords, ATS, Cultural Fit
- **Strengths** the resume already has aligned to the JD
- **Gaps** — explicit JD requirements NOT visible in the resume
- **Red flags** a recruiter would spot (gaps, inconsistencies, weak phrasing)
- **Company research** — check the company's website, LinkedIn page, Glassdoor, recent news, and current employees in similar roles. Note culture, tech stack, hiring patterns
- **Job market context** — search comparable JDs (Naukri, LinkedIn, Indeed, Instahyre) to identify the most-demanded keywords for this role family in India

### Step 2 — Complete Report Analysis
Deliver a structured Markdown report saved at:
`reports/<company>_<role>_<YYYY-MM-DD>.md`

**Sections:**
1. JD Summary (1 paragraph)
2. Match Score & Breakdown (table)
3. Strengths (bulleted, mapped to JD lines)
4. Gaps (bulleted, mapped to JD lines)
5. Red Flags & Recruiter Concerns
6. Company Intelligence (culture, tech, leadership, recent news)
7. Market Keyword Analysis (top 15 keywords from comparable JDs)
8. Recommended Resume Changes (high-priority list)
9. Interview Prep Hot-Spots (likely tough questions)

### Step 3 — Tailored Resume + Cover Letter + Change Log
Produce these deliverables:

| Deliverable | Path |
|-------------|------|
| Tailored resume (.docx) | `resumes/tailored/<company>_<role>_<YYYY-MM-DD>.docx` |
| Tailored resume (.pdf) | `resumes/tailored/<company>_<role>_<YYYY-MM-DD>.pdf` |
| Change log (.md) | `resumes/tailored/<company>_<role>_<YYYY-MM-DD>_changes.md` |
| Cover letter (.docx) | `resumes/tailored/<company>_<role>_<YYYY-MM-DD>_cover.docx` |
| Cover letter (.pdf) | `resumes/tailored/<company>_<role>_<YYYY-MM-DD>_cover.pdf` |

**Change log format (mandatory):**
```
### Change #N — [Section Name]
**Before:** <original text>
**After:**  <new text>
**Why:**   <which JD requirement this addresses + keyword/ATS rationale>
```

**Resume rules:**
- Keep it 1–2 pages max (PMs in India: 2 pages OK above 4 yrs experience)
- Preserve facts; rephrase, reorder, and reweight only — never fabricate
- Embed JD's exact keywords where truthful (ATS optimization)
- Lead bullets with action verb + measurable outcome
- Keep dates, company names, titles, education facts unchanged
- Use clean ATS-friendly format (no tables, no columns, no images)

**Cover letter rules:**
- 250–350 words
- 3 paragraphs: hook (why this company), value (top 3 matches), close (next step)
- Match company tone (formal vs. startup-casual based on JD/website signals)

## Final Output Step
After producing files, push to a new branch (`tailored-<company>-<role>-<date>`) and create a PR. Always link the PR in the chat reply.

## Things to Always Verify
- LinkedIn profile titles/dates match resume — flag mismatches
- No exaggeration of seniority (Consultant II ≠ Senior PM)
- Cover letter doesn't restate resume — it should add narrative
