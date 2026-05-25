# Resume Tailoring Workflow

**Owner:** Srikanth Chintala
**Trigger:** Whenever the user pastes a Job Description (JD) or shares a job posting URL.

## Standard Operating Procedure

When the user shares a JD, always perform these three steps **in this order** and produce the outputs described.

### Step 1 — Senior Recruiter Evaluation (Gap Analysis)

Act as a senior recruiter and evaluate the user's current resume against the JD.

**Required research before evaluation:**
- Visit the hiring company's official website (About, Careers, Blog, Engineering pages where relevant)
- Cross-reference the JD on other portals (LinkedIn Jobs, Indeed, Glassdoor, company-native ATS pages)
- Look up current employees in similar roles on LinkedIn — sample 3–5 profiles to identify the keywords, tools, certifications, and outcome patterns the company actually values
- Note any company-specific language, values, or tech stack hints

**Deliverables:**
- A keyword/skill match matrix: JD requirement → resume evidence → match status (Strong / Partial / Missing)
- A list of explicit gaps (skills, tools, certifications, outcome metrics, domain language)
- A list of phrasing/positioning improvements (not just missing skills, but how existing experience is framed)
- ATS keyword density observation (which JD keywords are under-represented in the resume)

### Step 2 — Complete Report Analysis

Produce a structured markdown report containing:
1. JD summary — role, level, must-haves, nice-to-haves, company context
2. Candidate fit score (0–100) with brief justification
3. Keyword match matrix (table)
4. Top 5 gaps with severity (High / Medium / Low)
5. Top 5 strengths to emphasize
6. Recommended changes (per resume section: Summary, Experience, Skills, Projects, Certifications, Education)
7. Suggested STAR/quantified bullet rewrites (before → after)
8. Interview prep hot-spots (likely questions based on the gaps)

Save the report into the repo at: `reports/<company>_<role-slug>_<YYYY-MM-DD>.md`

### Step 3 — Tailored Resume + Change Log

Modify the canonical resume and produce:
- A new tailored resume file in **.docx** format under `resumes/tailored/<company>_<role-slug>_<YYYY-MM-DD>.docx`
- A markdown change log next to it: `resumes/tailored/<company>_<role-slug>_<YYYY-MM-DD>_changes.md`

The change log must list every change with this format:
- **Section:** <Summary | Experience | Skills | etc.>
- **Before:** <original text>
- **After:** <new text>
- **Why:** <which JD requirement / recruiter insight this addresses>

Never overwrite the canonical resume — always create a new tailored copy.

## Defaults (override per request)

- **Resume length:** 1 page if <8 years experience, 2 pages otherwise
- **Style:** ATS-friendly (no tables, no graphics, single column, standard fonts)
- **Tone:** Outcome-first bullets using STAR with quantified results
- **Tense:** Past tense for previous roles, present tense for current role
- **File naming:** `<company>_<role-slug>_<YYYY-MM-DD>` lowercase, hyphenated

## What to ask the user when info is missing

- Which resume file is the canonical source of truth (if multiple versions exist)
- Target geography and visa status (affects keyword choices)
- Seniority preference if the JD spans a band
- LinkedIn profile URL (to cross-reference and keep aligned)
- Whether to also produce a tailored LinkedIn About / cover letter

## Out of scope unless asked

- Cover letter generation
- LinkedIn profile rewrite
- Interview coaching beyond the prep hot-spots in the report
