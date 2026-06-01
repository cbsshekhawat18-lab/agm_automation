# AGM Final Set — Usage & Logic Reference

This document explains the **workflow**, the **logic behind every section**, and **what each toggle changes in the output**. Read this once and you'll know exactly which cell in `Master_Input.xlsx` controls which line in `Final_Set.docx`.

> **Privacy note (CA / CS-grade):** This application runs entirely on your local machine. **No data ever leaves your computer.** No internet connection is used, no server is contacted, no telemetry is collected. Your client data is as confidential as anything else on your hard drive. See [README → Privacy & Confidentiality](README.md#-privacy--confidentiality--read-this-first) for the full statement.

## Table of Contents

1. [How to Use](#how-to-use)
2. [Workflow at a Glance](#workflow-at-a-glance)
3. [Sheets in Master_Input.xlsx](#sheets-in-master_inputxlsx)
4. [Sections of Final_Set.docx](#sections-of-final_setdocx)
5. [Toggle Reference (28 toggles)](#toggle-reference)
6. [Auto-Derived Behavior (no toggle needed)](#auto-derived-behavior)
7. [Validation Rules](#validation-rules)

---

## How to Use

### 1. Get the app
- **Mac**: download `AGM_Final_Set-macos-vX.Y.Z.zip` from [Releases](#), unzip, double-click `AGM_Final_Set`.
- **Windows**: download `AGM_Final_Set-windows-vX.Y.Z.zip`, unzip, double-click `AGM_Final_Set.exe`.

### 2. First run
The app sees there's no `Master_Input.xlsx` next to it, copies the demo template, opens it in Excel, and exits. **This is expected.** Fill the template with your client's data and save.

### 3. Second run
The app reads `Master_Input.xlsx` from its own folder, generates `output/Final_Set.docx`, and opens the doc automatically.

### 4. Per-client workflow
1. Make a copy of the app folder per client (e.g. `AGM_Acme/`, `AGM_Beta/`).
2. Edit each `Master_Input.xlsx` independently.
3. Run the app inside that folder. The doc lands in `output/Final_Set.docx` next to the binary.

---

## Workflow at a Glance

```
Master_Input.xlsx          loader.py            gen_final_set.py
┌────────────────┐        ┌─────────┐         ┌──────────────────┐
│ Company        │        │ load_   │         │ section_notice() │
│ Directors      │  ───►  │ master()├──data──►│ section_dr()     │
│ Shareholders   │        │         │         │ section_aoc2()   │
│ BoardMeetings  │        └─────────┘         │ ... 13 sections  │
│ Financials     │             │              └────────┬─────────┘
│ Toggles        │             ▼                       │
│ Auditor        │        ┌─────────┐                  ▼
│ RelatedParty   │        │validate │           Final_Set.docx
│ ...14 sheets   │        │_master()│                (output)
└────────────────┘        └─────────┘
```

**Key rule:** Excel = data only. Python = template text + logic. To change wording, edit Python. To change values, edit Excel.

---

## Sheets in Master_Input.xlsx

| Sheet | Required? | Purpose |
|---|---|---|
| `Company` | ✅ | Name, CIN, address, AGM details, signing directors, share capital, employees, RoC |
| `Directors` | ✅ | All directors on board during FY. `Regularise?=YES` marks Additional Directors to regularize |
| `Shareholders` | ✅ | All shareholders as on FY end |
| `BoardMeetings` | ✅ | One row per board meeting in the FY |
| `Financials` | ✅ | Revenue, expenses, taxes, PAT (current + previous year) |
| `Toggles` | ✅ | ~28 Yes/No applicability flags + one free-text (`FinancialUnit`) |
| `Auditor` | Conditional | Required only if `AuditorReappointment=YES` or `AuditorReportInDirReport=YES` |
| `RelatedParty` | Conditional | One row per RPT. Auto-flips `RPT_Required` |
| `AuditorRemarks` | Conditional | One row per remark. Auto-flips `AuditorRemarks` |
| `MaterialChanges` | Conditional | Material changes post-FY-end. Auto-flips `MaterialChangesPostFY` |
| `ShareCapitalChanges` | Conditional | Share-capital change descriptions. Auto-flips `ChangeInShareCapital` |
| `EGMMeetings` | Conditional | EGMs held during FY. Auto-flips `HeldEGM` |
| `BusinessNatureChanges` | Conditional | Business-nature change descriptions. Auto-flips `ChangeInBusinessNature` |
| `DirChanges` | Optional | Legacy free-form regularization entries (newer files use `Directors!Regularise?` instead) |

---

## Sections of Final_Set.docx

The doc is one continuous Word file with 13 logical sections. Each section is rendered by a function in [scripts/gen_final_set.py](scripts/gen_final_set.py).

### 1. NOTICE of AGM
- **Opening paragraph**: AGM number, members, date/day/time, venue. From `Company` tab.
- **ORDINARY BUSINESS**:
  - Item 1 (always): Adopt audited Balance Sheet for FY end date.
  - Item 2 (if `AuditorReappointment=YES`): Re-appoint auditor + resolution. Auditor details from `Auditor` tab.
- **SPECIAL BUSINESS** (only if regularizations or RPTs exist):
  - Items 3..N: One per `Directors!Regularise?=YES` row (director regularization resolution).
  - Items N+1..M: One per `RelatedParty` row (RPT approval resolution).
- **Closing**: Date / Place / single-signer block (First Signing Director).
- **NOTES**: 6 fixed legal notes (proxy, joint holders, inspection, etc.). Not editable from Excel.
- **EXPLANATORY STATEMENT under Section 102(1)**: Auto-included if special business exists. One paragraph per regularization, one block + table per RPT.
- **ROUTE MAP**: Placeholder page. User pastes map image after generation.

### 2. DIRECTOR'S REPORT (28 numbered sections)

This is the longest section. **§N below corresponds to the numbered heading in the doc.**

| § | Heading | What controls it |
|---|---|---|
| Opening | "Your directors have pleasure..." | `AGM Number`, `FY End Date` |
| 1 | Financial summary table | `Financials` tab + `FinancialUnit` toggle |
| 2 | Web address | `CompanyHasWebsite` toggle + `Company!Website` |
| 3 | Board Meetings | `BoardMeetings` tab |
| 4 | Directors' Responsibility Statement | Fixed (5 bullets) |
| 5 | Reporting of Frauds | `FraudsReported` toggle |
| 6 | Independent Director declaration | Fixed text (PDF-aligned) |
| 7 | Nomination & Remuneration Committee | Fixed text |
| 8 | Auditor's Remarks | `AuditorRemarks` toggle + tab |
| 9 | Loan/Guarantee under §186 | `Loan186Applicable` toggle (Yes/No values in table) |
| 10 | State of Company's affairs | Auto-branch on PAT sign (profit / loss) |
| 11 | Transfer to Reserves | `Financials!Amount Transferred to Reserves` |
| 12 | Dividend | `Financials!Dividend Amount` + PAT sign |
| 13 | Material changes post-FY | `MaterialChangesPostFY` + `MaterialChanges` tab |
| 14 | Risk management policy | `RiskMgmtPolicyExists` toggle |
| 15 | CSR | `CSRApplicable` toggle |
| 16 | Conservation of energy | Fixed NA table |
| 17 | Subsidiaries / JV / Associates | `HasSubsidiariesEtc` toggle |
| 18 | Rule 8(5) disclosures (6 sub-rows) | See breakdown below |
| 19 | Financial highlights paragraph | Auto-branch on PAT sign |
| 20 | Change in nature of business | `ChangeInBusinessNature` + `BusinessNatureChanges` tab |
| 21 | Directors and KMP | `ChangeInBoardComposition` toggle + per-regularization sentence |
| 22 | Deposits | Fixed (always "No deposit") |
| 23 | Significant orders by regulators | `SignificantOrdersByRegulators` toggle |
| 24 | Annual Evaluation | Fixed (always "Not applicable") |
| 25 | POSH (sexual harassment) | Fixed + complaint counts from `Company` tab |
| 26 | Maternity Benefit Act | Auto from Female employee count (override: `MaternityActApplicable`) |
| 27 | Number of Employees | Counts from `Company` tab |
| 28 | Others (13 sub-items) | See breakdown below |

**§18 Rule 8(5) breakdown:**

| Row | Controlled by |
|---|---|
| (i) Subsidiary status change | `SubsidiaryStatusChange` toggle |
| (ii) Independent Director opinion | Fixed (always "not required") |
| (iii) Internal Financial Controls | Fixed (PDF-aligned long statement) |
| (iv) Cost records | `CostRecordsApplicable` toggle |
| (v) Insolvency/Bankruptcy | `InsolvencyProceedings` toggle |
| (vi) One-time settlement valuation | `OneTimeSettlement` toggle |

**§28 Others breakdown:**

| # | Sub-item | Controlled by |
|---|---|---|
| 1 | Change of Name | `ChangeOfNameDuringYear` + `Company!Old Name`/`EGM Date` |
| 2 | Share Capital | `ChangeInShareCapital` + `ShareCapitalChanges` tab + Additional Capital Description |
| 3 | General Meetings (AGM + EGM) | `Company!Previous AGM Date` + `HeldEGM` + `EGMMeetings` tab |
| 4 | Auditors (Statutory / Secretarial / Cost) | `AuditorReportInDirReport`, `AuditorReappointment`, `SecretarialAuditorRequired`, `CostAuditorRequired` |
| 5 | Vigil Mechanism | `VigilMechanismRequired` toggle |
| 6 | Particulars of contracts (RPT) | `RPT_Required` toggle |
| 7 | Corporate Governance Certificate | Fixed (NA) |
| 8 | Management Discussion and Analysis | Fixed (NA) |
| 9 | Human Resources | Fixed |
| 10 | IEPF | Fixed |
| 11 | Secretarial Standards | Fixed |
| 12 | Designated Person | Fixed (uses `Designated Person Name/DIN` from Company tab) |
| 13 | Acknowledgements | Fixed |

- **Closing**: For and on behalf of BoD + two-signer block.

### 3. FORM AOC-2
- Section 1 (not at arm's length): Always 0, single empty Block-1 table for show.
- Section 2 (at arm's length): One Block-N table per `RelatedParty` row.
- Two-signer block.

### 4. LIST OF SHAREHOLDERS (landscape)
Table from `Shareholders` tab. Totals row at bottom.

### 5. ATTENDANCE SHEET OF MEMBERS in AGM
Empty signature columns; rows from `Shareholders` tab.

### 6. LIST OF DIRECTORS as on FY end
**Filtered**: only directors with `Appointment Date ≤ Current FY End Date`. Directors appointed after FY end appear in §21 trailing sentence instead.

### 7. DIRECTOR'S ATTENDANCE SHEET (landscape)
Grid: directors × board meeting dates. Attendance cells show whatever was typed in `BoardMeetings` (e.g. `P`, `NA`, blank).

### 8. CERTIFIED TRUE COPY — Director Regularization
One page per regularization (only if `RegulariseAdditionalDirector=YES`).

### 9. CERTIFIED TRUE COPY — Related Party Transactions
One page per RPT (only if `RPT_Required=YES`).

### 10. CERTIFIED TRUE COPY — Auditor Re-appointment
Single page (only if `AuditorReappointment=YES`).

### 11. INTIMATION LETTER to Auditor
On company letterhead. Only if `AuditorReappointment=YES`.

### 12. CONSENT LETTER cum Eligibility Certificate
On **auditor's letterhead** (placeholder). Only if `AuditorReappointment=YES`.

### 13. DETAILS OF DESIGNATED PERSON (RoC letter)
Only if `FirstDesignatedPersonDeclaration=YES`. Address from `Company!RoC Address`.

---

## Toggle Reference

Every toggle lives on the `Toggles` tab. **All values normalize to `YES` / `NO`** (also accepts `Y`/`N`/`TRUE`/`FALSE`/`1`/`0`). Missing or blank = `NO`.

The columns below mean:
- **Default** → what the doc shows if the toggle is `NO` (or missing).
- **When YES** → what changes when you set it to `YES`.
- **Where in doc** → which section is affected.

### Auditor-related

| Toggle | Default (NO) | When YES | Where in doc |
|---|---|---|---|
| `AuditorReappointment` | Notice has no auditor item, no auditor resolution, no intimation, no consent letter | Adds: Notice item 2, DR §28(4) Statutory paragraph, Resolution doc page, Intimation Letter page, Consent Letter page | Notice, DR §28(4), Resolution, Intimation, Consent |
| `AuditorReportInDirReport` | DR §28(4) Statutory section omitted | DR §28(4) shows "Pursuant to §139..." paragraph naming the firm + FRN | DR §28(4) |
| `AuditorRemarks` | DR §8 says "no observations, qualifications" | DR §8 shows a table of remarks (rows from `AuditorRemarks` tab) | DR §8 |
| `SecretarialAuditorRequired` | DR §28(4) says "not applicable" | DR §28(4) says Company has appointed a Secretarial Auditor | DR §28(4) |
| `CostAuditorRequired` | DR §28(4) says Cost Auditor "not applicable" | DR §28(4) says Company has appointed a Cost Auditor | DR §28(4) |

### Director / Board

| Toggle | Default (NO) | When YES | Where in doc |
|---|---|---|---|
| `RegulariseAdditionalDirector` | No special-business item for regularization, no Resolution page | Notice gets one Special Business item per Director with `Regularise?=YES`; DR §21 appends a sentence per regularization; one Resolution page per regularization | Notice, DR §21, Resolution |
| `ChangeInBoardComposition` | DR §21 says "no change in composition" | DR §21 says "there is change in composition" | DR §21 |

### Related Party

| Toggle | Default (NO) | When YES | Where in doc |
|---|---|---|---|
| `RPT_Required` | No RPT items in Notice, AOC-2 §2 count = 0 (no block tables), no RPT Resolution page, DR §28(6) says "not entered into" | One Notice item per `RelatedParty` row, one Explanatory Statement block per row, AOC-2 §2 shows one Block per row, DR §28(6) references Annexure-A, one Resolution page per row | Notice, Explanatory Statement, AOC-2, DR §28(6), Resolution |

### Section-specific applicability (DR)

| Toggle | Default (NO) | When YES | Where in doc |
|---|---|---|---|
| `CompanyHasWebsite` | DR §2 says "doesn't have any website" | DR §2 prints the website URL from `Company!Website` | DR §2 |
| `FraudsReported` | DR §5 says "no frauds were reported" | DR §5 says "details of fraud as follows: (Specify the Frauds...)" | DR §5 |
| `Loan186Applicable` | All 3 §9 questions = "No"; row 4 = "NA" | All 3 §9 questions = "Yes" | DR §9 |
| `MaterialChangesPostFY` | DR §13 says "no material changes" | DR §13 shows a table of changes (rows from `MaterialChanges` tab) | DR §13 |
| `RiskMgmtPolicyExists` | DR §14 says "no such elements of risk" | DR §14 says Company has implemented risk policy | DR §14 |
| `CSRApplicable` | DR §15 says "(Not Applicable)" | DR §15 says "CSR Committee constituted" | DR §15 |
| `HasSubsidiariesEtc` | DR §17 says "no subsidiary, JV or associate" | DR §17 shows a (placeholder) table | DR §17 |
| `SubsidiaryStatusChange` | §18(i) says "no company become or ceased" | §18(i) shows placeholder for company names | DR §18(i) |
| `CostRecordsApplicable` | §18(iv) says "not applicable" | §18(iv) says cost records are maintained | DR §18(iv) |
| `InsolvencyProceedings` | §18(v) says "no application/proceeding" | §18(v) shows placeholder for IBC details | DR §18(v) |
| `OneTimeSettlement` | §18(vi) says "no one-time settlement" | §18(vi) shows placeholder for OTS valuation difference | DR §18(vi) |
| `ChangeInBusinessNature` | DR §20 row: "No" + "NA" | DR §20 row: "Yes" + descriptions from `BusinessNatureChanges` tab | DR §20 |
| `SignificantOrdersByRegulators` | DR §23 says "no significant order" | DR §23 shows placeholder for order details | DR §23 |
| `MaternityActApplicable` | DR §26 says "Company does not have any female employee" | DR §26 says Company complies with the Act | DR §26 |
| `ChangeOfNameDuringYear` | DR §28(1) says "Company has not changed its name" | DR §28(1) shows the name change with old name and EGM date | DR §28(1) |
| `ChangeInShareCapital` | DR §28(2) says "no change in Share Capital" | DR §28(2) says "there has been a change" + appends `ShareCapitalChanges` rows | DR §28(2) |
| `HeldEGM` | DR §28(3)b says "had not held any EGM" | DR §28(3)b says EGMs were held + table from `EGMMeetings` tab | DR §28(3)b |
| `VigilMechanismRequired` | DR §28(5) says "not required" | DR §28(5) says "established the vigil mechanism" | DR §28(5) |
| `FirstDesignatedPersonDeclaration` | Last page (RoC Designated Person letter) NOT included | Adds the RoC letter as the final page | Last page |

### Free-text toggles

| Toggle | Values | Effect |
|---|---|---|
| `FinancialUnit` | `Thousand` / `Hundred` / `Lakh` | Inserted into "(In ___)" in DR §1, §10, §19. Default `Thousand` |

---

## Auto-Derived Behavior

These don't need a toggle — the loader sets them YES automatically when the data warrants:

| Behavior | Trigger |
|---|---|
| `RegulariseAdditionalDirector` auto-set YES | Any `Directors!Regularise?=YES` |
| `RPT_Required` auto-set YES | Any row in `RelatedParty` tab |
| `AuditorRemarks` auto-set YES | Any row in `AuditorRemarks` tab |
| `MaterialChangesPostFY` auto-set YES | Any row in `MaterialChanges` tab |
| `HeldEGM` auto-set YES | Any row in `EGMMeetings` tab |
| `ChangeInBusinessNature` auto-set YES | Any row in `BusinessNatureChanges` tab |
| `ChangeInShareCapital` auto-set YES | Any row in `ShareCapitalChanges` tab |
| Maternity branch flips to "complies" | `Female Employees Count > 0` (overrides toggle if set to NO) |
| Director list "as on FY end" filter | Excludes directors with `Appointment Date > Current FY End Date` |
| §21 trailing sentence per regularization | "After the end of financial year" if appt > FY end; else "During the year" |
| Capital descriptions | Auto-built from numeric inputs if Description field is blank |
| AGM Day / AGM Date in Words | Auto-derived from AGM Date if blank |
| PBT / PAT | Computed from Revenue, Expenses, Taxes if blank in Financials |
| State of affairs branch (§10) | "Profit" vs "Loss" wording chosen from PAT sign |
| §19 wording | "Net Profit after Tax (PAT)" vs "Net Loss" chosen from PAT sign |
| §12 Dividend wording | "do not propose" vs "net profit of Rs..." chosen from PAT sign |

---

## Validation Rules

`validate_master.py` runs before generation and **hard-stops** if anything required is missing.

### Required (always)
- Every label on `Company` and `Auditor` tabs ending with `*` must be filled (unless it's an auto-derived field like AGM Day).
- At least one row in `Directors`, `Shareholders`, `BoardMeetings`.

### Conditionally required
- `Auditor` tab is required if `AuditorReappointment=YES` or `AuditorReportInDirReport=YES`.
- `RelatedParty` rows required if `RPT_Required=YES`.
- Either `Directors!Regularise?=YES` or `DirChanges` rows required if `RegulariseAdditionalDirector=YES`.

### What's NOT validated (will silently propagate)
- Date sanity (appointment > cessation, AGM before FY end, etc.) — code accepts whatever the sheet says.
- Toggle/data contradictions (e.g. Maternity NO with Female=5) — auto-overridden where possible, otherwise rendered as-is.
- Internal contradictions inside one row (e.g. Date of Appointment in column E mentions one date while Description in column H quotes another).

Edit the Excel cells to fix the data; the code respects whatever you provide.

---

## Tips

- **One Master file per client.** Keep a separate copy of the app folder per company so each has its own `Master_Input.xlsx` and `output/`.
- **Re-run is safe.** Each run creates a NEW timestamped `Final_Set_YYYY-MM-DD_HH-MM-SS_IST.docx` in `output/` — previous outputs are preserved.
- **Logic changes** (e.g. wording tweaks) require editing Python in `scripts/gen_final_set.py`. Re-build the binary after editing.
- **Schema changes** to `Master_Input.xlsx` need matching loader updates in `scripts/loader.py`.

---

## License & Copyright

**© 2026 Chandrabhan Shekhawat · Gigai Kripa Services. All rights reserved.**

Released under the MIT License — see [LICENSE](LICENSE) for the full text. Full privacy policy at [PRIVACY.md](PRIVACY.md).
