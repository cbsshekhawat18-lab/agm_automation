# AGM Final Set — Privacy Policy

How **AGM Final Set** handles your data.

AGM Final Set is a desktop application that generates the complete AGM Final Set as one Microsoft Word document for an Indian Private Limited company — **running entirely on your computer**. This page explains exactly what data the application reads, where it lives, and what we do (and don't) do with it.

**Effective date:** June 1, 2026
**Publisher:** Chandrabhan Shekhawat — Gigai Kripa Services
**Contact:** [bug reports / privacy questions — see end of page]

---

## 🔒 100% on your device. Nothing uploaded. Ever.

AGM Final Set reads a single Excel file (`Master_Input.xlsx`) that you place next to the application, and writes a single Word file (`Final_Set.docx`) to a folder next to the application. **Your client data never leaves your computer.** There is no AGM Final Set server. There are no analytics. There are no third-party AI calls. There are no automatic update checks. The application makes **no network requests of any kind**.

---

## Single purpose

AGM Final Set's single purpose is to **help Indian CA / CS professionals generate the statutory AGM Final Set document** — Notice of AGM, Director's Report, Form AOC-2, Shareholder list, Director list, Attendance sheets, Resolutions, Intimation Letter, Consent Letter, and Designated Person letter — from a single Excel master input.

Every feature in the application serves this purpose. We do not bundle unrelated functionality.

---

## What AGM Final Set does

- Reads `Master_Input.xlsx` from the folder containing the application, **only when you launch the application**.
- Validates the data (required fields, date sanity, internal contradictions) and shows a structured report on the console.
- Generates a timestamped `Final_Set_YYYY-MM-DD_HH-MM-SS_IST.docx` in the `output/` subfolder of the application folder.
- Opens the generated document in your operating system's default Word viewer.
- Ships with two starter templates — `Master_Input_DEMO.xlsx` (pre-filled with fictional sample data) and `Master_Input_EMPTY.xlsx` (blank). You pick whichever you want to use.

That's it. The application exits when generation is complete.

---

## What AGM Final Set does **not** do

- **Send any of your data to any server** — there is no AGM Final Set backend.
- **Use cloud AI providers** — there is no AI in this product at all. Generation is deterministic Python code reading your Excel.
- **Make any network requests** — no telemetry, no crash reporting, no analytics endpoint, no feature-flag service, no auto-update check, no license-validation call.
- **Modify or delete files outside its own folder** — the application reads from and writes to its own folder only.
- **Read any file other than `Master_Input.xlsx`** — no scanning, no indexing, no other access.
- **Track you across the web** — the application is a self-contained binary; it has no browser, no embedded webview, no remote resources.
- **Sell, share, license, or transfer user data to any third party** — we never see your data, so we have nothing to sell.
- **Use user data to determine creditworthiness or for lending purposes.**
- **Use user data for advertising, retargeting, or any purpose unrelated to the single purpose stated above.**
- **Allow human review of user content** — no employee, contractor, or third party can see your local data.

---

## What data is stored locally

Everything AGM Final Set touches lives on **your computer's disk**, in the folder where you placed the application. It is never synced, uploaded, or transmitted. The table below covers all files involved.

| File | What it holds |
|---|---|
| `Master_Input.xlsx` | Your client's AGM data — company name, CIN, address, AGM date/venue, directors, shareholders, board meetings, financials, auditor details, related party transactions, applicability toggles. **You create and own this file.** |
| `output/Final_Set_*.docx` | Each generation produces a new timestamped Word document with the complete AGM final set. The application never deletes old outputs — you control retention. |
| `Master_Input_DEMO.xlsx` *(shipped)* | A fictional sample (ABC SAMPLE PRIVATE LIMITED) so you can try the app without preparing client data first. |
| `Master_Input_EMPTY.xlsx` *(shipped)* | A blank template you can rename and fill from scratch. |

There is no `chrome.storage`, no AppData, no `~/Library` data, no registry entry. The application reads and writes only inside its own folder.

To wipe everything: delete the application folder. That's the complete uninstall.

---

## Permissions & why we need them

This is a desktop binary. It uses only standard operating-system file I/O permissions; no special grants are requested.

| Permission | Why AGM Final Set needs it |
|---|---|
| Local file read | To read `Master_Input.xlsx` from the application's own folder. |
| Local file write | To write `output/Final_Set_*.docx` into the `output/` subfolder of the application's own folder. |
| Open a file | When generation is done, the application invokes the OS default viewer (`open` on macOS, `start` on Windows) to display the generated `.docx`. |

The application requests **no network access**, **no microphone**, **no camera**, **no contacts**, **no location**, **no system file scanning**, and **no admin/root**.

---

## Network calls

The application makes **zero network calls**. You can verify this two ways:

1. **Run it offline.** Disconnect from Wi-Fi entirely and the application works exactly the same.
2. **Audit the source.** The full source ships with the binary release and is published in the public GitHub repository. Search for `urllib`, `requests`, `http`, `socket`, `urlopen` — there are no outbound calls anywhere in the application logic.

There is no telemetry, no crash reporting, no analytics endpoint, no feature-flag service, no cloud LLM endpoint, no auto-update check, and no license server.

---

## Data retention & deletion

AGM Final Set keeps your data on your device for as long as the application folder exists. You can wipe it at any time:

- **Delete the generated documents** — remove files from the `output/` folder.
- **Delete the master input** — remove `Master_Input.xlsx` from the application folder.
- **Remove the application entirely** — delete the application folder. Wipes everything related to AGM Final Set.

Because there is no server, there is no separate copy of your data for us to delete, and no server-side request to file. **Local removal is total.**

---

## Limited Use of user data

In plain English: **AGM Final Set reads your `Master_Input.xlsx` only to generate the `Final_Set.docx` described above.** It is never transferred to anyone (including us — there is no "us" with a server), never used for ads, and never reviewed by a human.

The application performs no network communication of any kind. The only data movement is **disk → application memory → disk**, all on your computer.

---

## Open source — verify everything yourself

The full source code is published under the MIT License at [github.com/cbsshekhawat18/agm_automation](https://github.com/cbsshekhawat18/agm_automation). You can:

- Read every line that touches your data.
- Build the binary yourself from source instead of downloading the release.
- Fork it, modify it, audit it.

You don't have to trust us. You can verify.

---

## Changes to this policy

We may update this policy when the application gains new file-access behavior or new features. The **Effective date** at the top of this document reflects the most recent revision. Material changes (anything that affects what data is read, stored, or transmitted) will be called out in the next version's release notes.

---

## Contact, bug reports & privacy questions

**Publisher:** Chandrabhan Shekhawat — Gigai Kripa Services

For bug reports, feature requests, or privacy questions, open an issue at the project's GitHub repository: [github.com/cbsshekhawat18/agm_automation/issues](https://github.com/cbsshekhawat18/agm_automation/issues).

We don't operate a server — there's no support ticket system, no data-export endpoint, and no account to suspend. Because all data stays on your device, we have no copy to give you on request and nothing to delete server-side. Delete the application folder to wipe everything locally.

---

© 2026 Chandrabhan Shekhawat · **Gigai Kripa Services**. All rights reserved.

AGM Final Set is released under the MIT License. Source ships with the release package and is published on GitHub for reviewability.
