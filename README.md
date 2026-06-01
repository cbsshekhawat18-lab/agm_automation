# AGM Final Set Automation

Generates the **complete AGM Final Set** as one Microsoft Word document for an Indian Private Limited company from a single Excel master input.

The output (`output/Final_Set.docx`) contains, in order:

1. Notice of AGM (with Notes and Explanatory Statement)
2. Director's Report (28 numbered sections)
3. Form AOC-2 (Annexure A) — one block per related party
4. Equity List of Shareholders (landscape)
5. Attendance Sheet of Members in AGM
6. List of Directors as on FY end
7. Director's Attendance Sheet (landscape)
8. Certified True Copy — Director Regularization (one per regularized director)
9. Certified True Copy — Related Party Transactions (one per RPT)
10. Certified True Copy — Auditor Re-appointment
11. Intimation Letter to Auditor
12. Consent Letter cum Eligibility Certificate (auditor letterhead)
13. Details of Designated Person (RoC letter)

**Free to use** (MIT License). Source available, binaries on the [Releases page](#download).

---

## 🔒 Privacy & Confidentiality — Read This First

This application is designed for the strictest expectations of professional confidentiality. **Every CA / CS firm should be comfortable putting client data into it.**

- **100% offline.** The app runs entirely on your computer. There is no cloud component, no SaaS backend, no API call.
- **Zero data transmission.** The app makes **no network requests of any kind** — no telemetry, no analytics, no auto-update check, no crash reporting, no "phone-home".
- **No accounts, no login, no internet required.** Install, run offline, done. You can disconnect from Wi-Fi entirely and it still works.
- **Your files stay on your disk.** `Master_Input.xlsx` and `Final_Set.docx` live in the folder where you run the binary. Nothing is copied, synced, or backed up to anywhere.
- **No client data in the public release.** The downloadable binary ships with a fictional demo company (`ABC SAMPLE PRIVATE LIMITED`) — never any real Indian company's data.
- **Open source.** Every line of code that touches your data is in this repo. You can audit, fork, or self-build.
- **No license server.** This is MIT-licensed free software. There is no remote check that has to succeed for the app to work.

📄 **Full privacy policy:** [PRIVACY.md](PRIVACY.md) — what's read, what's stored, what's NOT done, permissions, retention, deletion.

If your firm's IT policy requires client data to stay on a single machine — this tool fits.

---

## Download

| Platform | File | Size |
|---|---|---|
| Windows | `AGM_Final_Set-windows-vX.Y.Z.zip` | ~15 MB |
| macOS   | `AGM_Final_Set-macos-vX.Y.Z.zip`   | ~14 MB |

Download from the [Releases](https://github.com/cbsshekhawat18-lab/agm_automation/releases) page. No Python install needed — the binary is self-contained.

---

## Quick Start

1. **Download** the zip for your platform, unzip anywhere (e.g. `Desktop/AGM/`).
2. **First run** — double-click `AGM_Final_Set` (Mac) or `AGM_Final_Set.exe` (Windows). The app copies a demo template called `Master_Input.xlsx` into the same folder and opens it.
3. **Fill the Excel** with your client's data (overwrite the demo values). Save.
4. **Second run** — double-click the app again. It reads your filled Master, generates `output/Final_Set.docx`, and opens it.
5. Open the doc in MS Word — review, paste a route-map image where the placeholder is, sign, and Save As PDF.

**Per-client tip:** copy the entire app folder once per client (e.g. `AGM_Acme/`, `AGM_Beta/`). Each gets its own `Master_Input.xlsx` and `output/`.

---

## What goes in the Excel?

Full reference in [USAGE.md](USAGE.md). Quick summary of sheets:

- `Company` — name, CIN, address, AGM details, signing directors, share capital
- `Directors` — every director on board during FY (set `Regularise?=YES` for Additional Directors being regularized)
- `Shareholders` — every shareholder as on FY end
- `BoardMeetings` — one row per board meeting
- `Financials` — revenue, expenses, taxes, PAT (current + previous year)
- `Toggles` — 28 Yes/No flags (e.g. `AuditorReappointment`, `HasSubsidiariesEtc`)
- `Auditor` — auditor details (only if reappointing)
- `RelatedParty`, `AuditorRemarks`, `MaterialChanges`, `EGMMeetings`, ... — tabular extras

Every cell maps to a specific paragraph in the doc. See [USAGE.md → Toggle Reference](USAGE.md#toggle-reference) for a full table of "if I change this cell, what changes in the doc".

---

## Running from source (developers only)

```bash
git clone https://github.com/cbsshekhawat18-lab/agm_automation
cd agm_automation
pip install -r requirements.txt
python generate_final_set.py        # uses master/Master_Input.xlsx by convention
# or:
python app_main.py                   # same logic, more user-friendly output
```

### Building a binary locally

```bash
pip install pyinstaller

# macOS
./build_macos.sh                     # produces dist/AGM_Final_Set-macos-dev.zip
VERSION=1.0.0 ./build_macos.sh       # tag a version

# Windows
build_windows.bat
```

### Project layout

```
agm_automation/
├── app_main.py                 # bundled entry (used by PyInstaller)
├── generate_final_set.py       # source-mode entry (existing CLI)
├── agm_final_set.spec          # PyInstaller config
├── scripts/
│   ├── loader.py               # reads Master_Input.xlsx
│   ├── gen_final_set.py        # builds the docx (one section per function)
│   ├── docx_helpers.py         # tables, signing blocks, letterhead
│   ├── validate_master.py      # required-field validation
│   └── share_capital_text.py   # auto-builds capital descriptions
├── master/Master_Input.xlsx    # your real master (gitignored if private)
├── sample/Master_Input_DEMO.xlsx  # bundled dummy data for public release
├── sample/build_demo.py        # regenerates the demo file
├── output/Final_Set.docx       # generated doc
├── build_macos.sh
├── build_windows.bat
├── .github/workflows/release.yml  # CI: build + release on tag push
├── LICENSE                     # MIT
├── README.md                   # this file
└── USAGE.md                    # detailed logic + toggle reference
```

---

## Releasing a new version

1. Bump version in your head (e.g. `v1.0.0`).
2. Tag and push:
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
3. GitHub Actions automatically:
   - Builds the macOS binary on a macOS runner
   - Builds the Windows binary on a Windows runner
   - Creates a GitHub Release named `v1.0.0`
   - Attaches both zips as downloadable assets
4. Share the release URL: `https://github.com/cbsshekhawat18-lab/agm_automation/releases/tag/v1.0.0`

No code signing yet, so users will see:
- **Windows**: SmartScreen "Unknown publisher" → click "More info" → "Run anyway"
- **macOS**: First launch needs right-click → "Open" → "Open" (Gatekeeper)

Add code signing later if you want to remove these warnings.

---

## Where to publish (for first-time release)

The cheapest, most professional path:

1. **Create a public GitHub repo** at `github.com/cbsshekhawat18-lab/agm_automation`.
2. **Push the code** — everything in this folder.
3. **Push a tag** (`git tag v1.0.0 && git push --tags`). The workflow builds both binaries and creates a public Release with download links.
4. **(Optional) Landing page** — turn on GitHub Pages from `Settings → Pages`. Point it at the `main` branch + `/docs` folder, or use a static site generator. Free.
5. **(Optional) Custom domain** — buy `agmfinalset.in` (~₹800/yr) and point a CNAME at the GitHub Pages site.

That's it. Anyone can now go to your Releases page and download the binary for their OS. No hosting fees, no payment gate, no app stores.

---

## License & Copyright

**© 2026 Chandrabhan Shekhawat · Gigai Kripa Services. All rights reserved.**

Released under the MIT License — see [LICENSE](LICENSE) for the full text. Free for personal and commercial use, modification, and redistribution; the only requirement is that you keep the copyright + license notice in any copy you distribute. No warranty.

## Contact

**Publisher:** Chandrabhan Shekhawat — Gigai Kripa Services
Bugs / feature requests / privacy questions: [open a GitHub issue](https://github.com/cbsshekhawat18-lab/agm_automation/issues).

## Contact

Bugs / feature requests: open a GitHub issue. Maintained by [@cbsshekhawat18-lab](https://github.com/cbsshekhawat18-lab).
