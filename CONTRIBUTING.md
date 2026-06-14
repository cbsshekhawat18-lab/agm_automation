# Contributing to AGM Final Set

Thanks for taking the time to contribute! AGM Final Set is a free, open-source desktop tool for Indian CA / CS firms. Bug reports, ideas, doc improvements, and pull requests are all welcome.

> 📜 By participating, you agree to our [Code of Conduct](CODE_OF_CONDUCT.md).
> 🔒 Please don't post real client data in issues or PRs — use the demo file (`Master_Input_DEMO.xlsx`) or fully fictional values.

---

## Ways to contribute

| What | Where | How |
|---|---|---|
| Found a bug | [Bug report issue](https://github.com/cbsshekhawat18-lab/agm_automation/issues/new?template=bug_report.yml) | Pick the **Bug Report** template, include the steps and your platform |
| Want a feature | [Feature request issue](https://github.com/cbsshekhawat18-lab/agm_automation/issues/new?template=feature_request.yml) | Pick the **Feature Request** template, describe the use case |
| Doc / wording typo | Pull Request | Edit the `.md` file and open a PR |
| Code change | Pull Request | See **Development setup** below |
| Found a security issue | [Security advisory](https://github.com/cbsshekhawat18-lab/agm_automation/security/advisories/new) | Please **do not** open a public issue — use the private advisory flow. See [SECURITY.md](SECURITY.md) |
| Have a question | [Discussions](https://github.com/cbsshekhawat18-lab/agm_automation/discussions) | Use Discussions, not Issues, for questions |

---

## Development setup

```bash
git clone https://github.com/cbsshekhawat18-lab/agm_automation
cd agm_automation
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

Run the CLI against your own master file:

```bash
# Uses master/Master_Input.xlsx by convention (gitignored — your real data
# stays out of the repo). For testing, copy the demo file there:
cp sample/Master_Input_DEMO.xlsx master/Master_Input.xlsx
python generate_final_set.py
# Output: output/Final_Set.docx
```

Or use the wrapped entry that simulates the bundled-binary flow:

```bash
python app_main.py
```

### Building a binary locally

```bash
pip install pyinstaller

# macOS
./build_macos.sh           # dist/AGM_Final_Set-macos-dev.zip
VERSION=1.0.6 ./build_macos.sh

# Windows
build_windows.bat
```

CI builds both Win+Mac binaries automatically when a `vX.Y.Z` tag is pushed. See [.github/workflows/release.yml](.github/workflows/release.yml).

---

## Project layout — quick map

```
agm_automation/
├── app_main.py                 # bundled-binary entry point (PyInstaller)
├── generate_final_set.py       # source-mode CLI (existing)
├── agm_final_set.spec          # PyInstaller config
├── scripts/
│   ├── loader.py               # reads Master_Input.xlsx into a dict
│   ├── gen_final_set.py        # generates Final_Set.docx (one func per section)
│   ├── docx_helpers.py         # tables, headings, signing blocks, letterhead
│   ├── validate_master.py      # required-field + sanity validation
│   └── share_capital_text.py   # auto-builds capital descriptions
├── sample/                     # demo + empty templates shipped publicly
└── master/Master_Input.xlsx    # YOUR private working file (gitignored)
```

**Two important rules baked into the design:**

1. **Excel = data only.** Template text and branching logic live in Python. If you want to change wording in the output doc, edit `scripts/gen_final_set.py`. Don't add a column to the Excel.
2. **No network calls — ever.** This app runs 100% locally. Don't add telemetry, auto-update checks, or analytics. See [PRIVACY.md](PRIVACY.md).

---

## Style guide

- **Python 3.13+**, type hints where useful, no exotic dependencies (`python-docx` + `openpyxl` only).
- Keep functions small. One section of the doc per function in `gen_final_set.py`.
- Don't write comments that just restate what the code does. Do write comments for *why* (a workaround, a hidden constraint, a legal quirk).
- Existing legal / boilerplate paragraphs are mostly verbatim from the reference PDF. Don't paraphrase them — match the wording.
- If you change a default for branching text (e.g. swap a toggle's YES vs NO branch), say so explicitly in the PR description so reviewers can sanity-check against the PDF.

---

## Pull request process

1. **Fork → branch → push → PR.** Keep one PR per logical change.
2. Fill in the PR template (it'll auto-load when you open the PR).
3. **Don't commit your real client master file.** The `.gitignore` blocks `master/Master_Input.xlsx`, but check `git diff --cached` before pushing.
4. If your change affects the generated doc, regenerate `output/Final_Set.docx` from the DEMO file and visually compare to the reference PDF.
5. CI runs on every push. A green build is required for merge.
6. Squash-merge by default. Commit message should follow `type(scope): summary` (e.g. `fix(notice): align item numbers with sub-paragraphs`).

---

## Releasing (maintainer only)

```bash
git tag v1.0.7 -m "v1.0.7 — short summary"
git push origin v1.0.7
# CI builds both binaries and publishes a GitHub Release automatically (~5 min).
```

---

## License

By submitting a contribution, you agree your work will be released under the project's [MIT License](LICENSE).
