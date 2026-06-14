# Security Policy

## TL;DR

Found a security issue in AGM Final Set? **Please don't open a public issue.** Use the private reporting channel below so it can be fixed before bad actors learn about it.

---

## Reporting a vulnerability

**Preferred — Private security advisory on GitHub:**
👉 https://github.com/cbsshekhawat18-lab/agm_automation/security/advisories/new

Only the maintainer can see private advisories. You can also collaborate with the maintainer on a fix inside the advisory before disclosure.

**Alternative — direct e-mail:**
If you can't use GitHub advisories, e-mail the maintainer (Chandrabhan Shekhawat — Gigai Kripa Services). The address is in the repo's git history (`git log --format='%ae' | sort -u`).

**Please include in your report:**

- A clear description of the issue and why it's a security concern (not just a bug)
- Steps to reproduce — exact OS, app version, sample data (use the demo file, **never real client data**)
- Impact assessment (what an attacker could do)
- Any suggested fix or workaround you've already explored

You'll typically get an acknowledgement within **3 business days**. We'll work on a fix in a private branch and coordinate disclosure with you.

---

## What's in scope

This project is a **desktop binary that runs 100% locally**. There's no server, no API, no auth, no cloud. The realistic attack surface is small and includes:

| In scope | Example |
|---|---|
| Code execution / command injection via the Excel input | A crafted `Master_Input.xlsx` that causes the binary to execute arbitrary code beyond docx generation |
| Path traversal in the loader | A crafted input that causes reads/writes outside the app folder |
| Document-based attacks against the output `.docx` | Generation produces a file that exploits Word/LibreOffice on open |
| Dependency vulnerabilities | A CVE in `python-docx`, `openpyxl`, or PyInstaller bootloader that affects this binary |
| Bundle integrity / supply-chain | Tampering with the GitHub Actions workflow or release artifacts |
| Privacy claims being violated | Any way the app could send data over the network (it shouldn't, ever — see [PRIVACY.md](PRIVACY.md)) |

## What's NOT in scope

- **Bugs that produce wrong AGM content but no security impact.** Open a normal [bug report](https://github.com/cbsshekhawat18-lab/agm_automation/issues/new?template=bug_report.yml) instead.
- **Antivirus / SmartScreen false positives** on unsigned binaries. These are a known issue with all unsigned PyInstaller executables; we'll address this if/when we buy code-signing certificates.
- **Issues that require physical access to the user's machine.** The threat model assumes an honest user on their own hardware.
- **Theoretical issues without a concrete exploit.** A working PoC against the current release is required.

---

## Supported versions

We support **the latest released version only**. There's no LTS branch — when a fix ships, it goes into the next release on the main branch.

| Version | Supported |
|---|---|
| Latest `vX.Y.Z` on the [Releases page](https://github.com/cbsshekhawat18-lab/agm_automation/releases/latest) | ✅ |
| Older versions | ❌ — please update |
| Local fork / source build | Best-effort, depends on whether you've changed the code that's affected |

The current latest version is always at:
👉 https://github.com/cbsshekhawat18-lab/agm_automation/releases/latest

---

## Coordinated disclosure

When you report an issue:

1. **Day 0:** You file the advisory.
2. **Within 3 business days:** Maintainer acknowledges.
3. **Negotiated window (typically 7–30 days):** We work on a fix together inside the private advisory. You can review the patch.
4. **On release:** We publish a new tagged version with the fix and a security note in the release notes. You're credited (unless you'd rather be anonymous — just tell us).
5. **CVE:** GitHub auto-issues a CVE if the advisory warrants one.

We commit to crediting researchers for valid reports. No bug-bounty programme right now (this is a free tool maintained by one person), but a public thank-you in the release notes is guaranteed.

---

## Thanks

To everyone who reports issues responsibly — your work makes this tool safer for every CA / CS firm using it. Thank you.
