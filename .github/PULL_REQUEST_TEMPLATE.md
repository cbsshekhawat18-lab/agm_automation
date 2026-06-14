<!--
Thanks for the pull request! Please fill in the sections below.
See CONTRIBUTING.md for the full process and project rules.
-->

## Summary

<!-- One or two sentences: what does this PR change and why? -->

## Type of change

<!-- Tick all that apply. -->

- [ ] 🐛 Bug fix (no behaviour change for unaffected users)
- [ ] ✨ New feature (adds a new section, toggle, or capability)
- [ ] 🎨 Style / formatting change to the generated doc
- [ ] 📝 Wording change to existing legal / boilerplate paragraphs
- [ ] 🛠 Build / CI / release process
- [ ] 📚 Documentation only
- [ ] ⚠️ Breaking change (existing Master_Input files behave differently)

## Linked issue

<!-- e.g. "Closes #42" or "Refs #42" — required for non-trivial changes. -->
Closes #

## How it was tested

<!--
Describe the manual / automated test steps. Specifics please:

- Which OS did you build / run on?
- Which Master_Input file did you use (DEMO / EMPTY / sanitised real client)?
- For doc-formatting changes, did you visually compare the output against the reference PDF?
- For loader changes, did you also regenerate the output to confirm no regressions?
-->

## Screenshots / output diff

<!--
For any doc-affecting change, attach a before/after screenshot or paste the
relevant paragraph of the generated Final_Set.docx.
-->

## Pre-merge checklist

- [ ] I read [CONTRIBUTING.md](../blob/main/CONTRIBUTING.md)
- [ ] My commit messages follow the `type(scope): summary` convention
- [ ] I did **not** commit any real client data (master/Master_Input.xlsx and similar are gitignored)
- [ ] I did **not** add any network calls, telemetry, analytics, or cloud dependencies
- [ ] CI is green (`Build (macos-latest)` + `Build (windows-latest)` checks pass)
- [ ] If the PR changes the generated doc, I've eye-balled the output against the reference PDF
- [ ] If the PR adds a toggle / column / paragraph, I updated [USAGE.md](../blob/main/USAGE.md)

## Additional notes

<!-- Anything reviewers should know — design tradeoffs, follow-up work, edge cases not covered. -->
