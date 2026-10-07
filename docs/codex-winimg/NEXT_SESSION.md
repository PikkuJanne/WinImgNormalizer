# Next session handoff

**M4-T06 is awaiting_owner; T075 is not_run.** Progress remains 27/28 tasks and
74/75 cases. Windows 10 testing is explicitly **out of project scope (D43)**;
do not request it or treat it as an acceptance/release blocker.

Read [publication evidence](evidence/M4-T06.json) and the
[exact remaining tag/public-release proposal](RELEASE_CANDIDATE.md).
Owner PR26 merge is already complete at 94dff68d9474e0d0218beba81353c699fc7a6c4b
(head fd6cd26dea4747b963313f9d9f11ffd3189b4458, merged 2026-10-07T05:01:46Z).
Do not request or execute that merge again. Reverify live clean feature/main,
canonical remote, current tags/releases and the containing checkpoint's actual
post-commit SHA/PR/CI; this file cannot embed its own future identity.

The remaining approval must name new tag **v1.0.0** at candidate
**8edbcbaeb3425ec3a52eeafde212c32553755af1** and a **public GitHub Release**, attaching
WinImgNormalizer-1.0.0-portable.zip (**184659 bytes**, SHA-256
**251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d**) plus exact build-provenance.json/SHA256SUMS.txt.
Task selection and a testing exclusion do not record exact publication permission.

Before approved publication, verify the actual tag/source/license/ZIP/manifest and
sidecars. Review preparation-time unreleased guide/changelog wording and necessary
draft-only builder/generator/website projection/guard transitions. Any changed
source/tag/artifact identity requires relevant tests, rebuild/smoke/rehash as needed
and new precise owner approval. Verify real public download bytes before updating
download metadata. Current draft gates cannot prove published-state consistency.
Website deployment/public owner-media reuse remain unapproved; no provider context
exists. Preserve D35/D37 and the separate WinImgNormalizer-main snapshot.

Use the clean feature checkout C:/projects/WinImgNormalizer; no reset, clean, stash,
rebase, force push, silent pull, default-branch mutation or subsequent task is implied.
