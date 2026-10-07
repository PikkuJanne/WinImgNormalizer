# Release candidate approval proposal — ready for owner review; publication pending

Candidate implementation and proposed tag target: **8edbcbaeb3425ec3a52eeafde212c32553755af1**.
Application version/proposed tag: **1.0.0 / v1.0.0**. The candidate is unreleased and
unsigned. This is a concrete proposal ready for owner review; publication remains
pending and this document records no permission or readiness to publish.

| Proposed public-release attachment | Actual bytes | SHA-256 |
|---|---:|---|
| WinImgNormalizer-1.0.0-portable.zip | 184659 | 251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d |
| build-provenance.json | 763 | 901598504d3bb60cdfa4a236873a07d3a7287cb8ff6b3c715b202b39865cefae |
| SHA256SUMS.txt | 190 | 2257dbcd00707a93f3e811c9d3ff0eafb59a2e65bb756c440e152f1642540497 |

The selected PS7 ZIP, both sidecars, seven exact Git inputs, generated manifest and
MIT/CC0 licenses are bound in [candidate evidence](evidence/M4-T05.json). Fresh
ordinary-account PS7/PS5.1 builds and selected-artifact extraction/native synthetic
media checks passed. Same-host archive/sidecar equality and cross-host decoded file
parity were measured; Framework and current .NET ZIP framing may differ. Matching
BAT checks establish invalid-source setup/native exit with closed stdin; extracted
PS1 media uses owned OutputParent. No physical drag, real Pictures, video decoding
or new subjective image-quality consent is implied. [Exact candidate CI](evidence/M4-T05-ci.json)
and desktop normal/control/analyzer results support the manually audited T074
proposal. All 18 requirements are reconciled with no unresolved required coverage.

The proposed operations require explicit owner approval naming this candidate,
artifact hash and the externally observed exact final PR URL/head:

1. Merge the successor draft PR containing this record-only feature checkpoint
   into main. Its actual PR identity and full final head SHA must be independently
   observed after commit/push and supplied with the approval request; these files
   cannot embed their own future checkpoint SHA or authenticate that merge.
2. Create the new tag v1.0.0 at the exact candidate commit above, so the selected
   ZIP's embedded source revision agrees with its tag. Before acting, verify no
   existing conflicting tag/release and verify the actual main/PR/source state.
3. Publish a public GitHub Release for that approved tag, attaching the exact
   three files in the table only after actual post-merge/tag source, license,
   manifest, provenance and SHA-256 verification. If source/tag/asset identity
   changes, rebuild/smoke/rehash as needed and obtain approval for the new identity.

Merge, tag and release authorization are **pending**; website deployment is
**not granted**. No operation above has been executed by M4-T05. The owner already
merged PR25 before this task; its dated main/candidate tree relation remains a
separate starting observation, not a merge performed by this audit.

M4-T06/T075 must verify the actual approved tag/source, attached ZIP/sidecars and
publicly downloaded bytes before changing release/download projections. The
selected ZIP's package manifest, build provenance, GETTING_STARTED.md and
CHANGELOG.md retain their preparation-time unreleased wording. Actual publication
parity and this wording must be reviewed under T075 before release. The
existing release/website metadata and Test-WebsiteHandoff.ps1 strictly accept draft
unpublished state; M4-T06 must review any required state-projection/guard changes
and relevant tests rather than assume today's draft gate validates publication.
The metadata generator and builder also deliberately support only draft preparation.
Publication-state guide/metadata/guard edits or changed tag/source/ZIP identity
require relevant tests, rebuild/smoke/recomputed SHA-256 as needed, and new exact
owner approval. They invalidate approval for this specifically named prepared
artifact; it cannot be relabeled as a containing evidence checkpoint or new source.
No public-document change is made by M4-T05. No
provider/domain/framework is supplied and no website deployment is approved.
Screenshots/comparisons remain empty; private owner images have no public reuse
approval. Version/publication fields remain unchanged and downloads unavailable.

Historical D37/T067 workflow/appearance acceptance remains **5a32fc6533c6f9fae1d36c271b515d1d1dd4d8ad**. It does
not supply candidate publication consent. Windows 10 remains untested and live
UNC remains owner-excluded D35. No verified signature or publisher authentication
claim is supplied by these checksums. Final record-only checkpoint sync/CI and the
successor PR are reported externally after commit and must be rechecked before
owner-approved operations. This proposal does not start M4-T06.
