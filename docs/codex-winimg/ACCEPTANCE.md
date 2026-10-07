# Acceptance ledger

| Gate | State | Evidence |
|---|---|---|
| M0 preparation | accepted historically; T001-T004 characterize known legacy defects; T005-T007 verify import/invocation/failure detection at 3b80ee94b8471a22560ea90a7583283c1613195c | [Characterization](evidence/M0-T02.json), [runner/import](evidence/M0-T03.json) |
| M1 safe output | accepted automated combined T028 gate at 4d7f9f2e7681902971b584560211139d555bf011; preceding corrected cases replayed in current candidate | [M1 gate](evidence/M1-T06.json), [current candidate](evidence/M4-T05.json) |
| M2 correct conversion | accepted automated combined T046 gate at 11ee247154a00d79aaa49a438d1f526457b22fdf; preceding conversion cases replayed in current candidate | [M2 gate](evidence/M2-T06.json), [current candidate](evidence/M4-T05.json) |
| M3 reliable Windows workflow | accepted automated T064-T066 gate at 17a405bfe0150607486248e845f93ba82e1d0d52; current 27-suite/67-case corpus passes at 8edbcbaeb3425ec3a52eeafde212c32553755af1 | [M3 gate](evidence/M3-T06.json), [current desktop](evidence/M4-T05.json), [current CI](evidence/M4-T05-ci.json) |
| M3 owner workflow/quality | accepted; T067 passed at 5a32fc6533c6f9fae1d36c271b515d1d1dd4d8ad; this remains narrow historical consent | [Direct owner acceptance](evidence/M3-T07-acceptance.json), [fixed Windows smoke](evidence/M3-T07-recheck-fixed-launcher.json), [validated repair](evidence/M3-T07-repair.json) |
| M4 release candidate | preparation accepted; manual T074 proposal names candidate 8edbcbaeb3425ec3a52eeafde212c32553755af1 and actual ZIP SHA-256 251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d; publication explicitly approved and completed | [Candidate audit](evidence/M4-T05.json), [candidate CI](evidence/M4-T05-ci.json), [approved publication record](RELEASE_CANDIDATE.md) |
| Merge/tag/release publication | owner approved exact v1.0.0 tag/source/three public assets; actual T075 passed; local main safely fast-forwarded to owner PR27 merge ccbf6626 | [Publication](evidence/M4-T06-publication.json), [desktop](evidence/M4-T06-validation.json), [CI](evidence/M4-T06-publication-ci.json) |
| Website deployment | not authorized; no provider context or approved public captures | [Held website preparation](evidence/M4-T04.json) |

The M0-M3 rows reconcile existing evidence; no new historical run or owner consent
is inferred. T074 remains a manual owner-kind proposal/audit case, with M4-T05
owner_gate false. The automated inventory remains 27 suites and 67 required IDs;
T001-T004 are characterization, T007 is the real failure control and T067 remains
the earlier named owner acceptance. Task progress is 28/28 and specified case
progress 75/75; T075 now records actual approved publication consistency.

Actual Windows 11 desktop and Windows Server 2025 CI observations are separate.
Windows 10 testing is owner-excluded D43 (2026-10-07), not a remaining gate;
historical Windows 10 results remain untested. Live UNC is owner-excluded D35, while local/lexical T039
remains mandatory. Checksum identity supplies no verified signature. This containing
record-only checkpoint's final SHA, live synchronization, PR and CI are independently
observed after commit; no future checkpoint result is certified here. Actual publication is separately
recorded; PR28 merge and website deployment retain separate owner gates.
