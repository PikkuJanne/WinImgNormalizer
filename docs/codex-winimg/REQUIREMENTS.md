# Accepted improvements — traceability

Every item in the accepted 18-item review is mapped below. Gate tasks replay preceding tests; primary implementation tasks are shown without the full-project gates.

| ID | Improvement | Primary tasks | Case coverage |
|---|---|---|---|
| R01 | Collision-free output names | M1-T03 | T017, T018, T019 |
| R02 | Destination exclusion and unique run roots | M1-T02 | T013, T014, T015, T016 |
| R03 | Validated temporary outputs and safe finalization | M1-T04 | T020, T021, T022, T023, T024 |
| R04 | Explicit multi-frame/page policy | M2-T01 | T029, T030, T031 |
| R05 | Colour management before metadata stripping | M2-T02 | T032, T033, T034, T035, T036 |
| R06 | Success-based heuristic duplicate handling | M1-T05 | T025, T026, T027 |
| R07 | Literal filename and root-path handling | M1-T01, M2-T03 | T008, T009, T010, T011, T012, T037, T038, T039 |
| R08 | Argument/dependency/destination preflight | M1-T01 | T008, T009, T010, T011, T012 |
| R09 | Exact byte caps, quality and warning status | M2-T04 | T040, T041, T042 |
| R10 | Actionable native diagnostics and justified retries | M2-T05 | T043, T044, T045 |
| R11 | Resilient discovery, links and ancillary failures | M3-T01 | T047, T048, T049, T050 |
| R12 | Resource limits and cancellation | M3-T02, M3-T03 | T051, T052, T053, T054, T055, T056 |
| R13 | Accurate counts, progress and local reports | M3-T04 | T057, T058, T059, T060 |
| R14 | Launcher outcomes and exit propagation | M3-T05 | T061, T062, T063 |
| R15 | Regression tests and target coverage | M0-T02, M0-T03, M3-T06 | T001, T002, T003, T004, T005, T006, T007, T064, T065, T066 |
| R16 | Minimal testable structure and CI checks | M0-T03, M3-T06 | T005, T006, T007, T064, T065, T066 |
| R17 | Versioned trusted distribution | M4-T01, M4-T02, M4-T06 | T068, T069, T070, T071, T075 |
| R18 | Accurate documentation and website presentation | M4-T03, M4-T04, M4-T06 | T072, T073, T075 |

## Cross-cutting invariants

Original bytes and source creation/modified timestamps are preserved by the tool; operating-system access-time side effects are not claimed controllable. Existing entry points, positional use and local-only operation remain. JPEG metadata policy, white transparency background and unchanged copied video bytes are explicit contracts. New tests and helpers are development assets, not new runtime frameworks.

The default byte cap remains 1,048,576 bytes. The scale sequence stays unchanged until a measured proposal is accepted. The agreed lightweight duplicate safeguard is success-only registration plus a length check; it is not proof of identical content. The first hardened release rejects a destination within its source rather than risking recursive self-ingestion. These changes must be visible in release notes and owner acceptance.

No added GUI, cloud service, content-hash database, global configuration replacement, automatic uploader, parallel worker pool or general-purpose media suite is part of this plan.
