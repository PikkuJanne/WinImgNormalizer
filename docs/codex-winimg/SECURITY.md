# Security and privacy constraints

The threat surface is local filesystem traversal, native image parsing, process
execution, temporary outputs, downloaded development dependencies and release
provenance. This is not a server product. No network input feature is being added.

Use literal paths and a verified executable. Do not use shell evaluation or accept
arbitrary remote URLs/coder expressions. Do not assume that an allowed filename
extension by itself makes the file safe. Apply practical resource/time limits,
respect restrictive installed policies, use supported patched codecs, and record
known residual risk. Per-job limits are not an exploit-proof sandbox. Current
ImageMagick security documentation and advisories should be checked before choosing
the development/release baseline. Do not copy an overly restrictive web policy that
breaks TIFF/HEIC, and do not weaken a user's global policy to make tests pass. [I3,I4]

Dependency installation during development is an explicit environment action, never
part of normal image processing. Prefer a documented trusted installation source,
record version/checksum or signature verification, and do not claim a dependency
is signed when it was not verified. The image tool itself has no telemetry, update
polling, uploader or account service in this plan. Website analytics, if later added
in another project, must not be confused with runtime behavior here.

No writes to source media, destructive cleanup, forced overwrite, ACL mutation,
admin escalation or machine-wide execution-policy weakening. Temporary cleanup
must use explicit owned paths. Retained final results and unrelated processes
survive cancellation. Denied folders are reported, not forcibly opened.

Git and CI use least privilege. Workflows should have `contents: read` unless a
separately approved release job genuinely needs more. Pin external actions to
reviewed full commit SHAs and record version comments; verify the SHAs rather than
inventing pins. Avoid `pull_request_target` with untrusted checked-out code. Never
expose write tokens/secrets to contributor-controlled scripts. Synthetic artifacts
only; no private archives, usernames, local paths, tokens or customer content.
Do not automatically alter repository security settings to satisfy a checklist. [G3]

Checksums detect corruption relative to a trusted checksum; they do not prove the
publisher's identity if artifact and checksum are replaced together. Authenticode
or artifact attestations may be added later with real credentials/provenance. No
mock signature, fabricated badge, permanent Defender exception, or blanket
Unrestricted-policy instruction is acceptable.

All CSV output must neutralize dangerous cell-leading text; quoting a formula is
not enough. Logs/reports may reveal filenames and source paths even when image
metadata is stripped. Keep those local or sanitize them for support. Copied videos
retain original metadata. Local output in a cloud-synced Pictures folder can be
uploaded by the user's sync software; say the tool itself does not upload, not that
no other software can ever synchronize the output.
