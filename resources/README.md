# Private runtime payload

`private-runtime.zip` is tracked ordinary Git data, not a Git LFS pointer. It
contains the exact redistributable runtime, integrity marker, licenses, notices,
SBOM and source information already shipped in the verified v1.0.1 portable ZIP.
`private-runtime.json` records its provenance; the application embeds the digest
pins and authenticates the archive before extraction and every cache before use.

A clean clone or tagged GitHub source archive runs offline without developer
staging. The first image run extracts into a private temporary sibling and
atomically publishes a per-user cache keyed by application version and digest.
The full portable release uses its expanded app-relative runtime and omits this
second compressed copy. No expanded engine is tracked in Git.

The existing runtime lock remains the dependency inventory. See
[third-party notices](../docs/release/THIRD_PARTY_NOTICES.md) and
[corresponding source information](../docs/release/IMAGEMAGICK_SOURCE.md).
