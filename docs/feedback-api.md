# Feedback API (Prime Productivity platform)

The contract between the products' in-app feedback forms (Prime Word: `apps/word/src/feedback.ts`) and the
Prime Productivity platform (its own repository). A local stand-in that follows it exactly:
`node scripts/feedback-mock.mjs` (then build or run Prime Word with `VITE_FEEDBACK_API=http://localhost:8787`).

## Principles

- **The product never holds a secret.** Calls are anonymous; the platform validates everything and rate-limits per
  installation (`X-Prime-Install`, a random id per install) and per IP.
- **Idempotent by report id.** The product makes the id (a UUID) and may send the same report any number of times
  (retries after a lost connection); the platform keeps one record per id.
- **Attachments go straight to private storage** through one-time upload URLs (serverless request bodies are
  limited to about 4.5 MB; attachments may be 10 MB).
- **Offline first on the product side:** a report is on disk before it is sent and leaves only when step 3 succeeds.

## 1. Announce: `POST /v1/feedback`

Headers: `Content-Type: application/json`, `X-Prime-Install: <uuid>`, `X-Prime-Product: prime-word`.

```json
{
  "id": "1b9d6bcd-bbfd-4b2d-9b5d-ab8dfbbd4bed",
  "product": "prime-word",
  "category": "bug | difficulty | idea | general",
  "description": "What happened… (3–10 000 characters)",
  "email": "tester@example.com",
  "followUp": true,
  "createdAt": "2026-11-17T09:30:00.000Z",
  "diagnostics": {
    "appVersion": "0.1.3", "os": "Windows 11 (15.0.0)", "desktop": true, "time": "…", "timeZone": "Africa/Accra",
    "language": "en-GB", "screen": "1920×1080 @1.25",
    "context": { "screen": "document", "mode": "print", "pages": 4, "words": 1210, "openDocuments": 2, "documentFormat": "docx" }
  },
  "attachments": [ { "file": "a0", "name": "screenshot-1.png", "type": "image/png", "size": 183204, "sha256": "…" } ]
}
```

Rules: at most 5 attachments, each at most 10 MB; `file` is `a0`, `a1`…; diagnostics is free-form JSON (store as jsonb).

Answers:

- `200 { "status": "received" | "complete", "uploads": [ { "file": "a0", "url": "<signed upload URL>", "headers": { … } } ] }`.
  `uploads` lists only attachments not stored yet (empty on a retry after they were uploaded).
- `400` / `413` / `422` with `{ "error": "…" }`: the report itself is wrong. The product keeps it and does not retry.
- `429` (with `Retry-After`), `5xx`, or no answer: the product retries later.

## 2. Upload: `PUT <signed upload URL>`

The attachment's bytes, `Content-Type` as announced (plus any `headers` given). The platform checks size and
SHA-256 against the announcement.

## 3. Complete: `POST /v1/feedback/<id>/complete`

Same headers. `200 { "status": "complete" }` once every announced attachment is stored (`409` otherwise: the product
retries from step 1, which hands out the missing upload URLs again). Completing again is harmless. On the first
completion the platform notifies the owner (push to the console, email).

## Platform side (to build in the platform repository)

- Storage: a `feedback` table (id uuid primary key, product, category, description, email, follow_up, created_at,
  received_at, completed_at, install_id, ip_hash, diagnostics jsonb, attachments jsonb, status), and a private
  bucket `feedback/<id>/<file>`.
- Limits: e.g. 10 reports per installation per hour, 30 per IP per hour; descriptions up to 10 000 characters.
- The products' CSP must allow the API's origin and the storage origin in `connect-src` (Prime Word:
  `apps/word/src-tauri/tauri.conf.json` → `app.security.csp`).
