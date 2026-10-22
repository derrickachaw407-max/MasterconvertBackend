# Academic-Home — server

The server behind Academic-Home (formerly Docente), the study app at
[getdocente.vercel.app](https://getdocente.vercel.app). The app itself (the web pages) is uploaded to Vercel
separately; this repository is only the server, and Render builds it from the `Dockerfile` on every push.

## What it does

| Area | What happens here |
|---|---|
| Conversions | Word, PDF, PowerPoint and Excel, both ways, with LibreOffice, pdf tools and Tesseract. A conversion mirrors the original: nothing is added, and converting back gives the original's words. |
| Slide Studio | Builds decks in the app's own templates, or keeps an imported deck's own design. |
| PDF and Word editors | Edits, signatures, forms, redaction and password protection (the last two on Pro). |
| Remy and Evidence-Based Writing | AI answers and drafts from Claude. When the question is about facts, every sentence carries its source — a real paper (OpenAlex, Semantic Scholar, Europe PMC) or a web page — and anything no source backs is left out. |
| Quizzes and decks | Practice quizzes, summaries and full decks drafted from notes. |
| Accounts | Email and password (codes by email), Google sign-in, password reset, account deletion. |
| Plans | Paystack on the web; the App Store and Google Play in the store apps. |
| Meetings | Academic-Home Meet, on LiveKit: waiting room, passcodes, encryption, presenting files, live quizzes, attendance. |

Check it's running: `https://<your server>/api/health` answers `{"status": "ok", …}`.

## Settings on Render (Environment)

Keys and passwords go **only** into Render's Environment page — never into a file in this repository.

| Setting | Needed for |
|---|---|
| `DATABASE_URL` | Required: the Postgres database. |
| `SECRET_KEY` | Required: signs sign-in tokens. |
| `ANTHROPIC_API_KEY` | Remy, writing, quizzes, summaries and decks. |
| `EMAIL_ADDRESS`, and `BREVO_API_KEY` or `EMAIL_APP_PASSWORD` | Confirmation codes and password-reset emails. |
| `PAYSTACK_SECRET_KEY`, `PAYSTACK_PUBLIC_KEY` | Paying for Pro on the web. |
| `GOOGLE_CLIENT_ID` | Google sign-in. |
| `SEMANTIC_SCHOLAR_API_KEY`, `OPENALEX_API_KEY` | Optional: faster, steadier paper searches. |
| `LIVEKIT_URL`, `LIVEKIT_API_KEY`, `LIVEKIT_API_SECRET` | Meetings (they stay switched off until all three are set). |
| `APPLE_*`, `GOOGLE_PLAY_*`, `GOOGLE_RTDN_TOKEN` | Purchases in the store apps. |
| `FRONTEND_URL` | Optional: the app's address in emails (default `https://getdocente.vercel.app`). |
| `EXTRA_ALLOWED_ORIGINS` | Optional: a custom domain for the app (any `*.vercel.app` address is already accepted). |
| `PROMO_PRO_UNTIL` | Optional: free Pro for everyone until this date and time (`off` to stop). |
| `SOFFICE_AT_ONCE` | Optional: LibreOffice conversions at the same time (default 1, for a 512 MB server). |

## Meetings (LiveKit)

1. Make a free project at [cloud.livekit.io](https://cloud.livekit.io) and create an API key.
2. On Render, add `LIVEKIT_URL` (the project's `wss://…` address), `LIVEKIT_API_KEY` and `LIVEKIT_API_SECRET`.
3. In LiveKit Cloud, under Settings → Webhooks, add `https://<your server>/api/meet/webhook`, signed with the same
   API key. Attendance and the monthly minutes count come from these webhooks; if none arrive while people are in a
   meeting, the server says so in its log.

Limits, all changeable on Render without a code change (defaults match LiveKit Cloud's free Build plan):
`MEET_MAX_PEOPLE` (100 in one meeting), `MEET_GLOBAL_MAX` (100 across all meetings at once),
`MEET_MONTH_MINUTES` (5,000 participant-minutes a month), `MEET_FREE_MINUTES` (60-minute meetings on Free),
`MEET_PRO_MINUTES` (300 on Pro) and `MEET_AUDIENCE_OVER` (above 30 people, newcomers listen until let speak).

## Security

- Sign-in travels as a token in the `Authorization` header, never a cookie.
- Rate limits use the visitor's real address from Render's edge (`True-Client-IP`), which visitors can't fake.
- Every answer carries `nosniff`, `X-Frame-Options: DENY`, `Referrer-Policy: no-referrer` and HSTS; signed-in
  answers aren't stored by caches.
- Meeting rooms are made only by the server. Only the host and the co-hosts they choose can let people in, mute or
  remove them; only the host can end or delete a meeting or name co-hosts. Files presented in a meeting are PDFs only.
- Payment errors never show Paystack's own messages or keys.

## Running the tests

The tests live with the app's source, outside this repository. Each release runs the full set — the app in a real
browser (light and dark, phone and laptop, contrast and accessibility checks) and this server against a real
Postgres and a real LiveKit server.
