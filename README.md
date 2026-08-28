# MasterConvert — Conversion Backend

Real document conversion API. Handles:
Word↔PowerPoint, PDF↔Word, PDF→PowerPoint, Word↔Excel, Word→PDF.

## What each conversion actually does

| Pair | Method | Fidelity |
|---|---|---|
| Word → PDF | LibreOffice headless rendering | Exact — true rendering, not reconstruction |
| Word → PowerPoint | Reads headings as slide titles, paragraphs as bullets | Structural rebuild — works well for notes-style docs |
| PowerPoint → Word | Reads slide titles/text into headings + bullets | Structural rebuild |
| PDF → Word | Extracts real text per page (pypdf) | Text only — layout/images aren't preserved. Scanned PDFs with no text layer will say so per page rather than guessing |
| PDF → PowerPoint | Each page becomes a full-image slide, with OCR'd text (Tesseract) attached as speaker notes | Visual fidelity is exact; slide text is notes, not editable text boxes (see below) |
| Excel → Word | Each sheet becomes a heading + table | Exact |
| Word → Excel | Tables copy directly; a table-less doc falls back to one paragraph per row | Exact for tables |

**Why PDF→PowerPoint uses notes instead of live text:** OCR is not 100% reliable, especially on
diagrams, equations, or unusual fonts. Silently putting OCR guesses into the visible slide risks
corrupting content without anyone noticing. Putting it in speaker notes gives you the extracted text
to copy from deliberately, while the slide itself stays a perfect visual copy of the original.

## AI endpoints (added since the original conversion-only version)

| Endpoint | Does |
|---|---|
| `/api/extract-text` | Pulls plain text from an uploaded file for the AI tools — real extraction, no AI involved |
| `/api/summarize` | Calls Claude to condense pasted/uploaded text into slide bullets |
| `/api/write` | Calls Claude to draft an explanatory paragraph on a topic. Deliberately never generates citations — an AI has no way to verify a specific source is real, and a fabricated citation in a student's work is a genuine harm, not a cosmetic bug |

These need `ANTHROPIC_API_KEY` set as an environment variable — see deployment
steps below. Without it, both endpoints return a clear error instead of
crashing, and the app shows that error as a toast.

## Deploying on Render (matches your original plan)

1. Push this folder to a GitHub repo (Render deploys from Git, not file upload)
2. In Render: **New → Web Service** → connect the repo
3. Render will detect the `Dockerfile` automatically — leave the environment as **Docker**
4. Instance type: at least the smallest **paid** tier. The free tier's ~512MB RAM is tight for
   LibreOffice; expect slow or failing conversions on Free specifically
5. In **Environment**, add `ANTHROPIC_API_KEY` with your real key — get one at
   [console.anthropic.com](https://console.anthropic.com)
6. Deploy. Render assigns a URL like `https://masterconvert-api.onrender.com`

## Deploying on Railway (alternative)

Same idea: push to GitHub, **New Project → Deploy from GitHub repo**, Railway detects the
Dockerfile automatically. Add `ANTHROPIC_API_KEY` under Variables. Railway injects
`PORT` itself — already handled in this Dockerfile.

## After deploying

Copy the live URL Render/Railway gives you and paste it into `API_BASE` near the top of the
frontend's script block in `index.html`. Until that's set, the app will show a toast telling
you it's not configured yet, rather than failing silently.

## Testing your deployment

```
curl https://your-backend-url.onrender.com/api/health
# should return {"status":"ok"}
```

## Security note before going further

This API currently accepts requests from anywhere (`Access-Control-Allow-Origin: *`) and has no
rate limiting or auth. Fine for testing. Before real customers hit it, lock `ALLOWED_ORIGIN` in
`app.py` to your actual domain, and add basic rate limiting — otherwise anyone who finds the URL
can run unlimited free conversions through it.
