# MasterConvert — Google Play Store Readiness Checklist

Based on current Google Play policy (checked August 2026). This covers what's
done, what's genuinely required, and what only you can do from here — none of
this can be completed by generating more code; the remaining steps need a
Google Play Console account and native Android build tooling.

## ✅ Already handled in the app files

- **PWA foundation** — `manifest.webmanifest` + service worker + HTTPS-ready
  icons. Google's own 2026 guidance calls a TWA wrapping a proper PWA "the
  compliant approach" for Policy 4.3 (Minimum Functionality) — this is exactly
  what you have.
- **Privacy policy** — `privacy.html`, written to match what the app actually
  does (temp file processing, Claude API for AI features, Google Play Billing
  for payments, no data selling).
- **Account/data deletion** — Account Settings has a working "Request Account
  Deletion" action. Google now requires this for any app with user accounts.
- **In-app privacy summary** — condensed version live in the Profile menu.

## ⚠️ The one big real issue: payments

Google Play requires **Google Play Billing** for in-app digital subscriptions,
with alternative billing allowed only for users in the US, UK, EU, and South
Korea (post-antitrust-ruling carve-outs) — and even then only after formally
enrolling in Google's program. Ghana isn't covered. Paystack (or any other
third-party processor) triggered *inside* the app risks rejection — this is
exactly what got KakaoTalk blocked from Play Store updates in South Korea.

**What's in the code now:** `initiateGooglePlayPurchase()` is a placeholder
that grants the plan directly, clearly commented as non-functional. Real
charging requires the native **Play Billing Library** (Kotlin/Java), which
only works from inside actual Android code — it cannot be built in HTML/JS.
When you wrap this as a TWA, this is the one piece that needs a native
developer (or you learning just enough Android to wire up Play Billing calls
that talk back to this function).

## 🔧 Steps only you can do (need your own accounts/tools)

1. **Host `privacy.html`** at a real URL (e.g. `masterconvert.app/privacy`) —
   Play Console requires a privacy policy *URL* in the store listing, not
   just in-app text.
2. **Package as a TWA.** The standard no-code-adjacent path:
   [PWABuilder.com](https://pwabuilder.com) — point it at your deployed
   masterconvert.app URL, it generates the Android project for you. The more
   manual path is Google's own [Bubblewrap CLI](https://github.com/GoogleChromeLabs/bubblewrap).
   Either way you'll need Android Studio installed to build the final signed
   `.aab` file.
3. **Digital Asset Links** — both tools above generate an `assetlinks.json`
   file. It must be hosted at `masterconvert.app/.well-known/assetlinks.json`
   — this is what cryptographically proves you own the site the TWA wraps.
4. **Google Play Console account** — $25 one-time fee, real ID verification.
5. **Content rating questionnaire** — mandatory; Google no longer allows
   unrated apps.
6. **Data Safety form** — mandatory, 14 data categories to declare (file
   uploads, email, usage data). Must match `privacy.html` exactly — mismatches
   are a common rejection reason. Since you're using Claude for AI features,
   declare that data flow explicitly.
7. **Target API level** — whoever builds the final `.aab` needs to target
   the current required API level (Play Console shows the exact number at
   build time — this changes periodically, so check it fresh rather than
   using a number from this document).
8. **Set `ANTHROPIC_API_KEY`** as an environment variable on your Render
   service, or the AI Tools features will return a clear error instead of
   working.

## Realistic sequencing

Steps 1–3 you can do this week with no new skills. Step 2 in particular
(PWABuilder) is genuinely close to one-click once your site is live. Step 6
(native Play Billing) is the one piece that's a real, separate scope of
native development work — everything else here is configuration, not code.
