<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://ai.google.dev/static/site-assets/images/share-ais-513315318.png" />
</div>

# Run and deploy your AI Studio app

This contains everything you need to run your app locally.

View your app in AI Studio: https://ai.studio/apps/65ce82e0-2ec2-4bbc-b6c6-17cd711c4b05

## Run Locally

**Prerequisites:**  Node.js


1. Install dependencies:
   `npm install`
2. Set the following in [.env.local](.env.local):
   - `GEMINI_API_KEY` — your Gemini API key (used for resume PDF summarization and interview audio summarization)
   - `GOOGLE_CLIENT_ID` — an OAuth 2.0 Client ID from Google Cloud Console (see below)
3. Run the app:
   `npm run dev`

## Google Sign-In / Drive setup

This app signs users in with their `bloom-firm.com` Google account and stores each user's KPI data as a JSON file in their own Google Drive (auto-shared domain-wide so teammates can see aggregated views). To set this up:

1. Create (or select) a Google Cloud project associated with the `bloom-firm.com` Workspace.
2. Enable the **Google Drive API** for that project.
3. Configure the **OAuth consent screen**: User Type = **Internal**, and add the scopes `openid`, `email`, `profile`, `https://www.googleapis.com/auth/drive.file`, `https://www.googleapis.com/auth/drive.readonly`.
4. Create an **OAuth Client ID** (application type: **Web application**), and add your dev/prod origins (e.g. `http://localhost:3100`) to "Authorized JavaScript origins".
5. Put the resulting Client ID into `.env.local` as `GOOGLE_CLIENT_ID`.

Notes:
- Only `@bloom-firm.com` accounts can sign in (enforced client-side via the Google userinfo response; the actual data-access boundary is Drive's domain-wide sharing, enforced by Google).
- The "チーム管理" (Teams) feature stores team membership in a single shared `kpi-manager-teams.json` file. Because of how the `drive.file` OAuth scope works, only the person who originally created that file can edit team membership — everyone else can view teams but not edit them.
