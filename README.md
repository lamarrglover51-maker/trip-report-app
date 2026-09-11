# Trip Report Builder

Builds the trip on-time-performance Excel report from a Google Sheet, with
filters for date range, carrier, and trip/load number.

- `build_trip_report.py` - the underlying logic + a CLI (`python build_trip_report.py --sheet ...`)
- `app.py` - a Streamlit web app wrapping the same logic with a UI

## Running locally

```
pip install -r requirements.txt
streamlit run app.py
```

Opens at http://localhost:8501, only on your machine. The rest of this
file is about making that same app reachable at a public link for
everyone else, via **Streamlit Community Cloud** (free, official
Streamlit hosting).

## Deploying so anyone with the link can use it

This takes three one-time steps: (1) create a Google service account so
the hosted app can read your sheet without a human login, (2) put the
code on GitHub, (3) deploy it on Streamlit Community Cloud. About
15–20 minutes total.

### 1. Create a Google service account (so the hosted app can read your Sheet)

The app currently logs into Google Sheets using *your* personal login
(a browser window pops up once, asks you to sign in). That only works
when *you're* the one running it on *your* machine - a server running
in the cloud has no browser and no "you" to log in as. A **service
account** is Google's fix for that: a robot account with its own key
file instead of a password, that only sees whatever sheets you
explicitly share with it.

1. Go to [console.cloud.google.com](https://console.cloud.google.com/)
   and either use the same project from your original OAuth setup, or
   create a new one.
2. Make sure the **Google Sheets API** is enabled for that project
   (APIs & Services > Library > search "Google Sheets API" > Enable).
3. Go to **IAM & Admin > Service Accounts > Create Service Account**.
   Give it any name (e.g. "trip-report-reader"). Skip the optional
   permission-granting steps - it doesn't need project-level roles.
4. Open the service account you just made > **Keys** tab > **Add Key >
   Create new key > JSON**. This downloads a `.json` file - **treat it
   like a password, don't share it or commit it to GitHub.**
5. Copy the service account's email address (looks like
   `something@your-project.iam.gserviceaccount.com` - it's in the JSON
   file as `client_email`, and on the service account's details page).
6. Open your Google Sheet > **Share** > paste that email in > give it
   **Viewer** access > Send. This is the only sheet the hosted app will
   ever be able to read - do this for any other sheet you want it to
   read too.

### 2. Put the code on GitHub

No `git` install needed - GitHub's website lets you upload files
directly.

1. Go to [github.com/new](https://github.com/new), create a new
   repository (Private is fine - Streamlit Community Cloud can deploy
   from private repos).
2. On the new repo's page, click **uploading an existing file**, and
   drag in every file from this folder **except**:
   - `credentials.json` / `credentials.json.json`
   - `.streamlit/secrets.toml` (if you made one for local testing)
   - any `.xlsx` files, `input.txt`
   - (the `.gitignore` in this folder already lists these - if you end
     up using real `git` instead of the web upload, it'll skip them
     automatically)
3. Commit the files.

### 3. Deploy on Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io/) and sign in
   (you can sign in with your GitHub account).
2. Click **New app**, pick the repo you just created, branch `main`,
   main file path `app.py`.
3. Before/after it deploys, open the app's **Settings > Secrets** and
   paste in something shaped like
   [.streamlit/secrets.toml.example](.streamlit/secrets.toml.example)
   in this repo, with real values:
   - `app_password` - make up a password; this is what gates the app
     (see below).
   - `[gcp_service_account]` - copy every field over from the JSON key
     file you downloaded in step 1.
4. Save. The app restarts and gives you a permanent link like
   `https://your-app-name.streamlit.app`.

### 4. Share it

Send the `*.streamlit.app` link plus the `app_password` you chose to
whoever needs it. Anyone with both can open the link, enter the
password once per browser session, pick their filters, and download
the report - no Google login, no local setup, on their end.

## Notes

- The hosted app can only read sheets you've explicitly shared with the
  service account's email (step 1.6) - sharing a *different* sheet with
  someone doesn't expose it through this app unless you also share it
  with the service account.
- To rotate the password, or if the service account key ever leaks,
  update the values in Settings > Secrets on Community Cloud - no
  redeploy or code change needed.
- Local `streamlit run app.py` and the CLI still work exactly as
  before, using your own Google login - the service account is only
  used when `gcp_service_account` is present in secrets (i.e. on the
  hosted deployment, or if you copy the example into a local
  `.streamlit/secrets.toml` to test this path yourself).
