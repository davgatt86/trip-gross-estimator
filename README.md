# Trip Gross Estimator — deploy guide

This is your Trip Gross Estimator packaged so it can live on the web at its own
address (e.g. `https://your-name.netlify.app`), which you and other skippers can
open on any phone. Your Anthropic API key stays hidden on the server — it is
never sent to anyone's phone.

There are two ways to get it online. **Option A (drag-and-drop) is easiest.**
**Option B (GitHub) is better if you want to make changes later.**

---

## What you need first (both options)

1. **An Anthropic API key.**
   - Go to https://console.anthropic.com → sign in → "API Keys" → "Create Key".
   - Copy the key (starts with `sk-ant-`). Keep it somewhere safe; you'll paste it once.
   - Note: API use is pay-as-you-go. Parsing a price sheet costs a fraction of a penny.

2. **A Netlify account** (free): https://app.netlify.com → sign up.

3. **This project folder** (the one containing `package.json`, `src`, `netlify`).

---

## Option A — Drag-and-drop (no coding tools)

Netlify can't build the project from a drag-and-drop of source code, so you first
build it into a `dist` folder on a computer, then drag that up. You need Node.js
installed (https://nodejs.org, "LTS" version).

1. Open a terminal **in this project folder** and run:
   ```
   npm install
   npm run build
   ```
   This creates a `dist` folder.

2. **But** drag-and-drop does NOT run the serverless function (the part that holds
   your key). So for the function to work you must use the Netlify CLI instead:
   ```
   npm install -g netlify-cli
   netlify deploy --build --prod
   ```
   - The first time, it opens a browser to log in and asks to create/link a site.
   - When it finishes it prints your live URL.

3. Set your API key on the site:
   - Go to https://app.netlify.com → your site → **Site configuration** →
     **Environment variables** → **Add a variable**.
   - Key: `ANTHROPIC_API_KEY`  Value: your `sk-ant-...` key. Save.
   - Then redeploy: `netlify deploy --build --prod` again (so the function picks up the key).

Done — open the URL on your phone.

---

## Option B — GitHub + Netlify (auto-builds, easiest to update)

1. Put this folder in a GitHub repository:
   - Create a free GitHub account, make a new repository, and upload these files
     (GitHub's website has an "upload files" button — you can drag the whole folder in,
     but **do not** upload `node_modules` if it exists).

2. Connect it to Netlify:
   - https://app.netlify.com → **Add new site** → **Import an existing project** →
     pick GitHub → choose your repo.
   - Build command: `npm run build`   Publish directory: `dist`
     (these are already set in `netlify.toml`, so it should auto-fill).
   - Click **Deploy**.

3. Add your API key:
   - Site → **Site configuration** → **Environment variables** → **Add a variable**.
   - Key: `ANTHROPIC_API_KEY`  Value: your `sk-ant-...` key. Save.
   - Trigger a redeploy: **Deploys** → **Trigger deploy** → **Deploy site**.

Now every time you change a file on GitHub, Netlify rebuilds automatically.

---

## Testing it locally first (optional)

If you have Node.js and want to try it on your own computer before deploying:
```
npm install
npm install -g netlify-cli
echo "ANTHROPIC_API_KEY=sk-ant-your-key" > .env
netlify dev
```
Open the address it prints (usually http://localhost:8888). `netlify dev` runs both
the app and the function together so uploads work.

---

## How the key stays safe

- The app in the browser calls **your own** address `/.netlify/functions/parse`.
- That function runs on Netlify's server, adds your `ANTHROPIC_API_KEY`, and calls
  Anthropic. The key lives only in Netlify's environment — it is never in the app
  code or sent to a phone.

---

## Day-to-day use

1. Open your Netlify URL on your phone.
2. **Boat tally** — upload the boat `.xlsx` (or PDF). It reads every size line.
3. **Peterhead prices** — upload the Don Fishing PDF. Auto-fills; check the grid.
4. **Hanstholm prices** — upload the Hanstholm PDF. Check the grid.
5. **Check mapping** — glance over the cards; amber = a price was missing and got
   substituted from the other market.
6. **Estimated gross** — see both markets side by side; **Export CSV** to keep a record.

Prices and tally are not saved between sessions — each visit starts fresh from the
built-in sample, and you upload the day's files. (If you later want trips to save,
that can be added.)

---

## If something doesn't work

- **"Server missing ANTHROPIC_API_KEY"** → the environment variable isn't set, or you
  didn't redeploy after setting it. Set it and redeploy.
- **A parse fails** → the on-screen message now says why (e.g. "server 429" = too many
  requests, wait a minute; "no JSON found" = try the PDF instead of a photo). Your
  current prices are kept; you can always edit any cell by hand.
- **The boat .xlsx won't read** → it expects the standard format (species row, then
  indented size rows, then Total). CSV also works.
