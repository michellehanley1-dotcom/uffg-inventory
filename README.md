# UFFG Inventory — setup

A phone app for tracking tools and parts. No app store, no monthly cost.
About 30 minutes to set up the first time.

You need: a Google account (use **ricardo@uffgconstruction.com** so the sign-in
emails come from the company, not a personal Gmail), and a free GitHub account.

---

## 1. Make the sheet

1. Go to **sheets.new** while signed in as the Google account you picked.
2. Name it **UFFG Inventory**.
3. **Extensions → Apps Script**. A code window opens in a new tab.
4. Delete whatever is in there. Open `apps-script/Code.gs` from this folder,
   copy all of it, paste it in.
5. Click **Save** (the disk icon).
6. In the function dropdown at the top, pick **setup**, then click **Run**.
   - Google will ask for permission. Click through: **Review permissions** →
     pick your account → **Advanced** → **Go to Untitled project (unsafe)** →
     **Allow**. That warning is normal for a script you wrote yourself.
7. Switch back to the sheet. You should see tabs: Team, Locations, Tools,
   Parts, Stock, Log, Auth, Sessions, Processed.

Team and Locations are pre-filled with Ricardo, Michelle, Garage and Ricky's Truck.

## 2. Publish the back end

1. Back in the Apps Script tab: **Deploy → New deployment**.
2. Click the gear next to "Select type" and choose **Web app**.
3. Set:
   - Description: `v1`
   - **Execute as: Me**
   - **Who has access: Anyone**
4. Click **Deploy**, approve again if asked.
5. Copy the **Web app URL**. It ends in `/exec`. You need it in the next step.

> "Anyone" sounds alarming but is required — phones hit this without a Google
> login. The script checks the Team tab before it returns anything, so a
> stranger with the link gets nothing.

**Check it works:** paste the URL in a browser with `?ping=1` on the end.
You should see version and row counts.

## 3. Point the app at it

Open `config.js` and paste your `/exec` URL between the quotes:

```js
var CONFIG = {
  SCRIPT_URL: "https://script.google.com/macros/s/AKfyc.../exec"
};
```

## 4. Put it online

1. At github.com: **+ → New repository**, name it `uffg-inventory`, **Public**, create.
2. **uploading an existing file**, then drag in *everything inside this folder*
   (not the folder itself). Commit.
3. **Settings → Pages**, Source: **Deploy from a branch**, Branch: **main**,
   folder: **/ (root)**, Save.
4. Two minutes later the address appears at the top of that page.

## 5. Add it to your phone

Open that address on the phone.

- **iPhone:** Share button → **Add to Home Screen**
- **Android:** three-dot menu → **Install app** / **Add to Home screen**

It then opens full screen with no browser bars, and works with no signal.

---

## Day-to-day

**Adding a person.** Add a row on the **Team** tab: email, name, their truck or
base, and `Lead` or `Manager`. They sign in with that email. Deleting the row
signs them out everywhere, immediately.

**Managers** can fix anything and tap a teammate under More to see what that
person sees. **Leads** can only move things to or from their own base.

**Adding job sites.** Anyone can add one from the app under More.

**Importing a list.** More → Import a list. One item per line, commas between
fields. Tools are `name, brand, model, ID, location`. Parts are
`name, brand, unit, low target`. Only the first field is required.

**Low stock.** Each part has a low target. The Overview flags anything at or
below it. Managers set it by tapping a part, then Set low target.

**No signal.** Changes queue on the phone and send themselves when signal comes
back. A banner shows how many are waiting. The same change can never be saved
twice, even if the phone retries.

---

## Rules to follow later

**Every time you change `index.html`, bump the cache name in `sw.js`.**
Change `uffg-inv-v1` to `-v2`, then `-v3`, and so on. Skip this and phones will
keep showing the old version with no obvious reason why.

**Nothing private in this repo — it's public.** The sheet link and script URL
are fine. Passwords, the clasp credentials, and any customer data are not.

**Never create a second deployment.** Use **Deploy → Manage deployments →**
pencil icon → **New version → Deploy**. Making a new deployment gives you a new
URL and the phones all break.

**Health check:** `your-script-url/exec?ping=1` shows the version and how many
rows are in each tab. First thing to check if something looks wrong.

---

## Optional: auto-deploy the back end

Lets you edit `apps-script/Code.gs` in GitHub and have it push itself to Google.
Skip this unless you're changing the back end often — pasting into the Apps
Script editor works fine.

In the repo, **Settings → Secrets and variables → Actions**, add three secrets:

| Secret | Where it comes from |
|---|---|
| `CLASPRC_JSON` | Run `npx @google/clasp login` on a computer, then copy the whole contents of `~/.clasprc.json` |
| `SCRIPT_ID` | Apps Script → Project Settings → Script ID |
| `DEPLOYMENT_ID` | Apps Script → Deploy → Manage deployments → the ID under your web app |

`.github/workflows/deploy.yml` then pushes to that **same** deployment on every
change, so the URL never moves.

---

## If something breaks

| What you see | What it usually is |
|---|---|
| "Please sign in again" | Their row was removed from Team, or 120 days passed |
| No code email arrives | Check the Team tab spelling; check spam; a free Gmail caps at 100/day, Workspace at 1,500 |
| App looks out of date | `sw.js` cache name wasn't bumped. Bump it and re-upload |
| Everything fails at once | Open `?ping=1`. If that fails, the deployment was replaced |
| "You can only move tools to or from…" | A Lead acting outside their base. Make them a Manager or have one do it |
