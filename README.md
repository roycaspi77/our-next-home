# Our Next Home 🏡

A shared list of the homes for sale that we send each other (Facebook posts, Yad2 ads), with a status, a call log and a chat for each one. It runs on Google Sheets and Apps Script plus GitHub Pages, costs nothing, and opens on PC and phone from a personal link.

## What it does
- **Table of homes:** link, city, street, plot m², built m², rooms, basement, description, price, ₪/m² (worked out from price and built area), seller, phone, status, next step, visit date.
- **Row colour = status:** 🟨 *Not seen*, 🟩 *Seen*, ⬜ *Not relevant*.
- **+ Add home → ✨ Analyze:** paste the link and press Analyze. The app reads the ad and fills **only what the ad says**. Anything it doesn't find stays blank for you to fill in.
  - **Yad2:** usually read straight from the link.
  - **Facebook:** Facebook doesn't let apps read posts, so also paste the **post text** or a **screenshot** (Ctrl+V on a PC, *choose* on a phone, up to 4). Screenshots are only read, never stored.
  - If the link is already in the list, Analyze tells you and offers to open that home.
- **Edit ✏️ / delete 🗑️**, or change status and basement right in the table.
- **💾 Save:** changes stay on your page (marked in blue) until you press Save (or Ctrl+S). If you both changed the same thing, the app shows both values and you pick one. Everything else is saved normally.
- **Filters:** a filter under every column header (text contains, min/max for numbers, yes/no, status). Type `-` to find blank cells. A search box searches everything. Click a header to sort. On a phone, tap **Filters**.
- **💬 Chat** per home, kept in the sheet. A red dot shows new messages from the other person. Messages are sent right away and don't need Save.
- **📞 Calls & next step** (in the chat window): a log of calls with the seller, plus the home's next step and visit date.
- **Phone:** 📞 calls, **WA** opens WhatsApp. **🗺️ / 🚗** open the address in Google Maps / Waze.
- The page picks up the other person's changes every minute.

## How it's built
- **`docs/`:** the screens, a static page on **GitHub Pages** (`https://roycaspi77.github.io/our-next-home/`). `config.js` holds the Apps Script `/exec` URL.
- **`apps-script/`:** the backend, a Google Apps Script bound to the Google Sheet (sheets `Homes`, `Chats`, `Calls`, `Users`). Analyze uses **Google Gemini** (free tier).
- Personal links look like `https://roycaspi77.github.io/our-next-home/#u=<token>`. The token goes after `#`, so it is never sent to GitHub. **Each link works like a password.**
- **Try it without Google:** open `docs/index.html?mock=1` (it uses made-up data in memory).

## First-time setup (about 10 minutes)
1. **Sheet + script:** create a Google Sheet named *Our Next Home* → **Extensions → Apps Script**. Replace `Code.gs` with [`apps-script/Code.gs`](apps-script/Code.gs). In **Project Settings** tick *Show "appsscript.json"*, then paste in [`apps-script/appsscript.json`](apps-script/appsscript.json). Save.
2. Choose **`setup`** in the toolbar → **Run** → approve the permissions. It creates the sheets and two people, `Roy` and `Wife`. You can rename them in the **Users** sheet (names show in the chat).
3. **Gemini key (free):** go to [aistudio.google.com/apikey](https://aistudio.google.com/apikey) → *Create API key*. In Apps Script → ⚙️ **Project Settings → Script Properties → Add script property**: `GEMINI_API_KEY` = the key. **Never put it on GitHub.**
4. **Deploy → New deployment → Web app**, with *Execute as: Me* and *Who has access: Anyone*. Copy the `/exec` URL into [`docs/config.js`](docs/config.js) and push.
5. Run **`printLinks`** (View → Logs shows the two links). Open yours, send your wife hers. On each phone: open the link → **Add to Home Screen**.

### Updating the script later
Paste the new `Code.gs` → Save → run `setup` (safe to run again; it only adds what's missing) → **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. The URL stays the same.

## Adding a column
Add one line to `COLUMNS` at the top of `Code.gs`, for example:
```js
{ key: 'parking', label: 'Parking', type: 'number', min: 0, max: 10, ai: 'Number of parking spaces (חניות)' },
```
Types: `text`, `longtext`, `number`, `money`, `checkbox`, `select` (with `options: [...]`), `date`, `phone`, `link`. Give it an `ai` description if Analyze should fill it, or leave `ai` out. Then run `setup` and deploy a new version. The page shows the new column by itself.

## Notes
- **No guessing:** Gemini is told to copy only what the ad states and to leave the rest blank. The script also drops numbers that don't make sense, and when it read text (not a screenshot) it drops a city, street, name or phone that doesn't appear in that text.
- **Gemini model:** `gemini-2.5-flash` by default. If Google retires it, Analyze says so; set a Script Property `GEMINI_MODEL` to a current model name.
- **Free quota:** Gemini's free tier allows plenty of analyses a day for two people. If it runs out, Analyze asks you to try again a minute later.
- If a link leaks, replace that person's `token` in the **Users** sheet with a new random string and send the new link.
- Tests for the Analyze helpers: `node .dev/test_helpers.js`.
