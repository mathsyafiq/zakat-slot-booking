# Interview Zakat Slot Booking

- `docs/index.html` — frontend, served by GitHub Pages
- `apps-script/Code.gs` — backend API (Google Apps Script + Google Sheet)

## 1. Deploy the backend (Apps Script)

1. Open your Google Sheet → **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with `apps-script/Code.gs`. Delete the old `Index.html` file from the project (no longer used).
3. Run `initialize()` once (authorise when asked) to create the `Slots` sheet. Fill in the Date/Time columns.
4. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Copy the Web App URL (ends in `/exec`).

> After changing `Code.gs` later, use **Deploy → Manage deployments → Edit → New version**; the URL stays the same.

## 2. Configure the frontend

In `docs/index.html`, set:

```js
var API_URL = 'https://script.google.com/macros/s/XXXX/exec';
```

## 3. Publish on GitHub Pages

```bash
git add . && git commit -m "Slot booking system"
git branch -M main
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```

Then on GitHub: **Settings → Pages → Source: Deploy from a branch → `main` / `/docs`**.
Your site will be at `https://<user>.github.io/<repo>/`.

## Notes

- The Web App URL is public by design; it only exposes slot listings and the booking action.
- Booked names/IDs/phones are never returned by the slots endpoint.
