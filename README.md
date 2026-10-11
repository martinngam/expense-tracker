# Expense Tracker (Google Apps Script + DeepSeek)

A mobile-friendly web app for several users. Each person signs in, takes a photo of a receipt (or uploads one), checks the details it reads, and saves them to your Google Sheet.

**Sheet columns:** Date | Pay To | Item | Amount | Category | User | Submitted At | Receipt (Drive link)

## How scanning works
1. The browser shrinks the photo to 1600px wide or less (JPEG), so uploads stay fast.
2. Google Drive's built-in OCR turns the image into text. This is free and needs no extra key.
3. **DeepSeek** (`deepseek-chat`) reads that text and returns the Date, Pay To, Item, Amount and Category as JSON.
4. The user checks or edits the fields, then saves.
5. A copy of the receipt image is kept in a Drive folder called `Expense Receipts`.

> DeepSeek's chat API only accepts text, not images. That's why the image goes through Drive OCR first.

## Setup (about 10 minutes)
1. Create a new Google Sheet, then open **Extensions → Apps Script**.
2. In the editor, delete the default `Code.gs` content and paste in **Code.gs**.
3. Click **+ → HTML**, name it `Index`, and paste in **Index.html**.
4. Turn on the manifest: **Project Settings → "Show appsscript.json"**. Replace its content with **appsscript.json** and change `timeZone` to your own time zone.
   (This also turns on the **Drive API** advanced service. You can check under *Services*.)
5. Save, then reload the Google Sheet. A new **Expense App** menu appears:
   - **1. Setup sheets**: creates the `Expenses` and `Users` tabs. Approve the permissions when asked.
   - **2. Set DeepSeek API key**: paste your key from https://platform.deepseek.com.
   - **Add / reset user**: run this once for each person. Answer "Yes" to the admin question for yourself.
6. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Copy the `/exec` URL and share it with your users. On a phone they can use "Add to Home Screen".

After you change the code, use **Deploy → Manage deployments → Edit → Version: New version**. That keeps the same URL.

## Multi-user notes
- Users sign in with a username and PIN. PINs are stored salted and hashed in the `Users` tab.
- To disable someone, set their `Active` value to `FALSE`.
- Normal users only see their own recent expenses. Admins see everyone's.
- After 5 wrong PINs, that username is locked for 15 minutes. Sessions last 6 hours.
- Because the app runs as **you**, users never need access to the sheet itself.

## Customising
- Categories: edit `CONFIG.CATEGORIES` in `Code.gs`.
- OCR language: change `ocrLanguage: 'en'` in `ocrImage_` (for example `'zh'` or `'ms'`).
