# Tax document requests (Dropbox file request → Outlook draft)

`send_tax_requests.py` reads your client list (Excel). For each client it:

1. Creates a Dropbox **file request** named `Tejas & Anita Girase - 2025 Tax Documents`. Uploads
   go to `/File requests/2025 Tax Documents/<that name>`. If an open request with the same
   name already exists, the script reuses it.
2. Writes the standard request email with the upload link in it.
   - **Returning client:** the standard email only.
   - **New client:** the email also asks for last year's (2024) federal and state return and a
     Driver's License / State ID. When the client is MFJ, it asks for both spouses' licenses.
3. Saves the email as a **draft** in Outlook, addressed to the client, with the spouse CC'd
   when a spouse email is listed. `templates/Document checklist & Questionnaire.xlsx` is
   attached. **Nothing is sent.** You review the drafts and send them yourself.
4. Logs each row in `<clients>_requests_log.xlsx` with the link, its status and any error.
   A re-run skips rows already marked Done. Use `--force` to redo them.

## Client list

Column names are matched loosely (case, spaces and dashes don't matter). See `sample_clients.xlsx`.

| Field | Accepted headers | Notes |
|---|---|---|
| email (required) | Email, Email Address, E-mail | Several addresses can be separated with `;` |
| first name (required*) | First Name, Taxpayer First Name | |
| last name | Last Name | |
| spouse first / last name | Spouse First Name, Spouse Last Name | Spouse last name defaults to the client's |
| spouse email | Spouse Email | CC'd on the draft |
| filing status | Filing Status | `MFJ` / `Married Filing Jointly`. If it's blank and a spouse is listed, the row is treated as MFJ |
| client type | Client Type, New/Returning | `New` or `Returning`. If it's blank, the row is treated as **returning**. Use `--default-new` to flip that |
| full name | Name, Client Name | *Can replace the first and last name columns. `Tejas & Anita Girase`, `Tejas Girase & Anita Patel` and `Girase, Tejas` are all understood |

If your headers are different, map them yourself, for example
`--column email="Primary Email" --column client_type="Prior Client?"`.

## One-time setup (Windows)

```powershell
pip install -r requirements.txt
```

**Dropbox app:** go to <https://www.dropbox.com/developers/apps> and choose *Create app*,
then *Scoped access* and *Full Dropbox*. On the **Permissions** tab, tick
`account_info.read`, `files.metadata.read`, `files.content.write`, `file_requests.read` and
`file_requests.write`, then click Submit. Copy the **App key** and run:

```powershell
python send_tax_requests.py dropbox-auth --app-key YOUR_APP_KEY
```

This saves a refresh token to `.dropbox_token.json`. Keep that file private: it has access
to your Dropbox. If your firm uses Dropbox Business and the client folders belong in the team
space, add `--team-space` to the `run` commands.

**Outlook:**
- **Classic Outlook for Windows** is the default (`--outlook desktop`). Outlook must be
  installed and signed in. Drafts appear in your Drafts folder. To send on behalf of a shared
  mailbox, add `--send-as MileHighCPAAdvisors@mh-cpas.com`.
- **New Outlook, Mac, or Outlook on the web:** these have no desktop automation, so use
  `--outlook graph --graph-client-id <id>`. This needs an Entra ID (Azure) app registration
  that is a public client with *Allow public client flows* turned on and the delegated
  `Mail.ReadWrite` permission. The first run asks you to sign in with a device code.
  `--graph-mailbox` saves the drafts to a shared mailbox.

## Running

```powershell
# 1. Preview only: no Dropbox or Outlook calls. Writes HTML previews next to the list.
python send_tax_requests.py run "2025 Clients.xlsx" --tax-year 2025 --dry-run

# 2. Test on one client, then check the draft in Outlook
python send_tax_requests.py run "2025 Clients.xlsx" --tax-year 2025 --only tejasgirasr@gmail.com

# 3. Everyone
python send_tax_requests.py run "2025 Clients.xlsx" --tax-year 2025
```

Other useful options: `--limit 10` for a test batch, `--deadline 2026-04-01` to set an upload
deadline (paid Dropbox plans only), `--sender-name "Your Name"`, `--dropbox-root "/Clients/2025"`,
`--no-attachment` and `--yes` to skip the confirmation prompt. Run `python send_tax_requests.py run -h`
to see every option.

The exit code is 0 when everything worked and 2 when some rows had errors. Those rows are in
the log. Fix them and re-run: rows already done are skipped.
