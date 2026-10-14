# Tax document requests (Dropbox file request → Outlook draft)

`send_tax_requests.py` reads your client list (Excel). For each client it:

1. Creates a Dropbox **file request** named `Tejas & Anita Girase - 2025 Tax Documents`. Uploads
   go to `/File requests/<client>` (e.g. `/File requests/Tejas & Anita Girase`). If an open request with the same
   name already exists, the script reuses it.
2. Writes the standard request email with the upload link in it.
   - **Returning client (MFJ or not):** the standard email only.
   - **New client:** the email also asks for last year's (2024) federal and state return and a
     Driver's License / State ID: both spouses' when MFJ, the client's own otherwise.
3. Saves the email as a **draft** in Outlook, addressed to the client, with the spouse CC'd
   when a spouse email is listed. `templates/Document checklist & Questionnaire.xlsx` is
   attached. **Nothing is sent.** You review the drafts and send them yourself.
4. Logs each row in `<clients>_requests_log.xlsx` with the link, its status and any error.
   A re-run skips rows already marked Done. Use `--force` to redo them.

## Client list (Drake Client Contact List)

Use `templates/Drake_Client_Contact_Template.xlsx`, rebuilt from the Drake "Client Contact
List" CSV export as described in its Notes tab. `sample_clients.xlsx` is a filled-in example.

- The script reads the **Individual Clients** tab. **Business Clients** is ignored, because these
  requests are for personal returns.
- The template's example row is skipped automatically, because its notes say "EXAMPLE ROW".
- **To:** Taxpayer Email. **CC:** Spouse Email. If Taxpayer Email is blank, the draft goes to
  the spouse instead and greets the spouse.
- **MFJ:** a client with a Spouse First Name or Spouse Email is treated as MFJ.
- **Names:** Drake's capitals (`JANE SAMPLE`) become `Jane Sample` in the email, subject and folder.
- **New vs returning:** I added one column to the template, **N: Client Type (New / Returning)**,
  with a New/Returning dropdown. A blank cell means returning. You can also skip the column and
  mark new clients on the command line: `--new client@email.com --new other@email.com`.
- If two rows share an email, only the first gets a draft and the second is logged as skipped.
- Rows with no email at all are logged as errors. They match the Still Missing tab.

Other spreadsheets also work. Headers are matched loosely, and `--column FIELD=HEADER` maps
anything unusual. Use `--sheet` to pick a different tab.

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
python send_tax_requests.py run "Drake Client Contact List.xlsx" --tax-year 2025 --dry-run

# 2. Test on one client, then check the draft in Outlook
python send_tax_requests.py run "Drake Client Contact List.xlsx" --tax-year 2025 --only tejasgirasr@gmail.com

# 3. Everyone
python send_tax_requests.py run "Drake Client Contact List.xlsx" --tax-year 2025
```

Other useful options: `--limit 10` for a test batch, `--deadline 2026-04-01` to set an upload
deadline (paid Dropbox plans only), `--sender-name "Your Name"`, `--dropbox-root "/Clients/2025"`,
`--no-attachment` and `--yes` to skip the confirmation prompt. Run `python send_tax_requests.py run -h`
to see every option.

The exit code is 0 when everything worked and 2 when some rows had errors. Those rows are in
the log. Fix them and re-run: rows already done are skipped.
