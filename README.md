# Off Rota Request Site

This local website replicates the June 2026 off-rota request Google Form fields:

- Email address
- Staff initials
- Dates to avoid from
- Dates to avoid to
- Whether the dates are part of continuous annual leave

## Run

Double-click `start-site.bat`, then open:

http://localhost:4173

## Deploy On Render

Create a Render web service from this folder or from a GitHub repository containing this folder.

Render settings:

```text
Build command: npm install
Start command: npm start
```

Environment variables:

```text
ADMIN_PASSWORD=your-secure-password
DATA_DIR=/opt/render/project/src/storage
```

Persistent disk:

```text
Mount path: /opt/render/project/src/storage
Size: 1 GB
```

The included `render.yaml` contains the same settings for Render Blueprint deployment.

## Admin

Default admin password:

`ChangeMe2026!`

The admin panel lets you update the deadline date and time and download the Excel workbook.

To use a different password, start the site from PowerShell like this:

```powershell
$env:ADMIN_PASSWORD="your-new-password"
.\start-site.bat
```

## Response Storage

Responses are stored in:

`data/off-rota-responses.xlsx`

The workbook is regenerated after each successful submission.
