1. Open `script.google.com`.
2. Create a new Apps Script project.
3. Replace the default code with [`Code.gs`](./Code.gs).
4. Update `spreadsheetId` and `sheetName` if needed.
5. Deploy:
   `Deploy -> New deployment -> Web app`
6. Execute as:
   `Me`
7. Who has access:
   `Anyone`
8. Copy the Web app URL.
9. Paste that URL into `Apps Script Web App URL` in this app.

The web app returns JSON like this:

```json
{
  "rows": [
    {
      "order": "1",
      "page_no": "1",
      "step": "1st-1",
      "side": "1st",
      "body": "text",
      "image": "",
      "stepStyle": {
        "backgroundColor": "#d94b67",
        "textColor": "#ffffff",
        "fontFamily": "Noto Serif JP"
      }
    }
  ]
}
```

Optional: create a `settings` sheet with these keys in column A and values in column B.

```text
step_font_family,MS Mincho
body_font_family,Yu Gothic
side_font_family,MS Mincho
page_no_font_family,Arial
step_font_scale,4
body_font_scale,2.5
```

Recommended meaning:

- `step_font_family`: font used for STEP
- `body_font_family`: font used for BODY
- `side_font_family`: font used for SIDE
- `page_no_font_family`: font used for page numbers
- `step_font_scale`: multiplier applied to spreadsheet font size for STEP
- `body_font_scale`: multiplier applied to spreadsheet font size for BODY
