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

Inside a `body` cell, wrap text in `[[box]]...[[/box]]` to place it on the
default red text box. Specify another background color after a colon, for
example `[[box:#1f78c8]]...[[/box]]`.

Optional: create a `settings` sheet with these keys in column A and values in column B.

```text
name,迷宮からの脱出
version,v1
step_font_family,MS Mincho
body_font_family,Yu Gothic
side_font_family,MS Mincho
page_no_font_family,Arial
step_font_scale,4
body_font_scale,2.5
side_width,40
```

Recommended meaning:

- `name`: text placed before `ヒント` in the suggested PDF filename
- `version`: version text placed after `ヒント`; enter the complete text, such as `v1`
- `step_font_family`: font used for STEP
- `body_font_family`: font used for BODY
- `side_font_family`: font used for SIDE
- `page_no_font_family`: font used for page numbers
- `step_font_scale`: multiplier applied to spreadsheet font size for STEP
- `body_font_scale`: multiplier applied to spreadsheet font size for BODY
- `side_width`: width of the SIDE column. Numbers are treated as `px`.

Optional: create a `side` sheet when you want to manage SIDE labels by id.

Recommended columns:

- `id`
- `text`
- `height`

Optional text direction columns:

- `direction` or `text_direction`
- `rotation` or `rotate`

Supported `direction` values:

- `horizontal` or `h`: regular horizontal text
- `vertical` or `v`: Japanese vertical text, including vertical prolonged-sound marks
- `upright` or `u`: vertical text with Latin letters kept upright
- `rotate_cw`: rotate a horizontal phrase clockwise
- `rotate_ccw`: rotate a horizontal phrase counterclockwise

The older `rotation` / `rotate` column remains supported with these values:

- `90`, `90deg`, `cw`, `clockwise`
- `-90`, `-90deg`, `ccw`, `counterclockwise`

To mix directions inside one `text` cell, wrap only the relevant portions:

```text
[v]コーヒー[/v][h]OPEN 10:00[/h]
第[tcy]12[/tcy]問
[cw]CHAPTER ONE[/cw]
```

Supported inline tags are `[h]...[/h]`, `[v]...[/v]`, `[u]...[/u]`,
`[upright]...[/upright]`, `[tcy]...[/tcy]`, `[cw]...[/cw]`, and
`[ccw]...[/ccw]`. Untagged text inherits `direction`.
Spreadsheet rich-text formatting such as font color and bold remains attached to
the visible text. An incomplete or mismatched tag is shown literally instead of
discarding text.
