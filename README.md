# Hint Book Formatter

Hint Book Formatter is a Vite + React app that reads data from Google Sheets or an Apps Script endpoint and renders print spreads plus booklet-order previews.

## User Manual

For the end-user operating guide in Japanese, see [MANUAL.md](./MANUAL.md).

## Local Development

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

Production builds use relative asset paths so the app works on GitHub Pages even when it is served from a repository subpath.

## Publish To GitHub Pages

1. Push this repository to GitHub.
2. Open `Settings` > `Pages`.
3. Set `Source` to `GitHub Actions`.
4. Push to `main` and `.github/workflows/deploy-pages.yml` will build and deploy the site.

The published URL is usually `https://<account>.github.io/<repository>/`.

## Deployment Notes

- If you load CSV data directly from Google Sheets, the sheet must be readable from each employee's browser.
- If you load data from Apps Script, the web app deployment and its CORS behavior must allow browser access.
- GitHub Pages visibility depends on your GitHub plan and repository visibility settings.

## Scripts

- `npm run dev`: start the dev server
- `npm run build`: run TypeScript checks and create a production build
- `npm run preview`: preview the production build locally
- `npm run lint`: run ESLint
