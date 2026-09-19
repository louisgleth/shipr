# Shipide GitBook publishing

Published and verified over HTTPS: https://docs.shipide.com

GitBook site: https://app.gitbook.com/o/SSngBtXf8qGV5zOw0w94/sites/site_Rft2W

GitBook fallback URL: https://shipide.gitbook.io/shipide-docs/

Configured in Cloudflare: CNAME `docs` to `5842f9da13-hosting.gitbook.io`, DNS only (not proxied), TTL Auto. GitBook detected the record, and the custom domain returned HTTP 200 with valid HTTPS on September 19, 2026.

The hosted OpenAPI source is https://portal.shipide.com/docs/openapi.json. GitBook is configured to import it under the name `shipide-api`; its URL import checks for updates every six hours. Guide content was imported as Markdown rather than enabling Git Sync.

Source content lives in `docs/`. `.gitbook.yaml` selects that root for a future Git Sync connection. `npm run docs:build` generates the importable OpenAPI schema. GitBook content can also be imported as a Markdown ZIP containing README.md, SUMMARY.md, the topic files, and openapi.json.

## Branding

- Site title: Shipide Developers
- Logo on dark background: `assets/shipide-logo-white.png`
- Logo on light background: `assets/shipide-logo-black.png`
- Favicon: `assets/favicon.png`
- Primary accent: `#7747E3`
- Dark background: `#00060F`
- Primary text: `#F1F4FB`
- Brand font: Creato Display; use the nearest available neutral sans-serif if the GitBook plan does not allow custom fonts.
- Code font: PT Mono where supported.
- Header links: Portal (`https://portal.shipide.com`), Support (`mailto:info@shipide.com`).

Set these in GitBook's site customization panel; `.gitbook.yaml` does not control hosted-site branding. Use GitBook's actual domain setup to obtain its CNAME target. Do not guess a target. Custom-domain support may require a paid GitBook site plan; do not purchase a plan without the user's approval.

Applied branding: both logo variants, favicon, `#7747E3` in light and dark modes, dark default mode with viewer toggle, and title "Shipide Developers". Main typography uses GitBook's Inter fallback; the original brand fonts are OTF/TTF while GitBook requires WOFF/WOFF2 and a higher plan for custom fonts. The account was placed on a 14-day trial by GitBook onboarding; no paid plan was purchased.

## API deployment

1. Run `npm run test:api`, `npm run docs:build`, and the existing admin analytics regression test.
2. Apply `wrangler d1 migrations apply shipide-customer-api --remote --config wrangler.shopify-api.toml`.
3. Dry-run and deploy `wrangler.shopify-api.toml`.
4. Deploy the frontend containing `developer.html`, `developer.js`, and `developer.css`.
5. Verify a test key can complete the sandbox lifecycle on the live service.

Live booking cannot be enabled until a carrier/provider integration and atomic billing settlement are implemented and tested. The API intentionally returns `carrier_not_configured` for live rate and label operations in the meantime.
