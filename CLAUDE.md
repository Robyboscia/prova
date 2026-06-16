# mentalpro.it — Contesto Progetto

## Struttura repo

```
main branch        → codice sorgente / sviluppo (questo branch)
gh-pages branch    → file serviti da GitHub Pages su mentalpro.it
worktree gh-pages  → /tmp/gh-pages-wt  (usa questo per editare il sito live)
```

**Regola fondamentale**: le modifiche al sito live vanno fatte nel worktree `/tmp/gh-pages-wt`, non qui su `main`. Dopo ogni modifica: `git add`, `git commit`, `git push origin gh-pages`.

Se il worktree non esiste:
```bash
git worktree add /tmp/gh-pages-wt gh-pages
```

## File principali

| File | Descrizione |
|------|-------------|
| `/tmp/gh-pages-wt/index.html` | Landing page mentalpro.it — file principale |
| `/tmp/gh-pages-wt/index-v1.html` | Vecchia versione — ha `noindex` — NON toccare |
| `/tmp/gh-pages-wt/assetto-biologico/index.html` | Altro prodotto — ha `noindex` — NON toccare |
| `/tmp/gh-pages-wt/robots.txt` | Disallows index-v1 e assetto-biologico |
| `/tmp/gh-pages-wt/sitemap.xml` | Solo https://mentalpro.it/ |
| `/tmp/gh-pages-wt/site.webmanifest` | PWA manifest con icone e colori brand |

## Brand & colori

- **Nome**: Mental Pro™
- **Colore primario**: `#1757c2` (blu)
- **Sfondo scuro**: `#0d1b2a`
- **Font titoli**: Fraunces (italic, weight 300/400) — file locali in `./fonts/`
- **Font corpo**: Inter — file locali in `./fonts/`
- **Dominio**: https://mentalpro.it/

## Ottimizzazioni già applicate

### Performance
- `contain: layout paint` + `will-change: transform` su navbar, sticky bar, urgency bar, popup
- Hero img: animazione senza delay (`0s forwards` invece di `0.45s both`)
- 4 `<link rel="preload">` per tutti i font critici (incluso Fraunces italic latin)
- `decoding="async"` su tutte le immagini lazy
- `prefers-reduced-motion` copre hero, fade-up, btn, card, sticky, popup
- `<noscript>` CSS fallback per elementi animati

### SEO
- `<link rel="canonical" href="https://mentalpro.it/">`
- `<meta name="theme-color" content="#1757c2">`
- `<meta name="robots" content="noindex, nofollow">` su index-v1 e assetto-biologico
- JSON-LD `@graph`: Organization, WebSite, WebPage, Product, FAQPage
- `robots.txt` e `sitemap.xml` presenti e corretti

### Analytics & Tracking
- Google Tag Manager: `GTM-NF8SMQR8` — script in `<head>` (il più in alto possibile), noscript subito dopo `<body>`
- GA4 e Microsoft Clarity sono gestiti DENTRO GTM (non più script standalone nell'HTML): tag "GA4 Config" (Google Tag, ID `G-F5N0FTWSCL`, con parametri evento condiviso `anonymize_ip: true` e `cookie_flags: SameSite=None;Secure`) e tag "Clarity" (HTML personalizzato, project ID `x7x1upcqtp`)
- Trigger comune: evento custom `deferred_tracking_load`, pushato nel dataLayer da uno script in `index.html` al primo evento utente (click/scroll/keydown/touchstart) oppure dopo 3s di timeout
- Per modificare/aggiungere tag (es. Meta Pixel) si lavora da dashboard GTM, non nel codice
- `product:price:amount` (47.00) e `product:price:currency` (EUR) in Open Graph per condivisioni social con prezzo

### Favicon & Icone
- `favicon.ico` (32×32), `favicon.svg`, `apple-touch-icon.png` (180×180)
- `icon-192.png`, `icon-512.png` per Android/PWA
- `site.webmanifest` collegato in `<head>`

### Tool esterni verificati
- **Google Search Console**: verificato, sitemap elaborata, 1 pagina indicizzata
- **Bing Webmaster Tools**: verificato via import GSC, sitemap inviata manualmente
- **Google Business Profile**: scheda "Mental Pro" creata e in verifica

## Da fare (non ancora implementato)

1. **Cookie consent banner** — obbligatorio GDPR/Garante per GA4
2. **Privacy Policy** — `/privacy-policy.html` (footer punta a `href="#"`)
3. **Termini di Servizio** — stesso problema footer
5. **Google Merchant Center** — free listings su Google Shopping, richiede account verificato + feed prodotti XML/CSV (`/product-feed.xml` da generare una volta verificato l'account)
6. **GA4 eventi conversione** — click CTA, scroll depth, tempo pagina
7. **Meta Pixel** — se si fanno campagne Facebook/Instagram
8. **OG image** — il file esiste (`images/og-image.jpg`, 1200×630) ma non è mai stata testata la condivisione social

## Note tecniche

- **Cache GitHub Pages**: hardcoded 4h — non modificabile via HTML. Per cache lunga: Cloudflare proxy con Edge Cache TTL = 1 anno
- **og:url**: usa sempre trailing slash (`https://mentalpro.it/`)
- Il worktree potrebbe non esistere a inizio sessione — ricrearlo con il comando sopra
