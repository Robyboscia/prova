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
- GA4 Measurement ID: `G-F5N0FTWSCL`
- Caricamento differito: si attiva al primo evento utente oppure dopo 3 secondi
- `anonymize_ip: true`, `cookie_flags: SameSite=None;Secure`

### Favicon & Icone
- `favicon.ico` (32×32), `favicon.svg`, `apple-touch-icon.png` (180×180)
- `icon-192.png`, `icon-512.png` per Android/PWA
- `site.webmanifest` collegato in `<head>`

### Tool esterni verificati
- **Google Search Console**: verificato, sitemap elaborata, 1 pagina indicizzata
- **Bing Webmaster Tools**: verificato via import GSC, sitemap inviata manualmente

## Da fare (non ancora implementato)

1. **Cookie consent banner** — obbligatorio GDPR/Garante per GA4
2. **Privacy Policy** — `/privacy-policy.html` (footer punta a `href="#"`)
3. **Termini di Servizio** — stesso problema footer
4. **Google Business Profile** — scheda brand su Google Maps
5. **GA4 eventi conversione** — click CTA, scroll depth, tempo pagina
6. **Meta Pixel** — se si fanno campagne Facebook/Instagram
7. **OG image** — il file esiste (`images/og-image.jpg`, 1200×630) ma non è mai stata testata la condivisione social

## Note tecniche

- **Cache GitHub Pages**: hardcoded 4h — non modificabile via HTML. Per cache lunga: Cloudflare proxy con Edge Cache TTL = 1 anno
- **og:url**: usa sempre trailing slash (`https://mentalpro.it/`)
- Il worktree potrebbe non esistere a inizio sessione — ricrearlo con il comando sopra
