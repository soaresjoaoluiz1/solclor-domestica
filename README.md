# Solclor Piscinas — Landing Page

Landing page do programa **Revendedor Autorizado Solclor** (linha de produtos para piscinas e domésticos).

**Produção:** https://piscinas.solclor.com.br

## Stack

- HTML + CSS + JS (single-file, sem build system)
- Imagens em WebP com fallback JPG (`<picture>` + `image-set()`)
- Google Apps Script pra webhook do formulário
- Deploy: VPS HostGator (cPanel + git pull)

## Estrutura

```
solclor-piscinas/
├── index.html             # LP completa (14 sections)
├── assets/                # imagens otimizadas
│   ├── favicon.png
│   ├── logo-solclor-produtos-limpeza.{webp,png}     # logo header
│   ├── logo-solclor-rodape.{webp,png}               # logo footer + menu mobile
│   ├── banner-solclor-produtos-limpeza.{webp,jpg}   # bg hero desktop
│   ├── bg-mobile-solclor-hero.{webp,jpg}            # bg hero mobile
│   └── bg-portfolio-piscinas.{webp,jpg}             # bg section piscinas
└── .gitignore
```

## Deploy

```bash
# na VPS Dros HostGator
cd /home/solclorpiscinas/public_html/
git pull origin main
```

## Otimizações aplicadas

- SEO completo (meta tags, Open Graph, Twitter Cards, JSON-LD Organization + Product + FAQPage)
- Imagens otimizadas (12 MB → ~470 KB pra 98% dos browsers modernos)
- `content-visibility: auto` nas sections abaixo da dobra
- Intersection Observer pra reveal on scroll
- `prefers-reduced-motion` respeitado
- Preload crítico das imagens above-the-fold
- Cabeçalho canonical, robots, hreflang pt-BR
