# adidas-informe-macro

Informe macroeconómico interactivo (HTML self-contained + Chart.js) que se
actualiza automáticamente desde un Google Sheet vía GitHub Actions.

## Arquitectura
- `docs/index.html` — deliverable principal, servido por GitHub Pages desde `main`.
- `scripts/fetch_and_build.py` — descarga el Excel de Drive, parsea, y actualiza
  **quirúrgicamente** solo el bloque `var CHARTS = {...};` del HTML (regex no-greedy,
  `count=1`). Nunca regenera el HTML completo.
- `config/sheet_config.json` — mapeo de pestañas del Sheet.
- Toda la data sale de la pestaña **"DATA"** (e IPC anual de **"Data 2"**).

## ⚠️ DATOS PROTEGIDOS — NO TOCAR SIN PEDIDO EXPLÍCITO

El gráfico **`pbg_var`** (Variación PBG 2025 vs. 2022, fuente Equilibra) vive en
`docs/index.html`, fuera del bloque `var CHARTS`, entre los marcadores:

```
// ╔═ DATOS PROTEGIDOS — MANTENIMIENTO MANUAL ... ╗
CHARTS['pbg_var'] = { ... };
// ── FIN DATOS PROTEGIDOS pbg_var ──
```

Reglas:
1. **NO** modificar, regenerar ni reordenar estos datos salvo que el usuario lo
   pida **explícitamente**.
2. En conflictos de **rebase/merge**, **conservar siempre la versión que contiene
   estos datos** (Equilibra, Neuquén +36.9% … Tierra del Fuego −13.7%). Nunca
   resolver con `--theirs`/`--ours` a ciegas sobre este bloque.
3. El script `fetch_and_build.py` **no** toca este bloque (su regex solo matchea
   `var CHARTS = {...};`). Si alguna vez se pierde, fue por un rebase mal resuelto.

## Convenciones
- Insights dinámicos: **sin guion largo** (— / –), para no delatar generación por IA.
- PNG se exporta con **fondo transparente**; JPG con fondo blanco.
- Colores condicionales: `CONDITIONAL_H_INV` → positivo verde, negativo rojo.
