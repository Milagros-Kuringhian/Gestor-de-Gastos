# Design: Tipografía minimalista + limpieza visual

Fecha: 2026-07-22  
Estado: Aprobado (enfoque 2)

## Objetivo

Hacer la PWA más minimalista y simple, cambiando principalmente la tipografía. La paleta de colores se mantiene.

## Decisiones

- **UI:** Plus Jakarta Sans (pesos 400, 500, 600, 700)
- **Números:** misma tipografía que el texto (sin monoespaciada), con `font-variant-numeric: tabular-nums`
- **Enfoque:** tipografía + limpieza ligera (no rediseño profundo)

## Cambios de tipografía

- Reemplazar Outfit e IBM Plex Mono por Plus Jakarta Sans
- Actualizar el link de Google Fonts en `pwa/index.html`
- Eliminar `--mono` o unificarlo con `--font`
- Montos (balance, historial, modal) heredan la UI font + tabular nums
- Reducir pesos 700 → 600 donde no aporten jerarquía clara

## Limpieza visual (colores intactos)

- Sombra más suave (`--shadow`)
- Grano más sutil (menor opacity)
- Radios un poco más contenidos (`--radius`, balance card, modal)
- Cards/botones con menos peso de sombra

## Fuera de alcance

- Cambios de layout, flujos o lógica JS
- Nueva paleta de colores
- Remoción total de gradientes del fondo o de la balance card

## Archivos

- `pwa/index.html` — carga de fuentes
- `pwa/styles.css` — tokens y tipografía/sombras/radios
