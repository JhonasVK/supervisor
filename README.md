# Informes COBRA — Repetido Reparado / Averías de Infancia

Dashboards web para dos indicadores de calidad de terreno, publicados en GitHub Pages.

- **Repetido Reparado**: reparaciones que vuelven a fallar dentro de 30 días (meta: no superar 4%).
- **Averías de Infancia**: instalaciones que generan una reparación dentro de su período de infancia (meta: no superar 2.5%).

## Estructura

```
Reiterados/
├── index.html                     ← página de entrada (elige el informe)
├── Dashboard_Reincidencias.html   ← Informe de Repetido Reparado
├── Dashboard_Infancia.html        ← Informe de Averías de Infancia
├── logo-cobra.png
├── generar_reincidencias.js       ← genera el Excel + Dashboard_Reincidencias.html
├── generar_infancia.js            ← genera el Excel + Dashboard_Infancia.html
├── generar_indice.js              ← genera index.html
├── Generar_Reporte_Reincidencias.bat  ← doble clic para correr los 3 scripts
└── bbdd/                          ← CSV de origen (NO se sube al repo, datos sensibles)
```

## Actualizar los datos

1. Reemplaza el CSV correspondiente dentro de `bbdd\` (mismo nombre o con el mes nuevo).
2. Doble clic en `Generar_Reporte_Reincidencias.bat`.
3. Se regeneran los Excel (locales, no se suben) y los 3 archivos HTML.
4. Sube los cambios a GitHub (`git add`, `git commit`, `git push`) para publicar la actualización.

## Privacidad

La carpeta `bbdd\` (CSV con datos de clientes) y los archivos `.xlsx` (incluyen RUT de técnicos)
están excluidos del repositorio vía `.gitignore`. Solo se publican los dashboards HTML, que
muestran nombres de técnicos y métricas de desempeño pero no RUT ni datos de clientes.
