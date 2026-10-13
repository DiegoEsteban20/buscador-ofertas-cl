# Buscador de ofertas para Chile (PowerShell)

Revisa ~40 fuentes de tiendas chilenas y de juegos digitales, detecta descuentos y **bajas de precio** respecto a la última revisión, y genera un Excel al iniciar sesión.

## Qué hace

- **Solotodo** (API pública): consolas, tarjetas gráficas, teclados, mouse, notebooks, monitores, celulares y más.
- **Falabella** (datos `__NEXT_DATA__` de la página): ropa, zapatillas, poleras de fútbol, relojes, figuras, LEGO, juegos físicos.
- Tiendas **VTEX** (API de catálogo): Reebok, Levi's, Asics, Totto.
- Juegos: **Steam** y **Nintendo eShop Chile** en pesos, **CheapShark** (GOG, Humble, Epic…) y juegos gratis de **Epic** y **GamerPower**.
- Guarda historial de precios y marca **BAJÓ / subió / nuevo**; hoja con las 5 mejores rebajas por categoría.
- Fuentes configurables en `Fuentes.csv` (activar, filtros, páginas) sin tocar el código.

## Uso

```powershell
.\BuscadorOfertas.ps1 -Forzar      # busca ahora
.\BuscadorOfertas.ps1 -Instalar    # automático al iniciar sesión
```

## Conceptos aplicados

Integración de varias fuentes (JSON, HTML con datos embebidos, regex), parser JSON compatible con PowerShell 5.1 para respuestas de más de 2 MB, sondeo previo de qué sitios responden desde la red del usuario, historial de precios y reportes en Excel.
