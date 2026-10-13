# Buscador de ofertas para Chile (PowerShell)

Revisa unas 40 fuentes de tiendas chilenas y de juegos digitales, detecta descuentos y **bajas de precio respecto a la última revisión**, y arma un Excel al iniciar sesión en Windows. Pensado para cazar ofertas de tecnología, ropa, coleccionables y videojuegos sin revisar tienda por tienda.

Funciona en Windows PowerShell 5.1 o PowerShell 7.

---

## Tabla de contenidos
- [Para qué sirve](#para-qué-sirve)
- [Qué revisa](#qué-revisa)
- [Qué hace, paso a paso](#qué-hace-paso-a-paso)
- [Instalación](#instalación)
- [Uso](#uso)
- [El Excel que genera](#el-excel-que-genera)
- [Configuración (Fuentes.csv)](#configuración-fuentescsv)
- [Dejarlo corriendo solo](#dejarlo-corriendo-solo)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Preguntas frecuentes](#preguntas-frecuentes)
- [Conceptos aplicados](#conceptos-aplicados)
- [Aviso](#aviso)

---

## Para qué sirve

Las ofertas buenas duran poco y están repartidas en muchas tiendas. Este script revisa todas de una, se acuerda de los precios de la vez anterior y te marca lo que **realmente bajó** (no lo que siempre dice "oferta"). Así, cada vez que prendes el PC, tienes un Excel con los mejores descuentos del día sin buscar nada.

---

## Qué revisa

- **Solotodo** (API pública): consolas, tarjetas gráficas, teclados, mouse, audífonos, notebooks, monitores, celulares, smartwatch, sillas gamer, TV, SSD, parlantes y más.
- **Falabella** (datos embebidos de la página): ropa de hombre, zapatillas, poleras de fútbol, relojes, perfumes, mochilas, figuras de anime, Funko, LEGO y juegos físicos.
- Tiendas **VTEX** (API de catálogo): Reebok, Levi's, Asics, Totto.
- **Juegos digitales**: Steam y Nintendo eShop Chile en pesos, y CheapShark (GOG, Humble, Epic, Fanatical…) en dólares.
- **Juegos gratis**: Epic Games y GamerPower.

---

## Qué hace, paso a paso

1. Recorre cada fuente activa de `Fuentes.csv`.
2. Trae los productos y sus precios actuales (indica si el precio es con tarjeta, quién vende, o si es gratis).
3. **Compara con la última revisión** y marca cada producto como **BAJÓ / subió / nuevo**.
4. Arma un Excel y lo abre solo, resaltando los mejores descuentos.
5. Guarda el historial de precios para la próxima comparación.

---

## Instalación

No necesitas instalar nada aparte. La primera vez, el módulo `ImportExcel` (para crear el Excel) se instala solo.

Doble clic en `BuscarAhora.cmd`.

---

## Uso

```powershell
.\BuscadorOfertas.ps1 -Forzar          # busca ahora, aunque haya buscado hace poco
.\BuscadorOfertas.ps1 -Instalar        # automático al iniciar sesión
.\BuscadorOfertas.ps1 -Solo falabella  # prueba solo las fuentes que contengan ese texto
.\BuscadorOfertas.ps1 -NoAbrir         # no abre el Excel al terminar
```

---

## El Excel que genera

| Hoja | Contenido |
| --- | --- |
| **Mejores** | Las 5 mayores rebajas de cada categoría (primero lo que bajó) |
| **Bajaron de precio** | Todo lo que está más barato que la vez anterior |
| **Todas** | Todos los productos encontrados |
| **Fuentes con error** | Tiendas que no respondieron (si hubo) |

Verde = 40 % o más de descuento. Amarillo = 15 % a 39 %.

---

## Configuración (Fuentes.csv)

Se abre con Excel o el Bloc de notas. Una fila por fuente:

- **Activo:** `SI` / `NO` para encender o apagar una fuente.
- **Falabella:** cambia lo que va después de `Ntt=` en la URL (ej. `Ntt=parrilla+electrica`).
- **Solotodo:** cambia el número de categoría en la URL.
- **Filtro:** palabras que debe tener el nombre del producto (separadas por `|`).
- **Paginas:** cuántas páginas revisar (más páginas = más productos, más lento).

---

## Dejarlo corriendo solo

`.\BuscadorOfertas.ps1 -Instalar` crea una tarea programada que, cada vez que inicias sesión (un par de minutos después), revisa las tiendas, arma el Excel en tu Escritorio y lo abre solo.

---

## Estructura del proyecto

```
buscador-ofertas-cl/
├─ BuscadorOfertas.ps1   # script principal
├─ BuscarAhora.cmd       # buscar ahora
├─ Fuentes.csv           # las ~40 fuentes, editable
├─ LEEME.txt             # ayuda rápida en español
└─ README.md
```

---

## Preguntas frecuentes

**Una tienda no trae nada.** Algunas tiendas bloquean las consultas automáticas; por eso ya no están adidas.cl, Nike, Decathlon ni PlayStation Store (esta última carga sus ofertas con JavaScript y no se puede leer). Puedes desactivar en `Fuentes.csv` cualquiera que falle.

**Quiero seguir un producto específico.** Agrega una fuente con el buscador de la tienda y un filtro con el nombre del producto.

**Los precios con tarjeta.** El Excel indica cuándo un precio es con tarjeta (por ejemplo CMR) para que no te confundas.

---

## Conceptos aplicados

Integración de varias fuentes distintas (JSON, HTML con datos embebidos, regex), parser JSON compatible con PowerShell 5.1 para respuestas de más de 2 MB, sondeo previo de qué sitios responden desde la red del usuario, historial de precios y reportes en Excel con formato condicional.

---

## Aviso

Herramienta de uso personal. Respeta los términos de cada tienda; no está pensada para consultas masivas ni para revender datos.
