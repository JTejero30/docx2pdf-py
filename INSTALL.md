# Instalación de `docx_a_pdf` (conversión .docx → PDF)

Esta funcionalidad convierte documentos Word (`.docx`) a PDF con **WeasyPrint**
(motor 100% Python a nivel de código). Las **fuentes van empaquetadas** dentro de
la librería, así que **no hay que instalar fuentes**. Lo único que cada máquina
necesita, además de Python, es el **runtime nativo de Pango/GTK** (librerías de
sistema que WeasyPrint usa por debajo). Hay que instalarlo **una sola vez** por
equipo.

> Resumen: `Python ≥ 3.9` + la librería por `pip` + el **runtime de GTK/Pango**.

---

## 1. Requisito de sistema: runtime de Pango/GTK (una vez por equipo)

### Windows
WeasyPrint necesita las DLL de GTK3/Pango. La forma recomendada (oficial de
WeasyPrint):

1. Instala **MSYS2**: https://www.msys2.org
2. Abre la consola *MSYS2 MinGW 64-bit* y ejecuta:
   ```bash
   pacman -S mingw-w64-x86_64-pango
   ```
3. Añade `C:\msys64\mingw64\bin` al **PATH** del sistema (o define
   `WEASYPRINT_DLL_DIRECTORIES=C:\msys64\mingw64\bin`).

> Alternativa: instalar el *GTK3 runtime* desde un instalador
> (https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer).
> MSYS2 es lo más fiable y soportado.

### Linux (Debian / Ubuntu)
```bash
sudo apt install libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz-subset0
```
(Las demás —glib, fontconfig, freetype, harfbuzz— entran como dependencias.)

### Linux (RHEL / Rocky / Fedora)
```bash
sudo dnf install pango
```

### macOS
```bash
brew install pango
```

> **¿Ya está instalado?** Comprueba sin instalar nada:
> ```bash
> python -c "import ctypes.util as u; print({l: bool(u.find_library(l)) for l in ['pango-1.0','pangoft2-1.0','harfbuzz','fontconfig','gobject-2.0']})"
> ```
> Si todo sale `True`, no necesitas instalar el runtime.

---

## 2. La librería Python

```bash
pip install air_pdf        # desde vuestro índice/repositorio interno
```
Arrastra automáticamente `weasyprint` y `lxml` (con sus wheels; sin compilar
nada). Recomendado usar un **entorno virtual** (`python -m venv .venv`), no
requiere permisos de administrador.

---

## 3. Comprobar que funciona

```python
from air_pdf import docx_a_pdf
docx_a_pdf("documento.docx", "salida.pdf")
```

Si falta el runtime de GTK/Pango, al llamar a `docx_a_pdf` verás un error del
tipo `cannot load library 'libpango-1.0...'` (Linux) o `'libgobject-2.0-0'`
(Windows) → vuelve al **paso 1**.

---

## Notas
- **Fidelidad**: WeasyPrint reproduce con alta fidelidad estilos, tablas,
  imágenes, cabeceras/pies, numeración y portada. La **paginación** es muy
  parecida a Word pero **no idéntica** (es otro motor de maquetado). Si el
  `.docx` se ha guardado con Word, conserva las pistas de salto de página y la
  paginación se ajusta mejor.
- **Fuentes**: van incrustadas en el PDF desde las que trae la librería
  (Carlito≈Calibri, Caladea≈Cambria, Gelasio≈Georgia, Liberation≈Arial/Times/
  Courier, Lato, Roboto). No hay que instalar fuentes en el equipo.
- **Sin conexión / sin red**: no llama a ningún servicio; todo es local.
