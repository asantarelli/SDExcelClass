# SDExcelClass

Wrapper Clarion para generacion de archivos Excel sin Microsoft Excel instalado.

Replica la API de **MSExcelClass / MSOfficeClass** de Softmasters, utilizando **SDExcelCtrl** — un control COM OLE basado en [ClosedXML 0.97.0](https://github.com/ClosedXML/ClosedXML) — en lugar de Microsoft Excel.

---

## Caracteristicas

- Genera archivos `.xlsx` nativos sin necesitar Microsoft Excel.
- API identica a MSExcelClass: mismos nombres de metodos, mismos patrones de uso.
- Usa OLE estandar de Clarion (`Window $ Control{'Metodo(args)'}`) — sin UltimateCOM ni librerias de terceros.
- Template Clarion incluido para integracion con el IDE.
- Deployment simple: 18 DLLs junto al EXE, sin registro COM en el sistema.
- Compatible con Clarion 11 (x86).
- Licencia MIT — uso libre en proyectos comerciales.

---

## Requisitos

| Componente | Version |
|---|---|
| Clarion | 11+ (x86) |
| Windows | 10 / 11 |
| .NET Framework | 4.7.2 (incluido en Windows 10/11) |
| SDExcelCtrl.dll | Incluido — ClosedXML 0.97.0 |

---

## Instalacion

### 1. Archivos Clarion

Copiar a `<Clarion>\LibSrc\Win\`:

```
SDExcelClass.inc
SDExcelClass.clw
```

### 2. Template

Copiar `SDExcelClass.tpl` a `<Clarion>\Template\` y registrar en el IDE:  
*Tools → Template Registry → Register*

### 3. DLLs

Copiar los siguientes archivos al directorio del `.EXE` de la aplicacion:

```
SDExcelCtrl.dll
ClosedXML.dll
ClosedXML.Parser.dll
DocumentFormat.OpenXml.dll
DocumentFormat.OpenXml.Framework.dll
ExcelNumberFormat.dll
Irony.dll
Microsoft.Bcl.HashCode.dll
RBush.dll
SixLabors.Fonts.dll
System.Buffers.dll
System.Diagnostics.DiagnosticSource.dll
System.IO.Packaging.dll
System.Memory.dll
System.Numerics.Vectors.dll
System.Runtime.CompilerServices.Unsafe.dll
System.ValueTuple.dll
XLParser.dll
```

> El control utiliza Registration-Free COM (manifest embebido). No se necesita `regasm` ni permisos de administrador.

> **Binding redirect requerido:** cada `.EXE` necesita un `.exe.config` con redirect de `System.Numerics.Vectors` (oldVersion="0.0.0.0-4.1.4.0" → newVersion="4.1.4.0", publicKeyToken=b03f5f7f11d50a3a).

---

## Uso rapido

### En el IDE

1. Agregar la extension global **SD Excel Extension Global** al nodo Global de la app.
2. En la ventana donde se va a usar Excel, agregar el control **SD Excel - SDExcelClass**.
3. En el `CODE` de la ventana, llamar al embed `SDExcelInit` antes de usar el objeto.
4. Al cerrar la ventana, llamar al embed `SDExcelKill`.

### Codigo de ejemplo

```clarion
! Inicializacion (via embed SDExcelInit)
SDExcel.Init(SELF, ?SDExcel)

! Crear workbook y escribir datos
SDExcel.CreateFile()
SDExcel.Select('A1')
SDExcel.Assign('Producto')
SDExcel.Select(1, 2)
SDExcel.Assign('Precio')

SDExcel.Assign(2, 1, 'Notebook')
SDExcel.Assign(2, 2, '1500.00')
SDExcel.SetNumberFormat(2, 2, '#,##0.00')
SDExcel.SetHAlignment(2, 2, SDExcel:Right)

! Formato de cabecera
SDExcel.Select('A1:B1')
SDExcel.SetHAlignment(SDExcel:Center)
SDExcel.SetColumnWidth('A', 30)
SDExcel.SetColumnWidth('B', 15)

! Guardar y cerrar
SDExcel.Save('C:\Reportes\Productos.xlsx')
SDExcel.Close(0)

! Cierre del control (via embed SDExcelKill)
SDExcel.Kill()
```

---

## Referencia rapida de la API

### Archivos
| Metodo | Descripcion |
|---|---|
| `CreateFile()` | Nuevo workbook en memoria |
| `CreateFile(FileName)` | Nuevo workbook y guarda |
| `OpenFile(FileName)` | Abre un `.xlsx` existente |
| `Save(FileName)` | Guarda el workbook |
| `Close(SaveChanges)` | Cierra el workbook |

### Seleccion
| Metodo | Descripcion |
|---|---|
| `Select(Range)` | Por rango (`'A1:C5'`) |
| `Select(Row, Column)` | Por posicion (1-based) |
| `SelectRow(Row)` | Fila completa |
| `SelectColumn(Column)` | Columna por letra |
| `SelectColumnByNumber(Column)` | Columna por numero |

### Escritura y lectura
| Metodo | Descripcion |
|---|---|
| `Assign(Contents)` | Escribe en la celda seleccionada |
| `Assign(Range, Contents)` | Escribe por rango |
| `Assign(Row, Col, Contents)` | Escribe por posicion |
| `Read()` | Lee celda seleccionada |
| `Read(Cell)` | Lee por rango |
| `Read(Row, Col)` | Lee por posicion |

### Formato
| Metodo | Descripcion |
|---|---|
| `SetNumberFormat(Format)` | Formato numerico |
| `SetHAlignment(Alignment)` | Alineacion horizontal |
| `SetVAlignment(Alignment)` | Alineacion vertical |
| `MergeCells()` / `MergeCells(Range)` | Combinar celdas |
| `SetWrapText(WrapText)` | Ajuste de texto |
| `SetColumnWidth(Width)` | Ancho de columna |
| `SetRowHeight(Height)` | Alto de fila |
| `AssignWithFormat(Contents, Format)` | Escribe valor y aplica formato en una llamada |

### Fuente
| Metodo | Descripcion |
|---|---|
| `SetBold(Bold)` / `ByRange` / `ByPos` | Negrita (1=si, 0=no) |
| `SetItalic(Italic)` / `ByRange` / `ByPos` | Italica |
| `SetFontSize(Size)` / `ByRange` / `ByPos` | Tamanio en puntos (LONG) |
| `SetFontName(Name)` / `ByRange` / `ByPos` | Nombre de fuente |
| `SetFontColor(Color)` / `ByRange` / `ByPos` | Color `'#RRGGBB'` |

### Fondo
| Metodo | Descripcion |
|---|---|
| `SetBgColor(Color)` / `ByRange` / `ByPos` | Color de fondo `'#RRGGBB'` |

### Bordes
| Metodo | Descripcion |
|---|---|
| `SetBorder(Style)` | Los 4 lados de la celda seleccionada |
| `SetTopBorder(Style)` | Solo superior |
| `SetBottomBorder(Style)` | Solo inferior — subrayados |
| `SetLeftBorder(Style)` | Solo izquierdo |
| `SetRightBorder(Style)` | Solo derecho |
| `SetBorderByRange(Range, Style)` | Todos los bordes del rango (ext + int) |
| `SetOutsideBorderByRange(Range, Style)` | Solo borde exterior |
| `SetInsideBorderByRange(Range, Style)` | Solo bordes interiores |
| `SetTopBorderByRange(Range, Style)` | Solo superior del rango |
| `SetBottomBorderByRange(Range, Style)` | Solo inferior del rango |

### Hojas
| Metodo | Descripcion |
|---|---|
| `AddSheet(Count)` | Agrega hojas |
| `SelectSheet(Sheet)` | Activa hoja por numero |
| `SelectSheetByName(Name)` | Activa hoja por nombre |
| `GetSheetCount()` | Total de hojas |
| `GetSheetName()` / `SetSheetName(...)` | Nombre de hoja |

### Constantes
```clarion
SDExcel:Left         ! 1 - Alineacion horizontal izquierda
SDExcel:Center       ! 2 - Alineacion horizontal centrada
SDExcel:Right        ! 3 - Alineacion horizontal derecha
SDExcel:VTop         ! 1 - Alineacion vertical superior
SDExcel:VCenter      ! 2 - Alineacion vertical centrada
SDExcel:VBottom      ! 3 - Alineacion vertical inferior
SDExcel:BorderNone   ! 0 - Sin borde
SDExcel:BorderThin   ! 1 - Borde fino
SDExcel:BorderMedium ! 2 - Borde medio
SDExcel:BorderThick  ! 3 - Borde grueso
```

---

## Nota crítica: Build All en Clarion

> **Después de actualizar `SDExcelClass.inc` o `SDExcelClass.clw`, siempre usar Build → Build All (recompilación completa) en el IDE Clarion — nunca Build incremental.**

Clarion mantiene la VMT (Virtual Method Table) de la clase por módulo. Si se agrega, elimina o reordena cualquier método en el INC y el proyecto se compila de forma incremental, los módulos que llaman a SDExcel quedan con una VMT desactualizada. El resultado es GPF en tiempo de ejecución: las llamadas a métodos aterrizan en la función equivocada, con parámetros del stack que no corresponden.

**Regla:** cambio en INC/CLW → ConvertToAnsi.ps1 → Build All en cada app que use SDExcelClass.

---

## Diferencias respecto a MSExcelClass

| Aspecto | MSExcelClass | SDExcelClass |
|---|---|---|
| Motor | Microsoft Excel (COM) | ClosedXML via SDExcelCtrl |
| Requiere Excel | Si | No |
| Herencia | MSOfficeClass | Standalone |
| PROP:Language | No aplica | Requerido antes de PROP:Create |
| SetColumnWidth / SetRowHeight / SetFontSize | REAL | LONG |
| Assign(Range/Pos, ...) | No mueve cursor | Mueve cursor a la celda target |
| Comunicacion COM | SendCommand / GetProperty | Llamadas OLE directas |

---

## Estructura del proyecto

```
SDExcelOleCtrl/
  SDExcelClass.inc       -- Declaracion de la clase (Clarion)
  SDExcelClass.clw       -- Implementacion (Clarion)
  SDExcelClass.tpl       -- Template para el IDE Clarion
  SDExcelClass.html      -- Documentacion de ayuda
  ConvertToAnsi.ps1      -- Utilidad: convierte a ANSI+CRLF y copia a LibSrc\Win
  README.md
  LICENSE
```

---

## Licencia

MIT — ver [LICENSE](LICENSE).

---

## Donaciones
Si esta extensión te es de utilidad, puede dejar una pequeña contribución en mi cuenta de PayPal informacion@sdigitales.com.ar

---

*SD Digitales — 2026*
