!ABCIncludeFile
OMIT('_EndOfInclude_',_SDExcelClassPresent_)
_SDExcelClassPresent_  EQUATE(1)
! Alineacion horizontal
SDExcel:Left    EQUATE(1)
SDExcel:Center  EQUATE(2)
SDExcel:Right   EQUATE(3)

! Alineacion vertical
SDExcel:VTop    EQUATE(1)
SDExcel:VCenter EQUATE(2)
SDExcel:VBottom EQUATE(3)

! Longitud maxima para valores STRING retornados
SDExcel:StrMax  EQUATE(8192)

! Estilos de borde
SDExcel:BorderNone   EQUATE(0)
SDExcel:BorderThin   EQUATE(1)
SDExcel:BorderMedium EQUATE(2)
SDExcel:BorderThick  EQUATE(3)

SDExcelClass  CLASS,TYPE,MODULE('SDExcelClass.CLW'),LINK('SDExcelClass.CLW',_ABCLinkMode_),DLL(_ABCDllMode_)
Control               SIGNED
Window                &WINDOW

! Ciclo de vida
Init                  PROCEDURE(WINDOW pWindow, SIGNED pControl),VIRTUAL
Start                 PROCEDURE(),VIRTUAL
Kill                  PROCEDURE(),VIRTUAL

! Archivos
CreateFile            PROCEDURE(),VIRTUAL
CreateFile            PROCEDURE(STRING FileName),VIRTUAL
OpenFile              PROCEDURE(STRING FileName),VIRTUAL
Save                  PROCEDURE(STRING FileName),VIRTUAL
Close                 PROCEDURE(BYTE SaveChanges=1),VIRTUAL

! Seleccion
Select                PROCEDURE(STRING Range),VIRTUAL
Select                PROCEDURE(LONG Row, LONG Column),VIRTUAL
SelectRow             PROCEDURE(LONG Row),VIRTUAL
SelectColumn          PROCEDURE(STRING Column),VIRTUAL
SelectColumnByNumber  PROCEDURE(LONG Column),VIRTUAL

! Escritura  (sobrecargada como MSExcelClass)
Assign                PROCEDURE(STRING Contents),VIRTUAL
Assign                PROCEDURE(STRING Range, STRING Contents),VIRTUAL
Assign                PROCEDURE(LONG Row, LONG Column, STRING Contents),VIRTUAL

! Lectura  (sobrecargada como MSExcelClass)
Read                  PROCEDURE(),STRING,VIRTUAL
Read                  PROCEDURE(STRING Cell),STRING,VIRTUAL
Read                  PROCEDURE(LONG Row, LONG Column),STRING,VIRTUAL

! Formato numerico
SetNumberFormat       PROCEDURE(STRING Format),VIRTUAL
SetNumberFormat       PROCEDURE(STRING Range, STRING Format),VIRTUAL
SetNumberFormat       PROCEDURE(LONG Row, LONG Column, STRING Format),VIRTUAL

! Alineacion horizontal
SetHAlignment         PROCEDURE(BYTE Alignment),VIRTUAL
SetHAlignment         PROCEDURE(STRING Range, BYTE Alignment),VIRTUAL
SetHAlignment         PROCEDURE(LONG Row, LONG Column, BYTE Alignment),VIRTUAL

! Alineacion vertical
SetVAlignment         PROCEDURE(BYTE Alignment),VIRTUAL
SetVAlignment         PROCEDURE(STRING Range, BYTE Alignment),VIRTUAL
SetVAlignment         PROCEDURE(LONG Row, LONG Column, BYTE Alignment),VIRTUAL

! Combinar celdas
MergeCells            PROCEDURE(),VIRTUAL
MergeCells            PROCEDURE(STRING Range),VIRTUAL

! Ajuste de texto
SetWrapText           PROCEDURE(BYTE WrapText),VIRTUAL
SetWrapText           PROCEDURE(STRING Range, BYTE WrapText),VIRTUAL
SetWrapText           PROCEDURE(LONG Row, LONG Column, BYTE WrapText),VIRTUAL

! Ancho de columna  (entero; el COM lo convierte a double internamente)
SetColumnWidth        PROCEDURE(LONG Width),VIRTUAL
SetColumnWidth        PROCEDURE(STRING Column, LONG Width),VIRTUAL
SetColumnWidthByNumber PROCEDURE(LONG Column, LONG Width),VIRTUAL

! Alto de fila
SetRowHeight          PROCEDURE(LONG Height),VIRTUAL
SetRowHeight          PROCEDURE(LONG Row, LONG Height),VIRTUAL

! Hojas
AddSheet              PROCEDURE(BYTE Count=1),VIRTUAL
SelectSheet           PROCEDURE(SHORT Sheet),VIRTUAL
SelectSheetByName     PROCEDURE(STRING Name),VIRTUAL
GetSheetCount         PROCEDURE(),SHORT,VIRTUAL
GetSheetIndex         PROCEDURE(),LONG,VIRTUAL
GetSheetIndex         PROCEDURE(STRING Name),LONG,VIRTUAL
GetSheetName          PROCEDURE(),STRING,VIRTUAL
GetSheetName          PROCEDURE(LONG Sheet),STRING,VIRTUAL
SetSheetName          PROCEDURE(SHORT Sheet, STRING Name),VIRTUAL

! Info
About                 PROCEDURE(),VIRTUAL

! Fuente
SetBold               PROCEDURE(BYTE Bold),VIRTUAL
SetBoldByRange        PROCEDURE(STRING Range, BYTE Bold),VIRTUAL
SetBoldByPos          PROCEDURE(LONG Row, LONG Column, BYTE Bold),VIRTUAL
SetItalic             PROCEDURE(BYTE Italic),VIRTUAL
SetItalicByRange      PROCEDURE(STRING Range, BYTE Italic),VIRTUAL
SetItalicByPos        PROCEDURE(LONG Row, LONG Column, BYTE Italic),VIRTUAL
SetFontSize           PROCEDURE(LONG Size),VIRTUAL
SetFontSizeByRange    PROCEDURE(STRING Range, LONG Size),VIRTUAL
SetFontSizeByPos      PROCEDURE(LONG Row, LONG Column, LONG Size),VIRTUAL
SetFontName           PROCEDURE(STRING Name),VIRTUAL
SetFontNameByRange    PROCEDURE(STRING Range, STRING Name),VIRTUAL
SetFontNameByPos      PROCEDURE(LONG Row, LONG Column, STRING Name),VIRTUAL
SetFontColor          PROCEDURE(STRING Color),VIRTUAL
SetFontColorByRange   PROCEDURE(STRING Range, STRING Color),VIRTUAL
SetFontColorByPos     PROCEDURE(LONG Row, LONG Column, STRING Color),VIRTUAL

! Fondo
SetBgColor            PROCEDURE(STRING Color),VIRTUAL
SetBgColorByRange     PROCEDURE(STRING Range, STRING Color),VIRTUAL
SetBgColorByPos       PROCEDURE(LONG Row, LONG Column, STRING Color),VIRTUAL

! Asignacion con formato
AssignWithFormat        PROCEDURE(STRING Contents, STRING Format),VIRTUAL
AssignWithFormatByRange PROCEDURE(STRING Range, STRING Contents, STRING Format),VIRTUAL
AssignWithFormatByPos   PROCEDURE(LONG Row, LONG Column, STRING Contents, STRING Format),VIRTUAL

! Bordes
SetBorder               PROCEDURE(BYTE Style),VIRTUAL
SetTopBorder            PROCEDURE(BYTE Style),VIRTUAL
SetBottomBorder         PROCEDURE(BYTE Style),VIRTUAL
SetLeftBorder           PROCEDURE(BYTE Style),VIRTUAL
SetRightBorder          PROCEDURE(BYTE Style),VIRTUAL
SetBorderByRange        PROCEDURE(STRING Range, BYTE Style),VIRTUAL
SetOutsideBorderByRange PROCEDURE(STRING Range, BYTE Style),VIRTUAL
SetInsideBorderByRange  PROCEDURE(STRING Range, BYTE Style),VIRTUAL
SetTopBorderByRange     PROCEDURE(STRING Range, BYTE Style),VIRTUAL
SetBottomBorderByRange  PROCEDURE(STRING Range, BYTE Style),VIRTUAL

              END!SDExcelClass

_EndOfInclude_
