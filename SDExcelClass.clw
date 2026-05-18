    MEMBER()
    INCLUDE('SDExcelClass.INC'),ONCE
    MAP
      SDExcel_Q(STRING pStr),STRING
    END!MAP
!=============================================================================
SDExcel_Q                 PROCEDURE(STRING pStr)
Result  CSTRING(8193)
Pos     LONG
Start   LONG
  CODE
    Result = pStr
    Start  = 1
    LOOP
      Pos = INSTRING('"', Result, 1, Start)
      IF NOT Pos THEN BREAK.
      Result = Result[1 : Pos - 1] & '""' & Result[Pos + 1 : LEN(Result)]
      Start  = Pos + 2
    END!LOOP
    RETURN Result

!=============================================================================
SDExcelClass.Init         PROCEDURE(WINDOW pWindow, SIGNED pControl)
  CODE
    SELF.Window  &= pWindow
    SELF.Control  = pControl
    SELF.Start()

!=============================================================================
SDExcelClass.Start        PROCEDURE()
  CODE
    SELF.Window $ SELF.Control{PROP:Language} = 0409H
    SELF.Window $ SELF.Control{PROP:Create} = 'SDExcelCtrl.SDExcelCtrl'

!=============================================================================
SDExcelClass.Kill         PROCEDURE()
  CODE
    SELF.Window $ SELF.Control{PROP:DeActivate}

!=============================================================================
! Archivos
!=============================================================================
SDExcelClass.CreateFile   PROCEDURE()
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'CreateFile()'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.CreateFile   PROCEDURE(STRING FileName)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'CreateFileAt("' & SDExcel_Q(FileName) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.OpenFile     PROCEDURE(STRING FileName)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'OpenFile("' & SDExcel_Q(FileName) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.Save         PROCEDURE(STRING FileName)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'Save("' & SDExcel_Q(FileName) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.Close        PROCEDURE(BYTE SaveChanges)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'Close(' & SaveChanges & ')'}
    LOCKTHREAD

!=============================================================================
! Seleccion
!=============================================================================
SDExcelClass.Select       PROCEDURE(STRING Range)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'Select("' & SDExcel_Q(Range) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.Select       PROCEDURE(LONG Row, LONG Column)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SelectByPos(' & Row & ',' & Column & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SelectRow    PROCEDURE(LONG Row)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SelectRow(' & Row & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SelectColumn PROCEDURE(STRING Column)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SelectColumn("' & SDExcel_Q(Column) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SelectColumnByNumber PROCEDURE(LONG Column)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SelectColumnByNumber(' & Column & ')'}
    LOCKTHREAD

!=============================================================================
! Escritura
!=============================================================================
SDExcelClass.Assign       PROCEDURE(STRING Contents)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'Assign("' & SDExcel_Q(Contents) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.Assign       PROCEDURE(STRING Range, STRING Contents)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'AssignByRange("' & SDExcel_Q(Range) & '","' & SDExcel_Q(Contents) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.Assign       PROCEDURE(LONG Row, LONG Column, STRING Contents)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'AssignByPos(' & Row & ',' & Column & ',"' & SDExcel_Q(Contents) & '")'}
    LOCKTHREAD

!=============================================================================
! Lectura
!=============================================================================
SDExcelClass.Read         PROCEDURE()
RetVal  STRING(SDExcel:StrMax)
  CODE
    UNLOCKTHREAD
    RetVal = SELF.Window $ SELF.Control{'Read()'}
    LOCKTHREAD
    RETURN RetVal

!=============================================================================
SDExcelClass.Read         PROCEDURE(STRING Cell)
RetVal  STRING(SDExcel:StrMax)
  CODE
    UNLOCKTHREAD
    RetVal = SELF.Window $ SELF.Control{'ReadByRange("' & SDExcel_Q(Cell) & '")'}
    LOCKTHREAD
    RETURN RetVal

!=============================================================================
SDExcelClass.Read         PROCEDURE(LONG Row, LONG Column)
RetVal  STRING(SDExcel:StrMax)
  CODE
    UNLOCKTHREAD
    RetVal = SELF.Window $ SELF.Control{'ReadByPos(' & Row & ',' & Column & ')'}
    LOCKTHREAD
    RETURN RetVal

!=============================================================================
! Formato numerico
!=============================================================================
SDExcelClass.SetNumberFormat PROCEDURE(STRING Format)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetNumberFormat("' & SDExcel_Q(Format) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetNumberFormat PROCEDURE(STRING Range, STRING Format)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetNumberFormatByRange("' & SDExcel_Q(Range) & '","' & SDExcel_Q(Format) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetNumberFormat PROCEDURE(LONG Row, LONG Column, STRING Format)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetNumberFormatByPos(' & Row & ',' & Column & ',"' & SDExcel_Q(Format) & '")'}
    LOCKTHREAD

!=============================================================================
! Alineacion horizontal
!=============================================================================
SDExcelClass.SetHAlignment PROCEDURE(BYTE Alignment)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetHAlignment(' & Alignment & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetHAlignment PROCEDURE(STRING Range, BYTE Alignment)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetHAlignmentByRange("' & SDExcel_Q(Range) & '",' & Alignment & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetHAlignment PROCEDURE(LONG Row, LONG Column, BYTE Alignment)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetHAlignmentByPos(' & Row & ',' & Column & ',' & Alignment & ')'}
    LOCKTHREAD

!=============================================================================
! Alineacion vertical
!=============================================================================
SDExcelClass.SetVAlignment PROCEDURE(BYTE Alignment)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetVAlignment(' & Alignment & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetVAlignment PROCEDURE(STRING Range, BYTE Alignment)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetVAlignmentByRange("' & SDExcel_Q(Range) & '",' & Alignment & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetVAlignment PROCEDURE(LONG Row, LONG Column, BYTE Alignment)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetVAlignmentByPos(' & Row & ',' & Column & ',' & Alignment & ')'}
    LOCKTHREAD

!=============================================================================
! Combinar celdas
!=============================================================================
SDExcelClass.MergeCells   PROCEDURE()
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'MergeCells()'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.MergeCells   PROCEDURE(STRING Range)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'MergeCellsByRange("' & SDExcel_Q(Range) & '")'}
    LOCKTHREAD

!=============================================================================
! Ajuste de texto
!=============================================================================
SDExcelClass.SetWrapText  PROCEDURE(BYTE WrapText)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetWrapText(' & WrapText & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetWrapText  PROCEDURE(STRING Range, BYTE WrapText)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetWrapTextByRange("' & SDExcel_Q(Range) & '",' & WrapText & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetWrapText  PROCEDURE(LONG Row, LONG Column, BYTE WrapText)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetWrapTextByPos(' & Row & ',' & Column & ',' & WrapText & ')'}
    LOCKTHREAD

!=============================================================================
! Ancho de columna
!=============================================================================
SDExcelClass.SetColumnWidth PROCEDURE(LONG Width)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetColumnWidth(' & Width & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetColumnWidth PROCEDURE(STRING Column, LONG Width)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetColumnWidthByName("' & SDExcel_Q(Column) & '",' & Width & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetColumnWidthByNumber PROCEDURE(LONG Column, LONG Width)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetColumnWidthByNumber(' & Column & ',' & Width & ')'}
    LOCKTHREAD

!=============================================================================
! Alto de fila
!=============================================================================
SDExcelClass.SetRowHeight PROCEDURE(LONG Height)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetRowHeight(' & Height & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetRowHeight PROCEDURE(LONG Row, LONG Height)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetRowHeightByNumber(' & Row & ',' & Height & ')'}
    LOCKTHREAD

!=============================================================================
! Hojas
!=============================================================================
SDExcelClass.AddSheet     PROCEDURE(BYTE Count)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'AddSheet(' & Count & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SelectSheet  PROCEDURE(SHORT Sheet)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SelectSheet(' & Sheet & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SelectSheetByName PROCEDURE(STRING Name)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SelectSheetByName("' & SDExcel_Q(Name) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.GetSheetCount PROCEDURE()
RetVal  SHORT
  CODE
    UNLOCKTHREAD
    RetVal = SELF.Window $ SELF.Control{'GetSheetCount()'}
    LOCKTHREAD
    RETURN RetVal

!=============================================================================
SDExcelClass.GetSheetIndex PROCEDURE()
RetVal  LONG
  CODE
    UNLOCKTHREAD
    RetVal = SELF.Window $ SELF.Control{'GetSheetIndex()'}
    LOCKTHREAD
    RETURN RetVal

!=============================================================================
SDExcelClass.GetSheetIndex PROCEDURE(STRING Name)
RetVal  LONG
  CODE
    UNLOCKTHREAD
    RetVal = SELF.Window $ SELF.Control{'GetSheetIndexByName("' & SDExcel_Q(Name) & '")'}
    LOCKTHREAD
    RETURN RetVal

!=============================================================================
SDExcelClass.GetSheetName PROCEDURE()
RetVal  STRING(SDExcel:StrMax)
  CODE
    UNLOCKTHREAD
    RetVal = SELF.Window $ SELF.Control{'GetSheetName()'}
    LOCKTHREAD
    RETURN RetVal

!=============================================================================
SDExcelClass.GetSheetName PROCEDURE(LONG Sheet)
RetVal  STRING(SDExcel:StrMax)
  CODE
    UNLOCKTHREAD
    RetVal = SELF.Window $ SELF.Control{'GetSheetNameByIndex(' & Sheet & ')'}
    LOCKTHREAD
    RETURN RetVal

!=============================================================================
SDExcelClass.SetSheetName PROCEDURE(SHORT Sheet, STRING Name)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetSheetName(' & Sheet & ',"' & SDExcel_Q(Name) & '")'}
    LOCKTHREAD

!=============================================================================
! Info
!=============================================================================
SDExcelClass.About        PROCEDURE()
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'About()'}
    LOCKTHREAD

!=============================================================================
! Fuente
!=============================================================================
SDExcelClass.SetBold      PROCEDURE(BYTE Bold)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBold(' & Bold & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetBoldByRange PROCEDURE(STRING Range, BYTE Bold)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBoldByRange("' & SDExcel_Q(Range) & '",' & Bold & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetBoldByPos PROCEDURE(LONG Row, LONG Column, BYTE Bold)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBoldByPos(' & Row & ',' & Column & ',' & Bold & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetItalic    PROCEDURE(BYTE Italic)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetItalic(' & Italic & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetItalicByRange PROCEDURE(STRING Range, BYTE Italic)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetItalicByRange("' & SDExcel_Q(Range) & '",' & Italic & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetItalicByPos PROCEDURE(LONG Row, LONG Column, BYTE Italic)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetItalicByPos(' & Row & ',' & Column & ',' & Italic & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetFontSize  PROCEDURE(LONG Size)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetFontSize(' & Size & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetFontSizeByRange PROCEDURE(STRING Range, LONG Size)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetFontSizeByRange("' & SDExcel_Q(Range) & '",' & Size & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetFontSizeByPos PROCEDURE(LONG Row, LONG Column, LONG Size)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetFontSizeByPos(' & Row & ',' & Column & ',' & Size & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetFontName  PROCEDURE(STRING Name)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetFontName("' & SDExcel_Q(Name) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetFontNameByRange PROCEDURE(STRING Range, STRING Name)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetFontNameByRange("' & SDExcel_Q(Range) & '","' & SDExcel_Q(Name) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetFontNameByPos PROCEDURE(LONG Row, LONG Column, STRING Name)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetFontNameByPos(' & Row & ',' & Column & ',"' & SDExcel_Q(Name) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetFontColor PROCEDURE(STRING Color)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetFontColor("' & SDExcel_Q(Color) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetFontColorByRange PROCEDURE(STRING Range, STRING Color)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetFontColorByRange("' & SDExcel_Q(Range) & '","' & SDExcel_Q(Color) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetFontColorByPos PROCEDURE(LONG Row, LONG Column, STRING Color)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetFontColorByPos(' & Row & ',' & Column & ',"' & SDExcel_Q(Color) & '")'}
    LOCKTHREAD

!=============================================================================
! Fondo
!=============================================================================
SDExcelClass.SetBgColor   PROCEDURE(STRING Color)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBgColor("' & SDExcel_Q(Color) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetBgColorByRange PROCEDURE(STRING Range, STRING Color)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBgColorByRange("' & SDExcel_Q(Range) & '","' & SDExcel_Q(Color) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetBgColorByPos PROCEDURE(LONG Row, LONG Column, STRING Color)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBgColorByPos(' & Row & ',' & Column & ',"' & SDExcel_Q(Color) & '")'}
    LOCKTHREAD

!=============================================================================
! Asignacion con formato
!=============================================================================
SDExcelClass.AssignWithFormat PROCEDURE(STRING Contents, STRING Format)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'AssignWithFormat("' & SDExcel_Q(Contents) & '","' & SDExcel_Q(Format) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.AssignWithFormatByRange PROCEDURE(STRING Range, STRING Contents, STRING Format)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'AssignWithFormatByRange("' & SDExcel_Q(Range) & '","' & SDExcel_Q(Contents) & '","' & SDExcel_Q(Format) & '")'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.AssignWithFormatByPos PROCEDURE(LONG Row, LONG Column, STRING Contents, STRING Format)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'AssignWithFormatByPos(' & Row & ',' & Column & ',"' & SDExcel_Q(Contents) & '","' & SDExcel_Q(Format) & '")'}
    LOCKTHREAD

!=============================================================================
! Bordes
!=============================================================================
SDExcelClass.SetBorder    PROCEDURE(BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBorder(' & Style & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetTopBorder PROCEDURE(BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetTopBorder(' & Style & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetBottomBorder PROCEDURE(BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBottomBorder(' & Style & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetLeftBorder PROCEDURE(BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetLeftBorder(' & Style & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetRightBorder PROCEDURE(BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetRightBorder(' & Style & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetBorderByRange PROCEDURE(STRING Range, BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBorderByRange("' & SDExcel_Q(Range) & '",' & Style & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetOutsideBorderByRange PROCEDURE(STRING Range, BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetOutsideBorderByRange("' & SDExcel_Q(Range) & '",' & Style & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetInsideBorderByRange PROCEDURE(STRING Range, BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetInsideBorderByRange("' & SDExcel_Q(Range) & '",' & Style & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetTopBorderByRange PROCEDURE(STRING Range, BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetTopBorderByRange("' & SDExcel_Q(Range) & '",' & Style & ')'}
    LOCKTHREAD

!=============================================================================
SDExcelClass.SetBottomBorderByRange PROCEDURE(STRING Range, BYTE Style)
  CODE
    UNLOCKTHREAD
    A$ = SELF.Window $ SELF.Control{'SetBottomBorderByRange("' & SDExcel_Q(Range) & '",' & Style & ')'}
    LOCKTHREAD

!=============================================================================
!=============================================================================
