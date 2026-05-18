#!-----------------------------------------------------------------------------
#! SDExcelClass.TPL  -  Template para SDExcelClass
#! Patron identico a SoftMasters Office Templates
#!
#! Instalacion:
#!   Copiar SDExcelClass.tpl a la carpeta de templates de Clarion
#!   Copiar SDExcelClass.inc y SDExcelClass.clw a Clarion\LibSrc\Win
#!   Registrar el template en el IDE (Template Registry)
#!-----------------------------------------------------------------------------
#TEMPLATE(SDExcelTemplates,'SD Excel - Control OLE (ClosedXML, sin MS Excel)'), FAMILY('ABC')
#EXTENSION(SDExcelGlobal,'SD Excel Extension Global'),Application
#SHEET
   #TAB('General')
      #DISPLAY('SD Excel - Control OLE (ClosedXML, sin MS Excel)')
   #endtab
#ENDSHEET
#AT(%AfterGlobalIncludes)
include('SDExcelClass.inc'),ONCE
#ENDAT

#AT(%GlobalData)
SDExcel SDExcelClass
#ENDAT
     
#!-----------------------------------------------------------------------------
#! CONTROL: SDExcelControl
#! Coloca un control OLE invisible en la ventana.
#! El programador luego usa el objeto en el codigo.
#!-----------------------------------------------------------------------------
#CONTROL(SDExcelControl,'SD Excel - SDExcelClass'),WRAP(OLE),MULTI
     CONTROLS
       OLE,AT(,,1,1),USE(?SDExcel),TRN,HIDE
       END
     END

#SHEET
  #TAB('&Acerca...')
    #BOXED('SD Excel - SDExcelClass'),SECTION
      #DISPLAY('')
      #DISPLAY('SDExcelClass')
      #DISPLAY('Control OLE para Excel sin Microsoft Excel')
      #DISPLAY('Utiliza ClosedXML (RegFree COM, x86)')
      #DISPLAY('')
      #DISPLAY('SD Digitales  -  2026')
    #ENDBOXED
  #ENDTAB
#ENDSHEET

#ATSTART
  #FIND(%ControlInstance,%ActiveTemplateInstance,%Control)
  #EQUATE(%SDExcelControl,%Control)
#ENDAT

#AT(%AfterWindowOpening)
SDExcel.Window &= %Window
SDExcel.Control = %SDExcelControl
#ENDAT

#CODE(SDExcelInit,'SD Excel - Init (activar control OLE)'),REQ(SDExcelControl)
SDExcel.Init(%Window,%SDExcelControl)

#CODE(SDExcelKill,'SD Excel - Kill (desactivar control OLE)'),REQ(SDExcelControl)
SDExcel.Kill()
