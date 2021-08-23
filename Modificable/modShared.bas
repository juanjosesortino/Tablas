Attribute VB_Name = "modShared"
Option Explicit
 
' declaraciones necesarias para obtener información acerca de las impresoras
 
Public Type PRINTER_INFO_1
   flags As Long
   prescription As Long
   Pane As Long
   Comment As Long
End Type
 
Private Type PRINTER_INFO_2
    pServerName As Long
    pPrinterName As Long
    pShareName As Long
    pPortName As Long
    pDriverName As Long
    pComment As Long
    pLocation As Long
    pDevMode As Long
    pSepFile As Long
    pPrintProcessor As Long
    pDatatype As Long
    pParameters As Long
    pSecurityDescriptor As Long
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type

Public Type PRINTER_INFO_4
   pPrinterName As Long
   pServerName As Long
   Attributes As Long
End Type


Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Boolean
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long)
Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
 
Public Declare Function GetDesktopWindow Lib "user32" () As Long     'Inc. 47913 pto 1
Public Declare Function ShellExecute Lib "shell32" _
    Alias "ShellExecuteA" _
   (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
  

'-------
'SIZEOFxxx are non-windows constants defined for this method
Public Const SIZEOFPRINTER_INFO_1 = 16
Public Const SIZEOFPRINTER_INFO_4 = 12

Public Const PRINTER_LEVEL1 = &H1
Public Const PRINTER_LEVEL4 = &H4

'EnumPrinters enumerates available printers,
'print servers, domains, or print providers.
Public Declare Function EnumPrinters Lib "winspool.drv" _
   Alias "EnumPrintersA" _
  (ByVal flags As Long, _
   ByVal Name As String, _
   ByVal Level As Long, _
   pPrinterEnum As Any, _
   ByVal cdBuf As Long, _
   pcbNeeded As Long, _
   pcReturned As Long) As Long

'EnumPrinters Parameters:
'Flags - Specifies the types of print objects that the function should enumerate.
Public Const PRINTER_ENUM_DEFAULT = &H1     'Windows 95: The function returns
                                            'information about the default printer.
Public Const PRINTER_ENUM_LOCAL = &H2       'function ignores the Name parameter,
                                            'and enumerates the locally installed
                                            'printers. Windows 95: The function will
                                            'also enumerate network printers because
                                            'they are handled by the local print provider
Public Const PRINTER_ENUM_CONNECTIONS = &H4 'Windows NT/2000: The function enumerates the
                                            'list of printers to which the user has made
                                            'previous connections
Public Const PRINTER_ENUM_NAME = &H8        'enumerates the printer identified by Name.
                                            'This can be a server, a domain, or a print
                                            'provider. If Name is NULL, the function
                                            'enumerates available print providers
Public Const PRINTER_ENUM_REMOTE = &H10     'Windows NT/2000: The function enumerates network
                                            'printers and print servers in the computer's domain.
                                            'This value is valid only if Level is 1
Public Const PRINTER_ENUM_SHARED = &H20     'enumerates printers that have the shared attribute.
                                            'Cannot be used in isolation; use an OR operation
                                            'to combine with another PRINTER_ENUM type
Public Const PRINTER_ENUM_NETWORK = &H40    'Windows NT/2000: The function enumerates network
                                            'printers in the computer's domain. This value is
                                            'valid only if Level is 1.

'''''''''''''''''''''''
'Name:
'If Level is 1, Flags contains PRINTER_ENUM_NAME, and Name is non-NULL,
'then Name is a pointer to a null-terminated string that specifies the
'name of the object to enumerate. This string can be the name of a server,
'a domain, or a print provider.
'
'If Level is 1, Flags contains PRINTER_ENUM_NAME, and Name is NULL, then
'the function enumerates the available print providers.
'
'If Level is 1, Flags contains PRINTER_ENUM_REMOTE, and Name is NULL, then
'the function enumerates the printers in the user's domain.
'
'If Level is 2 or 5, Name is a pointer to a null-terminated string that
'specifies the name of a server whose printers are to be enumerated. If
'this string is NULL, then the function enumerates the printers installed
'on the local machine.
'
'If Level is 4, Name should be NULL. The function always queries on
'the local machine.

'When Name is NULL, it enumerates printers that are installed on the
'local machine. These printers include those that are physically attached
'to the local machine as well as remote printers to which it has a
'network connection.

'''''''''''''''''''''''
'Level:
'Specifies the type of data structures pointed to by pPRinterEnum.
'Valid values are 1, 2, 4, and 5, which correspond to the
'PRINTER_INFO_1, PRINTER_INFO_2, PRINTER_INFO_4, and PRINTER_INFO_5
'data structures.
'
'Windows 95: The value can be 1, 2, or 5.
'
'Windows NT/Windows 2000: This value can be 1, 2, 4, or 5.

'''''''''''''''''''''''
'pPRinterEnum:
'Pointer to a buffer that receives an array of PRINTER_INFO_1,
'PRINTER_INFO_2, PRINTER_INFO_4, or PRINTER_INFO_5 structures.
'Each structure contains data that describes an available print object.
'
'If Level is 1, the array contains PRINTER_INFO_1 structures.
'If Level is 2, the array contains PRINTER_INFO_2 structures.
'If Level is 4, the array contains PRINTER_INFO_4 structures.
'If Level is 5, the array contains PRINTER_INFO_5 structures.
'
'The buffer must be large enough to receive the array of data
'structures and any strings or other data to which the structure
'members point. If the buffer is too small, the pcBNeeded parameter
'returns the required buffer size.
'
'Windows 95: The buffer cannot receive PRINTER_INFO_4 structures.
'It can receive any of the other types.

'''''''''''''''''''''''
'cbBuf
'Specifies the size, in bytes, of the buffer pointed to by pPRinterEnum.
'''''''''''''''''''''''
'pcBNeeded
'Pointer to a value that receives the number of bytes copied if the
'function succeeds or the number of bytes required if cbBuf is too small.
'''''''''''''''''''''''
'pcReturned
'Pointer to a value that receives the number of PRINTER_INFO_1,
'PRINTER_INFO_2, PRINTER_INFO_4, or PRINTER_INFO_5 structures that
'the function returns in the array to which pPRinterEnum points.


'PRINTER_INFO_4 returned Attribute values
Public Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Public Const PRINTER_ATTRIBUTE_DIRECT = &H2
Public Const PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800&
Public Const PRINTER_ATTRIBUTE_LOCAL = &H40
Public Const PRINTER_ATTRIBUTE_NETWORK = &H10
Public Const PRINTER_ATTRIBUTE_QUEUED = &H1
Public Const PRINTER_ATTRIBUTE_SHARED = &H8
Public Const PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400

'PRINTER_INFO_1 returned Flag values
Public Const PRINTER_ENUM_CONTAINER = &H8000&
Public Const PRINTER_ENUM_EXPAND = &H4000
Public Const PRINTER_ENUM_ICON1 = &H10000
Public Const PRINTER_ENUM_ICON2 = &H20000
Public Const PRINTER_ENUM_ICON3 = &H40000
Public Const PRINTER_ENUM_ICON4 = &H80000
Public Const PRINTER_ENUM_ICON5 = &H100000
Public Const PRINTER_ENUM_ICON6 = &H200000
Public Const PRINTER_ENUM_ICON7 = &H400000
Public Const PRINTER_ENUM_ICON8 = &H800000
 
'Fin declaración para impresoras
 
 
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const SPI_GETWORKAREA = 48

Public objTextBox As New AlgStdFunc.clsTextBoxEdit
 
'***********************************************************************
' Constantes Propias
'***********************************************************************
Public Const si                  As String = "Sí"
Public Const No                  As String = "No"
Public Const NullString          As String = ""
Public Const UNKNOWN_ERRORSOURCE As String = "[Fuente de Error Desconocida]"
Public Const KNOWN_ERRORSOURCE   As String = "[Fuente de Error Conocida]"

'Alertas
Public Const MUESTRA_A_PEDIDO                As Integer = 1
Public Const MUESTRA_SIEMPRE                 As Integer = 2
Public Const NO_PERMITE_FACTURAR             As Integer = 3
Public Const NO_PERMITE_REMITIR              As Integer = 4
Public Const NO_PERMITE_CERTIFICAR           As Integer = 5
Public Const NO_PERMITE_LIQUIDAR             As Integer = 6
Public Const NO_PERMITE_PAGAR                As Integer = 7
Public Const NO_PERMITE_FLETES               As Integer = 8
Public Const NO_PERMITE_EGRESO               As Integer = 9
Public Const IMPRIME_EN_CTA_CTE              As Integer = 10
Public Const NO_PERMITE_PEDIDOS              As Integer = 11
Public Const NO_PERMITE_CBTE_CONTRATO_COMPRA As Integer = 12   'TP 8242. Samsa
Public Const NO_PERMITE_CBTE_CONTRATO_VENTA  As Integer = 13   'TP 8242. Samsa
'***********************************************************************
Private ErrorLog                 As ErrType

'DEFINICIONES PARA EL DATAGRID CONTROL
Public Type udtGrid
   Ancho          As Integer
   Titulo         As String
   Campo          As String
   Alineacion     As MSDataGridLib.AlignmentConstants
   Formato        As String
   Permiso        As String                           'permiso requerido para visualizar la columnas
End Type

Public Type udtRefreshGrid
   Grilla            As MSDataGridLib.DataGrid
   ColsProperties()  As udtGrid
End Type

'DEFINICIONES PARA EL HFLEXGRID CONTROL
Public Type udtRefreshFLEXGrid
   Grilla            As MSHierarchicalFlexGridLib.MSHFlexGrid
   ColsProperties()   As udtGrid
End Type

Public Type udtFLEXGrid
   Ancho          As Integer
   Titulo         As String
   Campo          As String
   Alineacion     As MSHierarchicalFlexGridLib.AlignmentSettings
   Formato        As String
   Permiso        As String                           'permiso requerido para visualizar la columnas
End Type

Public Type udtSystemOptions
   iCacheSize                     As Integer       'valor del parámetro cachesize (registro del sistema)
   iZoom                          As Integer       'valor del Zoom por defecto en Vista Previa
   iFetchMode                     As alFetchMode   'indica el modo en el que vendran capturados los registros del server
   lngFetchLimit                  As Long          'si alFetchMode = 2, es el limite de registros recuperados en manera sincronica
   iFetchModeSearch               As alFetchMode   'indica el modo en el que vendran capturados los registros del server (para la busqueda)
   lngFetchLimitSearch            As Long          'si alFetchMode = 2, es el limite de registros recuperados en manera sincronica (para la busqueda)
   UseLocalCopy                   As String        'Sí=Usa copias locales; No=Usa copias locales (Vista-Lista, Navegación e Impresión)
   UseLocalCopySearch             As String        'Sí=Usa copias locales; No=Usa copias locales (para la búsqueda)
   AskOldLocalCopy                As String        'Sí=Pregunta si usa copias locales desactualizadas;(Vista-Lista, Navegación e Impresión)
   UseMRUEnterprise               As String        'Si=recuerda las ultimas empresas;No=No recuerda
   MaxMRUForms                    As Integer       'Dimension de la colecion MRUForms
End Type

Public Type udtRegistrySubKeys
   Environment                   As String
   DataBaseSettings              As String
   KeyMRUForms                   As String
   MRUEmpresas                   As String
   GridQueries                   As String
   NavigationQueries             As String
   PrintQueries                  As String
   QueryDBQueries                As String
   DataComboQueries              As String
End Type

Public Enum EnumRegistrySubKeys
   Environment = 0
   DataBaseSettings = 1
   KeyMRUForms = 2
   MRUEmpresas = 3
   GridQueries = 4
   NavigationQueries = 5
   PrintQueries = 6
   QueryDBQueries = 7
   DataComboQueries = 8
End Enum

Public Enum EnumSystemOptions
   iCacheSize = 0             'valor del parámetro cachesize (registro del sistema)
   iZoom = 1                  'valor del Zoom por defecto en Vista Previa
   iFetchMode = 2             'indica el modo en el que vendran capturados los registros del server
   lngFetchLimit = 3          'si alFetchMode = 2, es el limite de registros recuperados en manera sincronica
   iFetchModeSearch = 4       'indica el modo en el que vendran capturados los registros del server (para la busqueda)
   lngFetchLimitSearch = 5    'si alFetchMode = 2, es el limite de registros recuperados en manera sincronica (para la busqueda)
   UseLocalCopy = 6           'Sí=Usa copias locales; No=Usa copias locales (Vista-Lista, Navegación e Impresión)
   UseLocalCopySearch = 7     'Sí=Usa copias locales; No=Usa copias locales (para la búsqueda)
   AskOldLocalCopy = 8        'Sí=Pregunta si usa copias locales desactualizadas;(Vista-Lista, Navegación e Impresión)
   UseMRUEnterprise = 9       'Si=recuerda las ultimas empresas;No=No recuerda
   MaxMRUForms = 10           'Dimension de la colecion MRUForms
End Enum

Public Enum EnumIsValid
   Numerico = 1
   Fecha = 2
   Hora = 3
End Enum

Public Enum EnumMenu
   MenuCustom = 0
   MenuEdit = 1
   MenuTools = 2
   MenuWindow = 3
   MenuHelp = 4
   MenuFile = 5
End Enum

Public vValue                    As Variant
Public LoadingMRU                As Boolean
Public aAppReg()                                                     'matriz de aplicaciones registradas de Algortimo
Public aKeys()                                                       'matriz para la lectura de tablas
Public MRUForms                  As Collection                       'coleccion de forms mas frecuentemente usados

Public Enum alFetchMode
   alAsync = 1
   alSync = 2
   alTable = 3
End Enum

Public Enum ContextMenuEnum
   mnxNombre = 0
   mnxOrden = 1
   mnxForms = 2
   mnxCaption = 3
   mnxTarea = 4
   mnxClave = 5
End Enum

'tipos para el frmDialog y frmDynamicChild
Public Enum EnumButtonPressedDialog
   Cancel = 0
   accept = 1
   Advanced = 2
   Preview = 3
   Filter = 4
End Enum
Public Enum EnumEnabledDisabledControls
   Disabled = 0
   Enabled = 1
End Enum



'  constantes Modulos para Tabla COMPROBANTES_REIMPRESION
'  Atención: Estas constantes también están definidos en los
'            módulos especificos donde se graba esta tabla y donde
'            se imprimen los comprobantes
Public Const CRI_LIQUIDACIONES_1116  As Integer = 1


'  constantes para identificar los paneles del Status Bar del ABM Clasico
Public Const STB_PANEL1              As Integer = 1
Public Const STB_PANEL2              As Integer = 2
Public Const STB_PANEL3              As Integer = 3
Public Const STB_PANEL4              As Integer = 4
Public Const STB_PANEL5              As Integer = 5

'  constantes para la barra de estado de los ABM Clasico
Public Const STATE_FETCHING      As String = "Recuperando registros ..."  'constante para mensaje
Public Const STATE_SORTING       As String = "Ordenando"                  'constante para mensaje
Public Const STATE_PRINTING      As String = "Imprimiendo..."             'constante para mensaje
Public Const STATE_NORECORDS     As String = "Ningún registro"            'constante para mensaje
Public Const STATE_INSMODE       As String = "Agregando"                  'constante para mensaje
Public Const STATE_CUSTOMIZE     As String = "Vista Personalizada"        'constante para mensaje
Public Const STATE_NONE          As String = vbNullString                 'constante para mensaje

'  constantes para identificar los mensajes devueltos por Filter
Public Const MSG_CANCEL  As String = "CANCELFILTRO"
Public Const MSG_CONFIRM As String = "CONFIRMAFILTRO"
Public Const MSG_APPLY   As String = "APLICARFILTRO"

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
   
Public Const CB_RESETCONTENT = &H14B
Public Const CB_GETCOUNT = &H146
Public Const CB_GETITEMDATA = &H150
Public Const CB_SETITEMDATA = &H151
Public Const CB_GETLBTEXT = &H148
Public Const CB_ADDSTRING = &H143

'Mensajes enviados por la clase clsControls
Public Const CTL_BEFORE_QUERYDB = &H1
Public Const CTL_QUERYDB_RECORD_SELECTED = &H2
Public Const CTL_CALL_ADMIN = &H3
Public Const CTL_QUERY_USER            As Long = &H4
Public Const CTL_QUERY_CONTROLDATA     As Long = &H5
Public Const CTL_CALL_TOOLS            As Long = &H6
Public Const CTL_CONTEXT_MENU_CLICK    As Long = &H7
Public Const CTL_CONTEXT_MENU_AFTER_SET As Long = &H8
Public Const CTL_BEFORE_F3             As Long = &H9
Public Const CTL_AFTER_F3              As Long = &HA
Public Const CTL_BEFORE_F1             As Long = &HB
Public Const CTL_AFTER_F1              As Long = &HC
Public Const CTL_BEFORE_F2             As Long = &HD
Public Const CTL_AFTER_F2              As Long = &HE
Public Const CTL_BEFORE_F4             As Long = &HF
Public Const CTL_AFTER_F4              As Long = &H1A
Public Const CTL_BEFORE_F5             As Long = &H1B
Public Const CTL_AFTER_F5              As Long = &H1C
Public Const CTL_BEFORE_F6             As Long = &H1D
Public Const CTL_AFTER_F6              As Long = &H1E
Public Const CTL_BEFORE_F7             As Long = &H1F
Public Const CTL_AFTER_F7              As Long = &H2A
Public Const CTL_BEFORE_F8             As Long = &H2B
Public Const CTL_AFTER_F8              As Long = &H2C
Public Const CTL_BEFORE_F9             As Long = &H2D
Public Const CTL_AFTER_F9              As Long = &H2E
Public Const CTL_BEFORE_F10            As Long = &H2F
Public Const CTL_AFTER_F10             As Long = &H3A
Public Const CTL_BEFORE_F11            As Long = &H3B
Public Const CTL_AFTER_F11             As Long = &H3C
Public Const CTL_BEFORE_F12            As Long = &H3D
Public Const CTL_AFTER_F12             As Long = &H3E
Public Const CTL_AFTER_QUERYDB         As Long = &H3F
Public Const CTL_AFTER_SET_VALUE_DEF   As Long = &H4A

'Mensajes enviados por la Filter
Public Const FILTER_CALL_ADMIN = &H1
Public Const FILTER_QUERY_USER   As Long = &H2
Public Const FILTER_QUERY_CONTROLDATA   As Long = &H3

'constantes para el autosize column del ListView
Private Const LVM_SETCOLUMNWIDTH = &H1000 + 30
Private Const LVSCW_AUTOSIZE = 65535
Private Const LVSCW_AUTOSIZE_USEHEADER = 65534
Private Const LVM_FIRST = &H1000

'REPORT_BUTTON_CLICKED: Mensaje Enviado por frmReport - hay dos constantes. Esta se usa en los forms de listados y procesos,
'para identificar el evento. Y hay otra definida en modABM (ABM_REPORT_BUTTON_CLICKED)
'que se usa desde los forms de ABMs para identificar elmensaje en objABM_Message
Public Const REPORT_BUTTON_CLICKED = &H1 'despues de hacer click en unm boton opcional del report

'/ variables publicas seteadas por el proyecto Inicio
Public mvarMDIForm               As MDIForm
Public CUsuario                  As BOSeguridad.clsUsuario
Public CSysEnvironment           As AlgStdFunc.clsSysEnvironment
Public SystemOptions             As udtSystemOptions
Public RegistrySubKeys           As udtRegistrySubKeys
Public rstContextMenu            As ADODB.Recordset
Public rstMenu                   As ADODB.Recordset
Public rstVistasPersonalizadas   As ADODB.Recordset
Public rstVistasExportacion      As ADODB.Recordset

Public aMenuTools(9) As String                                                 'matriz elementos del menu Herramientas
Public aMenuHelp(4)  As String                                                'matriz elementos del menu Ayuda



Private miControlData               As DataShare.udtControlData 'TP 17409 INC 163891

'<TP 19870>
Public Const CARACTER_SEPARADOR_GRUPO      As String = "<F8>"

Public Type udtInfoCodigoTrazabilidad
   GTIN              As String
   Lote              As String
   NumeroSerie       As String
   FechaElaboracion  As Date
   FechaVencimiento  As Date
   CodigoHumanamenteLegible As String 'seria el codigo formateado con los parentecis
End Type
'</TP 19870>

'********************************************************************************************
' IMPORTANTE
'           Para poder utilizar este modulo de codigo en un proyecto es necesario que
' existan las siguientes referencias a los componentes de Algoritmo:
'  - DataAccess
'  - DataShare
'
' Instanciar las variables:
'     DSConnectionString = "STRING DE CONEXION"
'     Set ParametrosInternos = New clsParametrosInternos
'
' es necesario ademas que hayan sido ejecutados los siguientes metodos:
'  - InitDataAccess DSConnectionString
'  - ParametrosInternos.OpenParametros
'
' Si el proyecto no posee referencias a :
'  - Microsoft ActiveX Data Objects Library
'  - Crystal Reports SmartViewer
'  - Crystal Reports 7 ActiveX Designer RunTime
' entonces estas referencias deberan ser creadas.
'
'********************************************************************************************

Public Function GetData(ByVal strEmpresa As String, ByVal strSQL As String, Optional ByVal iCursorType As CursorTypeEnum, Optional ByVal iLockType As LockTypeEnum) As ADODB.Recordset
Dim rst As ADODB.Recordset

   
   If iCursorType = 0 Then
      iCursorType = adOpenForwardOnly
   End If
   
   If iLockType = 0 Then
      iLockType = adLockReadOnly
   End If
   
   Set rst = New ADODB.Recordset
  
   Select Case SystemOptions.iFetchMode
      Case alAsync
         rst.CursorLocation = adUseClient
         rst.CacheSize = SystemOptions.iCacheSize
         rst.Open TranslateSQL(strEmpresa, strSQL), GetSPMProperty(DBSConnectionString), iCursorType, iLockType, adAsyncFetch
      Case alSync
         Set rst = Fetch(strEmpresa, strSQL, iCursorType, iLockType)
      Case alTable
         If SystemOptions.lngFetchLimit > CountRecord(strEmpresa, strSQL) Then
            Set rst = Fetch(strEmpresa, strSQL, iCursorType, iLockType)
         Else
            rst.CursorLocation = adUseClient
            rst.CacheSize = SystemOptions.iCacheSize
            rst.Open TranslateSQL(strEmpresa, strSQL), GetSPMProperty(DBSConnectionString), iCursorType, iLockType, adAsyncFetch
         End If
      
   End Select
   
   Set GetData = rst
   
End Function

Public Sub ShowErrMsg(ByRef ErrorLog As ErrType)
Dim iErrNumber         As Long                          ' numero de error (sin vbObjectError)
Dim bAlgError          As Boolean                       ' identifica un error de Algoritmo
'Dim ix                 As Integer
Dim strSource          As String
Dim n                  As Integer
Dim frmMsg             As frmMsgBox
Dim StrMensaje         As String
Dim strDetalle         As String

   '  muestra en manera amigable un mensaje de error
   

   strSource = Trim(ErrorLog.source)
   
   bAlgError = True
   n = InStr(strSource, UNKNOWN_ERRORSOURCE)
   If n > 0 Then
      ' es un error generado por alguna aplicacion de Algoritmo
      bAlgError = False
   End If
   
   strSource = Replace(strSource, UNKNOWN_ERRORSOURCE, NullString)
   
   If bAlgError Then
      ' errores de Algortimo
      iErrNumber = ErrorLog.NumError - vbObjectError
      Select Case iErrNumber
         Case Is < 10000
            ' warnings de Algortimo
'            MsgBox ErrorLog.Descripcion & vbCrLf & vbCrLf & _
'                   "Origen: " & vbCrLf & vbCrLf & _
'                   strSource, vbOKOnly, App.ProductName
            
            
            Set frmMsg = New frmMsgBox
            
            StrMensaje = ErrorLog.Descripcion
            strDetalle = strSource
            
            frmMsg.ShowMsg App.ProductName, StrMensaje, strDetalle, Warning
            
         Case 10000 To 20000
            'Errores Severos de Algoritmo
'            MsgBox ErrorLog.Descripcion, vbExclamation, "Error Manager"
            
            Set frmMsg = New frmMsgBox
            
            StrMensaje = ErrorLog.Descripcion
            frmMsg.ShowMsg App.ProductName, StrMensaje, strDetalle, Error
            
            
      End Select
   Else
      ' errores no generados por Algoritmo
'      MsgBox "Se produjo el siguiente error:" & vbCrLf & vbCrLf & _
'             "Número     : " & ErrorLog.NumError & vbCrLf & vbCrLf & _
'             "Descripción: " & vbCrLf & ErrorLog.Descripcion & vbCrLf & vbCrLf & _
'             "Llamadas   : " & vbCrLf & _
'             strSource, vbExclamation, "Error Manager"
             
         Set frmMsg = New frmMsgBox
          
         ErrorLog.Descripcion = Replace(ErrorLog.Descripcion, vbCr, NullString)
         ErrorLog.Descripcion = Replace(ErrorLog.Descripcion, vbLf, NullString)
          
         StrMensaje = ErrorLog.Descripcion
         strDetalle = "Número     : " & ErrorLog.NumError & vbCrLf & strSource
         
         frmMsg.ShowMsg App.ProductName, StrMensaje, strDetalle, Error
            
             
   End If
   
   ' una vez visualizado el mensaje de error, este viene limpiado
   With ErrorLog
      .Modulo = NullString
      .NumError = 0
      .source = NullString
      .Descripcion = NullString
   End With
   
   Screen.MousePointer = vbDefault
   
End Sub

Public Sub LoadError(ByRef ErrLog As ErrType, ByVal strSource As String)
Dim PropBag As PropertyBag

   ' carga la información del error en la variable ErrorLog
   SetError ErrLog, App.ProductName, strSource
   
   Set PropBag = New PropertyBag
   With PropBag
      .WriteProperty "ERR_EMPRESA", ErrLog.Empresa
      .WriteProperty "ERR_APLICACION", ErrLog.Aplicacion
      .WriteProperty "ERR_COMENTARIO", ErrLog.COMENTARIO
      .WriteProperty "ERR_DESCRIPCION", ErrLog.Descripcion
      .WriteProperty "ERR_ERRORNATIVO", ErrLog.ErrorNativo
      .WriteProperty "ERR_FORM", ErrLog.Form
      .WriteProperty "ERR_MAQUINA", ErrLog.Maquina
      .WriteProperty "ERR_MODULO", ErrLog.Modulo
      .WriteProperty "ERR_NUMERROR", ErrLog.NumError
      .WriteProperty "ERR_SOURCE", ErrLog.source
      .WriteProperty "ERR_USUARIO", ErrLog.Usuario
      .WriteProperty "WRITE_ERROR", ErrLog.WriteError
   End With
   
   TrapError PropBag.Contents
   Set PropBag = Nothing
   
End Sub
Public Sub SetError(ByRef ErrLog As ErrType, ByVal strModuleName As String, ByVal strSource As String)

   With ErrLog
'      .Usuario = CUsuario.Usuario
      .Modulo = strSource
      .NumError = Err.Number
      .source = Err.source
      .Aplicacion = UCase(App.ProductName)
      .WriteError = si
      
      If InStr(.source, KNOWN_ERRORSOURCE) = 0 Then
         If InStr(.source, UNKNOWN_ERRORSOURCE) = 0 Then
            .source = UNKNOWN_ERRORSOURCE & vbCrLf & .source
         End If
      Else
         .source = Replace(.source, KNOWN_ERRORSOURCE, NullString)
         .WriteError = False
      End If
      
      If ErrLog.Form <> NullString Then strSource = ErrLog.Form & "." & strSource
         
      
      If InStr(.source, strModuleName) > 0 Then
         .source = .source & vbCrLf & "[" & strSource & "]"
      Else
         .source = .source & vbCrLf & strModuleName & "[" & strSource & "]"
      End If
      
      If .Descripcion = NullString Then
         .Descripcion = Err.Description
      End If

   End With

End Sub

Public Sub GetColumnsWidth(ByVal oControl As Object)
Dim ix As Integer

   '  para uso y consumo de MARCELO

   If TypeOf oControl Is DataGrid Then
      For ix = 0 To oControl.Columns.Count - 1
         Debug.Print oControl.Columns(ix).Width
      Next ix
   End If
   If TypeOf oControl Is ListView Then
      For ix = 1 To oControl.ColumnHeaders.Count
         Debug.Print oControl.ColumnHeaders(ix).Width
      Next ix
   End If
   
End Sub

Public Function MaxLen(aProps(), ByVal strField As String) As Integer

   MaxLen = Len(Formato("99999999999999,9999999999", FieldProperty(aProps, strField, dsDecimales), FieldProperty(aProps, strField, dsDimension) - FieldProperty(aProps, strField, dsDecimales))) + 1
   
End Function

Public Sub CopySubTree(SourceTV As TreeView, sourceND As Node, DestTV As TreeView, destND As Node)
Dim ix As Long, so As Node, de As Node
'Dim s As String

    ' rutina recursiva que copia o mueve todos los hijos de un nodo a otro nodo
    
    If sourceND.Children = 0 Then Exit Sub
    
    Set so = sourceND.Child
    For ix = 1 To sourceND.Children
'        s = so.key
'        so.key = ""
        ' agrega un nodo en el TreeView de destino
        Set de = DestTV.Nodes.Add(destND, tvwChild, so.Key, so.Text, so.Image, so.SelectedImage)
        de.ExpandedImage = so.ExpandedImage
        
        ' agrega todos los hijos de este nodo, en modo recursivo
        CopySubTree SourceTV, so, DestTV, de
        
        ' obtiene una referencia al proximo
        Set so = so.Next
    Next
    
End Sub

Public Sub DeleteSubTree(TV As TreeView, Node As Node)
Dim ix As Long, nd As Node, nd2 As Node
    
    ' recursivamente borra un subarbol de nodos
    
    ' si el nodo tiene uno o mas hijos, borra el primero
    Set nd = Node.Child
    For ix = 1 To Node.Children
        ' obtiene una referencia al proximo nodo, antes de borrarlo
        Set nd2 = nd.Next
        DeleteSubTree TV, nd
        Set nd = nd2
    Next
    ' borro el nodo
    TV.Nodes.Remove Node.Index
        
End Sub

Public Function IsEmptyRecordSet(ByVal rs As ADODB.Recordset) As Boolean

   ' determina si un recordset esta vacio
   
   If ((rs.BOF = True) And (rs.EOF = True)) Then
      IsEmptyRecordSet = True
   End If
   If ((rs.BOF = True Or rs.EOF = True) And (rs.RecordCount = 0)) Then
      IsEmptyRecordSet = True
   End If
   
End Function

Public Sub RefreshDataGrid(GridInfo As udtRefreshGrid)
Dim ix As Integer

   On Error GoTo GestErr
   
   For ix = 0 To UBound(GridInfo.ColsProperties)
      
      GridInfo.Grilla.Columns(ix).Width = GridInfo.ColsProperties(ix).Ancho
      GridInfo.Grilla.Columns(ix).Caption = GridInfo.ColsProperties(ix).Titulo
      GridInfo.Grilla.Columns(ix).DataField = GridInfo.ColsProperties(ix).Campo
      GridInfo.Grilla.Columns(ix).Alignment = GridInfo.ColsProperties(ix).Alineacion
      If Len(GridInfo.ColsProperties(ix).Formato) > 0 Then
         GridInfo.Grilla.Columns(ix).DataFormat.Format = GridInfo.ColsProperties(ix).Formato
      End If
      
      If Len(GridInfo.ColsProperties(ix).Permiso) > 0 Then
         If Not TaskIsEnabled(GridInfo.ColsProperties(ix).Permiso, CUsuario) Then
            GridInfo.Grilla.Columns(ix).Width = 0
         End If
       End If
      
   Next ix

   GridInfo.Grilla.Refresh

   Exit Sub
   
GestErr:
   LoadError ErrorLog, "RefreshDataGrid"
   ShowErrMsg ErrorLog
End Sub

Public Sub RefreshFLEXGrid(GridInfo As udtRefreshFLEXGrid)
Dim ix As Long
Dim iy As Long
   
   On Error GoTo GestErr
   
   For ix = LBound(GridInfo.ColsProperties) To UBound(GridInfo.ColsProperties)

      GridInfo.Grilla.ColWidth(ix + 1) = GridInfo.ColsProperties(ix).Ancho
      GridInfo.Grilla.TextMatrix(0, ix + 1) = GridInfo.ColsProperties(ix).Titulo
      GridInfo.Grilla.ColAlignment(ix + 1) = GridInfo.ColsProperties(ix).Alineacion
      
      If Len(GridInfo.ColsProperties(ix).Permiso) > 0 Then
         If Not TaskIsEnabled(GridInfo.ColsProperties(ix).Permiso, CUsuario) Then
            GridInfo.Grilla.ColWidth(ix) = 0
         End If
       End If
      
   Next ix

   For ix = 1 To GridInfo.Grilla.Rows - 1             'formateo de las columnas
   
      For iy = LBound(GridInfo.ColsProperties) To UBound(GridInfo.ColsProperties)
   
         
         If GridInfo.Grilla.ColWidth(iy) <> 0 Then
         
            If GridInfo.ColsProperties(iy).Formato <> NullString Then
            
               If IsDate(GridInfo.Grilla.TextMatrix(ix, iy + 1)) Then
                  GridInfo.Grilla.TextMatrix(ix, iy + 1) = Format(GridInfo.Grilla.TextMatrix(ix, iy + 1), GridInfo.ColsProperties(iy).Formato, , vbUseSystem)
               ElseIf IsNumeric(GridInfo.Grilla.TextMatrix(ix, iy + 1)) Then
                  GridInfo.Grilla.TextMatrix(ix, iy + 1) = Format(Val(GridInfo.Grilla.TextMatrix(ix, iy + 1)), GridInfo.ColsProperties(iy).Formato)
               Else
                  GridInfo.Grilla.TextMatrix(ix, iy + 1) = Format(GridInfo.Grilla.TextMatrix(ix, iy + 1), GridInfo.ColsProperties(iy).Formato)
               End If
               
            End If
                  
         End If
         
      Next iy
      
   Next ix

   GridInfo.Grilla.Refresh
   
   Exit Sub
GestErr:
   Resume Next
End Sub

Public Sub LoadGridArray(ByRef aArray() As udtGrid, ByVal iColumnsWidth As Integer, ByVal iAlignment As MSDataGridLib.AlignmentConstants, ByVal strField As String, ByVal strColumnsTitle As String, Optional strFormat As String, Optional strTarea As String)
Dim ix As Integer

   On Error Resume Next

   ix = -1
   ix = UBound(aArray)
   
   ReDim Preserve aArray(ix + 1)
   ix = UBound(aArray)
   
   aArray(ix).Alineacion = iAlignment
   aArray(ix).Ancho = iColumnsWidth
   aArray(ix).Campo = strField
   aArray(ix).Formato = strFormat
   aArray(ix).Permiso = strTarea
   aArray(ix).Titulo = strColumnsTitle

End Sub

Public Sub LoadFLEXGridArray(ByRef aArray() As udtFLEXGrid, ByVal iColumnsWidth As Integer, ByVal iAlignment As MSHierarchicalFlexGridLib.AlignmentSettings, ByVal strField As String, ByVal strColumnsTitle As String, Optional strFormat As String, Optional strTarea As String)
Dim ix As Integer

   On Error Resume Next

   ix = -1
   ix = UBound(aArray)
   
   ReDim Preserve aArray(ix + 1)
   ix = UBound(aArray)
   
   aArray(ix).Alineacion = iAlignment
   aArray(ix).Ancho = iColumnsWidth
   aArray(ix).Campo = strField
   aArray(ix).Formato = strFormat
   aArray(ix).Permiso = strTarea
   aArray(ix).Titulo = strColumnsTitle

End Sub

Public Function MakeSQLWhere(ByVal strTableName As String, ByVal strFieldList As String, ByRef aValueList() As String) As String
Dim ix                  As Integer
Dim ix1                 As Integer
Dim aArray()            As String
Dim aArray1()           As String
Dim aTableProperties()  As Variant                    'arreglo de propiedades de cada campo de la tabla
Dim strWhere            As String
Dim iTypeOfField        As Integer
Dim dups                As Integer
Dim strTemp             As String

'**************************************************************************************
'  strTableName : Nombre de la tabla
'  StrFieldList : Lista de campos separada por punto y coma
'  aValueList   : arreglo de lista de valores separada por punto y coma
'
'  El primer elemento de la lista strFieldList se corresponde con la lista de valores
'  del primer elemento del arreglo aValueList
'  Ejemplo:
'  strFieldList = "ESPECIE;COSECHA"
'  aValueList(0) = "TRIGO;MAIZ"
'  aValueList(1) = "98/99;99/00"
'
'    ? MakeSQLWhere(strFieldList,aValueList)
'
'  Devuelve:
'  (ESPECIE = 'TRIGO' AND COSECHA = '98/99') OR (ESPECIE = 'MAIZ' AND COSECHA = '99/00')
'
'**************************************************************************************

   'purgo la lista de valores dejando solo los valores distintos y quito los iguales
   For ix = 0 To UBound(aValueList, 1)
      
      aArray = Split(aValueList(ix), ";")
      
      dups = FilterDuplicates(aArray())
      If dups Then
        ReDim Preserve aArray(LBound(aArray) To UBound(aArray) - dups) As String
      End If
         
      'reconstruyo la lista y la vuelvo a asignar al arreglo
      For ix1 = 0 To UBound(aArray)
         If Len(strTemp) = 0 Then
            strTemp = strTemp & aArray(ix1)
         Else
            strTemp = strTemp & ";" & aArray(ix1)
         End If
         If Right(strTemp, 1) = ";" Then strTemp = Left(strTemp, Len(strTemp) - 1)
         
      Next ix1
      
      aValueList(ix) = strTemp
      
   Next ix

   ix = InStrCount(strFieldList, ";")
   ix1 = InStrCount(aValueList(0), ";")

   ReDim aArray(ix, ix1 + 1)

   'metto en la primer columna (la 0) los nombres de los campos
   aArray1 = Split(strFieldList, ";")
   For ix = 0 To UBound(aArray1)
      aArray(ix, 0) = aArray1(ix)
   Next ix

   'meto los datos en una matriz bidimensional
   
   ' aarray(0,0) = "NOMBRECAMPO1" aarray(0,1) = "VALORCAMPO1" aarray(0,2) = "VALORCAMPO2"
   ' aarray(1,0) = "NOMBRECAMPO2" aarray(1,1) = "VALORCAMPO1" aarray(0,2) = "VALORCAMPO2"
   
   For ix = 0 To UBound(aArray, 1)
   
      aArray1 = Split(aValueList(ix), ";")
   
      For ix1 = 0 To UBound(aArray1)
         
         aArray(ix, ix1 + 1) = aArray1(ix1)
         
      Next ix1
      
   Next ix

   'recupera información del diccionario y la guarda en el arreglo
   aTableProperties = GetTableInformation(strTableName)

   ' construyo la clausula WHERE
   strWhere = ""
   For ix = 1 To UBound(aArray, 2)
   
      For ix1 = 0 To UBound(aArray, 1)
         
         iTypeOfField = FieldProperty(aTableProperties, aArray(ix1, 0), dsTipoDato)
         
         If Len(aArray(ix1, ix)) = 0 Then
            strWhere = strWhere & "(" & strTableName & "." & aArray(ix1, 0) & " IS NULL) AND "
         Else
            Select Case iTypeOfField
              Case adChar, adVarChar
                  strWhere = strWhere & "(" & strTableName & "." & aArray(ix1, 0) & " = '" & aArray(ix1, ix) & "') AND "
               Case adNumeric
                  strWhere = strWhere & "(" & strTableName & "." & aArray(ix1, 0) & " = " & aArray(ix1, ix) & ") AND "
               Case adDBTimeStamp
                  strWhere = strWhere & "(" & strTableName & "." & aArray(ix1, 0) & " = TO_DATE(Format(" & """" & aArray(ix1, ix) & """" & ", " & """" & "yyyy-mm-dd" & """" & "), " & """" & "YYYY-MM-DD" & """" & ")" & ") AND "
            End Select
         End If
'                                                                                     TO_DATE('" & Format(aArray(ix1, ix), "yyyy-mm-dd") & "', 'YYYY-MM-DD')
      Next ix1
      
      If Right(strWhere, 5) = " AND " Then
         strWhere = Left(strWhere, Len(strWhere) - 5)
      End If
      
      strWhere = strWhere & " OR "
   Next ix
   
   
   If Right(strWhere, 4) = " OR " Then
      strWhere = Left(strWhere, Len(strWhere) - 4)
   End If
      
   MakeSQLWhere = strWhere
   
End Function

Public Function IsMRUForm(ByVal lngHndW As Long) As Boolean
Dim frm As Form

   '  determina si un forms esta cargado en la colección MRUForms
   
   If MRUForms Is Nothing Then Exit Function
   
   For Each frm In MRUForms

      If frm.hWnd = lngHndW Then IsMRUForm = True: Exit For

   Next frm
   
End Function

Public Sub ClearForm(ByVal frm As Form)
Dim cntrl As Control

   For Each cntrl In frm.Controls
      Select Case TypeName(cntrl)
         Case "RichTextBox", "TextBox"
            cntrl.Text = NullString
         Case "PowerMask"
            cntrl.Text = NullString
         Case "Label"
            If Left(cntrl.Name, 3) = "lbl" Then
               cntrl.Caption = NullString
            End If
         Case "CheckBox"
            cntrl.Value = vbUnchecked
         Case "DataPicker"
            cntrl.Value = Date
         Case "OptionButton"
            cntrl.Value = False
         Case "ListView"
            cntrl.ListItems.Clear
         Case "ComboBox"
            cntrl.ListIndex = -1
         Case "StatusBar"
            Dim p As Panel
            For Each p In cntrl.Panels
               p.Text = NullString
            Next p
         Case "Tab"
            cntrl.Tab = 0
         Case "DataCombo"
            cntrl.BoundText = NullString
      End Select
   Next cntrl

End Sub

Public Function IsValid(Optional ByVal DataType As EnumIsValid) As Boolean
Dim oActiveControl      As Control
Dim aTableProperties()  As Variant       'arreglo de propiedades de cada campo de la tabla
Dim iDimension          As Integer
Dim iEnteros            As Integer
Dim iDecimales          As Integer
Dim strFormato          As String

   '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   '  determina si el dato es válido para el control o no lo es.
   '  Si se desea bloquear el focus del control agregar un SetFocus
   '  Agregar un 'case' por cada control que se desee validar.
   '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

   IsValid = True

   Set oActiveControl = Screen.ActiveControl

   If DataType = 0 Then
      DataType = Numerico
   End If

   On Error Resume Next
   If Len(Trim(oActiveControl.Text)) = 0 Then Exit Function
   If Err Then Exit Function
   
   Select Case DataType
      Case Numerico 'el dato debe ser numèrico
         If (Not IsNumeric(Trim(oActiveControl.Text))) Then
            MsgBox "Valor no válido para este campo", vbExclamation, App.ProductName
            IsValid = False
            Exit Function
         End If
         
         'controlo si el numero de enteros y decimales ingresados conrresponde al definido en el diccionario
         If Len(oActiveControl.DataField) = 0 Then Exit Function
         aTableProperties = GetFieldInformation(oActiveControl.DataField)
         
         iDimension = FieldProperty(aTableProperties, oActiveControl.DataField, dsDimension)
         iEnteros = iDimension - FieldProperty(aTableProperties, oActiveControl.DataField, dsDecimales)
         iDecimales = FieldProperty(aTableProperties, oActiveControl.DataField, dsDecimales)
         
         If iEnteros < Len(CStr(xRound(Trim(oActiveControl.Text), 0))) Then
            strFormato = "+/-" & Formato("99999999999999,9999999999", iDecimales, iDimension - iDecimales)
            MsgBox "El número de enteros ingresados para este campo es demasiado grande. " & _
                   "El máximo valor permitido es " & strFormato, vbExclamation, App.ProductName
            IsValid = False
            Screen.ActiveControl.SetFocus
            Exit Function
         End If
         
      Case Fecha 'el dato debe ser una fecha válida
'         If (Len(Trim(oActiveControl.Text)) > 0) And (Not IsDate(Trim(oActiveControl.Text))) Then
'            MsgBox "No es una fecha válida", vbExclamation, App.ProductName
'            IsValid = False
'            Screen.ActiveControl.SetFocus
'            Exit Function
'         End If
         
         If (Len(Trim(oActiveControl.Text)) > 0) Then
            On Error Resume Next
            Dim dFecha As Date
            dFecha = DateValue(oActiveControl.Text)
            If Err Or UCase(CStr(dFecha)) = UCase("12:00:00 a.m.") Then
               MsgBox "No es una fecha válida", vbExclamation, App.ProductName
               IsValid = False
               Screen.ActiveControl.SetFocus
               Exit Function
            End If
         End If
         
      Case Hora 'el dato debe ser una hora válida
         If (Len(Trim(oActiveControl.Text)) > 0) Then
            On Error Resume Next
            Dim dHora As Date
            dHora = TimeValue(oActiveControl.Text)
            If Err Then
               MsgBox "No es una hora válida", vbExclamation, App.ProductName
               IsValid = False
               Screen.ActiveControl.SetFocus
               Exit Function
            End If
         End If
   End Select
   
End Function

Public Function FindNode(ByVal nods As Nodes, ByVal strNodeText, Optional ByVal StartNode As Integer) As Node
Dim nod  As Node
Dim ix   As Integer

   ' busca un nodo en un treeview. StartNode es el nodo de inicio, si es mayor que 0 inicia
   ' a partir del nodo cuyo indice es StartNode
   
   If StartNode > 0 Then
      Set nod = nods.Item(StartNode)
   
      For ix = nod.Index To nods.Count - 1
         If nod.Text = strNodeText Then FindNode = nod: Exit For
      Next ix
   Else
   
      For Each nod In nods
         If UCase(nod.Text) = UCase(strNodeText) Then Set FindNode = nod: Exit For
      Next nod
   End If
   

End Function

Public Function OpenComboDataSource(ByVal strEmpresa As String, ByVal strSQL As String) As ADODB.Recordset
Dim strTabla      As String
Dim strPrefijo    As String
Dim strRegistro   As String
Dim strNombreFile As Variant
Dim rst           As ADODB.Recordset
Dim dateTemp      As Variant
Dim FilterTemp    As Variant
Dim aArray()      As String
Dim SQLScomposer  As New AlgStdFunc.clsSQLScomposer
Dim objDataAccess As DataAccess.clsDataAccess

   strPrefijo = "dtc"
   strRegistro = RegistrySubKeys.DataComboQueries
            
   
   SQLScomposer.SQLInputString = strSQL
   
   Set rst = New ADODB.Recordset
   
   ' Pregunta si Usa Archivos Temporarios
   If SystemOptions.UseLocalCopy = No Then

      ' No Usa Archivos Temporarios
      Set rst = GetData(strEmpresa, strSQL, adOpenStatic, adLockPessimistic)
      
   Else
   
      ' Usa Archivos Temporarios
      
      aArray = Split(SQLScomposer.SQLFrom, ",")
      strTabla = aArray(0)
      
      ' Verifico la existencia en el registro de la clave del archivo
      strNombreFile = GetRegistryValue(HKEY_LOCAL_MACHINE, strRegistro & "\" & strTabla, "File", REG_SZ, "", True)
      If Len(strNombreFile) = 0 Then
         strNombreFile = strPrefijo & strTabla
         SetRegistryValue HKEY_LOCAL_MACHINE, strRegistro & "\" & strTabla, "File", REG_SZ, strNombreFile
      End If
      
      'Verifico la existencia del archivo temporario
      If Len(Dir(StoredQueries & strNombreFile)) = 0 Then
         'no existe archivo temporario
         Set rst = GetData(strEmpresa, strSQL, adOpenStatic, adLockPessimistic)
         
         'Si es sincrónico o es asincrónico pero ya terminó de recuperar los datos graba el temp
         If SystemOptions.iFetchMode = alSync Or rst.State <> adStateOpen + adStateFetching Then
            SaveTempFile rst, strRegistro, strTabla
         End If
         
      Else
         'existe el archivo temporario, obtengo la fecha de creación y filtro aplicado
         dateTemp = GetRegistryValue(HKEY_LOCAL_MACHINE, strRegistro & "\" & strTabla, "DateCreated", REG_SZ, "", True)
         FilterTemp = GetRegistryValue(HKEY_LOCAL_MACHINE, strRegistro & "\" & strTabla, "Filter", REG_SZ, "", True)
         
         Set objDataAccess = CreateObject("DataAccess.clsDataAccess")
         Select Case True
            Case dateTemp = vbNullString
               'fue eliminada manualmente la clave del registro
               Set rst = GetData(strEmpresa, strSQL, adOpenStatic, adLockPessimistic)
               
               If SystemOptions.iFetchMode = alSync Or rst.State <> adStateOpen + adStateFetching Then
                  SaveTempFile rst, strRegistro, strTabla
               End If
            
            Case CDate(dateTemp) < objDataAccess.GetLastUpdate(strEmpresa, SQLScomposer.SQLFrom)
               'existe el archivo temporario pero esta desactualizado
               
               ' siempre recupera información actualizada de la BD
               Set rst = GetData(strEmpresa, strSQL, adOpenStatic, adLockPessimistic)
               
               If SystemOptions.iFetchMode = alSync Or rst.State <> adStateOpen + adStateFetching Then
                  SaveTempFile rst, strRegistro, strTabla
               End If
               
            
            Case Else
               'existe, esta actualizado y ademas los filtros son iguales
               rst.Open StoredQueries & strNombreFile, , , , adCmdFile
         End Select
      
         Set objDataAccess = Nothing
      End If
      
   End If
   
   rst.Sort = Mid(SQLScomposer.SQLOrderBy, InStr(SQLScomposer.SQLOrderBy, ".") + 1)
   
   Set OpenComboDataSource = rst
   
   
End Function

Private Sub SaveTempFile(ByVal rst As ADODB.Recordset, ByVal strRegistro As String, ByVal strTabla As String)
Dim strNombreFile As String

   '-- salva el contenido del recordset localmente y efectua la oportuna registracion en el registro de windows

   strNombreFile = "dtc" & strTabla
   
   If Dir(StoredQueries, vbDirectory) = NullString Then MkDir StoredQueries
   
   If Len(Dir(StoredQueries & strNombreFile)) > 0 Then Kill StoredQueries & strNombreFile
   
   rst.Save StoredQueries & strNombreFile
   
   SetRegistryValue HKEY_LOCAL_MACHINE, strRegistro & "\" & strTabla, "DateCreated", REG_SZ, DateTimeServer
   
End Sub

Public Sub SetContextMenu(Optional strMenuName As String, _
                          Optional ByVal strMenuKeyAdmin As String, _
                          Optional ByVal mvarForm As Object)
                          
                          
'Dim ActiveForm As Form
'Dim ActiveControl As Control
'Dim ItemNumber As Integer
'Dim strCaption As String
'Dim bPrintSeparator As Boolean
'
'   'Carga el menu contextual strMenuName. Para los controles de tipo TEXT o DATACOMBO además
'   'agrega el menu standard para estos tipos de controles
'
'   If mvarForm Is Nothing Then
'      Set ActiveForm = Screen.ActiveForm
'   Else
'      Set ActiveForm = mvarForm
'   End If
'   Set ActiveControl = Screen.ActiveControl
'
'   If TypeOf ActiveControl Is TextBox Or _
'      TypeOf ActiveControl Is PowerMask Or _
'      TypeOf ActiveControl Is DataCombo Or _
'      TypeOf ActiveControl Is DataGrid Or _
'      TypeOf ActiveControl Is CRViewer Then
'
'
'      If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is PowerMask Or TypeOf ActiveControl Is DataCombo Then
'         If strMenuName <> "GRILLA" Then
'            ' si es un TextBox o DataCombo agrego el menu standard para estos controles
'            rstContextMenu.Filter = "MNX_NOMBRE = 'TEXT'"
'         Else
'            ' es un text oculto
'            rstContextMenu.Filter = "MNX_NOMBRE = 'GRILLA'"
'         End If
'      End If
'      If TypeOf ActiveControl Is DataGrid Then
'         ' si es una DataGrid agrego el menu standard para ese control
'         rstContextMenu.Filter = "MNX_NOMBRE = 'GRILLA'"
'      End If
'
'      If TypeOf ActiveControl Is CRViewer Then
'         ' si es una DataGrid agrego el menu standard para ese control
'         rstContextMenu.Filter = "MNX_NOMBRE = 'ZOOM_REPORT'"
'      End If
'
'      'empiezo la lectura del menu desplegable
'      Do While Not rstContextMenu.EOF
'
'         strCaption = rstContextMenu(ContextMenuEnum.mnxCaption)
'
'         '-- controlo si despues de este item es necesario un separador
'         If InStr(strCaption, "%S") > 0 Then
'            strCaption = Replace(strCaption, "%S", NullString)
'            bPrintSeparator = True
'         Else
'            bPrintSeparator = False
'         End If
'
'         If ItemNumber > 0 Then
'            'ya fue cargado el primer item
'            Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
'         End If
'         ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True
'         ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Caption = strCaption
'
'         'salvo la clave en el tag asi el programa llamador puede interrogar dicho tag
'         ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Tag = rstContextMenu(ContextMenuEnum.mnxClave)
'
'         ' habilito/deshabilito segun corresponda
'         If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is PowerMask Then
'
'            If rstContextMenu(ContextMenuEnum.mnxClave) = "FINDQUERYDB" Then
'               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True
'            End If
'
'            If rstContextMenu(ContextMenuEnum.mnxClave) = "ADMINISTRAR" Then
'
'               rstMenu.Filter = "MNU_CLAVE = '" & strMenuKeyAdmin & "'"
'               If Not rstMenu.EOF Then
''                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = ((TaskIsEnabled(rstMenu("MNU_TAREA").Value, CUsuario)))
'                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = ((TaskIsEnabled(rstMenu("MNU_CLAVE").Value, CUsuario)))
'               Else
'                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
'               End If
'
'            End If
'
'            If rstContextMenu(ContextMenuEnum.mnxClave) = "ACTUALIZAR" Then
'               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
'            End If
'
'         End If
'
'         If TypeOf ActiveControl Is DataCombo Then
'
'            If rstContextMenu(ContextMenuEnum.mnxClave) = "FINDQUERYDB" Then
'               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
'            End If
'
'            If rstContextMenu(ContextMenuEnum.mnxClave) = "ADMINISTRAR" Then
'
'               rstMenu.Filter = "MNU_CLAVE = '" & strMenuKeyAdmin & "'"
'               If Not rstMenu.EOF Then
''                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = ((TaskIsEnabled(rstMenu("MNU_TAREA").Value, CUsuario)))
'                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = ((TaskIsEnabled(rstMenu("MNU_CLAVE").Value, CUsuario)))
'               Else
'                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
'               End If
'
'            End If
'
'            If rstContextMenu(ContextMenuEnum.mnxClave) = "ACTUALIZAR" Then
'               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True
'            End If
'
'         End If
'
'
'         'ahora es el momento de cargar el separador
'         If bPrintSeparator Then
'
'            Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
'            ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True
'            ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Caption = "-"
'
'         End If
'
'         If rstContextMenu(ContextMenuEnum.mnxClave) = "FINDQUERYDB" Then
'            'seteo el frmFind
''            Set frmFind.CalledByForm = ActiveForm
''            Set frmFind.CalledByControl = ActiveControl
'
'         End If
'
'         ItemNumber = 1
'
'         rstContextMenu.MoveNext
'      Loop
'
'      If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is PowerMask Then
'
'         'Creo el menu Edit del control
'         AddEditContextMenu ActiveForm
'
'         ItemNumber = ActiveForm.mnuContextItem.UBound
'
'      End If
'
'   End If
'
'   If Len(strMenuName) = 0 Then Exit Sub
'
'   'creo el resto del menu
'   rstContextMenu.Filter = "MNX_NOMBRE = '" & strMenuName & "'"
'
'   If rstContextMenu.RecordCount > 0 Then
'      If ItemNumber > 0 Then
'         Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
'         ItemNumber = ActiveForm.mnuContextItem.UBound
'         ActiveForm.mnuContextItem(ItemNumber).Caption = "-"
'       End If
'   End If
'
'   Do While Not rstContextMenu.EOF
'
'      If IsNull(rstContextMenu(ContextMenuEnum.mnxForms)) Or InStr(rstContextMenu(ContextMenuEnum.mnxForms), ActiveForm.Name) > 0 Then
'
'         strCaption = rstContextMenu(ContextMenuEnum.mnxCaption)
'         If InStr(strCaption, "%S") > 0 Then
'            strCaption = Replace(strCaption, "%S", NullString)
'            bPrintSeparator = True
'         Else
'            bPrintSeparator = False
'         End If
'
'         If ItemNumber > 0 Then
'            'el menu ya posee mas de 1 item
'            Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
'            ItemNumber = ActiveForm.mnuContextItem.UBound
'            ActiveForm.mnuContextItem(ItemNumber).Caption = strCaption
'         Else
'            ActiveForm.mnuContextItem(ItemNumber).Caption = strCaption
'         End If
'
'         'si el caption termina con la secuenacia "%S" entonces trazo una linea
'         If bPrintSeparator Then
'
'            Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
'            ItemNumber = ActiveForm.mnuContextItem.UBound
'            ActiveForm.mnuContextItem(ItemNumber).Tag = NullString
'            ActiveForm.mnuContextItem(ItemNumber).Caption = "-"
'
'         End If
'
'         'salvo la clave en el tag asi el programa llamador puede interrogar dicho tag
'         ActiveForm.mnuContextItem(IIf(bPrintSeparator, ItemNumber - 1, ItemNumber)).Tag = rstContextMenu(ContextMenuEnum.mnxClave)
'
'         If Not IsNull(rstContextMenu(ContextMenuEnum.mnxForms)) Then
'            If (Not TaskIsEnabled(rstContextMenu(ContextMenuEnum.mnxForms), CUsuario)) Then ActiveForm.mnuContextItem(ItemNumber).Enabled = False
'         End If
'
'      End If
'
'      If ItemNumber = 0 Then ItemNumber = 1
'
'      rstContextMenu.MoveNext
'   Loop
'
'   rstContextMenu.Filter = adFilterNone
   
   
   
                           
Dim ActiveForm As Object
Dim ActiveControl As Control
Dim ItemNumber As Integer
Dim strCaption As String
Dim bPrintSeparator As Boolean

   'Carga el menu contextual strMenuName. Para los controles de tipo TEXT o DATACOMBO además
   'agrega el menu standard para estos tipos de controles
   
   ''''On Error GoTo GestErr
   
   On Error GoTo GestError
   
   If mvarForm Is Nothing Then
      Set ActiveForm = Screen.ActiveForm
   Else
      Set ActiveForm = mvarForm
   End If
   Set ActiveControl = Screen.ActiveControl

'   Set ActiveForm = mvarForm
'   Set ActiveControl = ActiveForm.ActiveControl

   'creo el menú particular si está definido
   rstContextMenu.Filter = "MNX_NOMBRE = '" & strMenuName & "'"

   If rstContextMenu.RecordCount > 0 Then
      If ItemNumber > 0 Then
         Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
         ItemNumber = ActiveForm.mnuContextItem.UBound
         ActiveForm.mnuContextItem(ItemNumber).Caption = "-"
         ActiveForm.mnuContextItem(ItemNumber).Tag = NullString
       End If
   End If
   
   Do While Not rstContextMenu.EOF
      
      If IsNull(rstContextMenu(ContextMenuEnum.mnxForms)) Or InStr(rstContextMenu(ContextMenuEnum.mnxForms), ActiveForm.Name) > 0 Then
      
         strCaption = rstContextMenu(ContextMenuEnum.mnxCaption)
         If InStr(strCaption, "%S") > 0 Then
            strCaption = Replace(strCaption, "%S", NullString)
            bPrintSeparator = True
         Else
            bPrintSeparator = False
         End If
         
         If ItemNumber > 0 Then
            'el menu ya posee mas de 1 item
            Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
            ItemNumber = ActiveForm.mnuContextItem.UBound
            ActiveForm.mnuContextItem(ItemNumber).Caption = strCaption
         Else
            ActiveForm.mnuContextItem(ItemNumber).Caption = strCaption
         End If
         
         'salvo la clave en el tag asi el programa llamador puede interrogar dicho tag
         ActiveForm.mnuContextItem(ItemNumber).Tag = rstContextMenu(ContextMenuEnum.mnxClave)
         
         'si el caption termina con la secuenacia "%S" entonces trazo una linea
         If bPrintSeparator Then
         
            Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
            ItemNumber = ActiveForm.mnuContextItem.UBound
            ActiveForm.mnuContextItem(ItemNumber).Tag = NullString
            ActiveForm.mnuContextItem(ItemNumber).Caption = "-"
         
         End If
         
         If Not IsNull(rstContextMenu(ContextMenuEnum.mnxForms)) Then
            If (Not TaskIsEnabled(rstContextMenu(ContextMenuEnum.mnxForms), CUsuario)) Then ActiveForm.mnuContextItem(ItemNumber).Enabled = False
         End If
         
      End If
      
      If ItemNumber = 0 Then ItemNumber = 1
      
      rstContextMenu.MoveNext
   Loop

   rstContextMenu.Filter = adFilterNone


   'Creo el resto del menú según el tipo de control
   'If bCreateDefaultMenu Then
      If TypeName(ActiveControl) = "TextBox" Or _
         TypeName(ActiveControl) = "PowerMask" Or _
         TypeName(ActiveControl) = "DataCombo" Or _
         TypeName(ActiveControl) = "DataGrid" Or _
         TypeName(ActiveControl) = "ComboBox" Or _
         TypeName(ActiveControl) = "OptionButton" Or _
         TypeName(ActiveControl) = "CheckBox" Then
'         TypeName(ActiveControl) = "CRViewer" Then
      
      
         Select Case TypeName(ActiveControl)
 '           Case "CRViewer"
               ' si es una CRViewer agrego el menu standard para ese control
 '              rstContextMenu.Filter = "MNX_NOMBRE = 'ZOOM_REPORT'"
               
            Case "DataGrid"
               ' si es una DataGrid agrego el menu standard para ese control
               rstContextMenu.Filter = "MNX_NOMBRE = 'GRILLA'"
            
            Case "TextBox", "PowerMask", "DataCombo"
                  ' si es un TextBox o DataCombo agrego el menu standard para estos controles
                  rstContextMenu.Filter = "MNX_NOMBRE = 'TEXT'"
               
            Case "CheckBox", "ComboBox", "OptionButton"
                  rstContextMenu.Filter = "MNX_NOMBRE = 'PERSONALIZAR'"
         End Select
         
         'empiezo la lectura del menu desplegable
         Do While Not rstContextMenu.EOF
            
            strCaption = rstContextMenu(ContextMenuEnum.mnxCaption)
            
            '-- controlo si despues de este item es necesario un separador
            If InStr(strCaption, "%S") > 0 Then
               strCaption = Replace(strCaption, "%S", NullString)
               bPrintSeparator = True
            Else
               bPrintSeparator = False
            End If
            
            If ItemNumber > 0 Then
               'ya fue cargado el primer item
               Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
            End If
            
            ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True
            ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Caption = strCaption
            
            'salvo la clave en el tag asi el programa llamador puede interrogar dicho tag
            ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Tag = rstContextMenu(ContextMenuEnum.mnxClave)
            
            ' habilito/deshabilito segun corresponda
            If TypeName(ActiveControl) = "TextBox" Or TypeName(ActiveControl) = "PowerMask" Then
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "FINDQUERYDB" Then
                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True   '''CanSearch
               End If
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "ADMINISTRAR" Then
                  If ActiveForm.Name = "frmCustomizeControl" Then
                 'If mvarForm.Name = "frmCustomizeControl" Or mvarCallerObjName = "frmCustomizeControl" Then
                     ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
                  Else
                     ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True  '''CanAdmin
                  End If
               End If
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "ACTUALIZAR" Then
                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
               End If
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "CUSTOMIZE" Then
                  If ActiveForm.Name = "frmCustomizeControl" Then
                 'If mvarForm.Name = "frmCustomizeControl" Or mvarCallerObjName = "frmCustomizeControl" Then
                     ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
                  End If
               End If
               
            End If
            
            If TypeName(ActiveControl) = "DataCombo" Then
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "FINDQUERYDB" Then
                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
               End If
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "ADMINISTRAR" Then
                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True ''CanAdmin
               End If
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "ACTUALIZAR" Then
                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True
               End If
               
            End If
            
                     
            'ahora es el momento de cargar el separador
            If bPrintSeparator Then
            
               Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True
               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Caption = "-"
               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Tag = NullString
            End If
            
            ItemNumber = 1
            
            rstContextMenu.MoveNext
         Loop
         
         If TypeName(ActiveControl) = "TextBox" Or _
            TypeName(ActiveControl) = "PowerMask" Then
         
            'Creo el menu Edit del control
            AddEditContextMenu ActiveForm
            
            ItemNumber = ActiveForm.mnuContextItem.UBound
   
         End If
         
      End If
   'End If

   Exit Sub

GestError:
   LoadError ErrorLog, "SetContextMenu"
   ShowErrMsg ErrorLog
   
End Sub

Public Sub UnloadContextMenu(ByVal frm As Form)
Dim ix As Integer
   
   'Descarga el menu contextual activo del form pasado como argumento
   
   For ix = frm.mnuContextItem.UBound To 1 Step -1
      Unload frm.mnuContextItem(ix)
   Next ix
   frm.mnuContextItem(0).Caption = NullString
   frm.mnuContextItem(0).Enabled = True
   
End Sub

Public Sub DataComboLoad(ByVal oCombo As DataCombo, ByVal strEmpresa As String, _
                         ByVal strTableName As String, ByVal BoundColumn As String, _
                         ByVal ListField As String, Optional ByRef rstSource As ADODB.Recordset, _
                         Optional DefaultValue As String, Optional OrderBy As String)
                         
Dim strSQL     As String
Dim strOrderBy As String
Dim strOldValue As String
Dim rst        As ADODB.Recordset

   strOldValue = oCombo.BoundText
   
   If Len(OrderBy) = 0 Then
      strOrderBy = ListField
   Else
      strOrderBy = OrderBy
   End If
   
   If rstSource Is Nothing Then
      If BoundColumn <> ListField Then
'         strSQL = "SELECT " & BoundColumn & ", " & ListField & " FROM " & strTableName & " ORDER BY " & strOrderBy
         strSQL = "SELECT DISTINCT " & BoundColumn & ", " & ListField & " FROM " & strTableName & " ORDER BY " & strOrderBy
      Else
'         strSQL = "SELECT " & BoundColumn & " FROM " & strTableName & " ORDER BY " & strOrderBy
         strSQL = "SELECT DISTINCT " & BoundColumn & " FROM " & strTableName & " ORDER BY " & strOrderBy
      End If
   
      Set rst = OpenComboDataSource(strEmpresa, strSQL)
      Set rstSource = CopyData(rst, AllRecords)
      If rstSource.RecordCount > 0 Then rstSource.MoveFirst
      
      Set rst = Nothing
   End If
   
   Set oCombo.RowSource = rstSource
   oCombo.BoundColumn = BoundColumn
   oCombo.ListField = ListField
   
   oCombo.BoundText = strOldValue
   
   If Len(DefaultValue) > 0 Then
      oCombo.BoundText = DefaultValue
   End If
   
End Sub
Public Sub CenterMDIActiveXChild(ByVal frmChild As Form)

   '--  centra el form MDIActiveX Child
   
   frmChild.Move (mvarMDIForm.ScaleWidth - frmChild.Width) / 2, (mvarMDIForm.ScaleHeight - frmChild.Height) / 2

End Sub

Public Sub SetCaption(ByVal frm As Form, Optional ByVal strEmpresa As String, Optional ByVal Title As String)

   If InStr(frm.Caption, "(") > 0 Then
      frm.Caption = Trim(Left(frm.Caption, InStr(frm.Caption, "(") - 1))
   End If
   
   'si no viene un titulo, uso el del form
   If Title = NullString Then Title = Trim(frm.Caption)
   
   ' si viene una empresa, agrego el nombre de la empresa
   If strEmpresa <> NullString Then
   
      strEmpresa = NombreEmpresa(strEmpresa)
      
      If InStr(Title, NombreEmpresa(strEmpresa)) > 0 Then
         frm.Caption = Left(Title, InStr(Title, "(") - 1)
      End If
      
      frm.Caption = Title & " (" & strEmpresa & ")"
      
   Else
   
      strEmpresa = NullString
      frm.Caption = Title
      
   End If
   

End Sub


Public Function FindKey(strKey As String, TV As TreeView) As Boolean
Dim nodX As Node
  
   ' busco si existe la clave strKey en el treeview TV
   
   On Error Resume Next
   
   Set nodX = TV.Nodes(strKey)
   FindKey = (Not nodX Is Nothing)
   Err.Clear
   
End Function

Public Function TextLostFocus(ctl As Object, objBusiness As Object, strPropLet As String, Optional strPropGet As String) As String
'Dim iDato   As Integer
'Dim strDato As String
'Dim lngDato As Long
'Dim v       As Variant

   On Error Resume Next
   
   If Len(Trim(ctl.Text)) = 0 Then
      TextLostFocus = NullString
      Exit Function
   End If
   
'   v = CallByName(objBusiness, strPropLet, VbGet)
   
   CallByName objBusiness, strPropLet, VbLet, ctl.Text
   
   If Len(strPropGet) > 0 Then
      TextLostFocus = CallByName(objBusiness, strPropGet, VbGet)
   Else
      TextLostFocus = NullString
   End If
      
'   Select Case TypeName(v)
'      Case "Integer"
'         iDato = v
'         CallByName objBusiness, strPropLet, VbLet, iDato
'      Case "String"
'         strDato = v
'         CallByName objBusiness, strPropLet, VbLet, strDato
'      Case "Long"
'         lngDato = v
'         CallByName objBusiness, strPropLet, VbLet, lngDato
'   End Select
   
End Function

Public Sub AddEditContextMenu(ByVal frm As Form)

   '-- agrea los comandos de edición al menu contextual del form

   'le informo al objeto cual es el TextBox que debe controlar
   objTextBox.TextBox = frm.ActiveControl

   'Creo los items del menu Edit y activo/desactivo los items segun corresponda
   With frm.mnuContextItem
   
      If .UBound > 0 Then
         Load frm.mnuContextItem(.UBound + 1)
         frm.mnuContextItem(.UBound).Caption = "-"
         Load frm.mnuContextItem(.UBound + 1)
      End If
      
      frm.mnuContextItem(.UBound).Caption = "&Deshacer"
      frm.mnuContextItem(.UBound).Tag = "DESHACER"
      frm.mnuContextItem(.UBound).Enabled = objTextBox.CanUndo
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Caption = "-"
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Caption = "C&ortar"
      frm.mnuContextItem(.UBound).Tag = "CORTAR"
      frm.mnuContextItem(.UBound).Enabled = objTextBox.CanCut
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Caption = "&Copiar"
      frm.mnuContextItem(.UBound).Tag = "COPIAR"
      frm.mnuContextItem(.UBound).Enabled = objTextBox.CanCopy
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Caption = "&Pegar"
      frm.mnuContextItem(.UBound).Tag = "PEGAR"
      frm.mnuContextItem(.UBound).Enabled = objTextBox.CanPaste
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Caption = "Eli&minar"
      frm.mnuContextItem(.UBound).Tag = "ELIMINAR"
      frm.mnuContextItem(.UBound).Enabled = objTextBox.CanDelete
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Caption = "-"
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Caption = "&Seleccionar Todo"
      frm.mnuContextItem(.UBound).Tag = "SELTODO"
      frm.mnuContextItem(.UBound).Enabled = objTextBox.CanSelectAll
      
   End With
   
End Sub

Public Sub CreateMenuItem(ByVal MenuType As EnumMenu, ByRef Form As Form, Optional ByVal aMenuItems As Variant, Optional ByVal mnuItem As Object)
Dim aMnuFileItems(8)    As String
Dim aMnuEditItems(12)   As String
Dim aMnuToolsItems(13)  As String
Dim aMnuWindowItems(2)  As String
Dim aMnuHelpItems(4)    As String
Dim ix                  As Integer
Dim iPos                As Integer

   On Error GoTo GestError
   
   Select Case MenuType
         
      Case MenuCustom

         For ix = LBound(aMenuItems) To UBound(aMenuItems)
            If ix > 0 Then
               Load mnuItem(ix)
            End If
            mnuItem(mnuItem.UBound).Caption = aMenuItems(mnuItem.UBound)
         Next ix

      Case MenuFile
         'Menu File
         aMnuFileItems(0) = "&Propiedades ..."
         aMnuFileItems(1) = "-"
         aMnuFileItems(2) = "&Configurar Página ..."
         aMnuFileItems(3) = "&Vista Preliminar"
         aMnuFileItems(4) = "&Imprimir ..."
         aMnuFileItems(5) = "_"
         aMnuFileItems(6) = "&Enviar ..."
         aMnuFileItems(7) = "-"
         aMnuFileItems(8) = "&Salir"
      
      Case MenuEdit
         'Menu Edición
         aMnuEditItems(0) = "&Deshacer" & vbTab & "Ctrl+Z"
         aMnuEditItems(1) = "-"
         aMnuEditItems(2) = "Cor&tar" & vbTab & "Ctrl+X"
         aMnuEditItems(3) = "&Copiar" & vbTab & "Ctrl+C"
         aMnuEditItems(4) = "&Pegar" & vbTab & "Ctrl+V"
         aMnuEditItems(5) = "&Eliminar" & vbTab & "Supr"
         aMnuEditItems(6) = "-"
         aMnuEditItems(7) = "&Seleccionar Todo" & vbTab & "Ctrl+E"
         aMnuEditItems(8) = "-"
         aMnuEditItems(9) = "&Buscar..." & vbTab & "Ctrl+F"
         aMnuEditItems(10) = "&Administrar" & vbTab & "Ctrl+A"
         aMnuEditItems(11) = "-"
         aMnuEditItems(12) = "A&ctualizar"

         For ix = LBound(aMnuEditItems) To UBound(aMnuEditItems)
            If ix > 0 Then
               Load Form.mnuEditItems(ix)
            End If
            Form.mnuEditItems(Form.mnuEditItems.UBound).Caption = aMnuEditItems(Form.mnuEditItems.UBound)
         Next ix

      Case MenuTools
         'Menu Herramientas
         
         aMnuToolsItems(0) = "&Cambio de Contraseña ...;frmCambioPWD"
         'aMnuToolsItems(1) = "&Establecer Ejercicio Contable Activo...;frmEjercicioContableActivo"
         aMnuToolsItems(2) = "&Iniciar sesión como un Usuario distinto;ReiniciarSesion"
         aMnuToolsItems(3) = "&Empresas ...;frmSeleccionEmpresas"
         aMnuToolsItems(4) = "-"
         aMnuToolsItems(5) = "&Reporte de Errores...;frmListadoErrores"
         'aMnuToolsItems(6) = "Editor &SQL...;frmSQLRun"
         aMnuToolsItems(7) = "-"
         aMnuToolsItems(8) = "&Buscar opción del menú...;BuscarMenu"
         aMnuToolsItems(9) = "Buscar &siguiente;BuscarSiguienteMenu"
         aMnuToolsItems(10) = "-"
         aMnuToolsItems(11) = "&Propiedades de la Ventana...;PropiedadesVentana"
         aMnuToolsItems(12) = "-"
         aMnuToolsItems(13) = "&Opciones...;frmOpciones"
         
         LoadMenu "Tools", Form, aMnuToolsItems

      Case MenuWindow
         'Menu Ventana
         aMnuWindowItems(0) = "&Cascada"
         aMnuWindowItems(1) = "&Mosaico Horizontal"
         aMnuWindowItems(2) = "Mosaico &Vertical"

         For ix = LBound(aMnuWindowItems) To UBound(aMnuWindowItems)
            If ix > 0 Then
               Load Form.mnuWindowItems(ix)
            End If
            Form.mnuWindowItems(Form.mnuWindowItems.UBound).Caption = aMnuWindowItems(Form.mnuWindowItems.UBound)
         Next ix

      Case MenuHelp
         'Menu Ayuda
         
         aMnuHelpItems(0) = "&Índice...;"
         aMnuHelpItems(1) = "-"
         aMnuHelpItems(2) = "&Soporte Técnico;"
         aMnuHelpItems(3) = "-"
         aMnuHelpItems(4) = "&Acerca de " & App.ProductName & "...;"

          For ix = 0 To UBound(aMnuHelpItems)
            If ix > 0 Then
               Load Form.mnuHelpItems(ix)
            End If
            iPos = InStr(aMnuHelpItems(ix), ";")
            If iPos > 0 Then
               Form.mnuHelpItems(ix).Caption = Left(aMnuHelpItems(ix), iPos - 1)
            Else
               Form.mnuHelpItems(ix).Caption = aMnuHelpItems(ix)
            End If
         Next ix

   End Select
   
   Exit Sub
   
GestError:
   LoadError ErrorLog, "CreateMenuItem"
   ShowErrMsg ErrorLog
End Sub

Public Sub LoadMenu(strMenuName As String, frm As Form, aMenuTools() As String)
Dim ix   As Integer
Dim iPos As Integer

   '  carga el menu indicado por el parametro strMenuname en el form frm
   
   On Error Resume Next
   
   Select Case strMenuName
      Case "Tools"
         
         '  menu Herramientas
         For ix = 0 To UBound(aMenuTools)
            If Len(aMenuTools(ix)) > 0 Then
               If ix > 0 Then
                  Load frm.mnuToolsItems(ix)
               End If
               iPos = InStr(aMenuTools(ix), ";")
               If iPos > 0 Then
                  frm.mnuToolsItems(ix).Caption = Left(aMenuTools(ix), iPos - 1)
               Else
                  frm.mnuToolsItems(ix).Caption = aMenuTools(ix)
               End If
            End If
         Next ix
      Case "Help"
       
       ' menu  Ayuda
       For ix = 0 To UBound(aMenuHelp)
         If ix > 0 Then
            Load frm.mnuAyudaItem(ix)
         End If
         iPos = InStr(aMenuHelp(ix), ";")
         If iPos > 0 Then
            frm.mnuAyudaItem(ix).Caption = Left(aMenuHelp(ix), iPos - 1)
         Else
            frm.mnuAyudaItem(ix).Caption = aMenuHelp(ix)
         End If
      Next ix
   End Select

End Sub

Public Function CallAdmin(ByVal ControlInfo As Variant, ByVal ControlData As Variant) As Long

   '-- llamo al Form para la administracion del control activo
   '-- ControlInfo es del tipo ControlType
   
   On Error GoTo GestErr
   
   CallAdmin = mvarMDIForm.CallAdmin(ControlInfo.MenuKeyAdmin, ControlData)
   
   Exit Function
   
GestErr:
   LoadError ErrorLog, "CallAdmin"
   ShowErrMsg ErrorLog
End Function

Public Function RemoveSpaces(ByVal strSource As String) As String
    RemoveSpaces = ReplaceText(strSource)
End Function

Public Function ReplaceText(ByVal SourceText, Optional ByVal StrToReplace = " ", Optional ByVal StrToInsert = "") As String
Dim RetS As String
Dim ix   As Integer

   If StrToReplace = StrToInsert Or StrToReplace = "" Then Exit Function
   RetS = SourceText
   ix = InStr(RetS, StrToReplace)
   Do While ix <> 0
       RetS = IIf(ix = 1, "", Left(RetS, ix - 1)) & StrToInsert & IIf(ix = Len(RetS) - Len(StrToReplace) + 1, "", Right(RetS, Len(RetS) - ix - Len(StrToReplace) + 1))
       ix = InStr(RetS, StrToReplace)
   Loop
   ReplaceText = RetS
    
End Function

Public Sub DuplicateComboBox(source As ComboBox, Target As ComboBox, _
                             Optional AppendMode As Boolean)
                             
Dim Index As Long
Dim itmData As Long
Dim numItems As Long
Dim sItemText As String
    
    ' prepare the receiving buffer
    sItemText = Space$(512)
    
    ' temporarily prevent updating
    LockWindowUpdate Target.hWnd
    
    ' reset target contents, if not in append mode
    If Not AppendMode Then
        SendMessage Target.hWnd, CB_RESETCONTENT, 0, ByVal 0&
    End If
    
    ' get the number of items in the source control
    numItems = SendMessage(source.hWnd, CB_GETCOUNT, 0&, ByVal 0&)
    
    For Index = 0 To numItems - 1
        ' get the item text
        SendMessage source.hWnd, CB_GETLBTEXT, Index, ByVal sItemText
        ' get the item data
        itmData = SendMessage(source.hWnd, CB_GETITEMDATA, Index, ByVal 0&)
        ' add the item text to the target list
        SendMessage Target.hWnd, CB_ADDSTRING, 0&, ByVal sItemText
        ' add the item data to the target list
        SendMessage Target.hWnd, CB_SETITEMDATA, Index, ByVal itmData
    Next
    
    ' allow redrawing
    LockWindowUpdate 0
    
End Sub

Public Function GetControlInformation(ByRef ControlData As udtControlData) As String
Dim PropBag As PropertyBag

   Set PropBag = New PropertyBag
   
   With PropBag
      .WriteProperty "EMPRESA", ControlData.Empresa
      .WriteProperty "SUCURSAL", ControlData.Sucursal
      .WriteProperty "USUARIO", ControlData.Usuario
      .WriteProperty "MAQUINA", ControlData.Maquina
      .WriteProperty "MENUKEY", ControlData.MenuKey
      .WriteProperty "EJERCICIO", ControlData.Ejercicio
   End With
   
   GetControlInformation = PropBag.Contents
   
End Function

Public Sub CenterForm(ByRef frm As Form)
Dim r As RECT
Dim lRes As Long
Dim lw As Long
Dim lh As Long

   lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, r, 0)

   If lRes Then
      With r
         .Left = Screen.TwipsPerPixelX * .Left
         .Top = Screen.TwipsPerPixelY * .Top
         .Right = Screen.TwipsPerPixelX * .Right
         .Bottom = Screen.TwipsPerPixelY * .Bottom
         lw = .Right - .Left
         lh = .Bottom - .Top
         
         frm.Move .Left + (lw - frm.Width) \ 2, .Top + (lh - frm.Height) \ 2
      End With
   End If

End Sub

Public Sub GetDefaultOptions(ByVal frm As Form)
      Dim Ctrl       As Control
      Dim ix         As Integer
'      Dim iOptionTrue As Integer
      Dim strKey     As String

10       On Error GoTo GestErr
   
20       strKey = "Dialogs\Defaults\" & CUsuario.Usuario & "\" & frm.MenuKey & "\"
   
30       ix = 0
40       For Each Ctrl In frm.Controls
50          If Ctrl.Tag = "Save" Then ix = ix + 1
60       Next Ctrl
   
70       If ix = 0 Then Exit Sub
   
80       ReDim aKeys(ix - 1, 1)
90       ix = 0
100      For Each Ctrl In frm.Controls

110         If Ctrl.Tag = "Save" Then
120            aKeys(ix, 0) = strKey & Ctrl.Name
130            ix = ix + 1
140         End If

150      Next Ctrl
   
         'leo las claves
160      GetKeyValues frm.ControlData.Empresa, aKeys
   
170      ix = 0
180      For Each Ctrl In frm.Controls

190         If Ctrl.Tag = "Save" Then
200            vValue = aKeys(ix, 1)
               'me fijo si tiene ya definido un valor default
210            If Not IsNull(vValue) Then
220               Select Case TypeName(Ctrl)
                     Case "TextBox", "RichTextBox", "PowerMask"
230                     Ctrl.Text = vValue
240                  Case "ComboBox"
250                     If Ctrl.Style = vbComboDropdownList Then
260                        Ctrl.ListIndex = ComboSearch(Ctrl, vValue)
270                     Else
280                        Ctrl.Text = vValue
290                     End If
                     Case "DataCombo"
                        Ctrl.BoundText = vValue
300                  Case "CheckBox"
310                     Ctrl.Value = IIf(vValue = "1", vbChecked, vbUnchecked)
320                  Case "OptionButton"
330                     On Error Resume Next
340                     Ctrl.Value = CBool(vValue)
350                     If Err.Number = 0 Then
360                        If Ctrl.Value = True Then
370                           Ctrl.Value = Not Ctrl.Value
380                           Ctrl.Value = Not Ctrl.Value
390                        End If
400                     End If
410                     On Error GoTo GestErr
420               End Select
430            End If
440            ix = ix + 1
450         End If
460      Next Ctrl

470      Exit Sub
   
GestErr:
480      LoadError ErrorLog, "GetDefaultOptions" & Erl
490      ShowErrMsg ErrorLog
End Sub

Public Sub GuardarOpciones(frm As Form)
Dim cntl                As Control
Dim aKeysToSave()       As String
Dim strKey              As String
Dim strValor            As String
Dim ix                  As Integer
Dim objTabla            As BOGeneral.clsTablas

   On Error GoTo GestErr
   
   Set objTabla = New BOGeneral.clsTablas
   
   strKey = "Dialogs\Defaults\" & CUsuario.Usuario & "\" & frm.MenuKey & "\"
   
   ix = 0
   For Each cntl In frm.Controls
      If cntl.Tag = "Save" Then ix = ix + 1
   Next cntl
   
   If ix = 0 Then Exit Sub
   
   ReDim aKeysToSave(ix - 1, 1)
   ix = 0
   
   For Each cntl In frm.Controls
      
      If cntl.Tag = "Save" Then
      
         Select Case TypeName(cntl)
            Case "TextBox", "RichTextBox", "ComboBox"
               strValor = cntl.Text
            Case "CheckBox"
               strValor = cntl.Value
            Case "OptionButton"
               strValor = cntl.Value
            Case "DataCombo"
               strValor = cntl.BoundText
         End Select
            
         aKeysToSave(ix, 0) = strKey & cntl.Name
         aKeysToSave(ix, 1) = strValor
         
         ix = ix + 1
      End If
   Next cntl
   
   Dim ControlData As udtControlData
   
   ControlData = frm.ControlData
   objTabla.ControlData = ControlData
'   objTabla.UpdateGlobal frm.ControlData.Empresa, aKeysToSave
   objTabla.UpdateGlobal NullString, aKeysToSave

   Exit Sub
   
GestErr:
   LoadError ErrorLog, "GuardarOpciones"
   ShowErrMsg ErrorLog

End Sub

Public Sub ListViewAdjustColumnWidth(LV As ListView, Optional AccountForHeaders As Boolean)
Dim col As Integer, lParam As Long
    
   'ajusta el ancho de cada columna del listview para que sea completamente visible
   'si el segundo parametro es true, tiene en cuenta el encabezamiento
    
    If AccountForHeaders Then
        lParam = LVSCW_AUTOSIZE_USEHEADER
    Else
        lParam = LVSCW_AUTOSIZE
    End If
    
    For col = 0 To LV.ColumnHeaders.Count
        SendMessage LV.hWnd, LVM_SETCOLUMNWIDTH, col, lParam
    Next

End Sub

Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
Dim c As ColumnHeader
 
 If Column Is Nothing Then
   For Each c In LV.ColumnHeaders
      SendMessage LV.hWnd, LVM_FIRST + 30, c.Index - 1, -1
   Next
 Else
   SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 
 LV.Refresh
 
End Sub


Public Sub SetFrameControls(ByVal FormObject As Form, _
                            ByVal FrameObject As Object, _
                            ByVal bEnabled As Boolean, _
                            Optional ByVal bInitialize As Boolean = True, _
                            Optional ByVal bSkipLabels As Boolean = False, _
                            Optional ByVal bExcludeMainFrame As Boolean = True, _
                            Optional ByVal strListNameExclude As String = NullString)
                            
Dim ctrlFocus As Control
Dim Ctrl As Control
Dim ctrl2 As Control
Dim Button As Button
Dim IndexFrame As Integer
Dim Index As Integer
Dim bFlag As Boolean

   ' Habilita/Deshabilita (con opcion inicializacion) de todos los controles de un frame
   
   On Error Resume Next
          
   Set ctrlFocus = FormObject.ActiveControl
   
   IndexFrame = FrameObject.Index
   If Err.Number <> 0 Then
      IndexFrame = -1
      Err.Clear
   End If
   
   'habilito/deshabilito el Frame (24/01/2003)
   If bExcludeMainFrame = False Then
      FrameObject.Enabled = bEnabled
   End If
   
   For Each Ctrl In FormObject.Controls
       
      If TypeOf Ctrl Is Frame Or TypeOf Ctrl Is SSTab Then
      
         If Ctrl.Name = FrameObject.Name Then
            
            If IndexFrame = -1 Then
               'el Frame no forma parte de un arreglo
               bFlag = True
            Else
               'el Frame no forma parte de un arreglo
               Index = Ctrl.Index
               If Err.Number <> 0 Then
                  'Ctrl es un frame que no forma parte del arreglo
                  bFlag = False
                  Err.Clear
               Else
                  If Index = IndexFrame Then
                     'Ctrl es el frame
                     bFlag = True
                  End If
               End If
               
            End If
               
            If bFlag Then
               bFlag = False
               For Each ctrl2 In FormObject.Controls
               
                  If (ctrl2.Container Is Ctrl) Then
                      
                      If Err.Number = 0 Then
                         Err.Clear
                         'el control admite la propiedad Container
                      
                         'si el control esta subclasado, cambia tambien el BackColor
                         If TypeOf ctrl2 Is Toolbar Then
                            'las toolbar no las desabilito, solo sus botones
                         Else
                           If InStr(strListNameExclude, ctrl2.Name) = 0 Then
                              If TypeOf ctrl2 Is Label And bSkipLabels Then
                                 ctrl2.Enabled = True
                              Else
                                 ctrl2.Enabled = bEnabled
                              End If
                           End If
                         End If
                         
                         If TypeOf ctrl2 Is Label Then
                            'el label hay que hacerlo a pulmon
                            If Left(ctrl2.Name, 3) = "lbl" Then
                               ctrl2.BackColor = IIf(bEnabled = True, LabelColor, DisabledColor)
                            End If
                         End If
                         
                         If bInitialize Then
                            
                            If TypeOf ctrl2 Is OptionButton Then
                                ctrl2.Value = False
                            ElseIf TypeOf ctrl2 Is CheckBox Then
                                ctrl2.Value = vbUnchecked
                            ElseIf TypeOf ctrl2 Is TextBox Then
                                ctrl2.Text = NullString
                            ElseIf TypeOf ctrl2 Is PowerMask Then
                                ctrl2.Text = NullString
                            ElseIf TypeOf ctrl2 Is ComboBox Then
                               If ctrl2.Style = vbComboDropdownList Then
                                  ctrl2.ListIndex = -1
                               Else
                                  ctrl2.Text = NullString
                               End If
                            ElseIf TypeOf ctrl2 Is DataCombo Then
                                ctrl2.BoundText = NullString
                            ElseIf TypeOf ctrl2 Is Label Then
                               If Left(ctrl2.Name, 3) = "lbl" Then
                                  ctrl2.Caption = NullString
                               End If
                            ElseIf TypeOf ctrl2 Is RichTextBox Then
                                ctrl2.Text = NullString
                            ElseIf TypeOf ctrl2 Is ListView Then
                                ctrl2.ListItems.Clear
                            ElseIf TypeOf ctrl2 Is Toolbar Then
                               For Each Button In ctrl2.Buttons
                                  Button.Enabled = False
                               Next Button
                            End If
                         
                         End If
                      
                      Else
                         Err.Clear
                      End If
                      
                  End If
                     
               Next
            
            End If
           
        End If
        
      End If
      
   Next
   
   'si despues de toda esta operacion ctrlFocus continua habilitado le doy el foco
   If ctrlFocus.Visible And ctrlFocus.Enabled Then ctrlFocus.SetFocus
   
End Sub

Public Function GetModuleName(ByVal SourceModule As Variant) As String

   Select Case VarType(SourceModule)
      Case vbObject:
         GetModuleName = TypeName(SourceModule)
      Case vbString:
         GetModuleName = SourceModule
      Case Else:
         GetModuleName = "<Modulo Desconocido>"
   End Select

End Function

Public Function PrinterIsValid(ByVal objImpresora As BOGeneral.clsImpresoras, ByRef ImpresoraFormulario As String, ByRef strCustomMess1 As String, ByRef strCustomMess2 As String) As Boolean
Dim bDefinidaLocalmente As Boolean
Dim objImpresoraLocal As BOGeneral.clsImpresoraLocal

   'valida que la impresora SharePrinterName (Ej. \\FERNANDO\LEXMARK) este definida localmente
   'y ademas que aún este compartida.
   'si impresora es válida devuelve el nombre con el que esta definida localmente, caso contrario
   'devuelve un mensaje
   
   PrinterIsValid = False
   
   strCustomMess1 = "La Impresora " & objImpresora.Codigo & " no se encuentra instalada en este Puesto de Trabajo. "
   strCustomMess2 = "Es posible instalarla o si lo desea, imprimir el formulario en una de las impresoras definidas localmente"
   
   If objImpresora Is Nothing Then Exit Function
   
   ' controlo si esta definida localmente
   bDefinidaLocalmente = False
   For Each objImpresoraLocal In objImpresora.ImpresorasLocales
   
      If UCase(objImpresoraLocal.Maquina) = UCase(Right(Machine, Len(Machine) - InStr(Machine, "@"))) Then
         bDefinidaLocalmente = True
         ImpresoraFormulario = objImpresoraLocal.NombreLocal
         Exit For
      End If
      
   Next objImpresoraLocal
   
   PrinterIsValid = bDefinidaLocalmente
   
   Exit Function


'    Select Case WinVersion
'      Case "Windows 95", "Windows 98"
'
'         For Each prt In Printers
'
'            ShareName = PrinterShareName(prt.Port)
'            If ShareName <> NullString Then
'               'es una impresora local compartida. Le agrego el nombre de mi PC
'               ShareName = "\\" & Machine & "\" & ShareName
'
'               If SharePrinterName = ShareName Then
'                  ImpresoraFormulario = prt.DeviceName
'                  PrinterIsValid = True
'                  Exit For
'               End If
'            End If
'            If UCase(SharePrinterName) = UCase(prt.Port) Then
'               'es una impresora de la red, correctamente instalada en mi PC
'               bNetPrinter = True
'               ImpresoraFormulario = prt.DeviceName
'               PrinterIsValid = True
'               Exit For
'            End If
'         Next prt
'
'      Case "Windows 2000", "Windows NT"
'          aInfo = Split(PrintersWinNT(SharePrinterName), ";")
'          PrinterIsValid = (IsArrayEmpty(aInfo) = False)
'          If IsArrayEmpty(aInfo) = False Then
'            bNetPrinter = (InStr(aInfo(1), "net ") > 0)
'          End If
'      Case Else
'          PrinterIsValid = False
'    End Select
'
'    If PrinterIsValid Then
'        '
'        ' controlo que aún este compartida
'        '
'        If bNetPrinter Then
'            Set dlgImpresoras = New ALGControls.NETResource
'            '
'            'chequeo que el recurso aún esta compartido
'            '
'            If dlgImpresoras.CheckResource(UCase(SharePrinterName), Impresoras) Then
'                Select Case WinVersion
'                   Case "Windows 95", "Windows 98"
'                        'ImpresoraFormulario = prt.DeviceName
'                   Case "Windows 2000", "Windows NT"
'                        ImpresoraFormulario = SharePrinterName
'                End Select
'               PrinterIsValid = True
'            Else
'               strCustomMess1 = "La Impresora de Red  " & SharePrinterName & " no es un Recurso Compartido. "
'               strCustomMess2 = "Si lo desea, es posible imprimir el formulario en una de las impresoras definidas localmente"
'               PrinterIsValid = False
'            End If
'        End If
'    End If
    
End Function

Public Function ExistItem(ByVal oCol As Collection, ByVal Key As Variant) As Boolean
Dim vValue As Variant

   On Error Resume Next
   vValue = oCol.Item(Key)
   ExistItem = (Err.Number = 0)
   Err.Clear
        
End Function

Public Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function

Public Function PrinterShareName(ByVal printerName As String) As String
Dim hPrinter As Long
Dim tPI As PRINTER_INFO_2
Dim abBuf() As Byte
Dim lBufLen As Long

   ' abro la impresora
   If OpenPrinter(printerName, hPrinter, 0&) Then

      ' obtengo la dimension del buffer
      GetPrinter hPrinter, 2, ByVal 0&, 0, lBufLen

      ReDim abBuf(0 To lBufLen)

      ' obtengo información de la impresora
      GetPrinter hPrinter, 2, abBuf(0), lBufLen + 1, lBufLen

      ' copio el array a la UDT
      MoveMemory tPI, abBuf(0), Len(tPI)

      ' obtengo el nombre compartido a partir de su puntero
      PrinterShareName = String$(lstrlenA(tPI.pShareName), 0)
      lstrcpyA PrinterShareName, tPI.pShareName

      ' cierro la impresora
      ClosePrinter hPrinter

   End If

'   Dim pntr As PRINTER_INFO_4
'
'   If OpenPrinter(PrinterName, hPrinter, 0&) Then
'
'      ' obtengo la dimension del buffer
'      GetPrinter hPrinter, 2, ByVal 0&, 0, lBufLen
'
'      ReDim abBuf(0 To lBufLen)
'
'      ' obtengo información de la impresora
'      GetPrinter hPrinter, 2, abBuf(0), lBufLen + 1, lBufLen
'
'      ' copio el array a la UDT
'      MoveMemory pntr, abBuf(0), Len(pntr)
'
'      ' obtengo el nombre compartido a partir de su puntero
'      PrinterShareName = String$(lstrlenA(pntr.pPrinterName), 0)
'      lstrcpyA PrinterShareName, pntr.pPrinterName
'
'      ' cierro la impresora
'      ClosePrinter hPrinter
'
'   End If
'

End Function

Public Function PrintersWinNT(ByVal printerName As String) As String
    
'   Dim Success As Boolean
   Dim cbRequired As Long
   Dim cbBuffer As Long
   Dim pntr() As PRINTER_INFO_4
   Dim nEntries As Long
   Dim c As Long
   Dim sAttr As String
   
  'To determine the required buffer size, call EnumPrinters with
  'cbBuffer set to zero. EnumPrinters fails, and Err.LastDLLError
  'returns ERROR_INSUFFICIENT_BUFFER, filling in the cbRequired
  'parameter with the size, in bytes, of the buffer required to
  'hold the array of structures and their data.
   Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, _
                     vbNullString, PRINTER_LEVEL4, _
                     0, 0, cbRequired, nEntries)
            
   
  'The strings pointed to by each PRINTER_INFO_4 struct's members
  'reside in memory after the end of the array of structs. So we're
  'not only allocating memory for the structs themselves, but all the
  'strings pointed to by each struct's member as well.
   ReDim pntr((cbRequired \ SIZEOFPRINTER_INFO_4))
       
  'Set cbBuffer equal to the size of the buffer
   cbBuffer = cbRequired
    
  'Enumerate the printers. If the function succeeds,
  'the return value is nonzero. If the function fails,
  'the return value is zero.
   If EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, _
                   vbNullString, PRINTER_LEVEL4, _
                   pntr(0), cbBuffer, _
                   cbRequired, nEntries) Then
              
      For c = 0 To nEntries - 1
           
         With pntr(c)
            
            sAttr = ""
            
            If GetStrFromPtrA(.pPrinterName) = printerName Then
            
                If (.Attributes And PRINTER_ATTRIBUTE_DEFAULT) Then sAttr = "default "
                If (.Attributes And PRINTER_ATTRIBUTE_DIRECT) Then sAttr = sAttr & "direct "
                If (.Attributes And PRINTER_ATTRIBUTE_ENABLE_BIDI) Then sAttr = sAttr & "bidirectional "
                If (.Attributes And PRINTER_ATTRIBUTE_LOCAL) Then sAttr = sAttr & "local "
                If (.Attributes And PRINTER_ATTRIBUTE_NETWORK) Then sAttr = sAttr & "net "
                If (.Attributes And PRINTER_ATTRIBUTE_QUEUED) Then sAttr = sAttr & "queued "
                If (.Attributes And PRINTER_ATTRIBUTE_SHARED) Then sAttr = sAttr & "shared "
                If (.Attributes And PRINTER_ATTRIBUTE_WORK_OFFLINE) Then sAttr = sAttr & "offline "
                
                PrintersWinNT = GetStrFromPtrA(.pPrinterName) & ";" & sAttr
                Exit For
            End If
            

         End With
              
      Next c
        
   Else: 'ctl.AddItem "Error enumerating printers."
   End If  'EnumPrinters
   
End Function

Public Function PrintersWin9x(ByVal printerName As String) As String
    
   Dim cbRequired As Long
   Dim cbBuffer As Long
   Dim pntr() As PRINTER_INFO_1
   Dim nEntries As Long
   Dim c As Long
   Dim sFlags As String
    
  'To determine the required buffer size, call EnumPrinters with
  'cbBuffer set to zero. EnumPrinters fails, and Err.LastDLLError
  'returns ERROR_INSUFFICIENT_BUFFER, filling in the cbRequired
  'parameter with the size, in bytes, of the buffer required to
  'hold the array of structures and their data.
   Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, _
                     vbNullString, PRINTER_LEVEL1, _
                     0, 0, cbRequired, nEntries)
                     
  'The strings pointed to by each PRINTER_INFO_4 struct's members
  'reside in memory after the end of the array of structs. So we're
  'not only allocating memory for the structs themselves, but all the
  'strings pointed to by each struct's member as well.
   ReDim pntr((cbRequired \ SIZEOFPRINTER_INFO_1))
       
  'Set cbBuffer equal to the size of the buffer
   cbBuffer = cbRequired
    
  'Enumerate the printers. If the function succeeds,
  'the return value is nonzero. If the function fails,
  'the return value is zero.
   If EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, _
                   vbNullString, PRINTER_LEVEL1, _
                   pntr(0), cbBuffer, _
                   cbRequired, nEntries) Then
          
      For c = 0 To nEntries - 1
           
         With pntr(c)

            sFlags = ""
               
            If GetStrFromPtrA(.Pane) = printerName Then
              'see Comments for info on these flags
               If (.flags And PRINTER_ENUM_CONTAINER) Then sFlags = "enumerable "
               If (.flags And PRINTER_ENUM_EXPAND) Then sFlags = sFlags & "expand "
               If (.flags And PRINTER_ENUM_ICON1) Then sFlags = sFlags & "icon1 "
               If (.flags And PRINTER_ENUM_ICON2) Then sFlags = sFlags & "icon2 "
               If (.flags And PRINTER_ENUM_ICON3) Then sFlags = sFlags & "icon3 "
               If (.flags And PRINTER_ENUM_ICON8) Then sFlags = sFlags & "icon8 "
               
               PrintersWin9x = GetStrFromPtrA(.Pane) & ";" & _
                            sFlags & ";" & _
                            GetStrFromPtrA(.prescription)
                             
            
               Exit For
            End If
            
         End With
      Next c
           
   Else: 'ctl.AddItem "Error enumerating printers."
   End If  'EnumPrinters
    
    
End Function

Public Sub SetControlFocus(ByVal objForm As Form, ByVal Ctrl As Control)
Dim cnt As Control
Dim TabIndex As Integer
Dim strNamesList As String

   ' intenta hacer un SetFocus en Ctrl. Si no es posible prueba con el siguiente
   ' y asi sucesivamente
   
   strNamesList = "TextBox PowerMask DataCombo ComboBox OptionButton ListView CommandButton SSTab"
   
   TabIndex = Ctrl.TabIndex
   
   On Error Resume Next
   
   Do While True
      If Ctrl.Enabled Then Ctrl.SetFocus: Exit Do
      
      If TabIndex > 300 Then Exit Do
      
      For Each cnt In objForm.Controls
         If cnt.TabIndex = TabIndex + 1 Then
            If InStr(strNamesList, TypeName(cnt)) > 0 Then
               Set Ctrl = cnt
               'TabIndex = Ctrl.TabIndex
               Exit For
            Else
               TabIndex = TabIndex + 1
            End If
         End If
      Next cnt
      
      TabIndex = TabIndex + 1
   Loop
   
End Sub

Public Function ReadTextFileContents(filename As String) As String
Dim fnum As Integer, IsOpen As Boolean

   On Error GoTo GestErr
   
   fnum = FreeFile()
   Open filename For Input As #fnum
   
   IsOpen = True
   ReadTextFileContents = Input(LOF(fnum), fnum)

GestErr:
   If IsOpen Then Close #fnum
   If Err Then MsgBox Err.Description, vbOKOnly, filename
End Function

Public Function xRound(ByVal nNum, Optional ByVal nDec As Integer = 0) As Variant
Dim n As Variant
Dim V1 As Variant
   
   If nNum = "" Then nNum = 0
   
   If nNum < 0 Then
      V1 = nNum
      nNum = Abs(nNum)
   End If
   
   n = CDec(nNum * 10 ^ nDec)
   n = n + IIf(nNum < 0, -0.5, 0.5)
   n = Int(n)
   n = n / 10 ^ nDec
   
   If V1 < 0 Then n = n * -1
   
   xRound = n
End Function
Public Function ConvertRstToCSV(ByVal rst As ADODB.Recordset, ByVal strDelimiter As String) As String
Dim fld  As ADODB.Field
'Dim ix   As Integer
Dim bUpdate As Boolean
Dim rst1    As ADODB.Recordset

   Set rst1 = CopyData(rst, AllRecords)
   
   If rst1.RecordCount > 0 Then
      rst1.MoveFirst
      
      Do While Not rst1.EOF
      
         bUpdate = False
         For Each fld In rst1.Fields
         
            If fld.Type = adDBTimeStamp And Not IsNull(rst1(fld.Name).Value) Then
               rst1(fld.Name).Value = DateValue(rst1(fld.Name).Value)
               bUpdate = True
            End If
            
         Next fld
         
         If bUpdate Then rst1.Update
         
         rst1.MoveNext
      Loop

      rst1.MoveFirst
      
      ConvertRstToCSV = rst1.GetString(adClipString, -1, strDelimiter, vbCrLf, "<NULL>")
   End If
   
End Function

Public Function GetUltimoDiaMes(ByVal mes As Integer, ByVal anio As Integer) As Date
   ' Creada por P.M.
   ' Esta funcion devuelve el ultimo día del mes y anio pasados como parametros
   Dim MesProx, AnioProx As Integer
   MesProx = IIf(mes = 12, 1, mes + 1)
   AnioProx = IIf(mes = 12, anio + 1, anio)
   GetUltimoDiaMes = CDate("01/" + Str(MesProx) + "/" + Str(AnioProx))
   GetUltimoDiaMes = DateAdd("d", -1, GetUltimoDiaMes)
End Function

'---------------------------------------------------------------------------------------
' Procedure : RestoreDefaultPrinter
' DateTime  : 23/02/2005 18:02
' Author    : tony
' Purpose   : setea StrDefaultDeviceName como default
'---------------------------------------------------------------------------------------
'
Public Sub RestoreDefaultPrinter(ByVal StrDefaultDeviceName As String)
Dim prt As Printer

   If StrDefaultDeviceName = NullString Then Exit Sub
   
   For Each prt In Printers
      If prt.DeviceName = StrDefaultDeviceName Then
         Set Printer = prt
         Exit Sub
      End If
   Next prt

End Sub
'---------------------------------------------------------------------------------------
' Procedure : GetNombreMaquina
' DateTime  : 12/01/2006 16:55
' Author    : fernando
' Purpose   : Buscar nombre PC segun el ultimo acceso a Auditoria
'---------------------------------------------------------------------------------------
Public Function GetNombreMaquina(ByVal strEmpresa As String, ByVal strUsuario As String) As String
      Dim rst1 As ADODB.Recordset
      Dim SQL  As String
   
10       On Error GoTo GestErr
   
20       SQL = " SELECT "
30       SQL = SQL & "        TRD_MAQUINA, "
40       SQL = SQL & "        TRD_COMENTARIO "
50       SQL = SQL & "   FROM TRANSACCIONES_AUDITORIA, "
60       SQL = SQL & "        (SELECT MAX (TRD_NUMERO_TRANSACCION) AS NUMERO "
70       SQL = SQL & "           FROM TRANSACCIONES_AUDITORIA "
80       SQL = SQL & "          WHERE TRD_USUARIO = '" & strUsuario & "') ULTIMO "
90       SQL = SQL & "  WHERE TRD_NUMERO_TRANSACCION = ULTIMO.NUMERO "
100      Set rst1 = Fetch(strEmpresa, SQL)

110      If rst1.RecordCount > 0 Then
120         If IIf(IsNull(rst1("TRD_COMENTARIO").Value), Space(1), rst1("TRD_COMENTARIO").Value) <> "Logout" Then
130            GetNombreMaquina = rst1("TRD_MAQUINA").Value
140         End If
150      End If
   
160      Exit Function

GestErr:
170      LoadError ErrorLog, "GetNombreMaquina" & Erl
180      ShowErrMsg ErrorLog

End Function

'TP 5525 INC 80382 se agrega el parametro strEmpresa
Public Function GetSqlAplicaciones(ByVal strTipoContrato As String, Optional ByVal strImprimeCuentaYOrden As String = No, Optional ByVal strEmpresa As String = NullString) As String
Dim strSQL As String

   
   
   strSQL = " SELECT   APLICACIONES.APC_TIPO_CONTRATO, APLICACIONES.APC_NUMERO_CONTRATO, APLICACIONES.APC_SUCURSAL_CARTA_PORTE,"
   strSQL = strSQL & "          APLICACIONES.APC_NUMERO_CARTA_PORTE, APLICACIONES.APC_VERIFICADO, "
   strSQL = strSQL & "             TO_CHAR (APLICACIONES.APC_SUCURSAL_CARTA_PORTE, '0000')"
   strSQL = strSQL & "          || TO_CHAR (APLICACIONES.APC_NUMERO_CARTA_PORTE, '0000000000') AS CARTA_PORTE," 'TP 2770 INC 64637
   strSQL = strSQL & "          APLICACIONES.APC_RENGLON, CONTRATOS.CNT_ESPECIE, CONTRATOS.CNT_COSECHA, ESPECIES_COSECHAS.ECO_DESCRIPCION,"
   strSQL = strSQL & "          CONTRATOS.CNT_CORREDOR, CORREDORES.ENT_NOMBRE AS NOMBRE_CORREDOR, CONTRATOS.CNT_CONTRATO_CORREDOR,"

  If strTipoContrato = "V" Then
      strSQL = strSQL & "          CONTRATOS.CNT_COMPRADOR, COMPRADORES.ENT_NOMBRE AS NOMBRE_COMPRADOR, CONTRATOS.CNT_CONTRATO_COMPRADOR,"
   Else
      strSQL = strSQL & "          CONTRATOS.CNT_VENDEDOR, VENDEDORES.ENT_NOMBRE AS NOMBRE_VENDEDOR, CONTRATOS.CNT_CONTRATO_VENDEDOR,"
   End If

   strSQL = strSQL & "          CONTRATOS.CNT_KILOS_MINIMOS, CONTRATOS.CNT_KILOS_MAXIMOS, CONTRATOS.CNT_FECHA_DESDE_ENTREGAS,"
   strSQL = strSQL & "          CONTRATOS.CNT_FECHA_HASTA_ENTREGAS, CARTAS_DE_PORTE.CPE_FECHA_ENVIO, CARTAS_DE_PORTE.CPE_FECHA_DESCARGA,"
   strSQL = strSQL & "          APLICACIONES.APC_CONTRATO_ORIGINAL, CARTAS_DE_PORTE.CPE_KILOS_BRUTOS, CARTAS_DE_PORTE.CPE_TARA,"
   strSQL = strSQL & "          ROUND(NVL(((APLICACIONES.APC_KILOS_NETOS * CARTAS_DE_PORTE.CPE_KILOS_ENVIADOS) / CARTAS_DE_PORTE.CPE_KILOS_NETOS), CARTAS_DE_PORTE.CPE_KILOS_ENVIADOS), 0) AS ENVIADOS_PROPORCIONALES,"
   strSQL = strSQL & "          APLICACIONES.APC_KILOS_DESCARGADOS,"
   strSQL = strSQL & "          ABS (NVL (CARTAS_DE_PORTE.CPE_KILOS_ENVIADOS, 0) - NVL (APLICACIONES.APC_KILOS_DESCARGADOS, 0)"
   strSQL = strSQL & "              ) AS DIFERENCIA_KILOS,"
   strSQL = strSQL & "          APLICACIONES.APC_PORCENTAJE_HUMEDAD, APLICACIONES.APC_MERMA_HUMEDAD, APLICACIONES.APC_MERMA_ZARANDEO,"
   strSQL = strSQL & "          APLICACIONES.APC_MERMA_VOLATIL, APLICACIONES.APC_KILOS_MERMA_HUMEDAD, APLICACIONES.APC_KILOS_MERMA_ZARANDEO,"
   strSQL = strSQL & "          APLICACIONES.APC_KILOS_MERMA_VOLATIL, APLICACIONES.APC_KILOS_NETOS, APLICACIONES.APC_KILOS_SERVICIOS,"
   strSQL = strSQL & "          CARTAS_DE_PORTE.CPE_SUCURSAL_TICKET, CARTAS_DE_PORTE.CPE_NUMERO_TICKET,"
   strSQL = strSQL & "             TO_CHAR (CARTAS_DE_PORTE.CPE_SUCURSAL_TICKET, '0000')"
   strSQL = strSQL & "          || TO_CHAR (CARTAS_DE_PORTE.CPE_NUMERO_TICKET, '00000000') AS TICKET,"
   strSQL = strSQL & "          CARTAS_DE_PORTE.CPE_LIQUIDA_VIAJE, CARTAS_DE_PORTE.CPE_TRANSPORTISTA,"
   strSQL = strSQL & "          TRANSPORTISTAS.ENT_NOMBRE AS NOMBRE_TRANSPORTISTA, CARTAS_DE_PORTE.CPE_EMPRESA_TRANSPORTE,"
   strSQL = strSQL & "          EMPRESAS_TRANSPORTE.ENT_NOMBRE AS NOMBRE_EMPRESA_TRANSPORTE, CARTAS_DE_PORTE.CPE_KILOMETROS,"
   strSQL = strSQL & "          APLICACIONES.APC_FLETE_FIJO, CARTAS_DE_PORTE.CPE_CODIGO_TARIFA_FLETE, CARTAS_DE_PORTE.CPE_TARIFA_FLETE,"
   strSQL = strSQL & "          CARTAS_DE_PORTE.CPE_PLANTA, CARTAS_DE_PORTE.CPE_PROCEDENCIA, CARTAS_DE_PORTE.CPE_DESTINO,"
   strSQL = strSQL & "          CARTAS_DE_PORTE.CPE_OBSERVACIONES,"
   strSQL = strSQL & "          DECODE (CARTAS_DE_PORTE.CPE_VERIFICADO, 'NI', 'Sí', CARTAS_DE_PORTE.CPE_VERIFICADO),"
   strSQL = strSQL & "          CARTAS_DE_PORTE.CPE_VIAJE_ENGANCHADO, CARTAS_DE_PORTE.CPE_PUNTO_INGRESO, CARTAS_DE_PORTE.CPE_NUMERO_INGRESO,"
   strSQL = strSQL & "             TO_CHAR (CARTAS_DE_PORTE.CPE_PUNTO_INGRESO, '0000')"
   strSQL = strSQL & "          || TO_CHAR (CARTAS_DE_PORTE.CPE_NUMERO_INGRESO, '00000000') AS NUMERO_INGRESO,"
   strSQL = strSQL & "          CARTAS_DE_PORTE.CPE_PATENTE_CAMION, CARTAS_DE_PORTE.CPE_PATENTE_ACOPLADO"
   
   'Desde aqui en adelante estos datos salen solo en exportacion excel
   If strImprimeCuentaYOrden = si Then
      strSQL = strSQL & "          , EGRESOS_PLANTAS.EPL_CUENTA_ORDEN_1, CYO1.ENT_NOMBRE AS CYO1_NOMBRE, EGRESOS_PLANTAS.EPL_CUENTA_ORDEN_2,"
      strSQL = strSQL & "          CYO2.ENT_NOMBRE AS CYO2_NOMBRE, EGRESOS_PLANTAS.EPL_CUENTA_ORDEN_3, CYO3.ENT_NOMBRE AS CYO3_NOMBRE"
   End If
      
   strSQL = strSQL & "        ,ENTIDADES.ENT_REGION as REGION_CUENTA "
   strSQL = strSQL & "        ,SUCURSALES.SUC_REGION as REGION_OPERACION"
   
   'TP 11100 INC 118822
   strSQL = strSQL & "        ,APLICACIONES.APC_CONCILIADA as CONCILIADA"
   
   strSQL = strSQL & " FROM     APLICACIONES,"
   strSQL = strSQL & "          CONTRATOS,"
   strSQL = strSQL & "          CARTAS_DE_PORTE,"
   strSQL = strSQL & "          ESPECIES_COSECHAS,"
   strSQL = strSQL & "          ENTIDADES COMPRADORES,"
   strSQL = strSQL & "          ENTIDADES CORREDORES,"
   strSQL = strSQL & "          ENTIDADES VENDEDORES,"
   strSQL = strSQL & "          ENTIDADES TRANSPORTISTAS,"
   strSQL = strSQL & "          ENTIDADES EMPRESAS_TRANSPORTE,"
   strSQL = strSQL & "          ANALISIS_CP_AFECTADAS,"
   strSQL = strSQL & "          MONEDAS"

   If strImprimeCuentaYOrden = si Then
      strSQL = strSQL & "          , EGRESOS_PLANTAS,"
      strSQL = strSQL & "          ENTIDADES CYO1,"
      strSQL = strSQL & "          ENTIDADES CYO2,"
      strSQL = strSQL & "          ENTIDADES CYO3"
   End If
   
   'TP 5525 INC 80382
   strSQL = strSQL & "          , PLANTAS , SUCURSALES, ENTIDADES "
   
   
   strSQL = strSQL & " WHERE    CARTAS_DE_PORTE.CPE_TIPO_MOVIMIENTO = APLICACIONES.APC_TIPO_CONTRATO"
   strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_SUCURSAL_CARTA_PORTE = APLICACIONES.APC_SUCURSAL_CARTA_PORTE"
   strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_NUMERO_CARTA_PORTE = APLICACIONES.APC_NUMERO_CARTA_PORTE"
   strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_RENGLON = APLICACIONES.APC_RENGLON"
   strSQL = strSQL & "      AND CONTRATOS.CNT_MONEDA = MONEDAS.MON_MONEDA(+)"
   strSQL = strSQL & "      AND CONTRATOS.CNT_TIPO_CONTRATO = APLICACIONES.APC_TIPO_CONTRATO"
   strSQL = strSQL & "      AND APLICACIONES.APC_TIPO_CONTRATO = '" & strTipoContrato & "'"
   strSQL = strSQL & "      AND APLICACIONES.APC_NUMERO_CONTRATO = CONTRATOS.CNT_NUMERO"
   strSQL = strSQL & "      AND COMPRADORES.ENT_TIPO_ENTIDAD(+) = 3"
   strSQL = strSQL & "      AND CONTRATOS.CNT_COMPRADOR = COMPRADORES.ENT_CODIGO(+)"
   strSQL = strSQL & "      AND VENDEDORES.ENT_TIPO_ENTIDAD(+) = 3"
   strSQL = strSQL & "      AND CONTRATOS.CNT_VENDEDOR = VENDEDORES.ENT_CODIGO(+)"
   strSQL = strSQL & "      AND CORREDORES.ENT_TIPO_ENTIDAD(+) = 3"
   strSQL = strSQL & "      AND CONTRATOS.CNT_CORREDOR = CORREDORES.ENT_CODIGO(+)"
   strSQL = strSQL & "      AND CONTRATOS.CNT_ESPECIE = ESPECIES_COSECHAS.ECO_ESPECIE"
   strSQL = strSQL & "      AND CONTRATOS.CNT_COSECHA = ESPECIES_COSECHAS.ECO_COSECHA"
   strSQL = strSQL & "      AND TRANSPORTISTAS.ENT_TIPO_ENTIDAD(+) = 3"
   strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_TRANSPORTISTA = TRANSPORTISTAS.ENT_CODIGO(+)"
   strSQL = strSQL & "      AND EMPRESAS_TRANSPORTE.ENT_TIPO_ENTIDAD(+) = 3"
   strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_EMPRESA_TRANSPORTE = EMPRESAS_TRANSPORTE.ENT_CODIGO(+)"
   strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_TIPO_MOVIMIENTO = ANALISIS_CP_AFECTADAS.ACF_TIPO_MOVIMIENTO(+)"
   strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_SUCURSAL_CARTA_PORTE = ANALISIS_CP_AFECTADAS.ACF_SUCURSAL_CARTA_PORTE(+)"
   strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_NUMERO_CARTA_PORTE = ANALISIS_CP_AFECTADAS.ACF_NUMERO_CARTA_PORTE(+)"
   strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_RENGLON = ANALISIS_CP_AFECTADAS.ACF_RENGLON(+)"
   'strSQL = strSQL & "      AND APLICACIONES.APC_VERIFICADO = '" & No & "' "
   
   'TP 5525 INC 80382
   strSQL = strSQL & "      AND NVL(CARTAS_DE_PORTE.CPE_PLANTA,0) = PLANTAS.PLA_PLANTA (+) "
   strSQL = strSQL & "      AND NVL(PLANTAS.PLA_SUCURSAL_OPERA,' ') = SUCURSALES.SUC_CODIGO (+)  "
   'TP 5525 INC 80382 por filtro
   If strTipoContrato = "V" Then
      strSQL = strSQL & "      AND NVL(COMPRADORES.ENT_CODIGO,' ') = ENTIDADES.ENT_CODIGO (+) "
      strSQL = strSQL & "      AND 3 =   ENTIDADES.ENT_TIPO_ENTIDAD (+)"
   Else
      strSQL = strSQL & "      AND NVL(VENDEDORES.ENT_CODIGO,' ') = ENTIDADES.ENT_CODIGO (+) "
      strSQL = strSQL & "      AND 3 =   ENTIDADES.ENT_TIPO_ENTIDAD (+) "
   End If
   
   
   If strImprimeCuentaYOrden = si Then
      strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_SUCURSAL_CARTA_PORTE = EGRESOS_PLANTAS.EPL_SUCURSAL_COMP_SALIDA(+)"
      strSQL = strSQL & "      AND CARTAS_DE_PORTE.CPE_NUMERO_CARTA_PORTE = EGRESOS_PLANTAS.EPL_NUMERO_COMP_SALIDA(+)"
      strSQL = strSQL & "      AND CYO1.ENT_TIPO_ENTIDAD(+) = 3"
      strSQL = strSQL & "      AND EGRESOS_PLANTAS.EPL_CUENTA_ORDEN_1 = CYO1.ENT_CODIGO(+)"
      strSQL = strSQL & "      AND CYO2.ENT_TIPO_ENTIDAD(+) = 3"
      strSQL = strSQL & "      AND EGRESOS_PLANTAS.EPL_CUENTA_ORDEN_2 = CYO2.ENT_CODIGO(+)"
      strSQL = strSQL & "      AND CYO3.ENT_TIPO_ENTIDAD(+) = 3"
      strSQL = strSQL & "      AND EGRESOS_PLANTAS.EPL_CUENTA_ORDEN_3 = CYO3.ENT_CODIGO(+)"
   End If


   strSQL = strSQL & " ORDER BY APLICACIONES.APC_NUMERO_CONTRATO,"
   strSQL = strSQL & "          APLICACIONES.APC_SUCURSAL_CARTA_PORTE,"
   strSQL = strSQL & "          APLICACIONES.APC_NUMERO_CARTA_PORTE,"
   strSQL = strSQL & "          APLICACIONES.APC_RENGLON"
   
   GetSqlAplicaciones = strSQL
   
End Function

Public Function GetMyObject(ByVal strComponentClass As String, Optional ByVal strServerName As String = NullString) As Object
10       On Error GoTo GestErr

         ' Sin este Objeto Local (que termina en Nothing) se queda vivo el objeto en el servidor
         Dim objetoLocal   As Object
         Dim ix As Integer
   
20       ix = 0
30       Set objetoLocal = CreateObject(strComponentClass, strServerName)
40       Set GetMyObject = objetoLocal
   
50       Set objetoLocal = Nothing
   
60       Exit Function

GestErr:
70       ix = ix + 1
80       If ix < 3 Then
90          Resume
100      End If
   
110      Set objetoLocal = Nothing
120      Set GetMyObject = Nothing

130      LoadError ErrorLog, "Objeto: " & strComponentClass & vbCrLf & "Servidor: " & strServerName
140      ShowErrMsg ErrorLog
End Function

Public Sub SetEjercicioPanel(ByVal frm As Object, ByVal strEjercicioCorriente As String, Optional iPanelWidth As Integer)
Dim pan As Panel

   On Error Resume Next
   
   Set pan = frm.stb1.Panels("EJERCICIO")
   
   If pan Is Nothing Then
      Set pan = frm.stb1.Panels.Add
      pan.Key = "EJERCICIO"
   End If
   
   pan.Text = strEjercicioCorriente
   pan.ToolTipText = "Haga Doble Click para cambiar el Ejercicio"
   If iPanelWidth = 0 Then
      pan.Width = 1200
   Else
      pan.Width = iPanelWidth
   End If
   pan.AutoSize = sbrNoAutoSize

End Sub

Public Sub CambiarEjercicio(ByVal mvarForm As Object, ByRef strEjercicioCorriente As String, ByVal strNuevoEjercicio As String)
Dim mvarControlData As udtControlData

   'se intentara reiniciar el form forzando una descarga de un form MRU
   'si en el tentativo de cambiar de ejercicio, el usuario decide
   '  Abortar:      no habo nada
   '  vbNo, vbYes:  cambio el ejercicio
   
   On Error Resume Next
   
   If strEjercicioCorriente = strNuevoEjercicio Then Exit Sub
   
   'agrego el form en la coleccion MRU para evitar la descarga durante el Unload
   MRUForms.Add Item:=mvarForm, Key:=mvarForm.MenuKey
   
   Unload mvarForm
   
   Select Case mvarForm.UnloadState
      Case vbCancel
         'el usuario aborto -> no hago nada
         Exit Sub
      Case vbNo, vbYes
         'el usuario decidio seguir adelante -> cambio el ejercicio
         strEjercicioCorriente = strNuevoEjercicio
      Case Else
         'el usuario no fue interrogado -> cambio el ejercicio
         strEjercicioCorriente = strNuevoEjercicio
   End Select
   
   
'   SaveSetting appname:="Algoritmo", Section:=mvarForm.Name, _
'            Key:="NuevoEjercicio", setting:=strNuevoEjercicio

   
   'agrego el form en la TaskBar
   If Not mvarMDIForm.MDITaskBar1.FormInTaskBar(mvarForm.hWnd) Then
      mvarMDIForm.MDITaskBar1.AddFormToTaskBar mvarForm
      CenterMDIActiveXChild mvarForm
   End If
   
   
   mvarControlData = mvarForm.ControlData
   mvarControlData.Ejercicio = strNuevoEjercicio
   mvarForm.ControlData = mvarControlData
   
   mvarForm.InitForm
   
   mvarForm.PostInitForm
   
End Sub
Public Function ExistField(ByVal strTabla As String, ByVal strCampo As String, ByVal strEmpresa As String) As Boolean
         Dim objDataAccess As DataAccess.clsDataFuncs
      
'TP 5525 INC 80382
      
10       On Error GoTo GestErr
   
20       Set objDataAccess = GetMyObject("DataAccess.clsDataFuncs")

30       ExistField = (objDataAccess.SerializedFetch(strEmpresa, strTabla & ";" & strCampo) = si)
   
40       Set objDataAccess = Nothing

50       Exit Function
GestErr:
60       LoadError ErrorLog, "ExistField" & Erl
   
70       Set objDataAccess = Nothing
   
80       ShowErrMsg ErrorLog
  
End Function

'BUG 9477. SAMSA - Error en orden de columnas
'Se estaba empesando a copiar por todos lados y no estaba completa, ahora esta un poquito mejor... no es infalible.
Public Sub AsignarOrdenDeColumnas(ByRef frm As Form, _
                                  ByRef lvw As ListView, _
                                  ByVal strListOrdenColumnas As String, _
                                  ByVal CONST_COLUMNAS_INTERCAMBIABLES As String, _
                                  ByVal strUpdateTableOrdenColumnas As String, _
                                  Optional ByVal strListAnchoColumnas As String = NullString, _
                                  Optional ByVal CONST_ANCHO_COLUMNAS_INTERCAMBIABLES As String = NullString, _
                                  Optional ByVal strUpdateTableAnchoColumnas As String = NullString)
                                 
         Dim aCol()        As String         'Arreglo con el orden de las columnas del registro de sistemas
         Dim aAnchosCol()  As String         'Arreglo con el ancho de las columnas del registro de sistemas
         Dim ColHeader     As ColumnHeader
         Dim aHeader()     As String         'Arreglo q contiene las propiedades text, width, alignment y tag del ListView
         Dim ic            As Integer
         'Variables para corregir el registro cuabndo se agregan columnas
         Dim objTabla      As BOGeneral.clsTablas
         Dim aConstante()  As String
         Dim aDefinidas()  As String
         Dim iss           As Integer
   
         Dim bEstablecerAnchos   As Boolean
         Dim iCountError         As Integer
         Dim mvarControlData     As udtControlData
   
10       On Error GoTo GestErr
    
20       mvarControlData = frm.ControlData
   
30       bEstablecerAnchos = (Len(strListAnchoColumnas) > 0 And Len(CONST_ANCHO_COLUMNAS_INTERCAMBIABLES) > 0)
   
         '----------------------------------------------------------------------------------------------------------------
         '                   Controlo cantidad de columnas definidas por registro contra la constante
         '----------------------------------------------------------------------------------------------------------------
40       If Len(CONST_COLUMNAS_INTERCAMBIABLES) <> Len(strListOrdenColumnas) Then
            '1. Si sobran o faltan caracteres, compruebo si es porque fatan columnas en el registro.
50          aConstante = Split(CONST_COLUMNAS_INTERCAMBIABLES, ",")
60          aDefinidas = Split(strListOrdenColumnas, ",")

70          If UBound(aDefinidas) <> UBound(aConstante) Then
               '2.1 Si la cantidad de columnas del registro de sistema es diferente a la constante intento agregarlas al final y ver si queda solucionado.
80             For iss = UBound(aDefinidas) To UBound(aConstante) - 1
90                strListOrdenColumnas = strListOrdenColumnas & "," & iss + 2
100            Next
110            If Len(CONST_COLUMNAS_INTERCAMBIABLES) <> Len(strListOrdenColumnas) Then
                  '2.1.1 El error continua, establesco orden por defecto.
120               strListOrdenColumnas = CONST_COLUMNAS_INTERCAMBIABLES
130            Else
                  '2.1.2 El error se soluciono, grabo el registro de sistema con las nuevas columnas al final
140               If Len(strUpdateTableOrdenColumnas) > 0 Then
150                  Set objTabla = New BOGeneral.clsTablas
160                  objTabla.ControlData = mvarControlData
170                  objTabla.UpdateKey mvarControlData.Empresa, strUpdateTableOrdenColumnas, strListOrdenColumnas
180               End If
190            End If
200         Else
               '2.2 No se cual puede ser el origen del error, establesco el orden por defecto.
210            strListOrdenColumnas = CONST_COLUMNAS_INTERCAMBIABLES
220         End If

230      End If
   
240      If bEstablecerAnchos Then
            '----------------------------------------------------------------------------------------------------------------
            '                    Controlo cantidad de anchos definidos pr sistema contra la constante
            '----------------------------------------------------------------------------------------------------------------
250         aConstante = Split(CONST_ANCHO_COLUMNAS_INTERCAMBIABLES, ",")
260         aDefinidas = Split(strListAnchoColumnas, ",")
270         If UBound(aDefinidas) <> UBound(aConstante) Then
               'Intento de agregar las columnas que faltan (esto supone que las columnas se agregan siempre al final del list view)
280            For iss = UBound(aDefinidas) To UBound(aConstante) - 1
290               strListAnchoColumnas = strListAnchoColumnas & "," & aConstante(iss + 1)
300            Next
310            aDefinidas = Split(strListAnchoColumnas, ",")
320            If UBound(aDefinidas) <> UBound(aConstante) Then
                  'el error persiste, pongo los anchos por defecto
330               strListAnchoColumnas = CONST_ANCHO_COLUMNAS_INTERCAMBIABLES
340            Else
                  'Se soluciono, grabar registro de sisteams con todas las columnas definidas.
350               If Len(strUpdateTableAnchoColumnas) > 0 Then
360                  If objTabla Is Nothing Then
370                     Set objTabla = New BOGeneral.clsTablas
380                     objTabla.ControlData = mvarControlData
390                  End If
400                  objTabla.UpdateKey mvarControlData.Empresa, strUpdateTableAnchoColumnas, strListAnchoColumnas
410               End If
420            End If
430         End If
440      End If
   
PorDefecto:

450      aCol = Split(strListOrdenColumnas, ",")
   
460      If bEstablecerAnchos Then aAnchosCol = Split(strListAnchoColumnas, ",")
   
470      If CONST_COLUMNAS_INTERCAMBIABLES <> strListOrdenColumnas Then
            'si el orden por defecto es distinto al del registro

480         ReDim aHeader(UBound(aCol), 3) As String

            'se graban las propiedades text, width, alignment y tag de las columnas del listView en el array aHeader()
490         For Each ColHeader In lvw.ColumnHeaders
500            If ColHeader.Index <> 1 Then
510               aHeader(ColHeader.Index - 2, 0) = ColHeader.Text
520               aHeader(ColHeader.Index - 2, 1) = CStr(ColHeader.Width)
530               aHeader(ColHeader.Index - 2, 2) = CStr(ColHeader.Alignment)
540               aHeader(ColHeader.Index - 2, 3) = CStr(ColHeader.Tag)
550            End If
560            If (ColHeader.Index - 2) = UBound(aHeader) Then Exit For
570         Next

            'se recorre el arreglo del registro de sistemas para asignar la nueva ubicacion y ancho de la columna
580         For ic = LBound(aCol) To UBound(aCol)
590            lvw.ColumnHeaders.Item(ic + 2).Text = aHeader(CInt(aCol(ic)) - 1, 0)
600            If bEstablecerAnchos Then
610               lvw.ColumnHeaders.Item(ic + 2).Width = CLng(aAnchosCol(ic))
620            Else
630               lvw.ColumnHeaders.Item(ic + 2).Width = CLng(aHeader(CInt(aCol(ic)) - 1, 1))
640            End If
650            lvw.ColumnHeaders.Item(ic + 2).Alignment = CInt(aHeader(CInt(aCol(ic)) - 1, 2))
660            lvw.ColumnHeaders.Item(ic + 2).Tag = aHeader(CInt(aCol(ic)) - 1, 3)
         
670            frm.AsignarVariablesPorLvw lvw.Name, aCol(ic), ic
680         Next ic
690      Else
            'sigue como antes
700         For ic = LBound(aCol) To UBound(aCol)
710            If bEstablecerAnchos Then lvw.ColumnHeaders.Item(ic + 2).Width = CLng(aAnchosCol(ic))
720            frm.AsignarVariablesPorLvw lvw.Name, aCol(ic), ic
730         Next ic
740      End If
   
750      Set objTabla = Nothing
   
760      Exit Sub
   
GestErr:
         ' si se define en el registro un index inexistente uso la configuracion por defecto
770      If Err.Number = 9 And InStr(UCase(Err.Description), UCase("fuera del intervalo")) > 0 And iCountError = 0 Then
780         strListOrdenColumnas = CONST_COLUMNAS_INTERCAMBIABLES
790         strListAnchoColumnas = CONST_ANCHO_COLUMNAS_INTERCAMBIABLES
      
800         iCountError = iCountError + 1
      
810         GoTo PorDefecto
820      Else
830         LoadError ErrorLog, "AsignarOrdenDeColumnas" & Erl
840         Set objTabla = Nothing
850         ShowErrMsg ErrorLog
860      End If
End Sub
Public Function ArmarPrevision(ByVal mvarControlData As Variant, ByVal pbParametros As PropertyBag) As PropertyBag
   Dim objGestionComercial       As Object 'puedo hacer referencia a aplications porque esta en todos los proyectos (es shared). si llegara a generar error en un futuro , paras a tipo object
   Dim hWndAdmin                 As Long
   Dim objGesComForms            As Object
   Dim ix                        As Integer
'   Dim frmComprobanteTesoreria   As frmEmisionCompTesoreria  ' Form
'   Dim objHistoricoTesoreria     As BOGesCom.clsHistoricoTesoreria 'CAMBIAR ANTES DE PROTEGER POR OBJECT
   
   Dim frmComprobanteTesoreria   As Form  ' Form
   Dim objHistoricoTesoreria     As Object  'CAMBIAR ANTES DE PROTEGER POR OBJECT
   
   
   On Error GoTo GestErr
   
   Set ArmarPrevision = New PropertyBag
   
   
   '------------------------------------------------------------------------------------------
   '  Armado de objeto
   '------------------------------------------------------------------------------------------
   Set objHistoricoTesoreria = CreateObject("BOGesCom.clsHistoricoTesoreria")
   objHistoricoTesoreria.ControlData = mvarControlData
   
   objHistoricoTesoreria.InitProperties
   
   With objHistoricoTesoreria
      'Cabecera:
      '.ComprobanteTesoreria = pbParametros.ReadProperty("COMPROBANTE_TESORERIA", NullString)
      '.FechaTesoreria = pbParametros.ReadProperty("FECHA", 0)
      '.Moneda = pbParametros.ReadProperty("MONEDA_CTA_CTE", NullString)
      '.TipoCambioEmision = pbParametros.ReadProperty("TIPO_CAMBIO", 0)
      '.FechaPago = pbParametros.ReadProperty("FECHA_PAGO", 0)
      '.TipoEntidad = pbParametros.ReadProperty("TIPO_ENTIDAD", 0)
      '.Entidad = pbParametros.ReadProperty("CODIGO_ENTIDAD", NullString)
      .Comentarios = pbParametros.ReadProperty("COMENTARIO", NullString)
      
   End With
   
   
   '------------------------------------------------------------------------------------------
   '  Obtener formulario
   '------------------------------------------------------------------------------------------
   Set objGestionComercial = mvarMDIForm.GetInstance("GestionComercial")
   
   hWndAdmin = mvarMDIForm.CallAdmin("EMISION_COMPROBANTES_TESORERIA", mvarControlData, True)
   
   DoEvents
         
   Set objGesComForms = objGestionComercial.CollectionForms
   
   DoEvents
         
   For ix = 0 To objGesComForms.Count - 1
       If objGesComForms(ix).hWnd = hWndAdmin Then
         Set frmComprobanteTesoreria = objGesComForms(ix)
         Exit For
      End If
   Next
   
   
   '------------------------------------------------------------------------------------------
   '  Completar objeto
   '------------------------------------------------------------------------------------------
   '...
   '..
   frmComprobanteTesoreria.CompletarComprobante pbParametros.ReadProperty("COMPROBANTE_TESORERIA", NullString), _
                                                pbParametros.ReadProperty("FECHA", 0), _
                                                pbParametros.ReadProperty("TIPO_ENTIDAD", 0), _
                                                pbParametros.ReadProperty("CODIGO_ENTIDAD", NullString), _
                                                , , , , , _
                                                pbParametros.ReadProperty("FECHA_PAGO", 0), _
                                                pbParametros.ReadProperty("IMPORTE", 0), _
                                                objHistoricoTesoreria, _
                                                , , , , _
                                                pbParametros.ReadProperty("MONEDA_CABECERA", NullString), _
                                                pbParametros.ReadProperty("TIPO_CAMBIO", 0), _
                                                True, _
                                                True, _
                                                pbParametros.ReadProperty("MONEDA_CTA_CTE", NullString), _
                                                pbParametros.ReadProperty("TIPO_CTA_CTE", NullString)
                                                

   Set objHistoricoTesoreria = frmComprobanteTesoreria.ComprobanteTesoreriaCompleto
   
   '-------------------------------------------------
   '  Armar valores de salida:
   '-------------------------------------------------
   ArmarPrevision.WriteProperty "COMPROBANTE_TESORERIA_COMPLETO", objHistoricoTesoreria.GetState(True)
   
   
   
   'Nothings:
   Set objHistoricoTesoreria = Nothing
   
   If Not frmComprobanteTesoreria Is Nothing Then Unload frmComprobanteTesoreria
   Set frmComprobanteTesoreria = Nothing
   
   Exit Function
GestErr:

   LoadError ErrorLog, "ArmarPrevision" & Erl
   
   Set objHistoricoTesoreria = Nothing
   If Not frmComprobanteTesoreria Is Nothing Then Unload frmComprobanteTesoreria
   Set frmComprobanteTesoreria = Nothing
   
   ShowErrMsg ErrorLog
End Function

'dejo comentada esta sub de prueba de emision que me habia olvidado de comentar en la protección anterior.
'''Public Sub TestArmarPrevision(m As Variant)
'''   Dim pbPar As PropertyBag
'''
'''   Set pbPar = New PropertyBag
'''
'''   With pbPar
'''      .WriteProperty "COMPROBANTE_TESORERIA", "LIQUI" '"PREF"
'''      .WriteProperty "FECHA", CDate(Format("08/08/2013", "DD/MM/YYYY"))
'''      .WriteProperty "MONEDA_CABECERA", "P" '"C"
'''      '.WriteProperty "FECHA_PAGO", CDate(Format("18/07/2013", "DD/MM/YYYY"))
'''      .WriteProperty "TIPO_CTA_CTE", "" '"B"
'''      .WriteProperty "IMPORTE", CDbl("40000") 'IMPORTE EN LA MONEDA DEL COMPROBANTE
'''      .WriteProperty "TIPO_ENTIDAD", CInt("0") 'CInt("3")
'''      .WriteProperty "CODIGO_ENTIDAD", NullString 'CStr("  11005")
'''      .WriteProperty "COMENTARIO", "Comentario de ejemplo (Pablo Pellegrini)"
'''      .WriteProperty "TIPO_CAMBIO", 5.5
'''      .WriteProperty "MONEDA_CTA_CTE", "" '"C"
'''
'''      .WriteProperty "CUENTA_BANCARIA_ORIGEN", "1"
'''      .WriteProperty "MEDIO_DE_PAGO_ORIGEN", "CB"
'''
'''      '<Liquidación>
'''      .WriteProperty "MEDIO_DE_PAGO_APLICACION", "CBD"
'''      .WriteProperty "CUENTA_BANCARIA_APLICACION", "10"
'''      'En la liqudiación la moneda del comprobante siempre sera Pesos, el importe dolar solo se informara si la moneda de la prefinanciación de la liquidación
'''      'es dolares
'''      .WriteProperty "IMPORTE_DOLAR", CDbl("100000") 'SOLO SE INFORMA SI HAY IMPORTE DOLAR Y SE UTILIZA ESTE PARA CALCULAR LA APLICACIÓN. SOLO SE UTILIZA CUANDO ES LIQUIDACIÓN.
'''      '</Liquidación>
'''
'''      .WriteProperty "TIPO_ORIGEN", "Liquidacion" '"Prefinanciacion"
'''
'''   End With
'''   dim objhistotesoreria as object
'''   Set objhistotesoreria = ArmarComprobanteTesoreria(m, pbPar)
'''
'''End Sub



Public Function ArmarComprobanteTesoreria(ByVal mvarControlData As Variant, ByVal pbParametros As PropertyBag) As Object 'BOGesCom.clsHistoricoTesoreria 'PropertyBag
         Dim objGestionComercial       As Object 'puedo hacer referencia a aplications porque esta en todos los proyectos (es shared). si llegara a generar error en un futuro , paras a tipo object
         Dim hWndAdmin                 As Long
         Dim objGesComForms            As Object
         Dim ix                        As Integer
         Dim frmComprobanteTesoreria   As Form  ' Form
         Dim objHistoricoTesoreria     As Object  'CAMBIAR ANTES DE PROTEGER POR OBJECT
   
   
10       On Error GoTo GestErr
   
      '   Set ArmarComprobanteTesoreria = New PropertyBag
   
   
         '------------------------------------------------------------------------------------------
         '  Armado de objeto
         '------------------------------------------------------------------------------------------
20       Set objHistoricoTesoreria = CreateObject("BOGesCom.clsHistoricoTesoreria")
30       objHistoricoTesoreria.ControlData = mvarControlData
   
40       objHistoricoTesoreria.InitProperties
   
50       With objHistoricoTesoreria
            'Cabecera:
            '.ComprobanteTesoreria = pbParametros.ReadProperty("COMPROBANTE_TESORERIA", NullString)
            '.FechaTesoreria = pbParametros.ReadProperty("FECHA", 0)
            '.Moneda = pbParametros.ReadProperty("MONEDA_CTA_CTE", NullString)
            '.TipoCambioEmision = pbParametros.ReadProperty("TIPO_CAMBIO", 0)
            '.FechaPago = pbParametros.ReadProperty("FECHA_PAGO", 0)
            '.TipoEntidad = pbParametros.ReadProperty("TIPO_ENTIDAD", 0)
            '.Entidad = pbParametros.ReadProperty("CODIGO_ENTIDAD", NullString)
60          .Comentarios = pbParametros.ReadProperty("COMENTARIO", NullString)
      
70       End With
   
   
         '------------------------------------------------------------------------------------------
         '  Obtener formulario
         '------------------------------------------------------------------------------------------
80       Set objGestionComercial = mvarMDIForm.GetInstance("GestionComercial")
   
90       hWndAdmin = mvarMDIForm.CallAdmin("EMISION_COMPROBANTES_TESORERIA", mvarControlData, True)
   
100      DoEvents
         
110      Set objGesComForms = objGestionComercial.CollectionForms
   
120      DoEvents
         
130      For ix = 0 To objGesComForms.Count - 1
140          If objGesComForms(ix).hWnd = hWndAdmin Then
150            Set frmComprobanteTesoreria = objGesComForms(ix)
160            Exit For
170         End If
180      Next
   
   
         '------------------------------------------------------------------------------------------
         '  Completar objeto
         '------------------------------------------------------------------------------------------
         '...
         '..
190      frmComprobanteTesoreria.CompletarComprobante pbParametros.ReadProperty("COMPROBANTE_TESORERIA", NullString), _
                                                      pbParametros.ReadProperty("FECHA", 0), _
                                                      pbParametros.ReadProperty("TIPO_ENTIDAD", 0), _
                                                      pbParametros.ReadProperty("CODIGO_ENTIDAD", NullString), _
                                                      , pbParametros.ReadProperty("TIPO_CTA_CTE", NullString), _
                                                      , , , _
                                                      pbParametros.ReadProperty("FECHA_PAGO", 0), _
                                                      pbParametros.ReadProperty("IMPORTE", 0), _
                                                      objHistoricoTesoreria, _
                                                      , , , , _
                                                      pbParametros.ReadProperty("MONEDA_CABECERA", NullString), _
                                                      pbParametros.ReadProperty("TIPO_CAMBIO", 0), _
                                                      True, _
                                                      False, _
                                                      pbParametros.ReadProperty("MONEDA_CTA_CTE", NullString), _
                                                      pbParametros.ReadProperty("TIPO_CTA_CTE", NullString), _
                                                      pbParametros.ReadProperty("TIPO_ORIGEN", NullString), _
                                                      pbParametros.ReadProperty("CUENTA_BANCARIA_ORIGEN", NullString), pbParametros.ReadProperty("MEDIO_DE_PAGO_ORIGEN", NullString), _
                                                      pbParametros.ReadProperty("CUENTA_BANCARIA_APLICACION", NullString), pbParametros.ReadProperty("MEDIO_DE_PAGO_APLICACION", NullString), pbParametros.ReadProperty("IMPORTE_DOLAR", 0), pbParametros.ReadProperty("DETALLE_MEDIOS_PAGO", NullString), _
                                                      pbParametros.ReadProperty("CUENTA_CONTABLE", NullString), pbParametros.ReadProperty("CUENTA_CONTABLE_CONCEPTO", NullString), pbParametros.ReadProperty("CENTRO_DE_COSTO", NullString)
                                                
                                                

200      Set objHistoricoTesoreria = frmComprobanteTesoreria.ComprobanteTesoreriaCompleto
   
         '-------------------------------------------------
         '  Armar valores de salida:
         '-------------------------------------------------
         'ArmarComprobanteTesoreria.WriteProperty "COMPROBANTE_TESORERIA_COMPLETO", objHistoricoTesoreria.GetState(True)
   
         'Antes retornaba el GetState, ahora retorna una instancia al objeto ya que hay propiedades ondemand que se pierden (una total poronga),por ej: el detalle de centros de costo en el comprobante de pago de intereses.
210      Set ArmarComprobanteTesoreria = objHistoricoTesoreria
   
         'Nothings:
220      Set objHistoricoTesoreria = Nothing
   
230      If Not frmComprobanteTesoreria Is Nothing Then Unload frmComprobanteTesoreria
240      Set frmComprobanteTesoreria = Nothing
   
250      Exit Function
GestErr:

260      LoadError ErrorLog, "ArmarComprobanteTesoreria" & Erl
   
270      Set objHistoricoTesoreria = Nothing
280      If Not frmComprobanteTesoreria Is Nothing Then Unload frmComprobanteTesoreria
290      Set frmComprobanteTesoreria = Nothing
   
300      Err.Raise ErrorLog.NumError, ErrorLog.source, ErrorLog.Descripcion
         'ShowErrMsg ErrorLog
End Function

Public Function LimpiarOCEmitidas(ByVal vData As Variant, ByVal strParametros As String) As String
      'TP 17409 INC 163891
   
         Dim aByte()          As Byte
         Dim pbParametros      As PropertyBag
         Dim pbSalida          As PropertyBag
         Dim SQL               As String
         Dim dFechaHasta       As Date
         Dim strPuntosEgresosIncluye  As String
         Dim strPuntosEgresosExcluye  As String
         Dim strWhereIncluye   As String
         Dim strWhereExcluye   As String
         Dim rstListaEgresos   As ADODB.Recordset
         Dim aPuntos()         As String
         Dim ix                As Integer
         Dim lngErrores        As Long
         Dim lngLimpiados      As Long
         Dim strCPsConErrores  As String
         Dim bErrors           As Boolean
         Dim objEgresoPlanta   As Object 'Shared esta en todos los proyecto. No se puede usar BOCereales.clsEgresoPlanta 'TP 17409 INC 163891
         
   
10       On Error GoTo GestErr
   
20       miControlData = vData 'Le paso el ControlData

30       Set pbParametros = New PropertyBag
40       aByte = strParametros
50       pbParametros.Contents = aByte

60       dFechaHasta = pbParametros.ReadProperty("FECHA_HASTA", 0)
70       strPuntosEgresosIncluye = pbParametros.ReadProperty("PUNTO_EGR_INCLUYE", NullString)
80       strPuntosEgresosExcluye = pbParametros.ReadProperty("PUNTO_EGR_EXCLUYE", NullString)
   
   
         'Armo la Where Punto Egresos Incluir
90       strWhereIncluye = NullString
100      aPuntos = Split(Replace(strPuntosEgresosIncluye, ",", ";"), ";")
110      For ix = LBound(aPuntos) To UBound(aPuntos)
120         If Trim(aPuntos(ix)) <> NullString Then
   
130            If strWhereIncluye = NullString Then
140               strWhereIncluye = " AND (   EGRESOS_PLANTAS.EPL_PUNTO_EGRESO = " & aPuntos(ix)
150            Else
160               strWhereIncluye = strWhereIncluye & " OR EGRESOS_PLANTAS.EPL_PUNTO_EGRESO = " & aPuntos(ix)
170            End If

180         End If
190      Next ix
   
200      If strWhereIncluye <> NullString Then
210         strWhereIncluye = strWhereIncluye & "  )  "
220      End If

         'Armo la Where Punto Egresos Excluir
230      strWhereExcluye = NullString
240      aPuntos = Split(Replace(strPuntosEgresosExcluye, ",", ";"), ";")
250      For ix = LBound(aPuntos) To UBound(aPuntos)
260         If Trim(aPuntos(ix)) <> NullString Then
270            strWhereExcluye = strWhereExcluye & " AND EGRESOS_PLANTAS.EPL_PUNTO_EGRESO <> " & aPuntos(ix) & "  "
280         End If
290      Next ix
   


         '*********************************************
         'Proceso de Obtencion de OC para Limpiar
         '*********************************************

300       SQL = "SELECT * FROM (SELECT EGRESOS_PLANTAS.EPL_PUNTO_EGRESO, EGRESOS_PLANTAS.EPL_NUMERO_EGRESO, "
310       SQL = SQL & "         MAX (EGRESOS_PLANTAS_ESTADOS.EPE_CODIGO_ESTADO) EPE_CODIGO_ESTADO "
320       SQL = SQL & "    FROM EGRESOS_PLANTAS, EGRESOS_PLANTAS_ESTADOS "
330       SQL = SQL & "   WHERE EGRESOS_PLANTAS.EPL_PUNTO_EGRESO = EGRESOS_PLANTAS_ESTADOS.EPE_PUNTO_EGRESO "
340       SQL = SQL & "     AND EGRESOS_PLANTAS.EPL_NUMERO_EGRESO = EGRESOS_PLANTAS_ESTADOS.EPE_NUMERO_EGRESO "
    
350       If strWhereIncluye <> NullString Then
360          SQL = SQL & strWhereIncluye
370       End If

380       If strWhereExcluye <> NullString Then
390          SQL = SQL & strWhereExcluye
400       End If

410       SQL = SQL & "     AND EGRESOS_PLANTAS.EPL_FECHA_HORA_OC <= TO_DATE ('" & dFechaHasta & "', 'DD/MM/YYYY') "
420       SQL = SQL & "   GROUP BY EGRESOS_PLANTAS.EPL_PUNTO_EGRESO, EGRESOS_PLANTAS.EPL_NUMERO_EGRESO "
430       SQL = SQL & "   ORDER BY EGRESOS_PLANTAS.EPL_PUNTO_EGRESO, EGRESOS_PLANTAS.EPL_NUMERO_EGRESO) EGRESOS "
440       SQL = SQL & " WHERE EGRESOS.EPE_CODIGO_ESTADO = " & 3 'EstadoOrdenCarga = 3

450       Set rstListaEgresos = Fetch(miControlData.Empresa, SQL)
    
460       On Error Resume Next
   
470       lngErrores = 0
480       lngLimpiados = 0
490       strCPsConErrores = NullString
500       bErrors = False
    
         '*********************************************
         'Proceso de Limpieza de OC Obtenidas
         '********************************************
   
510       Do While Not rstListaEgresos.EOF

520          Set objEgresoPlanta = CreateObject("BOCereales.clsEgresoPlanta")
530          objEgresoPlanta.ControlData = miControlData
540          objEgresoPlanta.LimpiarMovimientosPendientes rstListaEgresos("EPL_PUNTO_EGRESO").Value, rstListaEgresos("EPL_NUMERO_EGRESO").Value, dFechaHasta
550          If Err = 0 Then
560             Err.Clear
570             lngLimpiados = lngLimpiados + 1
580          Else
590             bErrors = True
600             strCPsConErrores = strCPsConErrores & "Egreso: " & rstListaEgresos("EPL_PUNTO_EGRESO").Value & "-" & rstListaEgresos("EPL_NUMERO_EGRESO").Value & vbCrLf & "Error: " & Err.Description & vbCrLf & "Fuente: " & Err.source & vbCrLf & String(73, "_") & vbCrLf
610             lngErrores = lngErrores + 1
620          End If
630          rstListaEgresos.MoveNext
640       Loop

650       On Error GoTo GestErr
 
660       strCPsConErrores = lngErrores & " errores" & vbCrLf & vbCrLf & strCPsConErrores
670       strCPsConErrores = rstListaEgresos.RecordCount - lngErrores & " Egresos eliminados" & vbCrLf & strCPsConErrores

   
680       Set pbSalida = New PropertyBag
   
690       pbSalida.WriteProperty "CANTIDAD_LIMPIADA", lngLimpiados
700       pbSalida.WriteProperty "CANTIDAD_ERRORES", lngErrores
710       pbSalida.WriteProperty "CP_ERRORES_DESC", strCPsConErrores
   
720       LimpiarOCEmitidas = pbSalida.Contents

730       If Not rstListaEgresos Is Nothing Then
740          If rstListaEgresos.State <> adStateClosed Then rstListaEgresos.Close
750       End If
760       Set rstListaEgresos = Nothing
   
   
770       Set objEgresoPlanta = Nothing

780       Exit Function

GestErr:
790       LoadError ErrorLog, "LimpiarOCEmitidas" & Erl
800       If Not rstListaEgresos Is Nothing Then
810          If rstListaEgresos.State <> adStateClosed Then rstListaEgresos.Close
820       End If
830       Set rstListaEgresos = Nothing
   
840       Set objEgresoPlanta = Nothing
    
850       ShowErrMsg ErrorLog
   
End Function




'---------------------------------------------------------------------------------------
' Procedure : GetInfoCodigoTrazabilidad
' DateTime  : 07/02/2014 10:46
' Author    : pablo.pellegrini
' TP        : 19870
'---------------------------------------------------------------------------------------
'Descripción:
'Esta funcion acepta por parametro un codigo de trazabilidad en uno de los dos formatos siguientes:
'  1)Humano: (01)07790001000019(11)121210(17)141212(10)8765(21)654
'  2)Sanner: 01077900010000191112121017141212108765<F8>21654
'La función busca primero los IA de longitud fija y luego los de longitud variable.
'La función busca los IA sin orden.
'La función retorna un string buffer de property bug con los datos:
'                                                                    GTIN
'                                                                    LOTE
'                                                                    NUMERO_SERIE
'                                                                    FECHA_ELABORACION
'                                                                    FECHA_VENCIMIENTO
'                                                                    CODIGO_FORMATEADO
'El caracter IA significa "Indicador de Aplicación".
Public Function GetInfoCodigoTrazabilidad(ByVal strCodigo As String) As udtInfoCodigoTrazabilidad
   'Constantes:
   Const PA                            As String = "("
   Const PC                            As String = ")"
'   Const CODIGO_SEPARADOR_GRUPO        As Integer = 119 'KeyCode = 119
   
   Const IA_CODIGOS_LONGITUD_FIJA      As String = "01,11,17"  'codigos a identificar de longitud fija
   Const IA_CODIGOS_LONGITUD_VARIABLE  As String = "10,21"     'codigos a identificar de longitud variable
   Const IA_LONGITUDES                 As String = "16,8,8"    'longitudes de las secciones de codigos a identificar de longitud fija concidewrando el codigo
   
   'Variables:
   Dim pb                           As PropertyBag
   Dim ix                           As Long
   Dim iy                           As Long
   Dim strCodigoWork                As String
   Dim strTemp                      As String
   
   Dim arrCodigosLongitudFija()     As String
   Dim arrCodigosLongitudVar()      As String
   Dim arrLongitudes()              As String
   
   Dim strCodigoFormateado          As String
   
   Dim lngPosicion                  As Long
   Dim lngLongitud                  As Long
   Dim strSegmento                  As String

   Dim bProcesar                    As Boolean
   Dim iDif                         As Integer
   Dim strCodigoIA                  As String
   Dim strCodigoIAFormat            As String
   
   Dim lngIAEncontradosLongitudFija As Long
   Dim lngIAEncontradosLongitudVar  As Long

   strCodigoWork = strCodigo
   
   Set pb = New PropertyBag
   
   '---------------------------------------------------
   '1. obtengo los de longitud fija
   '---------------------------------------------------
   arrCodigosLongitudFija = Split(IA_CODIGOS_LONGITUD_FIJA, ",")
   arrLongitudes = Split(IA_LONGITUDES, ",")
   
   lngPosicion = 1
   
   For iy = LBound(arrCodigosLongitudFija) To UBound(arrCodigosLongitudFija) 'este primer for es solo para que haga busqueda carteciana, o sea, en cualquier orden.
      For ix = LBound(arrCodigosLongitudFija) To UBound(arrCodigosLongitudFija)
         'tomo dos caracteres y me fijo si es un codigo de longitud fija, si es, obtengo la longitud a cortar
         bProcesar = False
         If Mid(strCodigoWork, lngPosicion, 2) = arrCodigosLongitudFija(ix) Then             'Estoy procesando en formato scanner
            bProcesar = True
            strCodigoIA = arrCodigosLongitudFija(ix)
            strCodigoIAFormat = PA & arrCodigosLongitudFija(ix) & PC
            iDif = 0
         Else
           If Mid(strCodigoWork, lngPosicion, 4) = PA & arrCodigosLongitudFija(ix) & PC Then 'Estoy procesando formato humano
              bProcesar = True
              strCodigoIA = PA & arrCodigosLongitudFija(ix) & PC                             'Agrego parentecis adelante y atras del codigo de IA ya que se buscarán los IA formateados "(01)", "(17)", etc..
              strCodigoIAFormat = strCodigoIA
              iDif = Len(strCodigoIA) - Len(arrCodigosLongitudFija(ix))
           End If
         End If
         
         If bProcesar Then
            lngLongitud = CLng(arrLongitudes(ix)) + iDif                                     'obtengo la longitud del segmento. iDif se suma porque debe conciderar los caracteres que se le agregan al codigo IA, son dos parentecis.
            strSegmento = Mid(strCodigoWork, lngPosicion, lngLongitud)                       'Obtiene el codigo de segmento(IA) y el contenido del segmento en si. Por ejemplo (01)55556159487987
            strSegmento = Right(strSegmento, Len(strSegmento) - Len(strCodigoIA))            'Obtengo solo elcontenido del segmento, le quito el codigo IA.
            strCodigoFormateado = strCodigoFormateado & strCodigoIAFormat & strSegmento      'Voy armando la cadena formateada del codigo de trazabilidad.
            pb.WriteProperty arrCodigosLongitudFija(ix), strSegmento                         'Agrego al PB la información
            lngPosicion = lngPosicion + lngLongitud                                          'Me posiciono en la proxima posición a analizar según la longitud del segmento actual.
         End If
      Next ix
   Next iy
   
   '--------------------------------------------------------
   '2. obtengo los de longitud variable
   '-------------------------------------------------------
   arrCodigosLongitudVar = Split(IA_CODIGOS_LONGITUD_VARIABLE, ",")
   For iy = LBound(arrCodigosLongitudVar) To UBound(arrCodigosLongitudVar) 'este primer for es solo para que haga busqueda carteciana, o sea, en cualquier orden.
      For ix = LBound(arrCodigosLongitudVar) To UBound(arrCodigosLongitudVar)
         bProcesar = False
         If Mid(strCodigoWork, lngPosicion, 2) = arrCodigosLongitudVar(ix) Then
            bProcesar = True
            strCodigoIA = arrCodigosLongitudVar(ix)
            strCodigoIAFormat = PA & arrCodigosLongitudVar(ix) & PC
            iDif = 0
         Else
           If Mid(strCodigoWork, lngPosicion, 4) = PA & arrCodigosLongitudVar(ix) & PC Then
              bProcesar = True
              strCodigoIA = PA & arrCodigosLongitudVar(ix) & PC
              strCodigoIAFormat = strCodigoIA
              iDif = Len(strCodigoIA) - Len(arrCodigosLongitudVar(ix))
           End If
         End If
         
         If bProcesar Then
            Dim strSG As String 'separador de grupo
            'Obtener la longitud: En este caso al ser variable la longitud depende de si es una codigo trazabilidad humano o scanner.
            'Los 3 posibles fines de seccion pueden ser:
            ' 1. Un caracter "("    = PA
            ' 2. Un caracter "<F8>" = CARACTER_SEPARADOR_GRUPO
            ' 3. Si no es ninguno de los anteriores es hasta el fin del codigo de trazabilidad.
            
            If InStr(lngPosicion, UCase(strCodigoWork), CARACTER_SEPARADOR_GRUPO) = 0 And InStr(lngPosicion + 1, strCodigoWork, PA) = 0 Then
               ' 3. Si no es ninguno de los anteriores es hasta el fin del codigo de trazabilidad.
               lngLongitud = (Len(strCodigoWork) - lngPosicion) + 1
               strSG = "" 'no importa este dato si entra por aqui porque es hasta el final del string
            Else
               If InStr(lngPosicion, UCase(strCodigoWork), CARACTER_SEPARADOR_GRUPO) > 0 Then
                  ' 2. Un caracter "<F8>" = CARACTER_SEPARADOR_GRUPO
                  strSG = CARACTER_SEPARADOR_GRUPO
                  lngLongitud = InStr(lngPosicion, UCase(strCodigoWork), strSG) - lngPosicion
               Else
                  If InStr(lngPosicion + 1, strCodigoWork, PA) > 0 Then
                     ' 1. Un caracter "("    = PA
                     strSG = PA
                     lngLongitud = InStr(lngPosicion + 1, strCodigoWork, strSG) - lngPosicion
                  End If
               End If
            End If
            
            strSegmento = Mid(strCodigoWork, lngPosicion, lngLongitud)
            strSegmento = Right(strSegmento, Len(strSegmento) - Len(strCodigoIA))
            strCodigoFormateado = strCodigoFormateado & strCodigoIAFormat & strSegmento
            pb.WriteProperty arrCodigosLongitudVar(ix), strSegmento
            
            If strSG = PA Then
               lngPosicion = lngPosicion + lngLongitud + 0           'No sumo al posicionamiento porque en teoria estoy parado en un parentecis que deberia ser analizado a continuación
            Else
               lngPosicion = lngPosicion + lngLongitud + Len(strSG)  'si el separador era <F8> sumo esto al posicionamiento para que la proxima interpretación sea desde el proximo codigo y no desde el <F8>.
            End If
         End If
      Next ix
   Next iy
   

   '-----------------------------
   '3. Armo datos de salida
   '-----------------------------
   If Len(pb.ReadProperty("01", "")) > 0 Then
      GetInfoCodigoTrazabilidad.GTIN = pb.ReadProperty("01", "")
   End If
   If Len(pb.ReadProperty("10", "")) > 0 Then
      GetInfoCodigoTrazabilidad.Lote = pb.ReadProperty("10", "")
   End If
   If Len(pb.ReadProperty("21", "")) > 0 Then
      GetInfoCodigoTrazabilidad.NumeroSerie = pb.ReadProperty("21", "")
   End If
   
   
   '----------------------
   'fechas
   Dim dFecha As Date
   If Len(pb.ReadProperty("11", "")) > 0 Then
      If IsNumeric(pb.ReadProperty("11", "")) Then
         strTemp = Format(InvertirSegmentoFecha(pb.ReadProperty("11")), "00-00-00")
         If IsDate(strTemp) Then
            dFecha = CDate(strTemp)
            dFecha = Format(dFecha, "dd/mm/yyyy")
            GetInfoCodigoTrazabilidad.FechaElaboracion = dFecha
         End If
      End If
   End If
   If Len(pb.ReadProperty("17", "")) > 0 Then
      If IsNumeric(pb.ReadProperty("17", "")) Then
         strTemp = Format(InvertirSegmentoFecha(pb.ReadProperty("17")), "00-00-00")
         If IsDate(strTemp) Then
            dFecha = CDate(strTemp)
            dFecha = Format(dFecha, "dd/mm/yyyy")
            GetInfoCodigoTrazabilidad.FechaVencimiento = dFecha
         End If
      End If
   End If
   
   GetInfoCodigoTrazabilidad.CodigoHumanamenteLegible = strCodigoFormateado
   
End Function
Private Function InvertirSegmentoFecha(ByRef strSegmentoFecha As String) As String
   Dim arr() As String
   Dim strSegmentoInvertido As String
   Dim ix As Integer
   
   arr = Split(Format(strSegmentoFecha, "00-00-00"), "-")
   
   For ix = UBound(arr) To LBound(arr) Step -1
      strSegmentoInvertido = strSegmentoInvertido & arr(ix)
   Next ix
   
   InvertirSegmentoFecha = strSegmentoInvertido
   
End Function
