Attribute VB_Name = "modShared"
Option Explicit
 
Public objTextBox As New AlgStdFunc.clsTextBoxEdit
 
'***********************************************************************
' Constantes Propias
'***********************************************************************
Public Const Si                  As String = "Sí"
Public Const No                  As String = "No"
Public Const NullString          As String = ""
Public Const UNKNOWN_ERRORSOURCE As String = "[Fuente de Error Desconocida]"
Public Const KNOWN_ERRORSOURCE   As String = "[Fuente de Error Conocida]"
'***********************************************************************
Private Const MODULE_NAME        As String = "[ModShare]"

'DEFINICIONES PARA EL DATAGRID CONTROL
Public Type udtGrid
   Ancho          As Integer
   Titulo         As String
   CAMPO          As String
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
   CAMPO          As String
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
   MRUForms                      As String
   MRUEmpresas                   As String
   GridQueries                   As String
   NavigationQueries             As String
   PrintQueries                  As String
   QueryDBQueries                As String
   DataComboQueries              As String
End Type

Public Enum EnumIsValid
   Numerico = 1
   Fecha = 2
   Hora = 3
End Enum

Public RegistrySubKeys           As udtRegistrySubKeys

Public vValue                    As Variant

Public SystemOptions             As udtSystemOptions

Public aAppReg()                                                     'matriz de aplicaciones registradas de Algortimo
Public aKeys()                                                       'matriz para la lectura de tablas
Public MRUForms                  As New Collection                   'coleccion de forms mas frecuentemente usados

Public Enum alFetchMode
   alAsync = 1
   alSync = 2
   alTable = 3
End Enum

Public Enum ContextMenuEnum
   mnxModulo = 0
   mnxNombre = 1
   mnxOrden = 2
   mnxForms = 3
   mnxCaption = 4
   mnxTarea = 5
   mnxClave = 6
End Enum

'  constantes para identificar los paneles del Status Bar del ABM Clasico
Public Const STB_PANEL1              As Integer = 1
Public Const STB_PANEL2              As Integer = 2
Public Const STB_PANEL3              As Integer = 3
Public Const STB_PANEL4              As Integer = 4

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

Public rstContextMenu As ADODB.Recordset
Public rstMenu        As ADODB.Recordset
Public rstVistasPersonalizadas   As ADODB.Recordset

Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetFocus Lib "USER32" (ByVal hWnd As Long) As Long

'Mensajes enviados por la clase clsControls
Public Const CTL_QUERYDB_IS_ACTIVE = &H1
Public Const CTL_QUERYDB_RECORD_SELECTED = &H2
Public Const CTL_CALL_ADMIN = &H3
Public Const CTL_QUERY_USER   As Long = &H4
Public Const CTL_QUERY_EMPRESA   As Long = &H5

'Mensajes enviados por la Filter
Public Const FILTER_CALL_ADMIN = &H1
Public Const FILTER_QUERY_USER   As Long = &H2
Public Const FILTER_QUERY_EMPRESA   As Long = &H3

Public Const WM_KEYDOWN = &H100
Public Const WM_CHAR = &H102
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_RBUTTONDOWN = &H204

Public Const EM_EMPTYUNDOBUFFER = 205
Public Const EM_CANUNDO = 198
Public Const EM_GETMODIFY = 184
Public Const EM_SETMODIFY = 185
Public Const EM_UNDO = 199

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

Public Function Dialog(strTipo As String, Optional strDefaultFileName As String, Optional strFilter As String, Optional iTop As Integer, Optional iLeft As Integer) As String

  On Error Resume Next
  
  If iTop <> 0 Then
    frmCommonDialog.Top = iTop
    frmCommonDialog.Left = iLeft
  Else
    frmCommonDialog.Top = 0
    frmCommonDialog.Left = 0
  End If

  frmCommonDialog.CommonDialog1.CancelError = True
  
  Select Case strTipo
    Case "ShowOpen"
      frmCommonDialog.CommonDialog1.FileName = strDefaultFileName
      frmCommonDialog.CommonDialog1.Filter = strFilter

      frmCommonDialog.CommonDialog1.ShowOpen
      
    Case "ShowSave"
      frmCommonDialog.CommonDialog1.FileName = strDefaultFileName
      frmCommonDialog.CommonDialog1.Filter = strFilter

      frmCommonDialog.CommonDialog1.ShowSave
      
  End Select
  
  If Err.Number = cdlCancel Then
    frmCommonDialog.CommonDialog1.FileName = NullString
    Exit Function
  End If
  
  Dialog = frmCommonDialog.CommonDialog1.FileName
  
  Unload frmCommonDialog

End Function

Public Function GetData(ByVal strEmpresa As String, ByVal strSQL As String, Optional ByVal iCursorType As CursorTypeEnum, Optional ByVal iLockType As LockTypeEnum) As ADODB.Recordset
Dim rst As ADODB.Recordset

   On Error GoTo GestErr
   
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
         Set rst = DataAccess.Fetch(strEmpresa, strSQL, iCursorType, iLockType)
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
   
   Exit Function
   
GestErr:

   LoadError "GetData"
   ShowErrMsg

End Function

Public Sub ReadSystemOptions()

   ' lectura de los parametros internos

   With SystemOptions
   
      '  CacheSize
      vValue = GetKeyValuePI("ADO\CacheSize")
      .iCacheSize = IIf(IsNull(vValue), 1, vValue)
      
      '  Zoom
      vValue = GetKeyValuePI("Opciones\Zoom Vista Previa\Valor Generico", 80)
      .iZoom = IIf(IsNull(vValue), 70, vValue)
      
      '  Fetch Mode
      vValue = GetKeyValuePI("Performance\FetchMode")
      .iFetchMode = IIf(IsNull(vValue), 1, vValue)
      If (.iFetchMode <> alAsync) And (.iFetchMode <> alSync) And (.iFetchMode <> alTable) Then
         MsgBox "El valor del parámetro 'Performance\FetchMode' admite los siguientes valores:" & vbCrLf & _
                "    1 - Fetch Asincrónico" & vbCrLf & _
                "    2 - Fetch Sincrónico" & vbCrLf & _
                "    3 - Variable según 'Performance\Limite Fetch Sincrónico'" & vbCrLf & _
                "En caso de omisión, asume la opción 2", vbInformation, App.ProductName
      End If
      
      If .iFetchMode = alTable Then
         vValue = GetKeyValuePI("Performance\Limite Fetch Sincronico")
         .lngFetchLimit = IIf(IsNull(vValue), 1000, vValue)
      End If
      
      
      '  Fetch Mode Busqueda
      vValue = GetKeyValuePI("Performance\FetchMode en Busqueda")
      .iFetchModeSearch = IIf(IsNull(vValue), 1, vValue)
      If (.iFetchModeSearch <> alAsync) And (.iFetchModeSearch <> alSync) And (.iFetchModeSearch <> alTable) Then
         MsgBox "El valor del parámetro 'Performance\FetchMode' admite los siguientes valores:" & vbCrLf & _
                "    1 - Fetch Asincrónico" & vbCrLf & _
                "    2 - Fetch Sincrónico" & vbCrLf & _
                "    3 - Variable según 'Performance\Limite Fetch Sincrónico'" & vbCrLf & _
                "En caso de omisión, asume la opción 2", vbInformation, App.ProductName
      End If
      
      If .iFetchModeSearch = alTable Then
         vValue = GetKeyValuePI("Performance\Limite Fetch Sincronico en Busqueda")
         .lngFetchLimitSearch = IIf(IsNull(vValue), 300, vValue)
      End If
      
      
      '  Usa copias locales
      vValue = GetKeyValuePI("Performance\Usa Copias Locales")
      .UseLocalCopy = IIf(IsNull(vValue), Si, vValue)
      If (.UseLocalCopy <> Si) And (.UseLocalCopy <> No) Then
         MsgBox "El valor del parámetro 'Performance\Usa Copias Locales' puede ser Sí o No:" & vbCrLf & _
                "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
      End If
   
      '  Usa copias locales en Búsquedas
      vValue = GetKeyValuePI("Performance\Usa Copias Locales en Busqueda")
      .UseLocalCopySearch = IIf(IsNull(vValue), Si, vValue)
      If (.UseLocalCopySearch <> Si) And (.UseLocalCopySearch <> No) Then
         MsgBox "El valor del parámetro 'Performance\Usa Copias Locales en Búsqueda' puede ser Sí o No:" & vbCrLf & _
                "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
      End If
   
      '  Pregunta si usa copias locales desactualizadas
      vValue = GetKeyValuePI("Performance\Usa Copias Locales Desactualizadas")
      .AskOldLocalCopy = IIf(IsNull(vValue), Si, vValue)
      If (.AskOldLocalCopy <> Si) And (.AskOldLocalCopy <> No) Then
         MsgBox "El valor del parámetro 'Performance\Usa Copias Locales Desactualizadas' puede ser Sí o No:" & vbCrLf & _
                "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
      End If
      
      '  Pregunta si usa copias locales desactualizadas
      vValue = GetKeyValuePI("Opciones\Empresas\Usa MRU de Empresas")
      .UseMRUEnterprise = IIf(IsNull(vValue), Si, vValue)
      If (.UseMRUEnterprise <> Si) And (.UseMRUEnterprise <> No) Then
         MsgBox "El valor del parámetro 'Opciones\Empresas\Usa MRU de Empresas' puede ser Sí o No:" & vbCrLf & _
                "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
      End If
      
      '  Dimension de MRUForms
      vValue = GetKeyValuePI("Performance\Dimension MRUForms")
      .MaxMRUForms = IIf(IsNull(vValue), 0, vValue)
      
   End With

End Sub

Public Sub ShowErrMsg()
Dim iErrNumber         As Long                          ' numero de error (sin vbObjectError)
Dim bAlgError          As Boolean                       ' identifica un error de Algoritmo
Dim ix                 As Integer
Dim strSource          As String
Dim n                  As Integer

   '  muestra en manera amigable un mensaje de error
   

   strSource = Trim(ErrorLog.Source)
   
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
            MsgBox ErrorLog.Descripcion, vbOKOnly, App.ProductName
         Case 10000 To 20000
            'Errores Severos de Algoritmo
            MsgBox ErrorLog.Descripcion, vbExclamation, "Error Manager"
      End Select
   Else
      ' errores no generados por Algoritmo
      MsgBox "Se produjo el siguiente error:" & vbCrLf & vbCrLf & _
             "Número     : " & ErrorLog.NumError & vbCrLf & vbCrLf & _
             "Descripción: " & vbCrLf & ErrorLog.Descripcion & vbCrLf & vbCrLf & _
             "Llamadas   : " & vbCrLf & _
             strSource, vbExclamation, "Error Manager"
   End If
   
   ' una vez visualizado el mensaje de error, este viene limpiado
   ClearError
   
End Sub

Public Sub LoadMenu(strMenuName As String, frm As Form)
Dim ix   As Integer
Dim iPos As Integer

   '  carga el menu indicado por el parametro strMenuname en el form frm
   
   On Error Resume Next
   
   Select Case strMenuName
      Case "Tools"
         
         '  menu Herramientas
         For ix = 0 To UBound(aMenuTools)
            If Len(aMenuTools(ix)) > 0 Then
               Load frm.mnuToolsItem(ix)
               iPos = InStr(aMenuTools(ix), ";")
               If iPos > 0 Then
                  frm.mnuToolsItem(ix).Caption = Left(aMenuTools(ix), iPos - 1)
               Else
                  frm.mnuToolsItem(ix).Caption = aMenuTools(ix)
               End If
            End If
         Next ix
      Case "Help"
       
       ' menu  Ayuda
       For ix = 0 To UBound(aMenuHelp)
          Load frm.mnuAyudaItem(ix)
          iPos = InStr(aMenuHelp(ix), ";")
          If iPos > 0 Then
             frm.mnuAyudaItem(ix).Caption = Left(aMenuHelp(ix), iPos - 1)
          Else
             frm.mnuAyudaItem(ix).Caption = aMenuHelp(ix)
          End If
      Next ix
   End Select

End Sub

Public Sub CallToolsItem(ByVal iItem As Integer)
Dim iPos    As Integer
Dim strForm As String
Dim Form    As Form

   ' llama una opcíon del menú Herramientas

   iPos = InStr(aMenuTools(iItem), ";")
   If iPos > 0 Then
      strForm = Mid(aMenuTools(iItem), iPos + 1)
   Else
      strForm = ""
   End If

   Select Case Left(strForm, 3)
      Case "frm"
         Set Form = Forms.Add(Trim(strForm))
         Form.Show
   End Select
   
   Select Case strForm
      Case "ReiniciarSesion"
         Call ReiniciarSesion
   End Select
   
End Sub

Public Sub ReiniciarSesion()
Dim strOldForm As String

   ' cierro todos los forms MDI children
   
   Screen.MousePointer = vbHourglass
   
   strOldForm = ""
   
   Do While Not (Forms(0).ActiveForm Is Nothing)
      If strOldForm <> Forms(0).ActiveForm.Name Then
         strOldForm = Forms(0).ActiveForm.Name
         Unload Forms(0).ActiveForm
      Else
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   Loop
   Screen.MousePointer = vbDefault
   
   '  inicio nuevo login
   FrmLogin.Show vbModal
    
   ' defino el application path para la clase clsEnvironment
   SetAppPath App.Path
   
   ReadSystemOptions
   
   With ErrorLog
      .Maquina = CSysEnvironment.Machine
      .Aplicacion = App.EXEName
   End With
   
   If UCase(App.EXEName) <> "SEGURIDAD" Then
      If App.StartMode = vbSModeStandalone Then
         ' mostrar menú
         frmMenu.Show
      End If
   End If
   
   
End Sub

Public Sub LoadError(ByVal strSource As String)

   ' carga la informaciòn del error en la variable ErrorLog

   SetError ErrorLog, App.ProductName, strSource
   
   TrapError ErrorLog
   
End Sub
Public Sub SetError(ByRef ErrLog As ErrType, ByVal strModuleName As String, ByVal strSource As String)

   With ErrLog
   
      .Modulo = strSource
      .NumError = Err.Number
      .Source = Err.Source
      .WriteError = True
      
      If InStr(.Source, KNOWN_ERRORSOURCE) = 0 Then
         If InStr(.Source, UNKNOWN_ERRORSOURCE) = 0 Then
            .Source = UNKNOWN_ERRORSOURCE & vbCrLf & .Source
         End If
      Else
         .Source = Replace(.Source, KNOWN_ERRORSOURCE, NullString)
         .WriteError = False
      End If
      
      If InStr(.Source, strModuleName) > 0 Then
         .Source = .Source & "[" & strSource & "]"
      Else
         .Source = .Source & vbCrLf & strModuleName & "[" & strSource & "]"
      End If
      
      .Descripcion = Err.Description

   End With

End Sub

Public Sub ClearError()

   ' carga la informaciòn del error en la variable ErrorLog
   
   With ErrorLog
      .Modulo = NullString
      .NumError = 0
      .Source = NullString
      .Descripcion = NullString
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
Dim s As String

    ' rutina recursiva que copia o mueve todos los hijos de un nodo a otro nodo
    
    If sourceND.Children = 0 Then Exit Sub
    
    Set so = sourceND.Child
    For ix = 1 To sourceND.Children
'        s = so.key
'        so.key = ""
        ' agrega un nodo en el TreeView de destino
        Set de = DestTV.Nodes.Add(destND, tvwChild, so.key, so.Text, so.Image, so.SelectedImage)
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

   For ix = 0 To UBound(GridInfo.ColsProperties)

      GridInfo.Grilla.Columns(ix).Width = GridInfo.ColsProperties(ix).Ancho
      GridInfo.Grilla.Columns(ix).Caption = GridInfo.ColsProperties(ix).Titulo
      GridInfo.Grilla.Columns(ix).DataField = GridInfo.ColsProperties(ix).CAMPO
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

End Sub

Public Sub RefreshFLEXGrid(GridInfo As udtRefreshFLEXGrid)
Dim ix As Integer

   For ix = 0 To UBound(GridInfo.ColsProperties)

      GridInfo.Grilla.ColWidth(ix + 1) = GridInfo.ColsProperties(ix).Ancho
      GridInfo.Grilla.TextMatrix(0, ix + 1) = GridInfo.ColsProperties(ix).Titulo
      GridInfo.Grilla.ColAlignment(ix + 1) = GridInfo.ColsProperties(ix).Alineacion
      
      If Len(GridInfo.ColsProperties(ix).Permiso) > 0 Then
         If Not TaskIsEnabled(GridInfo.ColsProperties(ix).Permiso, CUsuario) Then
            GridInfo.Grilla.ColWidth(ix) = 0
         End If
       End If
      
   Next ix

   GridInfo.Grilla.Refresh

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
   aArray(ix).CAMPO = strField
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
   aArray(ix).CAMPO = strField
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
Dim strWHERE            As String
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
   strWHERE = ""
   For ix = 1 To UBound(aArray, 2)
   
      For ix1 = 0 To UBound(aArray, 1)
         
         iTypeOfField = FieldProperty(aTableProperties, aArray(ix1, 0), dsTipoDato)
         
         If Len(aArray(ix1, ix)) = 0 Then
            strWHERE = strWHERE & "(" & strTableName & "." & aArray(ix1, 0) & " IS NULL) AND "
         Else
            Select Case iTypeOfField
              Case adChar, adVarChar
                  strWHERE = strWHERE & "(" & strTableName & "." & aArray(ix1, 0) & " = '" & aArray(ix1, ix) & "') AND "
               Case adNumeric
                  strWHERE = strWHERE & "(" & strTableName & "." & aArray(ix1, 0) & " = " & aArray(ix1, ix) & ") AND "
               Case adDBTimeStamp
                  strWHERE = strWHERE & "(" & strTableName & "." & aArray(ix1, 0) & " = TO_DATE(Format(" & """" & aArray(ix1, ix) & """" & ", " & """" & "yyyy-mm-dd" & """" & "), " & """" & "YYYY-MM-DD" & """" & ")" & ") AND "
            End Select
         End If
'                                                                                     TO_DATE('" & Format(aArray(ix1, ix), "yyyy-mm-dd") & "', 'YYYY-MM-DD')
      Next ix1
      
      If Right(strWHERE, 5) = " AND " Then
         strWHERE = Left(strWHERE, Len(strWHERE) - 5)
      End If
      
      strWHERE = strWHERE & " OR "
   Next ix
   
   
   If Right(strWHERE, 4) = " OR " Then
      strWHERE = Left(strWHERE, Len(strWHERE) - 4)
   End If
      
   MakeSQLWhere = strWHERE
   
End Function

Public Sub LoadMRUForms()
Dim aArray()            As Variant
Dim aMostUsedForms()    As Variant
Dim ix                  As Integer, ix1 As Integer
Dim MaxElem             As Integer
Dim strForm             As String
Dim MRUMaxItems         As Integer
Dim frm                 As Form


   '  aArray es un arreglo bidimensional
   '     1· Dimension: numero de veces que fue cargado el form en una sesion de trabajo
   '     2· Dimension: nombre del form
   
   Set MRUForms = Nothing

   'si no existe la clave la creo
   If Not CheckRegistryKey(HKEY_LOCAL_MACHINE, RegistrySubKeys.MRUForms) Then
      CreateRegistryKey HKEY_LOCAL_MACHINE, RegistrySubKeys.MRUForms
   End If

   'obtengo la estadistica de aperturas
   aArray = EnumRegistryValues(HKEY_LOCAL_MACHINE, RegistrySubKeys.MRUForms)
   
   If IsArrayEmpty(aArray) Then Exit Sub
   
   
   MRUMaxItems = IIf(UBound(aArray, 2) < SystemOptions.MaxMRUForms, UBound(aArray, 2), SystemOptions.MaxMRUForms)
   ReDim aMostUsedForms(MRUMaxItems)
   
   'lleno aMostUsedForms con los MRUmaxItems forms mas cargados
   For ix = 0 To MRUMaxItems
      MaxElem = LBound(aArray, 2)
      For ix1 = LBound(aArray, 2) To UBound(aArray, 2)
         
         If aArray(1, ix1) > aArray(1, MaxElem) Then
            MaxElem = ix1
         End If

      Next ix1
      
      If ix = 0 And aArray(1, MaxElem) = 1 Then
         'no hay forms con valores > 1
         Exit For
      End If
      
      If aArray(1, MaxElem) > 1 Then
         strForm = aArray(0, MaxElem)
         If Not ScanArray(aMostUsedForms, strForm) Then
            aMostUsedForms(ix) = aArray(0, MaxElem)
            aArray(1, MaxElem) = 0
         End If
      End If
      
   Next ix

   'cargo los forms mas frecuentemente cargados en la sesion anterior
   For ix = LBound(aMostUsedForms) To UBound(aMostUsedForms)
      If aMostUsedForms(ix) = NullString Then Exit For
      Set frm = Forms.Add(aMostUsedForms(ix))
      If Not frm.PropertlyLoaded Then Unload frm: Exit Sub
      
      MRUForms.Add Item:=frm, key:=aMostUsedForms(ix)
   Next ix
   
   'limpio el registro
   For ix1 = LBound(aArray, 2) To UBound(aArray, 2)
      DeleteRegistryValue HKEY_LOCAL_MACHINE, RegistrySubKeys.MRUForms, aArray(0, ix1)
   Next ix1
   
End Sub

Public Function IsMRUForm(ByVal lngHndW As Long) As Boolean
Dim frm As Form

   '  determina si un forms esta cargado en la colección MRUForms
   
   For Each frm In MRUForms

      If frm.hWnd = lngHndW Then IsMRUForm = True: Exit For

   Next frm
   
End Function

Public Sub AutoIncr(ByVal strFormName As String)
Dim iValue As Integer

   'incremento el parametro en el registro de configuración para la estadistica de forms mas usados
   
   iValue = GetRegistryValue(HKEY_LOCAL_MACHINE, RegistrySubKeys.MRUForms, strFormName, REG_DWORD, 0, False)
   If iValue = 0 Then
      'aún no ha sido grabado
      SetRegistryValue HKEY_LOCAL_MACHINE, RegistrySubKeys.MRUForms, strFormName, REG_DWORD, 1
   Else
      SetRegistryValue HKEY_LOCAL_MACHINE, RegistrySubKeys.MRUForms, strFormName, REG_DWORD, iValue + 1
   End If

End Sub

Public Sub SetRegistryEntries(Optional ByVal strUser As String)

   '  setea la ubicación de las claves del registro de windows
      
   With RegistrySubKeys
   
      If Len(strUser) = 0 Then
         .DataBaseSettings = "Software\Algoritmo\DataBaseSettings"
         .Environment = "Software\Algoritmo\Environment"
         .NavigationQueries = "Software\Algoritmo\MRU Queries\NavigationStoredQueries"
         .GridQueries = "Software\Algoritmo\MRU Queries\GridStoredQueries"
         .PrintQueries = "Software\Algoritmo\MRU Queries\PrintStoredQueries"
         .QueryDBQueries = "Software\Algoritmo\MRU Queries\QueryDBStoredQueries"
         .DataComboQueries = "Software\Algoritmo\MRU Queries\DataComboStoredQueries"
      Else
         .MRUEmpresas = "Software\Algoritmo\MRU Empresas\" & strUser
         .MRUForms = "Software\Algoritmo\MRU Forms\" & strUser & "\" & App.ProductName
      End If
      
   End With


End Sub

Public Sub ClearForm(ByVal frm As Form)
Dim cntrl As Control

   For Each cntrl In frm.Controls
      Select Case TypeName(cntrl)
         Case "RichTextBox", "TextBox", "PowerMask"
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
      End Select
   Next cntrl

End Sub

Public Sub UnloadForm()

   'si el form pertenece al MRU, lo escondo
   If IsMRUForm(Screen.ActiveForm.hWnd) Then
      MRUForms(Screen.ActiveForm.Name).Visible = False
   Else
     Unload Screen.ActiveForm
   End If

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
         
         If iEnteros < Len(CStr(Round(Trim(oActiveControl.Text), 0))) Then
            strFormato = "+/-" & Formato("99999999999999,9999999999", iDecimales, iDimension - iDecimales)
            MsgBox "El número de enteros ingresados para este campo es demasiado grande. " & _
                   "El máximo valor permitido es " & strFormato, vbExclamation, App.ProductName
            IsValid = False
            Screen.ActiveControl.SetFocus
            Exit Function
         End If
         
      Case Fecha 'el dato debe ser una fecha válida
         If (Len(Trim(oActiveControl.Text)) > 0) And (Not IsDate(Trim(oActiveControl.Text))) Then
            MsgBox "No es una fecha válida", vbExclamation, App.ProductName
            IsValid = False
            Screen.ActiveControl.SetFocus
            Exit Function
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
            
   
   Set rst = New ADODB.Recordset
   
   ' Pregunta si Usa Archivos Temporarios
   If SystemOptions.UseLocalCopy = No Then

      ' No Usa Archivos Temporarios
      Set rst = GetData(strEmpresa, strSQL, adOpenStatic, adLockPessimistic)
      
   Else
   
      ' Usa Archivos Temporarios
      
      SQLScomposer.SQLInputString = strSQL
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

Public Sub SetContextMenu(ByVal strMenuName As String, Optional ByVal strTaskAdmin As String, Optional ByVal strModulo As String)
Dim ActiveForm As Form
Dim ActiveControl As Control
Dim ItemNumber As Integer
Dim strCaption As String
Dim bPrintSeparator As Boolean

   'Carga el menu contextual strMenuName. Para los controles de tipo TEXT o DATACOMBO además
   'agrega el menu standard para estos tipos de controles
   
   Set ActiveForm = Screen.ActiveForm
   Set ActiveControl = ActiveForm.ActiveControl

   If TypeOf ActiveControl Is TextBox Or _
      TypeOf ActiveControl Is PowerMask Or _
      TypeOf ActiveControl Is DataCombo Or _
      TypeOf ActiveControl Is DataGrid Then
   
   
      If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is PowerMask Or TypeOf ActiveControl Is DataCombo Then
         If strMenuName <> "GRILLA" Then
            ' si es un TextBox o DataCombo agrego el menu standard para estos controles
            rstContextMenu.Filter = "MNX_MODULO = 'TODOS' AND MNX_NOMBRE = 'TEXT'"
         Else
            ' es un text oculto
            rstContextMenu.Filter = "MNX_MODULO = 'TODOS' AND MNX_NOMBRE = 'GRILLA'"
         End If
      End If
      If TypeOf ActiveControl Is DataGrid Then
         ' si es una DataGrid agrego el menu standard para ese control
         rstContextMenu.Filter = "MNX_MODULO = 'TODOS' AND MNX_NOMBRE = 'GRILLA'"
      End If
      
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
         If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is PowerMask Then
            
            If rstContextMenu(ContextMenuEnum.mnxClave) = "FINDQUERYDB" Then
               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = True
            End If
            
            If rstContextMenu(ContextMenuEnum.mnxClave) = "ADMINISTRAR" Then
               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = (Len(strTaskAdmin) > 0 And (TaskIsEnabled(strTaskAdmin, CUsuario)))
            End If
            
            If rstContextMenu(ContextMenuEnum.mnxClave) = "ACTUALIZAR" Then
               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
            End If
            
         End If
         
         If TypeOf ActiveControl Is DataCombo Then
            
            If rstContextMenu(ContextMenuEnum.mnxClave) = "FINDQUERYDB" Then
               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
            End If
            
            If rstContextMenu(ContextMenuEnum.mnxClave) = "ADMINISTRAR" Then
               ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = (Len(strTaskAdmin) > 0 And (TaskIsEnabled(strTaskAdmin, CUsuario)))
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
         
         End If
         
         If rstContextMenu(ContextMenuEnum.mnxClave) = "FINDQUERYDB" Then
            'seteo el frmFind
'            Set frmFind.CalledByForm = ActiveForm
'            Set frmFind.CalledByControl = ActiveControl
            
         End If
         
         ItemNumber = 1
         
         rstContextMenu.MoveNext
      Loop
      
      If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is PowerMask Then
      
         'Creo el menu Edit del control
         AddEditContextMenu ActiveForm
         
         ItemNumber = ActiveForm.mnuContextItem.UBound

      End If
      
   End If

   'creo el resto del menu
   If Len(strModulo) > 0 Then
      rstContextMenu.Filter = "(MNX_MODULO = '" & strModulo & "' AND MNX_NOMBRE = '" & strMenuName & "') OR (MNX_MODULO = '" & UCase(App.ProductName) & "' AND MNX_NOMBRE = '" & strMenuName & "')"
   Else
      rstContextMenu.Filter = "(MNX_MODULO = '" & UCase(App.ProductName) & "' AND MNX_NOMBRE = '" & strMenuName & "')"
   End If

   If rstContextMenu.RecordCount > 0 Then
      If ItemNumber > 0 Then
         Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
         ItemNumber = ActiveForm.mnuContextItem.UBound
         ActiveForm.mnuContextItem(ItemNumber).Caption = "-"
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
         
         'si el caption termina con la secuenacia "%S" entonces trazo una linea
         If bPrintSeparator Then
         
            Load ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound + 1)
            ItemNumber = ActiveForm.mnuContextItem.UBound
            ActiveForm.mnuContextItem(ItemNumber).Tag = NullString
            ActiveForm.mnuContextItem(ItemNumber).Caption = "-"
         
         End If
         
         'salvo la clave en el tag asi el programa llamador puede interrogar dicho tag
         ActiveForm.mnuContextItem(IIf(bPrintSeparator, ItemNumber - 1, ItemNumber)).Tag = rstContextMenu(ContextMenuEnum.mnxClave)
         
         If Not IsNull(rstContextMenu(ContextMenuEnum.mnxForms)) Then
            If (Not TaskIsEnabled(rstContextMenu(ContextMenuEnum.mnxForms), CUsuario)) Then ActiveForm.mnuContextItem(ItemNumber).Enabled = False
         End If
         
      End If
      
      If ItemNumber = 0 Then ItemNumber = 1
      
      rstContextMenu.MoveNext
   Loop

   rstContextMenu.Filter = adFilterNone
   
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

Public Function ShowForm(ByVal strFormName As String, ByVal strEmpresa As String, Optional ByVal strMenuKey As String, Optional ByVal hWndParent As Long) As Long
Dim frmForm As Form
   
   '  Muestra un form.
   '  Si el form que se desea visualizar existe en el MRUForms y no esta invisible entonces
   '  hago el Show de dicho form
   '  Si el form no esta en el MRUForms o bien esta pero se requiere una nueva instancia,
   '  entonces cargo la nueva instancia
   
   On Error Resume Next
   
   Set frmForm = MRUForms(Trim(strFormName))
   If Not (frmForm Is Nothing) Then
       'es un form presente en MRUForms
      If MRUForms(Trim(strFormName)).Visible = False Then
         MRUForms(Trim(strFormName)).Empresa = strEmpresa
         MRUForms(Trim(strFormName)).MenuKey = strMenuKey
         MRUForms(Trim(strFormName)).InitForm
         MRUForms(Trim(strFormName)).Visible = True
         MRUForms(Trim(strFormName)).PostInitForm
         ShowForm = MRUForms(Trim(strFormName)).hWnd
         Exit Function
      End If
   End If
   
   Set frmForm = Forms.Add(Trim(strFormName))
   If Err Then
      MsgBox "ShowForm no ha podido abrir el form " & strFormName, vbOKOnly, App.ProductName
      Exit Function
   End If
   
   If Not frmForm.PropertlyLoaded Then Unload frmForm: Exit Function
   
   frmForm.Empresa = strEmpresa
   frmForm.MenuKey = strMenuKey
   
   frmForm.InitForm
   frmForm.Visible = True
   frmForm.PostInitForm
   
   ShowForm = frmForm.hWnd
   
   Exit Function
   
GestErr:
   LoadError "ShowForm"
   ShowErrMsg
End Function

Public Sub DataComboLoad(oCombo As DataCombo, ByVal strEmpresa As String, ByVal strTableName As String, ByVal BoundColumn As String, ByVal ListField As String)
Dim strSQL  As String

   If BoundColumn <> ListField Then
      strSQL = "SELECT DISTINCT " & BoundColumn & ", " & ListField & " FROM " & strTableName & " ORDER BY " & ListField
   Else
      strSQL = "SELECT DISTINCT " & BoundColumn & " FROM " & strTableName & " ORDER BY " & ListField
   End If

   Set oCombo.RowSource = OpenComboDataSource(strEmpresa, strSQL)
   oCombo.BoundColumn = BoundColumn
   oCombo.ListField = ListField
   
End Sub
Public Sub CenterForm(ByVal frm As Form)
  
   '--  centra el form
   
   If frm.MDIChild Then
      If frm.WindowState = vbMaximized Or frm.WindowState = vbMinimized Then Exit Sub
   End If
  
   Select Case frm.MDIChild
    Case True
       frm.Left = (Forms(0).ScaleWidth - frm.Width) / 2
       frm.Top = (Forms(0).ScaleHeight - frm.Height) / 2
    Case False
       frm.Left = (Screen.Width - frm.Width) / 2
       frm.Top = (Screen.Height - frm.Height) / 2
    End Select
   
End Sub

Public Sub SetCaption(ByVal frm As Form, ByVal strEmpresa As String)

   If InStr(frm.Caption, NombreEmpresa(strEmpresa)) > 0 Then
      frm.Caption = Left(frm.Caption, InStr(frm.Caption, "(") - 1)
   End If
   frm.Caption = frm.Caption & " (" & NombreEmpresa(strEmpresa) & ")"

End Sub


Public Function FindKey(strKey As String, TV As TreeView) As Boolean
Dim nodX As Node
  
   ' busco si existe la clave strKey en el treeview TV
   
   On Error Resume Next
   
   Set nodX = TV.Nodes(strKey)
       
   If nodX Is Nothing Then
      FindKey = False
      Exit Function
   Else
      FindKey = True
      Exit Function
   End If

End Function

Public Function TextLostFocus(ctl As Object, objBusiness As Object, strPropLet As String, Optional strPropGet As String) As String

   If Len(Trim(ctl.Text)) = 0 Then
      TextLostFocus = NullString
      Exit Function
   End If
   
'   If IsChanged(ctl) Then
      CallByName objBusiness, strPropLet, VbLet, ctl.Text
      
      If Len(strPropGet) > 0 Then
         TextLostFocus = CallByName(objBusiness, strPropGet, VbGet)
      Else
         TextLostFocus = NullString
      End If
      
'   Else
'      TextLostFocus = CallByName(objBusiness, strPropGet, VbGet)
'   End If

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
      frm.mnuContextItem(.UBound).Enabled = objTextBox.canUndo
      
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

