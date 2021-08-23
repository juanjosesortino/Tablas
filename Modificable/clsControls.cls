VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ISubclass

Private WithEvents m_TextControls As clsControlItems
Attribute m_TextControls.VB_VarHelpID = -1
Private WithEvents Hooking As AlgStdFunc.clsHooking
Attribute Hooking.VB_VarHelpID = -1
'
' mensajes subclasados
Private Const WM_ENABLE = &HA
Private Const WM_SETTEXT = &HC
Private Const WM_KEYDOWN = &H100
'Private Const WM_SETFOCUS = &H7
'Private Const WM_KILLFOCUS = &H8
Private Const WM_CHAR = &H102
Private Const WM_RBUTTONDOWN = &H204

'mensajes personalizados subclasados
Private Const DISABLE_CONTROL_EDITION As Long = &H32
Private Const ENABLE_CONTROL_EDITION As Long = &H31

Private strUsaSubClassing2009             As String
'-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
'Esta clase permite interceptar y elaborar eventos tales como :
'
'  -  KeyDown (CheckKey), Validate (IsValid), keyPress (Mayusculas y Minusculas)
'  -  Búsqueda a traves de QUERYDB
'  -  Presentacion de menues contextuales para cada control
'  -  mnuContextItem_Click, mnuEdit_Click, mnuEditItem_Click y todo lo relacionado con
'     operaciones de Edicion, Búsqueda y Administracion (típicas para TextBox,PowerMask y DataCombo)
'
'Implementa una coleccion que mantiene informacion de cada control.
'-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --

'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Function CopyStringA Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long
'Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private WithEvents tmr1           As AlgStdFunc.clsTimer       'permite selecionar en manera asíncrona un item del menu contextual
Attribute tmr1.VB_VarHelpID = -1
Private ContextTag                As String                    'Conserva el Tag del item del menu contextual
Private objTextBox                As AlgStdFunc.clsTextBoxEdit
Private hWndForm                  As Long                      'es el hwnd del form que contiene los controles

Private objStdFuntions            As AlgStdFunc.clsStdFunctions
Private sComposer                 As AlgStdFunc.clsSQLScomposer

Event Messages(ByVal lngMessage As Long, ByRef Info As Variant)

Public Enum ValidEnum
   vldNumerico = 1
   vldFecha = 2
   vldHora = 3
End Enum

Public Enum eFormatField
   [_NoDefinido] = 0
   FormatSi = 1
   FormatNo = 2
   FormatSoloNumericos = 3
End Enum


Private Type QDBInfo
   Titulo As String
   TituloColumnas As String
   AnchoColumnas As String
   FormatoColumnas As String
   sql As String
End Type

Private objTablas                As BOGeneral.clsTablas
Private objTablasValores         As BOGeneral.clsTablasValores
Private mvarControlInfo          As Variant        'variable para interar en la coleccion
Private mvarForm                 As Object         'informa a la clase el objeto Form o Usercontrol al cual pertenecen los controles
Private mvarControlData          As Variant
Private mvarfrmFind              As Form           'necesaria para informar a la clase el form Find empleado
Private mvarUsuario              As Object         'usuario
Private mvarFullControlOnTextBox As Boolean        ' Default=False;True:los eventos Validate y LostFocus no se delegan a MValidatePointer.bas
Private aTableProperties         As Variant        'propiedades de los campo de la tabla
Private rstContextMenu           As ADODB.Recordset
Private mvarCallerObjName        As String         ' Nombre del Form/Clase que llama al dialog
Private mvarIsABM                As Boolean        ' indica si el form es un ABM o no

Private ControlInfo              As clsControlInfo  'variable de tipo
Private col                      As clsControlInfos 'coleccion interna de informacion de controles

Private Enum ContextMenuEnum
   mnxNombre = 0
   mnxOrden = 1
   mnxForms = 2
   mnxCaption = 3
   mnxTarea = 4
   mnxClave = 5
End Enum

Private mvarUserControlsChanged  As Boolean

Private mError                   As ErrRec
Private ErrorLog                 As ErrType                        'información del error generado
Private Const CLASSNAME          As String = "<clsControls>"

Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Const GA_ROOT = 2

Public Property Set Usuario(ByVal vData As Object)
    Set mvarUsuario = vData
End Property

Public Property Get Usuario() As Object
    Usuario = mvarUsuario
End Property
Public Property Set FormFind(ByVal vData As Object)
    Set mvarfrmFind = vData
End Property

Public Property Set Form(ByVal vData As Object)
Dim strFiltro     As String
Dim iCount        As Integer
Dim aSplit()      As String
Dim ix            As Integer
Dim strMenuKey    As String

   Set mvarForm = vData
   
    
   On Error GoTo GestErr
   
   If TypeName(mvarForm) = "frmDialog" Then
      If Len(mvarForm.MenuKey) > 0 Then
         strMenuKey = objStdFuntions.EliminaAcentos(mvarForm.MenuKey)
      End If
   Else
      strMenuKey = objStdFuntions.EliminaAcentos(mvarForm.MenuKey)
      
   End If
   
   If Len(strMenuKey) > 0 Then
      '/ si es un form distinto al dialogo leo sus propiedades

      strFiltro = "TAB_CLAVE1 = 'Forms' and TAB_CLAVE2 = '" & strMenuKey & "'"
      
'      If Err.Number = 438 Then Exit Property 'el form no admite la propiedad mvarForm.MenuKey
      
      On Error GoTo GestErr
      objTablas.ControlData = mvarForm.ControlData
      
      Set objTablasValores = objTablas.GetFromSPM(mvarForm.ControlData.Empresa, strFiltro)
      
      iCount = objTablasValores.Count
      For ix = 1 To iCount
         If objTablasValores.Count < ix Then Exit For
         aSplit = Split(objTablasValores.Item(ix).Tabla, "\")
         Select Case aSplit(3)
            Case CUsuario.Usuario, "ValorPredeterminado", "Fuente", "Color", "Color Fondo", "Control Edicion", "Usa Valores por Defecto", "Usa Fecha del Dia"
            Case Else
               objTablasValores.Remove ix
         End Select
      Next ix
   End If
   '
   ' determino si el form es un ABM o no
   '
   mvarIsABM = IsABM()
   
   Exit Property
   
GestErr:

   MsgBox Err.Description
   
End Property

Public Property Get Form() As Object
    Set Form = mvarForm
End Property

Public Property Get Controls() As Collection
    Set Controls = col
End Property

Public Sub Remove(ctrl As Object)
   
   '- remove un control de la clase Controls significa terminar con el subclassing
   '- y eliminarlo de la coleccion de controles
   
   If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "DataCombo" Or TypeName(ctrl) = "PowerMask" Then
   
      'remuevo los mensajes
      If strUsaSubClassing2009 = si Then
         Hooking.RemoveSubClass ctrl.hWnd
      Else
         DetachAllMessages ctrl.hWnd
      End If
   End If

   On Error Resume Next

   If TypeName(ctrl) = "PowerMask" Then
      'el control PowerMask no es subclasado. Sus secciones SI
      Dim Section As PowerMaskSection
      For Each Section In ctrl.Sections
      
         'remuevo los mensajes
         If strUsaSubClassing2009 = si Then
            Hooking.RemoveSubClass Section.Control.hWnd
         Else
            DetachPmkSectionMessages Section.Control.hWnd
         End If
         'remuevo la seccion de la colecciòn
         col.Remove CStr(Section.Control.hWnd)
      Next Section
   End If

   'lo remuevo de la colección si existe
   col.Remove CStr(ctrl.hWnd)
   
End Sub

Public Sub Add( _
         ctrl As Object, _
         Optional ByVal BoundField As String, _
         Optional ByVal ObjectProperty As String, _
         Optional ByVal Validar As ValidEnum, _
         Optional ByVal FindControlsList As String, _
         Optional ByVal FindBoundFieldList As String, _
         Optional ByVal Dimension As Integer, _
         Optional ByVal Decimales As Integer, _
         Optional ByVal FindWhere As String, _
         Optional ByVal Formatear As eFormatField = [_NoDefinido], _
         Optional ByVal Mascara As String _
               )

Dim ControlInfo As clsControlInfo
Dim aSplit()    As String
Dim ix          As Integer
Dim Dummy       As Variant
Dim strIDControl As String
Dim ControlItem As clsControlItem
Dim strMenuKey  As String


   On Error GoTo GestError

   
   ' agrego los TextBox y PowerMask a la colección con eventos
   Set ControlItem = New clsControlItem
         
   If FullControlOnTextBox = False Then
    
       If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "PowerMask" Then
    
          Set ControlItem.MyControl = ctrl
    
          m_TextControls.Add ControlItem
    
       End If
   
   End If
   
   ' me salvo en una variable local la MenuKey
   'strMenuKey = mvarForm.MenuKey
   
   strMenuKey = objStdFuntions.EliminaAcentos(mvarForm.MenuKey)
   
   
   Set ControlInfo = New clsControlInfo

   With ControlInfo
      
      .ControlEdicion = True
      
      If Len(BoundField) > 0 Then
      
         'EL CONTROL ES BOUND
      
         aTableProperties = GetFieldInformation(BoundField)
         
         'setea alineación, Maxlenght y DataField del control
         'IMPORTANTE: Estas líneas deben ser las primeras porque dentro
         '           SetControlProperties se setea la propiedad Alignment y
         '           esto cambia el hWnd del control.
         
         If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "PowerMask" Then
            SetControlProperties aTableProperties, ctrl, BoundField
         End If
         
         
         'datos del campo
         .BoundField = FieldProperty(aTableProperties, BoundField, dsCampo)
         .MayuscMinusc = FieldProperty(aTableProperties, .BoundField, dsCaseMode)
         .FieldDimension = FieldProperty(aTableProperties, .BoundField, dsDimension)
         .FieldEnteros = .FieldDimension - FieldProperty(aTableProperties, .BoundField, dsDecimales)
         .FieldDecimales = FieldProperty(aTableProperties, .BoundField, dsDecimales)
         .FieldValoresPosibles = FieldProperty(aTableProperties, .BoundField, dsValores)
         .FieldType = FieldProperty(aTableProperties, .BoundField, dsTipoDato)
         .Comentarios = FieldProperty(aTableProperties, .BoundField, dsComentarios)
         
         If TypeName(ctrl) = "TextBox" Then
            If Validar = 0 Then
               Select Case .FieldType
                  Case adNumeric
                     .Validar = vldNumerico
                  Case adDBTimeStamp
                     .Validar = vldFecha
                  Case Else
                     .Validar = 0
               End Select
            Else
               .Validar = Validar
            End If
         End If
         If TypeName(ctrl) = "TextBox" Then
            .MaxLenght = FieldProperty(aTableProperties, .BoundField, dsMaxLength)
         End If
         .ObjectProperty = ObjectProperty
         'datos para el menu contextual
         .ContextMenu = FieldProperty(aTableProperties, .BoundField, dsContextMenu)
         .ContextMenu = IIf(IsNull(.ContextMenu), NullString, .ContextMenu)
         .MenuKeyAdmin = FieldProperty(aTableProperties, .BoundField, dsMenuKeyAdmin)
         .MenuKeyAdmin = IIf(IsNull(.MenuKeyAdmin), NullString, .MenuKeyAdmin)
         
         .Tag = Null
 
         'formateo
         If TypeName(ctrl) = "TextBox" Then
            Select Case .FieldType
               Case adNumeric
                  'en caso de formatear, usa siempre la mascara definida en el diccionario
                  If Formatear = [_NoDefinido] Then
                     If .MenuKeyAdmin <> NullString Then
                        'asumo no formatear
                        .Formatear = [_NoDefinido]
                     Else
                        .Formatear = FormatSi
                        .MascaraFormato = Mascara
                     End If
                  Else
                     'si o no
                     .Formatear = Formatear
                     .MascaraFormato = Mascara
                  End If
               Case adDBTimeStamp
                  .Formatear = IIf((Formatear = [_NoDefinido]) Or (Formatear = FormatSi), FormatSi, FormatNo)
               Case Else
                  If Mascara <> NullString Then
                     .Formatear = Formatear
                     .MascaraFormato = Mascara
                  Else
                     .Formatear = FormatNo
                     .MascaraFormato = NullString
                  End If
            End Select
         End If
 
         'datos para la busqueda
         .FindIsEnabled = Not Len(FieldProperty(aTableProperties, .BoundField, dsQuery)) = 0
         .AdminIsEnabled = Not Len(.MenuKeyAdmin) = 0
         .FindDataLink = BoundField
         .FindControlsList = FindControlsList
         .FindWhere = FindWhere
         'si fue definida la busqueda y no fue definido la lista FindControlsList asumo
         'que se quiere actualizar el control Ctrl
         
         On Error Resume Next                   'Ctrl puede no ser un arreglo
                 
         If .FindControlsList = NullString Then
            Dim Index As Integer
            Index = ctrl.Index
            If Err.Number = 0 Then
               'es un arreglo (o matriz )
               .FindControlsList = ctrl.Name & "(" & Index & ")"
            Else
               .FindControlsList = ctrl.Name
            End If
         End If
         On Error GoTo GestError
         
         .FindBoundFieldList = IIf(FindBoundFieldList = NullString, .BoundField, FindBoundFieldList)
         
      Else
      
         'EL CONTROL ES UNBOUND
         
         .FieldDimension = Dimension
         .FieldEnteros = .FieldDimension - Decimales
         .FieldDecimales = Decimales
         If TypeName(ctrl) = "TextBox" Then
            Select Case Validar
               Case vldNumerico
                  .Validar = vldNumerico
                  .FieldType = adNumeric
                  .MaxLenght = Len(Formato("99999999999999,9999999999", Decimales, Dimension - Decimales)) + 1
               Case vldFecha
                  .Validar = vldFecha
                  .FieldType = adDBTimeStamp
               Case vldHora
                  .Validar = vldHora
                  .FieldType = adDBTimeStamp
               Case Else
                  .Validar = 0
                  .FieldType = adVarChar
                  .MaxLenght = Dimension
            End Select
            
         End If
         
      End If
      
      .hWnd = ctrl.hWnd
      Set .ctrl = ctrl
      
      .TypeOfControl = TypeName(ctrl)
      
      If TypeName(ctrl) = "ComboBox" Then
         If Len(.FieldValoresPosibles) > 0 Then
            aSplit = Split(.FieldValoresPosibles, ";")
            ctrl.Clear
            For ix = LBound(aSplit) To UBound(aSplit)
               ctrl.AddItem aSplit(ix)
            Next ix
         End If
      End If
         
      .Changed = False
         
      '/completo los miembros .ControlEdicion y .DefaultValue
      
      
      'controlo si el control forma parte de una matriz de controles
      On Error Resume Next
      strIDControl = .ctrl.Name & .ctrl.Index
      If Err Then strIDControl = .ctrl.Name
      Err.Clear
      
      .ControlEdicion = IIf(objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\" & CUsuario.Usuario & "\Control Edicion") = si, True, False)
      If Err.Number <> 0 Then
         .ControlEdicion = IIf(objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\Control Edicion") = si, True, False)
         Err.Clear
      End If
      If CUsuario.SysAdmin Or CUsuario.UserAdmin Then
         .ControlEdicion = True
      End If
      
      Select Case .FieldType
         Case adDBTimeStamp
            
            '/ si el campo es de tipo numerico me fijo si usa la Fecha de Dia como default
            
            Dummy = NullString
         
            Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\" & CUsuario.Usuario & "\Usa Fecha del Dia")
            If Err.Number <> 0 Then
               Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\Usa Fecha del Dia")
               Err.Clear
            End If
            If Dummy = si Then
               If .Validar = vldHora Then
                  .DefaultValue = Format(Time, "HH:MM")
               Else
                  .DefaultValue = Format(Date, "dd-mm-yyyy")
               End If
            Else
               .DefaultValue = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\" & CUsuario.Usuario & "\ValorPredeterminado")
               If Err.Number <> 0 Then
                  .DefaultValue = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\ValorPredeterminado")
                  If IsEmpty(.DefaultValue) Then .DefaultValue = NullString
                  Err.Clear
               Else
               '   .DefaultValue = NullString
               End If
               
            End If
            
         Case Else
      
            .DefaultValue = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\" & CUsuario.Usuario & "\ValorPredeterminado")
            If Err.Number <> 0 Then
               .DefaultValue = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\ValorPredeterminado")
               If IsEmpty(.DefaultValue) Then .DefaultValue = NullString
               Err.Clear
            End If
      End Select
      
      RaiseEvent Messages(CTL_AFTER_SET_VALUE_DEF, ControlInfo)
      
      Dummy = NullString
      Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\" & CUsuario.Usuario & "\Usa Valores por Defecto")
      If Dummy = si Then
         'el usuario interactivo usa para este control valores por defecto
         
         .ctrl.ForeColor = BlackColor
         Select Case TypeName(.ctrl)
            Case "OptionButton", "CheckBox"
               .ctrl.BackColor = GrayColor
            Case Else
               .ctrl.BackColor = WhiteColor
         End Select
         .ctrl.Font.Name = "MS Sans Serif"
         .ctrl.Font.Size = 8
         .ctrl.Font.Bold = False
         .ctrl.Font.Italic = False
         .ctrl.Font.Strikethrough = False
         .ctrl.Font.Underline = False
         
      Else
      
         Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\" & CUsuario.Usuario & "\Fuente")
         If Err.Number <> 0 Then
            Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\Fuente")
            Err.Clear
         End If
         If Len(Dummy) <> 0 Then
            aSplit = Split(Dummy, ";")
            If Len(aSplit(0)) Then
               .ctrl.Font.Name = aSplit(0)
            End If
            If Len(aSplit(1)) Then
               .ctrl.Font.Bold = IIf(aSplit(1) = "Bold", True, False)
            End If
            If Len(aSplit(2)) Then
               .ctrl.Font.Size = aSplit(2)
            End If
            If Len(aSplit(3)) Then
               .ctrl.Font.Italic = IIf(aSplit(3) = "Italic", True, False)
            End If
            If Len(aSplit(4)) Then
               .ctrl.Font.Strikethrough = IIf(aSplit(4) = "Tachado", True, False)
            End If
            If Len(aSplit(5)) Then
               .ctrl.Font.Underline = IIf(aSplit(5) = "Subrayado", True, False)
            End If
         End If
            
         Dummy = NullString
         Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\" & CUsuario.Usuario & "\Color")
         If Err.Number <> 0 Then
            Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\Color")
            Err.Clear
         End If
         If Len(Dummy) <> 0 Then
            .ctrl.ForeColor = Dummy
         End If
         
         Dummy = NullString
         Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\" & CUsuario.Usuario & "\Color Fondo")
         If Err.Number <> 0 Then
            Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\Color Fondo")
            Err.Clear
         End If
         If Len(Dummy) <> 0 Then
            .ctrl.BackColor = Dummy
         End If
         Select Case TypeName(.ctrl)
            Case "OptionButton", "CheckBox"
               If Len(Dummy) <> 0 Then
                  .BackColor = Dummy
               Else
                  .BackColor = GrayColor
               End If
            Case Else
               If Len(Dummy) <> 0 Then
                  .BackColor = Dummy
               Else
                  .BackColor = EnabledColor
               End If
         End Select
      
      End If
      
      On Error GoTo GestError
      
      '/
         
      'lo meto en la coleccion que mantiene informacion de los controles
      col.Add ControlInfo, CStr(.hWnd)
      
      Select Case TypeName(ctrl)
         Case "TextBox", "DataCombo", "PowerMask", "ComboBox", "OptionButton", "CheckBox", "RichTextBox"
         
            'informo los mensajes que serán subclasados
            If strUsaSubClassing2009 = si Then
               Hooking.AddSubClassing .hWnd
            Else
               AttachAllMessages .hWnd
            End If
         Case Else
            MsgBox "La clase " & TypeName(ctrl) & " no es aceptada por la clase clsControls", vbInformation, App.ProductName
            Exit Sub
      End Select
      
      If TypeName(ctrl) = "PowerMask" Then
         
         Dim Section As PowerMaskSection
         For Each Section In ctrl.Sections
         
            'informo los mensajes que serán subclasados
            If strUsaSubClassing2009 = si Then
               Hooking.AddSubClassing Section.Control.hWnd
            Else
               AttachPmkSectionMessages Section.Control.hWnd
            End If
            'Las secciones son colocadas en la colección simplemente para tener accesso
            'al al informacion del control "PowerMask" que las contiene
            
            Dim SectionInfo As clsControlInfo
            
            Set SectionInfo = New clsControlInfo
            
            With SectionInfo
               Set .ctrl = Section.Control
               .TypeOfControl = "PowerMaskSection"
               .hWnd = Section.Control.hWnd
               .hWndParent = ctrl.hWnd
               .FindIsEnabled = ControlInfo.FindIsEnabled
               .AdminIsEnabled = ControlInfo.AdminIsEnabled
               
               .BackColor = ControlInfo.BackColor
               .ControlEdicion = ControlInfo.ControlEdicion
               
               col.Add SectionInfo, CStr(.hWnd)
            End With
            
         Next Section
      End If
      
      If hWndForm = 0 Then
         'obtengo el hwnd del Form
         hWndForm = GetFormHwnd(.hWnd)
      End If
      
      If Not .ControlEdicion Then .ctrl.Enabled = False
      
   End With
   
   Exit Sub

GestError:
   If Err.Number = 438 Then Resume Next
   
   SetError mError, CLASSNAME, "Add"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Sub

Private Sub Class_Initialize()
   Set col = New clsControlInfos
   Set tmr1 = New AlgStdFunc.clsTimer
   
   Set objTablas = New BOGeneral.clsTablas
   Set objTablasValores = New BOGeneral.clsTablasValores
   
   Set sComposer = New AlgStdFunc.clsSQLScomposer
   
   Set objStdFuntions = New AlgStdFunc.clsStdFunctions
   
   Set rstContextMenu = ContextMenu
   
   Set m_TextControls = New clsControlItems
   
   'vs113sp
   Set objTextBox = New AlgStdFunc.clsTextBoxEdit
   'Si no tiene actualizado el dataShared.SPM
   
'   strUsaSubClassing2009 = UsaSubclassing2009 'Inc. 58776
   strUsaSubClassing2009 = RegistrySubKeys.KeyMRUForms
   
   If strUsaSubClassing2009 = si Then
      Set Hooking = New AlgStdFunc.clsHooking
      'DEFINO LISTA DE MENSAGES QUE PASAN POR CONSUME
      Hooking.ListMsgConsume = ";" & CStr(WM_CHAR) & ";" & CStr(WM_RBUTTONDOWN) & ";"
   End If
   
End Sub

Private Sub Class_Terminate()
   
   On Error GoTo GestErr
   
   If Not mvarfrmFind Is Nothing Then
      Unload mvarfrmFind
   End If
   
   '-- paro el subclasing de los controles
   If strUsaSubClassing2009 = si Then
      Set Hooking = Nothing
   Else
      If Not col Is Nothing Then
         If col.Count > 0 Then
            For Each mvarControlInfo In col
               DetachAllMessages mvarControlInfo.hWnd
            Next mvarControlInfo
         End If
      End If
   End If
   
   Set col = Nothing
   Set mvarForm = Nothing
   Set mvarfrmFind = Nothing
   Set tmr1 = Nothing
   Set m_TextControls = Nothing
   
   'vs113sp
   Set sComposer = Nothing
   Set rstContextMenu = Nothing
   Set objTextBox = Nothing
   Set mvarForm = Nothing
   Set mvarfrmFind = Nothing
   Set mvarUsuario = Nothing
   Set objStdFuntions = Nothing

   Exit Sub

GestErr:
   LoadError ErrorLog, "Class_Terminate"
   ShowErrMsg ErrorLog
   
End Sub

Public Function IsChanged(ByVal hWnd As Long) As Boolean

   On Error Resume Next
   
   ' si el "Form" pierde el foco devuelvo siempre false para evitar
   ' cualquier LostFocus dentro del form
   If Screen.ActiveForm.hWnd <> mvarForm.hWnd Then
      IsChanged = False: Exit Function
   End If
   
   Set ControlInfo = col.Item(CStr(hWnd))
   
   If Err.Number <> 0 Then
      MsgBox "El hwnd: " & hWnd & " no es un elemento de la clase clsControls", vbInformation, App.ProductName
      Exit Function
   End If
   
   Err.Clear
   
  Select Case ControlInfo.TypeOfControl

  ' una vez averiguado si cambio o no, lo pongo en false

   Case "TextBox"
      IsChanged = AlgStdFunc.IsChanged(hWnd)
   Case "PowerMask"
      'IsChanged = AlgStdFunc.IsChanged(hWnd)
      IsChanged = ControlInfo.ctrl.IsChanged()
  End Select
   
   
End Function

Private Function IsValid(Optional ByVal DataType As ValidEnum) As Boolean
Dim oActiveControl      As Control
'Dim iDimension          As Integer
'Dim iEnteros            As Integer
'Dim iDecimales          As Integer
Dim strFormato          As String
Dim ControlInfo         As clsControlInfo

   '--  determina la validez del dato para el control activo

   IsValid = True

   Set oActiveControl = mvarForm.ActiveControl

   If DataType = 0 Then DataType = vldNumerico

   On Error Resume Next
   If Len(Trim(oActiveControl.Text)) = 0 Then Exit Function
   If Err Then Exit Function
   
   Set ControlInfo = col.Item(CStr(oActiveControl.hWnd))
   
   If Err.Number <> 0 Then
      MsgBox oActiveControl.Name & " no es un elemento de la clase clsControls", vbInformation, App.ProductName
      Exit Function
   End If
   
   Err.Clear
   
   
   Select Case DataType
      Case vldNumerico 'el dato debe ser numèrico
         If (Not IsNumeric(Trim(oActiveControl.Text))) Then
            MsgBox "Valor no válido para este campo", vbExclamation, App.ProductName
            IsValid = False
            Exit Function
         End If
         
         'controlo si el numero de enteros y decimales ingresados conrresponde al definido en el diccionario
         If Len(oActiveControl.DataField) = 0 Then Exit Function
         
         If ControlInfo.FieldEnteros < Len(CStr(Abs(xRound(Trim(oActiveControl.Text), 0)))) Then
            strFormato = "+/-" & Formato("99999999999999,9999999999", ControlInfo.FieldDecimales, ControlInfo.FieldEnteros)
            MsgBox "El número de enteros ingresados para este campo es demasiado grande. El máximo valor permitido es " & strFormato, vbExclamation, App.ProductName
            IsValid = False
            Exit Function
         End If
         
      Case vldFecha 'el dato debe ser una fecha válida
'         If (Len(Trim(oActiveControl.Text)) > 0) And (Not IsDate(Trim(oActiveControl.Text))) Then
'            MsgBox "No es una fecha válida", vbExclamation, App.ProductName
'            IsValid = False
'            Exit Function
'         End If

         If (Len(Trim(oActiveControl.Text)) > 0) Then
            On Error Resume Next
            Dim dFecha As Date
            dFecha = DateValue(oActiveControl.Text)
            If Err Or UCase(CStr(dFecha)) = UCase("12:00:00 a.m.") Then
               MsgBox "No es una fecha válida", vbExclamation, App.ProductName
               IsValid = False
               Exit Function
            End If
         End If

      Case vldHora 'el dato debe ser una hora válida
         If (Len(Trim(oActiveControl.Text)) > 0) Then
            On Error Resume Next
            Dim dHora As Date
            dHora = TimeValue(oActiveControl.Text)
            If Err Then
               MsgBox ("No es una hora válida")
               IsValid = False
               Exit Function
            End If
         End If
   End Select
   
End Function

Public Sub mnuContextItem_Click(Index As Integer)

   ContextTag = mvarForm.mnuContextItem(Index).Tag
   tmr1.StartTimer 100
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MenuContextTag
' DateTime  : 02/07/2003 15:54
' Author    : tony
' Purpose   : a veces es el mismo form el que le setea el ContextTag
'---------------------------------------------------------------------------------------
'
Public Property Let MenuContextTag(vData As String)

   ContextTag = vData
   tmr1.StartTimer 100
   
End Property

Public Sub mnuToolItem_Click(Index As Integer)
   RaiseEvent Messages(CTL_CALL_TOOLS, Index)
End Sub
Public Sub mnuEdit_Click()

   If TypeName(mvarForm.ActiveControl) = "TextBox" Or TypeName(mvarForm.ActiveControl) = "PowerMask" Then
   
      objTextBox = mvarForm.ActiveControl
      
      With objTextBox
      
         'le informo al objeto cual es el TextBox que debe controlar
         .TextBox = mvarForm.ActiveControl
      
         'activo/desactivo los items según corresponda
         mvarForm.mnuEditItems(MNU_EDIT_DESHACER).Enabled = .CanUndo
   
         mvarForm.mnuEditItems(MNU_EDIT_CORTAR).Enabled = .CanCut
         mvarForm.mnuEditItems(MNU_EDIT_COPIAR).Enabled = .CanCopy
         mvarForm.mnuEditItems(MNU_EDIT_PEGAR).Enabled = .CanPaste
         mvarForm.mnuEditItems(MNU_EDIT_BORRAR).Enabled = .CanDelete
         mvarForm.mnuEditItems(MNU_EDIT_TODO).Enabled = .CanSelectAll
         
      End With
      
      'verifico si esta habilitada la búsqueda
      mvarForm.mnuEditItems(MNU_EDIT_BUSCAR).Enabled = CanSearch
   Else
   
      'desactivo los items para la edicion
      mvarForm.mnuEditItems(MNU_EDIT_DESHACER).Enabled = False

      mvarForm.mnuEditItems(MNU_EDIT_CORTAR).Enabled = False
      mvarForm.mnuEditItems(MNU_EDIT_COPIAR).Enabled = False
      mvarForm.mnuEditItems(MNU_EDIT_PEGAR).Enabled = False
      mvarForm.mnuEditItems(MNU_EDIT_BORRAR).Enabled = False
      mvarForm.mnuEditItems(MNU_EDIT_TODO).Enabled = False
   
      mvarForm.mnuEditItems(MNU_EDIT_BUSCAR).Enabled = False
   End If
   
   'verifico si esta habilitada la Administracion
   mvarForm.mnuEditItems(MNU_EDIT_ADMINISTRAR).Enabled = ((TypeName(mvarForm.ActiveControl) = "TextBox" Or TypeName(mvarForm.ActiveControl) = "PowerMask" Or TypeName(mvarForm.ActiveControl) = "DataCombo") And CanAdmin)
   
   'verifico si esta habilitada la Actualización
   mvarForm.mnuEditItems(MNU_EDIT_ACTUALIZAR).Enabled = (TypeName(mvarForm.ActiveControl) = "DataCombo")
   
   If TypeName(mvarForm.ActiveControl) = "RichTextBox" Then
      If Len(mvarForm.ActiveControl.SelText) > 0 Then mvarForm.mnuEditItems(MNU_EDIT_COPIAR).Enabled = True: mvarForm.mnuEditItems(MNU_EDIT_CORTAR).Enabled = True: mvarForm.mnuEditItems(MNU_EDIT_BORRAR).Enabled = True
      If Len(Clipboard.GetText) > 0 Then mvarForm.mnuEditItems(MNU_EDIT_PEGAR).Enabled = True
      mvarForm.mnuEditItems(MNU_EDIT_BUSCAR).Enabled = False: mvarForm.mnuEditItems(MNU_EDIT_ADMINISTRAR).Enabled = False: mvarForm.mnuEditItems(MNU_EDIT_ACTUALIZAR).Enabled = False
   End If
   
End Sub

Public Sub mnuEditItem_Click(Index As Integer)

   Select Case Index
   
      Case MNU_EDIT_DESHACER
         objTextBox.Undo
      Case MNU_EDIT_CORTAR
         objTextBox.Cut
      Case MNU_EDIT_COPIAR
         objTextBox.Copy
      Case MNU_EDIT_PEGAR
         objTextBox.Paste
      Case MNU_EDIT_BORRAR
         objTextBox.Delete
      Case MNU_EDIT_TODO
         objTextBox.SelectAll
   
      Case MNU_EDIT_BUSCAR
         FindQueryDB
      Case MNU_EDIT_ADMINISTRAR
         CallFormAdmin
      Case MNU_EDIT_ACTUALIZAR
      
         On Error Resume Next
         Dim Index2 As Integer
         Index2 = mvarForm.ActiveControl.Index
         If Err.Number = 0 Then
            'es un arreglo de controles
            mvarForm.SetComboList UCase(mvarForm.ActiveControl.Name) & ";" & Index2
         Else
            mvarForm.SetComboList UCase(mvarForm.ActiveControl.Name)
         End If
         
         Err.Clear
         On Error GoTo 0
         
   End Select
   
End Sub

Private Function CanAdmin() As Boolean
Dim ControlInfo   As clsControlInfo
Dim strClaveMenu  As String

   '--   Determina si el control activo esta habilitado para la Administracion y ademas
   '--   el usuario esta habilitado para la Administracion del campo
   On Error Resume Next
   Set ControlInfo = col.Item(CStr(mvarForm.ActiveControl.hWnd))
   If Err Then CanAdmin = False: Exit Function
         
   If Not ControlInfo.AdminIsEnabled Then
      'no esta habilitada la administracion
      CanAdmin = False
   Else
      If CUsuario.SysAdmin Or CUsuario.UserAdmin Then
         CanAdmin = True
      Else
         'verifico que el "usuario" este habilitado
         strClaveMenu = CUsuario.TareasAsignadas(ControlInfo.MenuKeyAdmin).Menu
         If strClaveMenu <> NullString Then
            CanAdmin = (Len(ControlInfo.MenuKeyAdmin) > 0 And (TaskIsEnabled(IIf(IsNull(rstMenu("MNU_CLAVE").Value), "", rstMenu("MNU_CLAVE").Value), mvarUsuario)))
         Else
            CanAdmin = False
         End If
      End If
   End If
   
End Function

Private Function CanSearch() As Boolean
Dim ControlInfo As clsControlInfo

   '-- Determina si el control activo esta habilitado para el Buscar. Si tiene definido una SQL
   On Error Resume Next
   Set ControlInfo = col.Item(CStr(mvarForm.ActiveControl.hWnd))
   If Err Then CanSearch = False: Exit Function
   
   CanSearch = ControlInfo.FindIsEnabled
   
End Function

Public Sub FindQueryDB()
      Dim ix                  As Integer
      Dim oCntl               As Control
      Dim oControl            As Control
      Dim fldBoundName        As String
'      Dim strKey              As String
      Dim aControlsNames()    As String
      Dim aControls()         As Control
      Dim aBoundField()       As String
      Dim bFindInGrid         As Boolean
      Dim iColumn             As Integer
      Dim ControlInfoSection  As clsControlInfo

10       On Error GoTo GestError

         '-- Esta Sub visualiza una QUERY. Utiliza la informacion del diccionario para visualizar
         '-- la SQL asociada al control activo

         'busco la posición del elemento buscado
         
20       On Error Resume Next
30       Set ControlInfo = col.Item(CStr(mvarForm.ActiveControl.hWnd))
40       If Err Then Exit Sub 'el control no admite busqueda

50       On Error GoTo GestError
         
60       RaiseEvent Messages(CTL_BEFORE_QUERYDB, ControlInfo)

70       Set oControl = mvarForm.ActiveControl
         
80       bFindInGrid = (TypeName(oControl) = "DataGrid")
         
90       aControlsNames = Split(ControlInfo.FindControlsList, ";")
100      aBoundField = Split(ControlInfo.FindBoundFieldList, ";")
         
         'completo el arreglo aControls con los controles que entran en juego
110      ReDim aControls(UBound(aControlsNames))
120      For ix = LBound(aControlsNames) To UBound(aControlsNames)
130         Select Case TypeName(mvarForm)
               Case "frmDialog"
                  'uso el propio GetControl
                  'si el control del dialog es un datracombo
                  'aplico el get con el contenido del tag ya que la funcion GetControl busca por el tag
                  If TypeName(oControl) = "DataCombo" Then
                     Set aControls(ix) = mvarForm.GetControl(oControl.Tag)
                  Else
                     Set aControls(ix) = mvarForm.GetControl(aControlsNames(ix))
                  End If
                  
150            Case Else
160               Set aControls(ix) = GetControl(mvarForm, aControlsNames(ix))
170         End Select
180      Next ix
            
         '-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
         '  obtengo la empresa de la siguiente manera: Si la propiedad Empresa no fue definida
         '  - Raiseo un evento al form/UserControl.
         '  - Luego la busco en la propiedad del form/UserControl
         '-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
190      If IsEmpty(mvarControlData) Then
200         RaiseEvent Messages(CTL_QUERY_CONTROLDATA, mvarControlData)
210      End If
220      If IsEmpty(mvarControlData) Then
230         mvarControlData = mvarForm.ControlData
240      End If
250      If IsEmpty(mvarControlData) Then                'en los controles, mvarControlData es un variant
260         Exit Sub
270      End If
280      If mvarControlData.Empresa = NullString Then    'en los forms, mvarControlData es un udt
290         MsgBox "En " & CLASSNAME & " ControlData no disponible"
300         Exit Sub
310      End If
               
320      Set mvarfrmFind = New frmFind
          
         'refreso los datos para la visualizacion
         
330      With mvarfrmFind
340         .ControlData = mvarControlData
            
            '
            ' si tiene una QDB alternativa, tiene prioridad sobre
            ' la definida en el diccionario de datos
            '
350         If Len(ControlInfo.FindAlternate.sql) > 0 Then
            
360            .sql = ControlInfo.FindAlternate.sql
370            .Titulo = ControlInfo.FindAlternate.Titulo
380            .TituloColumnas = ControlInfo.FindAlternate.TituloColumnas
390            .AnchoColumnas = ControlInfo.FindAlternate.AnchoColumnas
400            .FormatoColumnas = ControlInfo.FindAlternate.FormatoColumnas
410         Else
420            .TableField = ControlInfo.FindDataLink
               
430            If TypeName(oControl) = "PowerMask" Then
440               If FullControlOnTextBox = True Then
450                  For ix = 1 To oControl.Sections.Count
460                     Set ControlInfoSection = col.Item(CStr(oControl.Sections(ix).hWnd))
470                     If Len(ControlInfoSection.FindWhere) > 0 Then
480                        .Where = ControlInfoSection.FindWhere
490                        Exit For
500                     End If
510                  Next ix
520               Else
530                  .Where = ControlInfo.FindWhere
540               End If
550            Else
560               .Where = ControlInfo.FindWhere
570            End If
580         End If
            '
            ' cargo los datos en la grilla
            '
590         .LoadData
            
600      End With

610      If mvarfrmFind.Error <> "Error" Then
         
            'no se produjo ningun error en el control, entonces puedo mostrarlo
620         mvarfrmFind.Show vbModal

630         With mvarfrmFind.QueryDB1
            
640            If .RecordSelected Is Nothing Then Exit Sub
               
               'controlo si el usuario hizo alguna selección
650            If .RecordSelected.State = adStateOpen Then
                  
660               For ix = LBound(aControls) To UBound(aControls)
670                  Set oCntl = aControls(ix)
                                    
680                  DoEvents
                     
690                  fldBoundName = aBoundField(ix)
                  
700                  Select Case TypeName(oCntl)
                        Case "DataGrid"
710                        oCntl.col = iColumn
720                        oCntl.Text = .RecordSelected(fldBoundName)
                           
730                        Set ControlInfo = col.Item(CStr(oCntl.hWnd))
740                        If ControlInfo.FieldType = adNumeric Then oCntl.Text = FormatControl(oCntl, oCntl.Text)
                           
750                        oCntl.EditActive = True
760                     Case "TextBox", "PowerMask"
770                        If Not IsNull(.RecordSelected(fldBoundName)) And Len(.RecordSelected(fldBoundName)) > 0 Then
                           
780                           oCntl.Text = .RecordSelected(fldBoundName)
                              
790                        End If
                           
800                        If TypeName(oCntl) = "PowerMask" Then
810                           oCntl.SetModify True
820                        Else
830                           SetModify oCntl.hWnd, True
840                        End If
                           
850                     Case "DataCombo"
860                        If Not IsNull(.RecordSelected(fldBoundName)) And Len(.RecordSelected(fldBoundName)) > 0 Then
                           
870                           oCntl.BoundText = .RecordSelected(fldBoundName)
                              
880                        End If

                        Case "ComboBox"
                           If Not IsNull(.RecordSelected(fldBoundName)) And Len(.RecordSelected(fldBoundName)) > 0 Then
                              oCntl.ListIndex = ComboSearch(oCntl, .RecordSelected(fldBoundName))
                           Else
                              oCntl.ListIndex = -1
                           End If
                           
890                  End Select
                     
900                  SetFocus oCntl.hWnd
                     
910                  DoEvents
               
920               Next ix
                        
930               oControl.SetFocus
                  
940               If TypeName(oControl) = "PowerMask" Then
                     '
                     '  para abandonar el PowerMask necesito enviar algunos "{TAB}"
                     '
950                  For ix = oControl.ActiveSection.Index To oControl.Sections.Count - 1
960                     SendKeys "{TAB}"
970                  Next ix
                                    
980               Else
990                  SendKeys "{TAB}"
1000              End If

1010              RaiseEvent Messages(CTL_QUERYDB_RECORD_SELECTED, .RecordSelected)

1020           End If
               
1030        End With
            
1040     End If
         
1050     Unload mvarfrmFind
1060     Set mvarfrmFind = Nothing
         
1070     Exit Sub

GestError:
1080     If Err.Number = 3265 Then
1090        ErrorLog.Descripcion = fldBoundName & " es inválido para esta consulta. Verificar datos SetControls o Diccionario de Datos"
1100        ErrorLog.source = KNOWN_ERRORSOURCE & "FindQueryDB"
1110        ErrorLog.NumError = vbObjectError + Err.Number
1120     End If
         
1130     LoadError ErrorLog, "FindQueryDB(" & Erl & ")"
1140     ShowErrMsg ErrorLog

End Sub

Public Sub CallFormAdmin()

   '-- llamo al Form para la administracion del control activo
   On Error Resume Next
   
   Set ControlInfo = col.Item(CStr(mvarForm.ActiveControl.hWnd))
   
   If Err.Number <> 0 Then
      MsgBox mvarForm.ActiveControl.Name & " no es un elemento de la clase clsControls", vbInformation, App.ProductName
      Exit Sub
   End If
   
   Err.Clear
   
   RaiseEvent Messages(CTL_CALL_ADMIN, ControlInfo)
   
End Sub

Public Function FormatControl(ctrl As Object, Optional ByVal vData As Variant, _
                              Optional ByVal BlankWhenZero As Boolean = True, _
                              Optional ByVal Mascara As String) As String
                              
'Dim strFormato    As String
   '-- devuelve un dato numerico/fecha formateado segun el DataField del control

   On Error GoTo GestErr

   If IsMissing(vData) Then
      vData = ctrl.Text
   End If
   
   On Error Resume Next
   
   Set ControlInfo = col.Item(CStr(ctrl.hWnd))
   
   If Err.Number <> 0 Then
      MsgBox ctrl.Name & " no es un elemento de la clase clsControls", vbInformation, App.ProductName
      Exit Function
   End If
   
   On Error GoTo GestErr
   
   If TypeName(ctrl) = "PowerMask" Then
      Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, "FormatControl no es aplicable a campos PowerMask"
'      Exit Function
   End If

   'Si es un combo devuelvo el ListIndex
   If TypeName(ctrl) = "ComboBox" Then
      FormatControl = ComboSearch(ctrl, vData)
      Exit Function
   End If
   
   If TypeName(ctrl) = "TextBox" Then
      If ControlInfo.FieldType = adNumeric Then
         If Len(Mascara) = 0 Then
            If Len(ControlInfo.MascaraFormato) = 0 Then
               FormatControl = Formato(vData, ControlInfo.FieldDecimales, ControlInfo.MaxLenght)
            Else
               FormatControl = FormatData(vData, ControlInfo.MascaraFormato)
            End If
         Else
            FormatControl = FormatData(vData, Mascara)
         End If
      End If
      If ControlInfo.FieldType = adDBTimeStamp Then
         If ControlInfo.Validar = vldHora Then
            vData = IIf(Trim(vData) = NullString, 0, vData)
            On Error Resume Next
            FormatControl = IIf(vData = 0, NullString, FormatDateTime(vData, vbShortTime))
            On Error GoTo GestErr
         Else
            If UCase(CStr(vData)) = UCase("12:00:00 a.m.") Then
               FormatControl = NullString
            Else
               FormatControl = IIf(vData = 0, NullString, Format(vData, , , vbUseSystem))
            End If
         End If
      End If
      If ControlInfo.FieldType = adVarChar Or ControlInfo.FieldType = adChar Then
         If Len(Mascara) = 0 Then
            If Len(ControlInfo.MascaraFormato) = 0 Then
               FormatControl = vData
            Else
               FormatControl = IIf(vData = NullString, NullString, FormatData(vData, ControlInfo.MascaraFormato))
            End If
         Else
            FormatControl = IIf(vData = NullString, NullString, FormatData(vData, Mascara))
         End If
      End If
   
      If BlankWhenZero Then
         If Not FindChars(FormatControl, "123456789") Then
            FormatControl = NullString
         End If
      End If
      
   Else
   
      FormatControl = vData
      
   End If

   Exit Function

GestErr:
   SetError mError, CLASSNAME, "FormatControl"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
   
End Function

Public Function ControlValue(ByVal ctrl As Object, Optional ByVal vData As Variant) As Variant
      Dim ControlInfo As clsControlInfo

         '-- devuelve un dato numerico/fecha desformateado segun el DataField del control

10       On Error Resume Next
   
20       Set ControlInfo = col.Item(CStr(ctrl.hWnd))
30       If Err.Number <> 0 Then
40          MsgBox ctrl.Name & " no es un elemento de la clase clsControls", vbInformation, App.ProductName
50          Exit Function
60       End If
   
70       On Error GoTo GestErr
   
80       If InStr("TextBox DataCombo ComboBox PowerMask ", TypeName(ctrl)) = 0 Then
90          Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, "ControlValue no aplicable controles de tipo " & TypeName(ctrl)
      '      Exit Function
100      End If
   
110      If TypeName(ctrl) = "TextBox" Or _
            TypeName(ctrl) = "ComboBox" Or _
            TypeName(ctrl) = "PowerMask" Then
   
120         Select Case ControlInfo.FieldType
               Case adNumeric
130               Select Case True
                     Case ControlInfo.FieldDecimales = 0 And _
                          ControlInfo.FieldEnteros <= 4
140                     If Not (IsMissing(vData)) Then
150                        ControlValue = CToNumeric(Trim(vData), Entero)
160                     Else
170                        ControlValue = CToNumeric(Trim(ctrl.Text), Entero)
180                     End If
190                  Case ControlInfo.FieldDecimales = 0 And _
                          ControlInfo.FieldEnteros >= 5 And ControlInfo.FieldEnteros <= 9
200                     If Not (IsMissing(vData)) Then
210                        ControlValue = CToNumeric(vData, EnteroLargo)
220                     Else
230                        ControlValue = CToNumeric(ctrl, EnteroLargo)
240                     End If
250                  Case ControlInfo.FieldDecimales = 0 And _
                          ControlInfo.FieldEnteros >= 10
260                     If Not (IsMissing(vData)) Then
270                        ControlValue = CToNumeric(vData, DoblePrecision)
280                     Else
290                        ControlValue = CToNumeric(ctrl, DoblePrecision)
300                     End If
                        ' Comentamos esto el 19/12/2008 -  Inc 29214
      '               Case ControlInfo.FieldDecimales >= 0 And ControlInfo.FieldDecimales <= 4 And _
      '                    ControlInfo.FieldEnteros <= 3
      '                  If Not (IsMissing(vData)) Then
      '                     ControlValue = CToNumeric(vData, SimplePrecision)
      '                  Else
      '                     ControlValue = CToNumeric(Ctrl, SimplePrecision)
      '                  End If
310                  Case Else
320                     If Not (IsMissing(vData)) Then
330                        ControlValue = CToNumeric(vData, DoblePrecision)
340                     Else
350                        ControlValue = CToNumeric(ctrl, DoblePrecision)
360                     End If
370               End Select
      
380            Case adDBTimeStamp
         
390               If Not (IsMissing(vData)) Then
400                  If Len(Trim(vData)) = 0 Then
410                     ControlValue = 0
420                  Else
430                     ControlValue = CDate(Trim(vData))
440                  End If
               
450               Else
460                  If Len(Trim(ctrl.Text)) = 0 Then
470                     ControlValue = 0
480                  Else
490                     ControlValue = CDate(Trim(ctrl.Text))
500                  End If
510               End If
            
520            Case Else
530               ControlValue = ctrl.Text
540         End Select
      
      
550      End If
      
560      If TypeName(ctrl) = "DataCombo" Then
      
570         Select Case ControlInfo.FieldType
      
               Case adNumeric
580               Select Case True
                     Case ControlInfo.FieldDecimales = 0 And _
                          ControlInfo.FieldEnteros <= 4
590                     If Not (IsMissing(vData)) Then
600                        ControlValue = CToNumeric(vData, Entero)
610                     Else
620                        ControlValue = CToNumeric(ctrl.BoundText, Entero)
630                     End If
640                  Case ControlInfo.FieldDecimales = 0 And _
                          ControlInfo.FieldEnteros >= 5 And ControlInfo.FieldEnteros <= 9
650                     If Not (IsMissing(vData)) Then
660                        ControlValue = CToNumeric(vData, EnteroLargo)
670                     Else
680                        ControlValue = CToNumeric(ctrl.BoundText, EnteroLargo)
690                     End If
700                  Case ControlInfo.FieldDecimales = 0 And _
                          ControlInfo.FieldEnteros >= 10
710                     If Not (IsMissing(vData)) Then
720                        ControlValue = CToNumeric(vData, DoblePrecision)
730                     Else
740                        ControlValue = CToNumeric(ctrl.BoundText, DoblePrecision)
750                     End If
760                  Case ControlInfo.FieldDecimales >= 0 And ControlInfo.FieldDecimales <= 4 And _
                          ControlInfo.FieldEnteros <= 3
770                     If Not (IsMissing(vData)) Then
780                        ControlValue = CToNumeric(vData, SimplePrecision)
790                     Else
800                        ControlValue = CToNumeric(ctrl.BoundText, SimplePrecision)
810                     End If
820                  Case Else
830                     If Not (IsMissing(vData)) Then
840                        ControlValue = CToNumeric(vData, DoblePrecision)
850                     Else
860                        ControlValue = CToNumeric(ctrl.BoundText, DoblePrecision)
870                     End If
880               End Select
      
890            Case adChar, adVarChar
      
900               If Not (IsMissing(vData)) Then
910                  ControlValue = vData
920               Else
930                  ControlValue = ctrl.BoundText
940               End If
            
950         End Select
         
960      End If

970      Exit Function

GestErr:
980      SetError mError, CLASSNAME, "ControlValue" & " - " & ctrl.Name
990      Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Function

Private Sub SetError(ByRef ErrLog As ErrRec, ByVal strClassName As String, ByVal strSource As String)
   
   With ErrLog
   
      .Num = Err.Number
      .source = Err.source
      
      If InStr(.source, KNOWN_ERRORSOURCE) = 0 Then
         If InStr(.source, UNKNOWN_ERRORSOURCE) = 0 Then
            .source = UNKNOWN_ERRORSOURCE & vbCrLf & .source
         End If
      End If
      
      If InStr(.source, strClassName) > 0 Then
         .source = .source & "[" & strSource & "]"
      Else
         .source = .source & vbCrLf & strClassName & "[" & strSource & "]"
      End If
      
      .Description = IIf(.Description <> NullString, .Description, Err.Description)
      
   End With

End Sub

Public Sub SetContextMenu(ByVal strMenuName As String, Optional bCreateDefaultMenu As Boolean = True)
                           
Dim ActiveForm As Object
Dim ActiveControl As Control
Dim ItemNumber As Integer
Dim strCaption As String
Dim bPrintSeparator As Boolean

   'Carga el menu contextual strMenuName. Para los controles de tipo TEXT o DATACOMBO además
   'agrega el menu standard para estos tipos de controles
   
   On Error GoTo GestErr

   Set ActiveForm = mvarForm
   Set ActiveControl = ActiveForm.ActiveControl

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
            If (Not TaskIsEnabled(rstContextMenu(ContextMenuEnum.mnxForms), mvarUsuario)) Then ActiveForm.mnuContextItem(ItemNumber).Enabled = False
         End If
         
      End If
      
      If ItemNumber = 0 Then ItemNumber = 1
      
      rstContextMenu.MoveNext
   Loop

   rstContextMenu.Filter = adFilterNone


   'Creo el resto del menú según el tipo de control
   If bCreateDefaultMenu Then
      If TypeName(ActiveControl) = "TextBox" Or _
         TypeName(ActiveControl) = "PowerMask" Or _
         TypeName(ActiveControl) = "DataCombo" Or _
         TypeName(ActiveControl) = "DataGrid" Or _
         TypeName(ActiveControl) = "ComboBox" Or _
         TypeName(ActiveControl) = "OptionButton" Or _
         TypeName(ActiveControl) = "CheckBox" Then
'         TypeName(ActiveControl) = "CRViewer" Then
      
      
         Select Case TypeName(ActiveControl)
'            Case "CRViewer"
               ' si es una CRViewer agrego el menu standard para ese control
'               rstContextMenu.Filter = "MNX_NOMBRE = 'ZOOM_REPORT'"
               
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
                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = CanSearch
               End If
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "ADMINISTRAR" Then
                  If mvarForm.Name = "frmCustomizeControl" Or mvarCallerObjName = "frmCustomizeControl" Then
                     mvarForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
                  Else
                     ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = CanAdmin
                  End If
               End If
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "ACTUALIZAR" Then
                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
               End If
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "CUSTOMIZE" Then
                  If mvarForm.Name = "frmCustomizeControl" Or mvarCallerObjName = "frmCustomizeControl" Then
                     mvarForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
                  End If
               End If
               
            End If
            
            If TypeName(ActiveControl) = "DataCombo" Then
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "FINDQUERYDB" Then
                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = False
               End If
               
               If rstContextMenu(ContextMenuEnum.mnxClave) = "ADMINISTRAR" Then
                  ActiveForm.mnuContextItem(ActiveForm.mnuContextItem.UBound).Enabled = CanAdmin
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
   End If

   Exit Sub

GestErr:
   SetError mError, CLASSNAME, "SetContextMenu"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Sub

Public Sub UnloadContextMenu(ByVal frm As Object)
Dim ix As Integer
   
   'Descarga el menu contextual activo del form pasado como argumento
   
   For ix = frm.mnuContextItem.UBound To 1 Step -1
      Unload frm.mnuContextItem(ix)
   Next ix
   frm.mnuContextItem(0).Caption = NullString
   frm.mnuContextItem(0).Enabled = True
   
End Sub

Private Sub AddEditContextMenu(ByVal frm As Object)

   '-- agrea los comandos de edición al menu contextual del form

   'le informo al objeto cual es el TextBox que debe controlar
   objTextBox.TextBox = frm.ActiveControl

   'Creo los items del menu Edit y activo/desactivo los items segun corresponda
   With frm.mnuContextItem
   
      If .UBound > 0 Then
         Load frm.mnuContextItem(.UBound + 1)
         frm.mnuContextItem(.UBound).Enabled = True
         frm.mnuContextItem(.UBound).Caption = "-"
         frm.mnuContextItem(.UBound).Tag = NullString
         Load frm.mnuContextItem(.UBound + 1)
      End If
      
      frm.mnuContextItem(.UBound).Caption = "&Deshacer"
      frm.mnuContextItem(.UBound).Tag = "DESHACER"
      If TypeName(frm.ActiveControl) = "ComboBox" Then
         frm.mnuContextItem(.UBound).Enabled = False
      Else
         frm.mnuContextItem(.UBound).Enabled = objTextBox.CanUndo
      End If
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Enabled = True
      frm.mnuContextItem(.UBound).Caption = "-"
      frm.mnuContextItem(.UBound).Tag = NullString
      
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Enabled = True
      frm.mnuContextItem(.UBound).Caption = "C&ortar"
      frm.mnuContextItem(.UBound).Tag = "CORTAR"
      If TypeName(frm.ActiveControl) = "ComboBox" Then
         frm.mnuContextItem(.UBound).Enabled = False
      Else
         frm.mnuContextItem(.UBound).Enabled = objTextBox.CanCut
      End If
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Enabled = True
      frm.mnuContextItem(.UBound).Caption = "&Copiar"
      frm.mnuContextItem(.UBound).Tag = "COPIAR"
      If TypeName(frm.ActiveControl) = "ComboBox" Then
         frm.mnuContextItem(.UBound).Enabled = True
      Else
         frm.mnuContextItem(.UBound).Enabled = objTextBox.CanCopy
      End If
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Enabled = True
      frm.mnuContextItem(.UBound).Caption = "&Pegar"
      frm.mnuContextItem(.UBound).Tag = "PEGAR"
      If TypeName(frm.ActiveControl) = "ComboBox" Then
         frm.mnuContextItem(.UBound).Enabled = False
      Else
         frm.mnuContextItem(.UBound).Enabled = objTextBox.CanPaste
      End If
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Enabled = True
      frm.mnuContextItem(.UBound).Caption = "Eli&minar"
      frm.mnuContextItem(.UBound).Tag = "ELIMINAR"
      If TypeName(frm.ActiveControl) = "ComboBox" Then
         frm.mnuContextItem(.UBound).Enabled = False
      Else
         frm.mnuContextItem(.UBound).Enabled = objTextBox.CanDelete
      End If
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Enabled = True
      frm.mnuContextItem(.UBound).Caption = "-"
      frm.mnuContextItem(.UBound).Tag = NullString
      
      Load frm.mnuContextItem(.UBound + 1)
      frm.mnuContextItem(.UBound).Enabled = True
      frm.mnuContextItem(.UBound).Caption = "&Seleccionar Todo"
      frm.mnuContextItem(.UBound).Tag = "SELTODO"
      frm.mnuContextItem(.UBound).Enabled = objTextBox.CanSelectAll
      
   End With
   
End Sub

Private Sub tmr1_Timer()
Dim frm As Form
Dim Index As Integer
   
   tmr1.StopTimer
   
   Select Case ContextTag

      Case "DESHACER"
         objTextBox.Undo
      Case "CORTAR"
         objTextBox.Cut
      Case "COPIAR"
         objTextBox.Copy
      Case "PEGAR"
         objTextBox.Paste
      Case "ELIMINAR"
         objTextBox.Delete
      Case "SELTODO"
         objTextBox.SelectAll
      Case "CUSTOMIZE"

         Set frm = Forms.Add("frmCustomizeControl")
         DoEvents

         '/ asigno al frmCustomizeControl la propiedad ControlInfo
         Set frm.ControlInfo = col.Item(CStr(mvarForm.ActiveControl.hWnd))
         '/ asigno el form
         Set frm.ParentForm = mvarForm

         frm.Show vbModal

         Unload frm
         Set frm = Nothing

      Case "EDIT_HELP"
         
         CallHelp False

      Case "FINDQUERYDB"
         FindQueryDB
      Case "ADMINISTRAR"
         CallFormAdmin
      Case "ACTUALIZAR"

         On Error Resume Next

         Index = mvarForm.ActiveControl.Index
         If Err.Number <> 0 Then
            Err.Clear
            mvarForm.SetComboList UCase(mvarForm.ActiveControl.Name)
         Else
            Err.Clear
            mvarForm.SetComboList UCase(mvarForm.ActiveControl.Name), mvarForm.ActiveControl.Index
         End If

      Case Else
         RaiseEvent Messages(CTL_CONTEXT_MENU_CLICK, ContextTag)
   End Select

End Sub

Public Sub SetControlProperties(aTableProperties As Variant, ByVal oControl As Object, ByVal strFieldName As String)
Dim ix As Integer

   ix = InStr(strFieldName, ".")
   If ix > 0 Then strFieldName = Trim(Mid(strFieldName, ix + 1))

   If TypeName(oControl) = "TextBox" Then
      oControl.DataField = strFieldName
      If FieldProperty(aTableProperties, strFieldName, dsTipoDato) = adNumeric Then
         oControl.Alignment = vbRightJustify
      Else
         oControl.Alignment = vbLeftJustify
      End If
      oControl.MaxLength = FieldProperty(aTableProperties, strFieldName, dsMaxLength)
   End If
 
End Sub

Public Property Let ControlData(ByVal vData As Variant)
    mvarControlData = vData
End Property

Public Property Get ControlData() As Variant
    ControlData = mvarControlData
End Property

'---------------------------------------------------------------------------------------
' Procedure : FullControlOnTextBox
' DateTime  : 17/11/2003 09:50
' Author    : tony
' Purpose   : la gestion de LostFocus y Validate de los TextBox y PowerMask
'             no se delega a MValidatePointer (si = True)
'---------------------------------------------------------------------------------------
'
Public Property Let FullControlOnTextBox(ByVal vData As Boolean)
    mvarFullControlOnTextBox = vData
End Property

Public Property Get FullControlOnTextBox() As Boolean
    FullControlOnTextBox = mvarFullControlOnTextBox
End Property

Private Function RemoveChars(ByVal strData As String, ByVal CharList As String) As String
Dim ix As Integer

   'remueve c/uno de los caracteres presentes en CharList
   For ix = 1 To Len(CharList)
      strData = Replace(strData, Mid(CharList, ix, 1), "")
   Next ix
   
   RemoveChars = strData
   
End Function

Private Function FormatData(ByVal strData As String, ByVal strMascara As String) As String
Dim ix As Integer
'Dim iCount As Integer
Dim s As String
Dim s1 As String
Dim iPos As Integer
Dim MascaraSinFormato As String
Dim DatoSinFormato As String
'Dim aMascara() As String
'Dim aCodigo() As String
Dim EditChar As String
Dim c As String
Dim Align As String
   
   '/
   '  Reglas de formateo:
   '  9 : Marcador de posición de dígito (implicito alineamiento a la izquierda).
   '      Si la expresión tiene un dígito en la posición donde aparece el 9 en la cadena
   '      se muestra el dígito; de lo contrario, se muestra un cero en esa posición.
   '      Es incompatible alternar caracteres 9, 0, Z
   '  0 : igual que el 9 (implicito alineamiento a la derecha).
   '  Z : Marcador de posición de dígito (implicito alineamiento a la derecha).
   '      Si la expresión tiene un dígito en la posición donde aparece el Z en la cadena
   '      se muestra el dígito; de lo contrario, se muestra un blanco en esa posición.
   '      Es incompatible alternar caracteres 9, 0, Z
   '
   
   'si en la mascara hay almenos un nueve asumo aliniamiento a la izquierda.Caso contrario a la derecha
   Align = IIf(InStr(strMascara, "9"), "L", "R")
   
   'construyo un string de caracteres de edicion presentes en la mascara
   s = RemoveChars(strMascara, "90Z") 'Ej  si strMascara= "9.99.99#9-9" => "..#-"

   'quito los caracteres que repetidos Ej: ".#-"
   EditChar = s
   For ix = 1 To Len(EditChar)
      s = Replace(s, Mid(EditChar, ix, 1), "") & Mid(EditChar, ix, 1)
   Next ix
   
   EditChar = s

   If Len(strData) = 0 Then Exit Function
   If Len(strMascara) = 0 Then Exit Function
      
   'remuevo los caracteres reservado del dato y de la mascara
   DatoSinFormato = RemoveChars(strData, EditChar)
   MascaraSinFormato = RemoveChars(strMascara, EditChar)
   
   'reemplazo los caracteres de MascaraSinFormato con aquellos de DatosSinFormato
   
   If Align = "L" Then
      'ajusto a la izquierda
      
      'ajusto el dato a la izquierda
      s = Replace(MascaraSinFormato, String(Len(DatoSinFormato), "9"), DatoSinFormato, 1, 1)
      
      If Len(MascaraSinFormato) > Len(DatoSinFormato) Then
         'reemplazo los 9 de la mascara por x
         '(las x representan posiciones de la mascara que no fueron completadas con datos)
         s = Left(s, Len(DatoSinFormato)) & String(Len(MascaraSinFormato) - Len(DatoSinFormato), "x")
      End If
      
      'reemplazo las x por ceros
      s = Replace(s, "x", "0")
      
   Else
      'ajusto a la derecha
      s = IIf(InStr(strMascara, "Z"), PadL(DatoSinFormato, Len(MascaraSinFormato), " "), PadL(DatoSinFormato, Len(MascaraSinFormato), "0"))
   End If
   
   
   'le agrego los caracteres de edicion usados en el formato
   s1 = NullString
   iPos = 1
   For ix = 1 To Len(strMascara)
   
      c = Mid(strMascara, ix, 1)
      
      Select Case c
   
         Case "9", "Z", "0"
            s1 = s1 & Mid(s, iPos, 1)
            iPos = iPos + 1
            
         Case Else
            s1 = s1 & c
            
      End Select
   
   Next ix
   
   FormatData = s1
   
End Function

Private Function PadL(ByVal sTmp As String, lNbr As Long, Optional PadChar As String = " ") As String
   '/***********************************************************
   '  Devuelve un padded string ajustado
   '/***********************************************************
   PadL = Right(String(lNbr, PadChar) & RTrim(sTmp), lNbr)
End Function

Public Sub SetChanged(ByVal hWnd As Long, ByVal bStatus As Boolean)

   Set ControlInfo = col.Item(CStr(hWnd))
   
   ControlInfo.Changed = bStatus
   
'   col.Remove (CStr(hWnd))
'   col.Add ControlInfo, CStr(hWnd)
   
End Sub

Private Function GetFormHwnd(ByVal hWnd As Long) As Long
   'devuelve el Form del Control cuyo hwnd es hWnd
   GetFormHwnd = GetAncestor(hWnd, GA_ROOT)
End Function

Public Sub SetDefaults(Optional ByVal strControlList As String, Optional ByVal strControlListExclude As String)
'Dim bEnabled As Boolean
Dim bVisible      As Boolean
Dim ctrl          As Control
Dim l             As Long
Dim Index         As Integer
Dim MinTabIndex   As Long
Dim strIDControl  As String
Dim objCol        As New Collection
Dim lngWork       As Long
Dim Dummy         As Variant
Dim strMenuKey    As String
   '/
   'Proposito: Setea los valores de default de cada control en caso que los tuviera
   'Nota :
   '  Antes de operar es necesario clonar la coleccion col. No es posible usar esta coleccion
   'para iterar ya que durante el proceso de los distintos mensajes KILLFOCUS dicha colec-
   'cion cambia de contenido
   '/
   
   On Error GoTo GestErr
   
   'vs113sp
   If col Is Nothing Then Exit Sub
   If col.Count = 0 Then Exit Sub
   
   For Each mvarControlInfo In col
      objCol.Add mvarControlInfo
   Next mvarControlInfo
      
   Set ctrl = mvarForm.ActiveControl
   
   strMenuKey = objStdFuntions.EliminaAcentos(mvarForm.MenuKey)
   
   For Each mvarControlInfo In objCol
      
      With mvarControlInfo
      
         'controlo si el control forma parte de una matriz de controles
         On Error Resume Next
         strIDControl = .ctrl.Name & .ctrl.Index
         Index = .ctrl.Index
         If Err Then strIDControl = .ctrl.Name
         Err.Clear
         
'         On Error GoTo GestErr
         
         If Not (Len(strControlListExclude) > 0 And IsInList(strControlListExclude, strIDControl)) Then
         
            If Len(strControlList) = 0 Or _
               (Len(strControlList) > 0 And IsInList(strControlList, strIDControl)) Then
            
            
               If .DefaultValue <> NullString Then
                  Select Case TypeName(.ctrl)
                     Case "TextBox", "PowerMask"
                        
                        If TypeName(.ctrl) = "TextBox" And .FieldType = adNumeric Then
                           Dim c As Double
                           c = CDbl(.DefaultValue)
                           .ctrl.Text = c
                        Else
                           'si es de tipo fecha hora y valida hora, o sea si es hora
                           'Si usa hora del día, porque podria tener una hora cualquiera
                           If .FieldType = adDBTimeStamp And .Validar = vldHora Then
                              If Len(.DefaultValue) > 0 Then
                                 On Error Resume Next
                                 Dummy = NullString
                                 Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\" & CUsuario.Usuario & "\Usa Fecha del Dia")
                                 If Err.Number <> 0 Then
                                    Dummy = objTablasValores("Forms\" & strMenuKey & "\" & strIDControl & "\Usa Fecha del Dia")
                                    Err.Clear
                                 End If
                                 If Dummy = si Then
                                    .DefaultValue = Format(Time, "HH:MM")
                                 End If
                                 On Error GoTo GestErr
                              End If
                           End If
                           
                           .ctrl.Text = .DefaultValue
                           
                        End If
                        
                        SetModify .ctrl.hWnd, True
                        
                        'guardo el estado anterior
                        bVisible = .ctrl.Visible
      
                        lngWork = .ctrl.hWnd
                        'deshabilito el control de la edicion
                        l = SendMessage(.ctrl.hWnd, DISABLE_CONTROL_EDITION, 0&, 0&)
                        
                        'necesito  hacer visible y habilitar el control para forzar un SetFocus
                        .ctrl.Visible = True
                        .ctrl.Enabled = True
   
                        If .ctrl.Visible And .ctrl.Enabled Then
                           .ctrl.SetFocus
                            DoEvents         'desencadena eventos (--> cambiando el mvarControlInfo)
                        End If
   
                        Set mvarControlInfo = col.Item(CStr(lngWork))
                        
                        'restotuyo el estado anterior
                        .ctrl.Visible = bVisible
   
                        'habilito el control de la edicion
                       l = SendMessage(.ctrl.hWnd, ENABLE_CONTROL_EDITION, 0&, 0&)
                        
                        
                     Case "DataCombo"
                        If Len(.DefaultValue) > 0 Then .ctrl.BoundText = .DefaultValue
'                        If .DefaultValue <> 0 Then .Ctrl.BoundText = .DefaultValue
                     Case "ComboBox"
                        If .ctrl.Style = vbComboDropdownList Then
                           .ctrl.ListIndex = ComboSearch(.ctrl, .DefaultValue)
                        Else
                           .ctrl.Text = .DefaultValue
                        End If
                     Case "OptionButton"
                        .ctrl.Value = (.DefaultValue = si)
                     Case "CheckBox"
                        .ctrl.Value = IIf(.DefaultValue = si, vbChecked, vbUnchecked)
                  End Select
                  
               End If
            
            End If
         End If
      End With
            
   Next mvarControlInfo

   'si esta deshabilitada la edición, deshabilito el control
   For Each mvarControlInfo In objCol
      
      With mvarControlInfo
         If .ControlEdicion = False Then .ctrl.Enabled = False
      End With
         
   Next mvarControlInfo
   
   
   On Error Resume Next
   
   If Not ctrl Is Nothing Then
      ' el focus esta en un control
      If ctrl.Visible And ctrl.Enabled Then ctrl.SetFocus
   Else
      ' el focus no esta en ningun control perteneciente al form
      
      ' busco el control con el menor TabIndex existente
      Dim FirstControl As Control
      
      MinTabIndex = 32700
      For Each ctrl In mvarForm.Controls
         If ctrl.TabIndex < MinTabIndex Then
            If InStr("TextBox PowerMask DataCombo ComboBox CheckBox OptionButton", TypeName(ctrl)) > 0 Then
               MinTabIndex = ctrl.TabIndex
               Set FirstControl = ctrl
            End If
         End If
      Next ctrl
      
      FirstControl.SetFocus
   
   End If

'   mvarUserControlsChanged = False
   'Intenta arreglar pregunta "Desea guardar los cambios?" 03/05/07 Marcelo
   UserControlsChanged = False
   
   Set objCol = Nothing

   Exit Sub

GestErr:
   SetError mError, CLASSNAME, "SetDefaults"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Sub

Public Sub ResetPowerMask(ByVal hWnd As Long, ByVal PreviousHwndSectionsList As String)
Dim Section As PowerMaskSection
Dim aSplit() As String
Dim ix As Integer

   '/ Cuando cambia la propiedad Mask del PowerMask es necesario llamar este metodo
   '  para cambiar el viejo subclassing de las secciones por el subclassing de las nuevas.
   '  hWnd: es el Hwnd del PowerMask
   '  PreviousHwndSectionsList : es la list de hwnd de las secciones anteriores
   '/
   
   On Error GoTo GestErr
   
   Set ControlInfo = col.Item(CStr(hWnd))
   
   aSplit = Split(PreviousHwndSectionsList, ";")
   
   For ix = LBound(aSplit) To UBound(aSplit)
   
      'remuevo los mensajes
      If strUsaSubClassing2009 = si Then
         Hooking.RemoveSubClass CLng(aSplit(ix))
      Else
         DetachAllMessages CLng(aSplit(ix))
      End If
      'remuevo la seccion de la colecciòn
      col.Remove aSplit(ix)
   Next ix
   
   For Each Section In ControlInfo.ctrl.Sections
   
      'informo los mensajes que serán subclasados
      'AttachAllMessages Section.Control.hWnd
      If strUsaSubClassing2009 = si Then
         Hooking.AddSubClassing Section.Control.hWnd
      Else
         AttachPmkSectionMessages Section.Control.hWnd
      End If
      'Las secciones son colocadas en la colección simplemente para tener accesso
      'al al informacion del control "PowerMask" que las contiene
      
      Dim SectionInfo As New clsControlInfo
      
      With SectionInfo
         Set .ctrl = Section.Control
         .TypeOfControl = "PowerMaskSection"
         .hWnd = Section.Control.hWnd
         .hWndParent = ControlInfo.ctrl.hWnd
         .FindIsEnabled = ControlInfo.FindIsEnabled
         .AdminIsEnabled = ControlInfo.AdminIsEnabled
         
         .BackColor = ControlInfo.BackColor
         .ControlEdicion = ControlInfo.ControlEdicion
         
         col.Add SectionInfo, CStr(.hWnd)
      End With
      
   Next Section
   
   Exit Sub
   
GestErr:
   SetError mError, CLASSNAME, "ResetPowerMask"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Sub

Public Property Let CallerObjName(ByVal vData As String)
    mvarCallerObjName = vData
End Property
Public Property Get CallerObjName() As String
    CallerObjName = mvarCallerObjName
End Property

Private Sub AttachAllMessages(ByVal hWnd As Long)

   On Error GoTo GestErr
   
   AttachMessage Me, hWnd, WM_ENABLE
   AttachMessage Me, hWnd, WM_SETTEXT
   AttachMessage Me, hWnd, WM_KEYDOWN
   AttachMessage Me, hWnd, WM_CHAR
   AttachMessage Me, hWnd, WM_RBUTTONDOWN
   AttachMessage Me, hWnd, DISABLE_CONTROL_EDITION
   AttachMessage Me, hWnd, ENABLE_CONTROL_EDITION

   Exit Sub
   
GestErr:
   SetError mError, CLASSNAME, "AttachAllMessages"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Sub

Private Sub DetachAllMessages(ByVal hWnd As Long)

   On Error GoTo GestErr

   DetachMessage Me, hWnd, WM_ENABLE
   DetachMessage Me, hWnd, WM_SETTEXT
   DetachMessage Me, hWnd, WM_KEYDOWN
   DetachMessage Me, hWnd, WM_CHAR
   DetachMessage Me, hWnd, WM_RBUTTONDOWN
   DetachMessage Me, hWnd, DISABLE_CONTROL_EDITION
   DetachMessage Me, hWnd, ENABLE_CONTROL_EDITION

   Exit Sub
   
GestErr:
   SetError mError, CLASSNAME, "DetachAllMessages"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Sub

Private Sub AttachPmkSectionMessages(ByVal hWnd As Long)

   On Error GoTo GestErr

   AttachMessage Me, hWnd, WM_ENABLE
   AttachMessage Me, hWnd, WM_KEYDOWN
   AttachMessage Me, hWnd, WM_CHAR
   AttachMessage Me, hWnd, WM_RBUTTONDOWN
   AttachMessage Me, hWnd, DISABLE_CONTROL_EDITION
   AttachMessage Me, hWnd, ENABLE_CONTROL_EDITION

   Exit Sub
   
GestErr:
   SetError mError, CLASSNAME, "AttachPmkSectionMessages"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Sub

Private Sub DetachPmkSectionMessages(ByVal hWnd As Long)

   On Error GoTo GestErr

   DetachMessage Me, hWnd, WM_ENABLE
   DetachMessage Me, hWnd, WM_KEYDOWN
   DetachMessage Me, hWnd, WM_CHAR
   DetachMessage Me, hWnd, WM_RBUTTONDOWN
   DetachMessage Me, hWnd, DISABLE_CONTROL_EDITION
   DetachMessage Me, hWnd, ENABLE_CONTROL_EDITION

   Exit Sub
   
GestErr:
   SetError mError, CLASSNAME, "DetachPmkSectionMessages"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse

   Select Case CurrentMessage
      Case WM_RBUTTONDOWN, WM_CHAR ', WM_SETTEXT
         ISubclass_MsgResponse = emrConsume
      Case Else
         ISubclass_MsgResponse = emrPreprocess
   End Select

End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim ControlInfo As clsControlInfo
Static m As Boolean

   On Error Resume Next

   If col.Count = 0 Then Exit Function
   
   Set ControlInfo = col.Item(CStr(hWnd))

   Select Case iMsg

      Case WM_ENABLE
      
         If Not ControlInfo.ControlEdicion Then
            If IsNull(ControlInfo.Tag) Then
               'el control no admite edicion
               ControlInfo.ctrl.BackColor = DisabledColor
               ControlInfo.ctrl.Enabled = False
               Exit Function
            End If
         End If

         ControlInfo.ctrl.BackColor = IIf(Hex(wParam) = "FFFF" Or Hex(wParam) = "1" Or Hex(wParam) = "FFFFFFFF", ControlInfo.BackColor, DisabledColor)

      Case DISABLE_CONTROL_EDITION
      
         ControlInfo.Tag = "ControlEditionDisabled"
         Exit Function

      Case ENABLE_CONTROL_EDITION
      
         ControlInfo.Tag = Null
         Exit Function

      Case WM_KEYDOWN

         Select Case CInt(wParam)
            Case vbKeyF1
               RaiseEvent Messages(CTL_BEFORE_F1, ControlInfo)
               '
               '  llamo a la Ayuda
               '
               CallHelp True
               
               RaiseEvent Messages(CTL_AFTER_F1, ControlInfo)
            Case vbKeyF2
               RaiseEvent Messages(CTL_BEFORE_F2, ControlInfo)
               If ControlInfo.AdminIsEnabled Then
                  If CanAdmin Then CallFormAdmin
               End If
               RaiseEvent Messages(CTL_AFTER_F2, ControlInfo)
            Case vbKeyF3
               RaiseEvent Messages(CTL_BEFORE_F3, ControlInfo)
               If ControlInfo.FindIsEnabled Then FindQueryDB
               RaiseEvent Messages(CTL_AFTER_F3, ControlInfo)
            Case vbKeyF4
               RaiseEvent Messages(CTL_BEFORE_F4, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F4, ControlInfo)
            Case vbKeyF5
               RaiseEvent Messages(CTL_BEFORE_F5, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F5, ControlInfo)
            Case vbKeyF6
               RaiseEvent Messages(CTL_BEFORE_F6, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F6, ControlInfo)
            Case vbKeyF7
               RaiseEvent Messages(CTL_BEFORE_F7, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F7, ControlInfo)
            Case vbKeyF8
               RaiseEvent Messages(CTL_BEFORE_F8, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F8, ControlInfo)
            Case vbKeyF9
               RaiseEvent Messages(CTL_BEFORE_F9, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F9, ControlInfo)
            Case vbKeyF10
               RaiseEvent Messages(CTL_BEFORE_F10, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F10, ControlInfo)
            Case vbKeyF11
               RaiseEvent Messages(CTL_BEFORE_F11, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F11, ControlInfo)
            Case vbKeyF12
               RaiseEvent Messages(CTL_BEFORE_F12, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F12, ControlInfo)
         End Select

         CheckKey hWnd, CInt(wParam)
         
      Case WM_RBUTTONDOWN

         ' recibio el mensaje WM_CONTEXTMENU
   
         Set ControlInfo = col.Item(CStr(hWnd))
   
         SetFocus hWnd
   
         If ControlInfo.TypeOfControl = "PowerMaskSection" Then
            'si el control es una Section, entonces busco su PowerMask
            Set ControlInfo = col.Item(CStr(ControlInfo.hWndParent))
         End If
         
         If ControlInfo.ctrl.Enabled = False Then Exit Function
   
         '-- obtengo el usuario para que este pueda ser utilizado internamente en esta clase
         If mvarUsuario Is Nothing Then
            '-- si la propiedad Usuario no fue definida, la busco en el proyecto
            RaiseEvent Messages(CTL_QUERY_USER, mvarUsuario)
         End If
   
         SetContextMenu ControlInfo.ContextMenu
   
         RaiseEvent Messages(CTL_CONTEXT_MENU_AFTER_SET, ControlInfo)
   
         If ControlInfo.ContextMenu = "GRILLA" Then
            If Len(ControlInfo.FindDataLink) = 0 Then
               mvarForm.mnuContextItem(0).Enabled = False
            End If
         End If
   
         mvarForm.PopupMenu mvarForm.mnuContextMenu, vbPopupMenuRightButton
   
         UnloadContextMenu mvarForm
   
         SetFocus hWnd
         
         Exit Function

      Case WM_CHAR

         Set ControlInfo = col.Item(CStr(hWnd))
   
         'el campo es numèrico y posee cifras decimales
         Select Case True
         
            Case ControlInfo.FieldDecimales > 0
               wParam = CheckDecimal(CInt(wParam))
               ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
         
            Case ControlInfo.FieldType = adNumeric
               'para campos numéricos solo se permiten números
               If ControlInfo.FieldType = adNumeric Then
                  If wParam = vbKeyBack Or wParam = Asc("+") Or wParam = Asc("-") Or wParam = Asc(DecimalChar) Or ControlInfo.TypeOfControl = "DataCombo" Then 'Bug 5301
                     ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                  Else
                     If Not IsNumeric(Chr(wParam)) Then
                        wParam = 0
                        ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                     Else
                        ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                     End If
                  End If
               End If
   
            Case ControlInfo.MayuscMinusc = "U"
               'transformo en mayuscula
               wParam = Asc(UCase(Chr(wParam)))
               ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
         
            Case ControlInfo.MayuscMinusc = "L"
               'transformo en minuscula segun corresponde
               wParam = Asc(LCase(Chr(wParam)))
               ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            Case Else
               ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
               
         End Select
   
         If ControlInfo.TypeOfControl = "PowerMaskSection" Then
            Set ControlInfo = col.Item(CStr(ControlInfo.hWndParent))
            ' lo pongo modificado
            ControlInfo.ctrl.SetModify True
         End If
         If ControlInfo.TypeOfControl = "TextBox" Then
            ' lo pongo modificado
            SetModify hWnd, True
         End If

      Case WM_SETTEXT
         
         Dim s As String
        ' Debug.Assert ControlInfo.ctrl.Name <> "txtFechaDesdeLineaCredito"
         If m Then Exit Function
         
         ' llamo el procedimiento que formatea (si es requerido)
         
         m = True
         AplicarFormato ControlInfo
         m = False
         
   End Select

End Function

Public Sub AplicarFormato(ByRef ControlInfo As clsControlInfo)

   'formateo si es requerido
   Select Case ControlInfo.Formatear
      Case FormatSi
      
         ControlInfo.ctrl = FormatControl(ControlInfo.ctrl, ControlInfo.ctrl, True, ControlInfo.MascaraFormato)
         
      Case FormatSoloNumericos

         If IsNumeric(ControlInfo.ctrl) Then
            ControlInfo.ctrl = FormatControl(ControlInfo.ctrl, ControlInfo.ctrl, True, ControlInfo.MascaraFormato)
         End If

   End Select

End Sub

Private Sub m_TextControls_ItemGotFocus(ByVal nIndex As Long, ByVal sKey As String)
'
End Sub

'---------------------------------------------------------------------------------------
' Procedure : m_TextControls_ItemSetText
' DateTime  : 28/11/2002 12:58
' Author    : Algoritmo
' Purpose   : el evento ItemSetText se genera solo para los PowerMask cuando se asigna
'             un valor a la propiedad Text del PowerMask
'---------------------------------------------------------------------------------------
'
Private Sub m_TextControls_ItemSetText(ByVal nIndex As Long, ByVal sKey As String)

   ' ubico el ControlInfo
   If col Is Nothing Then Exit Sub   'vs113sp
   
   Set ControlInfo = col.Item(CStr(m_TextControls.Item(nIndex).hWnd))
   
End Sub

Private Sub m_TextControls_ItemLostFocus(ByVal nIndex As Long, ByVal sKey As String)

   ' ubico el ControlInfo
   If col Is Nothing Then Exit Sub   'vs113sp
   
   Set ControlInfo = col.Item(CStr(m_TextControls.Item(nIndex).hWnd))
   
   If AlgStdFunc.IsChanged(m_TextControls.Item(nIndex).hWnd) Then
       '
       ' Cambio --> llamo el procedimiento que formatea (si es requerido)
       '
       AplicarFormato ControlInfo
       
       SetModify m_TextControls.Item(nIndex).hWnd, False
   End If
      
End Sub

Private Sub m_TextControls_ItemValidate(ByVal nIndex As Long, ByVal sKey As String)
Dim Cancel As Boolean

   If col Is Nothing Then Exit Sub   'vs113sp
   
   Set ControlInfo = col.Item(CStr(m_TextControls.Item(nIndex).hWnd))

   If ControlInfo.Validar <> 0 Then
   
      'efectuo la validación
      Cancel = Not IsValid(ControlInfo.Validar)
      
      ' informo a la coleccion si el evento Validate debe ser Cancelado
      m_TextControls.Item(nIndex).Cancel = Cancel

   End If

End Sub

''Get ANSI string from pointer.
''lpzA - pointer to ANSI string
''nSize - size of string. If ommitted, string will be truncated up to first null char (chr(0))
'
'Private Function StrFromPtrA(ByVal lpszA As Long, Optional nSize As Long = 0) As String
'  Dim s As String, bTrim As Boolean
'  If nSize = 0 Then
'     nSize = lstrlenA(lpszA)
'     bTrim = True
'  End If
'  s = String(nSize, Chr$(0))
'  CopyStringA s, ByVal lpszA
'  If bTrim Then s = TrimNULL(s)
'  StrFromPtrA = s
'End Function
'
'Private Function TrimNULL(ByVal str As String) As String
'   If InStr(str, Chr$(0)) > 0& Then
'       TrimNULL = Left$(str, InStr(str, Chr$(0)) - 1&)
'   Else
'       TrimNULL = str
'   End If
'End Function
'
'---------------------------------------------------------------------------------------
' Procedure : UserControls
' DateTime  : 20/05/2003 16:56
' Author    : tony
' Purpose   : Si bReset = True copia el contenido de c/control en la propiedad OriginalValue
'             del objeto ControlInfo.
'             Si bReset = Falso devuelve un valor que indica si almenos uno de los controles
'             ha cambiado su valor original
'---------------------------------------------------------------------------------------
'
Private Function UserControls(Optional ByVal bReset As Boolean = False) As Boolean
Dim ctrl As Control
Dim vValue As Variant
Dim UserControlsChanged As Boolean

   On Error Resume Next
   '
   ' controlo si el Form es un ABM y en tal caso si el boton para salvar esta habilitado
   '
   If mvarIsABM = False Then Exit Function
   
'   If Not mvarForm.tbr1.Buttons(BUTTON_SAVE).Enabled Then Exit Function
'  esto ya no es valido pues con la vs113 se desabilita el boton save inmediatamente despues de pulsarlo
   
   If Not bReset And mvarUserControlsChanged Then UserControls = True: Exit Function
   
   For Each ControlInfo In col
      '
      ' creo referencia al control
      '
      Set ctrl = ControlInfo.ctrl
      '
      ' obtengo el valor original del control
      '
      vValue = ControlInfo.OriginalValue
      
      If ctrl.Visible = True Then
         
         If bReset Then
      
            If TypeOf ctrl Is OptionButton Then
                vValue = ctrl.Value
                
            ElseIf TypeOf ctrl Is CheckBox Then
                vValue = ctrl.Value
                
            ElseIf TypeOf ctrl Is TextBox Then
                vValue = ctrl.Text
                
            ElseIf TypeOf ctrl Is PowerMask Then
                vValue = ctrl.Text
                
            ElseIf TypeOf ctrl Is ComboBox Then
                vValue = ctrl.Text
                
            ElseIf TypeOf ctrl Is DataCombo Then
                vValue = ctrl.BoundText
                
            ElseIf TypeOf ctrl Is RichTextBox Then
                vValue = ctrl.Text
                
            End If
            '
            ' inicializo el valor original
            '
            ControlInfo.OriginalValue = vValue
            
         Else
         
            If TypeOf ctrl Is OptionButton Then
                UserControlsChanged = (vValue <> ctrl.Value)
            ElseIf TypeOf ctrl Is CheckBox Then
                UserControlsChanged = (vValue <> ctrl.Value)
            ElseIf TypeOf ctrl Is TextBox Then
                UserControlsChanged = (vValue <> ctrl.Text)
            ElseIf TypeOf ctrl Is PowerMask Then
                UserControlsChanged = (vValue <> ctrl.Text)
            ElseIf TypeOf ctrl Is ComboBox Then
                UserControlsChanged = (vValue <> ctrl.Text)
            ElseIf TypeOf ctrl Is DataCombo Then
                UserControlsChanged = (vValue <> ctrl.BoundText)
            ElseIf TypeOf ctrl Is RichTextBox Then
                UserControlsChanged = (vValue <> ctrl.Text)
            End If
            
            UserControls = UserControlsChanged
            If UserControlsChanged Then Exit For
            
         End If
         
      End If
      
   Next
   
   Set ctrl = Nothing
   
   Err.Clear
      
End Function

'---------------------------------------------------------------------------------------
' Procedure : UserControlsChanged
' DateTime  : 20/05/2003 17:07
' Author    : tony
' Purpose   : Devuelve True si almenos un control del la interface grafica cambio de valor
'---------------------------------------------------------------------------------------
'
Public Property Let UserControlsChanged(ByVal vData As Boolean)

   mvarUserControlsChanged = vData
   If vData = False Then UserControls bReset:=True
   
End Property
Public Property Get UserControlsChanged() As Boolean

   mvarUserControlsChanged = UserControls()
   UserControlsChanged = mvarUserControlsChanged
   
End Property

'---------------------------------------------------------------------------------------
' Procedure : IsABM
' DateTime  : 20/05/2003 18:01
' Author    : tony
' Purpose   : determina si el Form es un ABM
'---------------------------------------------------------------------------------------
'
Private Function IsABM() As Boolean
Dim objABM As Object

   On Error Resume Next

   Set objABM = mvarForm.objABM
   IsABM = (Err.Number = 0)
   
   Err.Clear
   
End Function

Public Property Get ControlsInfo() As clsControlInfos
    Set ControlsInfo = col
End Property
Public Function GetControlInfo(ByVal hWnd As Long) As clsControlInfo

    ' ubico el ControlInfo
    Set GetControlInfo = col.Item(CStr(hWnd))

End Function
Public Sub SetControlsInfo(ByVal hWnd As Long, ByVal vData As clsControlInfo)

   col.Remove (CStr(hWnd))
   col.Add vData, CStr(hWnd)
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CallHelp
' DateTime  : 19/10/2004 15:16
' Author    : tony
' Purpose   : Llama al Help Contextual
'---------------------------------------------------------------------------------------
'
Private Sub CallHelp(ByVal bLock As Boolean)
Dim Frm1 As Form
Dim Index As Integer
Dim strMenuKey As String

   Set Frm1 = Forms.Add("frmAyudaContextual")
   
   strMenuKey = objStdFuntions.EliminaAcentos(mvarForm.MenuKey)
   
   DoEvents

   On Error Resume Next
   Index = mvarForm.ActiveControl.Index
      
   Set ControlInfo = col.Item(CStr(mvarForm.ActiveControl.hWnd))

   If Err.Number <> 0 Then
      Err.Clear
      Frm1.InitAOL mvarForm.ControlData, strMenuKey & ";" & UCase(mvarForm.ActiveControl.Name), bLock, ControlInfo.BoundField
   Else
      Err.Clear
      Frm1.InitAOL mvarForm.ControlData, strMenuKey & ";" & strMenuKey & ";" & UCase(mvarForm.ActiveControl.Name) & ";" & mvarForm.ActiveControl.Index, bLock, ControlInfo.BoundField
   End If

   Frm1.Show vbModal

   Unload Frm1
   Set Frm1 = Nothing

End Sub

Private Sub Hooking_BeforeMessage(ByVal hWnd As Long, ByVal iMsg As Long, _
                                  wParam As Long, ByVal lParam As Long, _
                                  ByRef bCallWinProc As Boolean)
Dim ControlInfo As clsControlInfo
Static m As Boolean

   On Error Resume Next

   If col.Count = 0 Then Exit Sub
   
   Set ControlInfo = col.Item(CStr(hWnd))

   Select Case iMsg

      Case WM_ENABLE
      
         If Not ControlInfo.ControlEdicion Then
            If IsNull(ControlInfo.Tag) Then
               'el control no admite edicion
               ControlInfo.ctrl.BackColor = DisabledColor
               ControlInfo.ctrl.Enabled = False
               Exit Sub
            End If
         End If

         ControlInfo.ctrl.BackColor = IIf(Hex(wParam) = "FFFF" Or Hex(wParam) = "1" Or Hex(wParam) = "FFFFFFFF", ControlInfo.BackColor, DisabledColor)

      Case DISABLE_CONTROL_EDITION
      
         ControlInfo.Tag = "ControlEditionDisabled"
         Exit Sub

      Case ENABLE_CONTROL_EDITION
      
         ControlInfo.Tag = Null
         Exit Sub

      Case WM_KEYDOWN

         Select Case CInt(wParam)
            Case vbKeyF1
               RaiseEvent Messages(CTL_BEFORE_F1, ControlInfo)
               '
               '  llamo a la Ayuda
               '
               CallHelp True
               
               RaiseEvent Messages(CTL_AFTER_F1, ControlInfo)
            Case vbKeyF2
               RaiseEvent Messages(CTL_BEFORE_F2, ControlInfo)
               If ControlInfo.AdminIsEnabled Then
                  If CanAdmin Then CallFormAdmin
               End If
               RaiseEvent Messages(CTL_AFTER_F2, ControlInfo)
            Case vbKeyF3
               RaiseEvent Messages(CTL_BEFORE_F3, ControlInfo)
               If ControlInfo.FindIsEnabled Then FindQueryDB
               RaiseEvent Messages(CTL_AFTER_F3, ControlInfo)
            Case vbKeyF4
               RaiseEvent Messages(CTL_BEFORE_F4, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F4, ControlInfo)
            Case vbKeyF5
               RaiseEvent Messages(CTL_BEFORE_F5, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F5, ControlInfo)
            Case vbKeyF6
               RaiseEvent Messages(CTL_BEFORE_F6, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F6, ControlInfo)
            Case vbKeyF7
               RaiseEvent Messages(CTL_BEFORE_F7, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F7, ControlInfo)
            Case vbKeyF8
               RaiseEvent Messages(CTL_BEFORE_F8, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F8, ControlInfo)
            Case vbKeyF9
               RaiseEvent Messages(CTL_BEFORE_F9, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F9, ControlInfo)
            Case vbKeyF10
               RaiseEvent Messages(CTL_BEFORE_F10, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F10, ControlInfo)
            Case vbKeyF11
               RaiseEvent Messages(CTL_BEFORE_F11, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F11, ControlInfo)
            Case vbKeyF12
               RaiseEvent Messages(CTL_BEFORE_F12, ControlInfo)
               RaiseEvent Messages(CTL_AFTER_F12, ControlInfo)
         End Select

         CheckKey hWnd, CInt(wParam)
         
      Case WM_RBUTTONDOWN

         ' recibio el mensaje WM_CONTEXTMENU
   
         Set ControlInfo = col.Item(CStr(hWnd))
   
         SetFocus hWnd
   
         If ControlInfo.TypeOfControl = "PowerMaskSection" Then
            'si el control es una Section, entonces busco su PowerMask
            Set ControlInfo = col.Item(CStr(ControlInfo.hWndParent))
         End If
         
         If ControlInfo.ctrl.Enabled = False Then Exit Sub
   
         '-- obtengo el usuario para que este pueda ser utilizado internamente en esta clase
         If mvarUsuario Is Nothing Then
            '-- si la propiedad Usuario no fue definida, la busco en el proyecto
            RaiseEvent Messages(CTL_QUERY_USER, mvarUsuario)
         End If
   
         SetContextMenu ControlInfo.ContextMenu
   
         RaiseEvent Messages(CTL_CONTEXT_MENU_AFTER_SET, ControlInfo)
   
         If ControlInfo.ContextMenu = "GRILLA" Then
            If Len(ControlInfo.FindDataLink) = 0 Then
               mvarForm.mnuContextItem(0).Enabled = False
            End If
         End If
   
         mvarForm.PopupMenu mvarForm.mnuContextMenu, vbPopupMenuRightButton
   
         UnloadContextMenu mvarForm
   
         SetFocus hWnd
         
         Exit Sub

      Case WM_CHAR

         Set ControlInfo = col.Item(CStr(hWnd))
         
         'el campo es numèrico y posee cifras decimales
         Select Case True
         
            Case ControlInfo.FieldDecimales > 0
               wParam = CheckDecimal(CInt(wParam))
               'ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
         
            Case ControlInfo.FieldType = adNumeric
               'para campos numéricos solo se permiten números
               If ControlInfo.FieldType = adNumeric Then
                  If wParam = vbKeyBack Or wParam = Asc("+") Or wParam = Asc("-") Or wParam = Asc(DecimalChar) Or ControlInfo.TypeOfControl = "DataCombo" Then
                     'ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                  Else
                     If Not IsNumeric(Chr(wParam)) Then
                        wParam = 0
'                        ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                     Else
'                        ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                     End If
                  End If
               End If
   
            Case ControlInfo.MayuscMinusc = "U"
               'transformo en mayuscula
               wParam = Asc(UCase(Chr(wParam)))
'               ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
         
            Case ControlInfo.MayuscMinusc = "L"
               'transformo en minuscula segun corresponde
               wParam = Asc(LCase(Chr(wParam)))
'               ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            Case Else
'               ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
               
         End Select
   
         If ControlInfo.TypeOfControl = "PowerMaskSection" Then
            Set ControlInfo = col.Item(CStr(ControlInfo.hWndParent))
            ' lo pongo modificado
            ControlInfo.ctrl.SetModify True
         End If
         If ControlInfo.TypeOfControl = "TextBox" Then
            ' lo pongo modificado
            SetModify hWnd, True
         End If

      Case WM_SETTEXT
         
         Dim s As String
         
         If m Then Exit Sub
         
         ' llamo el procedimiento que formatea (si es requerido)
         
         m = True
         AplicarFormato ControlInfo
         m = False
         
   End Select

End Sub
'retorna el key de un combo box

'---------------------------------------------------------------------------------------
' Procedure : GetKey
' DateTime  : 03/01/2014 11:40
' Author    : pablo.pellegrini
' Purpose   : busca el texto del combo en el array del diccionario, si lo encuentra retorna la posición equivalente del key.
' TP        :
'---------------------------------------------------------------------------------------
'
Public Function GetKey(cmbControl As Object, Optional vDefaultValue As Variant) As Variant
   Dim CtrlInfo            As clsControlInfo
   Dim cmbCtrl             As ComboBox
   Dim arrDescripcoines()  As String
   Dim arrKeys()           As String
   Dim ix As Long
   
   On Error GoTo GestErr
   
   If TypeName(cmbControl) <> "ComboBox" Then
      Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, _
         "El control " & cmbControl.Name & " no es del tipo 'ComboBox'."
   End If
   
   On Error Resume Next
   Set CtrlInfo = col.Item(CStr(cmbControl.hWnd))
   
   If Err.Number <> 0 Then
      On Error GoTo GestErr
      Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, _
      cmbControl.Name & " no es un elemento de la clase clsControls"
   End If
   On Error GoTo GestErr
   
   Set cmbCtrl = cmbControl
   
   '-------------------------------------------------------------------------------------
   'Inicio de la logica:
   If Not IsEmpty(vDefaultValue) Then
      GetKey = ControlValue(cmbCtrl, vDefaultValue)
   Else
      GetKey = ControlValue(cmbCtrl, NullString)
   End If
   
   If cmbCtrl.ListCount = 0 Or cmbCtrl.ListIndex = -1 Then Exit Function
   
   If Len(CtrlInfo.FieldValoresPosibles) > 0 And Len(CtrlInfo.Comentarios) > 0 Then
      
      arrDescripcoines = Split(CtrlInfo.FieldValoresPosibles, ";")
      arrKeys = Split(CtrlInfo.Comentarios, ";")
      If UBound(arrDescripcoines) = UBound(arrKeys) Then
         For ix = LBound(arrDescripcoines) To UBound(arrDescripcoines)
            If cmbCtrl.Text = arrDescripcoines(ix) Then
               
               GetKey = arrKeys(ix)
               'GetKey = ControlValue(cmbCtrl, arrKeys(ix))
               Exit For
            End If
         Next
      Else
         Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, _
         "Cantidad de valores posibles incorrectos en el diccionario. Campo: " & CtrlInfo.BoundField
      End If
   Else
      Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, _
      "Falta información en diccionario. Campo: " & CtrlInfo.BoundField
   End If
   
   
   Exit Function

GestErr:
   SetError mError, CLASSNAME, "GetKey"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Function
'setea el combo box en función del key
Public Function SetKey(cmbControl As Object, vKey As Variant) As Variant
   Dim CtrlInfo            As clsControlInfo
   Dim cmbCtrl             As ComboBox
   Dim arrDescripcoines()  As String
   Dim arrKeys()           As String
   Dim bEncontroKey        As Boolean
   Dim ix As Long
   
   On Error GoTo GestErr
   
   SetKey = Empty
   
   If TypeName(cmbControl) <> "ComboBox" Then
      Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, _
         "El control " & cmbControl.Name & " no es del tipo 'ComboBox'."
   End If
   
   On Error Resume Next
   Set CtrlInfo = col.Item(CStr(cmbControl.hWnd))
   
   If Err.Number <> 0 Then
      On Error GoTo GestErr
      Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, _
      cmbControl.Name & " no es un elemento de la clase clsControls"
   End If
   On Error GoTo GestErr
   
   Set cmbCtrl = cmbControl
   
   '-------------------------------------------------------------------------------------
   'Inicio de la logica:
   If IsEmpty(vKey) Then
      cmbCtrl.ListIndex = -1
      Exit Function
   End If
   
   bEncontroKey = False
   
   If Len(CtrlInfo.FieldValoresPosibles) > 0 And Len(CtrlInfo.Comentarios) > 0 Then
      
      arrDescripcoines = Split(CtrlInfo.FieldValoresPosibles, ";")
      arrKeys = Split(CtrlInfo.Comentarios, ";")
      
      If UBound(arrDescripcoines) = UBound(arrKeys) Then
      
         For ix = LBound(arrKeys) To UBound(arrKeys)
            If vKey = arrKeys(ix) Then
               cmbCtrl.ListIndex = ComboSearch(cmbCtrl, arrDescripcoines(ix))
               bEncontroKey = True
               Exit For
            End If
         Next
      Else
         Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, _
         "Cantidad de valores posibles incorrectos en el diccionario. Campo: " & CtrlInfo.BoundField
      End If
   Else
      Err.Raise vbObjectError + 1, KNOWN_ERRORSOURCE & CLASSNAME, _
      "Falta información en diccionario. Campo: " & CtrlInfo.BoundField
   End If
   
   If Not bEncontroKey Then cmbCtrl.ListIndex = -1
   
   Exit Function

GestErr:
   SetError mError, CLASSNAME, "SetKey"
   Err.Raise Number:=mError.Num, source:=mError.source, Description:=mError.Description
End Function
'retorna una colección donde cada item tiene "Key" & ";" & "Descripcion", esto se podria usar para las querys donde se arman decode.
Public Function GetCollectionKeyValues(cmbControl As Object) As Collection
' Set ControlInfo = col.Item(CStr(hWnd))
End Function
