VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlInfoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private m_lhWnd                         As Long           'hwnd del control
Private m_vCtrl                         As Object         'control
Private m_sTypeOfControl                As String         'tipo de control
Private m_eValidar                      As ValidEnum      'validacion requerida para el tipo de dato
Private m_sBoundField                   As String         'Nombre del campo asociado al control
Private m_sObjectProperty               As String         'Nombre de la propiedad del objeto BO asociada al control
Private m_vMayuscMinusc                 As Variant        'valores posibles: "U":Ucase; "L":LCase;""
Private m_iFieldDimension               As Integer        'Dimension del campo definida en el diccionario
Private m_iFieldEnteros                 As Integer        'Numero de cifras enteras del campo definida en el diccionario
Private m_iFieldDecimales               As Integer        'Numero de cifras decimales del campo definida en el diccionario
Private m_iFieldType                    As Integer        'identifica el tipo de datos
Private m_iMaxLenght                    As Integer        'Numero Max de caracteres que pueden ser introducidos
Private m_vContextMenu                  As Variant        'nombre del menu contextual
Private m_sMenuKeyAdmin                 As String         'clave del menu para la administracion
Private m_bFindIsEnabled                As Boolean        'es true si tiene definida una query en el diccionario
Private m_bAdminIsEnabled               As Boolean        'es true si la administracion es posible (lo que no significa que el usuario este habilitado)
Private m_sFindDataLink                 As String         'entrada del diccionario para la búsqueda
Private m_sFindControlsList             As String         'Lista de nombres de controles para la búsqueda
Private m_sFindBoundFieldList           As String         'Lista de nombres de campos (de la Query) usados para actualizar los controles de FindControlsList
Private m_sFindWhere                    As String         'una clausula where por si se quiere filtrar la query del diccionario
Private m_FindAlternate                 As Variant        'Es posible definir una busqueda distinta en fase RunTime. Si existe, es la que será ejecutada
Private m_vFieldValoresPosibles         As Variant        'Valores Posibles
Private m_vFormatear                    As eFormatField   'formateo de Campo
Private m_sMascaraFormato               As String         'Si Formatear es True, aqui va la máscara de formato
Private m_lhWndParent                   As Long           'hwnd de Control PowerMask
Private m_bChanged                      As Boolean        'Indicador de modificacion del control
Private m_bControlEdicion               As Boolean        'Indica si el control permite la edicion
Private m_vDefaultValue                 As Variant        'Valor por Defecto atribuido al control
Private m_lBackColor                    As Long           'BackColor definido del control
Private m_vTag                          As Variant        'valor de cualquier dato adicional
Private m_vOriginalValue                As Variant        'valor original

Public Property Get hWnd() As Long

   hWnd = m_lhWnd

End Property

Public Property Let hWnd(ByVal lhWnd As Long)

   m_lhWnd = lhWnd

End Property

Public Property Get Ctrl() As Object

   Set Ctrl = m_vCtrl

End Property

Public Property Set Ctrl(ByVal vCtrl As Object)

   Set m_vCtrl = vCtrl

End Property

Public Property Get FindAlternate() As Variant

   FindAlternate = m_FindAlternate

End Property

Public Property Let FindAlternate(ByVal vData As Variant)

   m_FindAlternate = vData

End Property

Public Property Get TypeOfControl() As String

   TypeOfControl = m_sTypeOfControl

End Property

Public Property Let TypeOfControl(ByVal sTypeOfControl As String)

   m_sTypeOfControl = sTypeOfControl

End Property

Public Property Get Validar() As ValidEnum

   Validar = m_eValidar

End Property

Public Property Let Validar(ByVal eValidar As ValidEnum)

   m_eValidar = eValidar

End Property

Public Property Get BoundField() As String

   BoundField = m_sBoundField

End Property

Public Property Let BoundField(ByVal sBoundField As String)

   m_sBoundField = sBoundField

End Property

Public Property Get ObjectProperty() As String

   ObjectProperty = m_sObjectProperty

End Property

Public Property Let ObjectProperty(ByVal sObjectProperty As String)

   m_sObjectProperty = sObjectProperty

End Property

Public Property Get MayuscMinusc() As Variant

   MayuscMinusc = m_vMayuscMinusc

End Property

Public Property Let MayuscMinusc(ByVal vMayuscMinusc As Variant)

   m_vMayuscMinusc = vMayuscMinusc

End Property

Public Property Get FieldDimension() As Integer

   FieldDimension = m_iFieldDimension

End Property

Public Property Let FieldDimension(ByVal iFieldDimension As Integer)

   m_iFieldDimension = iFieldDimension

End Property

Public Property Get FieldEnteros() As Integer

   FieldEnteros = m_iFieldEnteros

End Property

Public Property Let FieldEnteros(ByVal iFieldEnteros As Integer)

   m_iFieldEnteros = iFieldEnteros

End Property

Public Property Get FieldDecimales() As Integer

   FieldDecimales = m_iFieldDecimales

End Property

Public Property Let FieldDecimales(ByVal iFieldDecimales As Integer)

   m_iFieldDecimales = iFieldDecimales

End Property

Public Property Get FieldType() As Integer

   FieldType = m_iFieldType

End Property

Public Property Let FieldType(ByVal iFieldType As Integer)

   m_iFieldType = iFieldType

End Property

Public Property Get MaxLenght() As Integer

   MaxLenght = m_iMaxLenght

End Property

Public Property Let MaxLenght(ByVal iMaxLenght As Integer)

   m_iMaxLenght = iMaxLenght

End Property

Public Property Get ContextMenu() As Variant

   ContextMenu = m_vContextMenu

End Property

Public Property Let ContextMenu(ByVal vContextMenu As Variant)

   m_vContextMenu = vContextMenu

End Property

Public Property Get MenuKeyAdmin() As String

   MenuKeyAdmin = m_sMenuKeyAdmin

End Property

Public Property Let MenuKeyAdmin(ByVal sMenuKeyAdmin As String)

   m_sMenuKeyAdmin = sMenuKeyAdmin

End Property

Public Property Get FindIsEnabled() As Boolean

   FindIsEnabled = m_bFindIsEnabled

End Property

Public Property Let FindIsEnabled(ByVal bFindIsEnabled As Boolean)

   m_bFindIsEnabled = bFindIsEnabled

End Property

Public Property Get AdminIsEnabled() As Boolean

   AdminIsEnabled = m_bAdminIsEnabled

End Property

Public Property Let AdminIsEnabled(ByVal bAdminIsEnabled As Boolean)

   m_bAdminIsEnabled = bAdminIsEnabled

End Property

Public Property Get FindDataLink() As String

   FindDataLink = m_sFindDataLink

End Property

Public Property Let FindDataLink(ByVal sFindDataLink As String)

   m_sFindDataLink = sFindDataLink

End Property

Public Property Get FindControlsList() As String

   FindControlsList = m_sFindControlsList

End Property

Public Property Let FindControlsList(ByVal sFindControlsList As String)

   m_sFindControlsList = sFindControlsList

End Property

Public Property Get FindBoundFieldList() As String

   FindBoundFieldList = m_sFindBoundFieldList

End Property

Public Property Let FindBoundFieldList(ByVal sFindBoundFieldList As String)

   m_sFindBoundFieldList = sFindBoundFieldList

End Property

Public Property Get FindWhere() As String

   FindWhere = m_sFindWhere

End Property

Public Property Let FindWhere(ByVal sFindWhere As String)

   m_sFindWhere = sFindWhere

End Property

Public Property Get FieldValoresPosibles() As Variant

   FieldValoresPosibles = m_vFieldValoresPosibles

End Property

Public Property Let FieldValoresPosibles(ByVal vFieldValoresPosibles As Variant)

   m_vFieldValoresPosibles = vFieldValoresPosibles

End Property

Public Property Get Formatear() As eFormatField

   Formatear = m_vFormatear

End Property

Public Property Let Formatear(ByVal eFormatear As eFormatField)

   m_vFormatear = eFormatear

End Property

Public Property Get MascaraFormato() As String

   MascaraFormato = m_sMascaraFormato

End Property

Public Property Let MascaraFormato(ByVal sMascaraFormato As String)

   m_sMascaraFormato = sMascaraFormato

End Property

Public Property Get hWndParent() As Long

   hWndParent = m_lhWndParent

End Property

Public Property Let hWndParent(ByVal lhWndParent As Long)

   m_lhWndParent = lhWndParent

End Property

Public Property Get Changed() As Boolean

   Changed = m_bChanged

End Property

Public Property Let Changed(ByVal bChanged As Boolean)

   m_bChanged = bChanged

End Property

Public Property Get ControlEdicion() As Boolean

   ControlEdicion = m_bControlEdicion

End Property

Public Property Let ControlEdicion(ByVal bControlEdicion As Boolean)

   m_bControlEdicion = bControlEdicion

End Property

Public Property Get DefaultValue() As Variant

   DefaultValue = m_vDefaultValue

End Property

Public Property Let DefaultValue(ByVal vDefaultValue As Variant)

   m_vDefaultValue = vDefaultValue

End Property

Public Property Get BackColor() As Long

   BackColor = m_lBackColor

End Property

Public Property Let BackColor(ByVal lBackColor As Long)

   m_lBackColor = lBackColor

End Property

Public Property Get Tag() As Variant

   Tag = m_vTag

End Property

Public Property Let Tag(ByVal vTag As Variant)

   m_vTag = vTag

End Property

Public Property Get OriginalValue() As Variant

   OriginalValue = m_vOriginalValue

End Property

Public Property Let OriginalValue(ByVal vOriginalValue As Variant)

   m_vOriginalValue = vOriginalValue

End Property
