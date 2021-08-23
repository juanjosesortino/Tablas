VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsControlItem"
Attribute VB_Ext_KEY = "Member0" ,"clsControlItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : clsControlItems
' DateTime  : 28/11/2002 12:05
' Author    : Algoritmo
' Purpose   : Junto con la clase clsControlItem permite implementar colecciones con
'             eventos
'---------------------------------------------------------------------------------------
Option Explicit

Private m_oCol As Collection

' indice de eventos que puedes ser raiseados por los controles
Public Enum EVENT_INDEXES
  cf_ItemGotFocus
  cf_ItemLostFocus
  cf_ItemValidate
  cf_ItemSetText                    ' (solo para el PowerMask)
  cf_ItemChange
End Enum

' eventos que serán raiseados al client
Event ItemGotFocus(ByVal nIndex&, ByVal sKey$)
Event ItemLostFocus(ByVal nIndex&, ByVal sKey$)
Event ItemValidate(ByVal nIndex&, ByVal sKey$)
Event ItemSetText(ByVal nIndex&, ByVal sKey$)
Event ItemChange(ByVal nIndex&, ByVal sKey$)

Public Property Get Item(vntIndexKey As Variant) As clsControlItem
Attribute Item.VB_UserMemId = 0
  Set Item = m_oCol(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = m_oCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    m_oCol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_oCol.[_NewEnum]
End Property
Private Sub Class_Initialize()
   Set m_oCol = New Collection
    
  ' agrego una referencia muda (puntero ilegal) al arreglo en el modulo MValidatePointer.BAS
  MValidatePointer.AddToLookupList ObjPtr(Me)
    
End Sub
Private Sub Class_Terminate()
    Set m_oCol = Nothing
    
  ' remuevo el puntero almacenado en el initialize event de la clase
  MValidatePointer.RemoveFromLookupList ObjPtr(Me)
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Add
' DateTime  : 28/11/2002 12:10
' Author    : Algoritmo
' Purpose   : este metodo agrega una nueva instancia de la clase Item
'---------------------------------------------------------------------------------------
'
Public Sub Add(oControlItem As clsControlItem)
Dim sItemsCount As String
  
  ' obtengo el numero de items en la coleccion (para generar una clave unica)
  sItemsCount = CStr(m_oCol.Count + 1)
    
  ' agrego el Item a la coleccion....
  With oControlItem
      ' paso al Item el puntero a esta clase
      .CollectionPointer = ObjPtr(Me)
      
      ' seteo la Clave del Item
      .Key = oControlItem.Name & sItemsCount
      
      ' agrego el Item dentro de la colección
      m_oCol.Add oControlItem, .Key
  End With
  
  ' destruyo la referencia local al Item
  Set oControlItem = Nothing
  
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FireEvent
' DateTime  : 28/11/2002 12:15
' Author    : Algoritmo
' Purpose   : este metodo es llamado desde el Item para raisear eventos al Client
'---------------------------------------------------------------------------------------
'
Friend Sub FireEvent(ByVal nEvent As EVENT_INDEXES, ByVal sKey$, ByVal lpObject&)
Dim i&, nIndex&
  
  ' determino el indice del objeto llamante.
  For i = 1 To m_oCol.Count
    If ObjPtr(m_oCol(i)) = lpObject Then
      nIndex = i
      Exit For
    End If
  Next
  
  If nIndex > 0 Then
    ' raiseo el evento
    Select Case nEvent
      Case cf_ItemGotFocus: RaiseEvent ItemGotFocus(ByVal nIndex&, ByVal sKey$)
      Case cf_ItemLostFocus: RaiseEvent ItemLostFocus(ByVal nIndex&, ByVal sKey$)
      Case cf_ItemValidate: RaiseEvent ItemValidate(ByVal nIndex&, ByVal sKey$)
      Case cf_ItemSetText: RaiseEvent ItemSetText(ByVal nIndex&, ByVal sKey$)
      Case cf_ItemChange: RaiseEvent ItemChange(ByVal nIndex&, ByVal sKey$)
    End Select
  End If
  
End Sub

