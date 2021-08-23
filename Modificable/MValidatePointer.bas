Attribute VB_Name = "MValidatePointer"
'---------------------------------------------------------------------------------------
' Module    : MValidatePointer
' DateTime  : 28/11/2002 15:23
' Author    : Algoritmo
' Purpose   : Este modulo permite raisear eventos desde objetos contenidos
'             en colecciones o arreglos sin crear referencias circulares
'---------------------------------------------------------------------------------------
Option Explicit

' arreglo que manitiene referencias (punteros) de los objetos que queremos controlar
Private m_alpCollections() As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

Public Sub AddToLookupList(ByVal lpObject&)
   ' encargo esta tarea al FindPointer
   Call FindPointer(lpObject, True)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RemoveFromLookupList
' DateTime  : 28/11/2002 15:27
' Author    : Algoritmo
' Purpose   : remueve del arreglo el puntero. Esta Sub es llamado durante el evento Terminate
'             de la clase "colección"
'---------------------------------------------------------------------------------------
'
Public Sub RemoveFromLookupList(ByVal lpObject&)
  
  Dim nItem&, nUbound&
  
  ' busco la ubicacion del puntero dentro del array
  nItem = IsPointerValid(lpObject)
  
  ' si fue encontrado, lo remuevo y hago un shift del resto de los items ....
  If nItem > (-1) Then
    nUbound = UBound(m_alpCollections)
    
    If nItem < nUbound Then
      ' grabo todos los items debajo del item y los muevo hacia arriba
      CopyMemory ByVal VarPtr(m_alpCollections(nItem)), ByVal VarPtr(m_alpCollections(nItem + 1)), (nUbound - nItem) * 4&
    End If
    
    
    If nUbound Then
      ' si hay elementos en el arreglo, los preservo
      ReDim Preserve m_alpCollections(nUbound - 1) As Long
    Else
      ' si es el ultimo elemento del array, me aseguro que el valor sea 0
      ReDim m_alpCollections(0) As Long
      m_alpCollections(0) = 0
    End If
  End If
  
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsPointerValid
' DateTime  : 28/11/2002 15:39
' Author    : Algoritmo
' Purpose   : controlo si el objeto ya existe en el arreglo. Si el objeto ya no existe
'             en el arreglo, entonces el objeto no existe
'---------------------------------------------------------------------------------------
'
Public Function IsPointerValid(ByVal lpObject&) As Long
   IsPointerValid = FindPointer(lpObject)
End Function

'---------------------------------------------------------------------------------------
' Procedure : FindPointer
' DateTime  : 28/11/2002 15:41
' Author    : Algoritmo
' Purpose   : función que provee un lookup veloz de los punteros del array
'---------------------------------------------------------------------------------------
'
Private Function FindPointer(ByVal lpObject&, Optional ByVal bAddIfNotFound As Boolean) As Long
  
  Static bInitialized As Boolean
  
  Dim i&, nLow&, nHigh&, nUbound&
  
  ' me aseguro que el arreglo este inicializado
  If bInitialized = False Then
    If bAddIfNotFound Then
      GoTo AddFirsItem
    Else
      FindPointer = (-1)
    End If
  End If
  
    
  nHigh = UBound(m_alpCollections)
  
  ' loop dentro del array buscando el puntero
  ' el arreglo esta ordenado en orden numerico de modo tale que
  ' podemos realizar la bussqueda velos
  Do
    ' divide y venceras!  Cada vez que loopeamos, divido la diferencia entre los
    ' ultimos items checkeados y busco entre 2 indices.  Esto es mucho mas rapido
    ' que lopear a trevz de la entera listalist aun con arreglos ordenados.
    i = nLow + ((nHigh - nLow) \ 2)
    
    ' ver como sKey relaciona al corriente index....
    Select Case m_alpCollections(i)
      Case Is = lpObject
        FindPointer = i
        Exit Do
        
      Case Is > lpObject: nHigh = i - 1
      Case Is < lpObject: nLow = i + 1
    End Select

    ' si el límite bajo de la búsqueda ha llegado a ser mayor que el límite superior,
    ' el itme no existe en el array. si bAddIfNotFound esta seteado ? el nuevo item
    ' sera agregado caso contrario devuelve Not Found Value
    If nLow > nHigh Then
      If bAddIfNotFound Then
      
AddFirsItem:

        ' me fijo si ha sido inicializado
        If Not bInitialized Then
          bInitialized = True
          ReDim m_alpCollections(0) As Long
        Else
          
          If m_alpCollections(0) <> 0 Then
            ReDim Preserve m_alpCollections(UBound(m_alpCollections) + 1) As Long
          End If
        
          nUbound = UBound(m_alpCollections)
          
          ' me fijo si debemos agregar este item debajo o arriba del indice 'i'
          Select Case m_alpCollections(i)
            Case Is < lpObject: i = i + 1
            Case Is > lpObject: i = i '<- incluido para autodocumentación
          End Select
          
          If i > nUbound Then i = nUbound
          
          If i < nUbound Then
            ' grabo todos los items debajo del item y los muevo hacia abajo
            CopyMemory ByVal VarPtr(m_alpCollections(i + 1)), ByVal VarPtr(m_alpCollections(i)), (nUbound - i) * 4&
          End If ' i < nUbound
        End If
        
        ' pongo el nuevo puntero en la posicion correcta
        m_alpCollections(i) = lpObject
        
      End If ' bAddIfNotFound
      
      
      ' return KEY_NOT_FOUND para informar al llamador que no ha sido encontrado
      FindPointer = (-1)
      Exit Do
    End If ' nLow > nHigh
  
  Loop
  
End Function

