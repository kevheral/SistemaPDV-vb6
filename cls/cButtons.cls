VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Colecci�n para los botones
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mColItems As New Collection

Public Event AddButton()
Public Event ClearButtons()
Public Event RemoveButton(lIndex As Integer)

Private mUC As ucMenu

Sub Init(pUcMenu As ucMenu)
Attribute Init.VB_MemberFlags = "40"
    Set mUC = pUcMenu
End Sub

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Agregar un nuevo bot�n a la colecci�n
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(sCaption As Variant, Optional sKey As Variant, Optional sToolTip As String = "") As cButton
    
    On Error GoTo err_handler
    
    ' nuevo bot�n
    Dim mItem As New cButton
    
    ' .. llenar los datos
    With mItem
    
        If mItem.ucMenu Is Nothing Then Set mItem.ucMenu = mUC
        
        .FlagMod = True
        .Index = Count + 1
        .Caption = sCaption
        .Enabled = True
        .ToolTipText = sToolTip
        
        ' .. comprobar si se pas� la clave
        If Not IsMissing(sKey) Then
           ' .. comprobar que sea un n�mero
           If IsNumeric(sKey) Then
              Set mItem = Nothing
              Err.Raise 35603, "Buttons", "La clave no es v�lida."
           Else
              .Key = sKey
              mColItems.Add mItem, CStr(sKey) ' agregar con clave
           End If
        Else
           mColItems.Add mItem ' agregar bot�n sin clave
        End If
        .FlagMod = False
        RaiseEvent AddButton
                
    End With
    ' retornar el bot�n ( Para poder referenciarlo >> Set xBtn = Menu.Buttons.add(datos))
    Set Add = mItem
Exit Function
err_handler:
DoEvents
Err.Raise Err.Number, "Buttons", Err.Description

End Function


' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Retornar la Cantidad de elementos actuales de la colecci�n
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Count() As Long
    
    Count = mColItems.Count

Exit Function
err_handler:
DoEvents
Err.Raise Err.Number, "Buttons", Err.Description
End Function

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Eliminar todos los Botones de la colecci�n
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
On Error GoTo err_handler
    
    ' mientras haya ....
    Do Until mColItems.Count = 0
       mColItems.Remove mColItems.Count ' .. eliminar el �ltimo hasta que no haya mas
    Loop
    
    ' Eliminar el SelectedItem, la colecci�n y los botones del UC
    RaiseEvent ClearButtons
    
    
Exit Sub
err_handler:
DoEvents
Err.Raise Err.Number, "Buttons", Err.Description
End Sub

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Eliminar un bot�n por el �ndice o por la clave
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Remove(ByVal vValue As Variant)
    
    On Error GoTo err_handler
    
    ' comprobar que se inicializ� la colecci�n
    If mColItems Is Nothing Then Exit Sub
    
    With mColItems
        ' comprobar que haya botones
        If .Count > 0 Then
            ' Recuperar el Index antes de eliminarlo
            Dim lIndex As Integer
            lIndex = mColItems(vValue).Index
            ' eliminarlo de la colecci�n
            Call .Remove(vValue)
           'Descargar todos los botones a partir de este �ndice
            RaiseEvent RemoveButton(lIndex)
        Else
           DoEvents
           Err.Raise 35600, "Buttons", "El �ndice est� fuera de los l�mites"
        End If
    End With
    
Exit Sub
err_handler:
DoEvents
Err.Raise Err.Number, "Buttons", Err.Description
End Sub

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Retornar el elemento de la colecci�n ( el bot�n ) por el indice o por la clave
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function item(ByVal vValue As Variant) As cButton
Attribute item.VB_UserMemId = 0
    
    On Error GoTo err_handler
    
    Set item = mColItems.item(vValue)
    
Exit Function
err_handler:
DoEvents
Err.Raise Err.Number, "Buttons", Err.Description
Set item = Nothing

End Function

Private Sub Class_Terminate()
    If Not mColItems Is Nothing Then Set mColItems = Nothing
    If Not mUC Is Nothing Then Set mUC = Nothing
End Sub
