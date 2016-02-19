VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Variables locales
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mKey                As String
Private mCaption            As String
Private mIndex              As Integer
Private mToolTipText        As String
Private mEnabled            As Boolean
Private mSelected           As Boolean
Private mFlagMod            As Boolean
Attribute mFlagMod.VB_VarMemberFlags = "40"
Private mUcMenu             As ucMenu

Event Change(lIndex As Integer, PropertyName As String, Value As Variant)

Property Get ucMenu() As ucMenu
Attribute ucMenu.VB_MemberFlags = "40"
    Set ucMenu = mUcMenu
End Property

Property Set ucMenu(objValue As ucMenu)
    Set mUcMenu = objValue
End Property

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ' Valor del item Seleccionado
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Selected() As Boolean
    Selected = mSelected
End Property

Property Let Selected(bValue As Boolean)
    mSelected = bValue
    If mFlagMod = False Then
       If Not mUcMenu Is Nothing Then
          Call mUcMenu.mUc_Click(Index)
       Else
          RaiseEvent Change(Index, "Selected", Nothing)
       End If
    End If
End Property

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ' Valor Enabled para cada boton
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Property Let Enabled(bValue As Boolean)
    mEnabled = bValue
    If mFlagMod = False Then
       If Not mUcMenu Is Nothing Then
          Call mUcMenu.mEnabledButtonByIndex(mIndex, bValue)
       Else
          RaiseEvent Change(mIndex, "Enabled", bValue)
       End If
    End If
End Property

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ' Flag para modificar las propiedades, pero no lanzar eventos, y para que el Key y el Index sea de solo lectura
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get FlagMod() As Boolean
    FlagMod = mFlagMod
End Property

Friend Property Let FlagMod(Value As Boolean)
    mFlagMod = Value
End Property

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ' ToolTip para cada botón
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ToolTipText() As String
    ToolTipText = mToolTipText
End Property

Property Let ToolTipText(sValue As String)
    mToolTipText = sValue
    If FlagMod = False Then
       If Not mUcMenu Is Nothing Then
          Call mUcMenu.mChangeToolTip(mIndex, sValue)
       Else
          RaiseEvent Change(mIndex, "ToolTipText", sValue)
       End If
    End If
End Property

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ' índice del botón
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Index() As Integer
    Index = mIndex
End Property

Property Let Index(iValue As Integer)
    If mFlagMod = False Then
       Err.Raise 383, "Button", "La propiedad es de solo lectura"
    End If
    mIndex = iValue
End Property

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ' Clave del botón
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Key() As String
    Key = mKey
End Property

Property Let Key(sKey As String)
    If mFlagMod = False Then
       Err.Raise 383, "Button", "La propiedad es de solo lectura"
    End If
    mKey = sKey
End Property

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ' Caption del botón
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get Caption() As String
    Caption = mCaption
End Property

Property Let Caption(sCaption As String)
    mCaption = sCaption
    If mFlagMod = False Then
       If Not mUcMenu Is Nothing Then
          Call mUcMenu.mModifyButtons(mIndex)
       Else
          RaiseEvent Change(mIndex, "Caption", sCaption)
       End If
    End If
End Property

Private Sub Class_Initialize()
    If Not mUcMenu Is Nothing Then Set mUcMenu = Nothing
End Sub
