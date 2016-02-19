VERSION 5.00
Begin VB.UserControl ucMenu 
   Alignable       =   -1  'True
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   ScaleHeight     =   795
   ScaleWidth      =   8235
   Begin VB.PictureBox picScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6840
      ScaleHeight     =   615
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      Begin VENTOS.ucBtnSkin ucScroll 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         Skin            =   0
         Caption         =   ""
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   0   'False
         CaptionAlign    =   0
         CaptionMargin   =   0
         ButtonType      =   0
         Object.ToolTipText     =   ""
         UseUnderLineMouseUp=   0   'False
         UseUnderLineMouseCheck=   0   'False
      End
      Begin VENTOS.ucBtnSkin ucScroll 
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         Skin            =   0
         Caption         =   ""
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   0   'False
         CaptionAlign    =   0
         CaptionMargin   =   0
         ButtonType      =   0
         Object.ToolTipText     =   ""
         UseUnderLineMouseUp=   0   'False
         UseUnderLineMouseCheck=   0   'False
      End
      Begin VB.Shape shapeScroll 
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4920
      Top             =   240
   End
   Begin VB.PictureBox picBtns 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   120
      ScaleHeight     =   600
      ScaleWidth      =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1800
      Begin VENTOS.ucBtnSkin uc 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Tag             =   "main"
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Skin            =   0
         Caption         =   ""
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   0   'False
         CaptionAlign    =   0
         CaptionMargin   =   0
         ButtonType      =   0
         Object.ToolTipText     =   ""
         UseUnderLineMouseUp=   0   'False
         UseUnderLineMouseCheck=   0   'False
      End
   End
End
Attribute VB_Name = "ucMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

' >> Descripción : Control para usar Menú con Scroll y Skins
' >> Autor       : Luciano Lodola - http://www.recursosvisualbasic.com.ar/

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Apis, Constantes , vars, tipos, enums
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()


Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

' Enum para mover el scroll al principio y al final
Enum eDir
    [First] = 0
    [Last] = 1
End Enum

' Enum de opciones para buscar un botón y seleccionarlo
Enum eOptFindButton
    [byIndex] = 0
    [byKey] = 1
    [byCaption] = 2
End Enum

Private WithEvents mParent                As Form
Attribute mParent.VB_VarHelpID = -1

Private mCurrentDirScroll                 As Integer        ' índice actual del botón de scroll
Private mSkin                             As eSkin          ' skin actual
Private WithEvents mButtons               As cButtons       ' Botones
Attribute mButtons.VB_VarHelpID = -1
Private WithEvents mSelectedItem          As cButton        ' Botón seleccionado
Attribute mSelectedItem.VB_VarHelpID = -1
Private mMarginButton                     As Long           ' margen del botón
Private mMarginCaption                    As Long           ' Margen del texto del botón
Private mShowFocusRect                    As Boolean        ' mostrar o no el rectángulo de enfoque
Private mEnabled                          As Boolean        ' Habilitar / Deshabilitar el menú
Private mToolTipText                      As String
Private mUseUnderLineMouseUp              As Boolean        ' Subrayar el caption en mouse UP
Private mUseUnderLineMouseCheck           As Boolean        ' Subrayar el caption cuando está seleccionado
Private mSmallChangeScroll                As Integer        ' valor para el movimiento del scroll

Private mForeColorNormal                  As OLE_COLOR      ' Colores de fuente
Private mForeColorUp                      As OLE_COLOR
Private mForeColorDown                    As OLE_COLOR
Private mForeColorDisabled                As OLE_COLOR
Private mForeColorCheck                   As OLE_COLOR
Private mUseCustomForeColor               As Boolean
Private mRedraw                           As Boolean        ' Habilitar / deshabilitar el repintado del UC
Private mEnabledDragMenu                  As Boolean        ' Habilitar / deshabilitar el Drag Drop


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Eventos
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Eventos de los botones de scroll
Public Event StopScroll(ByVal iDir As Integer)
Public Event Scroll(ByVal iDir As Integer)
Public Event ScrollKeyPress(ByVal ButtonIndex As Integer, KeyAscii As Integer)
Public Event ScrollKeyDown(ByVal ButtonIndex As Integer, KeyCode As Integer, Shift As Integer)
Public Event ScrollKeyUp(ByVal ButtonIndex As Integer, KeyCode As Integer, Shift As Integer)
Public Event ScrollMouseDown(ByVal ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ScrollMouseMove(ByVal ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ScrollMouseUp(ByVal ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ScrollMouseOver(ByVal ButtonIndex As Integer)
Public Event ScrollMouseOut(ByVal ButtonIndex As Integer)
Public Event ScrollClick(ByVal ButtonIndex As Integer)
Public Event ScrollBeforeClick(ByVal ButtonIndex As Integer)

'Eventos del contenedor del scroll
Public Event ScrollContainerMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ScrollContainerMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ScrollContainerMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Eventos de los botones
Public Event ButtonGotFocus(ByVal ButtonIndex As Integer)
Public Event ButtonLostFocus(ByVal ButtonIndex As Integer)
Public Event ButtonKeyPress(ByVal ButtonIndex As Integer, KeyAscii As Integer)
Public Event ButtonKeyDown(ByVal ButtonIndex As Integer, KeyCode As Integer, Shift As Integer)
Public Event ButtonKeyUp(ByVal ButtonIndex As Integer, KeyCode As Integer, Shift As Integer)
Public Event ButtonMouseDown(ByVal ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonMouseMove(ByVal ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonMouseUp(ByVal ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonMouseOver(ByVal ButtonIndex As Integer)
Public Event ButtonMouseOut(ByVal ButtonIndex As Integer)
Public Event ButtonClick(ByVal ButtonIndex As Integer, Button As cButton)
Public Event ButtonBeforeClick(ByVal ButtonIndex As Integer)

'Eventos del menú ( picBtns y UC )
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()
Public Event DblClick()
Public Event Resize()
Public Event Paint()


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'===============================================================================================================
' Fin de eventos
'===============================================================================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

' Funciones - Subs

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================



' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Deshabilitar / Habilitar el botón por el índice
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function mEnabledButtonByIndex(lIndex As Integer, bValue As Boolean) As Boolean
Attribute mEnabledButtonByIndex.VB_MemberFlags = "40"
    With mButtons(lIndex)
        .FlagMod = True
        .Enabled = bValue
        .FlagMod = False
    End With
    uc(lIndex).Enabled = bValue
    mEnabledButtonByIndex = True
End Function

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cambiar valores de los botones
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mSetPropertyButtons(PropertyName As String, Value As Variant)
    
    Dim xBtn As Control
    For Each xBtn In Controls
        
        If TypeOf xBtn Is ucBtnSkin Then
           Call CallByName(xBtn, PropertyName, VbLet, Value)
        End If
        If Not Ambient.UserMode Then
           uc(0).Width = picBtns.TextWidth(uc(0).Caption) + (mMarginCaption * 2)
        End If
    Next
End Sub


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Refrescar los botones y redimensionar
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Refresh()
    Dim xBtn As Control
    Redraw = False
    For Each xBtn In Controls
        If TypeOf xBtn Is ucBtnSkin Then
           xBtn.Refresh
        End If
    Next
    Redraw = True
    UserControl_Resize
End Sub


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cambiar ToolTipText
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function mChangeToolTip(lIndex As Integer, sValue As String) As Boolean
Attribute mChangeToolTip.VB_MemberFlags = "40"
    
    With mButtons(lIndex)
        .FlagMod = True
        .ToolTipText = sValue
        .FlagMod = False
    End With
    uc(lIndex).ToolTipText = sValue
    mChangeToolTip = True
    
End Function


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cambiar el Caption
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function mChangeCaption(lIndex As Integer, sValue As String) As Boolean
Attribute mChangeCaption.VB_MemberFlags = "40"
    With mButtons(lIndex)
        .FlagMod = True
        .Caption = sValue
        .FlagMod = False
    End With
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Dibujar el Skin en el PictureBox contenedor de los botones, en el de Scroll y cambiar el color del shape
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub mDrawSkinContainer()
Attribute mDrawSkinContainer.VB_MemberFlags = "40"
    'UC
    With UserControl
        .BackColor = uc(0).GetSkinsColors(Normal_Skin)
    End With
    ' Contenedor de los botones
    With picBtns
        Call uc(0).DrawSkin(.hdc, .ScaleWidth / 15, .ScaleHeight / 15, TS_NORMAL)
        .Refresh
    End With
    'Contenedor de las flechas de Scroll
    With picScroll
        Call uc(0).DrawSkin(.hdc, .ScaleWidth / 15, .ScaleHeight / 15, TS_NORMAL)
        .Refresh
    End With
    shapeScroll.BorderColor = uc(0).GetSkinsColors(Border_Normal_Skin)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desplazar el Scroll al principio y al final
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MoveScroll(lDir As eDir)
    If picScroll.Visible = False Then Exit Sub
    Select Case lDir
       ' Mostrar primer botón
       Case 0
         If ucScroll(0).Enabled Then
            picBtns.Left = 0
            Call mCheckScroll
            Call mDrawSkinContainer
         End If
       ' Mostrar últimobotón
       Case 1
         If ucScroll(1).Enabled Then
            picBtns.Left = (-picBtns.Width) + (UserControl.Width - picScroll.Width)
            Call mCheckScroll
            Call mDrawSkinContainer
         End If
    End Select
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' función para Buscar un botón por el índice, por el caption o por la clave
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function FindButton( _
    ByVal pValue As Variant, _
    lOpt As eOptFindButton, _
    Optional bSelectedItem As Boolean = True, _
    Optional bRaiseEventButtonClick As Boolean = True) As Boolean
    
    On Error GoTo error_handler
    
        
    Me.Redraw = False
    
    Dim i As Integer
    Dim lWidth As Long
    Dim bFind As Boolean
    
    ' recorrer los botones
    For i = 1 To uc.Count - 1
        
        ' Por el Index
        ' ''''''''''''''''''''''''''''''''''''''''''
        If lOpt = byIndex Then
           If i = (pValue) Then
              bFind = True
              Exit For
           End If
        End If
        
        ' Por el texto
        ' ''''''''''''''''''''''''''''''''''''''''''
        If lOpt = byCaption Then
           If Trim(LCase(uc(i).Caption)) = Trim(LCase(pValue)) Then
              bFind = True
              Exit For
           End If
        End If
        
        ' Buscart Por el Key
        ' ''''''''''''''''''''''''''''''''''''''''''
        If lOpt = byKey Then
           If Trim(LCase(uc(i).Tag)) = Trim(LCase(pValue)) Then
              bFind = True
              Exit For
           End If
        End If
        
        ' Mientras no se encuentre, almacenar el valor del Left
        lWidth = lWidth + uc(i).Width + mMarginButton
        
    Next
    
    ' si se encontró, cambiar la posiciuón Left del picture contenedor
    If bFind Then
        With picBtns
           .Left = -lWidth
        End With
        ' Comprobar que no se pase
        If (picBtns.Width - lWidth) < (UserControl.Width - picScroll.Width) Then
            picBtns.Left = -(picBtns.Width - UserControl.Width + picScroll.Width)
        End If
        UserControl_Resize
    End If
    
    ' Comprobar si se pasó el parámetro para seleccionar el item
    If bSelectedItem And bFind Then
       If lOpt = byIndex Then
          Call Me.SelectedByIndex(i, bRaiseEventButtonClick)
       End If
       If lOpt = byCaption Then
          Call Me.SelectedByCaption(uc(i).Caption, bRaiseEventButtonClick)
       End If
       If lOpt = byKey Then
          Call Me.SelectedByKey(uc(i).Tag, bRaiseEventButtonClick)
       End If
    End If
    
    Me.Redraw = True
    FindButton = bFind
    
    Exit Function
error_handler:
    Me.Redraw = True
    Err.Raise Err.Number, "UcMenu", Err.Description
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Función para Desplazar el Scroll mientras se esté presionando las flechas
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mSetScroll(pDir As Integer)
    
    With picBtns
    
        Select Case pDir
            ' Desplazar a la derecha ( Botón izquierdo )
            ' ''''''''''''''''''''''''''''''''''''''
            Case 0
                
              If .Left < 0 Then
                 ' Mover
                 .Left = .Left + mSmallChangeScroll
                 RaiseEvent Scroll(0)
              Else
                 ' terminar
                 Timer1.Enabled = False
                 ' deshabilitar el botón de scroll
                 ucScroll(mCurrentDirScroll).Enabled = False
                 .Left = 0
                 RaiseEvent StopScroll(0)
              End If
              
            ' Desplazar a la Izquierda ( Botón derecho )
            ' ''''''''''''''''''''''''''''''''''''''
            Case 1
              ' Comprobar que no se pase
              If (.Left) > ((-.Width) + (UserControl.Width - picScroll.Width)) Then
                  ' Mover
                  .Left = .Left - mSmallChangeScroll
                  RaiseEvent Scroll(1)
              Else
                 ' terminar
                 Timer1.Enabled = False
                 ' deshabilitar el botón de scroll
                 ucScroll(mCurrentDirScroll).Enabled = False
                 .Left = (-.Width) + (UserControl.Width - picScroll.Width)
                 RaiseEvent StopScroll(1)
              End If
        End Select
    
    End With
End Sub



' Recuperar colores de los skins para poder asociarlos a otros controles del formulario
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetSkinsColors(lImgStateButton As ePosSkin) As Long
    GetSkinsColors = uc(0).GetSkinsColors(lImgStateButton)
End Function

Private Sub mButtons_ClearButtons()
    Set SelectedItem = Nothing
    Call mClearButtons
End Sub

Private Sub mButtons_RemoveButton(lIndex As Integer)
    Call mRemoveButton(lIndex)
End Sub

Private Sub mSelectedItem_Change(lIndex As Integer, PropertyName As String, Value As Variant)
    
    With mSelectedItem
         Buttons(lIndex).FlagMod = True
         Buttons(lIndex).Caption = .Caption
         Buttons(lIndex).ToolTipText = .ToolTipText
         Buttons(lIndex).Enabled = .Enabled
         Buttons(lIndex).Selected = .Selected
         Buttons(lIndex).Index = .Index
         Buttons(lIndex).Key = .Key
         Buttons(lIndex).FlagMod = False
    End With
    
    Select Case PropertyName
        Case "ToolTipText": Call mChangeToolTip(lIndex, CStr(Value))
        Case "Enabled": Call mEnabledButtonByIndex(lIndex, CBool(Value))
        Case "Caption": Call mModifyButtons(lIndex)
        Case "Selected": Call mUc_Click(lIndex)
    End Select
    
    
End Sub

' timer para el Scroll mientras se está presionando las flechas
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Timer1_Timer()
    Call mSetScroll(CInt(mCurrentDirScroll))
End Sub


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sub para agregar un botón desde el la función Add > cls Buttons
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mButtons_AddButton()
    Dim lIndex As Integer
    
    ' Indice del enuevo botón
    lIndex = uc.Count
    ' Crearlo
    Load uc(lIndex)
    
    ' Flag para Deshabilitar el cambio de las propiedades ( caption , enabled, ...)
    uc(lIndex).bFlagNoUpdateBtn = True
    
    ' Asignar propiedades
    With uc(lIndex)
        .ToolTipText = mButtons(lIndex).ToolTipText
        .Caption = mButtons(lIndex).Caption
        .Tag = mButtons(lIndex).Key
        .Enabled = mButtons(lIndex).Enabled
         
         mButtons(lIndex).FlagMod = True
         mButtons(lIndex).Index = lIndex
         mButtons(lIndex).FlagMod = False
         
        .Width = picBtns.TextWidth(mButtons(lIndex).Caption) + (mMarginCaption * 2)
        
        ' .. si es el primer botón, colocarle el Left inicial
        If (lIndex - 1) = 0 Then
            If mMarginButton > 60 Then
               .Left = 60
            Else
               .Left = mMarginButton
            End If
        ' Calcular el left para el botón
        Else
            .Left = (uc(lIndex - 1).Left + uc(lIndex - 1).Width) + mMarginButton
        End If
        .Visible = True
    End With
    
    ' .. si es el primer botón, seleccionarlo
    If lIndex = 1 Then Call Me.SelectedByIndex(1, False)
    
    ' restablecer la actualización
    uc(lIndex).bFlagNoUpdateBtn = False
    uc(lIndex).Refresh ' refrescarlo
    
    UserControl_Resize

End Sub


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Función que devuelve el ancho total de los botones + los margenes de separación
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function mGetButtonsWidth() As Long
    Dim lWidth As Long
    Dim xBtn As Control
    If uc.Count > 0 Then lWidth = 60
    For Each xBtn In Controls
        If LCase(xBtn.Name) = LCase("uc") Then
           lWidth = lWidth + xBtn.Width + mMarginButton
        End If
    Next
    mGetButtonsWidth = lWidth
    
End Function


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sub para ocultar y mostrar el scroll
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub mCheckScroll()
Attribute mCheckScroll.VB_MemberFlags = "40"
    
    With picScroll
        ' Si el ancho de botones es mayor al ancho del Usercontrol ... hacer visible el scroll
        If (mGetButtonsWidth) > UserControl.Width Then
            .Visible = True
        Else
            .Visible = False
        End If
    End With
    
    ' posicionar el contenedor de los botones
    With picBtns
        .Top = 0
        .Height = Height
        ' Si no está el scroll, colocar el contenedor de botones con el mismo tamaño del UC
        If picScroll.Visible = False Then
            .Left = 0
            .Width = UserControl.Width
        End If
        ' Correr el left depicBtns al redimensionar desde la derecha
        If (.Width - (-.Left)) < (UserControl.Width - picScroll.Width) Then
           picScroll.Visible = False
           '.Left = 0
           '.Width = UserControl.Width
           .Left = .Left + (UserControl.Width - picScroll.Width) - (.Width - (-.Left))
        End If
    End With
    
    ' Habilitar / deshabilitar los botones de scroll
    If picScroll.Visible Then
        With picBtns
            If (.Left) <= -(.Width - Width + picScroll.Width) Then
                ucScroll(1).Enabled = False
            Else
                ucScroll(1).Enabled = True
            End If
            If .Left >= 0 Then
                ucScroll(0).Enabled = False
            Else
                ucScroll(0).Enabled = True
            End If
        End With
    End If
    
End Sub


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Para hacer Drag al menú
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub DragMenu()
    Dim lret As Long
    UserControl.MousePointer = 5
    lret = SendMessage(UserControl.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Eliminar todos los botones cuando se ejecuta el método clear
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub mClearButtons()
Attribute mClearButtons.VB_MemberFlags = "40"
    Dim xBtn As Control
    For Each xBtn In Controls
        With xBtn
            If LCase(.Name) = LCase("uc") Then
                If LCase(.Tag) <> LCase("Main") Then
                   Unload xBtn
                End If
            End If
        End With
    Next
    UserControl_Resize
End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Modificar botones
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub mModifyButtons(Optional lStartIndex As Integer = 1)
Attribute mModifyButtons.VB_MemberFlags = "40"
    On Error GoTo error_handler
    Dim i As Integer
    Redraw = False
    
    ' cambiar los botones a partir de lStartIndex
    For i = lStartIndex To uc.Count - 1
        
        With uc(i)
            ' Deshbilitar la actualización
            .bFlagNoUpdateBtn = True
            
            ' Establecer las propiedades
            Set .Font = uc(0).Font
            .Caption = mButtons(i).Caption
            .Tag = mButtons(i).Key
            .ToolTipText = mButtons(i).ToolTipText
            .Enabled = mButtons(i).Enabled
            .Width = picBtns.TextWidth(mButtons(i).Caption) + (mMarginCaption * 2)
            ' .. si es el primero, establecer el Left inicial
            If (i - 1) = 0 Then
               .Left = 60
            Else
                ' calcular el siguiente valor del Left
               .Left = (uc(i - 1).Left + uc(i - 1).Width) + mMarginButton
            End If
            ' refrescar los cambios
            .bFlagNoUpdateBtn = False
            .Refresh
        End With
    Next
    Call UserControl_Resize
    Redraw = True
    Exit Sub
error_handler:
    Redraw = True
    Err.Raise Err.Number, "UcMenu", Err.Description

End Sub


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sub para eliminar los botones cuando se modifican
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub mRemoveButton(iValue As Integer)
Attribute mRemoveButton.VB_MemberFlags = "40"
    
    Dim i As Integer
    
    For i = iValue To uc.Count - 1
        Unload uc(i)
    Next
    
    With Me
        For i = iValue To .Buttons.Count
           'Agregar
            mButtons_AddButton
        Next
        
        If .Buttons.Count > 0 Then
           Call .SelectedByIndex(.Buttons.Count)
        Else
           Set SelectedItem = Nothing
        End If
    End With
    
    Call UserControl_Resize

End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Función para seleccionar un botón por el índice
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SelectedByIndex(iIndex As Integer, Optional bRaiseEventClick As Boolean = True) As Boolean
    If (iIndex = 0) Or (iIndex > uc.Count - 1) Then
       SelectedByIndex = False
    Else
        SelectedByIndex = True
        ' lanzar el evento Click para cambiar el value y otras propiedades
        Call mUc_Click(iIndex, bRaiseEventClick)
    End If
End Function

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' función para seleccionar un botón por el caption
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SelectedByCaption(sCaption As String, Optional bRaiseEventClick As Boolean = True) As Integer
    Dim xBtn As Control
    ' recorrer los botones
    For Each xBtn In Controls
        If LCase(xBtn.Name) = "uc" Then
            ' comprobar el caption
            If Trim(LCase(xBtn.Caption)) = Trim(LCase(sCaption)) Then
               ' lanzar evento clic para seleccionarlo
               Call mUc_Click(xBtn.Index, bRaiseEventClick)
               SelectedByCaption = xBtn.Index ' retornar el índice
               Exit For
            End If
        End If
    Next
End Function

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' función para seleccionar un botón por la clave
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SelectedByKey(sKey As String, Optional bRaiseEventClick As Boolean = True) As Integer
    Dim xBtn As Control
    ' recorrer los botones
    For Each xBtn In Controls
        If LCase(xBtn.Name) = "uc" Then
            If LCase(xBtn.Tag) = LCase(sKey) Then
               ' lanzar evento clic para seleccionarlo
               Call mUc_Click(xBtn.Index, bRaiseEventClick)
               SelectedByKey = xBtn.Index ' retornar el índice
               Exit For
            End If
        End If
    Next
End Function

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Evento clic de los botones
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub mUc_Click(Index As Integer, Optional RaiseEventClic As Boolean = True)
Attribute mUc_Click.VB_MemberFlags = "40"
    On Error GoTo error_handler
    Static flag As Boolean
    
    If Not flag Then
        Me.Redraw = False
        
        Dim xBtn As Control
        For Each xBtn In Controls
            If LCase(xBtn.Name) = LCase("uc") Then
                xBtn.ButtonType = eNormalButton
            End If
        Next
        
        uc(Index).ButtonType = eCheckbox
                
        flag = True
        uc(Index).Value = True
                                        
        Dim i As Long
        For i = 1 To mButtons.Count
            mButtons(i).FlagMod = True
            mButtons(i).Selected = False
            mButtons(i).FlagMod = False
        Next
        
        
        mButtons(Index).FlagMod = True
        mButtons(Index).Selected = True
        mButtons(Index).FlagMod = False
        
        
        If mSelectedItem Is Nothing Then Set mSelectedItem = New cButton
        
        ' Asignar valores del Botón seleccionado
        With mSelectedItem
            
            .FlagMod = True
            .Caption = uc(Index).Caption
            .ToolTipText = uc(Index).ToolTipText
            .Key = uc(Index).Tag
            .Index = Index
            .Selected = True
            .Enabled = uc(Index).Enabled
            .FlagMod = False
            
        End With
        
        If RaiseEventClic Then
           RaiseEvent ButtonClick(Index, mSelectedItem)
        End If
        
        flag = False
    End If
    
    DoEvents
    Me.Redraw = True
    
Exit Sub
error_handler:
    Me.Redraw = True
    Err.Raise Err.Number, "UcMenu", Err.Description
End Sub


'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Eventos
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================


Private Sub mParent_Resize()
    Static bResizeFlag As Boolean
    
    If Not bResizeFlag Then
       bResizeFlag = True
       Me.Refresh
    End If

End Sub

Private Sub picBtns_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picBtns_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picBtns_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub ucScroll_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ScrollMouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub ucScroll_MouseOut(Index As Integer)
    RaiseEvent ScrollMouseOut(Index)
End Sub

Private Sub ucScroll_MouseOver(Index As Integer)
    UserControl.MousePointer = 0
    RaiseEvent ScrollMouseOver(Index)
End Sub

Private Sub ucScroll_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
    RaiseEvent ScrollMouseUp(Index, Button, Shift, X, Y)
End Sub


Private Sub uc_GotFocus(Index As Integer)
    RaiseEvent ButtonGotFocus(Index)
End Sub

Private Sub uc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent ButtonKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub uc_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent ButtonKeyPress(Index, KeyAscii)
End Sub

Private Sub uc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent ButtonKeyUp(Index, KeyCode, Shift)
End Sub

Private Sub uc_LostFocus(Index As Integer)
    RaiseEvent ButtonLostFocus(Index)
End Sub

Private Sub uc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ButtonBeforeClick(Index)
    RaiseEvent ButtonMouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub uc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ButtonMouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub uc_MouseOut(Index As Integer)
    RaiseEvent ButtonMouseOut(Index)
End Sub

Private Sub uc_MouseOver(Index As Integer)
    UserControl.MousePointer = 0
    RaiseEvent ButtonMouseOver(Index)
End Sub

Private Sub uc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ButtonMouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub ucScroll_Click(Index As Integer)
    RaiseEvent ScrollClick(Index)
End Sub

Private Sub ucScroll_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent ScrollKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub ucScroll_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent ScrollKeyPress(Index, KeyAscii)
End Sub

Private Sub ucScroll_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent ScrollKeyUp(Index, KeyCode, Shift)
End Sub

Private Sub picBtns_Click()
    RaiseEvent Click
End Sub

Private Sub picBtns_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If mEnabledDragMenu And (Extender.Align = 0) Then
        Call ReleaseCapture
        Call DragMenu
    End If
End Sub

Private Sub picScroll_Click()
    RaiseEvent Click
End Sub

Private Sub picScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ScrollContainerMouseDown(Button, Shift, X, Y)
End Sub

Private Sub picScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ScrollContainerMouseMove(Button, Shift, X, Y)
    
    If mEnabledDragMenu And (Extender.Align = 0) Then
       Call ReleaseCapture
       Call DragMenu
    End If
End Sub


Private Sub picScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent ScrollContainerMouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
       Set mParent = Extender.Parent
       uc(0).Width = 0
       uc(0).Visible = False
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mButtons = Nothing
    Set mSelectedItem = Nothing
    Set mParent = Nothing
End Sub

Private Sub ucScroll_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent ScrollMouseDown(Index, Button, Shift, X, Y)
    If Index = 0 Then ucScroll(1).Enabled = True
    If Index = 1 Then ucScroll(0).Enabled = True
    
    mCurrentDirScroll = Index
    Timer1.Enabled = True
    
    RaiseEvent ScrollMouseDown(Index, Button, Shift, X, Y)
    
End Sub

Private Sub uc_Click(Index As Integer)
    Call mUc_Click(Index, True)
End Sub

Private Sub UserControl_Initialize()
    If mButtons Is Nothing Then
       Set Buttons = New cButtons
       Call Buttons.Init(Me)
    End If
End Sub

Sub UserControl_Resize()
Attribute UserControl_Resize.VB_MemberFlags = "40"
    On Error GoTo error_handler
    With UserControl
        If .Height < 480 Then .Height = 480
    End With
    
    If Ambient.UserMode And uc.Count = 1 Then
       picScroll.Visible = False
       picBtns.Move 0, 0, UserControl.Width, UserControl.Height
       Call mDrawSkinContainer
    ElseIf Ambient.UserMode And uc.Count > 1 Then
       Call mResizeControls
       Call mResizeControls
       Call mDrawSkinContainer
    ElseIf Ambient.UserMode = False And uc.Count = 1 Then
       picScroll.Visible = False
       picBtns.Move 0, 0, UserControl.Width, UserControl.Height
       uc(0).Move 60, 60, picBtns.TextWidth(uc(0).Caption) + (mMarginCaption * 2), picBtns.Height - 120
       Call mDrawSkinContainer
    End If
    
    RaiseEvent Resize
    
   Exit Sub
error_handler:
    Me.Redraw = True
    Debug.Print Err.Description
End Sub

Private Sub mResizeControls()

    On Error GoTo error_handler
    With picScroll
        .Left = (UserControl.Width - .Width)
        .Top = 0
        .Height = UserControl.Height
    End With
    
    Dim lWidth As Long
    lWidth = mGetButtonsWidth
    
    With picBtns
        If lWidth > UserControl.Width Then .Width = lWidth
        If .Width < UserControl.Width Then .Width = UserControl.Width
    End With

    Call mCheckScroll
    
    Static oldHeightUC As Long
    
    If oldHeightUC <> UserControl.Height Then
       oldHeightUC = UserControl.Height
       If Ambient.UserMode Then
       Dim xUc As Control
           For Each xUc In Controls
               If LCase(xUc.Name) = LCase("uc") Then
                   xUc.Height = picBtns.Height - 120
                   xUc.Top = 60
               End If
           Next
       End If
    End If
    
    Exit Sub
error_handler:
    Me.Redraw = True
    Debug.Print Err.Description
End Sub


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Iniciar valores de propiedades
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserControl_InitProperties()
    EnabledDragMenu = True
    UseUnderLineMouseUp = True
    UseUnderLineMouseCheck = True
    SmallChangeScroll = 120
    CaptionMargin = 120
    MarginButton = 60
    Skin = Office2007BlueButton
    Enabled = True
    Extender.Align = 1
End Sub



' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Guardar valores de propiedades
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
    
        Call .WriteProperty("EnabledDragMenu", mEnabledDragMenu, True)
        Call .WriteProperty("UseCustomForeColor", mUseCustomForeColor, False)
        Call .WriteProperty("ForeColorNormal", mForeColorNormal, &H8000000F)
        Call .WriteProperty("ForeColorDown", mForeColorDown, &H8000000F)
        Call .WriteProperty("ForeColorUp", mForeColorUp, &H8000000F)
        Call .WriteProperty("ForeColorDisabled", mForeColorDisabled, &H8000000F)
        Call .WriteProperty("ForeColorCheck", mForeColorCheck, &H8000000F)
        Call .WriteProperty("SmallChangeScroll", mSmallChangeScroll, 120)
        Call .WriteProperty("CaptionMargin", mMarginCaption, 120)
        Call .WriteProperty("UseUnderLineMouseUp", mUseUnderLineMouseUp)
        Call .WriteProperty("UseUnderLineMouseCheck", mUseUnderLineMouseCheck)
        Call .WriteProperty("ToolTipText", mToolTipText)
        Call .WriteProperty("MarginButton", mMarginButton, 60)
        Call .WriteProperty("ShowFocusRect", mShowFocusRect, False)
        Call .WriteProperty("Font", uc(0).Font)
        Call .WriteProperty("Skin", mSkin)
        Call .WriteProperty("Enabled", mEnabled, True)
    End With
End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Leer valores de propiedades
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mEnabledDragMenu = .ReadProperty("EnabledDragMenu", True)
        mMarginCaption = .ReadProperty("CaptionMargin", 120)
        ShowFocusRect = .ReadProperty("ShowFocusRect", False)
        mSmallChangeScroll = .ReadProperty("SmallChangeScroll", 120)
        UseUnderLineMouseUp = .ReadProperty("UseUnderLineMouseUp", False)
        UseUnderLineMouseCheck = .ReadProperty("UseUnderLineMouseCheck", False)
        ToolTipText = .ReadProperty("ToolTipText", "")
        MarginButton = .ReadProperty("MarginButton", 60)
        Set uc(0).Font = .ReadProperty("Font", uc(0).Font)
        Set picBtns.Font = .ReadProperty("Font", uc(0).Font)
        Skin = .ReadProperty("Skin", 0)
        uc(0).Width = picBtns.TextWidth(uc(0).Caption) + (mMarginCaption * 2)
        mForeColorNormal = .ReadProperty("ForeColorNormal", &H8000000F)
        mForeColorDown = .ReadProperty("ForeColorDown", &H8000000F)
        mForeColorUp = .ReadProperty("ForeColorUp", &H8000000F)
        mForeColorDisabled = .ReadProperty("ForeColorDisabled", &H8000000F)
        mForeColorCheck = .ReadProperty("ForeColorCheck", &H8000000F)
        mUseCustomForeColor = .ReadProperty("UseCustomForeColor", False)
        Enabled = .ReadProperty("Enabled", True)
        
    End With
End Sub


'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Propiedades

' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'===================================================================================================
'///////////////////////////////////////////////////////////////////////////////////////////////////
'===================================================================================================


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Mostrar o no el rectángulo cuando un botón tiene el foco
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ShowFocusRect() As Boolean
    ShowFocusRect = mShowFocusRect
End Property

Property Let ShowFocusRect(bValue As Boolean)
    If mShowFocusRect <> bValue Then
       mShowFocusRect = bValue
       Call PropertyChanged("ShowFocusRect")
       Call mSetPropertyButtons("ShowFocusRect", bValue)
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Habilitar / Deshabilitar menú
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    If mEnabled <> newValue Then
       mEnabled = newValue
       Call PropertyChanged("Enabled")
       Call mSetPropertyButtons("Enabled", newValue)
    End If
End Property


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TooltipText del menú
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Let ToolTipText(sValue As String)
    mToolTipText = sValue
    picBtns.ToolTipText = sValue
    Call PropertyChanged("ToolTipText")
End Property

Property Get ToolTipText() As String
    ToolTipText = mToolTipText
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Subrayar o no el Caption cuando el mouse está encima
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get UseUnderLineMouseUp() As Boolean
     UseUnderLineMouseUp = mUseUnderLineMouseUp
End Property

Property Let UseUnderLineMouseUp(bValue As Boolean)
    If mUseUnderLineMouseUp <> bValue Then
       mUseUnderLineMouseUp = bValue
       Call PropertyChanged("UseUnderLineMouseUp")
       Call mSetPropertyButtons("UseUnderLineMouseUp", bValue)
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Subrayar o no el Caption cuando se encuentra seleccionado el botón
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get UseUnderLineMouseCheck() As Boolean
     UseUnderLineMouseCheck = mUseUnderLineMouseCheck
End Property

Property Let UseUnderLineMouseCheck(bValue As Boolean)
    If mUseUnderLineMouseCheck <> bValue Then
       mUseUnderLineMouseCheck = bValue
       Call PropertyChanged("UseUnderLineMouseCheck")
       Call mSetPropertyButtons("UseUnderLineMouseCheck", bValue)
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fuente
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Font() As Font
    Set Font = uc(0).Font
End Property

Public Property Set Font(ByRef newFont As Font)
    Set picBtns.Font = newFont
    Set uc(0).Font = newFont
    Call PropertyChanged("Font")
    Call mSetPropertyButtons("Font", newFont)
    
    If Not Ambient.UserMode Then
       uc(0).Width = picBtns.TextWidth(uc(0).Caption) + (mMarginCaption * 2)
    Else
       ' modificar la nueva fuente de los botones
       Call mModifyButtons
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Negrita
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = uc(0).FontBold
End Property

Public Property Let FontBold(ByVal newValue As Boolean)
    If uc(0).FontBold <> newValue Then
       uc(0).FontBold = newValue
       picBtns.FontBold = newValue
       If Ambient.UserMode Then mModifyButtons
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Italic
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = uc(0).FontItalic
End Property

Public Property Let FontItalic(ByVal newValue As Boolean)
    If uc(0).FontItalic <> newValue Then
       uc(0).FontItalic = newValue
       picBtns.FontItalic = newValue
       If Ambient.UserMode Then mModifyButtons
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Subrayado
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = uc(0).FontUnderline
End Property

Public Property Let FontUnderline(ByVal newValue As Boolean)
    If uc(0).FontUnderline <> newValue Then
       uc(0).FontUnderline = newValue
       picBtns.FontUnderline = newValue
       If Ambient.UserMode Then mModifyButtons
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tamaño de fuente
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FontSize() As Integer
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = uc(0).FontSize
End Property

Public Property Let FontSize(ByVal newValue As Integer)
    If uc(0).FontSize <> newValue Then
       uc(0).FontSize = newValue
       picBtns.FontSize = newValue
       If Ambient.UserMode Then mModifyButtons
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Nombre de fuente
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = uc(0).FontName
End Property

Public Property Let FontName(ByVal newValue As String)
    
    On Error Resume Next
    
    If uc(0).FontName <> newValue Then
       uc(0).FontName = newValue
       picBtns.FontName = newValue
       If Err.Number <> 0 Then
          uc(0).FontName = Ambient.Font.Name
          picBtns.FontName = Ambient.Font.Name
       End If

       On Error GoTo 0
       If Ambient.UserMode Then mModifyButtons
    End If
    
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Margen del texto del botón
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get CaptionMargin() As Long
    CaptionMargin = mMarginCaption
End Property

Property Let CaptionMargin(lValue As Long)
    If mMarginCaption <> lValue Then
       If lValue < 60 Then lValue = 60
       mMarginCaption = lValue
       Call PropertyChanged("CaptionMargin")
       Call mSetPropertyButtons("CaptionMargin", lValue)
       If Ambient.UserMode Then Call mModifyButtons
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Valor para el desplazamiento del Scroll
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get SmallChangeScroll() As Long
    SmallChangeScroll = mSmallChangeScroll
End Property

Property Let SmallChangeScroll(lValue As Long)
    If lValue <= 60 Then lValue = 60
    mSmallChangeScroll = lValue
    Call PropertyChanged("SmallChangeScroll")
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Valor de separación de los botones
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get MarginButton() As Long
    MarginButton = mMarginButton
End Property
Property Let MarginButton(lValue As Long)
    
    If mMarginButton <> lValue Then
       If lValue < 60 Then
          mMarginButton = 60
       Else
          mMarginButton = lValue
       End If
       If Ambient.UserMode Then Call mModifyButtons
    End If
End Property

Property Get Buttons() As cButtons
    Set Buttons = mButtons
End Property

Property Set Buttons(Value As cButtons)
    Set mButtons = Value
End Property


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Estilo del botón
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get Skin() As eSkin
    Skin = mSkin
End Property

Property Let Skin(skinValue As eSkin)

    If (skinValue = Link) Or (skinValue = CustomSkin) Then
       If Ambient.UserMode = False Then
          MsgBox "Esta opción no funciona para el menú.", vbInformation
       End If
    
       Exit Property
    End If
    
    If mSkin <> skinValue Then
       mSkin = skinValue
       Call PropertyChanged("Skin")
       Call mSetPropertyButtons("Skin", skinValue)
       Call mDrawSkinContainer
    End If
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Colores de fuentes para usar cuando la propiedad CustomForecolor está en True
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get ForeColorNormal() As OLE_COLOR
    ForeColorNormal = mForeColorNormal
End Property

Property Let ForeColorNormal(lValue As OLE_COLOR)
    If mForeColorNormal <> lValue Then
       mForeColorNormal = lValue
       Call PropertyChanged("ForeColorNormal")
       If mUseCustomForeColor Then
          Call mSetPropertyButtons("ForeColorNormal", lValue)
       End If
    End If
End Property

Property Get ForeColorUp() As OLE_COLOR
    ForeColorUp = mForeColorUp
End Property
Property Let ForeColorUp(lValue As OLE_COLOR)
    If mForeColorUp <> lValue Then
       mForeColorUp = lValue
       Call PropertyChanged("ForeColorUp")
       Call mSetPropertyButtons("ForeColorUp", lValue)
    End If
End Property

Property Get ForeColorDown() As OLE_COLOR
    ForeColorDown = mForeColorDown
End Property

Property Let ForeColorDown(lValue As OLE_COLOR)
    If mForeColorDown <> lValue Then
       mForeColorDown = lValue
       Call PropertyChanged("ForeColorDown")
       Call mSetPropertyButtons("ForeColorDown", lValue)
    End If
End Property

Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = mForeColorDisabled
End Property

Property Let ForeColorDisabled(lValue As OLE_COLOR)
    If mForeColorDisabled <> lValue Then
       mForeColorDisabled = lValue
       Call PropertyChanged("ForeColorDisabled")
       Call mSetPropertyButtons("ForeColorDisabled", lValue)
    End If
End Property

Property Get ForeColorCheck() As OLE_COLOR
    ForeColorCheck = mForeColorCheck
End Property

Property Let ForeColorCheck(lValue As OLE_COLOR)
    If mForeColorCheck <> lValue Then
       mForeColorCheck = lValue
       Call PropertyChanged("ForeColorCheck")
       Call mSetPropertyButtons("ForeColorCheck", lValue)
    End If
End Property

Property Get UseCustomForeColor() As Boolean
    UseCustomForeColor = mUseCustomForeColor
End Property

Property Let UseCustomForeColor(bValue As Boolean)
    
    If mUseCustomForeColor = bValue Then Exit Property
    
    mUseCustomForeColor = bValue
    Call PropertyChanged("UseCustomForeColor")
    
    Dim xUc As Control
        
    For Each xUc In Controls
        If TypeOf xUc Is ucBtnSkin Then
            With xUc
                If bValue Then
                   .bFlagNoUpdateBtn = True
                   .ColorSchemas = 0
                   .ForeColorNormal = mForeColorNormal
                   .ForeColorCheck = mForeColorCheck
                   .ForeColorDisabled = mForeColorDisabled
                   .ForeColorUp = mForeColorUp
                   .ForeColorDown = mForeColorDown
                   .bFlagNoUpdateBtn = False
                   .Refresh
                Else
                   .ColorSchemas = 1
                End If
            End With
       End If
    Next
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Habilitar el Drag Drop
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Property Get EnabledDragMenu() As Boolean
    EnabledDragMenu = mEnabledDragMenu
End Property

Property Let EnabledDragMenu(bValue As Boolean)
    mEnabledDragMenu = bValue
    PropertyChanged "EnabledDragMenu"
End Property


Property Get BorderColorSkinDefault() As Long
    BorderColorSkinDefault = uc(0).BorderColorSkinDefault
End Property

Property Get BackColorSkinDefault() As Long
Attribute BackColorSkinDefault.VB_MemberFlags = "400"
    BackColorSkinDefault = uc(0).BackColorSkinDefault
End Property
Property Get ForeColorDefault() As Long
    ForeColorDefault = uc(0).ForeColorNormal
End Property

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Botón actual
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get SelectedItem() As cButton
    If mSelectedItem Is Nothing Then Set mSelectedItem = New cButton
    Set SelectedItem = mSelectedItem
    
    If Not mSelectedItem Is Nothing Then
       If mSelectedItem.Index = 0 Then
          Set SelectedItem = Nothing
       End If
    End If
End Property

Property Set SelectedItem(objValue As cButton)
    Set mSelectedItem = objValue
End Property

Property Let SkinCustomPicturePath(stdPicPath As String)
    On Error GoTo error_handler
    Dim xBtn As Control
    Redraw = False
    For Each xBtn In Controls
        If TypeOf xBtn Is ucBtnSkin Then
            Set xBtn.SkinCustomPicture = LoadPicture(stdPicPath)
        End If
    Next
    
    Call mDrawSkinContainer
    Redraw = True
    
    Exit Property
error_handler:
Redraw = True
If Err.Number = 53 Then
   Err.Raise 53, "UcMenu", "No se ha encontrado el archivo " & stdPicPath
Else
   Err.Raise Err.Number, "UcMenu", Err.de
End If
End Property


' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Bloquear el repintado del UC
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Redraw() As Boolean
Attribute Redraw.VB_MemberFlags = "400"
    Redraw = mRedraw
End Property

Property Let Redraw(bValue As Boolean)
    DoEvents
    mRedraw = bValue
    If Not bValue Then
       LockWindowUpdate UserControl.hwnd
    Else
       LockWindowUpdate ByVal 0
    End If
    DoEvents
End Property

