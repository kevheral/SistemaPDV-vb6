VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Inicio 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9120
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   12435
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   12435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Proyecto1.ucMenu ucMenu1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   873
      ForeColorNormal =   0
      ForeColorDown   =   0
      ForeColorUp     =   0
      ForeColorDisabled=   0
      ForeColorCheck  =   0
      UseUnderLineMouseUp=   -1  'True
      UseUnderLineMouseCheck=   -1  'True
      Object.ToolTipText     =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Skin            =   16
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   11640
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Archivo"
      Begin VB.Menu mnuExit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuskin 
      Caption         =   "Skin"
      Begin VB.Menu mnuskins 
         Caption         =   "[VBNet]"
         Index           =   0
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[XPYellow] "
         Index           =   1
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[XPBlue] "
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[XP] "
         Index           =   3
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[XPPink] "
         Index           =   4
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[HotMail] "
         Index           =   5
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WindowsMedia] "
         Index           =   6
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WMPCobre] "
         Index           =   7
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WMPVerde] "
         Index           =   8
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Word2000] "
         Index           =   9
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[RoyaleXP] "
         Index           =   10
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WindowsLive] "
         Index           =   11
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WindowsLive2] "
         Index           =   12
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Google Picasa]"
         Index           =   13
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Office2007 Blue TabStrip]"
         Index           =   14
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Office2007 Blue Button]"
         Index           =   15
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Office2007 check green]"
         Index           =   16
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Hotmail2]"
         Index           =   17
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Windows vista]"
         Index           =   18
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[CBblue]"
         Index           =   19
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Royale xp TaskBar]"
         Index           =   20
      End
      Begin VB.Menu mnuskins 
         Caption         =   "WinVista2"
         Index           =   21
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Opciones"
      Begin VB.Menu mnualign 
         Caption         =   "Alinear Menu"
         Begin VB.Menu mnualignMent 
            Caption         =   "No alinear"
            Index           =   0
         End
         Begin VB.Menu mnualignMent 
            Caption         =   "Arriba"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnualignMent 
            Caption         =   "Abajo"
            Index           =   2
         End
      End
      Begin VB.Menu mnuText 
         Caption         =   "Texto"
         Begin VB.Menu mnuCambiarFuente 
            Caption         =   "Cambiar fuente"
         End
         Begin VB.Menu mnuUnderLine 
            Caption         =   "Subrayar enlace"
            Begin VB.Menu mnuUnderLines 
               Caption         =   "Subrayar en MouseUP"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuUnderLines 
               Caption         =   "Subrayar en Check"
               Checked         =   -1  'True
               Index           =   1
            End
         End
         Begin VB.Menu mnuForeColor 
            Caption         =   "Color de fuente"
            Begin VB.Menu mnuForeColors 
               Caption         =   "ForeColorNormal"
               Index           =   0
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "ForeColorUp"
               Index           =   1
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "ForecolorCheck"
               Index           =   2
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "ForeColorDisabled"
               Index           =   3
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "-"
               Index           =   4
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "No usar Forecolor"
               Checked         =   -1  'True
               Index           =   5
            End
         End
      End
      Begin VB.Menu mnuScroll2 
         Caption         =   "Scroll"
         Begin VB.Menu mnuScrolls 
            Caption         =   "Aumentar"
            Index           =   0
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuScrolls 
            Caption         =   "Disminuir"
            Index           =   1
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuScrolls 
            Caption         =   "Restaurar"
            Index           =   2
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu mnuMar 
         Caption         =   "Margen"
         Begin VB.Menu mnuMargin 
            Caption         =   "Margen del texto"
            Index           =   0
         End
         Begin VB.Menu mnuMargin 
            Caption         =   "Margen del botón"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEna 
         Caption         =   "Enabled"
         Begin VB.Menu mnuEnabled 
            Caption         =   "Habilitar control"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuEnabled 
            Caption         =   "Habilitar/Deshabilitar botón"
            Index           =   1
         End
      End
      Begin VB.Menu mnuDrag 
         Caption         =   "Drag"
         Begin VB.Menu mnuDragDrop 
            Caption         =   "Habilitar Drag Drop ( Debe estar el Align en 0)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuTam 
         Caption         =   "Tamaño del Menú"
         Begin VB.Menu mnuSizeMenu 
            Caption         =   "Chico"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuSizeMenu 
            Caption         =   "Medio"
            Index           =   1
         End
         Begin VB.Menu mnuSizeMenu 
            Caption         =   "Mas grande"
            Index           =   2
         End
      End
      Begin VB.Menu mnuLoadSkin 
         Caption         =   "Cargar Skin desde archivo"
         Begin VB.Menu mnuLoadSkins 
            Caption         =   "Skin rosa"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuScroll 
      Caption         =   "Scroll"
      Visible         =   0   'False
      Begin VB.Menu mnuMoveScroll 
         Caption         =   "Ir al principio"
         Index           =   0
      End
      Begin VB.Menu mnuMoveScroll 
         Caption         =   "Ir al final"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub loadButtons()


    On Error GoTo error_handler
    With ucMenu1
         
         .Redraw = False
         .Buttons.Add "VENTAS"
         .Buttons.Add "REPORTES"
         .Buttons.Add "EMPRESA"


         .Redraw = True
    End With

    Exit Sub
error_handler:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Form_Load()
    Call loadButtons
    Call mnuskins_Click(6)
    With ucMenu1
    .Height = 900
    End With
End Sub

Private Sub mnualignMent_Click(Index As Integer)
    Call CheckMenuChange(mnualignMent, Index)
    With ucMenu1
        .Align = Index
        If Index = 0 Then .Move 10, 50, Me.ScaleWidth / 2
    End With
End Sub

Private Sub mnuCambiarFuente_Click()
    
    Dim objFont As New StdFont
    With cd
        .CancelError = True
        .Flags = cdlCFScreenFonts
        .FontBold = ucMenu1.FontBold
        .FontItalic = ucMenu1.FontItalic
        .FontSize = ucMenu1.FontSize
        .FontName = ucMenu1.FontName
        
        On Error Resume Next
        .ShowFont
        
        If Not Err Then
           objFont.Bold = .FontBold
           objFont.Size = .FontSize
           objFont.Italic = .FontItalic
           objFont.Name = .FontName
           Set ucMenu1.Font = objFont
        On Error GoTo 0
        End If
    End With
End Sub

Private Sub mnuDragDrop_Click()
    With mnuDragDrop
        .Checked = Not .Checked
        ucMenu1.EnabledDragMenu = .Checked
    End With
End Sub

Private Sub mnuEnabled_Click(Index As Integer)
On Error GoTo error_handler
    Select Case Index
        Case 0
            mnuEnabled(0).Checked = Not mnuEnabled(0).Checked
            ucMenu1.Enabled = mnuEnabled(0).Checked
        Case 1
            Dim ret As Variant
            ret = InputBox("Indice del botón a deshabilitar")
            If IsNumeric(ret) Then
               ucMenu1.Buttons(CInt(ret)).Enabled = Not ucMenu1.Buttons(CInt(ret)).Enabled
            End If
    End Select
    
    Exit Sub
error_handler:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuForeColors_Click(Index As Integer)
    
    Dim lcolor As Long
    If Index <= 3 Then
       With cd
          .CancelError = True
          On Error Resume Next
          .ShowColor
          If Err.Number = 32755 Then Exit Sub
          On Error GoTo 0
          lcolor = .Color
       End With
    End If
    
    With ucMenu1
        Select Case Index
            Case 0
                .ForeColorNormal = lcolor
            Case 1
                .ForeColorUp = lcolor
            Case 2
                .ForeColorCheck = lcolor
            Case 3
                .ForeColorDisabled = lcolor
            Case 5
                .UseCustomForeColor = mnuForeColors(5).Checked
                 mnuForeColors(5).Checked = Not mnuForeColors(5).Checked
        End Select
    End With
    
End Sub

Private Sub mnuLoadSkins_Click(Index As Integer)
            
    On Error GoTo error_handler
    
    Dim spathSkin As String
    If Index = 0 Then spathSkin = App.Path & "\skin\skin1.bmp"
    With ucMenu1
        .UseCustomForeColor = True
         mnuForeColors(5).Checked = False
         
         Select Case Index
            Case 0
                .ForeColorUp = &H800000
                .ForeColorNormal = vbBlack
                .ForeColorCheck = vbBlack
                .ForeColorDown = vbBlack
                .ForeColorDisabled = RGB(190, 190, 190)
         End Select
        
        .SkinCustomPicturePath = spathSkin
    End With
    
    With ucMenu1
        Me.BackColor = .GetSkinsColors(Normal_Skin)
        Dim xCtrl As Control
        For Each xCtrl In Me.Controls
            
            On Error Resume Next
            If Not (TypeOf xCtrl Is ucMenu) Then
                xCtrl.BackColor = .GetSkinsColors(Normal_Skin)
                xCtrl.BorderColor = .GetSkinsColors(Border_Normal_Skin)
                xCtrl.ForeColor = .ForeColorNormal
            End If
            If TypeOf xCtrl Is ucBtnSkin Then
               xCtrl.UseCustomForeColor = True
               xCtrl.SkinCustomPicture = LoadPicture(spathSkin)
               xCtrl.ForeColorUp = .ForeColorUp
               xCtrl.ForeColorNormal = .ForeColorNormal
               xCtrl.ForeColorCheck = .ForeColorCheck
               xCtrl.ForeColorDown = .ForeColorDown
               xCtrl.ForeColorDisabled = .ForeColorDisabled
            End If
        Next
    End With
    
    Exit Sub
error_handler:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub mnuMargin_Click(Index As Integer)
    
    Dim ret As Variant
    With ucMenu1
       Select Case Index
           Case 0
             ret = InputBox("Ingresar valor" & vbNewLine & "Valor actual: " & CStr(.CaptionMargin), "Margen del texto")
             If IsNumeric(ret) Then ucMenu1.CaptionMargin = ret
           Case 1
             ret = InputBox("Ingresar valor" & vbNewLine & "Valor actual: " & CStr(.MarginButton), "Margen de separación del botón")
             If IsNumeric(ret) Then ucMenu1.MarginButton = ret
        End Select
    End With
End Sub

Private Sub mnuMoveScroll_Click(Index As Integer)
    With ucMenu1
        If Index = 0 Then Call .MoveScroll(First)
        If Index = 1 Then Call .MoveScroll(Last)
    End With
End Sub
Private Sub mnuScrolls_Click(Index As Integer)
    With ucMenu1
        Select Case Index
            Case 0: .SmallChangeScroll = .SmallChangeScroll + 25
            Case 1: .SmallChangeScroll = .SmallChangeScroll - 25
            Case 2: .SmallChangeScroll = 50
        End Select
    End With
End Sub
Private Sub mnuSizeMenu_Click(Index As Integer)
    
    Call CheckMenuChange(mnuSizeMenu, Index)
    
    With ucMenu1
        Select Case Index
            Case 0: .Height = 400
            Case 1: .Height = 615
            Case 2: .Height = 800
        End Select
    End With
End Sub

Private Sub mnuskins_Click(Index As Integer)
    
    Call CheckMenuChange(mnuskins, Index)
    
    With ucMenu1
         
         mnuForeColors(5).Checked = True
        .UseCustomForeColor = False
        .Skin = (CLng(Index))
        Me.BackColor = .BackColorSkinDefault
            
        On Error Resume Next
        Dim xCtrl As Control
        For Each xCtrl In Me.Controls
            
            If Not (TypeOf xCtrl Is ucMenu) Then
                xCtrl.Skin = CLng(Index)
                xCtrl.BackColor = .BackColorSkinDefault
                xCtrl.BorderColor = .BorderColorSkinDefault
                xCtrl.ForeColor = .ForeColorDefault
                xCtrl.FontName = .FontName
            End If
        Next
        On Error GoTo 0
        
    End With
End Sub
Private Sub CheckMenuChange(pMenu As Object, lIndex As Integer)
    Dim xMenu As Variant
    For Each xMenu In pMenu
        xMenu.Checked = False
    Next
    pMenu(lIndex).Checked = True
End Sub
Private Sub mnuUnderLines_Click(Index As Integer)
    mnuUnderLines(Index).Checked = Not mnuUnderLines(Index).Checked
    With ucMenu1
        .UseUnderLineMouseUp = mnuUnderLines(0).Checked
        .UseUnderLineMouseCheck = mnuUnderLines(1).Checked
    End With
End Sub

Private Sub txtFindBtn_GotFocus(Index As Integer)
    txtFindBtn(0).Text = ""
    txtFindBtn(1).Text = ""
End Sub

Private Sub txtFindBtn_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
       If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub ucBtn_Click(Index As Integer)

On Error GoTo error_handler

    Dim xBtn As cButton
    
    Select Case Index
        Case 0
            If txtTextButton = "" Then Exit Sub
            With ucMenu1
                Set xBtn = .Buttons.Add(txtTextButton.Text, , txtTooltip.Text)
                'xBtn.Selected = True
            End With
        Case 1
            With ucMenu1
                If Not .SelectedItem Is Nothing Then
                   .Buttons.Remove .SelectedItem.Index
                Else
                    MsgBox "No hay Item seleccionado", vbExclamation
                End If
            End With
        
        Case 2
            ucMenu1.Buttons.Clear
        
        Case 3
            With ucMenu1
                 
                 If Not .SelectedItem Is Nothing Then
                    If txtTextButton.Text <> "" Then
                       .Buttons(.SelectedItem.Index).Caption = txtTextButton.Text
                    End If
                    If txtTooltip.Text <> "" Then
                       .Buttons(.SelectedItem.Index).ToolTipText = txtTooltip.Text
                    End If
                 Else
                    MsgBox "No hay Item seleccionado", vbExclamation
                 End If
            End With
        Case 4
            Call loadButtons
    End Select
    
     Exit Sub
error_handler:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub ucBtnFind_Click()

On Error GoTo error_handler

    Dim bFind As Boolean
    With ucMenu1
        If Len(txtFindBtn(0)) Then
           bFind = .FindButton(txtFindBtn(0).Text, byCaption, True)
        ElseIf Len(txtFindBtn(1)) Then
           bFind = .FindButton(txtFindBtn(1).Text, byIndex, True)
        End If
    End With
    
    If Not bFind Then MsgBox "No se encontró el botón", vbInformation
    

    Exit Sub
error_handler:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub ucMenu1_ButtonClick( _
    ByVal ButtonIndex As Integer, _
    Button As cButton)
    
    With Button

        If .Caption = "VENTAS" Then
            Ventas.Show
        End If
        If .Caption = "REPORTES" Then
            Reportes.Show
        End If
        If .Caption = "EMPRESA" Then
            Ventas.Show
        End If
    End With
End Sub

Private Sub ucMenu1_ScrollContainerMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       Me.PopupMenu mnuScroll
    End If
End Sub
Private Sub ucMenu1_ButtonMouseOver(ByVal ButtonIndex As Integer)
    Me.Caption = ucMenu1.Buttons(ButtonIndex).Caption
End Sub
