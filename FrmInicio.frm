VERSION 5.00
Begin VB.Form FrmInicio 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Proyecto1.ucMenu ucMenu1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
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
End
Attribute VB_Name = "FrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call loadButtons
    Call mnuskins_Click(6)
    With ucMenu1
    .Height = 900
    End With
End Sub
Private Sub ucMenu1_ButtonClick( _
    ByVal ButtonIndex As Integer, _
    Button As cButton)
    
    With Button

        If .Caption = "VENTAS" Then
            FRMVENTAS.Show
        End If
    End With
End Sub
Private Sub loadButtons()

    On Error GoTo error_handler
    With ucMenu1
         
         .Redraw = False
         .Buttons.Add "VENTAS "
         .Buttons.Add "REPORTES"
         .Buttons.Add "EMPRESA"
         
         .Redraw = True
    End With

    Exit Sub
error_handler:
    MsgBox Err.Description, vbCritical

End Sub
Private Sub VScroll1_Change()
    Frame2.Top = -CSng(VScroll1.Value) * 120
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
