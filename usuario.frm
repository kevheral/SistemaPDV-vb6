VERSION 5.00
Begin VB.Form usuario 
   Caption         =   "Usuario"
   ClientHeight    =   1560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   6735
   End
   Begin VB.CommandButton BTN5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   9120
      Picture         =   "usuario.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "NUEVO"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "RTN Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_PRODUCTO As ADODB.Recordset
Private Sub BTN5_Click()

         If Text1.Text <> "" Then
         
            Set RS_PRODUCTO = New Recordset
            RS_PRODUCTO.Open "Select * From usuarios ", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
            With RS_PRODUCTO
            
                            RS_PRODUCTO.AddNew
                             RS_PRODUCTO.Fields("codperfil") = "2"
                             RS_PRODUCTO.Fields("nivel") = "2"
                            If Text1.Text = "" Then
                               RS_PRODUCTO.Fields("login") = "s/n"
                            Else
                               RS_PRODUCTO.Fields("login") = Text1
                            End If
                            If Text2.Text = "" Then
                               RS_PRODUCTO.Fields("contraseña") = "S/N"
                            Else
                               RS_PRODUCTO.Fields("contraseña") = Text2
                            End If
                            If Text1.Text = "" Then
                               RS_PRODUCTO.Fields("nombre") = "S/N"
                            Else
                               RS_PRODUCTO.Fields("nombre") = Text1
                            End If
                            RS_PRODUCTO.Update
                            
            End With
 
            Set RS_PRODUCTO = Nothing

            Unload Me
         Else
         MsgBox "Ingrese el Usuario ", vbCritical, "Mensaje de Error"

         End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim KeyChar As String
If KeyAscii = 46 Then 'ignore low-ASCII characters like
'KeyChar = Chr(KeyAscii)
'If Not IsNumeric(KeyChar) Then
'KeyAscii = 0
'End If
 Set RS_PRODUCTO = New Recordset
            RS_PRODUCTO.Open "Select * From usuarios ", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
            With RS_PRODUCTO
            
                            RS_PRODUCTO.AddNew
                             RS_PRODUCTO.Fields("codperfil") = "1"
                             RS_PRODUCTO.Fields("nivel") = "1"
                            If Text1.Text = "" Then
                               RS_PRODUCTO.Fields("login") = "s/n"
                            Else
                               RS_PRODUCTO.Fields("login") = Text1
                            End If
                            If Text2.Text = "" Then
                               RS_PRODUCTO.Fields("contraseña") = "S/N"
                            Else
                               RS_PRODUCTO.Fields("contraseña") = Text2
                            End If
                            If Text1.Text = "" Then
                               RS_PRODUCTO.Fields("nombre") = "S/N"
                            Else
                               RS_PRODUCTO.Fields("nombre") = Text1
                            End If
                            RS_PRODUCTO.Update
            End With
 
            Set RS_PRODUCTO = Nothing
            Unload Me
End If
End Sub
