VERSION 5.00
Begin VB.Form Cliente 
   Caption         =   "Cliente"
   ClientHeight    =   1560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BTN5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   955
      Left            =   9240
      Picture         =   "Cliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "NUEVO"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   6735
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
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN5_Click()
DIM_CLIENTE = Text1.Text
DIM_RTNCIENTE = Text2.Text
Ventas.addcliente
Unload Me
End Sub
