VERSION 5.00
Begin VB.Form Calcular1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "TABLA"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   Icon            =   "Calcular1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6780
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      Picture         =   "Calcular1.frx":628A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton BTN9 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   4920
      Picture         =   "Calcular1.frx":B374
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "SALIR"
      Top             =   720
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command2 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton Command5 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   1560
         TabIndex        =   13
         Top             =   1200
         Width           =   1300
      End
      Begin VB.CommandButton Command8 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   1560
         TabIndex        =   12
         Top             =   2160
         Width           =   1300
      End
      Begin VB.CommandButton Command10 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   1560
         TabIndex        =   11
         Top             =   3120
         Width           =   1300
      End
      Begin VB.CommandButton Command12 
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   3000
         TabIndex        =   8
         Top             =   3120
         Width           =   1300
      End
      Begin VB.CommandButton Command11 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   1300
      End
      Begin VB.CommandButton Command9 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   3000
         TabIndex        =   6
         Top             =   2160
         Width           =   1300
      End
      Begin VB.CommandButton Command7 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   1300
      End
      Begin VB.CommandButton Command6 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   3000
         TabIndex        =   4
         Top             =   1200
         Width           =   1300
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1300
      End
      Begin VB.CommandButton Command3 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1300
      End
   End
End
Attribute VB_Name = "Calcular1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN9_Click()
Unload Me
End Sub

Private Sub Command1_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 7
PUB_VALOR_C = Ventas.Text1
Else
Ventas.Text1 = Ventas.Text1 + "7"
PUB_VALOR_C = Ventas.Text1
End If
End Sub

Private Sub Command10_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = "00"
Else
Ventas.Text1 = Ventas.Text1 + "00"
End If
End Sub

Private Sub Command11_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 0
Else
Ventas.Text1 = Ventas.Text1 + "0"
End If
End Sub

Private Sub Command12_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = "000"
Else
Ventas.Text1 = Ventas.Text1 + "000"
End If
End Sub

Private Sub Command2_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 8
PUB_VALOR_C = Ventas.Text1
Else
Ventas.Text1 = Ventas.Text1 + "8"
PUB_VALOR_C = Ventas.Text1
End If
End Sub

Private Sub Command23_Click()
Ventas.Text1.Text = ""
Ventas.Text4.Text = ""
Ventas.Text2.Text = ""
End Sub

Private Sub Command3_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 9
PUB_VALOR_C = Ventas.Text1
Else
Ventas.Text1 = Ventas.Text1 + "9"
PUB_VALOR_C = Ventas.Text1
End If
End Sub

Private Sub Command4_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 4
PUB_VALOR_C = Ventas.Text1
Else
Ventas.Text1 = Ventas.Text1 + "4"
PUB_VALOR_C = Ventas.Text1
End If
End Sub

Private Sub Command5_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 5
PUB_VALOR_C = Ventas.Text1
Else
Ventas.Text1.Text = Ventas.Text1.Text + "5"
PUB_VALOR_C = Ventas.Text1
End If
End Sub

Private Sub Command6_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 6
PUB_VALOR_C = Ventas.Text1
Else
Ventas.Text1 = Ventas.Text1 + "6"
PUB_VALOR_C = Ventas.Text1
End If
End Sub

Private Sub Command7_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 1
PUB_VALOR_C = Ventas.Text1
Else
Ventas.Text1 = Ventas.Text1 + "1"
PUB_VALOR_C = Ventas.Text1
End If
End Sub

Private Sub Command8_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 2
PUB_VALOR_C = Ventas.Text1
Else
Ventas.Text1 = Ventas.Text1 + "2"
PUB_VALOR_C = Ventas.Text1
End If
End Sub

Private Sub Command9_Click()
If Ventas.Text1 = "" Then
Ventas.Text1 = 3
PUB_VALOR_C = Ventas.Text1
Else
Ventas.Text1 = Ventas.Text1 + "3"
PUB_VALOR_C = Ventas.Text1
End If
End Sub

