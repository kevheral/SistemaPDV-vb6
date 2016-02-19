VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Ingreso 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5130
   ClientLeft      =   -15
   ClientTop       =   315
   ClientWidth     =   9735
   Icon            =   "Ingreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "#"
      TabIndex        =   10
      Top             =   240
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4575
      Begin VB.CommandButton Command10 
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
         Left            =   1560
         TabIndex        =   11
         Top             =   3120
         Width           =   1300
      End
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   2160
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   5040
      TabIndex        =   12
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "Ingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cuenta As ADODB.Recordset
Dim DIM_SELECT
Private Sub Command1_Click()
Text1.Text = Text1 + "7"
End Sub

Private Sub Command10_Click()
Text1.Text = Text1 + "0"
End Sub

Private Sub Command2_Click()
Text1.Text = Text1 + "8"
End Sub

Private Sub Command3_Click()
Text1.Text = Text1 + "9"
End Sub

Private Sub Command4_Click()
Text1.Text = Text1 + "4"
End Sub

Private Sub Command5_Click()
Text1.Text = Text1 + "5"
End Sub

Private Sub Command6_Click()
Text1.Text = Text1 + "6"
End Sub

Private Sub Command7_Click()
Text1.Text = Text1 + "1"
End Sub

Private Sub Command8_Click()
Text1.Text = Text1 + "2"
End Sub

Private Sub Command9_Click()
Text1.Text = Text1 + "3"
End Sub

Private Sub Form_Load()
gsRutaBaseDatos = App.Path & "\DB.mdb"
AbrirDB

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''
    With ListView1
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NOMBRE", 5000
    End With
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''
Set Cuenta = New Recordset
Cuenta.Open "Select login From usuarios ", PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
With Cuenta
       If Cuenta.BOF = True And Cuenta.EOF = True Then
        ListView1.ListItems.Clear
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = ListView1.ListItems.Add(, , .Fields(0) & "")
                .MoveNext
            Loop
        End If
End With
 
Set Cuenta = Nothing

End Sub

Private Sub ListView1_DblClick()
usuario.Show
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
DIM_SELECT = Item.Text
End Sub

Private Sub Text1_Change()
Dim DIM_SQL
    With Text1

        If InStr(1, .Text, "'") <> 0 Or InStr(1, .Text, "[") <> 0 Or _
            InStr(1, .Text, "|") <> 0 Or InStr(1, .Text, """") <> 0 Or _
            InStr(1, .Text, "*") <> 0 Or InStr(1, .Text, "/") <> 0 Then
           .Text = ""
            Exit Sub
        Else
            Dim DIM_NOMBRE
            DIM_NOMBRE = .Text
            Set Cuenta = New Recordset
            DIM_SQL = Text1.Text
            DIM_SQLITEM = "SELECT * FROM usuarios where login like '" & DIM_SELECT & "'"
            DIM_SQLITEM = DIM_SQLITEM & " AND contraseña like  '" & DIM_SQL & "'"
        
        
            Cuenta.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
            With Cuenta
                If .EOF Then
                 '   MsgBox "No se localizo la Cuenta [" & a & "]", vbCritical, "Error de busqueda"
                Else
                Inicio.Show
                Unload Me
                End If
            End With
        End If
    End With



End Sub
