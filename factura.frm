VERSION 5.00
Begin VB.Form factura 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   6960
      Picture         =   "factura.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "SALIR"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   1200
      Picture         =   "factura.frx":4417
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "ELIMINAR"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton BTN5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   7440
      Picture         =   "factura.frx":8566
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "NUEVO"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   9240
      Top             =   720
   End
   Begin VB.TextBox TXT_CANTIDAD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   3855
   End
   Begin VB.Label lblST 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   11
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Label lblCB 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   870
      Left            =   4200
      TabIndex        =   10
      Top             =   4920
      Width           =   5385
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CAMBIO CLIENTE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   8
      Top             =   3360
      Width           =   5415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "DESCUENTOS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TOTAL VENTAS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TOTAL 15% ISV "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TOTAL SIN  ISV :"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   2640
      Width           =   5415
   End
End
Attribute VB_Name = "factura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PRVT_1 As New ADODB.Connection
Private PRVT_2 As New ADODB.Command
Private PRVT_3 As New ADODB.Recordset
Dim RS_TOTAL As ADODB.Recordset
Dim DIM_X
Dim DIM_Y
Dim Dsalida, Dvalor, dimp, dtotal, DIM_CODIGO, DIM_CANTIDAD, DIM_CAMBIO
Dim DIM_ISV, DIM_DESCUENTO, DIM_VALOR, DIM_GR
Dim DIM1
Dim DIM2
Dim PUB_42 As Boolean
Dim ErrNumber As String
Dim ErrSource As String
Dim ErrDescription As String
Dim DIM_INT_TIME As Boolean
Dim DIM_INT_TIME_1 As Boolean
Dim DIM_INT_TIME_2 As Boolean
Dim DIM_INT_TIME_3 As Boolean
Const TOP_MARGIN = 1
Const LEFT_MARGIN = 0

Private Sub BTN5_Click()
If TXT_CANTIDAD.Text = "" Then
DIM_INT_TIME = False
Else
    If Val(TXT_CANTIDAD.Text) < Val(PUB_59) Then
        MsgBox "!!!EL PAGO ES INSUFICIENTE!!!"
        TXT_CANTIDAD.Text = ""
        TXT_CANTIDAD.SetFocus
    Else
        DIM_CANTIDAD = TXT_CANTIDAD.Text
        dtotal = DIM_SUMTOTAL - TXT_CANTIDAD.Text
        lblCB.Caption = Format(dtotal, "#,##0.00")
        DIM_CAMBIO = lblCB
        DIM_SUMTOTAL = lblST
    End If
    DIM_INT_TIME = True
    cmdImprimir.SetFocus
End If
End Sub

Private Sub cmdImprimir_Click()
Timer1.Enabled = True
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Dim CAI1, CAI2


Printer.CurrentY = TOP_MARGIN
Printer.CurrentX = LEFT_MARGIN
Printer.Font.Size = 8
Printer.FontName = "Arial"
Printer.Font.Bold = True
Printer.Font.Size = 12
Dim DIM_TITULO
'Printer.FontName = "FontControl"
For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = DIM_EMPRESA
 HWidth = Printer.TextWidth(DIM_EMPRESA) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
Printer.Font.Size = 9

Printer.Print ""

For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "*** FACTURA DE VENTA ***"
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
Printer.Font.Size = 10
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
Printer.Print ""

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''

For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "*** CAI ***"
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
Printer.Print ""
 CAI1 = "7FF34D-88B807-F742B3"
 For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = CAI1
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
Printer.Print ""
 CAI2 = "F5AF95-D2A7A6-50"
  For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = CAI2
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
Printer.Print ""
 DIM_TITULO = "RTN : 08019008152819"

For i = 1 To 1   ' Set up two iterations.
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
Printer.Print ""
Printer.Print ""
Printer.Font.Size = 9
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "Francisco Morazan,D.C."
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
Printer.Print ""
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "Bo Concepcion, Entre 5y6 Ave Cll 7"
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
Printer.Print ""
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "Contiguo Escuela Lempira"
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
Printer.Print ""
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Printer.Print ""
DIM_TELEFONO = "22382972"
 DIM_TITULO = "Telefono : " & DIM_TELEFONO
For i = 1 To 1   ' Set up two iterations.
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Printer.Print ""
  For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "E-Mail: l_bendicion@hotmail.com"
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Printer.Print ""
Printer.Print ""
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''

DIM_TITULO = "Cajero  : " & PUB_USERNOMBRE
For i = 1 To 1   ' Set up two iterations.
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i

Printer.Print ""


  For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "Cliente :" & DIM_CLIENTE
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "RNT CLNTE :" & DIM_RTNCIENTE
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
Printer.Print ""
Printer.Print ""
Printer.Font.Size = 12
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''

For i = 1 To 1   ' Set up two iterations.

       ' NumAlt = 1 To 1   ' Set up two iterations.
         DIM_TITULO = "FACTURA #"
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '000-001-01-" &
 Printer.Print ""
For i = 1 To 1   ' Set up two iterations.

       ' NumAlt = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "000-001-01-" & DIM_NODOC
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Printer.Print ""
If DimOPTarjeta1 = "1" Then
    DIM_TITULO = "VENTA CON TARJETA "
    For i = 1 To 1   ' Set up two iterations.
     HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
     Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
     Printer.Print DIM_TITULO  ' Send new page.
    Next i
Else
    DIM_TITULO = "VENTA CON EFECTIVO "
    For i = 1 To 1   ' Set up two iterations.
     HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
     Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
     Printer.Print DIM_TITULO  ' Send new page.
    Next i
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Printer.Print ""
Printer.Print ""
Printer.Font.Size = 9
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
Printer.Print "Cant.....Producto.......Valor "
Printer.Print "_______________________________ "
Printer.Print ""
Printer.Font.Size = 10
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$like'" & Item.Text & "' "$$
PRVT_1.Open PUB_CONEXION_EASY
PRVT_2.ActiveConnection = PRVT_1
PRVT_2.CommandType = adCmdText
PRVT_2.CommandText = "SELECT Producto,Salida,Total,TAX  FROM INVSalida WHERE NDVentas like '" & DIM_NODOC & "'"
Set PRVT_3 = PRVT_2.Execute
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Do While Not PRVT_3.EOF
        ' Use rs!FieldName to get the data for
        ' the field named FieldName.
If PRVT_3!TAX = "GRABADO" Then
DIM_GR = "GR"
Else
DIM_GR = "EX"
End If
        Printer.CurrentX = LEFT_MARGIN
        Printer.Print PRVT_3!salida & "      " & PRVT_3!Producto & "     L. " & PRVT_3!Total & "     " & DIM_GR
        'Format$(rs!Titulo) & vbTab & Format$(rs!Formato) & vbTab & Format$(rs!FormatoCompresion) & vbTab & (rs!MinCDs) & vbTab & (rs!NumDVDs) & vbTab & Format$(rs!NumCDs) & vbTab & Format$(rs!Genero) & vbTab & Format$(rs!Subtitulos) & vbTab & Format$(rs!Idioma)
       ' See if we have filled the page.
        If Printer.CurrentY >= bottom_margin Then
            ' Start a new page.
            Printer.NewPage
            Printer.CurrentY = TOP_MARGIN
        End If
        
        PRVT_3.MoveNext
Loop

PRVT_3.Close
PRVT_1.Close
Set PRVT_1 = Nothing
Set PRVT_3 = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''

Printer.Font.Bold = True
Printer.Font.Size = 9
Dsalida = lblST
Dvalor = lblCB
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
Printer.Print "_______________________________ "
 DIM_TITULO = "Sin Impuesto = " & vbTab & vbTab & dimp
 Printer.Print DIM_TITULO  ' Send new page.
 
 DIM_TITULO = "Impuesto S/Ventas = " & vbTab & Format(DIM_SUMISV, "#,##0.00")
 Printer.Print DIM_TITULO  ' Send new page.
 DIM_TITULO = "Total Pagar  = " & vbTab & vbTab & Format(DIM_SUMTOTAL, "#,##0.00")
 Printer.Print DIM_TITULO  ' Send new page.
 DIM_TITULO = "Cambio = " & vbTab & vbTab & DIM_CAMBIO
 Printer.Print DIM_TITULO  ' Send new page.
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
Printer.ForeColor = RGB(0, 0, 0)
dtotal = Dvalor
'Printer.Print "Total = " & vbTab & dtotal
Printer.Print ""
Printer.Print "______________________________"
Printer.Print ""
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "Fecha de Venta " & Date
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "Fecha Limite : 06-08-2016 "
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
 DIM_TITULO = "La Factura es Beneficios de todos"
 Printer.Print DIM_TITULO  ' Send new page.

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''
For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "!EXIJALA¡"
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''

Printer.Print "______________________________ "
Printer.FontName = "Control"
Printer.Print "A"
Printer.EndDoc
NumAlt = NumAlt + 1
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''

End Sub

Private Sub Command13_Click()


End Sub

Private Sub Command12_Click()
TXT_CANTIDAD.Text = ""
End Sub

Private Sub Command23_Click()

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
PUB_42 = False

If IsNull(DIM_SUMTOTAL) Then
Else
lblST.Caption = DIM_SUMTOTAL
End If
'lblCB.Caption = LCB
DIM_CODIGO = CODCAJAA
DIM_INT_TIME_3 = True
DIM_INT_TIME_2 = False

Set RS_TOTAL = New Recordset
RS_TOTAL.Open "Select SUM(TOTAL),SUM(valor),SUM(ISV) from INVSALIDA where NDVentas like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
DIM_SUMTOTAL = RS_TOTAL.Fields(0)

DIM_SUMISV = RS_TOTAL.Fields(2)
DIM_SUMVALOR = RS_TOTAL.Fields(0)

DIM_SUMTOTAL = DIM_SUMTOTAL
lblST = Format(DIM_SUMTOTAL, "#,##0.00")
dimp = DIM_SUMTOTAL - DIM_SUMISV
Label9.Caption = Format(dimp, "#,##0.00")
Label7.Caption = Format(DIM_SUMISV, "#,##0.00")
'Label3.Caption = Format(DIM_SUMDESCUENTO, "#,##0.00")




'DIM_VALOR
'DIM_ISV
'DIM_DESCUENTO

DIM1 = DIM_SUMTOTAL

End Sub
Public Function CALCULAR()


DIM1 = DIM_SUMTOTAL - DIM_DESCUENTO

Label3.Caption = Format(DIM_DESCUENTO, "#,##0.00")
Label7.Caption = Format(DIM_ISV, "#,##0.00")
lblST = "Total Lps." & Format(DIM_SUMTOTAL, "#,##0.00")

'lblCB.Caption = Format(DIM2, "#,##0.00")

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Private Sub Timer1_Timer()
   Unload Me
End Sub

Private Sub TMR_1_Timer()
If DIM_INT_TIME = True Then
 If DIM_INT_TIME_1 = True Then
        cmdImprimir.BackColor = &H800000
     
            If DIM_INT_TIME_2 = True Then
               TMR_1.Enabled = False
            Else
               DIM_INT_TIME_1 = False
               'cmdAct.BackColor = &H8000000F
               Exit Sub
            End If
End If

If DIM_INT_TIME_1 = False Then
      cmdImprimir.BackColor = &H8000000F
        If DIM_INT_TIME_2 = True Then
           TMR_1.Enabled = False
        Else
           DIM_INT_TIME_1 = True
        End If
End If
cmdImprimir.BackColor = &H8000000F
End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If TXT_CANTIDAD.Text = "" Then
 If DIM_INT_TIME_1 = True Then
        TXT_CANTIDAD.BackColor = &H800000
     
            If DIM_INT_TIME_2 = True Then
               TMR_1.Enabled = False
            Else
               DIM_INT_TIME_1 = False
               'cmdAct.BackColor = &H8000000F
               Exit Sub
            End If
End If

If DIM_INT_TIME_1 = False Then
      TXT_CANTIDAD.BackColor = &H8000000F
        If DIM_INT_TIME_2 = True Then
           TMR_1.Enabled = False
        Else
           DIM_INT_TIME_1 = True
        End If
End If
Else
TXT_CANTIDAD.BackColor = &H8000000F
End If

End Sub
Private Sub TXT_CANTIDAD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If TXT_CANTIDAD.Text = "" Then
DIM_INT_TIME = False
Else
    If Val(TXT_CANTIDAD.Text) < Val(PUB_59) Then
        MsgBox "!!!EL PAGO ES INSUFICIENTE!!!"
        TXT_CANTIDAD.Text = ""
        TXT_CANTIDAD.SetFocus
    Else
        DIM_CANTIDAD = TXT_CANTIDAD.Text
        dtotal = lblST - TXT_CANTIDAD.Text
        lblCB.Caption = Format(dtotal, "#,##0.00")
        DIM_CAMBIO = lblCB
        DIM_SUMTOTAL = lblST
    End If
    DIM_INT_TIME = True
    cmdImprimir.SetFocus
End If
End If
End Sub
