VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMCUADRE 
   Caption         =   "CUADRE"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   Icon            =   "FRMCUADRE.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   13845
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "HOY"
      Height          =   735
      Left            =   9600
      TabIndex        =   11
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txtcaja 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """L."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   2
      EndProperty
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
      Left            =   6840
      TabIndex        =   9
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """L."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   2
      EndProperty
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
      Left            =   6840
      TabIndex        =   5
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """L."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   2
      EndProperty
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
      Left            =   6840
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton BTN9 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   12960
      Picture         =   "FRMCUADRE.frx":628A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "SALIR"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdRpt 
      Caption         =   "REPORTE"
      Enabled         =   0   'False
      Height          =   735
      Left            =   11160
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker txtfecha 
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy/mm/dd"
      Format          =   92143617
      CurrentDate     =   39173
   End
   Begin VB.Label LBLVAL1 
      Alignment       =   2  'Center
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   6120
      Width           =   12135
   End
   Begin VB.Label Label1 
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Tarjeta Grabado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Tarjeta Excento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "Efectivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "FRMCUADRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_CUADRE As ADODB.Recordset
Dim RS_TARJETA As ADODB.Recordset
Dim RS_TOTAL As ADODB.Recordset

Dim DIM_SQL As String
'''''''''''''''''''''''''''''''''''''''
Dim Dsalida, Dvalor, dimp, dtotal, resultado
Dim DIM_SUBTOTAL
Dim DIM_TOTAL
Dim DIMGRABADO
Dim DIM_TARJETA
Dim DIMEXCENTO
Dim DIMEFECTIVO
Dim DIM_INT_5
Dim DIM_CONEXION As String


Private Sub BTN9_Click()
Unload Me
End Sub

Private Sub cmdrpt_Click()
On Error Resume Next
Dim a, c
Dim COLORT
''''''''''''''''''''''''''''''''''''''''
Const TOP_MARGIN = 5
Const LEFT_MARGIN = 25

Printer.CurrentY = TOP_MARGIN
Printer.CurrentX = LEFT_MARGIN
'''''''''''''''''''''''''''''''''''''''''''
Printer.Font.Size = 10
Printer.FontName = "FontA1x1"
Printer.Font.Bold = True
'Printer.FontName = "FontControl"
Printer.Print "  ***CUADRE DE CAJA***"
Dim DIM_TITULO
'Printer.FontName = "FontControl"
For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = DIM_EMPRESA
 HWidth = Printer.TextWidth(DIM_EMPRESA) / 2   ' Get one-half width.
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
 Printer.Print DIM_TITULO  ' Send new page.
Next i

Printer.Font.Size = 10
Printer.Print "________________________________"

Printer.Print "Tienda = " & PUB_5
Printer.Print "Fecha = "; txtfecha.Value
Printer.Print "Doc....Producto.....Cant....Valor"
Printer.Print "________________________________"
Printer.Font.Bold = False

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set RS_TARJETA = New Recordset

'DIM_SQL = "select * from Ventas_Dia WHERE Fecha like '" & DIM_INT_5 & "'"      'DIM_DIA_DET = Format(DTPicker1, "dd mmmm yyyy")
Dim SQL_DIM_TARJETA
SQL_DIM_TARJETA = "1"
DIM_SQL = "select sum(total)FROM Ventas_Cuadre_Dia WHERE Fecha like '" & DIM_INT_5 & "'"
'DIM_SQL = DIM_SQL & "AND CAJA= " & DIM_CAJA
'DIM_SQL = DIM_SQL & "AND COLOR NOT LIKE '" & SQL_DIM_TARJETA & "'"
DIM_SQL = DIM_SQL & " AND COLOR IS NULL"
RS_TARJETA.Open DIM_SQL, PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DIM_EFECTIVO = RS_TARJETA.Fields(0)
Set RS_TARJETA = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set RS_TARJETA = New Recordset

'DIM_SQL = "select * from Ventas_Dia WHERE Fecha like '" & DIM_INT_5 & "'"      'DIM_DIA_DET = Format(DTPicker1, "dd mmmm yyyy")

SQL_DIM_TARJETA = "1"
DIM_SQL = "select sum(total)FROM Ventas_Cuadre_Dia WHERE Fecha like '" & DIM_INT_5 & "'"
'DIM_SQL = DIM_SQL & "AND CAJA= " & DIM_CAJA
DIM_SQL = DIM_SQL & "AND COLOR like '" & SQL_DIM_TARJETA & "'"

RS_TARJETA.Open DIM_SQL, PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DIM_TARJETA = RS_TARJETA.Fields(0)
Set RS_TARJETA = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''
Set RS_CUADRE = New Recordset
DIM_SQL = "select * FROM Ventas_Cuadre_Dia WHERE Fecha like '" & DIM_INT_5 & "' order by hora1"
'DIM_SQL = DIM_SQL & "AND CAJA= " & DIM_CAJA
'DIM_SQL = DIM_SQL & " order by hora1"
RS_CUADRE.Open DIM_SQL, PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dvalor = RS_CUADRE.Fields(0)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Do While Not RS_CUADRE.EOF
        ' Use rs!FieldName to get the data for
        ' the field named FieldName.
        If RS_CUADRE!Color = "1" Then
        COLORT = "TJ"
        Else
        COLORT = "EF"
        End If
        Printer.CurrentX = LEFT_MARGIN
        Printer.Print RS_CUADRE!nodoc & ".." & RS_CUADRE!Producto & ".." & RS_CUADRE!salida & ".." & RS_CUADRE!total & "..." & COLORT
        'Format$(rs!Titulo) & vbTab & Format$(rs!Formato) & vbTab & Format$(rs!FormatoCompresion) & vbTab & (rs!MinCDs) & vbTab & (rs!NumDVDs) & vbTab & Format$(rs!NumCDs) & vbTab & Format$(rs!Genero) & vbTab & Format$(rs!Subtitulos) & vbTab & Format$(rs!Idioma)
        ' See if we have filled the page.
        If Printer.CurrentY >= bottom_margin Then
            ' Start a new page.
            Printer.NewPage
            Printer.CurrentY = TOP_MARGIN
        End If
        RS_CUADRE.MoveNext
Loop

Set RS_CUADRE = Nothing
'''''''''''''''''''''''''''''''''''''''''''
Printer.Print "______________________________"



Set RS_CUADRE = New Recordset

DIM_SQL = "select SUM(total) FROM Ventas_Cuadre_Dia WHERE Fecha like '" & DIM_INT_5 & "'"
DIM_SQL = DIM_SQL & "AND CAJA= " & DIM_CAJA

RS_CUADRE.Open DIM_SQL, PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dvalor = RS_CUADRE.Fields(0)

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim resultado1, rslt
Printer.CurrentX = LEFT_MARGIN
rslt = Format(RS_CUADRE.Fields(0), "L ##,###.00")
resultado1 = RS_CUADRE.Fields(0) - DIM_DESCUENTO
Printer.Font.Bold = True
Printer.Font.Size = 10
Printer.Print "Descuento...."; Format(DIM_DESCUENTO, "L ##,###.00")
Printer.Print "Caja Segun Doc...."; rslt
Printer.Print "Valor Efectivo..."; Format(DIM_EFECTIVO, "L ##,###.00")
Printer.Print "Valor Tarjeta..."; Format(DIM_TARJETA, "L ##,###.00")
Printer.Print "Valor - Descuento..."; Format(resultado1, "L ##,###.00")
'Printer.Print "Valor Efectivo..."; Format(txtcaja, "L ##,###.00")


Printer.Print resultado
Printer.Print "_______________________________"

resultado = ""
resultado1 = ""
rslt = ""

Set RS_CUADRE = Nothing
Printer.FontName = "Control"
'Printer.Print "A"
Printer.EndDoc


'Exit Sub
'menerr:
'MsgBox "Hubo un error y no se pudo pudo hacer el backup", vbCritical + vbOKOnly, App.EXEName & ": Error"

End Sub

Private Sub Command1_Click()
On Error GoTo menerr
txtcaja.Text = ""
'lblval.Caption = ""
LBLVAL1.Caption = ""
Exit Sub
menerr:
MsgBox "Hubo un error y no se pudo pudo hacer el backup", vbCritical + vbOKOnly, App.EXEName & ": Error"

End Sub

Private Sub Command13_Click()
DIM_FECHADEL = Format(DTPicker3, "mm/dd/yyyy")
DIM_FECHAAL = Format(DTPicker1, "mm/dd/yyyy")
End Sub

Private Sub Command2_Click()

Set RS_TOTAL = New Recordset
DIM_SQL = "select sum(TOTAL)FROM INVSalida WHERE Fecha like '" & Date & "'"
DIM_SQL = DIM_SQL & " AND COLOR IS NULL"

RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic



If RS_TOTAL.Fields(0) = 0 Then
DIMEFECTIVO = "0"
Else
DIMEFECTIVO = RS_TOTAL.Fields(0)
End If
Set RS_TOTAL = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set RS_TOTAL = New Recordset

Dim SQL_DIM_TARJETA_EX, SQL_DIM_TARJETA_GR
SQL_DIM_TARJETA = "1"
SQL_DIM_TARJETA_EX = "EXCENTO"
DIM_SQL = "select sum(TOTAL) FROM INVSalida WHERE Fecha like '" & Date & "'"
DIM_SQL = DIM_SQL & "AND color LIKE '" & SQL_DIM_TARJETA & "'"
DIM_SQL = DIM_SQL & "AND tax like '" & SQL_DIM_TARJETA_EX & "'"


           'DIM_SQL = "DELETE * FROM INVSalida WHERE Codigo= " & Borrar_Codigo
           'DIM_SQL = DIM_SQL & "AND NDVentas LIKE '" & Borrar_NoDoc & "'"
           'DIM_SQL = DIM_SQL & "AND Hora1 LIKE '" & Borrar_Hora & "'"

RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
If RS_TOTAL.Fields(0) = 0 Or IsNull(RS_TOTAL.Fields(0)) Then
DIMEXCENTO = "0"
Else
DIMEXCENTO = RS_TOTAL.Fields(0)
End If
Set RS_TOTAL = Nothing

Set RS_TOTAL = New Recordset

SQL_DIM_TARJETA_GR = "GRABADO"
DIM_SQL = "select sum(TOTAL) FROM INVSalida WHERE Fecha like '" & Date & "'"
DIM_SQL = DIM_SQL & "AND color LIKE '" & SQL_DIM_TARJETA & "'"
DIM_SQL = DIM_SQL & "AND tax like '" & SQL_DIM_TARJETA_GR & "'"

RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
If RS_TOTAL.Fields(0) = 0 Or IsNull(RS_TOTAL.Fields(0)) Then
DIMGRABADO = "0"
Else
DIMGRABADO = RS_TOTAL.Fields(0)
End If

Set RS_TOTAL = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If txtcaja.Text = DIMEFECTIVO And Text2.Text = DIMEXCENTO And Text3.Text = DIMGRABADO Then
LBLVAL1 = "EL CUADRE DE VENTAS ESTA CORRECTO"
cmdRpt.Enabled = True
Else
LBLVAL1 = "EL CUADRE DE VENTAS ESTA INCORRECTO"
End If

End Sub

Private Sub DTPicker1_CloseUp()
DIM_FECHADEL = Format(DTPicker3, "mm/dd/yyyy")
DIM_FECHAAL = Format(DTPicker1, "mm/dd/yyyy")
DIM_FECHA_HOY = DTPicker1
DIM_SQL = "select SUM(total) from InvSalida where fecha between #" & DIM_FECHADEL & "# and #" & DIM_FECHAAL & "#"
Set RS_CUENTAS_INGRESOS = New Recordset
RS_CUENTAS_INGRESOS.Open DIM_SQL, PUB_CONEXION_PDV_TIENDA, adOpenStatic, adLockOptimistic
'Text11.Text = Format(RS_CUENTAS_INGRESOS.Fields(0), "#,##0.00")
Label13.Caption = Format(RS_CUENTAS_INGRESOS.Fields(0), "#,##0.00")
Set RS_CUENTAS_INGRESOS = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
DIM_SQL = "select fecha,format(Fecha,'dddd dd mmm yyyy'),format(Fecha,'dddd'),sum(total) from InvSalida where fecha between #" & DIM_FECHADEL & "# and #" & DIM_FECHAAL & "#"
DIM_SQL = DIM_SQL & " GROUP BY fecha"
Set RS_DEINFO = New Recordset
 RS_DEINFO.Open DIM_SQL, PUB_CONEXION_PDV_TIENDA, adOpenStatic, adLockOptimistic
 With RS_DEINFO
        If .RecordCount <> 0 Then
        If RS_DEINFO.BOF = True And RS_DEINFO.EOF = True Then
        LV3.ListItems.Clear
        Else
            LV3.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LV3.ListItems.Add(, , .Fields(0) & "")
                Items.SubItems(1) = .Fields(1) & ""
                Items.SubItems(2) = .Fields(2) & ""
                Items.SubItems(3) = Format$(.Fields(3), "#,##0.00") & ""
                DIM_LINEA = DIM_LINEA + 1
                .MoveNext
            Loop
        End If
         End If
    End With
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

End Sub

Private Sub Form_Load()
On Error Resume Next
DIM_X = PRINCIPAL.Width / 2
DIM_Y = PRINCIPAL.Height / 2
'rame1.Top = DIM_Y - 3000
'Frame1.Left = DIM_X - 5000
txtfecha.Value = Date
'Frame1.Visible = True
'Frame2.Visible = False
If PUB_RED = True Then
DIM_CONEXION = PUB_CONEXION_EASY
Else
DIM_CONEXION = PUB_CONEXION_EASY
End If
DIM_CAJA = "1"

DIM_INT_5 = Format(Date, "dd mmmm yyyy")
Dim DIM_SQL As String
Set RS_CUADRE = New Recordset

'DIM_SQL = "select * from Ventas_Dia WHERE Fecha like '" & DIM_INT_5 & "'"      'DIM_DIA_DET = Format(DTPicker1, "dd mmmm yyyy")
'DIM_SQL = DIM_SQL & "AND COLOR NOT LIKE '" & SQL_DIM_TARJETA & "'"



End Sub

Private Sub txtncaja_GotFocus()
txtcaja.Enabled = True
'lblval.Enabled = True

End Sub

Private Sub txtncaja_LostFocus()

End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub txtfecha_CloseUp()
'= Format(DTPicker1, "dd mmmm yyyy")
DIM_INT_5 = Format(txtfecha, "dd mmmm yyyy")
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''





Set RS_TOTAL = New Recordset
DIM_SQL = "select sum(TOTAL)FROM INVSalida WHERE Fecha like '" & txtfecha & "'"
DIM_SQL = DIM_SQL & " AND COLOR IS NULL"

RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic



If RS_TOTAL.Fields(0) = 0 Then
DIMEFECTIVO = "0"
Else
DIMEFECTIVO = RS_TOTAL.Fields(0)
End If
Set RS_TOTAL = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set RS_TOTAL = New Recordset

Dim SQL_DIM_TARJETA_EX, SQL_DIM_TARJETA_GR
SQL_DIM_TARJETA = "1"
SQL_DIM_TARJETA_EX = "EXCENTO"
DIM_SQL = "select sum(TOTAL) FROM INVSalida WHERE Fecha like '" & txtfecha & "'"
DIM_SQL = DIM_SQL & "AND color LIKE '" & SQL_DIM_TARJETA & "'"
DIM_SQL = DIM_SQL & "AND tax like '" & SQL_DIM_TARJETA_EX & "'"


           'DIM_SQL = "DELETE * FROM INVSalida WHERE Codigo= " & Borrar_Codigo
           'DIM_SQL = DIM_SQL & "AND NDVentas LIKE '" & Borrar_NoDoc & "'"
           'DIM_SQL = DIM_SQL & "AND Hora1 LIKE '" & Borrar_Hora & "'"

RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
If RS_TOTAL.Fields(0) = 0 Or IsNull(RS_TOTAL.Fields(0)) Then
DIMEXCENTO = "0"
Else
DIMEXCENTO = RS_TOTAL.Fields(0)
End If
Set RS_TOTAL = Nothing

Set RS_TOTAL = New Recordset
SQL_DIM_TARJETA = "1"
SQL_DIM_TARJETA_GR = "GRABADO"
DIM_SQL = "select sum(TOTAL) FROM INVSalida WHERE Fecha like '" & txtfecha & "'"
DIM_SQL = DIM_SQL & "AND color LIKE '" & SQL_DIM_TARJETA & "'"
DIM_SQL = DIM_SQL & "AND tax like '" & SQL_DIM_TARJETA_GR & "'"

RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
If RS_TOTAL.Fields(0) = 0 Or IsNull(RS_TOTAL.Fields(0)) Then
DIMGRABADO = "0"
Else
DIMGRABADO = RS_TOTAL.Fields(0)
End If

Set RS_TOTAL = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If txtcaja.Text = DIMEFECTIVO And Text2.Text = DIMEXCENTO And Text3.Text = DIMGRABADO Then
LBLVAL1 = "EL CUADRE DE VENTAS ESTA CORRECTO"
cmdRpt.Enabled = True
Else
LBLVAL1 = "EL CUADRE DE VENTAS ESTA INCORRECTO"
End If
End Sub

Private Sub txtfecha_LostFocus()
DIM_INT_5 = Format(txtfecha, "dd mmmm yyyy")
End Sub

