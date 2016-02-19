VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FBFD55C6-C23C-11D3-B65D-004005E66149}#1.0#0"; "swiftprint.ocx"
Begin VB.Form Eliminarventas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   10770
   ClientLeft      =   -420
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "EliminarVentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1000
      Left            =   8040
      Picture         =   "EliminarVentas.frx":5719A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton BTN9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1000
      Left            =   5640
      Picture         =   "EliminarVentas.frx":5B3A5
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton BTN8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1000
      Left            =   4590
      Picture         =   "EliminarVentas.frx":5F7BC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton BTN1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1000
      Index           =   2
      Left            =   2280
      Picture         =   "EliminarVentas.frx":6390B
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton BTN1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1000
      Index           =   1
      Left            =   1065
      Picture         =   "EliminarVentas.frx":67BE3
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton BTN1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1000
      Index           =   0
      Left            =   -30
      Picture         =   "EliminarVentas.frx":6BDE9
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1100
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton BTN1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1000
      Index           =   3
      Left            =   3375
      Picture         =   "EliminarVentas.frx":701F0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin MSComctlLib.ListView LST_INVT 
      Height          =   8895
      Left            =   4080
      TabIndex        =   0
      Top             =   1680
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   15690
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LV 
      Height          =   8895
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   15690
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
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
      Left            =   11400
      TabIndex        =   15
      Top             =   1440
      Width           =   3495
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
      Height          =   990
      Left            =   6120
      TabIndex        =   14
      Top             =   1440
      Width           =   4305
   End
   Begin SwiftPrintLib.SwiftPrint SpDoc 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin VB.Label LBL_DOC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   450
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VENTAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   450
      Left            =   4920
      TabIndex        =   3
      Top             =   1080
      Width           =   2130
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "Eliminarventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'FIXIT: Declare 'DIM_INT_1' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_1
'FIXIT: Declare 'DIM_INT_2' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_2
'FIXIT: Declare 'DIM_INT_3' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_3
'FIXIT: Declare 'DIM_INT_4' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_4
'FIXIT: Declare 'DIM_INT_5' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_5
'FIXIT: Declare 'DIM_INT_6' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_6
'FIXIT: Declare 'DIM_VALOR' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_VALOR, PUB_10
Dim ID As Long
    Dim ImgWidth As Long, ImgHeight As Long
    Dim imgType As Long
Dim DIM_INT_TIME_1 As Boolean
Dim DIM_INT_TIME_2 As Boolean
Dim DIM_INT_7 As Boolean
Dim Borrar_NoDoc
Dim DIM_ELIMINAR
Dim DIM_FORMA
Dim DIM_SQL
Dim DIM_FECHA
Dim DIM_INT_8 As Boolean
Dim RS_VENTAS As ADODB.Recordset
Dim RS_TOTAL As ADODB.Recordset
Dim RS_SALIDA As ADODB.Recordset
Dim RS_CLIENTEINFO As ADODB.Recordset
Dim RS_SALIDA_2 As ADODB.Recordset
Dim nocli As ADODB.Recordset
Dim RS_VASIO As ADODB.Recordset
Dim RS_BORRAR As ADODB.Recordset
Private PRVT_1 As New ADODB.Connection
Private PRVT_2 As New ADODB.Command
Private PRVT_3 As New ADODB.Recordset
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Private Sub refrescar()
'On Error GoTo menerr
Dim DIM_SQL
'If RS_ventas.State <> adStateClosed Then
If RS_VENTAS.EOF = True And RS_VENTAS.BOF = True Then
        PUBLIC_SUB_LOCK
        Limpiar_lstvDatos
        LBL_DOC.Caption = ""
        MsgBox "AGREGAR UNA NUEVA FACTURA"
        Modificar_Click
        Exit Sub
End If
   

        
        'If RS_VENTAS.EOF = True Or RS_VENTAS.BOF = True Then
        'RS_VENTAS.MoveLast
        'refrescar
        'End If



        'LBL_DOC = "DOC : " & PUB_10
        Set RS_SALIDA = New Recordset

        
 '       RS_SALIDA_2.Open "SELECT Codigo,Producto,Salida,Descuento,punitario,ISV,Valor,total,ClientE,NDVentas,ID,fecha,Hora1,tax,Tipo,Descripcion,cliente,Costo,NoDE,caja,DEI,FORMA,vendedor,TARJETA,COLOR FROM INVSalida WHERE ndventas like '" & RS_VENTAS.Fields("NDVentas") & "' ", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
'        DIM_SQL = "SELECT Codigo,Producto,Salida,punitario,valor,ncliente,cliente,fecha,Hora1,Descripcion,NDVentas,forma,total,Vendedor FROM INVSalida WHERE ndventas like '" & RS_VENTAS.Fields("NDVentas") & "' "

        DIM_SQL = "SELECT Codigo,Producto,Salida,punitario,NDVentas,TOTAL,ncliente,cliente,fecha,Hora1,Descripcion FROM INVSalida WHERE ndventas like '" & PUB_10 & "' "
        
        'ndventas like '" & RS_VENTAS.Fields("NDVentas") & "' "
        RS_SALIDA.Open DIM_SQL, PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic

        Carga_lstvDatos
        'MsgBox pic


End Sub
'''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'Private Const PUB_CONEXION_EASY = "DSN=EASI"
Private Sub cmdAct_Click()
'On Error GoTo menerr
DIM_INT_TIME_2 = True
PUB_42 = False
cmdAct.BackColor = &H8000000F
'RS_ventas.Update
'RS_SALIDA.Update
'RptCaja.WindowState = crptMaximized
'RptCaja.Action = 1
'''''''''''''''''''''''''''''''''''
    PRVT_1.Open PUB_CONEXION_EASY
    PRVT_2.ActiveConnection = PRVT_1
    PRVT_2.CommandType = adCmdText
    PRVT_2.CommandText = "SELECT SUM(Valor),SUM(Salida) As UnitsSold FROM InvSalida WHERE NDVentas like '" & PUB_10 & "' "
    Set PRVT_3 = PRVT_2.Execute
    '
   If Not IsNull(PRVT_3.Fields(0)) Then
    valor = PRVT_3.Fields(0)
    cantidad = PRVT_3.Fields(1)
'    Text2.Text = cantidad
    
    valor1 = Format(valor, "#,##0.00")
'Impuesto = valor1 * 12 / 100
'Impuesto1 = Format(Impuesto, "#,##0.00")
PUB_59 = Val(valor) '+ Val(Impuesto)
PUB_32 = Format(PUB_59, "#,##0.00")
'Text3.Text = Impuesto1
'Text4.Text = Subtotal1
'Text1.Text = valor1
   End If
    '
    PRVT_3.Close
    PRVT_1.Close
'''''''''''''''''''''''''''''
' If Val(cantidad) < Val(Subtotal1) Then
'     Dim a, DIM_INT_5, c
'     If TXT_CANTIDAD.Text = "" Then
'    Else
'    a = MsgBox("no se permite esta operacion", vbCritical, "CAMBIO")
'    End If
' Else
'    DIM_INT_5 = cantidad - Subtotal1
'    c = Format(DIM_INT_5, "#,##0.00")
'    'Text7.Text = c
'   LCB = c
'   lST = Subtotal1
' End If
DIM_INT_7 = True
DIM_INT_TIME_2 = False
PUB_7 = PUB_10
FRMIMPRIMIR.Show vbModal

End Sub

Private Sub Cmd_Buscar_Click()
'FIXIT: Declare 'a' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE
Dim a
a = InputBox("Buscar el documento", "BUSCAR")
Dim criterio1 As String

If a = "" Then
Else
'criterio1 = "[nombre]like '*" & a & "*'"
criterio1 = "[NoDoc]=" + a
RS_VENTAS.MoveFirst
RS_VENTAS.Find criterio1
refrescar
End If
End Sub


Private Sub Eliminar_Click()
'On Error GoTo menerr
FRMELIMINAR.Show vbModal
If AceptarE = True Then

    Dim Mens As Integer
      Mens = MsgBox("¿Desea borrar el registro?", vbYesNo + vbQuestion, "Atencion")
        If Mens = 6 Then

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.CurrentY = TOP_MARGIN
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.CurrentX = LEFT_MARGIN
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Font.Size = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.FontName = "FontA1x1"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Font.Bold = True
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Font.Size = 14
'FIXIT: Declare 'DIM_TITULO' con un tipo de datos de enlace en tiempo de compilación       FixIT90210ae-R1672-R1B8ZE
Dim DIM_TITULO
'Printer.FontName = "FontControl"
For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "*** FACTURA ELIMINADA ***"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print DIM_TITULO  ' Send new page.
Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Font.Size = 12
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Print "_______________________________"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For i = 1 To 1   ' Set up two iterations.
 DIM_TITULO = "No Factura = " & PUB_10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print DIM_TITULO  ' Send new page.
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Print "_______________________________ "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Font.Bold = False
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Font.Size = 9
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
PRVT_1.Open PUB_CONEXION_EASY
PRVT_2.ActiveConnection = PRVT_1
PRVT_2.CommandType = adCmdText
PRVT_2.CommandText = "SELECT Codigo,Producto,Salida,Valor  FROM InvSalida WHERE NDVentas like '" & PUB_10 & "' "

Set PRVT_3 = PRVT_2.Execute
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If PRVT_3.EOF Then
Else
Dsalida = PRVT_3.Fields(0)
End If

If PRVT_3.EOF Then
Else
Dvalor = PRVT_3.Fields(1)
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Do While Not PRVT_3.EOF
        ' Use rs!FieldName to get the data for
        ' the field named FieldName.
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.CurrentX = LEFT_MARGIN
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        Printer.Print PRVT_3!codigo & "   " & PRVT_3!Producto & "   Cant." & PRVT_3!salida & "   Lps." & PRVT_3!valor
        'Format$(rs!Titulo) & vbTab & Format$(rs!Formato) & vbTab & Format$(rs!FormatoCompresion) & vbTab & (rs!MinCDs) & vbTab & (rs!NumDVDs) & vbTab & Format$(rs!NumCDs) & vbTab & Format$(rs!Genero) & vbTab & Format$(rs!Subtitulos) & vbTab & Format$(rs!Idioma)
       ' See if we have filled the page.
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
        If Printer.CurrentY >= bottom_margin Then
            ' Start a new page.
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
            Printer.NewPage
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
            Printer.CurrentY = TOP_MARGIN
        End If
        
        PRVT_3.MoveNext
Loop

PRVT_3.Close
PRVT_1.Close
Set PRVT_1 = Nothing
Set PRVT_3 = Nothing
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Font.Size = 10
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Print "_______________________________ "
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.FontName = "Control"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.Print "A"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
Printer.EndDoc
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    PRVT_1.Open PUB_CONEXION_EASY
    PRVT_2.ActiveConnection = PRVT_1
    PRVT_2.CommandType = adCmdText
    PRVT_2.CommandText = "DELETE * FROM InvSalida WHERE ndventas like '" & RS_VENTAS.Fields("NDVentas") & "' "
    Set PRVT_3 = PRVT_2.Execute
    PRVT_1.Close
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'PRVT_1.Open PUB_CONEXION_EASY_A
    'PRVT_2.ActiveConnection = PRVT_1
    'PRVT_2.CommandType = adCmdText
    'PRVT_2.CommandText = "DELETE * FROM InvSalida WHERE NDVentaslike '" & RS_VENTAS.Fields("NDoc") & "' "
    'Set PRVT_3 = PRVT_2.Execute
    'PRVT_1.Close
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Set nocli = New Recordset
    DIM_SQL = "DELETE * FROM Ventas WHERE nodoc= " & RS_VENTAS.Fields("NDVentas")
    nocli.Open DIM_SQL, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
    'nocli.Close
    'RS_ventas.Delete
    RS_VENTAS.MovePrevious
    RS_VENTAS.MoveNext
    RS_VENTAS.MoveFirst
    
    If RS_VENTAS.State <> adStateClosed Then RS_VENTAS.Close
    RS_VENTAS.Open "select * from InvSalida Order by NDVentas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    refrescar
End If
End If

End Sub


Private Sub Form_Activate()
'Set pic = Nothing
If PUB_15 = True Then
    PUB_42 = True
    TMR_1.Enabled = True
    PUB_15 = False
End If
End Sub

Private Sub Form_Load()
'On Error GoTo menerr
'invisible

Set RS_VENTAS = New Recordset


Dim DIM_SQL, DIM_FORMA


     'DIM_FORMA = "CONTADO"
     DIM_SQL = "select * from INVSALIDA ORDER BY ndventas ASC "
      
RS_VENTAS.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'MsgBox PUB_CONEXION_EASY
lstvDatos_a_cero
PUB_42 = False
PUB_17 = True
refrescar

DIM_INT_7 = True
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Set RS_CUENTAS_INGRESOS = New Recordset

RS_CUENTAS_INGRESOS.Open "Select NoDoc4 from Ventas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

If RS_CUENTAS_INGRESOS.EOF = True Or RS_CUENTAS_INGRESOS.BOF = True Then
Else
lbl_total = Format(RS_CUENTAS_INGRESOS.Fields(0), "#,##0.00")
End If
   With LV
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NoDoc4", 5000

        
    End With
  With RS_CUENTAS_INGRESOS
        If RS_CUENTAS_INGRESOS.BOF = True And RS_CUENTAS_INGRESOS.EOF = True Then
        LV.ListItems.Clear
        Else
            LV.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LV.ListItems.Add(, , .Fields(0) & "")

                .MoveNext
            Loop
        End If
    End With
RS_CUENTAS_INGRESOS.Close
Set RS_CUENTAS_INGRESOS = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''


End Sub
Private Sub Form_Unload(Cancel As Integer)
'If RS_VENTAS.Fields("NDVentas") = "" Or RS_VENTAS.Fields("NDVentas") = 0 Then
'RS_InvSalida.CancelUpdate
DIM_DIRECCION = ""
DIM_NODOCB = ""
PUB_USUARIO = ""
PUB_VENDEDOR = ""
PUB_CLIENTE = ""
DIM_DIRECCIONCL = ""
DIM_FECHAPAGO = ""
PUB_CLIENTE = ""
DIM_CODCLIENTE = ""
                DIM_DIRECCIONCL = ""
                DIM_CODCLIENTE = ""
                DIM_CONTACTO = ""
                DIM_CIUDAD = ""
                DIM_DIRECCIONCL = ""
                PUB_CLIENTE = ""
End Sub

Private Sub Modificar_Click()
Dim i As Integer
For i = 0 To 3
  BTN1(i).Enabled = False
Next i

For i = 0 To 1
Command2.Enabled = True
Next i
'desabilitar_botones
End Sub

Private Sub Image3_Click()

Help.Show vbModal
End Sub

Private Sub LST_INVT_DblClick()
PUB_17 = False
FRM_INVT.Show vbModal
End Sub

Private Sub Nuevo_Click()
'On Error GoTo menerr
DIM_INT_7 = False
NUEVO.BackColor = &H8000000F
For i = 0 To 3
  BTN1(i).Enabled = True
Next i
For i = 0 To 1
Command2.Enabled = True
Next i
If RS_VENTAS.EOF = True Or RS_VENTAS.BOF = True Then
With RS_VENTAS
    .AddNew
        RS_VENTAS.Fields("NDOC") = "1"
        PUB_10 = "1"
        RS_VENTAS.Fields("fecha") = Date
        RS_VENTAS.Fields("tienda") = PUB_5
        RS_VENTAS.Fields("caja") = PUB_6
        RS_VENTAS.Fields("Diferenciacion") = 1
        RS_VENTAS.Update
        
        Set RS_SALIDA = New Recordset
'        RS_SALIDA.Open "SELECT Codigo,Producto,salida,Descuento,Valor,Saldo,NDVentas,Periodot,fecha,Hora1,caja FROM InvSalida WHERE NDVentaslike '" & RS_VENTAS.Fields("NDoc") & "' ", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
         RS_SALIDA.Open "SELECT Codigo,Producto,Salida,punitario,dventas,total,ncliente,cliente,fecha,Hora1,Descripcion,NDVentas,forma,Vendedor FROM InvSalida WHERE NDVentas like '" & RS_VENTAS.Fields("NDoc") & "' ", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
        'like '" & RS_VENTAS.Fields("NDoc") & "' "
        'like '" & RS_VENTAS.Fields("NDoc") & "' "
        Carga_lstvDatos
End With
Else
RS_VENTAS.MoveFirst
RS_VENTAS.MoveLast
PUB_10 = RS_VENTAS.Fields("NDVentas") + 1
Limpiar_lstvDatos
RS_VENTAS.AddNew
''''''''''''''''''''''''''''''''
RS_VENTAS.Fields("NDVentas") = PUB_10
RS_VENTAS.Fields("fecha") = Date
RS_VENTAS.Fields("tienda") = PUB_5
RS_VENTAS.Fields("caja") = PUB_6
'RS_ventas.Fields("Diferenciacion") = 1
RS_VENTAS.Update
End If
LST_INVT.SetFocus
PUB_17 = False
FRM_INVT.Show vbModal

End Sub

Private Sub Salir_Click()
Unload Me
End Sub

Public Sub ClearCell()

With GrdInventario
 .Columns(0).Text = ""
 .Columns(1).Text = ""
 .Columns(2).Text = ""
 .Columns(3).Text = ""
 .Columns(4).Text = ""
 .Columns(5).Text = ""
 Cancel = True
 .ReBind
 .Refresh
End With
End Sub
Public Sub CALCULAR()
'On Error GoTo menerr
'FIXIT: Declare 'a' and 'X' and 'c' and 'd' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim a, X, c, d
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If GrdInventario.Columns(2).Text <> "" Then
    a = GrdInventario.Columns(5).Text - GrdInventario.Columns(2)
     GrdInventario.Columns(5).Text = a
Else
    GrdInventario.Columns(5).Text = DIM_INT_3
End If
If GrdInventario.Columns(5) < 5 Then
    c = MsgBox("PELIGRO DEFICIENCIA EN INVENTARIO EN ESTE GRUPO", vbCritical, "INVENTARIO EN PELIGRO")
    
End If
   '    Cancel = True
   '     grdinventario.ReBind
   '     RS_SALIDA.CancelBatch
   
'saldo = ""


End Sub

Public Sub Imprimir()

End Sub

Public Sub Desc()
If IsNumeric(GrdInventario.Columns(5)) Then
    DIM_INT_4 = DIM_INT_5 - GrdInventario.Columns(5)
Else
    DIM_INT_4 = DIM_INT_5
End If
End Sub

Public Sub StatusB()
'On Error GoTo menerr
'FIXIT: Declare 'en' and 'sa' and 'de' and 'so' and 'fa' and 'DIM_TOTAL_ENTRADAS' and 'DIM_TOTAL_SALIDA' and 'saldo1' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim en, sa, de, so, fa, DIM_TOTAL_ENTRADAS, DIM_TOTAL_SALIDA, saldo1
    PRVT_1.Open PUB_CONEXION_EASY
    PRVT_2.ActiveConnection = PRVT_1
    PRVT_2.CommandType = adCmdText
    
    PRVT_2.CommandText = "SELECT SUM(Entrada),SUM(Faltante),SUM(Sobrante),SUM(Devolucion) As UnitsSold FROM InvSalida WHERE Codigo=" & DIM_INT_6
    Set PRVT_3 = PRVT_2.Execute
    '
    If Not IsNull(PRVT_3.Fields(0)) Then
    en = PRVT_3.Fields(0)
    sa = PRVT_3.Fields(1)
    fa = PRVT_3.Fields(2)
    so = PRVT_3.Fields(3)
    'de = PRVT_3.Fields(4)
    End If
    '
    PRVT_3.Close
    PRVT_1.Close
    stb.Panels.Item(1).Text = "Entrada : " & en
    stb.Panels.Item(2).Text = "Salida : " & sa
    stb.Panels.Item(3).Text = "Devolucion : " & de
    stb.Panels.Item(4).Text = "Sobrante : " & so
    stb.Panels.Item(5).Text = "Faltante : " & fa
  
    DIM_TOTAL_ENTRADAS = en + so
    DIM_TOTAL_SALIDA = sa + fa + de
    saldo1 = DIM_TOTAL_ENTRADAS - DIM_TOTAL_SALIDA
    stb.Panels.Item(6).Text = "Saldo : " & saldo1

End Sub
Public Sub PUBLIC_SUB_LOCK()
'Eliminar.Enabled = False
'cmdAct.Enabled = False

For i = 0 To 1
Command2.Enabled = False
Next i
For i = 0 To 3
  BTN1(i).Enabled = False
Next i

End Sub
Public Sub PUBLIC_SUB_UNLOCK()


End Sub
'FIXIT: Declare 'lstvDatos_Ingresar' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function lstvDatos_Ingresar()
'On Error GoTo menerr
With RS_SALIDA
.AddNew
.Fields("Codigo") = PUB_28
.Fields("Producto") = PUB_29
.Fields("Salida") = PUB_30
.Fields("Entrada") = "0"
.Fields("Faltante") = "0"
.Fields("Sobrante") = "0"
.Fields("Devolucion") = "0"
.Fields("Descuento") = PUB_66
.Fields("Valor") = PUB_32
'.Fields("Saldo") = PUB_41
.Fields("NDVentas") = PUB_10
.Fields("Periodot") = PUB_2 'Format(Date, "mmyyyy")
.Fields("Fecha") = Date
.Fields("Caja") = cajanom
.Fields("Hora1") = Format(Time, "Long Time")
.Fields("punitario") = PUB_42
.Fields("Descripcion") = "Ventas de Mercaderia en Tienda"
.Fields("Semana") = PUB_22
.Fields("mes") = PUB_23
.Fields("año") = PUB_31
.Fields("dia") = PUB_24
.Update
refrescar
Carga_lstvDatos
End With
Exit Function
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmventasinformacion,lvsingresar"
Close #1
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
End Function
'FIXIT: Declare 'Limpiar_lstvDatos' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function Limpiar_lstvDatos()
            LST_INVT.ListItems.Clear
End Function
'FIXIT: Declare 'Carga_lstvDatos' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function Carga_lstvDatos()
'On Error GoTo menerr
 With RS_SALIDA
        If RS_SALIDA.BOF = True And RS_SALIDA.EOF = True Then
        LST_INVT.ListItems.Clear
        Else
            LST_INVT.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LST_INVT.ListItems.Add(, , .Fields(0) & "")
                Items.SubItems(1) = .Fields(1) & ""
                Items.SubItems(2) = .Fields(2) & ""
                Items.SubItems(3) = .Fields(3) & ""
                Items.SubItems(4) = .Fields(4) & ""
                Items.SubItems(5) = .Fields(5) & ""
                Items.SubItems(6) = .Fields(6) & ""
                Items.SubItems(7) = .Fields(7) & ""
                Items.SubItems(8) = .Fields(8) & ""
                Items.SubItems(9) = .Fields(9) & ""
                Items.SubItems(10) = .Fields(10) & ""
                'Items.SubItems(11) = .Fields(11) & ""
                .MoveNext
            Loop
        End If
    End With
    Exit Function
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmventasinformacion,lvscargar"
Close #1
End Function
'FIXIT: Declare 'lstvDatos_a_cero' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function lstvDatos_a_cero()
'On Error GoTo menerr
'Aspecto de listview
    With LST_INVT
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Codigo", 1000
        .ColumnHeaders.Add , , "Producto", 2500
        .ColumnHeaders.Add , , "Salida", 500
        .ColumnHeaders.Add , , "Precio Unt", 1000
        .ColumnHeaders.Add , , "NDoc", 1000
        .ColumnHeaders.Add , , "Total", 1300
        .ColumnHeaders.Add , , "NCliente", 1300
        .ColumnHeaders.Add , , "Cliente", 1800
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Hora", 1000
        .ColumnHeaders.Add , , "Descripcion", 1300
        .ColumnHeaders.Add , , "Vendedor", 1000

    End With
        Exit Function
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmventasinformacion,lvscero"
Close #1
End Function
Private Sub TMR_1_Timer()
cmdAct.Enabled = True
If PUB_42 = True Then
 If DIM_INT_TIME_1 = True Then
        cmdAct.BackColor = &H800000
     
            If DIM_INT_TIME_2 = True Then
               TMR_1.Enabled = False
            Else
               DIM_INT_TIME_1 = False
               'cmdAct.BackColor = &H8000000F
               Exit Sub
            End If
End If

If DIM_INT_TIME_1 = False Then
      cmdAct.BackColor = &H8000000F
        If DIM_INT_TIME_2 = True Then
           TMR_1.Enabled = False
        Else
           DIM_INT_TIME_1 = True
        End If
End If
cmdAct.BackColor = &H8000000F
End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'cmdAct.Enabled = True
If PUB_17 = True Then
If DIM_INT_7 = True Then
 If DIM_INT_TIME_1 = True Then
        NUEVO.BackColor = &H800000
     
            If DIM_INT_TIME_2 = True Then
               TMR_1.Enabled = False
            Else
               DIM_INT_TIME_1 = False
               'cmdAct.BackColor = &H8000000F
               Exit Sub
            End If
End If

If DIM_INT_TIME_1 = False Then
      NUEVO.BackColor = &H8000000F
        If DIM_INT_TIME_2 = True Then
           TMR_1.Enabled = False
        Else
           DIM_INT_TIME_1 = True
        End If
End If
NUEVO.BackColor = &H8000000F
End If
End If
End Sub
'FIXIT: Declare 'VASIO' con un tipo de datos de enlace en tiempo de compilación            FixIT90210ae-R1672-R1B8ZE
Public Function VASIO()
'On Error GoTo menerr
'FIXIT: Declare 'DIM_SQL_VASIO' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim DIM_SQL_VASIO
'FIXIT: Declare 'DIM_A' con un tipo de datos de enlace en tiempo de compilación            FixIT90210ae-R1672-R1B8ZE
Dim DIM_A

Set nocli = New Recordset
Set RS_VASIO = New Recordset
Set RS_BORRAR = New Recordset

    DIM_SQL = "SELECT * FROM Ventas"
    nocli.Open DIM_SQL, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly

Do While Not nocli.EOF
DIM_A = nocli.Fields("nodoc")
    DIM_SQL_VASIO = "SELECT * FROM InvSalida WHERE ndventas like '" & DIM_A & "' "
    
    RS_VASIO.Open DIM_SQL_VASIO, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
    
    If RS_VASIO.EOF Then
           
    DIM_SQL_VASIO = "DELETE * FROM Ventas WHERE nodoclike '" & DIM_A & "' "
    RS_BORRAR.Open DIM_SQL_VASIO, PUB_CONEXION_EASY, adOpenStatic, adLockOptimistic
'    RS_BORRAR.Close
    End If
        RS_VASIO.Close
        RS_VENTAS.MoveFirst
        RS_VENTAS.MoveLast
nocli.MoveNext
Loop
nocli.Close
    Exit Function
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmventasinformacion,vasio"
Close #1
End Function

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim criterio1
criterio1 = Item.Text
PUB_10 = criterio1
'*****************************************************************************************************************
'criterio1 = "[nombre]like '*" & a & "*'"
criterio1 = "[NDVentas]=" + criterio1
RS_VENTAS.MoveFirst
'RS_VENTAS.Find criterio1

refrescar

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next


Set RS_SALIDA_2 = New Recordset
DIM_SQL = "SELECT Codigo,Producto,Salida,punitario,NDVentas,TOTAL,ncliente,cliente,fecha,Hora1,Descripcion,forma,tienda FROM INVSalida WHERE ndventas like '" & Text1.Text & "' "
RS_SALIDA_2.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic


lstvDatos_a_cero
 With RS_SALIDA_2
        If RS_SALIDA_2.BOF = True And RS_SALIDA_2.EOF = True Then
        LST_INVT.ListItems.Clear
        Else
            LST_INVT.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LST_INVT.ListItems.Add(, , .Fields(0) & "")
                Items.SubItems(1) = .Fields(1) & ""
                Items.SubItems(2) = .Fields(2) & ""
                Items.SubItems(3) = .Fields(3) & ""
                Items.SubItems(4) = .Fields(4) & ""
                Items.SubItems(5) = .Fields(5) & ""
                Items.SubItems(6) = .Fields(6) & ""
                Items.SubItems(7) = .Fields(7) & ""
                Items.SubItems(8) = .Fields(8) & ""
                Items.SubItems(9) = .Fields(9) & ""
                Items.SubItems(10) = .Fields(10) & ""
                'Items.SubItems(11) = .Fields(11) & ""
                .MoveNext
            Loop
        End If
    End With
RS_SALIDA_2.Close



End Sub
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************
'*****************************************************************************************************************


Private Sub BTN9_Click()
Unload Me
End Sub
Private Sub Command1_Click()
Dim DIMFORMA777

Set RS_SALIDA_2 = New Recordset
DIM_SQL = "select * from InvSalida where ndventas like '" & RS_VENTAS.Fields("NDVentas") & "' "

RS_SALIDA_2.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic


DIMFORMA777 = RS_SALIDA_2.Fields("FORMA")
        If IsNull(RS_SALIDA_2.Fields("CLIENTE")) Or RS_SALIDA_2.Fields("CLIENTE") = "0" Then
        DIM_DIRECCIONCL = "N/A"
        Else
        Set RS_CLIENTEINFO = New Recordset
        DIM_SQL = "select * from Clientes where nombre like '" & RS_SALIDA_2.Fields("CLIENTE") & "'"
        RS_CLIENTEINFO.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
                If RS_CLIENTEINFO.BOF = True And RS_CLIENTEINFO.EOF = True Then
                DIM_DIRECCIONCL = "N/A"
                DIM_DIRECCIONCL = ""
                DIM_CODCLIENTE = ""
                DIM_CONTACTO = ""
                DIM_CIUDAD = ""
                DIM_DIRECCIONCL = ""
                Else
                DIM_DIRECCIONCL = RS_CLIENTEINFO.Fields("DIRECCION")
                PUB_CLIENTE = RS_CLIENTEINFO.Fields("nombre")
                DIM_CODCLIENTE = RS_CLIENTEINFO.Fields("codigo")
                DIM_CONTACTO = RS_CLIENTEINFO.Fields("contacto")
                DIM_CIUDAD = RS_CLIENTEINFO.Fields("ciudad")
                DIM_DIRECCIONCL = RS_CLIENTEINFO.Fields("direccion")
                End If
        End If



RS_SALIDA_2.Close

DIM_NODOC = RS_VENTAS.Fields("NDVentas")
'PUB_FECHA = RS_VENTAS.Fields("fecha")
PUB_FECHA = LST_INVT.SelectedItem.ListSubItems.Item(8).Text
Dim NumLineas
NumLineas = 1
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    SpDoc.DocBegin
    
    'spDoc.SetTitle "Sample PreviewPrint Document", "PreviewPrintSample"
    SpDoc.WindowTitle = "Sample PreviewPrint Document - PreviewPrintSample"
    'spDoc.StartPage = 1
    SpDoc.FirstPage = 1
    'spDoc.Orientation = PPO_PORTRAIT
    SpDoc.PageOrientation = SPOR_PORTRAIT
    Dim nRows As Long, nCols As Long, nItem As Long
    SpDoc.Units = SPUN_LOMETRIC
        nPage = 1
        Dim nCentre As Long
    ' SpDoc.PaperSize = vbPRPSEnv14
    '= Envelope #14 5 x 11 1/2
    'Envelope #14 5 x 11 1/2
        SpDoc.TextAlign = SPTA_LEFT
         SpDoc.BackMode = SPBM_TRANSPARENT
    '''''''''''''''TELEFONO''''''''''''''''''''''
    SpDoc.SetFont "Arial", 250, SPFS_POINTS + 0, 0
    SpDoc.TextOut 800, 280, "FACTURA" '& nPage
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If DIM_TIENDA = "ARROZ" Then
'Set pic = LoadPicture(App.Path & "\iconos\ARROZ.bmp")
    nCentre = SpDoc.PageWidth / 2
    SpDoc.LoadImage App.Path & "\iconos\ARROZ.bmp", ID
    SpDoc.PlaceImage ID, 750, 150, 1200, 250, 10
'    SpDoc.LoadImage App.Path & "\iconos\ARROZ.bmp", id
'    SpDoc.PlaceImage id, 1700, 2000, 400, 1100, 10
           nRows = 50
    'draw the page title in black arial 24pt bold underline,
    'centred and starting 25mm down from the page top
    'spDoc.SetTextColor vbBlack
    SpDoc.ForeColor = vbBlack
    'spDoc.SetFont "Arial", 240, PPF_BOLD + PPF_UNDERLINE, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.SetFont "Arial", 150, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 50, nRows, "BENEFICIO DE ARROZ MATURAVE S.A." '& nPage
    
        ''''''''''''''''''CAI'''''''''''''''''''''''''''
    nRows = nRows + 10
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 2000, nRows, "CAI: 0AC6A0-605A1D-6D49AE-EE3C35-7D2D94-33" '& nPage
    nRows = nRows + 120
        ''''''''''''''''''''''''FACTURA
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.TextOut 2000, nRows, " Factura No:  000-001-01-" & DIM_NODOC
       ''''''''''''''''''''''''FECHA
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD + 0, 0
    nRows = nRows + 80
    SpDoc.TextOut 2000, nRows, " FECHA : " & PUB_FECHA

    ''''direccion local propio''''''''''''''''''''''''''''''
       SpDoc.SetFont "Arial", 80, SPFS_POINTS + 0, 0
   nRows = 180
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Francisco Morazan,D.C." '& nPage
  nRows = nRows + 40

    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "Kilometro 9 Carretera Olancho" '& nPage
  nRows = nRows + 40
    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "RTN: 08019007068267" '& nPage
  nRows = nRows + 40
    ''''''''''EMAIL''''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "E-Mail: maturave9@yahoo.com" '& nPage
  nRows = nRows + 40
      SpDoc.TextAlign = SPTA_RIGHT
    '''''''''''''''TELEFONO''''''''''''''''''''''
    SpDoc.TextOut 2000, nRows, "Telefono: 504 22916125"  '& nPage


End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If DIM_TIENDA = "MAIZ" Then

'Set pic = LoadPicture(App.Path & "\iconos\ARROZ.bmp")
    nCentre = SpDoc.PageWidth / 2
    SpDoc.LoadImage App.Path & "\iconos\MATURAVE1.bmp", ID
SpDoc.PlaceImage ID, 750, 150, 1200, 250, 10
'    SpDoc.LoadImage App.Path & "\iconos\ARROZ.bmp", id
'    SpDoc.PlaceImage id, 1700, 2000, 400, 1100, 10
           nRows = 50
    'draw the page title in black arial 24pt bold underline,
    'centred and starting 25mm down from the page top
    'spDoc.SetTextColor vbBlack
    SpDoc.ForeColor = vbBlack
    'spDoc.SetFont "Arial", 240, PPF_BOLD + PPF_UNDERLINE, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.SetFont "Arial", 150, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 50, nRows, "BENEFICIO DE GRANOS MATURAVE S.A." '& nPage
    
        ''''''''''''''''''CAI'''''''''''''''''''''''''''
    nRows = nRows + 10
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 2000, nRows, "CAI: D839C1-8C314F-43448B-0DE48B-7C8C31-2B" '& nPage
    nRows = nRows + 120
        ''''''''''''''''''''''''FACTURA
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.TextOut 2000, nRows, " Factura No:  000-002-01-" & DIM_NODOC
       ''''''''''''''''''''''''FECHA
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD + 0, 0
    nRows = nRows + 80
    SpDoc.TextOut 2000, nRows, " FECHA : " & PUB_FECHA
    ''''direccion local propio''''''''''''''''''''''''''''''
       SpDoc.SetFont "Arial", 80, SPFS_POINTS + 0, 0
   nRows = 180
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Francisco Morazan,D.C." '& nPage
  nRows = nRows + 40

    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "Kilometro 9 Carretera Olancho" '& nPage
  nRows = nRows + 40
    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "RTN: 08019007068256" '& nPage
  nRows = nRows + 40
    ''''''''''EMAIL''''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "E-Mail: maturave9@yahoo.com" '& nPage
  nRows = nRows + 40
    '''''''''''''''TELEFONO''''''''''''''''''''''
      SpDoc.TextAlign = SPTA_RIGHT
    '''''''''''''''TELEFONO''''''''''''''''''''''
    SpDoc.TextOut 2000, nRows, "Telefono: 504 22916125"  '& nPage
        SpDoc.TextAlign = SPTA_LEFT


End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If DIM_TIENDA = "CONCENTRADO" Then
'Set pic = LoadPicture(App.Path & "\iconos\ARROZ.bmp")
    nCentre = SpDoc.PageWidth / 2
    SpDoc.LoadImage App.Path & "\iconos\FACOCA.bmp", ID
SpDoc.PlaceImage ID, 750, 150, 1200, 250, 10
'    SpDoc.LoadImage App.Path & "\iconos\ARROZ.bmp", id

'    SpDoc.PlaceImage id,1200, 150, 700, 250, 10
           nRows = 50
    'draw the page title in black arial 24pt bold underline,
    'centred and starting 25mm down from the page top
    'spDoc.SetTextColor vbBlack
    SpDoc.ForeColor = vbBlack
    'spDoc.SetFont "Arial", 240, PPF_BOLD + PPF_UNDERLINE, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.SetFont "Arial", 150, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 50, nRows, "FABRICA DE CONCENTRADOS CARMEN S.A." '& nPage
    
        ''''''''''''''''''CAI'''''''''''''''''''''''''''
    nRows = nRows + 10
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 2000, nRows, "CAI: DD66C7-C51456-6A45AB-3A3566-534BBC-E0" '& nPage
    nRows = nRows + 120
        ''''''''''''''''''''''''FACTURA
    SpDoc.SetFont "Arial", 120, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.TextOut 2000, nRows, " Factura No:  000-002-01-" & DIM_NODOC
       ''''''''''''''''''''''''FECHA
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD + 0, 0
    nRows = nRows + 80
    SpDoc.TextOut 2000, nRows, " FECHA : " & PUB_FECHA
    ''''direccion local propio''''''''''''''''''''''''''''''
       SpDoc.SetFont "Arial", 80, SPFS_POINTS + 0, 0
   nRows = 180
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Francisco Morazan,D.C." '& nPage
  nRows = nRows + 40

    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "Kilometro 9 Carretera Olancho" '& nPage
  nRows = nRows + 40
    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "RTN: 08019007068278" '& nPage
  nRows = nRows + 40
    ''''''''''EMAIL''''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "E-Mail: maturave9@yahoo.com" '& nPage
  nRows = nRows + 40
    '''''''''''''''TELEFONO''''''''''''''''''''''
       SpDoc.TextAlign = SPTA_RIGHT
    '''''''''''''''TELEFONO''''''''''''''''''''''
    SpDoc.TextOut 2000, nRows, "Telefono: 504 22916125"  '& nPage
        SpDoc.TextAlign = SPTA_LEFT

End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If DIM_TIENDA = "HUEVO" Then

'Set pic = LoadPicture(App.Path & "\iconos\ARROZ.bmp")
    nCentre = SpDoc.PageWidth / 2
    SpDoc.LoadImage App.Path & "\iconos\GRAVASI.bmp", ID
SpDoc.PlaceImage ID, 1200, 150, 800, 250, 10
'    SpDoc.LoadImage App.Path & "\iconos\ARROZ.bmp", id
'    SpDoc.PlaceImage id, 1700, 2000, 400, 1100, 10
           nRows = 50
    'draw the page title in black arial 24pt bold underline,
    'centred and starting 25mm down from the page top
    'spDoc.SetTextColor vbBlack
    SpDoc.ForeColor = vbBlack
    'spDoc.SetFont "Arial", 240, PPF_BOLD + PPF_UNDERLINE, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.SetFont "Arial", 150, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 50, nRows, "GRAVASI S.A." '& nPage
    
        ''''''''''''''''''CAI'''''''''''''''''''''''''''
    nRows = nRows + 10
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 2000, nRows, "CAI: 9A9B1E-79673D-5B4C82-6EB5E2-EEC05D-7B" '& nPage
    nRows = nRows + 120
        ''''''''''''''''''''''''FACTURA
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.TextOut 2000, nRows, " Factura No:  000-002-01-" & DIM_NODOC
       ''''''''''''''''''''''''FECHA
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD + 0, 0
    nRows = nRows + 80
    SpDoc.TextOut 2000, nRows, " FECHA : " & PUB_FECHA
    ''''direccion local propio''''''''''''''''''''''''''''''
       SpDoc.SetFont "Arial", 80, SPFS_POINTS + 0, 0
   nRows = 180
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Francisco Morazan,D.C." '& nPage
  nRows = nRows + 40

    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "Kilometro 9 Carretera Olancho" '& nPage
  nRows = nRows + 40
    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "RTN: 08039008133334" '& nPage
  nRows = nRows + 40
    ''''''''''EMAIL''''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "E-Mail: maturave9@yahoo.com" '& nPage
  nRows = nRows + 40
    '''''''''''''''TELEFONO''''''''''''''''''''''
      SpDoc.TextAlign = SPTA_RIGHT
    '''''''''''''''TELEFONO''''''''''''''''''''''
    SpDoc.TextOut 2000, nRows, "Telefono: 504 22916125"  '& nPage
        SpDoc.TextAlign = SPTA_LEFT

End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If DIM_TIENDA = "MATURAVE" Then


'Set pic = LoadPicture(App.Path & "\iconos\ARROZ.bmp")
    nCentre = SpDoc.PageWidth / 2
    SpDoc.LoadImage App.Path & "\iconos\MATURAVE.bmp", ID
SpDoc.PlaceImage ID, 1200, 150, 800, 250, 10
'    SpDoc.LoadImage App.Path & "\iconos\ARROZ.bmp", id
'    SpDoc.PlaceImage id, 1700, 2000, 400, 1100, 10
           nRows = 50
    'draw the page title in black arial 24pt bold underline,
    'centred and starting 25mm down from the page top
    'spDoc.SetTextColor vbBlack
    SpDoc.ForeColor = vbBlack
    'spDoc.SetFont "Arial", 240, PPF_BOLD + PPF_UNDERLINE, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.SetFont "Arial", 150, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 50, nRows, "RAMOS VELAZQUEZ MARCO TULIO" '& nPage
    
        ''''''''''''''''''CAI'''''''''''''''''''''''''''
    nRows = nRows + 10
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 2000, nRows, "CAI: 3688A1-ADC612-F845B8-F9CA6B-F8BCA2-EF" '& nPage
    nRows = nRows + 120
        ''''''''''''''''''''''''FACTURA
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.TextOut 2000, nRows, " Factura No:  000-002-01-" & DIM_NODOC
       ''''''''''''''''''''''''FECHA
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD + 0, 0
    nRows = nRows + 80
    SpDoc.TextOut 2000, nRows, " FECHA : " & PUB_FECHA
    ''''direccion local propio''''''''''''''''''''''''''''''
       SpDoc.SetFont "Arial", 80, SPFS_POINTS + 0, 0
   nRows = 180
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Francisco Morazan,D.C." '& nPage
  nRows = nRows + 40

    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "Kilometro 9 Carretera Olancho" '& nPage
  nRows = nRows + 40
    ''''''''''RTN''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "RTN: 08051983000924" '& nPage
  nRows = nRows + 40
    ''''''''''EMAIL''''''''''''''''''''''''''''''
    SpDoc.TextOut 50, nRows, "E-Mail: maturave9@yahoo.com" '& nPage
  nRows = nRows + 40
    '''''''''''''''TELEFONO''''''''''''''''''''''
      SpDoc.TextAlign = SPTA_RIGHT
    '''''''''''''''TELEFONO''''''''''''''''''''''
    SpDoc.TextOut 2000, nRows, "Telefono: 504 22916125"  '& nPage
        SpDoc.TextAlign = SPTA_LEFT

End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
Set RS_TOTAL = New Recordset
DIM_SQL = "Select * From Clientes where Codigo like '%" & DIM_CODCLIENTE & "%'"
RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
With RS_TOTAL
    If .EOF Then
        '.MoveFirst
        Exit Sub
    End If
    If IsEmpty(DIM_CODCLIENTE) Then
    DIM_CODCLIENTE = ""
    DIM_CONTACTO = ""
    PUB_CLIENTE = ""
    DIM_CIUDAD = ""
    DIM_DIRECCIONCL = ""
  SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOL, 0
    nRows = nRows + 50
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "CLIENTE :       Anonimo  "
'    SpDoc.FillSolidRect 2050, 1000, 80, 1050, Gray
'    SpDoc.FillSolidRect 2050, 2120, 80, 2300, vbWhite
      
''''''''''''''''''''''''CLIENTE
 
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 1000, nRows, "RTN CLIENTE :   0   "
    Else
    DIM_CODCLIENTE = .Fields("codigo")
    DIM_CONTACTO = .Fields("contacto")
    DIM_CIUDAD = .Fields("ciudad")
    DIM_DIRECCIONCL = .Fields("direccion")
        PUB_CLIENTE = .Fields("NOMBRE")
    
    nRows = nRows + 50
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "CLIENTE : " & PUB_CLIENTE
      
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 1000, nRows, "RTN CLIENTE :  " & DIM_CODCLIENTE

    nRows = nRows + 25
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Direccion : " & DIM_DIRECCIONCL
    
    End If
End With
Set RS_TOTAL = Nothing
    SpDoc.TextAlign = SPTA_LEFT

    SpDoc.FillSolidRect 2050, 450, 50, 455, vbBlack
 
''''''''''''''''''''''''CLIENTE
    SpDoc.SetFont "Arial", 100, SPFS_POINTS + 0, 0
    nRows = nRows + 60
    SpDoc.TextAlign = SPTA_LEFT
   ' SpDoc.ForeColor = vbWhite 'RGB(255, 255, 255)
    SpDoc.TextOut 50, nRows, "  Codigo             Producto                        Isv          Cantidad            PrecioU            Forma           Precio Total"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   nRows = nRows + 20
    SpDoc.FillSolidRect 2050, 510, 50, 520, vbBlack
         nRows = nRows + 75
   ' SpDoc.FillSolidRect 2050, 1300, 200, 1400, vbWhite
   
       SpDoc.SetFont "Arial", 70, SPFS_POINTS + 0, 0
       
    SpDoc.SetFont "Arial", 90, SPFS_POINTS + 0, 0
    SpDoc.ForeColor = vbBlack
SpDoc.BackMode = SPBM_TRANSPARENT


Set RS_TOTAL = New Recordset
' "Select SUM(TOTAL),SUM(ISV),SUM(DESCUENTO),SUM(valor) from INVSALIDA where ndventas  =  " & DIM_NODOCC, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'DIM_SUMTOTAL = RS_TOTAL.Fields(0)
RS_TOTAL.Open "SELECT Codigo,Producto,ISV,Salida,punitario,FORMA,TOTAL  FROM INVSalida where NDVentas like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Do While Not RS_TOTAL.EOF
        Printer.CurrentX = LEFT_MARGIN
        SpDoc.SetFont "Arial", 100, SPFS_POINTS + 0, 0
            
         SpDoc.TextOut 110, nRows, RS_TOTAL!codigo
         SpDoc.TextOut 300, nRows, RS_TOTAL!Producto
         SpDoc.TextOut 750, nRows, RS_TOTAL!Isv

         SpDoc.TextOut 900, nRows, RS_TOTAL!salida
         SpDoc.TextOut 1150, nRows, RS_TOTAL!PUNITARIO
        SpDoc.TextOut 1400, nRows, RS_TOTAL!FORMA
         If RS_TOTAL!Isv = 0 Or RS_TOTAL!Isv = "" Then
        SpDoc.SetFont "Arial", 110, SPFS_POINTS + 0, 0
         SpDoc.TextOut 1600, nRows, "...Lps." & Format(RS_TOTAL!total, "#,##0.00")
        SpDoc.SetFont "Arial", 100, SPFS_POINTS + 0, 0
        Else
        SpDoc.SetFont "Arial", 110, SPFS_POINTS + 0, 0
        DIMVI = RS_TOTAL!Isv
        DIMVV = RS_TOTAL!total
        DIMVS = DIMVV - DIMVI
         SpDoc.TextOut 1600, nRows, "...Lps." & Format(DIMVS, "#,##0.00")
        SpDoc.SetFont "Arial", 100, SPFS_POINTS + 0, 0
        End If
         nRows = nRows + 40
          SpDoc.FillSolidRect 2050, nRows, 50, nRows + 5, vbBlack
        nRows = nRows + 35
        RS_TOTAL.MoveNext
        NumLineas = NumLineas + 1
Loop

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''

Set RS_TOTAL = Nothing



Set RS_TOTAL = New Recordset
' "Select SUM(TOTAL),SUM(ISV),SUM(DESCUENTO),SUM(valor) from INVSALIDA where ndventas  =  " & DIM_NODOCC, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'DIM_SUMTOTAL = RS_TOTAL.Fields(0)
RS_TOTAL.Open "SELECT sum(total),sum(isv)  FROM INVSalida where ndventas  like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

If NumLineas = 1 Then
'nRows = nRows + 200
'SpDoc.FillSolidRect 2050, nRows, 80, nRows + 8, vbBlack
End If
If NumLineas = 2 Then
nRows = nRows + 200
SpDoc.FillSolidRect 2050, nRows, 50, nRows + 8, vbBlack
End If
If NumLineas = 3 Then
nRows = nRows + 80
SpDoc.FillSolidRect 2050, nRows, 50, nRows + 8, vbBlack
End If
If NumLineas = 4 Then
nRows = nRows + 80
SpDoc.FillSolidRect 2050, nRows, 50, nRows + 8, vbBlack
End If
If NumLineas = 5 Then
nRows = nRows + 50
SpDoc.FillSolidRect 2050, nRows, 50, nRows + 8, vbBlack
End If
If NumLineas = 6 Then
nRows = nRows + 50
SpDoc.FillSolidRect 2050, nRows, 50, nRows + 8, vbBlack
End If
If NumLineas = 7 Then
nRows = nRows + 50
SpDoc.FillSolidRect 2050, nRows, 50, nRows + 8, vbBlack
End If
DL1 = nRows
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SpDoc.FillSolidRect 2050, 450, 2060, DL1, vbBlack
SpDoc.FillSolidRect 50, 450, 55, DL1, vbBlack
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''TOTAL
    SpDoc.ForeColor = vbBlack
    DIM_z = nRows
    nRows = nRows + 30

If DIM_TIENDA = "MATURAVE" Then
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Fecha limite de Emision : 05 de Febrero del 2016 "
End If
If DIM_TIENDA = "ARROZ" Then
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Fecha limite de Emision : 05 de Febrero del 2016 "
End If
If DIM_TIENDA = "MAIZ" Then
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Fecha limite de Emision : 28 de Marzo del 2016 "
End If
If DIM_TIENDA = "FACOCA" Then
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Fecha limite de Emision : 05 de Noviembre del 2015 "
End If
If DIM_TIENDA = "GRAVASI" Then
    SpDoc.SetFont "Arial", 80, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 50, nRows, "Fecha limite de Emision : 05 de Febrero del 2016 "
End If






    SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.TextAlign = SPTA_RIGHT
    SpDoc.SetFont "Arial", 100, SPFS_POINTS, 0
    'spDoc.SetTextAlign PPA_NOUPDATECP + PPA_CENTER + PPA_TOP
    'SpDoc.TextAlign = SPTA_NOUPDATECP + SPTA_CENTER + SPTA_TOP
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.TextOut 1500, nRows, "IMPORTE ISV "
    SpDoc.TextAlign = SPTA_RIGHT
        SpDoc.SetFont "Arial", 110, SPFS_POINTS + SPFO_BOLD, 0
    If IsNull(RS_TOTAL.Fields(1)) Or RS_TOTAL.Fields(1) = 0 Then
    DIM3 = Format(RS_TOTAL.Fields(0), "#,##0.00")
    Else
    DIM3 = RS_TOTAL.Fields(0) - RS_TOTAL.Fields(1)
    End If
    SpDoc.TextOut 2000, nRows, Format(DIM3, "#,##0.00")
    nRows = nRows + 50
    SpDoc.TextAlign = SPTA_LEFT
    SpDoc.FillSolidRect 2050, nRows, 1300, nRows + 5, vbBlack
    DL3 = nRows - 80
        nRows = nRows + 50
            SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.TextOut 1500, nRows, "ISV 15 %"
    SpDoc.TextAlign = SPTA_RIGHT
        SpDoc.SetFont "Arial", 110, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.TextOut 2000, nRows, Format(RS_TOTAL.Fields(1), "#,##0.00")
    nRows = nRows + 50
    SpDoc.TextAlign = SPTA_LEFT
        SpDoc.SetFont "Arial", 100, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.FillSolidRect 2050, nRows, 1300, nRows + 5, vbBlack
        nRows = nRows + 50
    SpDoc.TextOut 1500, nRows, "TOTAL A PAGAR  "
    SpDoc.TextAlign = SPTA_RIGHT
        SpDoc.SetFont "Arial", 120, SPFS_POINTS + SPFO_BOLD, 0
    'DIM3 = RS_TOTAL.Fields(0) + RS_TOTAL.Fields(1)
    SpDoc.TextOut 2000, nRows, Format(RS_TOTAL.Fields(0), "#,##0.00")
    
    



    SpDoc.TextAlign = SPTA_LEFT
    Label9.Caption = DIM3
        SpDoc.SetFont "Arial", 120, SPFS_POINTS + SPFO_BOLD, 0
    '    SpDoc.TextOut 50, DIM_z + 75, NroEnLetras(Label9.Caption)
        

        
    SpDoc.SetFont "Arial", 150, SPFS_POINTS + SPFO_BOLD + 0, 0
    SpDoc.TextOut 50, DIM_z + 150, "***" & DIMFORMA777 & "***"
    
               nRows = nRows + 25
    SpDoc.SetFont "Arial", 120, SPFS_POINTS, 0
    SpDoc.TextOut 50, nRows, "_______________________"
    SpDoc.TextOut 600, nRows, "_______________________"
        nRows = nRows + 50
    SpDoc.TextOut 50, nRows, "     Firma Cliente"
    SpDoc.TextOut 600, nRows, "     Entregada por"
    
    
    SpDoc.TextAlign = SPTA_RIGHT
    nRows = nRows + 50
        SpDoc.SetFont "Arial", 75, SPFS_POINTS + SPFO_BOLD, 0
    SpDoc.FillSolidRect 2050, nRows, 1300, nRows + 9, vbBlack
        SpDoc.FillSolidRect 2050, DL3, 2060, nRows, vbBlack
    SpDoc.FillSolidRect 1300, DL3, 1310, nRows, vbBlack
    SpDoc.TextAlign = SPTA_LEFT

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''

Set RS_TOTAL = Nothing

''''''''''''''''''''''''CLIENTE

    
''''''''''''''''''''''''CLIENTE

    
    
        SpDoc.SetFont "Arial", 150, SPFS_POINTS + SPFO_BOLD, 0
    nRows = nRows + 70
    SpDoc.TextAlign = SPTA_CENTER
    SpDoc.TextOut 1100, nRows, "LA FACTURA ES BENEFICIOS DE TODOS EXIJALA "
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '    SpDoc.SetFont "Arial", 100, SPFS_POINTS, 0
    'rgLines(0, 0) = SPPN_SOLID:      rgLines(1, 0) = "Solid"
    'rgLines(0, 1) = SPPN_DASH:       rgLines(1, 1) = "Dash"
    'rgLines(0, 2) = SPPN_DOT:        rgLines(1, 2) = "Dot"
    'rgLines(0, 3) = SPPN_DASHDOT:    rgLines(1, 3) = "DashDot"
    'rgLines(0, 4) = SPPN_DASHDOTDOT: rgLines(1, 4) = "DashDotDot"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                DrawPicture strPath & "\skins\logomobile300.jpg", SPIA_STRETCH
        SpDoc.DoPrintPreview




'RS_CLIENTEINFO.Close
DIM_DIRECCION = ""
DIM_NODOCB = ""
PUB_USUARIO = ""
PUB_VENDEDOR = ""
PUB_CLIENTE = ""
DIM_DIRECCIONCL = ""
DIM_FECHAPAGO = ""
PUB_CLIENTE = ""
DIM_CODCLIENTE = ""
                DIM_DIRECCIONCL = ""
                DIM_CODCLIENTE = ""
                DIM_CONTACTO = ""
                DIM_CIUDAD = ""
                DIM_DIRECCIONCL = ""
                PUB_CLIENTE = ""
                PUB_FECHA = ""
DIMFORMA777 = ""
End Sub
Private Sub BTN8_Click()
'On Error GoTo menerr
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim DOCELIMINAR

'Busca la impresora de PDF y ponla como
'predetermianda dentro de esta aplicacion.
Dim oPrn As Printer

'Busca en todas las imrpesoras.
For Each oPrn In Printers
'Busca el generador de PDF.
If oPrn.DeviceName = "BIXOLON SRP-275" Then
'Se encontro, pon esta impresora como predeterminada
'y sal del FOR Loop.
Set Printer = oPrn
Exit For
End If
Next

'Aqui la aplicacion usara la impresora preterminada.
'Al salirte de la aplicacion WIndows sigue usando la
'impresora predeterminada desde siempre.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_ELIMINADO As ADODB.Recordset


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
DIM_A = MsgBox("DESEA ELIMINAR TODA LA TRANSACCION", vbYesNo)
    If DIM_A = vbYes Then
    
            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''
                        Borrar_NoDoc = LST_INVT.SelectedItem.ListSubItems.Item(9).Text
                        Printer.CurrentY = TOP_MARGIN
                        Printer.CurrentX = LEFT_MARGIN
                        Printer.Font.Size = 9
                        Printer.FontName = "Arial"
                        Printer.Font.Bold = True
                        Printer.Font.Size = 10

                        'Printer.FontName = "FontControl"
                        For i = 1 To 1   ' Set up two iterations.
                         DIM_TITULO = "*** FACTURA ELIMINADA ***"
                         HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
                         Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
                         Printer.Print DIM_TITULO  ' Send new page.
                        Next i
                        Printer.Font.Size = 9
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Printer.Print "_______________________________"
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        For i = 1 To 1   ' Set up two iterations.
                         DIM_TITULO = "No Factura = " & RS_VENTAS.Fields("NDVentas")
                         HWidth = Printer.TextWidth(DIM_TITULO) / 2   ' Get one-half width.
                         Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
                         Printer.Print DIM_TITULO  ' Send new page.
                        Next i
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Printer.Print "_______________________________ "
                        Printer.Font.Bold = False
                        Printer.Font.Size = 9
                         '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
                         DOCELIMINAR = PUB_10
                        Set RS_ELIMINADO = New Recordset
                        PUB_SQL = "select Codigo,Producto,Salida,Total,forma  FROM INVSalida WHERE ndventas like '" & PUB_10 & "' "
                        RS_ELIMINADO.Open PUB_SQL, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
                        Printer.Print
                        With RS_ELIMINADO
                        Printer.Print !FORMA
                         Printer.Print "________________________________"
                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    Do While Not .EOF
                                            ' Use rs!FieldName to get the data for
                                            ' the field named FieldName.
                                            Printer.CurrentX = LEFT_MARGIN
                                            Printer.Print "Codigo:  " & !codigo
                                            Printer.Print "Producto: " & !Producto
                                            Printer.Print "Cantidad: " & !salida
                                            Printer.Print "Total: " & !total
                                            Printer.Print "________________________________"
                                            'Format$(rs!Titulo) & vbTab & Format$(rs!Formato) & vbTab & Format$(rs!FormatoCompresion) & vbTab & (rs!MinCDs) & vbTab & (rs!NumDVDs) & vbTab & Format$(rs!NumCDs) & vbTab & Format$(rs!Genero) & vbTab & Format$(rs!Subtitulos) & vbTab & Format$(rs!Idioma)
                                           ' See if we have filled the page.
                                            If Printer.CurrentY >= bottom_margin Then
                                                ' Start a new page.
                                                Printer.NewPage
                                                Printer.CurrentY = TOP_MARGIN
                                            End If
                                            
                                            .MoveNext
                                    Loop
                                    

                        End With
                        RS_ELIMINADO.Close
                        Set RS_ELIMINADO = Nothing
                        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''

                                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                Printer.FontName = "Control"
                                                
                                                Printer.EndDoc
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
            
                    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
                    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
                    Set RS_ELIMINADO = New Recordset
                    PUB_SQL = "DELETE * FROM INVSalida WHERE ndventas like '" & PUB_10 & "' "
                    RS_ELIMINADO.Open PUB_SQL, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
            
                    Set RS_ELIMINADO = Nothing
                    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
                    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
                    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''

                    Dim RS_ELIMINAR As ADODB.Recordset
                    Set RS_ELIMINAR = New Recordset
                    
                    PUB_SQL = "DELETE * FROM INVSalida1  where NDVENTAS like '" & PUB_10 & "'"
                    'DIM_SQLITEM = DIM_SQLITEM & " AND HORA1 like '" & Borrar_Hora & "'"
                    RS_ELIMINAR.Open PUB_SQL, PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
                    Set RS_ELIMINAR = Nothing
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
                    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''
                    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            
            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            With RS_SALIDA
            .AddNew
            .Fields("codigo") = "ELIMINADO"
            .Fields("producto") = "ELIMINADO"
            '.Fields("vendedor") = "ELIMINADO"
            .Fields("total") = "0"
            '.Fields("valor") = "0"
            '.Fields("NoDE") = "0"
            .Fields("NDVentas") = PUB_10
            '.Fields("TIENDA") = gsRutaBaseDatos

            .Fields("Fecha") = Date
            '.Fields("Forma") = "CONTADO"
            '.Fields("Hora1") = Format(Time, "Long Time")
            '.Fields("Descripcion") = "FACTURA ELIMINADA"
            
            .Update
            
                            
            
            Set RS_SALIDA = New Recordset
            DOCELIMINAR = ""
            RS_SALIDA.Open "SELECT Codigo,Producto,Salida,punitario,ndventas,total,ncliente,cliente,fecha,Hora1,Descripcion,NDVentas,forma,Vendedor  FROM INVSalida", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
            RS_SALIDA.MoveFirst
            RS_SALIDA.MoveLast
            Carga_lstvDatos
            End With



Else
End If


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''

End Sub
'Dim ban As Integer
Private Sub BTN1_Click(Index As Integer)
Select Case Index
    Case 0
    
       RS_VENTAS.MoveFirst
       refrescar
       
       BTN1(0).Enabled = False
       BTN1(1).Enabled = False
       BTN1(2).Enabled = True
       BTN1(3).Enabled = True
    Case 1
    
       RS_VENTAS.MovePrevious
       
       BTN1(2).Enabled = True
       BTN1(3).Enabled = True
       If RS_VENTAS.BOF = True Then
        RS_VENTAS.MoveFirst
        refrescar
        BTN1(0).Enabled = False
        BTN1(1).Enabled = False
       Else
         refrescar
       End If
    Case 2
   
        RS_VENTAS.MoveNext
    
       BTN1(0).Enabled = True
       BTN1(1).Enabled = True
       If RS_VENTAS.EOF = True Then
         BTN1(2).Enabled = False
         BTN1(3).Enabled = False
         RS_VENTAS.MoveLast
         refrescar
       Else
        refrescar
       End If
    Case 3
  
       RS_VENTAS.MoveLast
       
       BTN1(0).Enabled = True
       BTN1(1).Enabled = True
       BTN1(2).Enabled = False
       BTN1(3).Enabled = False
       refrescar
End Select


'DIM_INT_7 = False
End Sub

Private Sub Command5_Click()
If IsEmpty(PUB_CONEXION_EASY) Then
Set pic = Nothing
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "ARROZ"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\ARROZ.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE ARROZ"
Else
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "ARROZ"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\ARROZ.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing
'MsgBox pic
PubNegocio = "MATURAVE ARROZ"

End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Set RS_CUENTAS_INGRESOS = New Recordset

RS_CUENTAS_INGRESOS.Open "Select NoDoc4 from Ventas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

If RS_CUENTAS_INGRESOS.EOF = True Or RS_CUENTAS_INGRESOS.BOF = True Then
Else
lbl_total = Format(RS_CUENTAS_INGRESOS.Fields(0), "#,##0.00")
End If
   With LV
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NoDoc4", 5000

        
    End With
  With RS_CUENTAS_INGRESOS
        If RS_CUENTAS_INGRESOS.BOF = True And RS_CUENTAS_INGRESOS.EOF = True Then
        LV.ListItems.Clear
        Else
            LV.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LV.ListItems.Add(, , .Fields(0) & "")

                .MoveNext
            Loop
        End If
    End With
RS_CUENTAS_INGRESOS.Close
Set RS_CUENTAS_INGRESOS = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''

'*****************************************************************************************************************
'*****************************************************************************************************************
 LST_INVT.Width = Screen.Width - 500
  LST_INVT.Height = Screen.Height - 2300
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from InvSalida Order by NDVentas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'MsgBox PUB_CONEXION_EASY
lstvDatos_a_cero

refrescar

End Sub

Private Sub Command2_Click()
If IsEmpty(PUB_CONEXION_EASY) Then
Set pic = Nothing
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "MAIZ"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_MAIZ & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\MATURAVE1.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE MAIZ"
Else
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "MAIZ"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_MAIZ & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\MATURAVE1.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE MAIZ"

End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Set RS_CUENTAS_INGRESOS = New Recordset

RS_CUENTAS_INGRESOS.Open "Select NoDoc4 from Ventas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

If RS_CUENTAS_INGRESOS.EOF = True Or RS_CUENTAS_INGRESOS.BOF = True Then
Else
lbl_total = Format(RS_CUENTAS_INGRESOS.Fields(0), "#,##0.00")
End If
   With LV
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NoDoc4", 5000

        
    End With
  With RS_CUENTAS_INGRESOS
        If RS_CUENTAS_INGRESOS.BOF = True And RS_CUENTAS_INGRESOS.EOF = True Then
        LV.ListItems.Clear
        Else
            LV.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LV.ListItems.Add(, , .Fields(0) & "")

                .MoveNext
            Loop
        End If
    End With
RS_CUENTAS_INGRESOS.Close
Set RS_CUENTAS_INGRESOS = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''

'*****************************************************************************************************************
'*****************************************************************************************************************
 LST_INVT.Width = Screen.Width - 500
  LST_INVT.Height = Screen.Height - 2300
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from InvSalida Order by NDVentas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

lstvDatos_a_cero

refrescar

End Sub
Private Sub Command3_Click()
If IsEmpty(PUB_CONEXION_EASY) Then
Set pic = Nothing
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "CONCENTRADO"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_CONCENTRADO & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\FACOCA.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "FACOCA"
Else
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "CONCENTRADO"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_CONCENTRADO & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\FACOCA.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "FACOCA"

End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Set RS_CUENTAS_INGRESOS = New Recordset

RS_CUENTAS_INGRESOS.Open "Select NoDoc4 from Ventas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

If RS_CUENTAS_INGRESOS.EOF = True Or RS_CUENTAS_INGRESOS.BOF = True Then
Else
lbl_total = Format(RS_CUENTAS_INGRESOS.Fields(0), "#,##0.00")
End If
   With LV
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NoDoc4", 5000

        
    End With
  With RS_CUENTAS_INGRESOS
        If RS_CUENTAS_INGRESOS.BOF = True And RS_CUENTAS_INGRESOS.EOF = True Then
        LV.ListItems.Clear
        Else
            LV.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LV.ListItems.Add(, , .Fields(0) & "")

                .MoveNext
            Loop
        End If
    End With
RS_CUENTAS_INGRESOS.Close
Set RS_CUENTAS_INGRESOS = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''

'*****************************************************************************************************************
'*****************************************************************************************************************
 LST_INVT.Width = Screen.Width - 500
  LST_INVT.Height = Screen.Height - 2300
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from InvSalida Order by NDVentas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

lstvDatos_a_cero

refrescar

End Sub
Private Sub Command4_Click()
If IsEmpty(PUB_CONEXION_EASY) Then
Set pic = Nothing
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "HUEVO"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HUEVO & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\GRAVASI.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE HUEVO"
Else
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "HUEVO"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HUEVO & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\GRAVASI.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE HUEVO"

End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Set RS_CUENTAS_INGRESOS = New Recordset

RS_CUENTAS_INGRESOS.Open "Select NoDoc4 from Ventas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

If RS_CUENTAS_INGRESOS.EOF = True Or RS_CUENTAS_INGRESOS.BOF = True Then
Else
lbl_total = Format(RS_CUENTAS_INGRESOS.Fields(0), "#,##0.00")
End If
   With LV
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NoDoc4", 5000

        
    End With
  With RS_CUENTAS_INGRESOS
        If RS_CUENTAS_INGRESOS.BOF = True And RS_CUENTAS_INGRESOS.EOF = True Then
        LV.ListItems.Clear
        Else
            LV.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LV.ListItems.Add(, , .Fields(0) & "")

                .MoveNext
            Loop
        End If
    End With
RS_CUENTAS_INGRESOS.Close
Set RS_CUENTAS_INGRESOS = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''

'*****************************************************************************************************************
 LST_INVT.Width = Screen.Width - 500
  LST_INVT.Height = Screen.Height - 2300
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from InvSalida Order by NDVentas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

lstvDatos_a_cero

refrescar

End Sub
Private Sub Command6_Click()
If IsEmpty(PUB_CONEXION_EASY) Then
Set pic = Nothing
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "HARINA"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HARINA & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\MATURAVE.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE HUEVO"
Else
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "HARINA"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HARINA & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\MATURAVE.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE HARINA"

End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Set RS_CUENTAS_INGRESOS = New Recordset

RS_CUENTAS_INGRESOS.Open "Select NoDoc4 from Ventas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

If RS_CUENTAS_INGRESOS.EOF = True Or RS_CUENTAS_INGRESOS.BOF = True Then
Else
lbl_total = Format(RS_CUENTAS_INGRESOS.Fields(0), "#,##0.00")
End If
   With LV
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NoDoc4", 5000

        
    End With
  With RS_CUENTAS_INGRESOS
        If RS_CUENTAS_INGRESOS.BOF = True And RS_CUENTAS_INGRESOS.EOF = True Then
        LV.ListItems.Clear
        Else
            LV.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LV.ListItems.Add(, , .Fields(0) & "")

                .MoveNext
            Loop
        End If
    End With
RS_CUENTAS_INGRESOS.Close
Set RS_CUENTAS_INGRESOS = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''

'*****************************************************************************************************************

'*****************************************************************************************************************
 LST_INVT.Width = Screen.Width - 500
  LST_INVT.Height = Screen.Height - 2300
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from InvSalida Order by NDVentas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

lstvDatos_a_cero

refrescar

End Sub
Private Sub Command7_Click()
If IsEmpty(PUB_CONEXION_EASY) Then
Set pic = Nothing
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "MATURAVE"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_MATURAVE & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\MATURAVE.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE HUEVO"
Else
PUB_CONEXION_EASY.Close


gsRutaBaseDatos = "MATURAVE"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_MATURAVE & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
PUB_CONEXION_EASY.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\MATURAVE.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing
'MsgBox pic
PubNegocio = "MATURAVE "

End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Set RS_CUENTAS_INGRESOS = New Recordset

RS_CUENTAS_INGRESOS.Open "Select NoDoc4 from Ventas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

If RS_CUENTAS_INGRESOS.EOF = True Or RS_CUENTAS_INGRESOS.BOF = True Then
Else
lbl_total = Format(RS_CUENTAS_INGRESOS.Fields(0), "#,##0.00")
End If
   With LV
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NoDoc4", 5000

        
    End With
  With RS_CUENTAS_INGRESOS
        If RS_CUENTAS_INGRESOS.BOF = True And RS_CUENTAS_INGRESOS.EOF = True Then
        LV.ListItems.Clear
        Else
            LV.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LV.ListItems.Add(, , .Fields(0) & "")

                .MoveNext
            Loop
        End If
    End With
RS_CUENTAS_INGRESOS.Close
Set RS_CUENTAS_INGRESOS = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''

'*****************************************************************************************************************

'*****************************************************************************************************************
 LST_INVT.Width = Screen.Width - 500
  LST_INVT.Height = Screen.Height - 2300
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from InvSalida Order by NDVentas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'RS_VENTAS.Open "select * from InvSalida Order by NDVentas", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

lstvDatos_a_cero

refrescar

End Sub
