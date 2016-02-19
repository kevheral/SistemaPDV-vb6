VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Inventario 
   BackColor       =   &H00FFFFFF&
   Caption         =   "AGREGAR CLIENTES"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox text6 
      Height          =   315
      ItemData        =   "frminventario.frx":0000
      Left            =   2040
      List            =   "frminventario.frx":0002
      TabIndex        =   7
      Text            =   "tegucigalpa"
      Top             =   3960
      Width           =   3495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5535
      Left            =   8640
      TabIndex        =   10
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9763
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton BTN9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   6495
      Picture         =   "frminventario.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "SALIR"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton BTN8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   4065
      Picture         =   "frminventario.frx":441B
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "ELIMINAR"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton BTN7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   5280
      Picture         =   "frminventario.frx":856A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "MODIFICAR"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton BTN6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   1635
      Picture         =   "frminventario.frx":C792
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "GUARDAR"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton BTN5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   360
      Picture         =   "frminventario.frx":10A76
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "NUEVO"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   2880
      Picture         =   "frminventario.frx":14D8A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "ACTUALIZAR"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   280
      Left            =   2040
      TabIndex        =   2
      Top             =   1860
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   280
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   280
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   3500
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   280
      Left            =   2040
      TabIndex        =   6
      Top             =   3540
      Width           =   3500
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   280
      Left            =   2040
      TabIndex        =   9
      Top             =   4800
      Width           =   6255
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   280
      Left            =   2040
      TabIndex        =   8
      Top             =   4380
      Width           =   3500
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   280
      Left            =   2040
      TabIndex        =   4
      Top             =   2700
      Width           =   3500
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   280
      Left            =   2040
      TabIndex        =   5
      Top             =   3120
      Width           =   3500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   24
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   23
      Top             =   1860
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   22
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   21
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Movil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   20
      Top             =   3540
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   19
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   18
      Top             =   4800
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   17
      Top             =   4380
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono Contaco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "Inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim RS_PRODUCTO As ADODB.Recordset
Dim RS_LV As ADODB.Recordset
Dim RS_CODCLIENTE As ADODB.Recordset
Dim RS_PERFIL_USUARIO As ADODB.Recordset

'FIXIT: Declare 'DIM_INT_1' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_1
'FIXIT: Declare 'UserNuevo' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim UserNuevo
'FIXIT: Declare 'log' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Dim log
'FIXIT: Declare 'logcaja' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim logcaja
Dim DIM_INT_TIME_1 As Boolean

Private Sub refrescar()
'On Error GoTo menerr
'RS_PRODUCTO.Close
Set RS_PRODUCTO = New Recordset
Set RS_PERFIL_USUARIO = New Recordset
RS_PRODUCTO.Open "Select * from Clientes", STR_CONEXION, adOpenKeyset, adLockOptimistic
If IsNull(RS_PRODUCTO.Fields("Nombre")) Then
Text1.Text = ""
Else
Text1.Text = RS_PRODUCTO.Fields("nombre")
End If
If IsNull(RS_PRODUCTO.Fields("codigo")) Then
Text2.Text = ""
Else
Text2.Text = RS_PRODUCTO.Fields("codigo")
End If
If IsNull(RS_PRODUCTO.Fields("TELEFONO")) Then
Text3.Text = ""
Else
Text3.Text = RS_PRODUCTO.Fields("TELEFONO")
End If
If IsNull(RS_PRODUCTO.Fields("TELCONTACTO")) Then
Text8.Text = ""
Else
Text8.Text = RS_PRODUCTO.Fields("TELCONTACTO")
End If
If IsNull(RS_PRODUCTO.Fields("FAX")) Then
Text9.Text = ""
Else
Text9.Text = RS_PRODUCTO.Fields("FAX")
End If
If IsNull(RS_PRODUCTO.Fields("MOVIL")) Then
Text4.Text = ""
Else
Text4.Text = RS_PRODUCTO.Fields("MOVIL")
End If
If IsNull(RS_PRODUCTO.Fields("CIUDAD")) Then
Text6.Text = ""
Else
Text6.Text = RS_PRODUCTO.Fields("CIUDAD")
End If
If IsNull(RS_PRODUCTO.Fields("CONTACTO")) Then
Text5.Text = ""
Else
Text5.Text = RS_PRODUCTO.Fields("CONTACTO")
End If
If IsNull(RS_PRODUCTO.Fields("DIRECCION")) Then
Text7.Text = "0"
Else
Text7.Text = RS_PRODUCTO.Fields("DIRECCION")
End If

Set RS_LV = New Recordset
RS_LV.Open "select CODIGO,NOMBRE,CIUDAD from clientes order by nombre,ciudad", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
 With RS_LV
        If .BOF = True And .EOF = True Then
        ListView1.ListItems.Clear
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = ListView1.ListItems.Add(, , .Fields(0) & "")
                Items.SubItems(1) = .Fields(1) & ""
                Items.SubItems(2) = .Fields(2) & ""
                .MoveNext
            Loop
        End If
    End With
    
Set RS_LV = Nothing


Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,refrescar"
Close #1
End Sub

Private Sub BTN1_Click()

If RS_PRODUCTO.BOF = True And RS_PRODUCTO.EOF = True Then
Else
 RS_PRODUCTO.MoveFirst

 refrescar
End If
End Sub







Private Sub BTN5_Click()
'On Error GoTo menerr
PUB_42 = False
'Nuevo.BackColor = &H8000000F
LIMPIAR
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set RS_CLIENTES = New Recordset
RS_CLIENTES.Open "Select * from CLIENTES order by codigo", STR_CONEXION, adOpenKeyset, adLockOptimistic

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set RS_CODCLIENTE = New Recordset
RS_CODCLIENTE.Open "Select * from Clientes order by codigo", STR_CONEXION, adOpenKeyset, adLockOptimistic
If RS_CODCLIENTE.EOF = False Or RS_CODCLIENTE.BOF = False Then
With RS_CODCLIENTE
    .MoveFirst
    .MoveLast
    Text2.Text = .Fields("CODIGO") + 1
End With
Set RS_CODCLIENTE = Nothing
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
habilitar
Text2.Enabled = False
Text1.SetFocus
BTN6.Enabled = True
'Command24.Enabled = False
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,btn5"
Close #1
End Sub

Private Sub BTN6_Click()
'On Error GoTo menerr
Dim i As Integer
Dim ban As Integer
ban = 0
Dim criterio As String
Select Case Index
      Case 0
         If Text1.Text <> "" Then
         
                RS_PRODUCTO.AddNew
                 RS_PRODUCTO.Fields("nombre") = Text1.Text
                RS_PRODUCTO.Fields("codigo") = Text2.Text
                RS_PRODUCTO.Fields("telefono") = Text3.Text
                RS_PRODUCTO.Fields("movil") = Text4.Text
                RS_PRODUCTO.Fields("contacto") = Text5.Text
                 RS_PRODUCTO.Fields("telcontacto") = Text8.Text
                 RS_PRODUCTO.Fields("fax") = Text9.Text
                If Text6.Text = "" Then
                   RS_PRODUCTO.Fields("ciudad") = "s/n"
                Else
                   RS_PRODUCTO.Fields("ciudad") = Text6
                End If
                If Text7.Text = "" Then
                   RS_PRODUCTO.Fields("direccion") = "S/N"
                Else
                   RS_PRODUCTO.Fields("direccion") = Text7
                End If
                RS_PRODUCTO.Update

         
         Else
         MsgBox "Ingrese la Descripcion", vbCritical, "Mensaje de Error"
         ban = 1
         End If
      Case 1
            RS_PRODUCTO.CancelUpdate
            RS_PRODUCTO.MoveFirst
    refrescar
End Select
If ban <> 1 Then
refrescar
End If
deshabilitar

Set RS_LV = New Recordset
RS_LV.Open "select CODIGO,NOMBRE,CIUDAD from clientes order by nombre,ciudad", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
 With RS_LV
        If .BOF = True And .EOF = True Then
        ListView1.ListItems.Clear
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = ListView1.ListItems.Add(, , .Fields(0) & "")
                Items.SubItems(1) = .Fields(1) & ""
                Items.SubItems(2) = .Fields(2) & ""
                .MoveNext
            Loop
        End If
    End With
    
Set RS_LV = Nothing
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,btn6"
Close #1
End Sub

Private Sub BTN7_Click()
'On Error GoTo menerr
habilitar
'BTN6.Enabled = False
Command24.Enabled = True
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,btn7"
Close #1
End Sub

Private Sub BTN8_Click()

'On Error GoTo menerr
n = MsgBox("Esta Seguro que desea eliminar el registro?", vbYesNo, "Confirme Eliminacion")
If n = vbYes Then
Set RS_TOTAL = New Recordset
Dim PUB_SQL
PUB_SQL = "DELETE  FROM Clientes WHERE codigo= " & Text2.Text
RS_TOTAL.Open PUB_SQL, STR_CONEXION, adOpenKeyset, adLockOptimistic
Set RS_LV = Nothing
Set RS_LV = New Recordset
RS_LV.Open "select CODIGO,NOMBRE,CIUDAD from clientes order by nombre,ciudad", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
 With RS_LV
        If .BOF = True And .EOF = True Then
        ListView1.ListItems.Clear
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = ListView1.ListItems.Add(, , .Fields(0) & "")
                Items.SubItems(1) = .Fields(1) & ""
                Items.SubItems(2) = .Fields(2) & ""
                .MoveNext
            Loop
        End If
    End With
    
Set RS_LV = Nothing
End If
refrescar
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,btn8"
Close #1
End Sub





Private Sub Command24_Click()

'On Error GoTo menerr



Set RS_TOTAL = New Recordset
RS_TOTAL.Open "Select * from Clientes where codigo like '" & Text2 & "'", STR_CONEXION, adOpenKeyset, adLockOptimistic
'DIM_SQLSUM = DIM_SQLSUM & "and  ctagrupo LIKE '" & DIMString & "'"

With RS_TOTAL
         If Text1.Text = "" Then
            .Fields("nombre") = "s/n"
         Else
            .Fields("nombre") = Text1
         End If
         
         
         
         If Text2.Text = "" Then
            .Fields("codigo") = "0"
         Else
            .Fields("codigo") = Text2
         End If
         
         
         If Text3.Text = "" Then
            .Fields("telefono") = "0"
         Else
            .Fields("telefono") = Text3
         End If
         
         
         If Text4.Text = "" Then
            .Fields("MOVIL") = "0"
         Else
            .Fields("MOVIL") = Text4
         End If
         

         If Text5.Text = "" Then
            .Fields("contacto") = "s/n"
         Else
            .Fields("contacto") = Text5
         End If



         If Text8.Text = "" Then
            .Fields("telcontacto") = "0"
         Else
            .Fields("telcontacto") = Text8
         End If
         
          
         If Text9.Text = "" Then
            .Fields("FAX") = "0"
         Else
            .Fields("FAX") = Text9
         End If
          
          
         If Text6.Text = "" Then
            .Fields("ciudad") = "s/n"
         Else
            .Fields("ciudad") = Text6
         End If
         If Text7.Text = "" Then
            .Fields("direccion") = "S/N"
         Else
            .Fields("direccion") = Text7
         End If
         .Update

End With
Set RS_TOTAL = Nothing


refrescar

deshabilitar
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,command24"
Close #1
End Sub



Private Sub Form_Load()
Set pic = Nothing
If IsEmpty(STR_CONEXION) Then

STR_CONEXION.Close
habilitar

gsRutaBaseDatos = "HARINA"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HARINA & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\maturavebmp.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE HARINA"

Else

STR_CONEXION.Close

gsRutaBaseDatos = "HARINA"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HARINA & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\maturavebmp.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE HARINA"

End If

'*****************************************************************************************************************
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from Ventascrd Order by NoDoc", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
refrescar

'FIXIT: Declare 'Z' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE

'*****************************************************************************************************************

Dim Z
PUB_42 = True
deshabilitar
End Sub
Public Function lstvDatos_a_cero()
On Error GoTo menerr
'Aspecto de listview
    With ListView1
        .View = lvwReport
        .ColumnHeaders.Clear
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "COIGO", 500
        .ColumnHeaders.Add , , "NOMBRE", 2500
        .ColumnHeaders.Add , , "CIUDAD", 2500
    End With
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''
Exit Function
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,lvscero"
Close #1
End Function
Private Sub Guardar_Click()

End Sub

Private Sub Modificar_Click()
On Error GoTo menerr
Dim i As Integer
If RS_PRODUCTO.Fields("BORRAR") = True Then
MsgBox "NO PUEDE MODIFICAR ESTE USUARIO"
RS_PRODUCTO.MoveFirst
Else
txtUsuario.Enabled = True
txtContraseña.Enabled = True
TXTNOMBRE.Enabled = True
TXTAPELLIDO.Enabled = True

'desabilitar_botones
End If
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,modificar"
Close #1
End Sub


Private Sub LIMPIAR()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""

'DCACCESO.Text
'DtCaja.ListField -1
End Sub
Private Sub Salir_Click()

Unload Me
FrmInicio3.Show
'RS_PRODUCTO.Close
'cajas.Close
End Sub

Private Sub habilitar()

Text1.Enabled = True
Text8.Enabled = True
Text2.Enabled = True
Text9.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True

'DCACCESO.Text
'DtCaja.ListField -1
End Sub
Private Sub deshabilitar()

Text1.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text6.Enabled = False
Text7.Enabled = False

'DCACCESO.Text
'DtCaja.ListField -1
End Sub












Private Sub Image5_Click()

Unload Me
FrmInicio3.Show
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo menerr
Dim Cuenta As ADODB.Recordset
Set Cuenta = New Recordset
Cuenta.Open "SELECT * FROM Clientes where codigo like'" & Item.Text & "' ", STR_CONEXION, adOpenStatic, adLockReadOnly
With Cuenta
    If .EOF Then
        '.MoveFirst
        MsgBox "No se localizo la Cuenta [" & a & "]", vbCritical, "Error de busqueda"
        Exit Sub
    End If
If IsNull(Cuenta.Fields("Nombre")) Then
Text1.Text = ""
Else
Text1.Text = Cuenta.Fields("nombre")
End If
If IsNull(Cuenta.Fields("codigo")) Then
Text2.Text = ""
Else
Text2.Text = Cuenta.Fields("codigo")
End If
If IsNull(Cuenta.Fields("TELEFONO")) Then
Text3.Text = ""
Else
Text3.Text = Cuenta.Fields("TELEFONO")
End If
If IsNull(Cuenta.Fields("TELCONTACTO")) Then
Text8.Text = ""
Else
Text8.Text = Cuenta.Fields("TELCONTACTO")
End If
If IsNull(Cuenta.Fields("FAX")) Then
Text9.Text = ""
Else
Text9.Text = Cuenta.Fields("FAX")
End If
If IsNull(Cuenta.Fields("MOVIL")) Then
Text4.Text = ""
Else
Text4.Text = Cuenta.Fields("MOVIL")
End If
If IsNull(Cuenta.Fields("CIUDAD")) Then
Text6.Text = ""
Else
Text6.Text = Cuenta.Fields("CIUDAD")
End If
If IsNull(Cuenta.Fields("CONTACTO")) Then
Text5.Text = ""
Else
Text5.Text = Cuenta.Fields("CONTACTO")
End If
If IsNull(Cuenta.Fields("DIRECCION")) Then
Text7.Text = "0"
Else
Text7.Text = Cuenta.Fields("DIRECCION")
End If
End With
Set Cuenta = Nothing
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,lvsitemclick"
Close #1
End Sub

Public Sub TextSelected()
'On Error GoTo menerr
Dim i As Integer
'FIXIT: Declare 'oMyTextBox' con un tipo de datos de enlace en tiempo de compilación       FixIT90210ae-R1672-R1B8ZE
Dim oMyTextBox As Object


Set oMyTextBox = Screen.ActiveControl
If TypeName(oMyTextBox) = "TextBox" Then
i = Len(oMyTextBox.Text)
oMyTextBox.SelStart = 0
oMyTextBox.SelLength = i
End If
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "frmaddcliente,textselect"
Close #1
End Sub

Private Sub Text1_Change()
'FIXIT: Declare 'DIM_SQL' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim DIM_SQL
    With Text1

        If InStr(1, .Text, "'") <> 0 Or InStr(1, .Text, "[") <> 0 Or _
            InStr(1, .Text, "|") <> 0 Or InStr(1, .Text, """") <> 0 Or _
            InStr(1, .Text, "*") <> 0 Or InStr(1, .Text, "/") <> 0 Then
           .Text = ""
            Exit Sub
        Else
'FIXIT: Declare 'DIM_NOMBRE' con un tipo de datos de enlace en tiempo de compilación       FixIT90210ae-R1672-R1B8ZE
            Dim DIM_NOMBRE
            DIM_NOMBRE = .Text
            DIM_SQL = "Select Nombre From Clientes where Nombre like '%" & Text1 & "%'"
            'like '%Circul%'",'*" & a & "*'"
        End If
    End With
    Dim Cuenta As ADODB.Recordset
Set Cuenta = New Recordset




'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''
    With ListView1
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "CODIGO", 1200
        .ColumnHeaders.Add , , "NOMBRE", 5000
        .ColumnHeaders.Add , , "CIUDAD", 2500
        .ColumnHeaders.Add , , "DIRECCION", 2500
    End With
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''

Cuenta.Open "Select * From Clientes where Nombre like '%" & Text1 & "%'", STR_CONEXION, adOpenStatic, adLockReadOnly
With Cuenta
       If Cuenta.BOF = True And Cuenta.EOF = True Then
        ListView1.ListItems.Clear
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = ListView1.ListItems.Add(, , .Fields(0) & "")
                Items.SubItems(1) = .Fields(1) & ""
                                Items.SubItems(2) = .Fields(7) & ""
                Items.SubItems(3) = .Fields(10) & ""
                .MoveNext
            Loop
        End If
End With
 
Set Cuenta = Nothing

End Sub

Private Sub Text5_GotFocus()

TextSelected
End Sub

Private Sub Text5_LostFocus()

TextSelected
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
'On Error GoTo menerr
Set pic = Nothing
If IsEmpty(STR_CONEXION) Then

STR_CONEXION.Close
habilitar

gsRutaBaseDatos = "ARROZ"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

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

STR_CONEXION.Close

gsRutaBaseDatos = "ARROZ"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\ARROZ.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE ARROZ"

End If

'*****************************************************************************************************************
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from Ventascrd Order by NoDoc", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
refrescar
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "F FrmNCredito,command1"
Close #1

End Sub
Private Sub Command2_Click()
'On Error GoTo menerr
Set pic = Nothing
If IsEmpty(STR_CONEXION) Then

STR_CONEXION.Close
habilitar

gsRutaBaseDatos = "MAIZ"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_MAIZ & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\maturavebmp.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE MAIZ"

Else

STR_CONEXION.Close

gsRutaBaseDatos = "ARROZ"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_MAIZ & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\maturavebmp.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE MAIZ"

End If

'*****************************************************************************************************************
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from Ventascrd Order by NoDoc", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
refrescar
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "F FrmNCredito,command1"
Close #1

End Sub

Private Sub Command3_Click()
'On Error GoTo menerr
Set pic = Nothing
If IsEmpty(STR_CONEXION) Then

STR_CONEXION.Close
habilitar

gsRutaBaseDatos = "CONCENTRADO"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_CONCENTRADO & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\FACOCA.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE CONCENTRADO"

Else

STR_CONEXION.Close

gsRutaBaseDatos = "CONCENTRADO"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_CONCENTRADO & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\FACOCA.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE CONCENTRADO"

End If

'*****************************************************************************************************************
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from Ventascrd Order by NoDoc", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
refrescar
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "F FrmNCredito,command1"
Close #1

End Sub

Private Sub Command4_Click()
'On Error GoTo menerr
Set pic = Nothing
If IsEmpty(STR_CONEXION) Then

STR_CONEXION.Close
habilitar

gsRutaBaseDatos = "HUEVO"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HUEVO & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

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

STR_CONEXION.Close

gsRutaBaseDatos = "HUEVO"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HUEVO & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

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

'*****************************************************************************************************************
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from Ventascrd Order by NoDoc", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
refrescar
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "F FrmNCredito,command1"
Close #1

End Sub

Private Sub Command6_Click()
On Error Resume Next
Set pic = Nothing
If IsEmpty(STR_CONEXION) Then

STR_CONEXION.Close
habilitar

gsRutaBaseDatos = "HARINA"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HARINA & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\maturavebmp.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE HARINA"

Else

STR_CONEXION.Close

gsRutaBaseDatos = "HARINA"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_HARINA & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\maturavebmp.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE HARINA"

End If

'*****************************************************************************************************************
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from Ventascrd Order by NoDoc", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
refrescar
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "F FrmNCredito,command1"
Close #1

End Sub
Private Sub Command7_Click()
'On Error GoTo menerr
Set pic = Nothing
If IsEmpty(STR_CONEXION) Then

STR_CONEXION.Close
habilitar

gsRutaBaseDatos = "HARINA"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_MATURAVE & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\FACOCA.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE MATURAVE"

Else

STR_CONEXION.Close

gsRutaBaseDatos = "HARINA"
    gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_MATURAVE & ";Persist Security Info=False"
'gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos_ARROZ & ";Persist Security Info=False"

STR_CONEXION.ConnectionString = gsCadenaConexion
STR_CONEXION.Open



Set RS_EMPRESA = New Recordset
RS_EMPRESA.Open "Select *from EMPRESA ", STR_CONEXION, adOpenKeyset, adLockOptimistic

DIM_DIRECCION = RS_EMPRESA.Fields("direccion")
DIM_TELEFONO = RS_EMPRESA.Fields("TELEFONO1")
DIM_TIENDA = RS_EMPRESA.Fields("TIENDA")
DIM_EMPRESA = RS_EMPRESA.Fields("empresa")
DIM_RTN = RS_EMPRESA.Fields("rtn")
Set pic = LoadPicture(App.Path & "\iconos\FACOCA.bmp")
RS_EMPRESA.Close
Set RS_EMPRESA = Nothing

PubNegocio = "MATURAVE MATURAVE"

End If

'*****************************************************************************************************************
Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from Ventascrd Order by NoDoc", STR_CONEXION, adOpenKeyset, adLockOptimistic
lstvDatos_a_cero
refrescar
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "F FrmNCredito,command1"
Close #1

End Sub

Private Sub BTN9_Click()
Unload Me
FrmInicio3.Show
End Sub


