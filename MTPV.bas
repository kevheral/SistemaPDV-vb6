Attribute VB_Name = "MTPV"
Public PUB_CONEXION_EASY As New ADODB.Connection


Public DIM_USUARIO As Boolean

Public PUB_ODBC_PDV
Public PUB_ODBC_PRINCIPAL

Public gsRutaBaseDatos As String
Public gsPasswordUsua As String

Public adoComd As Command
Public rsComd As Recordset
Public Items As ListItem
Public DIM_NODOC
Public DIM_DIRECCION
Public DIM_CAJA
Public DIM_TELEFONO
Public DIM_NIVEL
Public DIM_TIENDA
Public DIM_EMPRESA
Public PUB_VALOR_C
Public DIM_RTN
Public DIM_TIPO
Public DIM_IMPUESTO
Public DIM_DOCFIN
Public DIM_REST
Public PUB_CANTIDAD1
Public DIM_CLIENTE
Public DIM_RTNCIENTE
Public DimNumTarjeta
Public DimOPTarjeta1
Public DimOPTarjeta
Public DimNumTarjetaOper
Public DIM_NUMERO
Public DIM_AceptarE As Boolean
Public DIM_SUMTOTAL
Public DIM_SUBTOTAL
Public DIM_SUMDESCUENTO
Public DIM_SUMISV
Public DIM_VIEJO_FORMA As Boolean
Public DIM_SUMVALOR
Public NumAlt
Public DIM_FORMAVENTA
Public PUB_IMPUESTO
Public PUB_CANTIDAD As Boolean
Public PUB_RED As Boolean
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''
Public PUB_USERNOMBRE
Public PUB_USERCODIGO
Public PUB_USERCAJA
Public PUB_USUARIO
Public PUB_NIVEL


Public Sub AbrirDB()
'On Error GoTo menerr
'PUB_ODBC_PDV = "DSN=PDV1"
'PUB_CONEXION_EASY.Open PUB_ODBC_PDV
'PUB_CONEXION_EASY.Open PUB_ODBC_PDV, , "IGLESIA"

'PUB_ODBC_PRINCIPAL = "DSN=PRINCIPAL"
'PUB_CONEXION_PRINCIPAL.Open PUB_ODBC_PRINCIPAL

'PUB_RED = True
'Exit Sub
'menerr:

Set PUB_CONEXION_EASY = New ADODB.Connection
Set adoComd = New ADODB.Command
Set rsComd = New ADODB.Recordset '

gsCadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; data source=" & gsRutaBaseDatos & ";Persist Security Info=False"
    PUB_CONEXION_EASY.ConnectionString = gsCadenaConexion
    PUB_CONEXION_EASY.Properties("Jet OLEDB:Database Password") = gsPasswordUsua
    PUB_CONEXION_EASY.Open
'PUB_RED = False
End Sub

