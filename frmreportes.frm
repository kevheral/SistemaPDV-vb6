VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FBFD55C6-C23C-11D3-B65D-004005E66149}#1.0#0"; "swiftprint.ocx"
Begin VB.Form Reportes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "REPORTES"
   ClientHeight    =   11565
   ClientLeft      =   45
   ClientTop       =   90
   ClientWidth     =   19110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11565
   ScaleWidth      =   19110
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   38
      Text            =   "Combo1"
      Top             =   4200
      Width           =   5055
   End
   Begin VB.Frame Frame4 
      Height          =   2295
      Left            =   5160
      TabIndex        =   28
      Top             =   9240
      Width           =   11775
      Begin VB.CommandButton Command22 
         Caption         =   "Reporte por Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   7680
         TabIndex        =   34
         Top             =   1440
         Width           =   3500
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Informes Tipo  Dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   3500
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Informes Tipo Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   3960
         TabIndex        =   30
         Top             =   600
         Width           =   3500
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Informes Tipo Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   7680
         TabIndex        =   29
         Top             =   600
         Width           =   3500
      End
      Begin VB.Label Label4 
         Caption         =   "PRODUCTOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2295
      Left            =   5160
      TabIndex        =   23
      Top             =   6840
      Width           =   11775
      Begin VB.CommandButton Command18 
         Caption         =   "Reporte Condensado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Width           =   3500
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cliente Deudores"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   3960
         TabIndex        =   36
         Top             =   1440
         Width           =   3500
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Reporte por Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   7680
         TabIndex        =   27
         Top             =   1440
         Width           =   3500
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clientes Dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   3500
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Clientes Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   3960
         TabIndex        =   25
         Top             =   600
         Width           =   3500
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Clientes Saldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   7680
         TabIndex        =   24
         Top             =   600
         Width           =   3500
      End
      Begin VB.Label Label5 
         Caption         =   "CLIENTES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   5160
      TabIndex        =   17
      Top             =   4320
      Width           =   11775
      Begin VB.CommandButton Command16 
         Caption         =   "Vendedores Condesado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4320
         TabIndex        =   35
         Top             =   1560
         Width           =   3500
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Reporte por Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   8040
         TabIndex        =   22
         Top             =   1560
         Width           =   3500
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Vendedores Dias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   720
         TabIndex        =   21
         Top             =   840
         Width           =   3500
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Vendedores Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4320
         TabIndex        =   20
         Top             =   840
         Width           =   3500
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Vendedores Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   8040
         TabIndex        =   19
         Top             =   840
         Width           =   3500
      End
      Begin VB.Label Label3 
         Caption         =   "VENDEDORES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   5160
      TabIndex        =   9
      Top             =   1320
      Width           =   11775
      Begin VB.CommandButton Command20 
         Caption         =   "Ventas Ciudad 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   8040
         TabIndex        =   41
         Top             =   2160
         Width           =   3500
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Ventas Ciudad 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4320
         TabIndex        =   40
         Top             =   2160
         Width           =   3500
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ventas Credito Dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   720
         TabIndex        =   16
         Top             =   1440
         Width           =   3500
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ventas Credito Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4320
         TabIndex        =   15
         Top             =   1440
         Width           =   3500
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ventas Ciudad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   720
         TabIndex        =   14
         Top             =   2160
         Width           =   3500
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ventas Contado Dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   3500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ventas Contado Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   4320
         TabIndex        =   12
         Top             =   720
         Width           =   3500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ventas Unidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   8040
         TabIndex        =   11
         Top             =   720
         Width           =   3500
      End
      Begin VB.Label Label2 
         Caption         =   "VENTAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton BTN9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   5040
      Picture         =   "frmreportes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "SALIR"
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   3000
      Width           =   5055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1920
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   89915393
      CurrentDate     =   39731
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   89915393
      CurrentDate     =   39731
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   39
      Top             =   3600
      Width           =   2415
   End
   Begin SwiftPrintLib.SwiftPrint SpDoc 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "AL :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "DEL :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_REPORTES As ADODB.Recordset
Dim RS_REPORTES1 As ADODB.Recordset
Dim RS_REPORTES2 As ADODB.Recordset
Dim RS_REPORTES3 As ADODB.Recordset
Dim RS_REPORTES4 As ADODB.Recordset
Dim RS_ELIMINAR As ADODB.Recordset
Dim RS_TOTAL As ADODB.Recordset
Dim RS_Fecha As ADODB.Recordset
 Dim DIM_SQL As String
  Dim DIM_SQLSEL As String
  Dim DIM_SQLSUM As String
    Dim DIM_SQLP As String
  Dim DIM_CLIENTES
  Dim NomVendedor
  Dim DIMTITULOPAGINA
  Dim RptTitle As String
   Dim nFooterTop As Integer
   Dim DimPie As Integer
  Dim DIM_TITULORPT
Dim Fecha_Inicial As String
Dim Fecha_Final As String
Dim hora_Inicial As String
Dim hora_final As String
Dim oPrn As Printer
'FIXIT: Declare 'DIM_DIA_DET' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Dim DIM_DIA_DET
'FIXIT: Declare 'DIM_MES_DET' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Dim DIM_MES_DET
'FIXIT: Declare 'DIM_DIA_REPORT' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim DIM_DIA_REPORT
'FIXIT: Declare 'DIM_DIA' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim DIM_DIA
'FIXIT: Declare 'DIM_FORMA' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_FORMA, DIM_FRMCREDITO
'FIXIT: Declare 'DIM_MES' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim DIM_MES
'FIXIT: Declare 'DIM_AÑO' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim DIM_AÑO
'FIXIT: Declare 'DIM_WEEK' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Dim DIM_WEEK
'FIXIT: Declare 'DIM_CAJA' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Dim DIM_CAJA

Private Sub Command16_Click()
SpDoc.DocClearPage
SpDoc.DocBegin

     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = Command13.Caption
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY VENDEDOR ASC "
       DIM_SQL = "select * from Ventas_Mes where mes like '" & DIM_MES & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "VENDEDOR"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select VENDEDOR from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
  'DIM_SQL = "select VENDEDOR,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      DIM_SQL = "select NODE from Ventas_Mes where mes like '" & DIM_MES & "'"
      DIM_SQL = DIM_SQL & " GROUP BY NODE ORDER BY NODE ASC "
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("node").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                    Dim Cuenta As ADODB.Recordset
                    Set Cuenta = New Recordset
                    Cuenta.Open "Select * From VENDEDORES where CODIGO like '%" & .Fields("NODE").Value & "%'", PUB_CONEXION_ADMIN, adOpenStatic, adLockReadOnly
                    'MsgBox PUB_CONEXION_ADMIN
                    SpDoc.TextOut 119, nRows, Cuenta.Fields("NOMBRE").Value
                    Set Cuenta = Nothing
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQL = "select sum(salida),sum(total) from InvSalida where node like '" & .Fields("NODE").Value & "'"
                                    
                                         'DIM_SQL = "select VENDEDOR,producto,salida,forma,total from Ventas_Mes where mes like '" & DIM_MES & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "VENDEDOR"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, "Cantidad total : " & RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, "Ventas Totales Vendedor : " & Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"


Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop

DIM_SQLSUM = "select SUM(total) from Invsalida"

 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 45, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 1500, nRows, "TOTAL VENTAS ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2000, nRows, "0"
    Else
    SpDoc.TextOut 2000, nRows, "Lps. " & Format(.Fields(0).Value, "#,##0.00")
    End If

End With
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview

End Sub

Private Sub Command17_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
DIM_FORMA = "CREDITO"
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = "COMPRAS POR CLIENTE"
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY VENDEDOR ASC "
       DIM_SQL = "select * from Ventas_Mes where mes like '" & DIM_MES & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "VENDEDOR"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select VENDEDOR from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
  'DIM_SQL = "select VENDEDOR,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      DIM_SQL = "select nombre from ClientesDts GROUP BY nombre"
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("Nombre").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("Nombre").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQL = "select sum(cantidad),sum(valor) from ClientesDts where nombre like '" & .Fields("nombre").Value & "' "
                                    
                                         'DIM_SQL = "select VENDEDOR,producto,salida,forma,total from Ventas_Mes where mes like '" & DIM_MES & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "VENDEDOR"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, "Cantidad : " & RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, "Compras por Cliente : " & Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"


Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop

DIM_SQLSUM = "select SUM(valor) from ClientesDts"

 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 45, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 1500, nRows, "TOTAL COMPRADO ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2000, nRows, "0"
    Else
    SpDoc.TextOut 2000, nRows, "Lps. " & Format(.Fields(0).Value, "#,##0.00")
    End If

End With
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview


SpDoc.DocClearPage
SpDoc.DocBegin

     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
'Dim nRows As Long, nCols As Long, nItem As Long
    'Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    'Dim X As Long, Y As Long, nIdx As Long
    'Dim center As Long, lMaxY As Long
    'Dim strText As String, CharsDrawn As Long
    'Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = "COMPRAS POR CLIENTE"
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY VENDEDOR ASC "
       DIM_SQL = "select * from Ventas_Mes where mes like '" & DIM_MES & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "VENDEDOR"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select VENDEDOR from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
  'DIM_SQL = "select VENDEDOR,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      DIM_SQL = "select nombre from ClientesDts GROUP BY nombre"
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("Nombre").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("Nombre").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQL = "select sum(cantidad),sum(valor) from ClientesDts where nombre like '" & .Fields("nombre").Value & "' "
                                    
                                         'DIM_SQL = "select VENDEDOR,producto,salida,forma,total from Ventas_Mes where mes like '" & DIM_MES & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "VENDEDOR"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, "Cantidad : " & RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, "Compras por Cliente : " & Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"


Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop

DIM_SQLSUM = "select SUM(valor) from ClientesDts"

 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 45, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 1500, nRows, "TOTAL COMPRADO ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2000, nRows, "0"
    Else
    SpDoc.TextOut 2000, nRows, "Lps. " & Format(.Fields(0).Value, "#,##0.00")
    End If

End With
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview



End Sub

Private Sub Command18_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
DIM_FORMA = "CREDITO"
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = "COMPRAS POR CLIENTE"
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY VENDEDOR ASC "
       DIM_SQL = "select * from Ventas_Mes where mes like '" & DIM_MES & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "VENDEDOR"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select VENDEDOR from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
  'DIM_SQL = "select VENDEDOR,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      DIM_SQL = "select nombre from ClientesDts GROUP BY nombre"
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("Nombre").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("Nombre").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQL = "select sum(cantidad),sum(valor) from ClientesDts where nombre like '" & .Fields("nombre").Value & "' "
                                    
                                         'DIM_SQL = "select VENDEDOR,producto,salida,forma,total from Ventas_Mes where mes like '" & DIM_MES & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "VENDEDOR"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, "Cantidad : " & RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, "Compras por Cliente : " & Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"


Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop

DIM_SQLSUM = "select SUM(valor) from ClientesDts"

 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 45, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 1500, nRows, "TOTAL COMPRADO ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2000, nRows, "0"
    Else
    SpDoc.TextOut 2000, nRows, "Lps. " & Format(.Fields(0).Value, "#,##0.00")
    End If

End With
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview


End Sub


Private Sub Command19_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT
      DIM_TITULORPT = " VENTAS POR CIUDAD " & DIM_DIA_DET
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

  DIM_SQL = "select * from InvSalida "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY cliente ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "Cliente"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
DIM_SQL = "select CIUDAD from clienteciudad "
  'DIM_SQL = "select cliente,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
DIM_SQL = DIM_SQL & " GROUP BY CIUDAD ORDER BY CIUDAD ASC "
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("CIUDAD").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("CIUDAD").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQLSEL = "select nombre,unidades,producto,valor from clientesciudad2 where ciudad like '" & .Fields("ciudad").Value & "'"
                                    'DIM_SQLSEL = DIM_SQLSEL & " GROUP BY cliente "
                                    'DIM_SQL = DIM_SQL & " GROUP BY cliente ORDER BY cliente ASC "
                                    RS_REPORTES1.Open DIM_SQLSEL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "Cliente"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3

                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                            
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                            
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    
DIM_SQLSUM = "select SUM(valor) from clientesciudad2 where ciudad like '" & .Fields("ciudad").Value & "'"
Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
End Sub

Private Sub Command20_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT
      DIM_TITULORPT = " VENTAS POR CIUDAD " & DIM_DIA_DET
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

  DIM_SQL = "select * from InvSalida "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY cliente ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "Cliente"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
DIM_SQL = "select CIUDAD from ClientesCiudad3 "
  'DIM_SQL = "select cliente,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
DIM_SQL = DIM_SQL & " GROUP BY CIUDAD ORDER BY CIUDAD ASC "
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("CIUDAD").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("CIUDAD").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQLSEL = "select vendedor,nombre,unidades,valor from ClientesCiudad3 where ciudad like '" & .Fields("ciudad").Value & "'"
                                    'DIM_SQLSEL = DIM_SQLSEL & " GROUP BY cliente "
                                    'DIM_SQL = DIM_SQL & " GROUP BY cliente ORDER BY cliente ASC "
                                    RS_REPORTES1.Open DIM_SQLSEL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "Cliente"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3


                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                            
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                            
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    
DIM_SQLSUM = "select SUM(valor) from clientesciudad2 where ciudad like '" & .Fields("ciudad").Value & "'"
Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
End Sub


Private Sub Command22_Click()
SpDoc.DocClearPage
SpDoc.DocBegin

     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
RptTitle = "REPORTE DE PRODUCTO POR DIA"
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

  DIM_SQL = "select * from InvSalida ORDER BY PRODUCTO ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY VENDEDOR ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "VENDEDOR"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
DIM_SQL = "select fecha from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
DIM_SQL = DIM_SQL & " GROUP BY fecha ORDER BY fecha ASC  "
'DIM_SQL = "select VENDEDOR from InvSalida GROUP BY VENDEDOR ORDER BY VENDEDOR ASC "
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("fecha").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("fecha").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQLSEL = "select producto from InvSalida WHERE fecha like '" & .Fields("fecha").Value & "' "
                                    DIM_SQLSEL = DIM_SQLSEL & " GROUP BY producto "
                                    'DIM_SQL = DIM_SQL & " AND producto like '" & DIM_FORMA & "' ORDER BY producto ASC "
                                    'ORDER BY producto ASC "
                                    'DIM_SQL = DIM_SQL & " GROUP BY VENDEDOR ORDER BY VENDEDOR ASC "
                                    RS_REPORTES1.Open DIM_SQLSEL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        
                                        
                                        
                                        
                                        
                                                                If IsNull(RS_REPORTES1.Fields(0).Value) Then
                                                                SpDoc.TextOut 119, nRows, "0"
                                                                Else
                                                                SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(0).Value
                                                                End If
                                                                'nCols = nCols + 350
                                                                
                                                                                DIM_SQLP = "select SUM(salida) from InvSalida WHERE fecha like '" & .Fields("fecha").Value & "' "
                                                                                DIM_SQLP = DIM_SQLP & " AND producto like '" & RS_REPORTES1.Fields("producto").Value & "' "
                                                                                Set RS_REPORTES3 = New Recordset
                                                                                RS_REPORTES3.Open DIM_SQLP, PUB_CONEXION_EASY

                                                                                    If IsNull(RS_REPORTES3.Fields(0).Value) Then
                                                                                    SpDoc.TextOut 1500, nRows, "0"
                                                                                    Else
                                                                                    SpDoc.TextOut 1500, nRows, RS_REPORTES3.Fields(0).Value
                                                                                    End If

                                                                                
                                                                                Set RS_REPORTES3 = Nothing
                                                                                
                                                                                
                                                                                
                                                                                
                                                                                Set RS_REPORTES3 = New Recordset
                                                                                DIM_SQLP = "select SUM(total) from InvSalida WHERE fecha like '" & .Fields("fecha").Value & "' "
                                                                                DIM_SQLP = DIM_SQLP & " AND producto like '" & RS_REPORTES1.Fields("producto").Value & "' "
                                                                                Set RS_REPORTES3 = New Recordset
                                                                                RS_REPORTES3.Open DIM_SQLP, PUB_CONEXION_EASY

                                                                                    If IsNull(RS_REPORTES3.Fields(0).Value) Then
                                                                                    SpDoc.TextOut 1800, nRows, "0"
                                                                                    Else
                                                                                    SpDoc.TextOut 1800, nRows, Format(RS_REPORTES3.Fields(0).Value, "#,##0.00")
                                                                                    End If

                                                                                
                                                                                Set RS_REPORTES3 = Nothing
                                                                
                                                                

                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    
DIM_SQLSUM = "select SUM(salida),SUM(total) from InvSalida where  fecha like '" & .Fields("fecha").Value & "' "
Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 1000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 1500, nRows, "0"
    Else
    SpDoc.TextOut 1500, nRows, .Fields(0).Value
    End If
    If IsNull(.Fields(1).Value) Then
    SpDoc.TextOut 1800, nRows, "0"
    Else
    SpDoc.TextOut 1800, nRows, Format(.Fields(1).Value, "#,##0.00")
    End If
End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''
Fecha_Inicial = Format(DTPicker1, "mm/dd/yyyy")
Fecha_Final = Format(DTPicker2, "mm/dd/yyyy")
       
hora_Inicial = Format(DTPicker3, "mm/dd/yyyy")
hora_final = Format(DTPicker4, "mm/dd/yyyy")

DTPicker1.Value = Date
DTPicker2.Value = Date
DIM_DIA_DET = Format(DTPicker1, "dd mmmm yyyy")
DIM_MES_DET = Format(DTPicker1, "mmmm yyyy")
DIM_DIA_REPORT = Format(DTPicker1, "dd mmm yy")
DIM_DIA = Format(DTPicker1, "dd mmm yyyy")
DIM_MES = Format(DTPicker1, "mmmm yyyy")
DIM_AÑO = Format(DTPicker1, "yyyy")
DIM_WEEK = Format(DTPicker1, "WW")
DIM_CAJA = Combo5
SpDoc.DocBegin
'SpDoc.DocClearPage
SpDoc.WindowOwner = Me.hwnd

'*****************************************************************************************************************
Set RS_TOTAL = New Recordset

DIM_SQL = "select nombre from Inventario01 GROUP BY nombre "
RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'PUB_CONEXION_ADMIN
'PUB_CONEXION_ADMINdf
Combo3.Clear
Do Until RS_TOTAL.EOF = True


    If IsNull(RS_TOTAL.Fields(0).Value) Then

    Else
    Combo3.AddItem RS_TOTAL.Fields("nombre")
    End If


RS_TOTAL.MoveNext

Loop
Set RS_TOTAL = Nothing
'*****************************************************************************************************************
'*****************************************************************************************************************
Set RS_TOTAL = New Recordset

DIM_SQL = "select CLIENTE from InvSalida GROUP BY CLIENTE "
RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'PUB_CONEXION_ADMIN
'PUB_CONEXION_ADMINdf
Combo1.Clear
Do Until RS_TOTAL.EOF = True


    If IsNull(RS_TOTAL.Fields(0).Value) Then

    Else
    Combo1.AddItem RS_TOTAL.Fields("CLIENTE")
    End If



RS_TOTAL.MoveNext

Loop
Set RS_TOTAL = Nothing
'*****************************************************************************************************************
Set RS_TOTAL = New Recordset
DIM_SQL = "select VENDEDOR from InvSalida GROUP BY VENDEDOR "
RS_TOTAL.Open DIM_SQL, PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'PUB_CONEXION_ADMIN
'PUB_CONEXION_ADMINdf
Combo2.Clear
Do Until RS_TOTAL.EOF = True

    If IsNull(RS_TOTAL.Fields(0).Value) Then

    Else
    Combo2.AddItem RS_TOTAL.Fields("VENDEDOR")
    End If
RS_TOTAL.MoveNext

Loop
Set RS_TOTAL = Nothing
'BIXOLON SRP-275

'Busca en todas las imrpesoras.
For Each oPrn In Printers
'Busca el generador de PDF.
If oPrn.DeviceName = "EPSON LX-300+II ESC/P" Then
'If oPrn.DeviceName = "EPSON LX-300+II" Then
'Se encontro, pon esta impresora como predeterminada
'y sal del FOR Loop.
Set Printer = oPrn
Exit For
End If
Next

End Sub
Private Sub BTN9_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim DIMSALDO, DIMSUBTOTAL, DIMISV
  SpDoc.DocClearPage
SpDoc.DocBegin
      DIM_TITULORPT = "VENTAS POR DIA CONTADO  " & DIM_DIA_DET
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset
     DIM_FORMA = "CONTADO"
  DIM_SQL = "select fecha,producto,salida,total,cliente,ndventas,Talla,forma from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY ndventas ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
    SpDoc.TextOut 139, 210, "Fecha"
    SpDoc.TextOut 269, 210, "Producto"
    SpDoc.TextOut 700, 210, "Cantidad"
    SpDoc.TextOut 900, 210, "Valor"
    SpDoc.TextOut 1100, 210, "Cliente"
    SpDoc.TextOut 1550, 210, "NDoc"
    SpDoc.TextOut 1800, 210, "Forma"



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
  DIM_SQL = "select fecha,producto,salida,total,cliente,ndventas,Talla,forma from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY ndventas ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    num_fields = .Fields.Count
    For i = 0 To num_fields - 1
    Select Case i
    Case 0
    'SpDoc.TextOut 139, 210, "Fecha"
    'SpDoc.TextOut 269, 210, "Producto"
    'SpDoc.TextOut 700, 210, "Cantidad"
    'SpDoc.TextOut 900, 210, "Valor"
    'SpDoc.TextOut 1050, 210, "Cliente"
    'SpDoc.TextOut 1650, 210, "NDoc"
    'SpDoc.TextOut 1800, 210, "Forma"
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 119, nRows, "0"
    Else
    SpDoc.TextOut 119, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 1
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 299, nRows, "0"
    Else
    SpDoc.TextOut 299, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 2
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 750, nRows, "0"
    Else
    SpDoc.TextOut 750, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 3
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 850, nRows, "0"
    Else
    SpDoc.TextOut 850, nRows, Format(.Fields(i).Value, "#,##0.00")
    End If
    'nCols = nCols + 350
    
    Case 4
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1000, nRows, "0"
    Else
    SpDoc.TextOut 1000, nRows, .Fields(i).Value
    End If
    
    Case 5
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1600, nRows, "0"
    Else
    SpDoc.TextOut 1600, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 6
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1700, nRows, "0"
    Else
    SpDoc.TextOut 1700, nRows, .Fields(i).Value
    End If
    Case 7
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1850, nRows, "0"
    Else
    SpDoc.TextOut 1850, nRows, Format(.Fields(i).Value, "#,##0.00")
    End If
    'nCols = nCols + 350
    End Select
    Next i
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 32
    Loop
    
    
End With

Set RS_REPORTES = Nothing
     DIM_SQLSUM = "select SUM(total),sum(isv) from InvSalida where fecha between DateValue('" & Format(DTPicker1, "Short Date") & "') AND DateValue('" & Format(DTPicker2, "Short Date") & "')"
       DIM_SQLSUM = DIM_SQLSUM & " AND forma Like '" & DIM_FORMA & "'"



Set RS_REPORTES = New Recordset
RS_REPORTES.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES
nRows = nRows + 32
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 1800, nRows, "0"
    Else
    DIMSUBTOTAL = .Fields(0).Value
    SpDoc.TextOut 1800, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If
nRows = nRows + 50
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL ISV 15%..."
    If IsNull(.Fields(1).Value) Then
    SpDoc.TextOut 1800, nRows, "0"
    Else
    DIMISV = .Fields(1).Value
    SpDoc.TextOut 1800, nRows, Format(.Fields(1).Value, "#,##0.00")
    End If
nRows = nRows + 45
DIMSALDO = DIMSUBTOTAL + DIMISV
SpDoc.TextOut 269, nRows, "TOTAL ..."
    SpDoc.TextOut 1800, nRows, Format(DIMSALDO, "#,##0.00")
'    SpDoc.TextOut 1539, nRows, Format(.Fields(1).value, "#,##0.00")
'    SpDoc.TextOut 1839, nRows, Format(.Fields(2).value, "#,##0.00")

End With

Set RS_REPORTES = Nothing


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
      
End Sub

Private Sub Command10_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT
      DIM_TITULORPT = "TIPO POR DIA  " & DIM_DIA_DET
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

  DIM_SQL = "select * from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY TALLA ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "TALLA"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
DIM_SQL = "select TALLA from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
  'DIM_SQL = "select TALLA,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
DIM_SQL = DIM_SQL & " GROUP BY TALLA ORDER BY TALLA ASC "
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("TALLA").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("TALLA").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQLSEL = "select TALLA,producto,salida,forma,total from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
                                      'DIM_SQL = "select TALLA,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
                                    DIM_SQLSEL = DIM_SQLSEL & " AND TALLA like '%" & .Fields("TALLA").Value & "%'"
                                    'DIM_SQLSEL = DIM_SQLSEL & " GROUP BY TALLA "
                                    'DIM_SQL = DIM_SQL & " GROUP BY TALLA ORDER BY TALLA ASC "
                                    RS_REPORTES1.Open DIM_SQLSEL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "TALLA"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    
DIM_SQLSUM = "select SUM(total) from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
DIM_SQLSUM = DIM_SQLSUM & " AND TALLA like '%" & .Fields("TALLA").Value & "%'"
Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command11_Click()
SpDoc.DocClearPage
SpDoc.DocBegin

     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = Command11.Caption
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY TALLA ASC "
       DIM_SQL = "select * from Ventas_Mes where mes like '" & DIM_MES & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "TALLA"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select TALLA from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
  'DIM_SQL = "select TALLA,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
      DIM_SQL = "select TALLA from Ventas_Mes where mes like '" & DIM_MES & "'"
      DIM_SQL = DIM_SQL & " GROUP BY TALLA ORDER BY TALLA ASC "
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("TALLA").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("TALLA").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQLSEL = "select TALLA,producto,salida,forma,VALOR from Ventas_Mes where mes like '" & DIM_MES & "'"

                                      'DIM_SQL = "select TALLA,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
                                    DIM_SQLSEL = DIM_SQLSEL & " AND TALLA like '%" & .Fields("TALLA").Value & "%'"
                                    
                                         'DIM_SQL = "select TALLA,producto,salida,forma,total from Ventas_Mes where mes like '" & DIM_MES & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQLSEL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "TALLA"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"
DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"
DIM_SQLSUM = DIM_SQLSUM & " AND TALLA like '%" & .Fields("TALLA").Value & "%'"

 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command12_Click()
SpDoc.DocClearPage
SpDoc.DocBegin

     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = Command12.Caption
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY TALLA ASC "
       DIM_SQL = "select * from Ventas_año where año like '" & DIM_AÑO & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_año where mes like '" & DIM_año & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "TALLA"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select TALLA from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
  'DIM_SQL = "select TALLA,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
      DIM_SQL = "select TALLA from Ventas_año where año like '" & DIM_AÑO & "'"
      DIM_SQL = DIM_SQL & " GROUP BY TALLA ORDER BY TALLA ASC "
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_año where mes like '" & DIM_año & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("TALLA").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("TALLA").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQL = "select TALLA,producto,salida,forma,valor from Ventas_año where año like '" & DIM_AÑO & "'"

                                      'DIM_SQL = "select TALLA,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY TALLA ASC "
                                    DIM_SQL = DIM_SQL & " AND TALLA like '%" & .Fields("TALLA").Value & "%'"
                                    
                                         'DIM_SQL = "select TALLA,producto,salida,forma,total from Ventas_año where mes like '" & DIM_año & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "TALLA"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_año where mes like '" & DIM_año & "'"
DIM_SQLSUM = "select SUM(valor) from Ventas_año where año like '" & DIM_AÑO & "'"
DIM_SQLSUM = DIM_SQLSUM & " AND TALLA like '%" & .Fields("TALLA").Value & "%'"

 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command13_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT
      DIM_TITULORPT = "VENTAS POR VENDEDOR " & DIM_DIA_DET
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

  DIM_SQL = "select * from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY cliente ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "Cliente"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
DIM_SQL = "select vendedor from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
  'DIM_SQL = "select cliente,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
DIM_SQL = DIM_SQL & " GROUP BY vendedor ORDER BY vendedor ASC "
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("vendedor").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("vendedor").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQLSEL = "select cliente,producto,salida,forma,total from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
                                      'DIM_SQL = "select cliente,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
                                    DIM_SQLSEL = DIM_SQLSEL & " AND vendedor like '" & .Fields("vendedor").Value & "'"
                                    'DIM_SQLSEL = DIM_SQLSEL & " GROUP BY cliente "
                                    'DIM_SQL = DIM_SQL & " GROUP BY cliente ORDER BY cliente ASC "
                                    RS_REPORTES1.Open DIM_SQLSEL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "Cliente"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                            
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                            
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    
DIM_SQLSUM = "select SUM(total) from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
DIM_SQLSUM = DIM_SQLSUM & " AND vendedor like '" & .Fields("vendedor").Value & "'"
Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
End Sub


Private Sub Command14_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT
      DIM_TITULORPT = "VENTAS POR VENDEDOR " & DIM_DIA_DET
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

  DIM_SQL = "select * from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY cliente ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "Cliente"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
DIM_SQL = "select vendedor from InvSalida "
  'DIM_SQL = "select cliente,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
DIM_SQL = DIM_SQL & " GROUP BY vendedor ORDER BY vendedor ASC "
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("vendedor").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("vendedor").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQLSEL = "select cliente,producto,salida,forma,total from InvSalida where  vendedor like '" & .Fields("vendedor").Value & "'"
                                    'DIM_SQLSEL = DIM_SQLSEL & " GROUP BY cliente "
                                    'DIM_SQL = DIM_SQL & " GROUP BY cliente ORDER BY cliente ASC "
                                    RS_REPORTES1.Open DIM_SQLSEL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "Cliente"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                            
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                            
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    
DIM_SQLSUM = "select SUM(total) from InvSalida where  vendedor like '" & .Fields("vendedor").Value & "'"
Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
End Sub


Private Sub Command15_Click()
Dim DIMSALDO, DIMSUBTOTAL, DIMISV
  SpDoc.DocClearPage
SpDoc.DocBegin
      DIM_TITULORPT = "VENTAS POR UNIDADES  " & DIM_DIA_DET
      Dim DIMUNIDAD, DIMVALORP, DIMRESULTADOP
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   
SpDoc.SetFont "Arial", 35, SPFO_BOLD + SPFS_UNITS, 0
Set RS_REPORTES = New Recordset
     DIM_FORMA = "CONTADO"
  DIM_SQL = "select fecha,vendedor,producto,salida,VALOR from VENTASUNIDADV where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY ndventas ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
    SpDoc.TextOut 139, 210, "Fecha"
    SpDoc.TextOut 339, 210, "Vendedor"
    SpDoc.TextOut 600, 210, "Producto"
    SpDoc.TextOut 950, 210, "Cantidad"
    SpDoc.TextOut 1200, 210, "Valor"
    SpDoc.TextOut 1500, 210, "Porcentarje"
    


Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
  DIM_SQL = "select fecha,vendedor,producto,salida,VALOR from VENTASUNIDADv  where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY ndventas ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    num_fields = .Fields.Count
    For i = 0 To num_fields - 1
    Select Case i
    Case 0
    'SpDoc.TextOut 139, 210, "Fecha"
    'SpDoc.TextOut 269, 210, "Producto"
    'SpDoc.TextOut 700, 210, "Cantidad"
    'SpDoc.TextOut 900, 210, "Valor"
    'SpDoc.TextOut 1050, 210, "Cliente"
    'SpDoc.TextOut 1650, 210, "NDoc"
    'SpDoc.TextOut 1800, 210, "Forma"
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 119, nRows, "0"
    Else
    SpDoc.TextOut 119, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 1
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 350, nRows, "0"
    Else
    SpDoc.TextOut 350, nRows, .Fields(i).Value
    End If
    Case 2
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 550, nRows, "0"
    Else
    SpDoc.TextOut 550, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 3
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1000, nRows, "0"
    Else
    SpDoc.TextOut 1000, nRows, .Fields(i).Value
    DIMUNIDAD = .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 4
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1200, nRows, "0"
    Else
    SpDoc.TextOut 1200, nRows, Format(.Fields(i).Value, "#,##0.00")
    DIMVALORP = .Fields(i).Value
    
    DIMRESULTADOP = DIMUNIDAD / DIMVALORP
    SpDoc.TextOut 1500, nRows, Format(DIMRESULTADOP, "#,##0.000000")
    End If

    
    'nCols = nCols + 350
    End Select
    

    Next i
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 40
        DIMUNIDAD = ""
    DIMVALORP = ""
    DIMRESULTADOP = ""
    Loop
    
    
End With

Set RS_REPORTES = Nothing
     DIM_SQLSUM = "select SUM(SALIDA),sum(VALOR) from VENTASUNIDADv where fecha between DateValue('" & Format(DTPicker1, "Short Date") & "') AND DateValue('" & Format(DTPicker2, "Short Date") & "')"
      ' DIM_SQLSUM = DIM_SQLSUM & " AND forma Like '" & DIM_FORMA & "'"



Set RS_REPORTES = New Recordset
RS_REPORTES.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES
nRows = nRows + 50
    SpDoc.SetFont "Arial", 60, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL UNIDADES ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 1500, nRows, "0"
    Else
    DIMSUBTOTAL = .Fields(0).Value
    SpDoc.TextOut 1500, nRows, .Fields(0).Value
    End If
nRows = nRows + 70
    SpDoc.SetFont "Arial", 60, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL VALOR"
    If IsNull(.Fields(1).Value) Then
    SpDoc.TextOut 1500, nRows, "0"
    Else
    DIMISV = .Fields(1).Value
    SpDoc.TextOut 1500, nRows, Format(.Fields(1).Value, "#,##0.00")
    End If
nRows = nRows + 70
DIMSALDO = DIMSUBTOTAL / DIMISV
SpDoc.TextOut 269, nRows, "TOTAL ..."
    SpDoc.TextOut 1500, nRows, Format(DIMSALDO, "#,##0.0000")
'    SpDoc.TextOut 1539, nRows, Format(.Fields(1).value, "#,##0.00")
'    SpDoc.TextOut 1839, nRows, Format(.Fields(2).value, "#,##0.00")

End With

Set RS_REPORTES = Nothing


SpDoc.DoPrintPreview

End Sub

Private Sub Command2_Click()
 SpDoc.DocClearPage
SpDoc.DocBegin
      DIM_TITULORPT = "VENTAS POR MES CONTADO"
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
        SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = Command2.Caption
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset
     DIM_FORMA = "CONTADO"

  DIM_SQL = "select mes,producto,salida,VALOR,cliente,Talla,forma from Ventas_Mes where mes like '" & DIM_MES & "'"
  DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "'"
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
    SpDoc.TextOut 139, 210, "Fecha"
    SpDoc.TextOut 269, 210, "Producto"
    SpDoc.TextOut 700, 210, "Cantidad"
    SpDoc.TextOut 900, 210, "Valor"
    SpDoc.TextOut 1100, 210, "Cliente"
    SpDoc.TextOut 1550, 210, "NDoc"
    SpDoc.TextOut 1800, 210, "Forma"



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
  DIM_SQL = "select mes,producto,salida,VALOR,cliente,Talla,forma from Ventas_Mes where mes like '" & DIM_MES & "'"
  DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "'  "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    num_fields = .Fields.Count
    For i = 0 To num_fields - 1
    Select Case i
    Case 0
    'SpDoc.TextOut 139, 210, "Fecha"
    'SpDoc.TextOut 269, 210, "Producto"
    'SpDoc.TextOut 700, 210, "Cantidad"
    'SpDoc.TextOut 900, 210, "Valor"
    'SpDoc.TextOut 1050, 210, "Cliente"
    'SpDoc.TextOut 1650, 210, "NDoc"
    'SpDoc.TextOut 1800, 210, "Forma"
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 119, nRows, "0"
    Else
    SpDoc.TextOut 119, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 1
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 299, nRows, "0"
    Else
    SpDoc.TextOut 299, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 2
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 750, nRows, "0"
    Else
    SpDoc.TextOut 750, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 3
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 850, nRows, "0"
    Else
    SpDoc.TextOut 850, nRows, Format(.Fields(i).Value, "#,##0.00")
    End If
    'nCols = nCols + 350
    
    Case 4
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1000, nRows, "0"
    Else
    SpDoc.TextOut 1000, nRows, .Fields(i).Value
    End If
    
    Case 5
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 2000, nRows, "0"
    Else
    SpDoc.TextOut 2000, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 6
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 2200, nRows, "0"
    Else
    SpDoc.TextOut 2200, nRows, .Fields(i).Value
    End If
    Case 7
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 2400, nRows, "0"
    Else
    SpDoc.TextOut 2400, nRows, Format(.Fields(i).Value, "#,##0.00")
    End If
    'nCols = nCols + 350
    End Select
    Next i
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 32
    Loop
    
    
End With

Set RS_REPORTES = Nothing

  DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"
  DIM_SQLSUM = DIM_SQLSUM & " AND forma like '" & DIM_FORMA & "'  "


Set RS_REPORTES = New Recordset
RS_REPORTES.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES
nRows = nRows + 32
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 1800, nRows, "0"
    Else
    SpDoc.TextOut 1800, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If
nRows = nRows + 50

'    SpDoc.TextOut 1539, nRows, Format(.Fields(1).value, "#,##0.00")
'    SpDoc.TextOut 1839, nRows, Format(.Fields(2).value, "#,##0.00")

End With

Set RS_REPORTES = Nothing


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command23_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = Command23.Caption
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY VENDEDOR ASC "
       DIM_SQL = "select * from InvSalida where cliente like '" & Combo1.Text & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "VENDEDOR"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select VENDEDOR from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
  'DIM_SQL = "select VENDEDOR,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      DIM_SQL = "select CLIENTE from InvSalida where cliente like '" & Combo1.Text & "'"
      DIM_SQL = DIM_SQL & " GROUP BY CLIENTE ORDER BY CLIENTE ASC "
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("CLIENTE").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("CLIENTE").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQL = "select cliente,producto,salida,forma,total from Invsalida where cliente like '" & Combo1.Text & "'"

                                      'DIM_SQL = "select VENDEDOR,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
                                    'DIM_SQL = DIM_SQL & " AND VENDEDOR like '" & .Fields("VENDEDOR").value & "' "
                                    
                                         'DIM_SQL = "select VENDEDOR,producto,salida,forma,total from Ventas_Mes where mes like '" & DIM_MES & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "VENDEDOR"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"
DIM_SQLSUM = "select SUM(total) from InvSalida where cliente like '" & Combo1.Text & "'"
'DIM_SQLSUM = DIM_SQLSUM & " AND VENDEDOR like '" & .Fields("VENDEDOR").value & "' "

 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command24_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = Command24.Caption
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY VENDEDOR ASC "
       DIM_SQL = "select * from InvSalida where vendedor like '" & Combo2.Text & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "VENDEDOR"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select VENDEDOR from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
  'DIM_SQL = "select VENDEDOR,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      DIM_SQL = "select VENDEDOR from InvSalida where vendedor like '" & Combo2.Text & "'"
      DIM_SQL = DIM_SQL & " GROUP BY VENDEDOR "
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("VENDEDOR").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("VENDEDOR").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQL = "select cliente,producto,salida,NDVENTAS,forma,total,ID from Invsalida where vendedor like '" & Combo2.Text & "'"

                                      'DIM_SQL = "select VENDEDOR,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
                                    'DIM_SQL = DIM_SQL & " AND VENDEDOR like '" & .Fields("VENDEDOR").value & "' "
                                    
                                         'DIM_SQL = "select VENDEDOR,producto,salida,forma,total from Ventas_Mes where mes like '" & DIM_MES & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "VENDEDOR"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, "Cliente : " & RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        
                                        nRows = nRows + 32
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, "Producto : " & RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1019, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1019, nRows, "Cantidad : " & RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        
                                       If RS_REPORTES1.Fields("FORMA") = "CONTADO" Then
                                        
                                        Set RS_REPORTES4 = New Recordset
                                        DIM_SQL = "select NDVENTAS from InvSalida where ID like '" & RS_REPORTES1.Fields("ID") & "'"
                                        RS_REPORTES4.Open DIM_SQL, PUB_CONEXION_EASY
                                        SpDoc.TextOut 1419, nRows, "Doc : " & RS_REPORTES4.Fields(0)
                                        Set RS_REPORTES4 = Nothing

                                      Else
                                        
                                        
                                        Set RS_REPORTES4 = New Recordset
                                        DIM_SQL = "select NDVENTASC from InvSalida where ID like '" & RS_REPORTES1.Fields("ID") & "'"
                                        RS_REPORTES4.Open DIM_SQL, PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
                                        SpDoc.TextOut 1419, nRows, "Doc : " & RS_REPORTES4.Fields(0)
                                        Set RS_REPORTES4 = Nothing
                                        
                                        'nCols = nCols + 350
                                    End If
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1719, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1719, nRows, "Forma : " & RS_REPORTES1.Fields(i).Value
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, "Valor : Lps." & Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        'SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        'SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"
DIM_SQLSUM = "select SUM(total) from InvSalida where vendedor like '" & Combo2.Text & "'"
'DIM_SQLSUM = DIM_SQLSUM & " AND VENDEDOR like '" & .Fields("VENDEDOR").value & "' "

 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 1500, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2000, nRows, "0"
    Else
    SpDoc.TextOut 2000, nRows, "Lps. " & Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command3_Click()
Dim DIMSALDO, DIMSUBTOTAL, DIMISV
  SpDoc.DocClearPage
SpDoc.DocBegin
      DIM_TITULORPT = "VENTAS POR UNIDADES  " & DIM_DIA_DET
      Dim DIMUNIDAD, DIMVALORP, DIMRESULTADOP
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   
SpDoc.SetFont "Arial", 40, SPFO_BOLD + SPFS_UNITS, 0
Set RS_REPORTES = New Recordset
     DIM_FORMA = "CONTADO"
  DIM_SQL = "select fecha,producto,salida,VALOR from VENTASUNIDAD where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY ndventas ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
    SpDoc.TextOut 139, 210, "Fecha"
    SpDoc.TextOut 339, 210, "Producto"
    SpDoc.TextOut 1000, 210, "Cantidad"
    SpDoc.TextOut 1300, 210, "Valor"
    SpDoc.TextOut 1600, 210, "Porcentaje"
    


Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
  DIM_SQL = "select fecha,producto,salida,VALOR from VENTASUNIDAD  where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY ndventas ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    num_fields = .Fields.Count
    For i = 0 To num_fields - 1
    Select Case i
    Case 0
    'SpDoc.TextOut 139, 210, "Fecha"
    'SpDoc.TextOut 269, 210, "Producto"
    'SpDoc.TextOut 700, 210, "Cantidad"
    'SpDoc.TextOut 900, 210, "Valor"
    'SpDoc.TextOut 1050, 210, "Cliente"
    'SpDoc.TextOut 1650, 210, "NDoc"
    'SpDoc.TextOut 1800, 210, "Forma"
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 119, nRows, "0"
    Else
    SpDoc.TextOut 119, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 1
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 350, nRows, "0"
    Else
    SpDoc.TextOut 350, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 2
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1000, nRows, "0"
    Else
    SpDoc.TextOut 1000, nRows, .Fields(i).Value
    DIMUNIDAD = .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 3
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1200, nRows, "0"
    Else
    SpDoc.TextOut 1200, nRows, Format(.Fields(i).Value, "#,##0.00")
    DIMVALORP = .Fields(i).Value
    
    DIMRESULTADOP = DIMUNIDAD / DIMVALORP
    SpDoc.TextOut 1500, nRows, Format(DIMRESULTADOP, "#,##0.0000")
    End If

    
    'nCols = nCols + 350
    End Select
    

    Next i
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 40
        DIMUNIDAD = ""
    DIMVALORP = ""
    DIMRESULTADOP = ""
    Loop
    
    
End With

Set RS_REPORTES = Nothing
     DIM_SQLSUM = "select SUM(SALIDA),sum(VALOR) from VENTASUNIDAD where fecha between DateValue('" & Format(DTPicker1, "Short Date") & "') AND DateValue('" & Format(DTPicker2, "Short Date") & "')"
      ' DIM_SQLSUM = DIM_SQLSUM & " AND forma Like '" & DIM_FORMA & "'"



Set RS_REPORTES = New Recordset
RS_REPORTES.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES
nRows = nRows + 50
    SpDoc.SetFont "Arial", 60, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL UNIDADES ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 1500, nRows, "0"
    Else
    DIMSUBTOTAL = .Fields(0).Value
    SpDoc.TextOut 1500, nRows, .Fields(0).Value
    End If
nRows = nRows + 70
    SpDoc.SetFont "Arial", 60, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL VALOR"
    If IsNull(.Fields(1).Value) Then
    SpDoc.TextOut 1500, nRows, "0"
    Else
    DIMISV = .Fields(1).Value
    SpDoc.TextOut 1500, nRows, Format(.Fields(1).Value, "#,##0.00")
    End If
nRows = nRows + 70
DIMSALDO = DIMSUBTOTAL / DIMISV
SpDoc.TextOut 269, nRows, "TOTAL ..."
    SpDoc.TextOut 1500, nRows, Format(DIMSALDO, "#,##0.00")
'    SpDoc.TextOut 1539, nRows, Format(.Fields(1).value, "#,##0.00")
'    SpDoc.TextOut 1839, nRows, Format(.Fields(2).value, "#,##0.00")

End With

Set RS_REPORTES = Nothing


SpDoc.DoPrintPreview
End Sub

Private Sub Command4_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
          DIM_TITULORPT = "VENTAS POR DIA CREDITO  " & DIM_DIA_DET
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT
     DIM_FORMA = "CREDITO"
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

  DIM_SQL = "select fecha,producto,salida,total,cliente,ndventasc,Talla,forma from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY ndventasc ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
    SpDoc.TextOut 139, 210, "Fecha"
    SpDoc.TextOut 269, 210, "Producto"
    SpDoc.TextOut 700, 210, "Cantidad"
    SpDoc.TextOut 900, 210, "Valor"
    SpDoc.TextOut 1100, 210, "Cliente"
    SpDoc.TextOut 1550, 210, "NDoc"
    SpDoc.TextOut 1800, 210, "Forma"



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
  DIM_SQL = "select fecha,producto,salida,total,cliente,ndventasc,Talla,forma from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY ndventasc ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    num_fields = .Fields.Count
    For i = 0 To num_fields - 1
    Select Case i
    Case 0
    'SpDoc.TextOut 139, 210, "Fecha"
    'SpDoc.TextOut 269, 210, "Producto"
    'SpDoc.TextOut 700, 210, "Cantidad"
    'SpDoc.TextOut 900, 210, "Valor"
    'SpDoc.TextOut 1050, 210, "Cliente"
    'SpDoc.TextOut 1650, 210, "NDoc"
    'SpDoc.TextOut 1800, 210, "Forma"
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 119, nRows, "0"
    Else
    SpDoc.TextOut 119, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 1
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 299, nRows, "0"
    Else
    SpDoc.TextOut 299, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 2
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 750, nRows, "0"
    Else
    SpDoc.TextOut 750, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 3
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 850, nRows, "0"
    Else
    SpDoc.TextOut 850, nRows, Format(.Fields(i).Value, "#,##0.00")
    End If
    'nCols = nCols + 350
    
    Case 4
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1000, nRows, "0"
    Else
    SpDoc.TextOut 1000, nRows, .Fields(i).Value
    End If
    
    Case 5
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 2000, nRows, "0"
    Else
    SpDoc.TextOut 2000, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 6
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 2200, nRows, "0"
    Else
    SpDoc.TextOut 2200, nRows, .Fields(i).Value
    End If
    Case 7
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 2400, nRows, "0"
    Else
    SpDoc.TextOut 2400, nRows, Format(.Fields(i).Value, "#,##0.00")
    End If
    'nCols = nCols + 350
    End Select
    Next i
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 32
    Loop
    
    
End With

Set RS_REPORTES = Nothing
     DIM_SQLSUM = "select SUM(total),sum(isv) from InvSalida where fecha between DateValue('" & Format(DTPicker1, "Short Date") & "') AND DateValue('" & Format(DTPicker2, "Short Date") & "')"
       DIM_SQLSUM = DIM_SQLSUM & " AND forma Like '" & DIM_FORMA & "'"



Set RS_REPORTES = New Recordset
RS_REPORTES.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES
nRows = nRows + 32
nRows = nRows + 32
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 1800, nRows, "0"
    Else
    DIMSUBTOTAL = .Fields(0).Value
    SpDoc.TextOut 1800, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If
nRows = nRows + 50
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL ISV 15%..."
    If IsNull(.Fields(1).Value) Then
    SpDoc.TextOut 1800, nRows, "0"
    Else
    DIMISV = .Fields(1).Value
    SpDoc.TextOut 1800, nRows, Format(.Fields(1).Value, "#,##0.00")
    End If
nRows = nRows + 45
DIMSALDO = DIMSUBTOTAL + DIMISV
SpDoc.TextOut 269, nRows, "TOTAL ..."
    SpDoc.TextOut 1800, nRows, Format(DIMSALDO, "#,##0.00")

End With

Set RS_REPORTES = Nothing


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub


Private Sub Command5_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
      DIM_TITULORPT = "VENTAS POR MES CREDITO"
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
        SpDoc.PageOrientation = SPOR_LANDSCAPE
            RptTitle = Command5.Caption
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset
     DIM_FORMA = "CREDITO"

  DIM_SQL = "select mes,producto,salida,valor,cliente,Talla,forma from Ventas_Mes where mes like '" & DIM_MES & "'"
  DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
    SpDoc.TextOut 139, 210, "Fecha"
    SpDoc.TextOut 269, 210, "Producto"
    SpDoc.TextOut 700, 210, "Cantidad"
    SpDoc.TextOut 900, 210, "Valor"
    SpDoc.TextOut 1100, 210, "Cliente"
    SpDoc.TextOut 1550, 210, "NDoc"
    SpDoc.TextOut 1800, 210, "Forma"



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
  DIM_SQL = "select mes,producto,salida,valor,Talla,forma from Ventas_Mes where mes like '" & DIM_MES & "'"
  DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    num_fields = .Fields.Count
    For i = 0 To num_fields - 1
    Select Case i
    Case 0
    'SpDoc.TextOut 139, 210, "Fecha"
    'SpDoc.TextOut 269, 210, "Producto"
    'SpDoc.TextOut 700, 210, "Cantidad"
    'SpDoc.TextOut 900, 210, "Valor"
    'SpDoc.TextOut 1050, 210, "Cliente"
    'SpDoc.TextOut 1650, 210, "NDoc"
    'SpDoc.TextOut 1800, 210, "Forma"
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 119, nRows, "0"
    Else
    SpDoc.TextOut 119, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 1
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 299, nRows, "0"
    Else
    SpDoc.TextOut 299, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 2
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 750, nRows, "0"
    Else
    SpDoc.TextOut 750, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 3
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 850, nRows, "0"
    Else
    SpDoc.TextOut 850, nRows, Format(.Fields(i).Value, "#,##0.00")
    End If
    'nCols = nCols + 350
    
    Case 4
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1000, nRows, "0"
    Else
    SpDoc.TextOut 1000, nRows, .Fields(i).Value
    End If
    
    Case 5
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1600, nRows, "0"
    Else
    SpDoc.TextOut 1600, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 6
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1700, nRows, "0"
    Else
    SpDoc.TextOut 1700, nRows, .Fields(i).Value
    End If
    Case 7
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1850, nRows, "0"
    Else
    SpDoc.TextOut 1850, nRows, Format(.Fields(i).Value, "#,##0.00")
    End If
    'nCols = nCols + 350
    End Select
    Next i
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 32
    Loop
    
    
End With

Set RS_REPORTES = Nothing

  DIM_SQLSUM = "select SUM(valor)  from Ventas_Mes where mes like '" & DIM_MES & "'"
  DIM_SQLSUM = DIM_SQLSUM & " AND forma like '" & DIM_FORMA & "'"


Set RS_REPORTES = New Recordset
RS_REPORTES.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES
nRows = nRows + 32
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 1800, nRows, "0"
    Else
    SpDoc.TextOut 1800, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If
nRows = nRows + 50

'    SpDoc.TextOut 1539, nRows, Format(.Fields(1).value, "#,##0.00")
'    SpDoc.TextOut 1839, nRows, Format(.Fields(2).value, "#,##0.00")

End With

Set RS_REPORTES = Nothing


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command6_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
      DIM_TITULORPT = "VENTAS POR CIUDAD"
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
        'SpDoc.PageOrientation = SPOR_LANDSCAPE
            RptTitle = Command6.Caption
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset
     DIM_FORMA = "CREDITO"

  DIM_SQL = "select ciudad,producto,unidades,valor from ClientesCiudad1 "
  'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "'  "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
    SpDoc.TextOut 139, 210, "Ciudad"
    SpDoc.TextOut 700, 210, "Producto"
    SpDoc.TextOut 1300, 210, "Cantidad"
    SpDoc.TextOut 1500, 210, "Valor"




Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
  DIM_SQL = "select ciudad,producto,unidades,valor from ClientesCiudad1 "
  'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "'  "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    num_fields = .Fields.Count
    For i = 0 To num_fields - 1
    Select Case i
    Case 0
    'SpDoc.TextOut 139, 210, "Fecha"
    'SpDoc.TextOut 269, 210, "Producto"
    'SpDoc.TextOut 700, 210, "Cantidad"
    'SpDoc.TextOut 900, 210, "Valor"
    'SpDoc.TextOut 1050, 210, "Cliente"
    'SpDoc.TextOut 1650, 210, "NDoc"
    'SpDoc.TextOut 1800, 210, "Forma"
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 119, nRows, "0"
    Else
    SpDoc.TextOut 119, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 1
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 700, nRows, "0"
    Else
    SpDoc.TextOut 700, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 2
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1300, nRows, "0"
    Else
    SpDoc.TextOut 1300, nRows, .Fields(i).Value
    End If
    'nCols = nCols + 350
    Case 3
    If IsNull(.Fields(i).Value) Then
    SpDoc.TextOut 1500, nRows, "0"
    Else
    SpDoc.TextOut 1500, nRows, Format(.Fields(i).Value, "#,##0.00")
    End If
    'nCols = nCols + 350
    
    
    'nCols = nCols + 350
    End Select
    Next i
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 32
    Loop
    
    
End With

Set RS_REPORTES = Nothing

  DIM_SQLSUM = "select SUM(UNIDADES),SUM(valor)   from ClientesCiudad1 "
  'DIM_SQLSUM = DIM_SQLSUM & " AND forma like '" & DIM_FORMA & "'  "


Set RS_REPORTES = New Recordset
RS_REPORTES.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES
nRows = nRows + 50
    SpDoc.SetFont "Arial", 60, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL UNIDADES ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 1500, nRows, "0"
    Else
    DIMSUBTOTAL = .Fields(0).Value
    SpDoc.TextOut 1500, nRows, .Fields(0).Value
    End If
nRows = nRows + 70
    SpDoc.SetFont "Arial", 60, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 269, nRows, "TOTAL VALOR"
    If IsNull(.Fields(1).Value) Then
    SpDoc.TextOut 1500, nRows, "0"
    Else
    DIMISV = .Fields(1).Value
    SpDoc.TextOut 1500, nRows, Format(.Fields(1).Value, "#,##0.00")
    End If

'    SpDoc.TextOut 1539, nRows, Format(.Fields(1).value, "#,##0.00")
'    SpDoc.TextOut 1839, nRows, Format(.Fields(2).value, "#,##0.00")

End With

Set RS_REPORTES = Nothing


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command7_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT
      DIM_TITULORPT = "VENTAS POR DIA CLIENTES  " & DIM_DIA_DET
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = DIM_TITULORPT
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

  DIM_SQL = "select * from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY cliente ASC "
 
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "Cliente"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
DIM_SQL = "select cliente from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
  'DIM_SQL = "select cliente,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
DIM_SQL = DIM_SQL & " GROUP BY cliente ORDER BY cliente ASC "
RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("Cliente").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("Cliente").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQLSEL = "select cliente,producto,salida,forma,total from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
                                      'DIM_SQL = "select cliente,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
                                    DIM_SQLSEL = DIM_SQLSEL & " AND cliente like '" & .Fields("cliente").Value & "'"
                                    'DIM_SQLSEL = DIM_SQLSEL & " GROUP BY cliente "
                                    'DIM_SQL = DIM_SQL & " GROUP BY cliente ORDER BY cliente ASC "
                                    RS_REPORTES1.Open DIM_SQLSEL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "Cliente"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    
DIM_SQLSUM = "select SUM(total) from InvSalida where fecha between #" & Format(DTPicker1, "mm/dd/yyyy") & "# and #" & Format(DTPicker2, "mm/dd/yyyy") & "#"
DIM_SQLSUM = DIM_SQLSUM & " AND cliente like '" & .Fields("cliente").Value & "'"
Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command8_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = Command8.Caption
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY cliente ASC "
       DIM_SQL = "select * from Ventas_Mes where mes like '" & DIM_MES & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "Cliente"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select cliente from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
  'DIM_SQL = "select cliente,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
      DIM_SQL = "select cliente from Ventas_Mes where mes like '" & DIM_MES & "'"
      DIM_SQL = DIM_SQL & " GROUP BY cliente ORDER BY cliente ASC "
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("Cliente").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("Cliente").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQLSEL = "select cliente,producto,salida,forma,valor from Ventas_Mes where mes like '" & DIM_MES & "'"

                                      'DIM_SQL = "select cliente,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY cliente ASC "
                                    DIM_SQLSEL = DIM_SQLSEL & " AND cliente like '" & .Fields("cliente").Value & "'"
                                    
                                         'DIM_SQL = "select cliente,producto,salida,forma,total from Ventas_Mes where mes like '" & DIM_MES & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQLSEL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "Cliente"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1100, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 2
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"
DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"
DIM_SQLSUM = DIM_SQLSUM & " AND cliente like '" & .Fields("cliente").Value & "'"

 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQLSUM, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 2000, nRows, "TOTAL REPORTE ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2500, nRows, "0"
    Else
    SpDoc.TextOut 2500, nRows, Format(.Fields(0).Value, "#,##0.00")
    End If

End With

Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview
    
Exit Sub
menerr:
Open App.Path & "\ERRORES\MISERRORES.txt" For Append As #1
Print #1, "FRMREPORTE,reporteie"
Close #1
End Sub

Private Sub Command9_Click()
SpDoc.DocClearPage
SpDoc.DocBegin
DIM_FRMCREDITO = "CREDITO"
Dim DIMCredito, DIMAbono, DIMSALDO
     'spDoc.PageOrientation = cboOrientation.ListIndex + SPOR_PORTRAIT

      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

'On Error GoTo menerr
Dim nRows As Long, nCols As Long, nItem As Long
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long

    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    

    nFooterTop = Bottom - 380
   
    SpDoc.PageOrientation = SPOR_LANDSCAPE
    RptTitle = "COMPRAS POR CLIENTE"
    DIMTITULOPAGINA = RptTitle
 PrintEncabezado
   
   

Set RS_REPORTES = New Recordset

 ' DIM_SQL = "select * from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
      'DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' ORDER BY VENDEDOR ASC "
       DIM_SQL = "select * from Ventas_Mes where mes like '" & DIM_MES & "'"
      'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY
 'ReporteIE DIM_SQL, DIM_SQLSUM, DIM_TITULORPT
If RS_REPORTES.RecordCount = 0 Then
MsgBox "NO HAY DATOS PARA MOSTRAR"
SpDoc.DocClearPage
Exit Sub
End If


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nCols = 139
  '  SpDoc.TextOut 139, 210, "VENDEDOR"
  '  SpDoc.TextOut 269, 210, "Producto"
  '  SpDoc.TextOut 700, 210, "Cantidad"
'    SpDoc.TextOut 900, 210, "NoDoc"
  '  SpDoc.TextOut 1100, 210, "Forma"
  '  SpDoc.TextOut 1550, 210, "Valor"

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''



Set RS_REPORTES = Nothing

Set RS_REPORTES = New Recordset
'DIM_SQL = "select VENDEDOR from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
  'DIM_SQL = "select VENDEDOR,producto,total,Talla,forma from InvSalida where fecha between DateValue('" & Format(Fecha_Inicial, "Short Date") & "') AND DateValue('" & Format(Fecha_Final, "Short Date") & "') ORDER BY VENDEDOR ASC "
   '   DIM_SQL = "select NODE from Ventas_Mes where mes like '" & DIM_MES & "'"
   '   DIM_SQL = DIM_SQL & " GROUP BY NODE ORDER BY NODE ASC "
      
      
      DIM_SQL = "select nombre from ClientesDts where forma like '" & DIM_FRMCREDITO & "'"
      DIM_SQL = DIM_SQL & " GROUP BY nombre"
     ' DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"

RS_REPORTES.Open DIM_SQL, PUB_CONEXION_EASY

With RS_REPORTES

                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
    nRows = 270
    nCols = 139
    Do While Not .EOF
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0


                If IsNull(.Fields("Nombre").Value) Then
                SpDoc.TextOut 119, nRows, "0"
                Else
                SpDoc.TextOut 119, nRows, .Fields("Nombre").Value
                End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    nRows = nRows + 75
    SpDoc.SetFont "Arial", 30, SPFO_BOLD + SPFS_UNITS, 0
                                    
                                    Set RS_REPORTES1 = Nothing
                                    
                                    Set RS_REPORTES1 = New Recordset
                                    DIM_SQL = "select sum(cantidad),sum(valor) from ClientesDts where nombre like '" & .Fields("nombre").Value & "' "
                                    DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FRMCREDITO & "'"
                                         'DIM_SQL = "select VENDEDOR,producto,salida,forma,total from Ventas_Mes where mes like '" & DIM_MES & "'"
      '
 
                                    
                                    RS_REPORTES1.Open DIM_SQL, PUB_CONEXION_EASY
                                    

                                    

                                        Do While Not RS_REPORTES1.EOF
                                        num_fields = RS_REPORTES1.Fields.Count
                                        For i = 0 To num_fields - 1
                                        Select Case i
                                        Case 0
                                        'SpDoc.TextOut 139, 210, "Fecha"
                                        'SpDoc.TextOut 269, 210, "Producto"
                                        'SpDoc.TextOut 700, 210, "Cantidad"
                                        'SpDoc.TextOut 900, 210, "Valor"
                                        'SpDoc.TextOut 1050, 210, "VENDEDOR"
                                        'SpDoc.TextOut 1650, 210, "NDoc"
                                        'SpDoc.TextOut 1800, 210, "Forma"
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 119, nRows, "0"
                                        Else
                                        SpDoc.TextOut 119, nRows, "Cantidad : " & RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 1
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1100, nRows, "0"
                                        Else
                                        DIMCredito = RS_REPORTES1.Fields(i).Value
                                        SpDoc.TextOut 1100, nRows, "Compras por Cliente : " & Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        Case 2

                                        
                                        'nCols = nCols + 350
                                        Case 3
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2000, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2000, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        
                                        Case 4
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 2500, nRows, "0"
                                        Else
                                        SpDoc.TextOut 2500, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        
                                        Case 5
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1600, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1600, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        'nCols = nCols + 350
                                        Case 6
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1700, nRows, RS_REPORTES1.Fields(i).Value
                                        End If
                                        Case 7
                                        If IsNull(RS_REPORTES1.Fields(i).Value) Then
                                        SpDoc.TextOut 1850, nRows, "0"
                                        Else
                                        SpDoc.TextOut 1850, nRows, Format(RS_REPORTES1.Fields(i).Value, "#,##0.00")
                                        End If
                                        'nCols = nCols + 350
                                        End Select
                                        Next i
                                        
                                        DIM_FORMA = "ABONO"
                                        Set RS_REPORTES4 = New Recordset
                                        DIM_SQL = "select SUM(VALOR) from ClientesDts where nombre like '" & .Fields("nombre").Value & "' "
                                        DIM_SQL = DIM_SQL & " AND forma like '" & DIM_FORMA & "' "
                                        RS_REPORTES4.Open DIM_SQL, PUB_CONEXION_EASY
                                        'SpDoc.TextOut 2000, nRows, "Abono : " & RS_REPORTES4.Fields(0)
                                        If IsNull(RS_REPORTES4.Fields(0)) Then
                                        SpDoc.TextOut 1700, nRows, "0"
                                        Else
                                        DIMAbono = RS_REPORTES4.Fields(0)
                                        SpDoc.TextOut 1700, nRows, "Abono : Lps." & Format(RS_REPORTES4.Fields(0), "#,##0.00")
                                        End If
                                        Set RS_REPORTES4 = Nothing
                                        DIMSALDO = DIMCredito - DIMAbono
                                        SpDoc.TextOut 2000, nRows, "Saldo : Lps." & Format(DIMSALDO, "#,##0.00")
                                        
                                        RS_REPORTES1.MoveNext
                                        
                                        nRows = nRows + 32
                                        Loop
                                        
                                        

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''

    'DIM_SQLSUM = "select SUM(valor) from Ventas_Mes where mes like '" & DIM_MES & "'"


Set RS_REPORTES2 = Nothing
    
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
                                                            DimPie = nFooterTop
                                                            If DimPie <= nRows Then
                                                            SpDoc.Page = SpDoc.Page + 1
                                                            PrintPageNumber
                                                            PrintEncabezado
                                                            nRows = 270
                                                            End If
   
    .MoveNext
    
                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
                        
    nRows = nRows + 50
    
    


    
Loop

DIM_SQL = "select SUM(valor) from ClientesDts where forma like '" & DIM_FRMCREDITO & "' "
 

Set RS_REPORTES2 = New Recordset
RS_REPORTES2.Open DIM_SQL, PUB_CONEXION_EASY


                        DimPie = nFooterTop
                        If DimPie <= nRows Then
                        SpDoc.Page = SpDoc.Page + 1
                        PrintPageNumber
                        PrintEncabezado
                        nRows = 270
                        End If
With RS_REPORTES2
nRows = nRows + 50
    SpDoc.SetFont "Arial", 45, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.TextOut 1500, nRows, "TOTAL COMPRADO ....."
    If IsNull(.Fields(0).Value) Then
    SpDoc.TextOut 2000, nRows, "0"
    Else
    SpDoc.TextOut 2000, nRows, "Lps. " & Format(.Fields(0).Value, "#,##0.00")
    End If

End With
    
    
End With

Set RS_REPORTES = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''


SpDoc.DoPrintPreview


End Sub


Private Sub PrintEncabezado()
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim X As Long, Y As Long, nIdx As Long
    Dim center As Long, lMaxY As Long
    Dim strText As String, CharsDrawn As Long
    Dim TextHeight As Long, TextWidth As Long
    
    'set up the page in preparation
    SpDoc.Units = SPUN_LOMETRIC
    'PrintPageOutline
    
    'get the printable space on the page, then set the margins to 10mm or
    'the printable area whichever is greatest
    SpDoc.GetPrintableArea Left, Top, Right, Bottom
    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)
    center = Left + ((Right - Left) / 2)
    
    '----------------------------------------------------------------
    'draw the SwiftPrint title
    '----------------------------------------------------------------
    Y = 300
    SpDoc.SetPen SPPN_NULL, 0, 0
    SpDoc.SetBrush SPBR_SOLID, RGB(232, 232, 255)
    SpDoc.Rectangle Left, Top, Right, Top + Y / 2
    
    SpDoc.SetFont "Arial", 45, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.BackMode = SPBM_TRANSPARENT
    SpDoc.TextAlign = SPTA_TOP + SPTA_CENTER + SPTA_NOUPDATECP
    SpDoc.TextOut center, Top, DIM_EMPRESA
    
    SpDoc.SetFont "Arial", 45, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.BackMode = SPBM_TRANSPARENT
    SpDoc.TextAlign = SPTA_TOP + SPTA_CENTER + SPTA_NOUPDATECP
    SpDoc.TextOut center, Top + 70, DIMTITULOPAGINA
    'draw the standard page title
    'DrawTitle Left, Top, Right, Bottom, False
    
    '----------------------------------------------------------------
    'draw page number at top of page
    '----------------------------------------------------------------
    SpDoc.SetFont "Arial", 50, SPFO_BOLD + SPFS_UNITS, 0
    SpDoc.ForeColor = RGB(0, 0, 0)
    SpDoc.BackMode = SPBM_TRANSPARENT
    SpDoc.TextAlign = SPTA_TOP + SPTA_RIGHT + SPTA_NOUPDATECP
    SpDoc.TextOut Right - 10, Top + 10, "Page: " & SpDoc.Page
    
    SpDoc.SetPen SPPN_SOLID, 0, RGB(0, 0, 0)
    SpDoc.SetBrush SPBR_NULL, 0
    SpDoc.Rectangle Left, Top, Right, 200
    
    SpDoc.SetPen SPPN_SOLID, 0, RGB(0, 0, 0)
    SpDoc.SetBrush SPBR_NULL, 0
    SpDoc.Rectangle Left, Top, Right, 250
    'SpDoc.Rectangle Left, Top, Right, Top + y
    'draw the all around rectangle
    SpDoc.SetPen SPPN_SOLID, 0, RGB(0, 0, 0)
    SpDoc.SetBrush SPBR_NULL, 0
    SpDoc.Rectangle Left, Top, Right, Bottom
    
    'draw the vertical copyright statement
 '   strText = "Easy Accounting 2010 & SwiftPrint"
 '   SpDoc.SetFont "Arial", 65, SPFS_POINTS, 900
 '   SpDoc.ForeColor = RGB(0, 0, 0)
 '   SpDoc.BackMode = SPBM_TRANSPARENT
 '   SpDoc.TextAlign = SPTA_BOTTOM + SPTA_RIGHT + SPTA_NOUPDATECP
 '   SpDoc.TextOut Right, Bottom, strText
    
    'reset the y and top values
    Y = Top + 300
    Top = Y
    
    'adjust the margins now
    Top = Top + 20
    Left = Left + 20
    Right = Right - 20
    Bottom = Bottom - 20
    
    '----------------------------------------------------------------
    'draw the page footer
    '----------------------------------------------------------------

    nFooterTop = Bottom - 380
    
    'draw the purple rectangle
 '   SpDoc.SetPen SPPN_SOLID, 0, RGB(0, 0, 0)
 ''   SpDoc.SetBrush SPBR_SOLID, RGB(232, 232, 255)
 '   SpDoc.RoundRect Left, nFooterTop + 250, Right, Bottom, 75, 75
    
    SpDoc.SetFont "Arial", 80, SPFS_POINTS, 0
    SpDoc.ForeColor = RGB(0, 0, 0)
    SpDoc.TextAlign = SPTA_TOP + SPTA_LEFT + SPTA_NOUPDATECP
     
 '   strText = "Direccion: " & vbTab & DIM_DIRECCION & vbCrLf & "Email:" & vbTab & "brodie@iname.com"
 '   SpDoc.TextOutEx strText, Left + 160, nFooterTop + 280, center, Bottom - 20, SPTO_LEFT + SPTO_VCENTER + SPTO_WORDBREAK, 8, CharsDrawn
 '   strText = "Telefono:" & vbTab & DIM_TELEFONO & vbCrLf & "Tienda :" & vbTab & DIM_TIENDA
 '   SpDoc.TextOutEx strText, center + 120, nFooterTop + 280, Right - 20, Bottom - 20, SPTO_LEFT + SPTO_VCENTER + SPTO_WORDBREAK, 8, CharsDrawn
        
    'adjust the margins now
    Top = Top + 20
    Left = Left + 20
    Right = Right - 20
    Bottom = Bottom - 20
End Sub


Private Function Max(ByVal L1 As Long, ByVal L2 As Long) As Long
    Max = IIf(L1 > L2, L1, L2)
End Function

Private Function Min(ByVal L1 As Long, ByVal L2 As Long) As Long
    Min = IIf(L1 < L2, L1, L2)
End Function
Private Sub PrintPageNumber()
    Dim Left As Long, Right As Long, Top As Long, Bottom As Long
    
    'get the printable space on the page, then set the margins to 10mm or
    'the printable area whichever is greatest
    SpDoc.GetPrintableArea Left, Top, Right, Bottom
    Left = Max(Left, 60)
    Top = Max(Top, 60)
    Right = Min(Right - 30, SpDoc.PageWidth - 60)
    Bottom = Min(Bottom, SpDoc.PageHeight - 60)

    SpDoc.SetFont "Arial", 80, SPFS_POINTS, 0
    SpDoc.TextAlign = SPTA_RIGHT + SPTA_TOP + SPTA_NOUPDATECP
    SpDoc.TextOut Right - 10, Top + 10, "Page " & SpDoc.Page & " of " & SpDoc.NumPages
End Sub

