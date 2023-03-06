VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ListIVAVentas 
   Caption         =   "Listado de IVA [Ventas]"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   10245
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   5175
      Begin VB.TextBox txtTotNeto 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   13
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotIVA 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotTotal 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotBruto 
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotDescuentos 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Neto"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Bruto"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Descuentos"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   16
      Top             =   5400
      Width           =   1815
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         Picture         =   "ListIVAVentas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   " Salir "
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Picture         =   "ListIVAVentas.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   " Imprimir "
         Top             =   0
         Width           =   855
      End
   End
   Begin MSComCtl2.DTPicker DTFechaD 
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   99221505
      CurrentDate     =   39603
   End
   Begin MSComCtl2.DTPicker DTFechaH 
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   99221505
      CurrentDate     =   39603
   End
   Begin VB.CommandButton cmdBuscar 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Picture         =   "ListIVAVentas.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Aplicar"
      Top             =   360
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   4335
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7646
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      FixedCols       =   0
      BackColorFixed  =   12640511
      ForeColorFixed  =   0
      ForeColorSel    =   16777215
      BackColorBkg    =   8421504
      BackColorUnpopulated=   12632256
      GridColor       =   0
      GridColorFixed  =   0
      GridColorUnpopulated=   8421504
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "ListIVAVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub OrdenaFlex()
    
'    Flex.FormatString = "id|Fecha|Tipo|PtoVenta|Número|Cliente|CUIT|Neto|IVA 10,5%|IVA 21%|IIBB|Total"
'    Flex.ColWidth(0) = 1000
'    Flex.ColWidth(1) = 1000
'    Flex.ColWidth(2) = 1000
'    Flex.ColWidth(3) = 1000
'    Flex.ColWidth(4) = 1000
'    Flex.ColWidth(5) = 1000
'    Flex.ColWidth(6) = 1000
'    Flex.ColWidth(7) = 1000
'    Flex.ColWidth(8) = 1000
'    Flex.ColWidth(9) = 1000
'    Flex.ColWidth(10) = 1000
'    Flex.ColWidth(11) = 1000
'    Flex.ColWidth(12) = 1000
'    Flex.ColWidth(13) = 1000
    
End Sub

Sub CalcFechas()
    
    'Calcula el priemr y último día del mes anterior
    Mes = Format(Date, "m") - 1
    Año = Format(Date, "yyyy")
    
    DTFechaD.Value = DateSerial(Año, Mes + 0, 1)
    DTFechaH.Value = DateSerial(Año, Mes + 1, 0)
    
End Sub

Public Sub cmdBuscar_Click()
    
    Flex.Clear
    Flex.Rows = 2
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT ventas.id, ventas.fecha, tipocomprobante.nombre, ventas.ptoventa, ventas.numerocomprobante, clientes.nombre AS cliente, clientes.numerodocumento, ventas.totalneto, "
    SQL = SQL & "REPLACE((SELECT SUM(CAST(REPLACE(ivaimp,',','.') AS DECIMAL(10,2))) FROM ventasd WHERE idventa = ventas.id AND ventasd.idtipoiva = 4),'.',',') AS 'IVA 10,5%', "
    SQL = SQL & "REPLACE((SELECT SUM(CAST(REPLACE(ivaimp,',','.') AS DECIMAL(10,2))) FROM ventasd WHERE idventa = ventas.id AND ventasd.idtipoiva = 5),'.',',') AS 'IVA 21%', "
    SQL = SQL & "ventas.totaltributos, ventas.Total "
    SQL = SQL & "From VENTAS "
    SQL = SQL & "Inner Join tipocomprobante ON tipocomprobante.id = ventas.idtipocomprobante "
    SQL = SQL & "Inner Join clientes ON clientes.id = ventas.idcliente"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        Set Flex.DataSource = Recordset
    End If
    
    OrdenaFlex
    
End Sub

Private Sub cboProvincia_Click()
    
    cboIdProvincia.ListIndex = cboProvincia.ListIndex
    
End Sub

Private Sub cmdImprimir_Click()
    
    'ActualizarDR
    
    drIVAVentas.Sections("ReportHeader").Controls("rptDesde").Caption = DTFechaD.Value
    drIVAVentas.Sections("ReportHeader").Controls("rptHasta").Caption = DTFechaH.Value
    drIVAVentas.Sections("ReportFooter").Controls("rptNeto").Caption = FormatCurrency(txtTotNeto.Text)
    drIVAVentas.Sections("ReportFooter").Controls("rptIVA").Caption = FormatCurrency(txtTotIVA.Text)
    drIVAVentas.Sections("ReportFooter").Controls("rptIVA10").Caption = FormatCurrency(txtTotIVAMono.Text)
    drIVAVentas.Sections("ReportFooter").Controls("rptTotal").Caption = FormatCurrency(txtTotTotal.Text)
    
    drIVAVentas.Orientation = rptOrientLandscape
    
    drIVAVentas.Show
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub DTFechaD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        DTFechaH.SetFocus
    End If
    
End Sub

Private Sub DTFechaH_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        cmdBuscar.SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    'Posiciona el Form en Pantalla
    Me.Left = (zMain.ScaleWidth - Me.Width) / 2
    Me.Top = (zMain.ScaleHeight - Me.Height) / 2
    
    'Carga la fecha en los DTPickers
    CalcFechas
    
End Sub

Private Sub Form_Resize()
    
    'Ordena los objetos cuando se redimensiona el form
    If Me.Height > 2500 Then
        Flex.Width = Me.ScaleWidth - 240
        Flex.Height = Me.ScaleHeight - Flex.Top - 855
        Frame1.Top = Flex.Height + Flex.Top + 120
        Frame2.Top = Frame1.Top
    End If
    
    If Me.Width > 8340 Then
        Frame1.Left = Me.ScaleWidth - Frame1.Width - 120
    End If
    
End Sub

