VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ListIVACompras 
   Caption         =   "Listado de IVA [Compras]"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   12345
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBuscar 
      Height          =   495
      Left            =   3240
      Picture         =   "ListIVACompras.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   " Aplicar "
      Top             =   360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   10920
      TabIndex        =   9
      Top             =   5400
      Width           =   1335
      Begin VB.CommandButton cmdSalir 
         Height          =   615
         Left            =   720
         Picture         =   "ListIVACompras.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   615
         Left            =   0
         Picture         =   "ListIVACompras.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   560
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   10695
      Begin VB.TextBox txtIVAMono 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPerIVA 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtIIBB 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9600
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtIVA21 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtIVA27 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtIVA10 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtNoGrav 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtNeto 
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "IVA Mono"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   27
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Perc. IVA"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   24
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "IIBB"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   23
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   22
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "IVA 21"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "IVA 27"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "IVA 10,5"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "No Grav."
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Neto"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker DTFechaD 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   90243073
      CurrentDate     =   39603
   End
   Begin MSComCtl2.DTPicker DTFechaH 
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   90243073
      CurrentDate     =   39603
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7646
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      FixedCols       =   0
      BackColorFixed  =   12640511
      ForeColorFixed  =   0
      BackColorSel    =   8388608
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "ListIVACompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub OrdenaFlex()

ListIVACompras.MSHFlexGrid1.FormatString = "Tipo / Fact Nº|Fecha|Proveedor|CUIT|Neto|No Grav.|IVA 10,5%|IVA 21%|IVA 27%|IVA Mono|Percep. IVA|IIBB|Total"
ListIVACompras.MSHFlexGrid1.ColWidth(0) = 1700  'Tipo
ListIVACompras.MSHFlexGrid1.ColWidth(1) = 1200  'Fecha
ListIVACompras.MSHFlexGrid1.ColWidth(2) = 2500  'Proveedor
ListIVACompras.MSHFlexGrid1.ColWidth(3) = 1200  'CUIT
ListIVACompras.MSHFlexGrid1.ColWidth(4) = 1200  'Neto
ListIVACompras.MSHFlexGrid1.ColWidth(5) = 1200  'Concepto no Grav.
ListIVACompras.MSHFlexGrid1.ColWidth(6) = 1200  'IVA 10,5%
ListIVACompras.MSHFlexGrid1.ColWidth(7) = 1200  'IVA 21%
ListIVACompras.MSHFlexGrid1.ColWidth(8) = 1200  'IVA 27%
ListIVACompras.MSHFlexGrid1.ColWidth(9) = 1200  'IVA Mono
ListIVACompras.MSHFlexGrid1.ColWidth(10) = 1200 'Percep. IVA
ListIVACompras.MSHFlexGrid1.ColWidth(11) = 1200 'IIBB
ListIVACompras.MSHFlexGrid1.ColWidth(12) = 1200 'Total

End Sub

Sub CalcFechas()

'Calcula el priemr y último día del mes anterior
Mes = Format(Date, "m") - 1
Año = Format(Date, "yyyy")

DTFechaD.Value = DateSerial(Año, Mes + 0, 1)
DTFechaH.Value = DateSerial(Año, Mes + 1, 0)

End Sub

Public Sub cmdBuscar_Click()
Dim Total, Desc1Pesos, Desc2Pesos, Neto, IVA21 As Single

'Borra todo de la tabla temporal
Set rsBorraTemp = New ADODB.Recordset
SQL = "DELETE FROM temp"
rsBorraTemp.Open SQL, Data, adOpenKeyset, adLockOptimistic

'Abre la tabla temporal
Set rsTemp = New ADODB.Recordset
SQL = "SELECT * from temp;"
rsTemp.Open SQL, Data, adOpenKeyset, adLockOptimistic

'Limpia los campos de Texto
txtNeto.Text = "0,00"
txtNoGrav.Text = "0,00"
txtIVA10.Text = "0,00"
txtIVA21.Text = "0,00"
txtIVA27.Text = "0,00"
txtIVAMono.Text = "0,00"
txtPerIVA.Text = "0,00"
txtIIBB.Text = "0,00"
txtTotal.Text = "0,00"

'Selecciona las facturas entre el rango de fecha seleccionado
Set Recordset = New ADODB.Recordset
SQL = "SELECT * FROM proveedoresdf WHERE fechaiva BETWEEN  '" & Format(DTFechaD.Value, "yyyy-mm-dd") & "' AND '" & Format(DTFechaH.Value, "yyyy-mm-dd") & "' ORDER BY fecha ASC"
Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
Do While Not Recordset.EOF
    rsTemp.AddNew
    If Left(Recordset!Tipo, 15) = "Nota de Crédito" Then
        rsTemp!col1 = "NC" & Trim(Right(Recordset!Tipo, 2)) & "/" & Recordset!idreal
    ElseIf Left(Recordset!Tipo, 14) = "Nota de Débito" Then
        rsTemp!col1 = "ND" & Trim(Right(Recordset!Tipo, 2)) & "/" & Recordset!idreal
    Else
        rsTemp!col1 = Trim(Right(Recordset!Tipo, 2)) & "/" & Recordset!idreal
    End If
    rsTemp!col2 = Format(Recordset!fecha, "dd/mm/yyyy")
    'Verifica que la factura no esté cancelada
    If Recordset!idProveedor <> "0" Then
        'Carga los datos del proveedor
        Set rsCli = New ADODB.Recordset
        SQL = "SELECT * FROM proveedores WHERE id = '" & Recordset!idProveedor & "'"
        rsCli.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not rsCli.BOF And Not rsCli.EOF Then
            rsTemp!col3 = rsCli!razon
            rsTemp!col4 = rsCli!cuit
        End If
        rsCli.Close
        
        ' - Ingresa los datos en un registro de la base
        If Recordset!impinterno <> "" Then
            impinterno = Recordset!impinterno
        Else
            impinterno = 0
        End If
        If Recordset!otrosimp <> "" Then
            otrosimp = Recordset!otrosimp
        Else
            otrosimp = 0
        End If
        rsTemp!col5 = Format(Recordset!Bruto, "0.00")
        rsTemp!col6 = Format(CDec(impinterno), "0.00")
        If Recordset!iva10 <> "" Then
            rsTemp!col7 = Format(CDec(Recordset!iva10), "0.00")
        Else
            rsTemp!col7 = "0,00"
        End If
        If Recordset!IVA21 <> "" Then
            rsTemp!col8 = Format(CDec(Recordset!IVA21), "0.00")
        Else
            rsTemp!col8 = "0,00"
        End If
        If Recordset!ivadif <> "" Then
            rsTemp!col9 = Format(CDec(Recordset!ivadif), "0.00")
        Else
            rsTemp!col9 = "0,00"
        End If
        If Right(Recordset!Tipo, 1) = "C" Then 'IVA Mono
            a = Format(CDec(Recordset!Total) / 1.21, "0.00")
            rsTemp!col13 = Format(CDec(Recordset!Total) - CDec(a), "0.00")
            a = 0
        Else
            rsTemp!col13 = "0,00"
        End If
        If Recordset!periva <> "" Then
            rsTemp!col10 = Format(CDec(Recordset!periva), "0.00")
        Else
            rsTemp!col10 = "0,00"
        End If
        If Recordset!IIBB <> "" Then
            rsTemp!col11 = Format(CDec(Recordset!IIBB), "0.00")
        Else
            rsTemp!col11 = "0,00"
        End If
        rsTemp!col12 = Format(CDec(Recordset!Total), "0.00")
        
        'Calcula los Totales:
        If Left(Recordset!Tipo, 15) = "Nota de Crédito" Then
            'Si es Nota de Crédito resta
            txtNeto.Text = Format(CDec(txtNeto.Text) - CDec(rsTemp!col5), "0.00")
            txtNoGrav.Text = Format(CDec(txtNoGrav.Text) - CDec(rsTemp!col6), "0.00")
            txtIVA10.Text = Format(CDec(txtIVA10.Text) - CDec(rsTemp!col7), "0.00")
            txtIVA21.Text = Format(CDec(txtIVA21.Text) - CDec(rsTemp!col8), "0.00")
            txtIVA27.Text = Format(CDec(txtIVA27.Text) - CDec(rsTemp!col9), "0.00")
            txtIVAMono.Text = Format(CDec(txtIVAMono.Text) - CDec(rsTemp!col13), "0.00")
            txtPerIVA.Text = Format(CDec(txtPerIVA.Text) - CDec(rsTemp!col10), "0.00")
            txtIIBB.Text = Format(CDec(txtIIBB.Text) - CDec(rsTemp!col11), "0.00")
            txtTotal.Text = Format(CDec(txtTotal.Text) - CDec(rsTemp!col12), "0.00")
        Else
            'Si es Factura o Nota de Débito, Suma
            txtNeto.Text = Format(CDec(txtNeto.Text) + CDec(rsTemp!col5), "0.00")
            txtNoGrav.Text = Format(CDec(txtNoGrav.Text) + CDec(rsTemp!col6), "0.00")
            txtIVA10.Text = Format(CDec(txtIVA10.Text) + CDec(rsTemp!col7), "0.00")
            txtIVA21.Text = Format(CDec(txtIVA21.Text) + CDec(rsTemp!col8), "0.00")
            txtIVA27.Text = Format(CDec(txtIVA27.Text) + CDec(rsTemp!col9), "0.00")
            txtIVAMono.Text = Format(CDec(txtIVAMono.Text) + CDec(rsTemp!col13), "0.00")
            txtPerIVA.Text = Format(CDec(txtPerIVA.Text) + CDec(rsTemp!col10), "0.00")
            txtIIBB.Text = Format(CDec(txtIIBB.Text) + CDec(rsTemp!col11), "0.00")
            txtTotal.Text = Format(CDec(txtTotal.Text) + CDec(rsTemp!col12), "0.00")
        End If
    Else '(Si la factura está cancelada)
        rsTemp!col2 = "ANULADA"
        rsTemp!col3 = "-"
        rsTemp!col4 = "-"
        rsTemp!col5 = "-"
        rsTemp!col6 = "-"
        rsTemp!col7 = "-"
        rsTemp!col8 = "-"
        rsTemp!col9 = "-"
        rsTemp!col10 = "-"
        rsTemp!col11 = "-"
        rsTemp!col12 = "-"
        rsTemp!col13 = "-"
    End If
    
    rsTemp.Update
    Recordset.MoveNext
Loop
Recordset.Close
rsTemp.Close

'Prepara el flex
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
OrdenaFlex

'Muestra el resultado en el flexGrid
Set rsTemp = New ADODB.Recordset
SQL = "SELECT col1,col2,col3,col4,col5,col6,col7,col8,col9,col13,col10,col11,col12 from temp;"
rsTemp.Open SQL, Data, adOpenKeyset, adLockOptimistic
If Not rsTemp.BOF And Not rsTemp.EOF Then
    Set MSHFlexGrid1.DataSource = rsTemp
End If
rsTemp.Close
OrdenaFlex

End Sub

Private Sub cmdImprimir_Click()

'Actualiza la conexión de el DataEnvironment
'ActualizarDR

'Escribe el encabezado en el Data Report
drIVACompras.Sections("ReportHeader").Controls("rptDesde").Caption = DTFechaD.Value
drIVACompras.Sections("ReportHeader").Controls("rptHasta").Caption = DTFechaH.Value

'Escribe los totales
drIVACompras.Sections("ReportFooter").Controls("rptNeto").Caption = Format(txtNeto.Text, "0.00")
drIVACompras.Sections("ReportFooter").Controls("rptNoGrav").Caption = Format(txtNoGrav.Text, "0.00")
drIVACompras.Sections("ReportFooter").Controls("rptIVA10").Caption = Format(txtIVA10.Text, "0.00")
drIVACompras.Sections("ReportFooter").Controls("rptIVA21").Caption = Format(txtIVA21.Text, "0.00")
drIVACompras.Sections("ReportFooter").Controls("rptIVA27").Caption = Format(txtIVA27.Text, "0.00")
drIVACompras.Sections("ReportFooter").Controls("rptIVAMono").Caption = Format(txtIVAMono.Text, "0.00")
drIVACompras.Sections("ReportFooter").Controls("rptPerIVA").Caption = Format(txtPerIVA.Text, "0.00")
drIVACompras.Sections("ReportFooter").Controls("rptIIBB").Caption = Format(txtIIBB.Text, "0.00")
drIVACompras.Sections("ReportFooter").Controls("rptTotal").Caption = Format(txtTotal.Text, "0.00")

drIVACompras.Orientation = rptOrientLandscape

drIVACompras.Show

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
Me.Left = (zPrincipal.ScaleWidth - Me.Width) / 2
Me.Top = (zPrincipal.ScaleHeight - Me.Height) / 2

'Carga la fecha en los DTPickers
CalcFechas

End Sub

Private Sub Form_Resize()

'Ordena los objetos cuando se redimensiona el form
If Me.Height > 2500 Then
    MSHFlexGrid1.Width = Me.ScaleWidth - 240
    MSHFlexGrid1.Height = Me.ScaleHeight - MSHFlexGrid1.Top - 855
    Frame1.Top = MSHFlexGrid1.Height + MSHFlexGrid1.Top + 120
    Frame2.Top = Frame1.Top
End If
If Me.Width > Frame1.Width + Frame2.Width + 360 Then
    Frame1.Left = Me.ScaleWidth - Frame1.Width - 120
End If

End Sub
