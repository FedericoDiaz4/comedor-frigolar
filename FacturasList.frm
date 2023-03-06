VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form VentaList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9615
   Begin VB.TextBox txtCodCliente 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   2
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   1680
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
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
      Height          =   615
      Left            =   7680
      Picture         =   "FacturasList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Modificar "
      Top             =   5040
      Width           =   855
   End
   Begin VB.ComboBox cboTipoComprobante 
      Height          =   360
      ItemData        =   "FacturasList.frx":058A
      Left            =   120
      List            =   "FacturasList.frx":058C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.ComboBox cboCliente 
      Height          =   360
      ItemData        =   "FacturasList.frx":058E
      Left            =   4440
      List            =   "FacturasList.frx":05A1
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   5055
   End
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
      Left            =   8640
      Picture         =   "FacturasList.frx":05D3
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Salir "
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
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
      Left            =   6720
      Picture         =   "FacturasList.frx":0B5D
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Nuevo "
      Top             =   5040
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7011
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
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
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
   Begin VB.Label Label12 
      Caption         =   "Código Cliente"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "Nombre Cliente"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "VentaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboCliente_Click()
    
    Cargar
    
End Sub

Private Sub cboTipoComprobante_Click()
    
    Cargar
    
End Sub

Private Sub cmdAnular_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    Select Case MsgBox("¿DESEA ANULAR LA FACTURA?", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
        Case vbNo: Exit Sub
    End Select
    
    Set rsDelete = New ADODB.Recordset
    SQL = "SELECT * FROM facturasdf WHERE id = " & Flex.TextMatrix(Flex.Row, 0)
    rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsDelete.BOF And Not rsDelete.EOF Then
        rsDelete!idCliente = 0
        rsDelete!Bruto = "0,00"
        rsDelete!Total = "0,00"
        rsDelete!Estado = "ANULADO"
        rsDelete.Update
    End If
    rsDelete.Close
    
    Cargar
    
End Sub

Private Sub cmdModificar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    Venta.Modificando = True
    Venta.ID = Flex.TextMatrix(Flex.Row, 0)
    
    Set rsVenta = New ADODB.Recordset
    SQL = "SELECT * FROM ventas WHERE id = " & Venta.ID
    rsVenta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Venta.DTFechaComprobante.Value = rsVenta!fecha
    Venta.cboTipoComprobante.Text = getData(rsVenta!idtipocomprobante, "nombre", "tipocomprobante")
    Venta.cboPtoVenta.Text = rsVenta!ptoventa
    Venta.txtNumComprobante.Text = Format(rsVenta!numerocomprobante, "0000000")
    Venta.txtCodCliente.Text = getData(rsVenta!idCliente, "codigo", "clientes")
    If rsVenta!cae <> "" Then
        Venta.txtCAE = rsVenta!cae
        Venta.DTVencimientoCAE = rsVenta!fechavtocae
    End If
    rsVenta.Close
    
    Unload Me
    Venta.Show
    Venta.Cargar
    
End Sub

Private Sub cmdNuevo_Click()
    
    Venta.Modificando = False
    Venta.ID = 0
    Venta.Show
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Flex_Click()
    
    Tipo = Flex.TextMatrix(Flex.Row, 2)
    
'    If Tipo = "REMITO" Then
        cmdModificar.Enabled = True
'    Else
'        cmdModificar.Enabled = False
'    End If
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    CargaCombo "tipocomprobante", "nombre", "id", cboTipoComprobante
    CargaCombo "clientes", "nombre", "nombre", cboCliente, True, "id<>0"
    cboCliente.ListIndex = 0
    cboTipoComprobante.ListIndex = 0
    Cargar
    
End Sub

Sub OrdenaFlex()
    
    Flex.FormatString = "id|Fecha|Tipo|Nombre|Neto|Total"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 1200
    Flex.ColWidth(2) = 1350
    Flex.ColWidth(3) = 4200
    Flex.ColWidth(4) = 1100
    Flex.ColWidth(5) = 1100
    
End Sub

Sub Cargar()
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT f.id, f.fecha, t.nombre, c.nombre, f.totalneto, f.total FROM ventas AS f Inner Join clientes AS c ON f.idcliente = c.id Inner Join tipocomprobante AS t ON f.idtipocomprobante = t.id WHERE 0 = 0 "
    If cboCliente.ListIndex <> 0 And cboCliente.ListIndex <> -1 Then
        SQL = SQL & " AND f.idcliente = '" & cboCliente.ItemData(cboCliente.ListIndex) & "'"
    End If
    If cboTipoComprobante.ListIndex <> -1 Then
        SQL = SQL & " AND f.idtipocomprobante = '" & cboTipoComprobante.ItemData(cboTipoComprobante.ListIndex) & "'"
    End If
    SQL = SQL & " ORDER BY fecha DESC"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        Set Flex.DataSource = Recordset
    Else
        Flex.Clear
        Flex.Rows = 2
    End If
    Recordset.Close
    
    OrdenaFlex
    
End Sub

Private Sub txtCodCliente_Change()
    
    Set rsCli = New ADODB.Recordset
    SQL = "SELECT nombre FROM clientes WHERE codigo = '" & txtCodCliente.Text & "' AND eliminado <> 1"
    rsCli.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsCli.BOF And Not rsCli.EOF Then
        If rsCli!nombre <> "" Then
            cboCliente.Text = rsCli!nombre
        End If
        rsCli.Close
    Else
        cboCliente.ListIndex = -1
    End If
    
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub
