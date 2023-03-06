VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form RecibosList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
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
   ScaleWidth      =   10335
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
      Left            =   8400
      Picture         =   "RecibosList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Nuevo"
      Top             =   5040
      Width           =   855
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
      Left            =   9360
      Picture         =   "RecibosList.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Salir "
      Top             =   5040
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
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
   Begin VB.Frame frmBuscar 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.ComboBox cboVendedor 
         Height          =   360
         ItemData        =   "RecibosList.frx":0B14
         Left            =   5760
         List            =   "RecibosList.frx":0B27
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.ComboBox cboCliente 
         Height          =   360
         ItemData        =   "RecibosList.frx":0B59
         Left            =   120
         List            =   "RecibosList.frx":0B6C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   5760
         TabIndex        =   7
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   675
      End
   End
End
Attribute VB_Name = "RecibosList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboCliente_Click()
    
    Cargar
    
End Sub

Private Sub cboVendedor_Click()
    
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
        rsDelete!bruto = "0,00"
        rsDelete!Total = "0,00"
        rsDelete!Estado = "ANULADO"
        rsDelete.Update
    End If
    rsDelete.Close
    
    Cargar
    
End Sub

Private Sub cmdImportar_Click()

    Cargar

End Sub

Private Sub cmdImprimir_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    a = Flex.TextMatrix(Flex.Row, 0)
    i = 1
    p = 1
    
    Set drRecibos = Nothing
    
    Set rsRecibos = New ADODB.Recordset
    SQL = "SELECT * FROM recibosdr WHERE id = " & a
    rsRecibos.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    With drRecibos.Sections("ReportHeader")
        
        'Encabezado
        .Controls("lblFecha").Caption = Format(rsRecibos!fecha, "dd/mm/yyyy")
        .Controls("lblNomCli").Caption = getData(rsRecibos!idCliente, "nombre", "clientes")
        .Controls("lblIdCli").Caption = "(" & getData(rsRecibos!idCliente, "codigo", "clientes") & ")"
        .Controls("lblDirCli").Caption = getData(rsRecibos!idCliente, "calle", "clientes")
        .Controls("lblLocCli").Caption = getData(rsRecibos!idCliente, "localidad", "clientes")
        .Controls("lblTelCli").Caption = getData(rsRecibos!idCliente, "telefono", "clientes")
        
        'Imputaciones
        Set rsRecibosi = New ADODB.Recordset
        SQL = "SELECT * FROM recibosi WHERE id = " & a
        rsRecibosi.Open SQL, Data, adOpenKeyset, adLockOptimistic
        Do While Not rsRecibosi.EOF
            .Controls("lblTipoFac" & i).Caption = rsRecibosi!Tipo
            .Controls("lblNumFac" & i).Caption = Format(rsRecibosi!idFactura, "00000000")
            .Controls("lblFechaFac" & i).Caption = Format(rsRecibosi!fecha, "dd/mm/yyyy")
            .Controls("lblTotalFac" & i).Caption = Format(rsRecibosi!Total, "0.00")
            .Controls("lblDescuentoFac" & i).Caption = Format(rsRecibosi!DescPesos, "0.00")
            .Controls("lblSaldoFac" & i).Caption = Format(rsRecibosi!Saldo, "0.00")
            .Controls("lblAbonoFac" & i).Caption = Format(rsRecibosi!Abonado, "0.00")
            rsRecibosi.MoveNext
            i = i + 1
        Loop
        rsRecibosi.Close
        
        'Total
        .Controls("txtTotalFacturas").Caption = Format(rsRecibos!totalImp, "0.00")
        
        
        'Pagos
        Set rsRecibosp = New ADODB.Recordset
        SQL = "SELECT * FROM recibosp WHERE id = " & a
        rsRecibosp.Open SQL, Data, adOpenKeyset, adLockOptimistic
        Do While Not rsRecibosp.EOF
            .Controls("lblTipoPago" & p).Caption = rsRecibosp!Tipo
            .Controls("lblDetallePago" & p).Caption = rsRecibosp!detalle
            .Controls("lblFechaPago" & p).Caption = Format(rsRecibosp!fecha, "dd/mm/yyyy")
            .Controls("lblImportePago" & p).Caption = Format(rsRecibosp!importe, "0.00")
            rsRecibosp.MoveNext
            p = p + 1
        Loop
        rsRecibosp.Close
        
        'Total
        .Controls("txtTotalPagos").Caption = Format(rsRecibos!totalPag, "0.00")
        
    End With
    
    rsRecibos.Close
    
    drRecibos.Orientation = rptOrientLandscape
    drRecibos.Show
    
End Sub

Private Sub cmdModificar_Click()

    idRecibo = Val(Flex.TextMatrix(Flex.Row, 0))
    
    If idRecibo = "" Or idRecibo = "0" Then
        Exit Sub
    End If
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM recibosdr WHERE id = " & idRecibo
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        RecibosD.idRecibo = idRecibo
        RecibosD.Modificar = True
        RecibosD.Show
        RecibosD.DTFecha.Value = Recordset!fecha
        RecibosD.DTFecha.Enabled = False
        RecibosD.cboIdCli.Text = Recordset!idCliente
        RecibosD.cboCodCli.Enabled = False
        RecibosD.cboNomCli.Enabled = False
        RecibosD.cboIdVendedor.Text = Recordset!idVendedor
        RecibosD.cboVendedor.Enabled = False
        RecibosD.txtComision.Text = CDec(Recordset!comision)
    End If
    Unload Me

End Sub

Private Sub cmdNuevo_Click()
    
    RecibosD.Modificar = False
    RecibosD.Show
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    CargaCombo "clientes", "nombre", "nombre", cboCliente, True, "id<>0"
    CargaCombo "vendedores", "nombre", "nombre", cboVendedor, True, "id<>0"
    cboCliente.ListIndex = 0
    cboVendedor.ListIndex = 0
    Cargar
    
End Sub

Sub OrdenaFlex()
    
    Flex.FormatString = "id|Fecha|Cliente|Vendedor|Descuentos|Imputaciones|Pagos"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 1200
    Flex.ColWidth(2) = 3000
    Flex.ColWidth(3) = 2000
    Flex.ColWidth(4) = 1200
    Flex.ColWidth(5) = 1200
    Flex.ColWidth(6) = 1200
    
End Sub

Sub Cargar()
    
    Set Recordset = New ADODB.Recordset
    'SQL = "SELECT r.id,DATE_FORMAT(r.fecha,'%d/%m/%Y'),c.nombre,v.nombre,r.totaldes,r.totalimp,r.totalpag FROM recibosdr AS r Inner Join clientes AS c ON r.idcliente = c.id Inner Join vendedores AS v ON r.idvendedor = v.id WHERE r.id <> 0 "
    SQL = "SELECT r.id,r.fecha,c.nombre,v.nombre,r.totaldes,r.totalimp,r.totalpag FROM recibosdr AS r Inner Join clientes AS c ON r.idcliente = c.id Inner Join vendedores AS v ON r.idvendedor = v.id WHERE r.id <> 0 "
    
    If cboCliente.ListIndex <> 0 And cboCliente.ListIndex <> -1 Then
        SQL = SQL & " AND r.idcliente = '" & cboCliente.ItemData(cboCliente.ListIndex) & "'"
    End If
    
    If cboVendedor.ListIndex <> 0 And cboVendedor.ListIndex <> -1 Then
        SQL = SQL & " AND r.idvendedor = '" & cboVendedor.ItemData(cboVendedor.ListIndex) & "'"
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


