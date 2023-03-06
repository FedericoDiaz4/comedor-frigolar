VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CompraC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Seleccionar Factura"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   3330
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
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
      Left            =   6720
      Picture         =   "CompraC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cargar Factura desde un Pedido seleccionado"
      Top             =   2760
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12640511
      ForeColorFixed  =   -2147483640
      GridColorFixed  =   -2147483630
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
   Begin VB.Label lblCliente 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor :"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "CompraC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idCliente As Integer
Public idRecibo As Integer
Public idPago As Integer 'ver esto

Private Sub cmdOk_Click()

    Dim idFactura As String
    
    idFactura = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
    
    ''Copia los datos de la factura en Pagos
    Set rsFacturas = New ADODB.Recordset
    SQL = "SELECT * FROM compras where id ='" & idFactura & "'"
    rsFacturas.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsFacturas.BOF And Not rsFacturas.EOF Then
        
        'Escribe el Tipo
        Pagos.txtTipoFac.Text = rsFacturas!tipo & " " & rsFacturas!letra
        
        'Escribe el id
        Pagos.idFac = rsFacturas!ID
        
        'Escribe el número
        Pagos.txtNumFac = rsFacturas!numero
        
        'Escribe la Fecha
        Pagos.txtFechaFac.Text = rsFacturas!fecha
        
        'Escribe el total
        Pagos.txtTotalFac.Text = rsFacturas!Total
        
        'Escribe el saldo
        Pagos.txtSaldoFac.Text = rsFacturas!Saldo
        
        'Escribe el abonado
        Pagos.txtAbonoFac.Text = rsFacturas!Saldo
        Pagos.txtAbonoFac.SelStart = 0
        Pagos.txtAbonoFac.SelLength = 10
        
    Else
        Call MsgBox("La Factura ingresada no existe.", vbExclamation, "Atención")
        Exit Sub
    End If
    rsFacturas.Close
    Unload Me
    
    Pagos.Show
    Pagos.txtDescFac.SetFocus

End Sub

Private Sub Form_Load()

    'Posiciona el form
    Me.Top = 1000
    Me.Left = (zMain.Width / 2) - (Me.Width / 2)
    
    MSHFlexGrid1.Rows = 2
    OrdenaFlex
    
    'Selecciona las facturas en CTA CTE
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM compras WHERE idproveedor = '" & idCliente & "' AND estado = 'CTACTE' ORDER BY fecha ASC"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not Recordset.EOF
    
        'Busca que la factura no haya sido ya abonada en el recibo actual
        Set rsImp = New ADODB.Recordset
        SQL = "SELECT * FROM pagosi where id = " & idRecibo & " AND idfactura = '" & Recordset!ID & "'"
        rsImp.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not rsImp.BOF And Not rsImp.EOF Then
            MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
        Else
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = Recordset!ID
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = Recordset!tipo & " " & Recordset!letra & "/" & Recordset!numero
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 2) = Format(Recordset!fecha, "dd/mm/yyyy")
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 3) = Recordset!Total
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 4) = Recordset!Saldo
        End If
        rsImp.Close
        Recordset.MoveNext
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    
    Loop
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
    Recordset.Close

End Sub

Public Sub OrdenaFlex()

    MSHFlexGrid1.FormatString = "id|ID|Fecha|Total|Saldo"
    MSHFlexGrid1.ColWidth(0) = 0
    MSHFlexGrid1.ColWidth(1) = 2200
    MSHFlexGrid1.ColWidth(2) = 1500
    MSHFlexGrid1.ColWidth(3) = 1500
    MSHFlexGrid1.ColWidth(4) = 1500
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Pagos.Show

End Sub

Private Sub MSHFlexGrid1_Click()

    cmdOk_Click

End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        cmdOk_Click
    End If

End Sub
