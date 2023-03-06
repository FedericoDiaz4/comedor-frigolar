VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ChequesList 
   Caption         =   "Cheques"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   9855
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   9855
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
      Left            =   8880
      Picture         =   "ChequesList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Salir "
      Top             =   5040
      Width           =   855
   End
   Begin VB.CheckBox chkCartera 
      Caption         =   "Sólo cheques en cartera"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Value           =   1  'Checked
      Width           =   2655
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
      Left            =   6960
      Picture         =   "ChequesList.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Nuevo"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
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
      Left            =   7920
      Picture         =   "ChequesList.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Eliminar"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame frmBuscar 
      Caption         =   "Buscar"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton cmdBuscar 
         Height          =   580
         Left            =   7425
         Picture         =   "ChequesList.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   130
         Width           =   615
      End
      Begin VB.ComboBox cboBuscar 
         Height          =   360
         ItemData        =   "ChequesList.frx":1628
         Left            =   120
         List            =   "ChequesList.frx":1635
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   250
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   345
         Left            =   2400
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   250
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
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
         Format          =   89849857
         CurrentDate     =   36526
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   345
         Left            =   4920
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   250
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
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
         Format          =   89849857
         CurrentDate     =   36526
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   9615
      _ExtentX        =   16960
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
End
Attribute VB_Name = "ChequesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboBuscar_Click()
    
    On Error Resume Next
    txtBuscar.Text = ""
    txtBuscar.SetFocus
    
End Sub

Private Sub cmdEliminar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    Select Case MsgBox("¿DESEA ELIMINAR EL CLIENTE " & Flex.TextMatrix(Flex.Row, 1) & "?", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
        Case vbNo: Exit Sub
    End Select
    
    Set rsDelete = New ADODB.Recordset
    SQL = "UPDATE clientes SET eliminado = 1 WHERE id = " & Flex.TextMatrix(Flex.Row, 0)
    rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Private Sub cmdModificar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    Clientes.Nuevo = False
    Clientes.ID = Flex.TextMatrix(Flex.Row, 0)
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM clientes WHERE id = " & Clientes.ID
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Clientes.txtCodigo.Text = Format(Recordset!codigo, "0000")
    Clientes.txtNombre.Text = Recordset!nombre
    Clientes.txtCUIT.Text = Recordset!cuit
    Clientes.txtLocalidad.Text = Recordset!localidad
    Clientes.txtTelefono.Text = Recordset!telefono
    Clientes.txtTelParticular.Text = Recordset!telparticular
    Clientes.txtDireccion.Text = Recordset!domicilio
    Clientes.txtCP.Text = Recordset!cp
    Clientes.txtEmail.Text = Recordset!email
    Clientes.txtEmail2.Text = Recordset!email2
    Clientes.cboIVA.Text = Recordset!tipoiva
    Clientes.txtContacto.Text = Recordset!contacto
    Clientes.txtFax.Text = Recordset!fax
    Clientes.txtDescuento.Text = Recordset!descuentohabitual
    Clientes.cboProvincia.Text = Recordset!provincia
    Clientes.txtObs.Text = Recordset!obs
    
    Recordset.Close
    
    Unload Me
    Clientes.Show
    
End Sub

Private Sub cmdNuevo_Click()
    
    Clientes.Nuevo = True
    Unload Me
    Clientes.Show
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    cboBuscar.ListIndex = 0
    DTDesde.Value = "01/01/2013"
    DTHasta.Value = Date
    Cargar
    
End Sub

Sub OrdenaFlex()
    
'    Flex.FormatString = "id|Código|Nombre|CUIT|Teléfono|Localidad"
'    Flex.ColWidth(0) = 0
'    Flex.ColWidth(1) = 880
'    Flex.ColWidth(2) = 3000
'    Flex.ColWidth(3) = 1800
'    Flex.ColWidth(4) = 1800
'    Flex.ColWidth(5) = 1800
    
End Sub

Sub Cargar()
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT id, emision, banco, ncheque, importe FROM cheques WHERE eliminado = 0"
    If txtBuscar.Text <> "" Then
        SQL = SQL & " AND " & cboBuscar.Text & " BETWEEN '" & DTDesde.Value & "' AND '" & DTHasta.Value & "' "
    End If
    If chkCartera.Value = vbChecked Then
        SQL = SQL & " AND idproveedor = 0"
    End If
    SQL = SQL & " ORDER BY " & cboBuscar.Text
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

Private Sub txtBuscar_Change()
    
    Cargar
    
End Sub

