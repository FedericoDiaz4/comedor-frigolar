VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ProveedoresList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Proveedores"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   5760
   ScaleWidth      =   8655
   Begin VB.CommandButton cmdExportar 
      Caption         =   "E&xportar"
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
      Left            =   1080
      Picture         =   "ProveedoresList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
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
      Left            =   7680
      Picture         =   "ProveedoresList.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Salir "
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdArticulos 
      Caption         =   "&Artículos"
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
      Left            =   120
      Picture         =   "ProveedoresList.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Left            =   4800
      Picture         =   "ProveedoresList.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Left            =   6720
      Picture         =   "ProveedoresList.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Eliminar"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
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
      Left            =   5760
      Picture         =   "ProveedoresList.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Modificar"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame frmBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.ComboBox cboBuscar 
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
         ItemData        =   "ProveedoresList.frx":213C
         Left            =   120
         List            =   "ProveedoresList.frx":214F
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   275
         Width           =   1815
      End
      Begin VB.TextBox txtBuscar 
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
         Left            =   2040
         TabIndex        =   2
         Top             =   275
         Width           =   6255
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   8415
      _ExtentX        =   14843
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
Attribute VB_Name = "ProveedoresList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboBuscar_Click()
    
    On Error Resume Next
    txtBuscar.Text = ""
    txtBuscar.SetFocus
    
End Sub

Private Sub cmdArticulos_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        
        Exit Sub
    End If
    
    ArticulosPrecios.idProveedor = Flex.TextMatrix(Flex.Row, 0)
    ArticulosPrecios.Show
    
End Sub

Private Sub cmdEliminar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    Select Case MsgBox("Desea ELIMINAR el proveedor" & Flex.TextMatrix(Flex.Row, 1) & "?" _
                       & vbCrLf & "Esta acción también eliminará la lista de precios asociados               " _
                       , vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
    
        Case vbNo: Exit Sub
    End Select
    
    'Elimina el proveedor
    Set rsDelete = New ADODB.Recordset
    SQL = "SELECT * FROM proveedores WHERE id = " & Flex.TextMatrix(Flex.Row, 0)
    rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsDelete.BOF And Not rsDelete.EOF Then
        rsDelete!eliminado = 1
        rsDelete.Update
    End If
    rsDelete.Close
    
    'Elimina la lista asociada
    Set rsDelete = New ADODB.Recordset
    SQL = "DELETE FROM articulosc WHERE idpro = " & Flex.TextMatrix(Flex.Row, 0)
    rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Private Sub cmdExportar_Click()
    
    Exportar_Excel "C:\", Flex
    
End Sub

Private Sub cmdModificar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    Proveedores.Nuevo = False
    Proveedores.ID = Flex.TextMatrix(Flex.Row, 0)
    
    Set rsGuardar = New ADODB.Recordset
    SQL = "SELECT * FROM proveedores WHERE id = " & Proveedores.ID
    rsGuardar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Proveedores.txtCodigo.Text = rsGuardar!codigo
    Proveedores.txtNombre.Text = rsGuardar!nombre
    Proveedores.txtDireccion.Text = rsGuardar!domicilio
    Proveedores.txtCUIT.Text = rsGuardar!cuit
    Proveedores.txtTelefono1.Text = rsGuardar!telefono
    Proveedores.txtTelefono2.Text = rsGuardar!tel1
    Proveedores.txtCelular.Text = rsGuardar!celular
    Proveedores.txtFax.Text = rsGuardar!fax
    Proveedores.txtContacto.Text = rsGuardar!contacto
    Proveedores.cboProvincia.Text = rsGuardar!provincia
    Proveedores.txtLocalidad.Text = rsGuardar!localidad
    Proveedores.txtCP.Text = rsGuardar!cp
    Proveedores.cboIVA.Text = rsGuardar!tipoiva
    Proveedores.txtEmail1.Text = rsGuardar!email
    Proveedores.txtEmail2.Text = rsGuardar!email2
    Proveedores.txtCotizacionDolar.Text = rsGuardar!cotizaciondolar
    Proveedores.txtCotizacionEuro.Text = rsGuardar!cotizacioneuro
    Proveedores.txtObs.Text = rsGuardar!tel2
    Proveedores.txtCPago.Text = rsGuardar!cpago
    Proveedores.txtFPago.Text = rsGuardar!fpago
    
    rsGuardar.Close
    
    Unload Me
    Proveedores.Show
    
End Sub

Private Sub cmdNuevo_Click()
    
    Proveedores.Nuevo = True
    Proveedores.Show
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    
    initForm Me
    cboBuscar.ListIndex = 0
    Cargar
    
End Sub

Sub OrdenaFlex()
    
    Flex.FormatString = "id|Código|Nombre|CUIT|Teléfono"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 1000
    Flex.ColWidth(2) = 3600
    Flex.ColWidth(3) = 1700
    Flex.ColWidth(4) = 1700
    
End Sub

Sub Cargar()
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT id, codigo, nombre, cuit, tel1 FROM proveedores WHERE eliminado = 0"
    If txtBuscar.Text <> "" Then
        SQL = SQL & " AND " & cboBuscar.Text & " LIKE '" & txtBuscar.Text & "%'"
    End If
    SQL = SQL & " ORDER BY " & cboBuscar.Text '& " LIMIT 100"
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
