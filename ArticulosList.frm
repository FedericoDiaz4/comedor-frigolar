VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ArticulosList 
   Caption         =   "Artículos"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   12015
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBotones 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   8160
      ScaleHeight     =   615
      ScaleWidth      =   3735
      TabIndex        =   4
      Top             =   5040
      Width           =   3735
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   2880
         Picture         =   "ArticulosList.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " Salir "
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   615
         Left            =   0
         Picture         =   "ArticulosList.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Nuevo"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   615
         Left            =   1920
         Picture         =   "ArticulosList.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Eliminar"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   615
         Left            =   960
         Picture         =   "ArticulosList.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Modificar"
         Top             =   0
         Width           =   855
      End
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
      Width           =   11775
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
         ItemData        =   "ArticulosList.frx":1628
         Left            =   120
         List            =   "ArticulosList.frx":1632
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   250
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
         Top             =   250
         Width           =   9615
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
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
Attribute VB_Name = "ArticulosList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Entrada As String

Private Sub cboBuscar_Click()
    
    On Error Resume Next
    txtBuscar.Text = ""
    
End Sub

Private Sub cboBuscar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmdEliminar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    Select Case MsgBox("¿DESEA ELIMINAR EL ARTICULO " & Flex.TextMatrix(Flex.Row, 1) & "?", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
        Case vbNo: Exit Sub
    End Select
    
    Set rsDelete = New ADODB.Recordset
    SQL = "UPDATE articulos SET eliminado = 1 WHERE id = " & Flex.TextMatrix(Flex.Row, 0)
    rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Private Sub cmdModificar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    Articulos.Nuevo = False
    Articulos.ID = Flex.TextMatrix(Flex.Row, 0)
    
    Set rsCargarArt = New ADODB.Recordset
    SQL = "SELECT * FROM articulos WHERE id = " & Articulos.ID
    rsCargarArt.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Articulos.txtCodigo.Text = rsCargarArt!codigo
    Articulos.txtNombre.Text = rsCargarArt!nombre
    Articulos.cboIVA.Text = getData(rsCargarArt!idTipoIVA, "nombre", "tipoiva")
    Articulos.txtPrecio.Text = Format(rsCargarArt!Precio, "0.00")
    
    rsCargarArt.Close
    
    Unload Me
    Articulos.Show
    
End Sub

Private Sub cmdNuevo_Click()
    
    Articulos.Nuevo = True
    Articulos.ID = 0
    Articulos.Show
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Sub Flex_Click()
    
    If Entrada = "COMPRA" Then
        
        If Flex.TextMatrix(Flex.Row, 0) = "" Then
            Exit Sub
        End If
        Compra.txtCodArticulo.Text = Flex.TextMatrix(Flex.Row, 1)
        Compra.txtCantidad.SetFocus
        Unload Me
        
    ElseIf Entrada = "VENTA" Then
        
        If Flex.TextMatrix(Flex.Row, 0) = "" Then
            Exit Sub
        End If
        Venta.txtCodArticulo.Text = Flex.TextMatrix(Flex.Row, 1)
        Venta.txtCantidad.SetFocus
        Unload Me
    
    ElseIf Entrada = "FALTANTES" Then
        
        If Flex.TextMatrix(Flex.Row, 0) = "" Then
            Exit Sub
        End If
        ArticulosFaltantes.txtCodArticulo.Text = Flex.TextMatrix(Flex.Row, 1)
        ArticulosFaltantes.Add
        Unload Me
        
    End If
    
End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Flex_Click
    End If
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    cboBuscar.ListIndex = 1
    Cargar
    
End Sub

Sub OrdenaFlex()
    
    Flex.FormatString = "id|Código|Nombre|IVA|Precio"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 1300
    Flex.ColWidth(2) = 7700
    Flex.ColWidth(3) = 1300
    Flex.ColWidth(4) = 1300
    
End Sub

Sub Cargar()
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT a.id, a.codigo, a.nombre, i.nombre, a.precio FROM articulos AS a Inner Join tipoiva AS i ON i.id = a.idtipoiva WHERE a.eliminado = 0 "
    If txtBuscar.Text <> "" Then
        SQL = SQL & " AND a." & cboBuscar.Text & " LIKE '" & txtBuscar.Text & "%'"
    End If
    SQL = SQL & " ORDER BY a." & cboBuscar.Text '& " LIMIT 250"
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

Private Sub Form_Resize()
    
    If Me.ScaleHeight < 2000 Or Me.ScaleWidth < 2000 Then
        Exit Sub
    End If

    Const Margen = 120

    Flex.Width = Me.ScaleWidth - 2 * Margen
    If Entrada = "" Then
        Flex.Height = Me.ScaleHeight - picBotones.Height - frmBuscar.Height - 4 * Margen
        picBotones.Left = Me.ScaleWidth - picBotones.Width - Margen
        picBotones.Top = Flex.Height + frmBuscar.Height + Margen * 3
    Else
        Flex.Height = Me.ScaleHeight - frmBuscar.Height - 3 * Margen
    End If
    
End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then
        Flex.SetFocus
    End If
    
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 13 Then
        KeyAscii = 0
        Cargar
    End If
    
End Sub
