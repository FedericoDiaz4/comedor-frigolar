VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EmpleadosList 
   Caption         =   "Listado de Empleados"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   9855
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9855
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBotones 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6000
      ScaleHeight     =   615
      ScaleWidth      =   3735
      TabIndex        =   4
      Top             =   5040
      Width           =   3735
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
         Left            =   0
         Picture         =   "EmpleadosList.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " Nuevo"
         Top             =   0
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
         Left            =   1920
         Picture         =   "EmpleadosList.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Eliminar"
         Top             =   0
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
         Left            =   960
         Picture         =   "EmpleadosList.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Modificar"
         Top             =   0
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
         Left            =   2880
         Picture         =   "EmpleadosList.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Salir "
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
      Width           =   9615
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
         Width           =   7455
      End
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
         ItemData        =   "EmpleadosList.frx":1628
         Left            =   120
         List            =   "EmpleadosList.frx":1632
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   250
         Width           =   1815
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
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
Attribute VB_Name = "EmpleadosList"
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
    
    Select Case MsgBox("¿DESEA ELIMINAR EL EMPLEADO " & Flex.TextMatrix(Flex.Row, 2) & "?", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
        Case vbNo: Exit Sub
    End Select
    
    Set rsDelete = New ADODB.Recordset
    SQL = "UPDATE empleados SET eliminado = 1 WHERE id = " & Flex.TextMatrix(Flex.Row, 0)
    rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Private Sub cmdModificar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    Empleados.Nuevo = False
    Empleados.id = Flex.TextMatrix(Flex.Row, 0)
    
    Set rsCliente = New ADODB.Recordset
    SQL = "SELECT * FROM empleados WHERE id = " & Empleados.id
    rsCliente.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Empleados.txtNombre.Text = rsCliente!nombre
    Empleados.txtNumero.Text = rsCliente!numerodocumento
    Empleados.txtNroLegajo.Text = rsCliente!nrolegajo
    Empleados.txtCuil.Text = rsCliente!Cuil
    rsCliente.Close
    
    Unload Me
    Empleados.Show
    
End Sub

Private Sub cmdNuevo_Click()
    
    Empleados.Nuevo = True
    Unload Me
    Empleados.Show
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    cboBuscar.ListIndex = 0
    Cargar
    
End Sub

Sub ordenaFlex()
    
    Flex.FormatString = "id|Num Doc|Nombre|Nro Legajo|Cuil"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 1400
    Flex.ColWidth(2) = 4000
    Flex.ColWidth(3) = 1100
    Flex.ColWidth(4) = 1650
    
End Sub

Sub Cargar()
    
    Flex.Clear
    Flex.Rows = 2
    Flex.Cols = 5
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT e.id, e.numerodocumento, e.nombre, e.nrolegajo, e.cuil FROM empleados as e WHERE e.eliminado = 0"
    If txtBuscar.Text <> "" Then
        SQL = SQL & " AND e." & cboBuscar.Text & " LIKE '%" & txtBuscar.Text & "%'"
    End If
    SQL = SQL & " ORDER BY e." & cboBuscar.Text
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        'Set Flex.DataSource = Recordset
        Do While Not Recordset.EOF
            Flex.TextMatrix(Flex.Rows - 1, 0) = Recordset!id
            Flex.TextMatrix(Flex.Rows - 1, 1) = Format(Recordset!numerodocumento, "00,000,000")
            Flex.TextMatrix(Flex.Rows - 1, 2) = Recordset!nombre
            Flex.TextMatrix(Flex.Rows - 1, 3) = Recordset!nrolegajo
            Flex.TextMatrix(Flex.Rows - 1, 4) = Format(Recordset!Cuil, "00-00000000-0")
            Flex.Rows = Flex.Rows + 1
            Recordset.MoveNext
        Loop
    Else
        Flex.Clear
        Flex.Rows = 2
    End If
    Flex.Rows = Flex.Rows - 1
    Recordset.Close
    
    ordenaFlex
    
End Sub

Private Sub Form_Resize()
    
    If Me.ScaleHeight < 2000 Or Me.ScaleWidth < 2000 Then
        Exit Sub
    End If

    Const Margen = 120

    Flex.Width = Me.ScaleWidth - 2 * Margen
    Flex.Height = Me.ScaleHeight - picBotones.Height - frmBuscar.Height - 4 * Margen

    picBotones.Left = Me.ScaleWidth - picBotones.Width - Margen
    picBotones.Top = Flex.Height + frmBuscar.Height + Margen * 3

End Sub

Private Sub txtBuscar_Change()
    
    Cargar
    
End Sub

' Instala el hook en el MSFlexGrid
Private Sub Flex_GotFocus()
    HookForm Flex
End Sub

' elimina el hook
Private Sub Flex_LostFocus()
    UnHookForm Flex
End Sub
