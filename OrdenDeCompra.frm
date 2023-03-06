VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form OrdenDeCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Órdenes de Compras"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
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
   ScaleHeight     =   4335
   ScaleWidth      =   8175
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
      Left            =   5280
      Picture         =   "OrdenDeCompra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Eliminar "
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "&Finalizar"
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
      Left            =   4320
      Picture         =   "OrdenDeCompra.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Finalizar Orden"
      Top             =   3600
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
      Left            =   3360
      Picture         =   "OrdenDeCompra.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Nueva Orden "
      Top             =   3600
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
      Left            =   6240
      Picture         =   "OrdenDeCompra.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Imprimir "
      Top             =   3600
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
      Left            =   7200
      Picture         =   "OrdenDeCompra.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Salir "
      Top             =   3600
      Width           =   855
   End
   Begin VB.CheckBox chkPendientes 
      Caption         =   "Sólo Pendientes"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   2640
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4657
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   3
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
      ScrollBars      =   2
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
      _Band(0).Cols   =   3
   End
End
Attribute VB_Name = "OrdenDeCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkPendientes_Click()
    
    Cargar
    
End Sub

Private Sub cmdEliminar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Or Flex.TextMatrix(Flex.Row, 3) <> "PENDIENTE" Then Exit Sub
    
    Dim IdOC As Single
    IdOC = Flex.TextMatrix(Flex.Row, 0)
    
    Select Case MsgBox("¿DESEA ELIMINAR ESTA ORDEN DE COMPRA? EL PROCESO ES IRREVERSIBLE.        ", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
        
        Case vbNo: Exit Sub
        
    End Select
    
    
    'Agrega a faltantes los artículos
    Set rsOCD = New ADODB.Recordset
    SQL = "SELECT * FROM ocomprasd WHERE idorden = " & IdOC
    rsOCD.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not rsOCD.BOF And Not rsOCD.EOF
        
        'Lo agrega a la tabla articulosfaltantes
        Set rsFaltante = New ADODB.Recordset
        SQL = "SELECT * FROM articulosfaltantes WHERE idart = " & rsOCD!idArt
        rsFaltante.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        If rsFaltante.BOF And rsFaltante.EOF Then
            rsFaltante.AddNew
            rsFaltante!idArt = rsOCD!idArt
            rsFaltante!idPro = getProveedorRecomendado(rsOCD!idArt) 'le asigna el proveedor recomendado
        End If
        
        rsFaltante!Cantidad = rsOCD!Cantidad
        rsFaltante.Update
        rsFaltante.Close
                
        rsOCD.MoveNext
    Loop
    rsOCD.Close
    
    
    Set rsOC = New ADODB.Recordset
    SQL = "DELETE FROM ocompras WHERE id = " & IdOC
    rsOC.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Set rsOCD = New ADODB.Recordset
    SQL = "DELETE FROM ocomprasd WHERE idorden = " & IdOC
    rsOCD.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Private Sub cmdFinalizar_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Or Flex.TextMatrix(Flex.Row, 3) <> "PENDIENTE" Then Exit Sub
    
    Dim IdOC As Single
    IdOC = Flex.TextMatrix(Flex.Row, 0)
    
    OrdenDeCompraD.Cargar IdOC
    
    Unload Me
    
End Sub

Private Sub cmdImprimir_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then Exit Sub
    
    Dim IdOC As Single
    IdOC = Flex.TextMatrix(Flex.Row, 0)
    
    imprimirOrdenDeCompra IdOC
    
End Sub


Private Sub cmdNuevo_Click()
    
    ArticulosFaltantes.Show
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    Cargar
    
End Sub

Sub Cargar()
    
    Set rsCargar = New ADODB.Recordset
    SQL = "SELECT oc.id, p.nombre, oc.fecha, estado FROM ocompras AS oc Inner Join proveedores AS p ON oc.idpro = p.id "
    If chkPendientes Then
        SQL = SQL & " WHERE oc.estado = 'PENDIENTE'"
    End If
    rsCargar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsCargar.BOF And Not rsCargar.EOF Then
        Set Flex.DataSource = rsCargar
    Else
        Flex.Clear
        Flex.Rows = 2
    End If
    rsCargar.Close
    
    OrdenaFlex
    
End Sub

Private Sub OrdenaFlex()
    
    Flex.FormatString = "Código|Proveedor|Fecha|Estado"
    Flex.ColWidth(0) = 1000
    Flex.ColWidth(1) = 4000
    Flex.ColWidth(2) = 1300
    Flex.ColWidth(3) = 1300
    
End Sub
