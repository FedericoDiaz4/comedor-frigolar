VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ArticulosPrecios 
   Caption         =   "Artículos"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13815
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
   ScaleHeight     =   6375
   ScaleWidth      =   13815
   WindowState     =   2  'Maximized
   Begin VB.Frame frmOk 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   12240
      TabIndex        =   25
      Top             =   5280
      Width           =   2775
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   1920
         Picture         =   "ArticulosPrecios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   " Salir "
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   615
         Left            =   960
         Picture         =   "ArticulosPrecios.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton cmdRevisada 
         Caption         =   "&Revisada"
         Height          =   615
         Left            =   0
         Picture         =   "ArticulosPrecios.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblRevisada 
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
         TabIndex        =   26
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frmEliminar 
      Caption         =   "Eliminar"
      Height          =   975
      Left            =   6720
      TabIndex        =   36
      Top             =   5280
      Width           =   1095
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
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
         Picture         =   "ArticulosPrecios.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   " Eliminar"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox cboEdit 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   10
      Left            =   11160
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboEdit 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   9
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboEdit 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   8
      ItemData        =   "ArticulosPrecios.frx":1628
      Left            =   8040
      List            =   "ArticulosPrecios.frx":162A
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   2
      Left            =   240
      TabIndex        =   29
      Text            =   "ARTICULO"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame frmProveedor 
      Caption         =   "Asignar a otro proveedor"
      Height          =   975
      Left            =   7920
      TabIndex        =   22
      Top             =   5280
      Width           =   4095
      Begin VB.CommandButton cmdModificarPro 
         Caption         =   "&Aplicar"
         Height          =   615
         Left            =   3120
         Picture         =   "ArticulosPrecios.frx":162C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboProveedorAsignar 
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "ArticulosPrecios.frx":1BB6
         Left            =   120
         List            =   "ArticulosPrecios.frx":1BB8
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmModificar 
      Caption         =   "Modificar"
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   6495
      Begin VB.CommandButton cmdModificarPrecio 
         Caption         =   "&Aplicar"
         Height          =   615
         Left            =   5520
         Picture         =   "ArticulosPrecios.frx":1BBA
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtValor 
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
         Left            =   3720
         TabIndex        =   20
         Top             =   480
         Width           =   1680
      End
      Begin VB.ComboBox cboTipo 
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "ArticulosPrecios.frx":2144
         Left            =   1920
         List            =   "ArticulosPrecios.frx":2151
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   480
         Width           =   1680
      End
      Begin VB.ComboBox cboAModificar 
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "ArticulosPrecios.frx":2179
         Left            =   120
         List            =   "ArticulosPrecios.frx":219B
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   1680
      End
      Begin VB.ComboBox cboValor 
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "ArticulosPrecios.frx":2226
         Left            =   3720
         List            =   "ArticulosPrecios.frx":2233
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   1680
      End
   End
   Begin VB.Frame frmFiltro 
      Caption         =   "Filtrar"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Filtrar"
         Height          =   615
         Left            =   11640
         Picture         =   "ArticulosPrecios.frx":225B
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Salir "
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboRubro 
         Height          =   360
         Left            =   9840
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtHasta 
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
         Left            =   8040
         TabIndex        =   9
         Top             =   480
         Width           =   1680
      End
      Begin VB.TextBox txtDesde 
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
         Left            =   6240
         TabIndex        =   8
         Top             =   480
         Width           =   1680
      End
      Begin VB.ComboBox cboProveedor 
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "ArticulosPrecios.frx":27E5
         Left            =   1920
         List            =   "ArticulosPrecios.frx":27E7
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtCodProveedor 
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
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1680
      End
      Begin VB.Label Label4 
         Caption         =   "Rubro"
         Height          =   255
         Left            =   9840
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Artículo Hasta"
         Height          =   255
         Left            =   8040
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo Desde"
         Height          =   255
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Código Proveedor"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   4
      Left            =   3360
      TabIndex        =   14
      Text            =   "PRECIO"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   15
      Text            =   "DESCUENTOS"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   16
      Text            =   "IVA"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   3
      Left            =   1800
      TabIndex        =   13
      Text            =   "CODIGO"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   15658734
      ForeColor       =   0
      Cols            =   8
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
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
End
Attribute VB_Name = "ArticulosPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idProveedor As Integer
Dim rsCargar As ADODB.Recordset
Dim filaActual As Integer

Private Sub cboAModificar_Click()
    
    cboTipo.Text = "ESTABLECER"
    
    Select Case cboAModificar.Text
        
        Case "PRECIO"
            txtValor.Visible = True
            cboValor.Visible = False
        
        Case "DESCUENTO", "IVA"
            txtValor.Visible = True
            cboValor.Visible = False
        
        Case "RUBRO"
            txtValor.Visible = False
            cboValor.Visible = True
            CargaCombo "rubros", "nombre", "nombre", cboValor
        
        Case "SUBRUBRO"
            txtValor.Visible = False
            cboValor.Visible = True
            CargaCombo "subrubros", "nombre", "nombre", cboValor
        
        Case "MONEDA"
            txtValor.Visible = False
            cboValor.Visible = True
            cboValor.Clear
            cboValor.AddItem "$"
            cboValor.AddItem "U$S"
            cboValor.AddItem "€"
                        
    End Select
    
End Sub

Private Sub cboEdit_GotFocus(Index As Integer)
    
    'Pinta la fila de color
    columnaNueva = Flex.Col
    filaNueva = Flex.Row
    
    Flex.Col = 1
5    Flex.Row = filaActual
    Flex.CellBackColor = RGB(238, 238, 238)
    
    Flex.Row = filaNueva
    Flex.CellBackColor = RGB(200, 250, 200)
    Flex.Col = columnaNueva
    filaActual = filaNueva
    
End Sub

Private Sub cboEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        Case 13
            ActualizarReg Index
            If Flex.Row < Flex.Rows - 1 Then
                Flex.Row = Flex.Row + 1
                Flex.Col = Index
                Flex_Click
            End If
            
        Case vbKeyLeft
            ActualizarReg Index
            If Flex.Col = 2 And Flex.Row > 1 Then
                Flex.Row = Flex.Row - 1
                Flex.Col = 6
            Else
                Flex.Col = Flex.Col - 1
            End If
            Flex_Click
        
        Case vbKeyRight
            ActualizarReg Index
            If Flex.Col = 6 And Flex.Row < Flex.Rows - 1 Then
                Flex.Row = Flex.Row + 1
                Flex.Col = 2
            Else
                Flex.Col = Flex.Col + 1
            End If
            Flex_Click
            
    End Select
End Sub

Private Sub cboProveedor_Click()
    
    If cboProveedor.ListIndex <> -1 Then
        txtCodProveedor.Text = getData(cboProveedor.ItemData(cboProveedor.ListIndex), "codigo", "proveedores")
    End If
    
End Sub

Private Sub cboProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cboRubro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmdEliminar_Click()
    
    Dim idArtC As Long
    
    Dim sMensaje As String
    sMensaje = "ESTÁ POR ELIMINAR TODOS LOS ARTÍCULOS SELECCIONADOS, LA OPERACIÓN ES IRREVERSIBLE         "
    Select Case MsgBox(sMensaje _
                       & vbCrLf & "¿DESEA CONTINUAR?           " _
                       , vbYesNo Or vbQuestion Or vbDefaultButton1, "ATENCIÓN")
    
        Case vbNo: Exit Sub
    End Select
    
    If Flex.Rows <= 2 And Flex.TextMatrix(1, 0) = "" Then
        Call MsgBox("NO HAY ARTÍCULOS SELECCIONADOS", vbExclamation, App.Title)
        Exit Sub
    End If
    
    VerificarConexion
    
    'Oculta los texts de edición
    For i = 2 To 6
        txtEdit(i).Visible = False
    Next i
    For i = 8 To 10
        cboEdit(i).Visible = False
    Next i
    
    'Muestra la barra
    zMain.pBar.Value = 0
    zMain.sBar.Height = 0
    zMain.pBar.Height = 255
    zMain.pBar.Max = Flex.Rows - 1
    
    'Recorre el Flex y modifica los registros
    For i = 1 To Flex.Rows - 1
        
        idArtC = Flex.TextMatrix(i, 0)
        
        'Actualiza la tabla de la base
        Set rsUpd = New ADODB.Recordset
        SQL = "DELETE FROM articulosc WHERE id = " & idArtC & ";"
        rsUpd.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        zMain.pBar.Value = i
        
    Next i
    
    'Oculta la barra
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    zMain.sBar.Height = 255
    
    Cargar
    
End Sub

Private Sub cmdGuardar_Click()
    
    Dim rsTemp  As ADODB.Recordset
    
    Select Case MsgBox("¿ESTÁ SEGURO QUE DESEA GUARDAR LOS CAMBIOS SOBRE LA LISTA?          ", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
        
        Case vbNo: Exit Sub
        
    End Select
    
    VerificarConexion
    
    Set rsTemp = New ADODB.Recordset
    SQL = "SELECT * FROM temp_proxart_" & nTmp & ""
    rsTemp.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        Do While Not rsTemp.EOF
            
            Set Recordset = New ADODB.Recordset
            SQL = "UPDATE articulos Inner Join articulosc ON articulos.id = articulosc.idart "
            SQL = SQL & "SET nombre = '" & rsTemp!col2 & "', articulosc.codigo = '" & rsTemp!col3 & "', articulosc.precio = '" & rsTemp!col4 & "', articulosc.descuento = '" & rsTemp!col5 & "', articulosc.iva = '" & rsTemp!col6 & "', articulosc.total = '" & rsTemp!col7 & "', articulos.rubro = '" & rsTemp!col8 & "', articulos.subrubro = '" & rsTemp!col9 & "', articulosc.moneda = '" & rsTemp!col10 & "' "
            SQL = SQL & "WHERE articulosc.id = " & rsTemp!col0 & ""
            Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
            
            rsTemp.MoveNext
        Loop
        
    rsTemp.Close
    
    Call MsgBox("CAMBIOS GUARDADOS          ", vbInformation, App.Title)
    
End Sub

Private Sub cmdModificarPrecio_Click()
    
    Dim idArtC As Long
    Dim sMensaje As String
    Dim sPrecio As String
    Dim IVA As Double
    Dim sDesc As String
    Dim sTotal As String
    Dim Total As Double
    
    sMensaje = "ESTÁ POR "
    
    If cboTipo.ListIndex = 0 Then
        sMensaje = sMensaje & "AUMENTAR EL PRECIO DE LOS ARTÍCULOS SELECCIONADOS UN " & txtValor.Text & "%"
    ElseIf cboTipo.ListIndex = 1 Then
        sMensaje = sMensaje & "REDUCIR EL PRECIO DE LOS ARTÍCULOS SELECCIONADOS UN " & txtValor.Text & "%"
    ElseIf cboTipo.ListIndex = 2 Then
        If txtValor.Visible Then
            sMensaje = sMensaje & "ESTABLECER TODOS LOS " & cboAModificar.Text & "S DE LOS ARTÍCULOS SELECCIONADOS EN " & txtValor.Text
        Else
            sMensaje = sMensaje & "ESTABLECER TODOS LOS " & cboAModificar.Text & "S DE LOS ARTÍCULOS SELECCIONADOS EN " & cboValor.Text
        End If
    End If
    sMensaje = sMensaje & "         "
    
    Select Case MsgBox(sMensaje _
                       & vbCrLf & "¿DESEA CONTINUAR?           " _
                       , vbYesNo Or vbQuestion Or vbDefaultButton1, "ATENCIÓN")
    
        Case vbNo: Exit Sub
    End Select
    
    If Flex.Rows <= 2 And Flex.TextMatrix(1, 0) = "" Then
        Call MsgBox("NO HAY ARTÍCULOS SELECCIONADOS", vbExclamation, App.Title)
        Exit Sub
    End If
    
    'Oculta los texts de edición
    For i = 2 To 6
        txtEdit(i).Visible = False
    Next i
    For i = 8 To 10
        cboEdit(i).Visible = False
    Next i
    
    'Muestra la barra
    zMain.pBar.Value = 0
    zMain.sBar.Height = 0
    zMain.pBar.Height = 255
    zMain.pBar.Max = Flex.Rows - 1
    
    VerificarConexion
    
    'Recorre el Flex y modifica los registros
    For i = 1 To Flex.Rows - 1
        
        idArtC = Flex.TextMatrix(i, 0)
        sPrecio = Flex.TextMatrix(i, 4)
        sDesc = Flex.TextMatrix(i, 5)
        IVA = Flex.TextMatrix(i, 6)
        Rubro = Flex.TextMatrix(i, 8)
        subrubro = Flex.TextMatrix(i, 9)
        moneda = Flex.TextMatrix(i, 10)
        
        If cboAModificar.Text = "PRECIO" Then
            
            'Calcula el nuevo precio
            If cboTipo.ListIndex = 0 Then
                'sPrecio = Format(sPrecio + (CDec(sPrecio) * CDec(txtValor.Text) / 100), "0.000")
                sPrecio = CalculaRecargos(sPrecio, txtValor.Text)
            ElseIf cboTipo.ListIndex = 1 Then
                'sPrecio = Format(sPrecio - (CDec(sPrecio) * CDec(txtValor.Text) / 100), "0.000")
                sPrecio = CalculaDescuentos(sPrecio, txtValor.Text)
            ElseIf cboTipo.ListIndex = 2 Then
                sPrecio = Format(CDec(txtValor.Text), "0.000")
            End If
            
        ElseIf cboAModificar.Text = "DESCUENTO" Then
            
            sDesc = txtValor.Text
            
        ElseIf cboAModificar.Text = "IVA" Then
            
            IVA = txtValor.Text
            
        ElseIf cboAModificar.Text = "RUBRO" Then
            
            Rubro = cboValor.Text
            
        ElseIf cboAModificar.Text = "SUBRUBRO" Then
            
            subrubro = cboValor.Text
            
        ElseIf cboAModificar.Text = "MONEDA" Then
            
            moneda = cboValor.Text
            
        End If
        
        'Calcula el total
        Total = CalculaDescuentos(sPrecio, sDesc)
        Total = Format(Total + ((Total * IVA) / 100), "0.000")
        sPrecio = Format(sPrecio, "0.000")
        sTotal = Format(Total, "0.000")
        
        'Actualiza la tabla de la base
        Set rsUpd = New ADODB.Recordset
        SQL = "UPDATE temp_proxart_" & nTmp & " SET col4 = '" & sPrecio & "', col5 = '" & sDesc & "', col6 = '" & IVA & "', col7 = '" & sTotal & "', col8 = '" & Rubro & "', col9 = '" & subrubro & "', col10 = '" & moneda & "' WHERE col0 = " & idArtC & ";"
        rsUpd.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        zMain.pBar.Value = i
        
    Next i
    
    'Oculta la barra
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    zMain.sBar.Height = 255
        
    cboAModificar.ListIndex = 0
    cboTipo.ListIndex = -1
    txtValor.Text = ""
    
    MuestraTemp
    
End Sub

Private Sub cmdModificarPro_Click()
    
    Dim idPro As Integer
    Dim sMensaje As String
    
    If cboProveedorAsignar.ListIndex = -1 Then Exit Sub
    idPro = cboProveedorAsignar.ItemData(cboProveedorAsignar.ListIndex)
    
    sMensaje = "¿DESEA ASIGNAR LOS ARTÍCULOS SELECCIONADOS A " & cboProveedorAsignar.Text & "?"
    
    Select Case MsgBox(sMensaje _
                       & vbCrLf & "¿DESEA CONTINUAR?           " _
                       , vbYesNo Or vbQuestion Or vbDefaultButton1, "ATENCIÓN")
    
        Case vbNo: Exit Sub
    End Select
    
    If Flex.Rows <= 2 And Flex.TextMatrix(1, 0) = "" Then
        Call MsgBox("NO HAY ARTÍCULOS SELECCIONADOS", vbExclamation, App.Title)
        Exit Sub
    End If
    
    'Muestra la barra
    zMain.pBar.Value = 0
    zMain.sBar.Height = 0
    zMain.pBar.Height = 255
    zMain.pBar.Max = Flex.Rows - 1
    
    VerificarConexion
    
    'Recorre el Flex y genera los nuevos registros
    For i = 1 To Flex.Rows - 1
        
        'Obtiene el id del Artículo
        Set Recordset = New ADODB.Recordset
        SQL = "SELECT idart FROM articulosc WHERE id = " & Flex.TextMatrix(i, 0)
        Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        'Verifica si el proveedor ya tiene el artículo
        Set rsUpd = New ADODB.Recordset
        SQL = "SELECT idart,idpro FROM articulosc WHERE idart = " & Recordset!idArt & " AND idpro = '" & idPro & "'"
        rsUpd.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If rsUpd.BOF And rsUpd.EOF Then
            rsUpd.AddNew
        End If
        
        rsUpd!idArt = Recordset!idArt
        rsUpd!idPro = idPro
        rsUpd.Update
        rsUpd.Close
        
        zMain.pBar.Value = i
        
    Next i
    
    'Oculta la barra
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    zMain.sBar.Height = 255
    
    cboProveedor.Text = cboProveedorAsignar.Text
    Cargar
    
End Sub

Private Sub cmdOk_Click()
    
    Cargar
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub cmdRevisada_Click()
    
    If cboProveedor.Text = "" Then
        Call MsgBox("Debe ingresar un proveedor     ", vbExclamation, App.Title)
        OrdenaFlex
        Exit Sub
    End If
    
    VerificarConexion
    
    Set rsPro = New ADODB.Recordset
    SQL = "UPDATE proveedores SET listarevisada = '" & Format(Date, "yyyy-mm-dd") & "' WHERE id = " & cboProveedor.ItemData(cboProveedor.ListIndex)
    rsPro.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Sub Flex_Click()
    
    On Error Resume Next
    
    'Posiciona los TextBox
    For i = 2 To 6
        txtEdit(i).Visible = False
        txtEdit(i).Top = Flex.Top + Flex.CellTop
        txtEdit(i).Left = Flex.ColPos(i) + 175
        txtEdit(i).Width = Flex.ColWidth(i) - 60
        txtEdit(i).Text = Flex.TextMatrix(Flex.Row, i)
    Next i
    For i = 8 To 10
        cboEdit(i).Visible = False
        cboEdit(i).Top = Flex.Top + Flex.CellTop - 50
        cboEdit(i).Left = Flex.ColPos(i) + 130
        cboEdit(i).Width = Flex.ColWidth(i) + 30
        cboEdit(i).Text = Flex.TextMatrix(Flex.Row, i)
    Next i
    
    'Le da el foco al correspondiente
    If 2 <= Flex.Col And Flex.Col <= 6 Then
        txtEdit(Flex.Col).Visible = True
        txtEdit(Flex.Col).SetFocus
    End If
    If 8 <= Flex.Col And Flex.Col <= 10 Then
        cboEdit(Flex.Col).Visible = True
        cboEdit(Flex.Col).SetFocus
    End If
    
End Sub

Private Sub Flex_DblClick()
    
    i = Flex.TextMatrix(Flex.Row, 0)
    
    codigoArt = Flex.TextMatrix(Flex.Row, 1)
    
    If i = "" Then Exit Sub
    
    Select Case MsgBox("¿DESEA ELIMINAR EL ARTÍCULO " & codigoArt & " DEL PROVEEDOR SELECCIONADO?" _
                       & vbCrLf & "EL CAMBIO ES IRREVERSIBLE" _
                       , vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
    
    
        Case vbNo: Exit Sub
    
    End Select
    
    VerificarConexion
    
    Set Recordset = New ADODB.Recordset
    SQL = "DELETE FROM temp_proxart_" & nTmp & " WHERE col0 = '" & i & "'"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Set Recordset = New ADODB.Recordset
    SQL = "DELETE FROM articulosc WHERE id = '" & i & "'"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    MuestraTemp
    
End Sub

Private Sub Flex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If y < 270 And x < 5800 Then
        MuestraTemp "col" & Flex.Col
    End If
    
End Sub

Private Sub Flex_Scroll()
    
    'Oculta los texts de edición
    For i = 2 To 6
        txtEdit(i).Visible = False
    Next i
    For i = 8 To 10
        cboEdit(i).Visible = False
    Next i
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    
    'Carga los Combos
    CargaCombo "proveedores", "nombre", "nombre", cboProveedor
    CargaCombo "proveedores", "nombre", "nombre", cboProveedorAsignar
    CargaCombo "rubros", "nombre", "nombre", cboRubro
    CargaCombo "rubros", "nombre", "nombre", cboEdit(8)
    CargaCombo "subrubros", "nombre", "nombre", cboEdit(9)
    
    cboEdit(10).AddItem "$"
    cboEdit(10).AddItem "U$S"
    cboEdit(10).AddItem "€"
    
    cboAModificar.ListIndex = 0
    
    'Carga el Flex
    Cargar
    
End Sub

Sub OrdenaFlex()
    
    Flex.FormatString = "id|Código|Artículo|Código|Precio|Descuentos|IVA|Total|Rubro|SubRubro|Moneda|Proveedor|Predeterminado||||"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 1000
    Flex.ColWidth(2) = 4800
    Flex.ColWidth(3) = 1400
    Flex.ColWidth(4) = 1200
    Flex.ColWidth(5) = 1500
    Flex.ColWidth(6) = 1200
    Flex.ColWidth(7) = 1200
    Flex.ColWidth(8) = 1400
    Flex.ColWidth(9) = 1400
    Flex.ColWidth(10) = 1000
    Flex.ColWidth(11) = 3870
    Flex.ColWidth(12) = 1600
    Flex.ColWidth(13) = 0
    Flex.ColWidth(14) = 0
    Flex.ColWidth(15) = 0
    Flex.ColWidth(16) = 0
    
    Flex.ColAlignment(3) = 2
    
End Sub

Sub Cargar()
    
    If cboProveedor.Text = "" And txtDesde.Text = "" And txtHasta.Text = "" And cboRubro.Text = "" Then
        OrdenaFlex
        Exit Sub
    End If
    
    VerificarConexion
    
    'Vacía la tabla temporal
    BorraTemp
    
    'Copia el resultado a la tabla temporal
    Set rsCargar = New ADODB.Recordset
    SQL = "INSERT INTO temp_proxart_" & nTmp & " (col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12) "
    SQL = SQL & "SELECT c.id, a.codigo, a.nombre, c.codigo, c.precio, c.descuento, c.iva, c.total, a.rubro, a.subrubro, c.moneda, p.nombre, c.predet FROM articulosc AS c Inner Join articulos AS a ON c.idart = a.id Inner Join proveedores AS p ON c.idpro = p.id "
    
    'Aplica los filtros
    If cboProveedor.Text <> "" Then
        SQL = SQL & "AND c.idpro = " & cboProveedor.ItemData(cboProveedor.ListIndex) & " "
    End If
    If txtDesde.Text <> "" And txtHasta.Text <> "" Then
        SQL = SQL & "AND a.codigo BETWEEN '" & txtDesde.Text & "' AND '" & txtHasta & "' "
    End If
    If cboRubro.Text <> "" Then
        SQL = SQL & "AND a.rubro = '" & Trim(cboRubro.Text) & "' "
    End If
    If Orden = "" Then
        SQL = SQL & "ORDER BY a.codigo;"
    Else
        SQL = SQL & "ORDER BY " & Orden & ";"
    End If
    
    'Muestra el resultado en el flex
    rsCargar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    MuestraTemp
    
    'Muestra la última actualización de la lista
    If cboProveedor.Text <> "" Then
        
        Set rsPro = New ADODB.Recordset
        SQL = "SELECT listarevisada FROM proveedores WHERE id = " & cboProveedor.ItemData(cboProveedor.ListIndex)
        rsPro.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not IsNull(rsPro!listarevisada) Then
            lblRevisada.Caption = rsPro!listarevisada
        Else
            lblRevisada.Caption = ""
        End If
        If rsPro!listarevisada = Date Then
            cmdRevisada.Enabled = False
        Else
            cmdRevisada.Enabled = True
        End If
        rsPro.Close
        
    Else
        
        cmdRevisada.Enabled = False
        lblRevisada.Caption = ""
        
    End If
    
End Sub

Sub BorraTemp()
    
    'Vacía la tabla temporal
    Set rsDel = New ADODB.Recordset
    SQL = "DELETE FROM temp_proxart_" & nTmp & ""
    rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
End Sub

Sub MuestraTemp(Optional Orden As String)
    
    'Carga la tabla temporal en el flex
    
    VerificarConexion
    cmdOk.Enabled = False
    
    Set rsCargar = New ADODB.Recordset
    SQL = "SELECT * FROM temp_proxart_" & nTmp & " "
    If Orden <> "" Then
    SQL = SQL & "ORDER BY " & Orden
    End If
    rsCargar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Not rsCargar.BOF And Not rsCargar.EOF Then
        
        Set Flex.DataSource = rsCargar
        OrdenaFlex
        
        'Color de celdas
        For i = 1 To Flex.Rows - 1
            For a = 2 To 10
                If a <> 7 Then
                    Flex.Row = i
                    Flex.Col = a
                    Flex.CellBackColor = vbWhite
                End If
            Next a
        Next i
        
    Else
        Flex.Clear
        Flex.Rows = 2
        OrdenaFlex
    End If
    rsCargar.Close
    
    cmdOk.Enabled = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error Resume Next
    rsCargar.Close
    
End Sub

Private Sub Form_Resize()
    
    If Me.ScaleHeight = 0 Then Exit Sub

    frmFiltro.Width = Me.ScaleWidth - 240
    Flex.Width = Me.ScaleWidth - 240

    Flex.Height = Me.ScaleHeight - frmFiltro.Height - frmModificar.Height - 500
    frmOk.Top = Flex.Top + Flex.Height + 240
    frmOk.Left = Flex.Width - frmOk.Width
    frmModificar.Top = Flex.Top + Flex.Height + 120
    frmEliminar.Top = Flex.Top + Flex.Height + 120
    frmProveedor.Top = Flex.Top + Flex.Height + 120
    
End Sub

Private Sub txtCodProveedor_Change()
    
    VerificarConexion
    
    Set rsPro = New ADODB.Recordset
    SQL = "SELECT * FROM proveedores WHERE codigo = '" & txtCodProveedor.Text & "' AND eliminado <> 1"
    rsPro.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsPro.BOF And Not rsPro.EOF Then
        If rsPro!nombre <> "" Then
            cboProveedor.Text = rsPro!nombre
        End If
        rsPro.Close
    Else
        cboProveedor.ListIndex = -1
    End If
    
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDesde_LostFocus()
    
    If txtDesde.Text <> "" Then
        txtDesde.Text = Format(txtDesde.Text, "000000")
    End If
    
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    
    'Pinta la fila de color
    columnaNueva = Flex.Col
    filaNueva = Flex.Row
    
    Flex.Col = 1
    Flex.Row = filaActual
    Flex.CellBackColor = RGB(238, 238, 238)
    
    Flex.Row = filaNueva
    Flex.CellBackColor = RGB(200, 250, 200)
    Flex.Col = columnaNueva
    filaActual = filaNueva
    
    txtEdit(Index).SelStart = 0
    txtEdit(Index).SelLength = Len(txtEdit(Index).Text)
    
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        Case vbKeyDown, 13
            ActualizarReg Index
            If Flex.Row < Flex.Rows - 1 Then
                Flex.Row = Flex.Row + 1
                Flex.Col = Index
                Flex_Click
            End If
            
        Case vbKeyUp
            ActualizarReg Index
            If Flex.Row > 1 Then
                Flex.Row = Flex.Row - 1
                Flex.Col = Index
                Flex_Click
            End If
        
        Case vbKeyLeft
            ActualizarReg Index
            If Flex.Col = 2 And Flex.Row > 1 Then
                Flex.Row = Flex.Row - 1
                Flex.Col = 6
            Else
                Flex.Col = Flex.Col - 1
            End If
            Flex_Click
        
        Case vbKeyRight
            ActualizarReg Index
            If Flex.Col = 6 And Flex.Row < Flex.Rows - 1 Then
                Flex.Row = Flex.Row + 1
                Flex.Col = 2
            Else
                Flex.Col = Flex.Col + 1
            End If
            Flex_Click
            
    End Select
    
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index > 3 And Index < 7 Then
        CambiaPunto txtEdit(Index), KeyAscii, "+"
    End If
    
End Sub

Sub ActualizarReg(Index As Integer)
    
    Dim idArtC As Long
    
    'Guarda el ID
    idArtC = Flex.TextMatrix(Flex.Row, 0)
    
    'Sólo actualiza si se modifica el TextBox
    If Index < 7 Then
        If txtEdit(Index).Text = Flex.TextMatrix(Flex.Row, Flex.Col) Then
            Exit Sub
        End If
    Else
        If cboEdit(Index).Text = Flex.TextMatrix(Flex.Row, Flex.Col) Then
            Exit Sub
        End If
    End If
    
    'Calcula el total
    If Index > 3 And Index < 7 Then
        If txtEdit(4).Text <> "" And txtEdit(5).Text <> "" And txtEdit(6).Text <> "" Then
            Tot = CalculaDescuentos(txtEdit(4).Text, txtEdit(5).Text)
            Tot = Format(Tot + ((Tot * CDec(txtEdit(6).Text)) / 100), "0.000")
        Else
            Tot = "0,000"
        End If
    Else
        Tot = Flex.TextMatrix(Flex.Row, 7)
    End If
    
    VerificarConexion
    
    'Actualiza la tabla de la base
    Set rsUpd = New ADODB.Recordset
    SQL = "UPDATE temp_proxart_" & nTmp & " SET col7 = '" & Tot & "', "
    
    Select Case Index
    Case 2: SQL = SQL & "col2 = '" & txtEdit(Index).Text & "' "
    Case 3: SQL = SQL & "col3 = '" & txtEdit(Index).Text & "' "
    Case 4: SQL = SQL & "col4 = '" & Format(txtEdit(Index).Text, "0.000") & "' "
    Case 5: SQL = SQL & "col5 = '" & txtEdit(Index).Text & "' "
    Case 6: SQL = SQL & "col6 = '" & txtEdit(Index).Text & "' "
    Case 8: SQL = SQL & "col8 = '" & cboEdit(Index).Text & "' "
    Case 9: SQL = SQL & "col9 = '" & cboEdit(Index).Text & "' "
    Case 10: SQL = SQL & "col10 = '" & cboEdit(Index).Text & "' "
    End Select
    
    SQL = SQL & "WHERE col0 = " & idArtC & ";"
    rsUpd.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Actualiza la información del flex
    If Index = 4 Then
        Flex.TextMatrix(Flex.Row, Flex.Col) = Format(txtEdit(Index).Text, "0.000")
    ElseIf Index < 7 Then
        Flex.TextMatrix(Flex.Row, Flex.Col) = txtEdit(Index).Text
    Else
        Flex.TextMatrix(Flex.Row, Flex.Col) = cboEdit(Index).Text
    End If
    Flex.TextMatrix(Flex.Row, 7) = Tot
    
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtHasta_LostFocus()
    
    If txtHasta.Text <> "" Then
        txtHasta.Text = Format(txtHasta.Text, "000000")
    End If
    
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    
    CambiaPunto txtValor, KeyAscii, "+"
    
End Sub
