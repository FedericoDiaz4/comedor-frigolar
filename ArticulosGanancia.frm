VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ArticulosGanancia 
   Caption         =   "Listado de Precios y Ganancia"
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
      TabIndex        =   23
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
      Text            =   "% CTA CTE"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
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
      Text            =   "% MOSTRADOR"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame frmFiltro 
      Caption         =   "Filtrar"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.CheckBox chkPred 
         Caption         =   "Sólo Predeterminados"
         Height          =   495
         Left            =   5520
         TabIndex        =   7
         Top             =   330
         Value           =   1  'Checked
         Width           =   1815
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
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1680
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
         Left            =   1920
         TabIndex        =   5
         Top             =   480
         Width           =   1680
      End
      Begin VB.ComboBox cboRubro 
         Height          =   360
         Left            =   3720
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Filtrar"
         Height          =   615
         Left            =   7440
         Picture         =   "ArticulosGanancia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " Filtrar "
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo Desde"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Artículo Hasta"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Rubro"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmOk 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   11880
      TabIndex        =   20
      Top             =   5280
      Width           =   1815
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
         Left            =   960
         Picture         =   "ArticulosGanancia.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   " Salir "
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
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
         Picture         =   "ArticulosGanancia.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   " Guardar "
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
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frmModificar 
      Caption         =   "Modificar"
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Width           =   4695
      Begin VB.ComboBox cboAModificar 
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "ArticulosGanancia.frx":109E
         Left            =   120
         List            =   "ArticulosGanancia.frx":10A8
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   1680
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   18
         Top             =   480
         Width           =   1680
      End
      Begin VB.CommandButton cmdModificarPrecio 
         Caption         =   "&Aplicar"
         Height          =   615
         Left            =   3720
         Picture         =   "ArticulosGanancia.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   " Aplicar "
         Top             =   240
         Width           =   855
      End
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
      TabIndex        =   13
      Text            =   "ARTICULO"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboEdit 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   6
      ItemData        =   "ArticulosGanancia.frx":164E
      Left            =   6480
      List            =   "ArticulosGanancia.frx":1650
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboEdit 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   7
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboEdit 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   8
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   16777215
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
Attribute VB_Name = "ArticulosGanancia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idProveedor As Integer
Dim rsCargar As ADODB.Recordset

Private Sub cboAModificar_Click()
    
    If cboAModificar.Text = "DESCUENTO" Or cboAModificar = "IVA" Then
        cboTipo.Text = "ESTABLECER"
    End If
    
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
            If Flex.Col = 4 And Flex.Row > 1 Then
                Flex.Col = 2
            ElseIf Flex.Col = 2 And Flex.Row > 1 Then
                Flex.Row = Flex.Row - 1
                Flex.Col = 8
            Else
                Flex.Col = Flex.Col - 1
            End If
            Flex_Click
            
        Case vbKeyRight
            ActualizarReg Index
            If Flex.Col = 2 And Flex.Row < Flex.Rows - 1 Then
                Flex.Col = 4
            ElseIf Flex.Col = 8 And Flex.Row < Flex.Rows - 1 Then
                Flex.Row = Flex.Row + 1
                Flex.Col = 2
            Else
                Flex.Col = Flex.Col + 1
            End If
            Flex_Click
            
    End Select
    
End Sub

Private Sub cboRubro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub chkPred_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmdGuardar_Click()
    
    Dim rsTemp  As ADODB.Recordset
    
    Select Case MsgBox("¿ESTÁ SEGURO QUE DESEA GUARDAR LOS CAMBIOS SOBRE LA LISTA?          ", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
        
        Case vbNo: Exit Sub
        
    End Select
    
    Set rsTemp = New ADODB.Recordset
    SQL = "SELECT * FROM temp_proxart_" & nTmp & ""
    rsTemp.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        Do While Not rsTemp.EOF
            
            Set Recordset = New ADODB.Recordset
            SQL = "UPDATE articulos Inner Join articulosc ON articulos.id = articulosc.idart "
            SQL = SQL & "SET nombre = '" & rsTemp!col2 & "', articulosc.codigo = '" & rsTemp!col3 & "', articulos.porcventa1 = '" & rsTemp!col4 & "', articulos.porcventa2 = '" & rsTemp!col5 & "', articulos.rubro = '" & rsTemp!col6 & "', articulos.subrubro = '" & rsTemp!col7 & "', articulosc.moneda = '" & rsTemp!col8 & "' "
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
    
    sMensaje = sMensaje & "ESTÁ POR ESTABLECER TODOS LOS PORCENTAJES DE LOS ARTÍCULOS SELECCIONADOS EN " & txtValor.Text & "         "
    
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
    For i = 2 To 5
        txtEdit(i).Visible = False
    Next i
    For i = 6 To 8
        cboEdit(i).Visible = False
    Next i
    
    'Recorre el Flex y modifica los registros
    For i = 1 To Flex.Rows - 1
        
        idArtC = Flex.TextMatrix(i, 0)
        pMost = Flex.TextMatrix(i, 4)
        pCCte = Flex.TextMatrix(i, 5)
        
        If cboAModificar.Text = "% MOSTRADOR" Then
            
            pMost = txtValor.Text
            
        ElseIf cboAModificar.Text = "% CTA CTE" Then
            
            pCCte = txtValor.Text
            
        End If
        
        'Actualiza la tabla de la base
        Set rsUpd = New ADODB.Recordset
        SQL = "UPDATE temp_proxart_" & nTmp & " SET col4 = '" & pMost & "', col5 = '" & pCCte & "' WHERE col0 = " & idArtC & ";"
        rsUpd.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
    Next i
    
    cboAModificar.ListIndex = 0
    txtValor.Text = ""
    
    MuestraTemp
    
End Sub

Private Sub cmdOk_Click()
    
    Cargar
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Sub Flex_Click()
    
    On Error Resume Next
    
    'Posiciona los TextBox
    For i = 2 To 5
        txtEdit(i).Visible = False
        txtEdit(i).Top = Flex.Top + Flex.CellTop
        txtEdit(i).Left = Flex.ColPos(i) + 175
        txtEdit(i).Width = Flex.ColWidth(i) - 60
        txtEdit(i).Text = Flex.TextMatrix(Flex.Row, i)
    Next i
    For i = 6 To 8
        cboEdit(i).Visible = False
        cboEdit(i).Top = Flex.Top + Flex.CellTop - 50
        cboEdit(i).Left = Flex.ColPos(i) + 130
        cboEdit(i).Width = Flex.ColWidth(i) + 30
        cboEdit(i).Text = Flex.TextMatrix(Flex.Row, i)
    Next i
    
    'Le da el foco al correspondiente
    If 2 <= Flex.Col And Flex.Col <> 3 And Flex.Col <= 5 Then
        txtEdit(Flex.Col).Visible = True
        txtEdit(Flex.Col).SetFocus
    End If
    If 6 <= Flex.Col And Flex.Col <= 8 Then
        cboEdit(Flex.Col).Visible = True
        cboEdit(Flex.Col).SetFocus
    End If
    
End Sub

Private Sub Flex_DblClick()
    
    i = Flex.TextMatrix(Flex.Row, 0)
    
    If i = "" Then Exit Sub
    
    Select Case MsgBox("¿DESEA ELIMINAR EL ARTÍCULO DEL PROVEEDOR SELECCIONADO?" _
                       & vbCrLf & "EL CAMBIO ES IRREVERSIBLE" _
                       , vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
    
    
        Case vbNo: Exit Sub
    
    End Select
    
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
    For i = 2 To 5
        txtEdit(i).Visible = False
    Next i
    For i = 6 To 8
        cboEdit(i).Visible = False
    Next i
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    
    'Carga los Combos
    CargaCombo "rubros", "nombre", "nombre", cboRubro
    CargaCombo "rubros", "nombre", "nombre", cboEdit(6)
    CargaCombo "subrubros", "nombre", "nombre", cboEdit(7)
    
    cboEdit(8).AddItem "$"
    cboEdit(8).AddItem "U$S"
    cboEdit(8).AddItem "€"
    
    cboAModificar.ListIndex = 0
    
    'Carga el Flex
    Cargar
    
End Sub

Sub OrdenaFlex()
    
    Flex.FormatString = "id|Código|Artículo|Precio|% Mostrador|% Cta Cte|Rubro|SubRubro|Moneda|Proveedor||||||"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 1000
    Flex.ColWidth(2) = 4800
    Flex.ColWidth(3) = 1200
    Flex.ColWidth(4) = 1400
    Flex.ColWidth(5) = 1400
    Flex.ColWidth(6) = 1400
    Flex.ColWidth(7) = 1400
    Flex.ColWidth(8) = 1000
    Flex.ColWidth(9) = 3870
    Flex.ColWidth(10) = 0
    Flex.ColWidth(11) = 0
    Flex.ColWidth(12) = 0
    Flex.ColWidth(13) = 0
    Flex.ColWidth(14) = 0
    Flex.ColWidth(15) = 0
    
    Flex.ColAlignment(3) = 2
    
End Sub

Sub Cargar()
    
    If txtDesde.Text = "" And txtHasta.Text = "" And cboRubro.Text = "" Then
        OrdenaFlex
        Exit Sub
    End If
    
    'Vacía la tabla temporal
    BorraTemp
    
    'Copia el resultado a la tabla temporal
    Set rsCargar = New ADODB.Recordset
    SQL = "INSERT INTO temp_proxart_" & nTmp & " (col0, col1, col2, col3, col4, col5, col6, col7, col8, col9) "
    SQL = SQL & "SELECT c.id, a.codigo, a.nombre, c.total, a.porcventa1 '% Mostrador', a.porcventa2 '% Cta Cte', a.rubro, a.subrubro, c.moneda, p.nombre FROM articulosc AS c Inner Join articulos AS a ON c.idart = a.id Inner Join proveedores AS p ON c.idpro = p.id "
    
    'Aplica los filtros
    If txtDesde.Text <> "" And txtHasta.Text <> "" Then
        SQL = SQL & "AND a.codigo BETWEEN '" & txtDesde.Text & "' AND '" & txtHasta & "' "
    End If
    If cboRubro.Text <> "" Then
        SQL = SQL & "AND a.rubro = '" & Trim(cboRubro.Text) & "' "
    End If
    If chkPred.Value = 1 Then
        SQL = SQL & "AND c.predet = 'SÍ' "
    End If
    If Orden = "" Then
        SQL = SQL & "ORDER BY a.codigo;"
    Else
        SQL = SQL & "ORDER BY " & Orden & ";"
    End If
    
    'Muestra el resultado en el flex
    rsCargar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    MuestraTemp
    
End Sub
Sub BorraTemp()
    
    'Vacía la tabla temporal
    Set rsDel = New ADODB.Recordset
    SQL = "DELETE FROM temp_proxart_" & nTmp & ""
    rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
End Sub

Sub MuestraTemp(Optional Orden As String)
    
    'Carga la tabla temporal en el flex
    
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
            'For a = 2 To 10
                Flex.Row = i
                Flex.Col = 1
                Flex.CellBackColor = &HEEEEEE
                Flex.Col = 3
                Flex.CellBackColor = &HEEEEEE
                Flex.Col = 9
                Flex.CellBackColor = &HEEEEEE
            'Next a
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
    
    txtEdit(Index).SelStart = 0
    txtEdit(Index).SelLength = Len(txtEdit(Index).Text)
    
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
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
            If Flex.Col = 4 And Flex.Row > 1 Then
                Flex.Col = 2
            ElseIf Flex.Col = 2 And Flex.Row > 1 Then
                Flex.Row = Flex.Row - 1
                Flex.Col = 8
            Else
                Flex.Col = Flex.Col - 1
            End If
            Flex_Click
        
        Case vbKeyRight
            ActualizarReg Index
            If Flex.Col = 2 And Flex.Row < Flex.Rows - 1 Then
                Flex.Col = 4
            ElseIf Flex.Col = 8 And Flex.Row < Flex.Rows - 1 Then
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
    If Index < 6 Then
        If txtEdit(Index).Text = Flex.TextMatrix(Flex.Row, Flex.Col) Then
            Exit Sub
        End If
    Else
        If cboEdit(Index).Text = Flex.TextMatrix(Flex.Row, Flex.Col) Then
            Exit Sub
        End If
    End If
    
'    'Calcula el total
'    If Index > 3 And Index < 7 Then
'        If txtEdit(4).Text <> "" And txtEdit(5).Text <> "" And txtEdit(6).Text <> "" Then
'            Tot = CalculaDescuentos(txtEdit(4).Text, txtEdit(5).Text)
'            Tot = Format(Tot + ((Tot * CDec(txtEdit(6).Text)) / 100), "0.000")
'        Else
'            Tot = "0,000"
'        End If
'    Else
'        Tot = Flex.TextMatrix(Flex.Row, 7)
'    End If
    
    'Actualiza la tabla de la base
    Set rsUpd = New ADODB.Recordset
    SQL = "UPDATE temp_proxart_" & nTmp & " SET "
    
    Select Case Index
    Case 2: SQL = SQL & "col2 = '" & txtEdit(Index).Text & "' "
    Case 4: SQL = SQL & "col4 = '" & txtEdit(Index).Text & "' "
    Case 5: SQL = SQL & "col5 = '" & txtEdit(Index).Text & "' "
    Case 6: SQL = SQL & "col6 = '" & cboEdit(Index).Text & "' "
    Case 7: SQL = SQL & "col7 = '" & cboEdit(Index).Text & "' "
    Case 8: SQL = SQL & "col8 = '" & cboEdit(Index).Text & "' "
    End Select
    
    SQL = SQL & "WHERE col0 = " & idArtC & ";"
    rsUpd.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Actualiza la información del flex
    If Index < 6 Then
        Flex.TextMatrix(Flex.Row, Flex.Col) = txtEdit(Index).Text
    Else
        Flex.TextMatrix(Flex.Row, Flex.Col) = cboEdit(Index).Text
    End If
    
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
