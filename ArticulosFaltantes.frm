VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ArticulosFaltantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos Faltantes"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
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
   ScaleHeight     =   5415
   ScaleWidth      =   11655
   Begin VB.Frame Frame 
      Caption         =   "Ordenar"
      Height          =   1215
      Left            =   9000
      TabIndex        =   13
      Top             =   480
      Width           =   2535
      Begin VB.OptionButton optOrdenar 
         Caption         =   "Proveedor"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton optOrdenar 
         Caption         =   "Artículo"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.ComboBox cboProveedor 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   735
      Left            =   10200
      Picture         =   "ArticulosFaltantes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " Generar Órdenes"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtArticulo 
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
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Width           =   3960
   End
   Begin VB.TextBox txtCodArticulo 
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
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   1680
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexArt 
      Height          =   2640
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPro 
      Height          =   1440
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2540
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
   Begin VB.TextBox txtCantidad 
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
      Height          =   315
      Left            =   7080
      TabIndex        =   6
      Top             =   480
      Width           =   1680
   End
   Begin VB.Label Label4 
      Caption         =   "Proveedor:"
      Height          =   255
      Left            =   9000
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      X1              =   8880
      X2              =   8880
      Y1              =   0
      Y2              =   5400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Proveedores por artículo:"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   2565
   End
   Begin VB.Label Label1 
      Caption         =   "Agregar:"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Artículo"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Código"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "ArticulosFaltantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cantidadOrdenes As Integer

Private Sub cboProveedor_Click()
    
    If cboProveedor.Text = "Todos" Then
        cmdGenerar.Caption = "GENERAR (" & getCantidadOrdenes & ")"
        cmdGenerar.Enabled = False
    Else
        cmdGenerar.Caption = "GENERAR (1)"
        cmdGenerar.Enabled = True
    End If
    
End Sub

Private Sub cmdGenerar_Click()
    
    Dim i As Integer
    Dim idPro As Integer
    Dim IdOC As Single
    Dim Art As String
    
    Select Case MsgBox("SE VAN A " & cmdGenerar.Caption & " ÓRDENES DE COMPRA           " _
                       & vbCrLf & "¿DESEA CONTINUAR?" _
                       , vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
    
        Case vbNo: Exit Sub
    
    End Select
    
    Dim Imprimir As Boolean
    Imprimir = MsgBox("¿DESEA IMPRIMIR LAS ORDENES DE COMPRA GENERADAS?         ", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title) = vbYes
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM articulosfaltantes "
    If cboProveedor.ListIndex <> -1 And cboProveedor.ListIndex <> 0 Then
        SQL = SQL & "WHERE idpro = " & cboProveedor.ItemData(cboProveedor.ListIndex)
    End If
    SQL = SQL & " ORDER BY idpro"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not Recordset.EOF
        
        If idPro <> Recordset!idPro Then
            
            'Borra los faltantes que ya fueron movidos
            Set rsDel = New ADODB.Recordset
            SQL = "DELETE FROM articulosfaltantes WHERE idpro = " & idPro
            rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
            
            If Imprimir Then imprimirOrdenDeCompra IdOC
            
            idPro = Recordset!idPro
            
            Set rsOC = New ADODB.Recordset
            SQL = "INSERT INTO ocompras (id,idpro,fecha,estado) VALUES(NULL," & idPro & ",'" & Format(Date, "YYYY-MM-DD") & "','PENDIENTE');"
            rsOC.Open SQL, Data, adOpenKeyset, adLockOptimistic
            
            Set rsOC = New ADODB.Recordset
            SQL = "SELECT LAST_INSERT_ID() id;"
            rsOC.Open SQL, Data, adOpenKeyset, adLockOptimistic
            IdOC = rsOC!ID
            rsOC.Close
            
            i = i + 1
            
        End If
        
        Art = getData(Recordset!idArt, "nombre", "articulos")
        
        Set rsOC = New ADODB.Recordset
        SQL = "INSERT INTO ocomprasd (id,idorden,idart,art,cantidad,estado) VALUES(NULL," & IdOC & "," & Recordset!idArt & ",'" & Art & "'," & Recordset!Cantidad & ",'PENDIENTE');"
        rsOC.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        Recordset.MoveNext
    Loop
    
    'Borra los faltantes que ya fueron movidos
    Set rsDel = New ADODB.Recordset
    SQL = "DELETE FROM articulosfaltantes WHERE idpro = " & idPro
    rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Imprimir Then imprimirOrdenDeCompra IdOC
    
    Recordset.Close
    
    Call MsgBox("PROCESO FINALIZADO" _
                & vbCrLf & "ÓRDENES DE COMPRA GENERADAS: " & i & "       " _
                , vbInformation, App.Title)
    
    Cargar
    
End Sub

Private Sub FlexArt_Click()
    
    If FlexArt.TextMatrix(FlexArt.Row, 0) = "" Then Exit Sub
    
    Dim idPro As Single
    idPro = FlexArt.TextMatrix(FlexArt.Row, 5)
    
    Dim idArt As Single
    idArt = FlexArt.TextMatrix(FlexArt.Row, 1)
    
    Dim codArt As String
    codArt = FlexArt.TextMatrix(FlexArt.Row, 2)
    
    CargarProvedoresXArticulo codArt, codArt
    
    '"idart|idpro|Proveedor|Precio|Descuentos|IVA|Total|"
    
    Set rsTmp = New ADODB.Recordset
    SQL = "SELECT col14, col15, col10, col4, col5, col6, col7 FROM temp_proxart_" & nTmp & " ORDER BY CAST(REPLACE(col7,',','.') AS DECIMAL)"
    rsTmp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsTmp.BOF And Not rsTmp.EOF Then
        Set FlexPro.DataSource = rsTmp
    Else
        FlexPro.Clear
        FlexPro.Rows = 2
    End If
    rsTmp.Close
    
'    Set rsPro = New ADODB.Recordset
'    SQL = "SELECT articulosc.idart, articulosc.idpro, proveedores.nombre, articulosc.precio, articulosc.descuento, articulosc.iva, articulosc.total, ''  "
'    SQL = SQL & "FROM articulosc Inner Join proveedores ON articulosc.idpro = proveedores.id WHERE idart = " & idArt & " ORDER BY CAST(REPLACE(total,',','.') AS DECIMAL);"
'    rsPro.Open SQL, Data, adOpenKeyset, adLockOptimistic
'
'    If Not rsPro.BOF And Not rsPro.EOF Then
'        Set FlexPro.DataSource = rsPro
'    Else
'        FlexPro.Clear
'        FlexPro.Rows = 2
'    End If
'    rsPro.Close
    
    OrdenaFlexPro
    
    'Marca el proveedor seleccionado
    For i = 1 To FlexPro.Rows - 1
        If FlexPro.TextMatrix(i, 1) = idPro Then
            FlexPro.TextMatrix(i, 7) = "<-"
        Else
            FlexPro.TextMatrix(i, 7) = ""
        End If
    Next
    
    txtCodArticulo.Text = FlexArt.TextMatrix(FlexArt.Row, 2)
    txtCantidad.Text = FlexArt.TextMatrix(FlexArt.Row, 4)
    txtCantidad.SetFocus
    
End Sub

Private Sub FlexArt_DblClick()
    
    If FlexArt.TextMatrix(FlexArt.Row, 0) = "" Then Exit Sub
    
    Dim idFaltante As Single
    idFaltante = FlexArt.TextMatrix(FlexArt.Row, 0)
    
    Select Case MsgBox("¿Desea eliminar el artículo de la lista de faltantes?    ", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
    
        Case vbNo: Exit Sub
    
    End Select
    
    Set rsDel = New ADODB.Recordset
    SQL = "DELETE FROM articulosfaltantes WHERE id = " & idFaltante
    rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
    FlexPro.Clear
    FlexPro.Rows = 2
    OrdenaFlexPro
    
End Sub

Private Sub FlexArt_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        FlexArt_Click
    End If
    
End Sub

Private Sub FlexPro_Click()
    
    If FlexPro.TextMatrix(FlexPro.Row, 1) = "" Then Exit Sub
    
    Dim idPro As Single
    idPro = FlexPro.TextMatrix(FlexPro.Row, 1)
    
    Dim idArt As Single
    idArt = FlexPro.TextMatrix(FlexPro.Row, 0)
    
    Set Recordset = New ADODB.Recordset
    SQL = "UPDATE articulosfaltantes SET idpro = " & idPro & " WHERE idart = " & idArt
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
    Cargar
    
    'Busca la posicion actual del artículo
    For i = 1 To FlexArt.Rows - 1
        If idArt = FlexArt.TextMatrix(i, 1) Then
            rowArt = i
        End If
    Next
    
    FlexArt.Row = rowArt
    FlexArt_Click
    
End Sub

Private Sub FlexPro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        FlexPro_Click
    End If
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    Cargar
    OrdenaFlexPro
    
End Sub

Sub Cargar()
    
    Set rsCargar = New ADODB.Recordset
    SQL = "SELECT articulosfaltantes.id, articulos.id, articulos.codigo Código, articulos.nombre Artículo, articulosfaltantes.cantidad, articulosfaltantes.idpro, proveedores.nombre Proveedor "
    SQL = SQL & "FROM articulosfaltantes Inner Join articulos ON articulosfaltantes.idart = articulos.id Inner Join proveedores ON articulosfaltantes.idpro = proveedores.id "
    If optOrdenar(0).Value = True Then
        SQL = SQL & "ORDER BY articulos.codigo;"
    Else
        SQL = SQL & "ORDER BY proveedores.nombre, articulos.codigo;"
    End If
    
'    Clipboard.Clear
'    Clipboard.SetText SQL
    
    rsCargar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsCargar.BOF And Not rsCargar.EOF Then
        Set FlexArt.DataSource = rsCargar
    Else
        FlexArt.Clear
        FlexArt.Rows = 2
    End If
    rsCargar.Close
    
    OrdenaFlexArt
    
    'Cuenta la cantidad de ordenes de compra que se generarían
    Set rsCant = New ADODB.Recordset
    SQL = "SELECT DISTINCT af.idpro, p.nombre FROM articulosfaltantes AS af Left Join proveedores AS p ON af.idpro = p.id ORDER BY nombre"
    rsCant.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    cantidadOrdenes = rsCant.RecordCount
    
    cmdGenerar.Caption = "Generar (" & cantidadOrdenes & ")"
    
    cboProveedor.Clear
    cboProveedor.AddItem "Todos"
    cboProveedor.ItemData(cboProveedor.NewIndex) = 0
    
    Do While Not rsCant.EOF
        cboProveedor.AddItem rsCant!nombre
        cboProveedor.ItemData(cboProveedor.NewIndex) = rsCant!idPro
        rsCant.MoveNext
    Loop
    rsCant.Close
    
    cboProveedor.Text = "Todos"
    
End Sub

Function getCantidadOrdenes()

    'Cuenta la cantidad de ordenes de compra que se generarían
    Set rsCant = New ADODB.Recordset
    SQL = "SELECT DISTINCT idpro FROM articulosfaltantes"
    rsCant.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    getCantidadOrdenes = rsCant.RecordCount
    
    rsCant.Close
    
End Function

Sub OrdenaFlexArt()
    
    FlexArt.FormatString = "id|idart|Código|Artículo|Cantidad|idpro|Proveedor"
    FlexArt.ColWidth(0) = 0
    FlexArt.ColWidth(1) = 0
    FlexArt.ColWidth(2) = 1000
    FlexArt.ColWidth(3) = 4200
    FlexArt.ColWidth(4) = 1000
    FlexArt.ColWidth(5) = 0
    FlexArt.ColWidth(6) = 2100
    
End Sub

Sub OrdenaFlexPro()
    
    FlexPro.FormatString = "idart|idpro|Proveedor|Precio|Descuentos|IVA|Total|"
    FlexPro.ColWidth(0) = 0
    FlexPro.ColWidth(1) = 0
    FlexPro.ColWidth(2) = 3300
    FlexPro.ColWidth(3) = 1100
    FlexPro.ColWidth(4) = 1200
    FlexPro.ColWidth(5) = 1000
    FlexPro.ColWidth(6) = 1100
    FlexPro.ColWidth(7) = 600
    
End Sub

Private Sub optOrdenar_Click(Index As Integer)
    
    Cargar
    
End Sub

Private Sub txtCantidad_GotFocus()
    
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Add
    Else
        CambiaPunto txtCantidad, KeyAscii
    End If
    
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtCodArticulo_Change()
    
    'Busca el art si se ingresa el código completo
    If Len(txtCodArticulo.Text) < 6 And txtCodArticulo.Text <> "000000" Then
        txtArticulo.Text = ""
        Exit Sub
    End If
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT id, nombre FROM articulos WHERE codigo = '" & txtCodArticulo.Text & "' AND eliminado <> 1 LIMIT 1"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Recordset.BOF And Recordset.EOF Then
        txtArticulo.Text = ""
        txtCantidad.Text = ""
        Exit Sub
    End If
    
    idArticulo = Recordset!ID
    txtArticulo.Text = Recordset!nombre
    Recordset.Close
    
    'Busca si ya se cargó faltante
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT id, cantidad FROM articulosfaltantes WHERE idart = '" & idArticulo & "';"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        txtCantidad.Text = Recordset!Cantidad
        txtCantidad.SetFocus
    End If
    Recordset.Close
    
End Sub

Private Sub txtCodArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        If txtArticulo.Text = "" And txtCodArticulo.Text <> "000000" And txtCodArticulo.Text <> "0" Then
            ArticulosList.Entrada = "FALTANTES"
            ArticulosList.Height = 5520
            ArticulosList.txtBuscar.Text = txtCodArticulo.Text
            ArticulosList.txtBuscar.SelStart = Len(txtCodArticulo.Text)
            ArticulosList.Show
            ArticulosList.Cargar
            ArticulosList.txtBuscar.SetFocus
        Else
            PasarFoco
            KeyAscii = 0
        End If
        
    End If

End Sub

Sub Add()
    
    'Si se seleccionó un artículo
    Set rsArt = New ADODB.Recordset
    SQL = "SELECT id FROM articulos WHERE codigo = '" & txtCodArticulo.Text & "' AND eliminado = 0"
    rsArt.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If rsArt.BOF And rsArt.EOF Then
        Call MsgBox("Primero debe seleccionar un artículo.      ", vbExclamation, "")
        Exit Sub
    End If
    
    If txtCantidad = "" Then Exit Sub
    
    'Verifica que no haya una orden pendiente con el artículo
    Set rsOCD = New ADODB.Recordset
    SQL = "SELECT * FROM ocomprasd WHERE idart = " & rsArt!ID & " AND estado = 'PENDIENTE'"
    rsOCD.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsOCD.BOF And Not rsOCD.EOF Then
        Call MsgBox("EL ARTÍCULO YA HA SIDO PEDIDO" _
                    & vbCrLf & "LA ORDEN DE COMPRA Nº " & rsOCD!idorden & " ESTÁ PENDIENTE    " _
                    , vbInformation Or vbDefaultButton1, App.Title)
        
        Exit Sub
    End If
    rsOCD.Close
    
    'No permite cargarlo si no tiene ningún proveedor
    If getProveedorRecomendado(rsArt!ID) = 0 Then
        Call MsgBox("EL ARTÍCULO INGRESADO NO TIENE NINGÚN PROVEEDOR ASIGNADO   ", vbExclamation, App.Title)
        Exit Sub
    End If
    
    'Lo agrega a la tabla articulosfaltantes
    Set rsFaltante = New ADODB.Recordset
    SQL = "SELECT * FROM articulosfaltantes WHERE idart = " & rsArt!ID
    rsFaltante.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If rsFaltante.BOF And rsFaltante.EOF Then
        rsFaltante.AddNew
        rsFaltante!idArt = rsArt!ID
        rsFaltante!Cantidad = txtCantidad.Text
        rsFaltante!idPro = getProveedorRecomendado(rsArt!ID) 'le asigna el proveedor recomendado
        rsFaltante.Update
    Else
        If txtCantidad.Text <> "0" Then
            rsFaltante!Cantidad = txtCantidad.Text
            rsFaltante.Update
        Else
            Select Case MsgBox("¿Desea eliminar el artículo de la lista de faltantes?    ", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
                Case vbNo: Exit Sub
            End Select
            rsFaltante.Delete
        End If
    End If
    rsFaltante.Close
    
    rsArt.Update
    rsArt.Close
    
    Cargar
    
    txtCodArticulo.Text = ""
    txtArticulo.Text = ""
    txtCantidad.Text = ""
    txtCodArticulo.SetFocus
    
End Sub
