VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CuentaCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
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
   ScaleHeight     =   5040
   ScaleWidth      =   8775
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
      Left            =   7800
      Picture         =   "CuentaCompra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   " Salir "
      Top             =   4320
      Width           =   855
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
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
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
      Left            =   6840
      Picture         =   "CuentaCompra.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " Exportar a Excel"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtSaldo 
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
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Text            =   "0,00"
      Top             =   4560
      Width           =   1935
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
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "CuentaCompra.frx":0B14
      Left            =   1320
      List            =   "CuentaCompra.frx":0B16
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   4695
   End
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
      Left            =   8040
      Picture         =   "CuentaCompra.frx":0B18
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.ComboBox cboTipo 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "CuentaCompra.frx":10A2
      Left            =   6120
      List            =   "CuentaCompra.frx":10AC
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   1815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5953
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12640511
      ForeColorFixed  =   -2147483640
      GridColorFixed  =   -2147483630
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
   Begin VB.ComboBox cboIdCli 
      Height          =   360
      ItemData        =   "CuentaCompra.frx":10C5
      Left            =   4320
      List            =   "CuentaCompra.frx":10C7
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Saldo Cuenta"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LabelPro 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "CuentaCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub calculaCtaCte()
    
    Dim SaldoCtaCte As Single
    Dim SaldoFacRes As Single
    Dim SaldoFac As Single
    
    'Borra todos los registros de las tablas temporales
    Set dbBorraTemp = New ADODB.Recordset
    dbBorraTemp.Open "DELETE FROM temp", Data, adOpenKeyset, adLockOptimistic
    Set dbBorraTemp2 = New ADODB.Recordset
    dbBorraTemp.Open "DELETE FROM temp2", Data, adOpenKeyset, adLockOptimistic
    
    Flex.Clear
    Flex.Rows = 2
    OrdenaFlex
    
    If cboProveedor.ListIndex = -1 Then
        Exit Sub
    End If
    
    Set rsTemp = New ADODB.Recordset
    SQL = "SELECT * FROM temp2"
    rsTemp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Mueve las facturas del cliente a una tabla temporal
    Set rsFacturas = New ADODB.Recordset
    SQL = "SELECT * FROM compras WHERE idproveedor = '" & cboProveedor.ItemData(cboProveedor.ListIndex) & "' "
    If cboTipo.Text = "RESUMIDO" Then
        SQL = SQL & "AND estado = 'CTACTE' "
    End If
    SQL = SQL & "ORDER BY fecha"
    rsFacturas.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not rsFacturas.EOF
        rsTemp.AddNew
        rsTemp!col0 = rsFacturas!DateTime
        rsTemp!col1 = rsFacturas!Tipo & " " & rsFacturas!letra
        rsTemp!col2 = rsFacturas!Numero
        rsTemp!col3 = Format(rsFacturas!fecha, "dd/mm/yyyy")
        rsTemp!col4 = Format(rsFacturas!Total, "0.00")
        rsTemp!col5 = Format(rsFacturas!Saldo, "0.00")
        rsTemp.Update
        rsFacturas.MoveNext
    Loop
    rsFacturas.Close
    
    If cboTipo.Text <> "RESUMIDO" Then
        'Mueve los pagos del cliente a la tabla temporal
        Set rspagos = New ADODB.Recordset
        SQL = "SELECT * FROM pagosdr WHERE idproveedor = '" & cboProveedor.ItemData(cboProveedor.ListIndex) & "' ORDER BY fecha"
        rspagos.Open SQL, Data, adOpenKeyset, adLockOptimistic
        Do While Not rspagos.EOF
            rsTemp.AddNew
            rsTemp!col0 = rspagos!DateTime
            rsTemp!col1 = "PAGO"
            rsTemp!col2 = Format(rspagos!ID, "00000000")
            rsTemp!col3 = Format(rspagos!fecha, "dd/mm/yyyy")
            rsTemp!col4 = Format(rspagos!Total, "0.00")
            rsTemp.Update
            rspagos.MoveNext
        Loop
        rspagos.Close
        rsTemp.Close
    End If
    
    Set rsImprimir = New ADODB.Recordset
    SQL = "SELECT * FROM temp"
    rsImprimir.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Calcula el saldo de cliente y lo muestra en el flex
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM temp2 ORDER BY col3, col0"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Not Recordset.BOF And Not Recordset.EOF Then
        Do While Not Recordset.EOF
            rsImprimir.AddNew
            Flex.TextMatrix(Flex.Rows - 1, 1) = Recordset!col1          'Tipo
            Flex.TextMatrix(Flex.Rows - 1, 2) = Recordset!col2          'Número
            Flex.TextMatrix(Flex.Rows - 1, 0) = Recordset!col3          'Fecha
            rsImprimir!col2 = Recordset!col1
            rsImprimir!col3 = Recordset!col2
            rsImprimir!col1 = Recordset!col3
            
            If Left(Recordset!col1, 7) = "FACTURA" Or Left(Recordset!col1, 11) = "NOTA DÉBITO" Then
                Flex.TextMatrix(Flex.Rows - 1, 3) = Recordset!col4      'Debe
                Flex.TextMatrix(Flex.Rows - 1, 4) = " "
                rsImprimir!col4 = Recordset!col4
                SaldoCtaCte = SaldoCtaCte + CDec(Recordset!col4)
                If cboTipo.Text = "RESUMIDO" Then
                    SaldoCtaCte = SaldoCtaCte + CDec(Recordset!col5)
                End If
            ElseIf Recordset!col1 = "PAGO" Or Left(Recordset!col1, 12) = "NOTA CRÉDITO" Then
                Flex.TextMatrix(Flex.Rows - 1, 3) = " "
                Flex.TextMatrix(Flex.Rows - 1, 4) = Recordset!col4      'Haber
                rsImprimir!col5 = Recordset!col4
                SaldoCtaCte = SaldoCtaCte - CDec(Recordset!col4)
            End If
            
            If Recordset!col1 <> "PAGO" Then
                Flex.TextMatrix(Flex.Rows - 1, 5) = Recordset!col5       'Saldo Fac.
                rsImprimir!col6 = Recordset!col5
                SaldoFacRes = SaldoFacRes + CDec(Recordset!col5)
            End If
            
            If cboTipo.Text = "ANALÍTICO" Then
                Flex.TextMatrix(Flex.Rows - 1, 6) = Format(SaldoCtaCte, "0.00")   'Saldo Cta.
                rsImprimir!col7 = Format(SaldoCtaCte, "0.00")
                txtSaldo.Text = Format(SaldoCtaCte, "0.00")
            Else
                Flex.TextMatrix(Flex.Rows - 1, 6) = Format(SaldoFacRes, "0.00")   'Saldo Cta.
                rsImprimir!col7 = Format(SaldoFacRes, "0.00")
                txtSaldo.Text = Format(SaldoFacRes, "0.00")
            End If
            
            rsImprimir.Update
            Recordset.MoveNext
            Flex.Rows = Flex.Rows + 1
            
        Loop
        Flex.Rows = Flex.Rows - 1
    Else
        Flex.Rows = 2
    End If
    rsImprimir.Close
    Recordset.Close
    
End Sub

Public Sub OrdenaFlex()
    
    Flex.FormatString = "Fecha|Tipo|Número|Debe|Haber|Saldo Fac.|Saldo Cta."
    Flex.ColWidth(0) = 1250
    Flex.ColWidth(1) = 1450
    Flex.ColWidth(2) = 1100
    Flex.ColWidth(3) = 1100
    Flex.ColWidth(4) = 1100
    Flex.ColWidth(5) = 1100
    Flex.ColWidth(6) = 1100
    
End Sub

Private Sub cboProveedor_Click()

    If cboProveedor.ListIndex <> -1 Then
        txtCodProveedor.Text = getData(cboProveedor.ItemData(cboProveedor.ListIndex), "codigo", "proveedores")
    End If
    
End Sub

Private Sub cboProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        cmdOk_Click
    End If
    
End Sub

Private Sub cmdExcel_Click()

Exportar_Excel "C:\Cuenta Corriente.xls", Flex

'Dim ApExcel As Variant
'Dim nFila As Integer
'Dim Path As String
'Dim Desc
'
'Path = "C:\Cuenta Corriente.xls"
'
''Elimina el listado antiguo (si existe).
'If Len(Dir(Path)) > 0 Then
'    Kill Path
'End If
'
'Set ApExcel = CreateObject("Excel.application")
'
''ApExcel.Workbooks.Open Path     'Abre un Libro
'ApExcel.Workbooks.Add            'Crea un Libro
'
''Define los margenes de impresión
'ApExcel.Worksheets(1).PageSetup.LeftMargin = 13
'ApExcel.Worksheets(1).PageSetup.RightMargin = 13
'ApExcel.Worksheets(1).PageSetup.TopMargin = 28
'ApExcel.Worksheets(1).PageSetup.BottomMargin = 28
'
''Formato al encabezado
'ApExcel.Range("A1:D1").Merge           'Combina las celdas
'ApExcel.Cells(1, 1).Font.Size = 16     'Tamaño de Fuente
'ApExcel.Range("A1").HorizontalAlignment = xlCenter 'Alínea el texto
'ApExcel.Range("D:D").HorizontalAlignment = xlRight 'Alínea el texto
'ApExcel.Cells(1, 1).Value = "FRIGORÍFICO MERLO"
'
''escribe los titulos (fila, columna)
'ApExcel.Range("3:3").HorizontalAlignment = xlCenter 'Alínea el texto
'ApExcel.Range("3:3").Font.Bold = True 'Titulos en negrita
'ApExcel.Cells(3, 1).Value = "Nº Original"
'ApExcel.Cells(3, 2).Value = "Descripción"
'ApExcel.Cells(3, 3).Value = "Nº Art"
'ApExcel.Cells(3, 4).Value = "Precio"
'
'nFila = 4
'
''Barra de progreso
'Set rsArt = New ADODB.Recordset
'SQL = "SELECT * FROM articulos WHERE activo = 'True'"
'rsArt.Open SQL, Data, adOpenKeyset, adLockOptimistic
'    zMain.pBar.Max = rsArt.RecordCount
'rsArt.Close
'zMain.pBar.Value = 0
'zMain.sBar.Height = 0
'zMain.pBar.Height = 255
'
''Selecciona todos los grupos
'Set rsGrupos = New ADODB.Recordset
'SQL = "SELECT * FROM grupos ORDER BY id"
'rsGrupos.Open SQL, Data, adOpenKeyset, adLockOptimistic
'Do While Not rsGrupos.EOF
'
'    'Escribe el nombre del Grupo
'    ApExcel.Range("A" & nFila & ":D" & nFila).Merge
'    ApExcel.Range("A" & nFila).HorizontalAlignment = xlCenter
'    ApExcel.Cells(nFila, 1).Interior.Color = RGB(150, 200, 255)
'    ApExcel.Cells(nFila, 1).Font.Size = 13
'    ApExcel.Cells(nFila, 1).Value = rsGrupos!grupo
'    nFila = nFila + 1
'
'    'Seleccciona los artículos de ese grupo
'    Set Recordset = New ADODB.Recordset
'    SQL = "SELECT * FROM articulos WHERE grupo = '" & rsGrupos!grupo & "' AND activo = 'True'"
'    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
'    Do While Not Recordset.EOF
'        'Escribe el contenido
'        ApExcel.Cells(nFila, 1).Value = Recordset!idor
'        ApExcel.Cells(nFila, 2).Value = Recordset!descripcion
'        ApExcel.Cells(nFila, 3).Value = Recordset!ID
'        If Monotributo = True Then
'            Ini = Recordset!precioventa
'            conDesc = Ini - (Ini * (Desc / 100))
'            ApExcel.Cells(nFila, 4).Value = Format(conDesc * CDec("1,21"), "0.00")
'        Else
'            ApExcel.Cells(nFila, 4).Value = FormatCurrency(Recordset!precioventa)
'        End If
'        zMain.pBar.Value = zMain.pBar.Value + 1
'        nFila = nFila + 1
'        Recordset.MoveNext
'    Loop
'    Recordset.Close
'    rsGrupos.MoveNext
'
'Loop
'rsGrupos.Close
'
''Oculta la barra
'zMain.pBar.Value = 0
'zMain.pBar.Height = 0
'zMain.sBar.Height = 255
'
''Autoajusta el ancho de la columna descripción
'ApExcel.Range("B:B").EntireColumn.AutoFit
'
''Muestra los bordes de las celdas
'ApExcel.Range("A3:D" & (nFila - 1)).Borders.Color = RGB(50, 50, 50)
'
''Guardar libro
'ApExcel.ActiveWorkbook.SaveAs Path
'
''Muestra el Archivo
'ApExcel.Visible = True
'
'Set ApExcel = Nothing

'    If cboIdCli.Text = "" Then
'        Call MsgBox("Debe seleccionar un Cliente.", vbExclamation, "Atención")
'        cboProveedor.SetFocus
'        Exit Sub
'    End If
'
'    calculaCtaCte
'
'    drCtaCte1.Sections("ReportHeader").Controls("rptCliente").Caption = cboProveedor.Text
'
'    ActualizarDR
'    drCtaCte1.Show
    
End Sub

Private Sub cmdExport_Click()

End Sub

Public Sub cmdOk_Click()
    
    calculaCtaCte
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    'Posiciona el form
    Me.Top = (zMain.ScaleHeight / 2) - (Me.Height / 2)
    Me.Left = (zMain.ScaleWidth / 2) - (Me.Width / 2)
    
    'Carga los proveedores en los combos
    CargaCombo "proveedores", "nombre", "nombre", cboProveedor
    
    cboTipo.ListIndex = 0
    
    calculaCtaCte
    
End Sub

Private Sub txtCodProveedor_Change()
    
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
        cmdOk_Click
    End If

End Sub
