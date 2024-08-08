VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form InformeZeta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe Excel"
   ClientHeight    =   1560
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
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
   ScaleHeight     =   1560
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
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
      Left            =   3480
      Picture         =   "informeZeta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Salir "
      Top             =   840
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   129368065
      CurrentDate     =   43717
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   129368065
      CurrentDate     =   43717
   End
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
      Left            =   2520
      Picture         =   "informeZeta.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Exportar "
      Top             =   840
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   10095
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   17806
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
   Begin VB.Label lblHasta 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblFechaDesde 
      Caption         =   "Desde"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "InformeZeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()

    Flex.Clear

    Set rsMenu = New ADODB.Recordset
    SQL = "SELECT id, codigo, nombre, precio FROM menus WHERE eliminado = 0 ORDER BY nombre"
    rsMenu.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsMenu.BOF And Not rsMenu.EOF Then
        Set Flex.DataSource = rsMenu
    Else
        Flex.Rows = 2
        Flex.Cols = 6
    End If
    
    Flex.Cols = 6
    ordenaFlex
    
    Zeta True
    
    Zeta False

End Sub

Private Sub Excel(nombreHoja As String)
    
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
    Dim Encabezado() As String
    
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    
    Encabezado = Split(Flex.FormatString, "|")
    
    For i = LBound(Encabezado) To UBound(Encabezado)
        o_Hoja.Cells(1, i + 1).Value = Encabezado(i)
    Next i
    
    o_Excel.Range("1:1").Font.Bold = True 'Encabezado en negrita
    
    ' -- Bucle para Exportar los datos
    With Flex
        For Fila = 2 To .Rows
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix(Fila - 1, Columna)
            Next
        Next
    End With
    
    o_Excel.Range("A:A").EntireColumn.AutoFit
    o_Excel.Range("B:B").EntireColumn.AutoFit
    o_Excel.Range("C:C").EntireColumn.AutoFit
    o_Excel.Range("D:D").EntireColumn.AutoFit
    o_Excel.Range("E:E").EntireColumn.AutoFit
    o_Excel.Range("F:F").EntireColumn.AutoFit
    
    o_Excel.Worksheets("Hoja1").Delete
    o_Excel.Worksheets("Hoja2").Delete
    o_Excel.Worksheets("Hoja3").Delete
    
    o_Hoja.Name = nombreHoja
    
    'o_Libro.Close True, sOutputPath
    o_Excel.Visible = True
    
    ' -- Cerrar Excel
    'o_Excel.Quit
    
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    'Exportar_Excel = True
    
End Sub

Private Sub CargoDetalle(idcomida As Double)

    Set rsDetalle = New ADODB.Recordset
    SQL = "SELECT * FROM comidasd WHERE idcomida = " & idcomida
    rsDetalle.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not rsDetalle.EOF
        For Fila = 1 To Flex.Rows - 1
            If Flex.TextMatrix(Fila, 0) = rsDetalle!idArt Then
                Flex.TextMatrix(Fila, 4) = CDbl(Flex.TextMatrix(Fila, 4)) + rsDetalle!Cantidad
                Flex.TextMatrix(Fila, 5) = Format(CInt(Flex.TextMatrix(Fila, 4)) * CDbl(Flex.TextMatrix(Fila, 3)), "$ 0.00")
            End If
        Next
        rsDetalle.MoveNext
    Loop
    rsDetalle.Close
    
End Sub

Private Sub Zeta(blanco As Boolean)

    Dim Cantidad As Double
    Dim total As Double
    
    'Cargo cantidades y totales 0 en el flex
    For Fila = 1 To Flex.Rows - 1
        Flex.TextMatrix(Fila, 4) = 0
        Flex.TextMatrix(Fila, 5) = Format(0, "$ 0.00")
    Next
    
    Cantidad = 0
    total = 0
    
    Set rsZeta = New ADODB.Recordset
    If blanco Then
        SQL = "SELECT id FROM comidas WHERE idempleado <> 9999 "
        If dtpDesde.Value = dtpHasta.Value Then
            SQL = SQL & "AND date(fecha) = '" & Format(dtpDesde.Value, "yyyy-mm-dd") & "' "
        Else
            SQL = SQL & "AND date(fecha) BETWEEN '" & Format(dtpDesde.Value, "yyyy-mm-dd") & "' AND '" & Format(dtpHasta.Value, "yyyy-mm-dd") & "' "
        End If
    Else
        SQL = "SELECT id FROM comidas WHERE idempleado = 9999 "
        If dtpDesde.Value = dtpHasta.Value Then
            SQL = SQL & "AND date(fecha) = '" & Format(dtpDesde.Value, "yyyy-mm-dd") & "' "
        Else
            SQL = SQL & "AND date(fecha) BETWEEN '" & Format(dtpDesde.Value, "yyyy-mm-dd") & "' AND '" & Format(dtpHasta.Value, "yyyy-mm-dd") & "' "
        End If
    End If
    rsZeta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Not rsZeta.EOF And Not rsZeta.BOF Then
        Do While Not rsZeta.EOF
            Cantidad = Cantidad + 1
            CargoDetalle rsZeta!id
            rsZeta.MoveNext
        Loop
    Else
        Call MsgBox("No existen tickets en el período seleccionado", vbExclamation, "No hay datos")
        Exit Sub
    End If
    
    For Fila = 1 To Flex.Rows - 1
        total = total + CDbl(Flex.TextMatrix(Fila, 5))
    Next
    
    rsZeta.Close
    
    If blanco Then
        Excel ("blanco")
    Else
        Excel ("Negro")
    End If
        

End Sub

Sub ordenaFlex()

    Flex.FormatString = "id|Codigo|Nombre|Precio|Cantidad|Total"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 1100
    Flex.ColWidth(2) = 2000
    Flex.ColWidth(3) = 1000
    Flex.ColWidth(4) = 1000
    Flex.ColWidth(5) = 1000

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.Width = 4590
    Me.Height = 2955
    initForm Me
    If Day(Now) < 16 Then
        dtpDesde.Value = DateSerial(Year(Now), Month(Now), 1)
        dtpHasta.Value = DateSerial(Year(Now), Month(Now), 15)
    Else
        dtpDesde.Value = DateSerial(Year(Now), Month(Now), 16)
        dtpHasta.Value = DateSerial(Year(Now), Month(Now) + 1, 0)
    End If
    
    
End Sub
