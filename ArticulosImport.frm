VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ArticulosImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos - Importar"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
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
   ScaleHeight     =   4695
   ScaleWidth      =   9615
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
      Left            =   8640
      Picture         =   "ArticulosImport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Salir "
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtTotal 
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
      Enabled         =   0   'False
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1560
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
      Left            =   7680
      Picture         =   "ArticulosImport.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Guardar"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdCargar 
      Height          =   360
      Left            =   7800
      Picture         =   "ArticulosImport.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   360
   End
   Begin VB.ComboBox cboProveedor 
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "ArticulosImport.frx":109E
      Left            =   1680
      List            =   "ArticulosImport.frx":10A0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   6015
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
      TabIndex        =   1
      Top             =   480
      Width           =   1560
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5106
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
   Begin VB.Label Label1 
      Caption         =   "Total de Artículos"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "ArticulosImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub LeerExcel()
    
    Dim Conexion As ADODB.Connection
    Dim RsExcel As ADODB.Recordset
    
    Hoja = "Hoja1$"
    
    'Vacía y abre la tabla temporal
    Set Recordset = New ADODB.Recordset
    SQL = "DELETE FROM temp_art_import"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM temp_art_import"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Muestra la barra
    zMain.pBar.Value = 0
    zMain.sBar.Height = 0
    zMain.pBar.Height = 255
    
    'Abre el Excel
    Set Conexion = New ADODB.Connection
    Conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & App.Path & "\PreciosCompra.xls" & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
    
    Set RsExcel = New ADODB.Recordset
    With RsExcel
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
    SQL = "SELECT * FROM [" & Hoja & "]"
    RsExcel.Open SQL, Conexion, , , adCmdText
    zMain.pBar.Max = RsExcel.RecordCount
    
    'Lo copia a la tabla temporal
    Do While Not RsExcel.EOF
        If RsExcel!codigo <> "" Then
            Recordset.AddNew
            Recordset!codigo = RsExcel!codigo
            Recordset!Precio = Format(RsExcel!Precio, "0.000")
            Recordset.Update
        End If
        zMain.pBar.Value = zMain.pBar.Value + 1
        RsExcel.MoveNext
    Loop
    RsExcel.Close
    
    'Oculta la barra
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    zMain.sBar.Height = 255
    
End Sub

Private Function ExportarASQL() As Boolean
    
    'Exporta la hoja de excel de los criterios
    'a una tabla de SQL (us_InvTablaExcel)
    Dim cnn As ADODB.Connection
    Dim lNumRegAfect As Long
    Dim strSQL As String
    
    'Eliminar la tabla
    Set rsDrop = New ADODB.Recordset
    SQL = "DELETE FROM us_InvTablaExcel"
    rsDrop.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    ' Abrimos una conexión con el libro de trabajo
    Set cnn = New ADODB.Connection
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & App.Path & "\PreciosCompra.xls"
        .Properties("Extended Properties") = "Excel 8.0;HDR=Yes;"
        .Open
    End With
    
    ' Importamos utilizando una cadena ODBC
    strSQL = "SELECT * INTO [ODBC;Driver={MySQL ODBC 3.51 Driver};" _
    & "SERVER=naty;" _
    & "PORT=3306;" _
    & "DATABASE=bulonera;" _
    & "UID=root;" _
    & "PWD=cmsis00;" & "].us_InvTablaExcel " & _
    "FROM [Hoja1$]"
    
    ' Ejecutamos la consulta
    cnn.Execute strSQL, lNumRegAfect, adExecuteNoRecords
    
    MsgBox "Número de registros afectados: " & lNumRegAfect
    
    'Cerramos la conexión
    cnn.Close
    
    ExportarASQL = True

End Function

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

Private Sub cmdCargar_Click()
    
    If cboProveedor.ListIndex = -1 Then
        cboProveedor.SetFocus
        Exit Sub
    End If
    
    LeerExcel
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT articulosc.codigo, articulos.nombre, articulosc.precio, temp_art_import.precio FROM articulos "
    SQL = SQL & "Inner Join articulosc ON articulos.id = articulosc.idart "
    SQL = SQL & "Inner Join temp_art_import ON articulosc.codigo = temp_art_import.codigo "
    SQL = SQL & "WHERE articulosc.idpro = " & cboProveedor.ItemData(cboProveedor.ListIndex) & " "
    SQL = SQL & "ORDER BY articulosc.codigo"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        Set Flex.DataSource = Recordset
        txtTotal.Text = Recordset.RecordCount
    Else
        Flex.Clear
        Flex.Rows = 2
        txtTotal.Text = ""
        Call MsgBox("NO SE HAN ENCONTRADO ARTÍCULOS QUE COINCIDAN CON LOS DEL PROVEEDOR SELECCIONADO    ", vbInformation, App.Title)
    End If
    Recordset.Close
    
    OrdenaFlex
    
End Sub

Private Sub cmdGuardar_Click()
    
    If cboProveedor.ListIndex = -1 Then
        cboProveedor.SetFocus
        Exit Sub
    End If
    
    Dim Total As Double
    Dim IVA As Double
    
    'Muestra la barra
    zMain.pBar.Value = 0
    zMain.sBar.Height = 0
    zMain.pBar.Height = 255
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM temp_art_import"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    zMain.pBar.Max = Recordset.RecordCount
    Do While Not Recordset.EOF
        
        Set rsArtC = New ADODB.Recordset
        SQL = "SELECT * FROM articulosc WHERE codigo = '" & Recordset!codigo & "' AND idpro = " & cboProveedor.ItemData(cboProveedor.ListIndex)
        rsArtC.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If rsArtC.RecordCount = 1 Then
            
            IVA = rsArtC!IVA
            
            'Calcula el total
            Total = CalculaDescuentos(Recordset!Precio, rsArtC!descuento)
            Total = Format(Total + ((Total * IVA) / 100), "0.000")
            
            rsArtC!Precio = Recordset!Precio
            rsArtC!Total = Format(Total, "0.000")
            
            rsArtC.Update
        ElseIf rsArtC.RecordCount <> 0 Then
            a = rsArtC.RecordCount
            For i = 1 To rsArtC.RecordCount
                Select Case MsgBox("¿DESEA ACTUALIZAR EL PRECIO DE '" & Left(getData(rsArtC!idArt, "nombre", "articulos"), 25) & "' " _
                                   & vbCrLf & "DE   '$" & rsArtC!Precio & "'   A   '$" & Recordset!Precio & "' ?" _
                                   , vbYesNo Or vbQuestion Or vbDefaultButton2, "CÓDIGO DUPLICADO")
                    Case vbYes
                        
                        
                        IVA = rsArtC!IVA
                        
                        'Calcula el total
                        Total = CalculaDescuentos(Recordset!Precio, rsArtC!descuento)
                        Total = Format(Total + ((Total * IVA) / 100), "0.000")
                        
                        rsArtC!Precio = Recordset!Precio
                        rsArtC!Total = Format(Total, "0.000")
                        rsArtC.Update
                End Select
                rsArtC.MoveNext
            Next i
        End If
        rsArtC.Close
        
        zMain.pBar.Value = zMain.pBar.Value + 1
        Recordset.MoveNext
    Loop
    Recordset.Close
    
    'Oculta la barra
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    zMain.sBar.Height = 255
    
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    CargaCombo "proveedores", "nombre", "nombre", cboProveedor
    
    OrdenaFlex
    
    'ExportarASQL
    
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

Sub OrdenaFlex()
    
    Flex.FormatString = "Código|Nombre|Precio Actual|Precio Nuevo"
    Flex.ColWidth(0) = 1500
    Flex.ColWidth(1) = 4550
    Flex.ColWidth(2) = 1500
    Flex.ColWidth(3) = 1500
    
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub
