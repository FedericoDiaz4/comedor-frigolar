Attribute VB_Name = "Module1"
Global Recordset As ADODB.Recordset
Global Data As ADODB.Connection
Global SQL As String
Global nTmp As String
Global idcomida As Double
Public idEmpleado As Integer
Dim inicio As Boolean
Private Declare Sub InitCommonControls Lib "comctl32" ()

'Option Explicit

' declaraciones del api
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    wParam As Any, lParam As Any) As Long
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Constantes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_VSCROLL As Integer = &H115

Dim PrevProc As Long


Sub Main()
    
    'Conexión a la base de datos
    Call ConectarDB
    
    'Número de tabla temporal
    nTmp = Format(Aleatorio(0, 9999), "0000")
    
'    'Crea la tabla temporal
'    Set rsTmp = New ADODB.Recordset
'    SQL = "CREATE TABLE `temp_proxart_" & nTmp & "` (  `col0` varchar(255) default '',  `col1` varchar(255) default NULL,  `col2` varchar(255) default NULL,  `col3` varchar(255) default NULL,  `col4` varchar(255) default NULL,  `col5` varchar(255) default NULL,  `col6` varchar(255) default NULL,  `col7` varchar(255) default NULL,  `col8` varchar(255) default NULL,  `col9` varchar(255) default NULL,  `col10` varchar(255) default NULL,  `col11` varchar(255) default NULL,  `col12` varchar(255) default NULL,  `col13` varchar(255) default NULL,  `col14` varchar(255) default NULL,  `col15` varchar(255) default NULL) ENGINE=InnoDB DEFAULT CHARSET=utf8"
'    rsTmp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Muestra el formulario principal
    zMain.Show
    
End Sub

Sub ConectarDB()
    
    Set Data = New ADODB.Connection
    Data.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
    & "SERVER=localhost;" _
    & "PORT=3306;" _
    & "UID=root;" _
    & "PWD=cmsis00;" _
    & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    Data.CursorLocation = adUseClient
    Data.Open
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = 'ventas';"
    Recordset.Open SQL, Data
'    If Recordset.BOF And Recordset.EOF Then
'
'        Set rsCrear = New ADODB.Recordset
'        SQL = "CREATE DATABASE 'ventas'; SELECT DATABASE 'ventas';"
'        rsCrear.Open SQL, Data
'
'    End If
    Recordset.Close
    Data.Close
    
    'Conexión a la base de datos
    Set Data = New ADODB.Connection
    Data.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
    & "SERVER=localhost;" _
    & "PORT=3306;" _
    & "DATABASE=comedorfrigolar;" _
    & "UID=root;" _
    & "PWD=cmsis00;" _
    & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    Data.CursorLocation = adUseClient
    Data.Open
    
End Sub

'---------------------------------------------------------------------------------------
' \\ -- Verifica si la conexión está activa. Si no lo está, la reconecta.
'---------------------------------------------------------------------------------------

Sub VerificarConexion()
    
    If Data.State = 0 Then
        ConectarDB
    End If
    
End Sub

Function Centrar(texto As String, Cantidad As Integer) As String
    
    Dim espacios As Integer
    espacios = (Cantidad - Len(texto)) / 2
    
    For i = 0 To espacios
        texto = " " & texto
    Next
    
    Centrar = texto
    
End Function

Function Derecha(texto As String, Cantidad As Integer) As String

    Dim espacios As Integer
    espacios = (Cantidad - Len(texto))
    
    For i = 0 To espacios
        texto = " " & texto
    Next

    Derecha = texto

End Function

Function Detalle(Producto As String, Cantidad As Integer, total As String)

    Dim texto As String
    Dim espacios As Integer
    
    texto = Producto
    espacios = 17 - (Len(Producto))
    
    For i = 0 To espacios
        texto = texto & " "
    Next
    
    espacios = 4 - (Len(Cantidad))
    
    For i = 0 To espacios
        texto = texto & " "
    Next
    texto = texto & Cantidad
    
    espacios = 9 - (Len(total))
    
    For i = 0 To espacios
        texto = texto & " "
    Next
    
    texto = texto & total

    Detalle = texto

End Function

Function DetalleZ(Producto As String, Cantidad As Integer)

    Dim texto As String
    Dim espacios As Integer
    
    texto = Producto
    espacios = 20 - (Len(Producto))
    
    For i = 0 To espacios
        texto = texto & " "
    Next

    espacios = 4 - (Len(Cantidad))
    
    For i = 0 To espacios
        texto = texto & " "
    Next
    
    texto = texto & Cantidad

    DetalleZ = texto
    
End Function

Sub Format000()
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT id, precio, total FROM articulosc"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not Recordset.EOF
        If Recordset!Precio <> Format(Recordset!Precio, "0.000") Or Recordset!total <> Format(Recordset!total, "0.000") Then
            Recordset!Precio = Format(Recordset!Precio, "0.000")
            Recordset!total = Format(Recordset!total, "0.000")
            Recordset.Update
        End If
        Recordset.MoveNext
    Loop
    Recordset.Close
    
    MsgBox "listo"
    
End Sub

Public Function Rellena(ByVal sCadena As String, iLen As Integer, Optional DerIzq) As String
    
    'Si no se especifica DerIzq, se justifica a la Izquierda
    'Si se especifica, un valor 0 lo justifica a la Izquierda, otro cualquiera a la Derecha
    'Izquierda 0 (vbLeftJustify), Derecha 1 (vbRightJustify), Centrado 2 (vbCenter)
    
    Dim sTmp As String
    Dim vDerIzq As Integer
    
    sTmp = Space$(iLen)
    If IsMissing(DerIzq) Then
        vDerIzq = vbLeftJustify
    Else
        vDerIzq = CInt(DerIzq)
    End If
    sCadena = Trim$(sCadena)
    
    'Si no se va a mostrar todo
    If (Len(sCadena) > iLen) Then
        sCadena = Left$(sCadena, iLen)
    End If
    
    If vDerIzq = vbRightJustify Then
        RSet sTmp = sCadena
    ElseIf vDerIzq = vbLeftJustify Then
        LSet sTmp = sCadena
    Else    'Centrado
        sTmp = sCadena
        Do While Len(sTmp) < iLen
            sTmp = " " & sTmp & " "
        Loop
    sTmp = Left$(sTmp, iLen)
    End If
    
    Rellena = sTmp
    
End Function

Public Function Exportar_Excel1(sOutputPath As String, FlexGrid As Object, Optional Formateado As Boolean) As Boolean
    
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
    
    Encabezado = Split(FlexGrid.FormatString, "|")
    
    For i = 1 To UBound(Encabezado)
        o_Hoja.Cells(1, i).Value = Encabezado(i)
    Next i
    
    o_Excel.Range("1:1").Font.Bold = True 'Encabezado en negrita
    
    ' -- Bucle para Exportar los datos
    With FlexGrid
        zMain.pBar.Max = .Rows - 1
        For Fila = 2 To .Rows
            For Columna = 1 To .Cols - 1
                o_Hoja.Cells(Fila, Columna).Value = .TextMatrix(Fila - 1, Columna)
            Next
            zMain.pBar.Value = zMain.pBar.Value + 1
        Next
    End With
    
    
    
    If Formateado Then
        o_Excel.Range("G:G").Format = moneda
    End If
    
    o_Excel.Range("A:A").EntireColumn.AutoFit
    o_Excel.Range("B:B").EntireColumn.AutoFit
    o_Excel.Range("C:C").EntireColumn.AutoFit
    o_Excel.Range("D:D").EntireColumn.AutoFit
    o_Excel.Range("E:E").EntireColumn.AutoFit
    o_Excel.Range("F:F").EntireColumn.AutoFit
    o_Excel.Range("G:G").EntireColumn.AutoFit
    o_Excel.Range("H:H").EntireColumn.AutoFit
    o_Excel.Range("I:I").EntireColumn.AutoFit
    
    
    'Oculta la barra
    
    'o_Libro.Close True, sOutputPath
    o_Excel.Visible = True
    
    ' -- Cerrar Excel
    'o_Excel.Quit
    
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel1 = True
    
End Function


' -------------------------------------------------------------------------------------------
' \\ -- Función para crear un nuevo libro con el contenido del Grid
' -------------------------------------------------------------------------------------------
Public Function Exportar_Excel(sOutputPath As String, Optional desde As Date, Optional hasta As Date) As Boolean
    
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim nombre      As String
    Dim Cantidad    As Integer
    Dim Totalpersona As Double
    
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    
    Set o_Hoja = o_Libro.Worksheets.Add
    
    Set rsconsulta = New ADODB.Recordset
    SQL = "SELECT c.id, p.nombre empleado, p.nrolegajo nrodoc, c.fecha, c.precio precio, p.idtipo FROM comidas as c "
    SQL = SQL & "INNER JOIN empleados as p ON p.id = c.idempleado "
    SQL = SQL & "AND date(c.fecha) BETWEEN '" & Format(desde, "yyyy-MM-dd") & "' AND '" & Format(hasta, "yyyy-MM-dd") & "'"
    SQL = SQL & "ORDER BY p.nrolegajo, c.fecha"
    rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    o_Hoja.Name = Left(Replace(Replace("Detalle de Comidas", "*", ""), "/", " "), 30)
    
    total = rsconsulta.RecordCount
    
    Fila = 3
    
    zMain.pBar.Value = 0
    zMain.pBar.Height = 255
    If (total <> 0) Then
        zMain.pBar.Max = total
    End If
    
    ' Cargo Encabezado
    o_Hoja.Cells(Fila, 1).Value = "Empresa"
    o_Hoja.Cells(Fila, 2).Value = "Apellido, Nombre"
    o_Hoja.Cells(Fila, 3).Value = "Nro Legajo"
    o_Hoja.Cells(Fila, 4).Value = "Puesto"
    o_Hoja.Cells(Fila, 5).Value = "Retiró"
    o_Hoja.Cells(Fila, 6).Value = "Dia"
    o_Hoja.Cells(Fila, 7).Value = "Cantidad"
    o_Hoja.Cells(Fila, 8).Value = "Importe"
    
    Cantidad = 0
    Fila = Fila + 1
    ' Encabezado en negrita
    o_Excel.Range("1:1").Font.Bold = True
    o_Excel.Range("A1:B1").interior.Color = RGB(166, 166, 166)
    o_Excel.Range("3:3").Font.Bold = True
        
    ' Cargo Contenido
    Do While Not rsconsulta.EOF
    
        o_Hoja.Cells(Fila, 2).Value = rsconsulta!empleado
        o_Hoja.Cells(Fila, 3).Value = rsconsulta!nrodoc
        o_Hoja.Cells(Fila, 6).Value = rsconsulta!fecha
        o_Hoja.Cells(Fila, 6).numberformat = "dd/mm/yyyy hh:mm:ss"
        o_Hoja.Cells(Fila, 7).Value = "1"
        o_Hoja.Cells(Fila, 8).Value = CDbl(rsconsulta!Precio)
        
        Cantidad = Cantidad + 1
        Fila = Fila + 1
        zMain.pBar.Value = zMain.pBar.Value + 1
        
        rsconsulta.MoveNext
        
    Loop

    rsconsulta.Close
    
    'Autoajusta el ancho de las columnas corte y depósito
    o_Excel.Range("A:A").EntireColumn.AutoFit
    o_Excel.Range("B:B").EntireColumn.AutoFit
    o_Excel.Range("C:C").EntireColumn.AutoFit
    o_Excel.Range("D:D").EntireColumn.AutoFit
    o_Excel.Range("E:E").EntireColumn.AutoFit
    o_Excel.Range("F:F").EntireColumn.AutoFit
    o_Excel.Columns("G").ColumnWidth = 10.67

    o_Hoja.Cells(Fila, 7).Value = Cantidad

    o_Hoja.Cells(1, 1).Value = "Total"
    o_Hoja.Cells(1, 2).Value = Cantidad
    
    o_Excel.Worksheets("Hoja1").Delete
    o_Excel.Worksheets("Hoja2").Delete
    o_Excel.Worksheets("Hoja3").Delete
    'o_Libro.Close True, sOutputPath
    o_Excel.Visible = True
    
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    
    ' -- Cerrar Excel
    'o_Excel.Quit

    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
    
End Function

Public Function Exportar_Excel_Totales(sOutputPath As String, Optional desde As Date, Optional hasta As Date, Optional idTipo As Integer, Optional idEmpresa As Integer) As Boolean
    
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim nombre      As String
    Dim Cantidad    As Integer
    Dim Totalpersona As Double
    
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    
    Set o_Hoja = o_Libro.Worksheets.Add
    
    'Set rsconsulta = New ADODB.Recordset
    'SQL = "SELECT c.id, p.nombre empleado, p.nrolegajo nrodoc, c.fecha, c.precio, p.idtipo FROM comidas as c "
    'SQL = SQL & "INNER JOIN empleados as p ON p.id = c.idempleado "
    'SQL = SQL & "AND date(c.fecha) BETWEEN '" & Format(desde, "yyyy-MM-dd") & "' AND '" & Format(hasta, "yyyy-MM-dd") & "'"
    'SQL = SQL & "ORDER BY p.nrolegajo, c.fecha"
    'rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Set rsconsulta = New ADODB.Recordset
    SQL = "SELECT c.id, p.id, p.nombre empleado, p.nrolegajo nrodoc, c.fecha, sum(c.precio) PRECIO, p.idtipo, t.beneficio FROM comidas as c "
    SQL = SQL & "INNER JOIN empleados as p ON p.id = c.idempleado "
    SQL = SQL & "INNER JOIN tipos as t ON p.idtipo = t.id "
    If (idTipo <> 0) Then
        SQL = SQL & "AND p.idtipo = " & idTipo & " "
    End If
    If (idEmpresa <> 0) Then
        SQL = SQL & "AND p.idempresa = " & idEmpresa & " "
    End If
    SQL = SQL & "AND date(c.fecha) BETWEEN '" & Format(desde, "yyyy-MM-dd") & "' AND '" & Format(hasta, "yyyy-MM-dd") & "'"
    SQL = SQL & "GROUP BY p.id ORDER BY p.nrolegajo, c.fecha"
    rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    
    o_Hoja.Name = Left(Replace(Replace("Detalle de Comidas", "*", ""), "/", " "), 30)
    
    total = rsconsulta.RecordCount
    
    Fila = 3
    
    zMain.pBar.Value = 0
    zMain.pBar.Height = 255
    If (total <> 0) Then
        zMain.pBar.Max = total
    End If
    
    ' Cargo Encabezado
    o_Hoja.Cells(Fila, 1).Value = "Empresa"
    o_Hoja.Cells(Fila, 2).Value = "Apellido, Nombre"
    o_Hoja.Cells(Fila, 3).Value = "Nro Legajo"
    o_Hoja.Cells(Fila, 4).Value = "Puesto"
    o_Hoja.Cells(Fila, 5).Value = "Retiró"
    o_Hoja.Cells(Fila, 6).Value = "Dia"
    o_Hoja.Cells(Fila, 7).Value = "Cantidad"
    o_Hoja.Cells(Fila, 8).Value = "Importe"
    o_Hoja.Cells(Fila, 9).Value = "Beneficio"
    o_Hoja.Cells(Fila, 10).Value = "Final"
    o_Hoja.Cells(Fila, 11).Value = "Beneficio A Facturar"
    

    
    Cantidad = 0
    Fila = Fila + 1
    ' Encabezado en negrita
    o_Excel.Range("1:1").Font.Bold = True
    o_Excel.Range("A1:B1").interior.Color = RGB(166, 166, 166)
    o_Excel.Range("3:3").Font.Bold = True
        
    ' Cargo Contenido
    Do While Not rsconsulta.EOF
    
        o_Hoja.Cells(Fila, 2).Value = rsconsulta!empleado
        o_Hoja.Cells(Fila, 3).Value = rsconsulta!nrodoc
        o_Hoja.Cells(Fila, 6).Value = rsconsulta!fecha
        o_Hoja.Cells(Fila, 6).numberformat = "dd/mm/yyyy hh:mm:ss"
        o_Hoja.Cells(Fila, 7).Value = "1"
        o_Hoja.Cells(Fila, 8).Value = CDbl(rsconsulta!Precio)
        If (rsconsulta!idTipo = 1) Then
            o_Hoja.Cells(Fila, 9).Value = CDbl(0)
            o_Hoja.Cells(Fila, 10).Value = CDbl(rsconsulta!Precio)
        ElseIf (rsconsulta!idTipo = 2) Then
            o_Hoja.Cells(Fila, 9).Value = "-" & CDbl(rsconsulta!beneficio)
            If (rsconsulta!Precio < rsconsulta!beneficio) Then
                o_Hoja.Cells(Fila, 10).Value = CDbl(0)
            Else
                o_Hoja.Cells(Fila, 10).Value = CDbl(rsconsulta!Precio) - CDbl(rsconsulta!beneficio)
            End If
        ElseIf (rsconsulta!idTipo = 3) Then
            o_Hoja.Cells(Fila, 9).Value = "-" & CDbl(rsconsulta!Precio)
            o_Hoja.Cells(Fila, 10).Value = CDbl(0)
        ElseIf (rsconsulta!idTipo = 4) Then
            o_Hoja.Cells(Fila, 9).Value = CDbl(0)
            o_Hoja.Cells(Fila, 10).Value = CDbl(rsconsulta!Precio)
        ElseIf (rsconsulta!idTipo = 5) Then
            o_Hoja.Cells(Fila, 9).Value = "-" & CDbl(rsconsulta!beneficio)
            If (rsconsulta!Precio < rsconsulta!beneficio) Then
                o_Hoja.Cells(Fila, 10).Value = CDbl(0)
            Else
                o_Hoja.Cells(Fila, 10).Value = CDbl(rsconsulta!Precio) - CDbl(rsconsulta!beneficio)
            End If
        ElseIf (rsconsulta!idTipo = 6) Then
            o_Hoja.Cells(Fila, 9).Value = "-" & CDbl(rsconsulta!beneficio)
            If (rsconsulta!Precio < rsconsulta!beneficio) Then
                o_Hoja.Cells(Fila, 10).Value = CDbl(0)
            Else
                o_Hoja.Cells(Fila, 10).Value = CDbl(rsconsulta!Precio) - CDbl(rsconsulta!beneficio)
            End If
        ElseIf (rsconsulta!idTipo = 7) Then
            o_Hoja.Cells(Fila, 9).Value = "-" & CDbl(rsconsulta!beneficio)
            If (rsconsulta!Precio < rsconsulta!beneficio) Then
                o_Hoja.Cells(Fila, 10).Value = CDbl(0)
            Else
                o_Hoja.Cells(Fila, 10).Value = CDbl(rsconsulta!Precio) - CDbl(rsconsulta!beneficio)
            End If
        End If
        If (o_Hoja.Cells(Fila, 2).Value <> "EFECTIVO" Or o_Hoja.Cells(Fila, 2).Value <> "TRANSFERENCIA") Then
            If CDbl((o_Hoja.Cells(Fila, 9).Value) * -1) >= CDbl(o_Hoja.Cells(Fila, 8).Value) Then
                o_Hoja.Cells(Fila, 11).Value = CDbl(o_Hoja.Cells(Fila, 9).Value * -1)
            Else
                o_Hoja.Cells(Fila, 11).Value = CDbl(o_Hoja.Cells(Fila, 9).Value * -1) + CDbl(o_Hoja.Cells(Fila, 8).Value)
            End If
        Else
            o_Hoja.Cells(Fila, 11).Value = CDbl(0)
        End If

        Cantidad = Cantidad + 1
        Fila = Fila + 1
        zMain.pBar.Value = zMain.pBar.Value + 1
        
        rsconsulta.MoveNext
        
    Loop

    rsconsulta.Close
    
    'Autoajusta el ancho de las columnas corte y depósito
    o_Excel.Range("A:A").EntireColumn.AutoFit
    o_Excel.Range("B:B").EntireColumn.AutoFit
    o_Excel.Range("C:C").EntireColumn.AutoFit
    o_Excel.Range("D:D").EntireColumn.AutoFit
    o_Excel.Range("E:E").EntireColumn.AutoFit
    o_Excel.Range("F:F").EntireColumn.AutoFit
    o_Excel.Columns("G").ColumnWidth = 10.67
    o_Excel.Range("H:H").EntireColumn.AutoFit
    o_Excel.Range("I:I").EntireColumn.AutoFit

    o_Hoja.Cells(Fila, 7).Value = Cantidad

    o_Hoja.Cells(1, 1).Value = "Total"
    o_Hoja.Cells(1, 2).Value = Cantidad
    
    o_Excel.Worksheets("Hoja1").Delete
    o_Excel.Worksheets("Hoja2").Delete
    o_Excel.Worksheets("Hoja3").Delete
    'o_Libro.Close True, sOutputPath
    o_Excel.Visible = True
    
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    
    ' -- Cerrar Excel
    'o_Excel.Quit

    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel_Totales = True
    
End Function


Public Function Exportar_ExcelDNI(sOutputPath As String, Optional desde As Date, Optional hasta As Date) As Boolean
    
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim nombre      As String
    Dim Cantidad    As Integer
    Dim Totalpersona As Double
    
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add

    Set o_Hoja = o_Libro.Worksheets.Add
    
    Set rsconsulta = New ADODB.Recordset
    SQL = "SELECT c.id, c.idempleado, e.nombre empleado, e.numerodocumento nrodoc, c.fecha, sum(c.precio) precio, count(c.id) cant FROM comidas as c "
    SQL = SQL & "INNER JOIN empleados as e ON e.id = c.idempleado "
    SQL = SQL & "AND date(c.fecha) BETWEEN '2023-03-01' AND '2023-03-15' "
    SQL = SQL & "GROUP BY c.idempleado ORDER BY c.idempleado "
    rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    o_Hoja.Name = Left(Replace(Replace("Total de Comidas", "*", ""), "/", " "), 30)
    
    total = rsconsulta.RecordCount
    
    Fila = 3
    
    zMain.pBar.Value = 0
    zMain.pBar.Height = 255
    If (total <> 0) Then
        zMain.pBar.Max = total
    End If
    
    ' Cargo Encabezado
    o_Hoja.Cells(Fila, 1).Value = "Empresa"
    o_Hoja.Cells(Fila, 2).Value = "Apellido, Nombre"
    o_Hoja.Cells(Fila, 3).Value = "Nro Documento"
    o_Hoja.Cells(Fila, 4).Value = "Puesto"
    o_Hoja.Cells(Fila, 5).Value = "Retiró"
    o_Hoja.Cells(Fila, 6).Value = "Dia"
    o_Hoja.Cells(Fila, 7).Value = "Cantidad"
    o_Hoja.Cells(Fila, 8).Value = "Importe"
    

    
    Cantidad = 0
    Fila = Fila + 1
    ' Encabezado en negrita
    o_Excel.Range("1:1").Font.Bold = True
    o_Excel.Range("A1:B1").interior.Color = RGB(166, 166, 166)
    o_Excel.Range("3:3").Font.Bold = True
        
    ' Cargo Contenido
    Do While Not rsconsulta.EOF
    
        o_Hoja.Cells(Fila, 2).Value = rsconsulta!empleado
        o_Hoja.Cells(Fila, 3).Value = rsconsulta!nrodoc
        o_Hoja.Cells(Fila, 6).Value = rsconsulta!fecha
        o_Hoja.Cells(Fila, 6).numberformat = "dd/mm/yyyy hh:mm:ss"
        o_Hoja.Cells(Fila, 7).Value = CDbl(rsconsulta!cant)
        o_Hoja.Cells(Fila, 8).Value = CDbl(rsconsulta!Precio)

        Cantidad = Cantidad + CDbl(rsconsulta!cant)
        Fila = Fila + 1
        zMain.pBar.Value = zMain.pBar.Value + 1
        
        rsconsulta.MoveNext
        
    Loop

    rsconsulta.Close
    
    'Autoajusta el ancho de las columnas corte y depósito
    o_Excel.Range("A:A").EntireColumn.AutoFit
    o_Excel.Range("B:B").EntireColumn.AutoFit
    o_Excel.Range("C:C").EntireColumn.AutoFit
    o_Excel.Range("D:D").EntireColumn.AutoFit
    o_Excel.Range("E:E").EntireColumn.AutoFit
    o_Excel.Range("F:F").EntireColumn.AutoFit
    o_Excel.Columns("G").ColumnWidth = 10.67

    o_Hoja.Cells(Fila, 7).Value = Cantidad

    o_Hoja.Cells(1, 1).Value = "Total"
    o_Hoja.Cells(1, 2).Value = Cantidad
    
    o_Excel.Worksheets("Hoja1").Delete
    o_Excel.Worksheets("Hoja2").Delete
    o_Excel.Worksheets("Hoja3").Delete
    'o_Libro.Close True, sOutputPath
    o_Excel.Visible = True
    
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    
    ' -- Cerrar Excel
    'o_Excel.Quit

    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_ExcelDNI = True
    
End Function


' -------------------------------------------------------------------
' \\ -- Eliminar objetos para liberar recursos
' -------------------------------------------------------------------
Public Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
    
End Sub

Function CambiaPunto(Text As TextBox, KeyAscii As Integer, Optional Char As String)
    
    Dim i As Integer
    Dim Coma As Integer
    
    'Si hay un punto escribe una coma
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    
    'Busca la coma en la cadena
    Coma = InStr(Text.Text, ",")
    
    'Si acepta varios importes, se elimina la restricción de la cantidad de comas
    If InStrRev(Text.Text, "+") > InStrRev(Text.Text, ",") = True Then
        Coma = 0
    End If
    
    'Si ya hay una coma, no acepta otra
    If Coma <> 0 Then
        If InStr("0123456789" & Char, Chr(KeyAscii)) = 0 Then
            If KeyAscii <> 8 Then KeyAscii = 0
        End If
    
    'Si no hay una coma, acepta una coma
    Else
        If InStr("0123456789," & Char, Chr(KeyAscii)) = 0 Then
            If KeyAscii <> 8 Then KeyAscii = 0
        End If
    End If
    
End Function

Function SoloEnteros(Text As TextBox, KeyAscii As Integer)
    
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        If KeyAscii <> 8 Then KeyAscii = 0
    End If
    
End Function

Function getData(id As Variant, field As String, table As String)
    
    Dim rsData As ADODB.Recordset
    
    getData = ""
    
    Set rsData = New ADODB.Recordset
    SQL = "SELECT id, " & table & "." & field & " AS data FROM " & table & " WHERE id = '" & id & "'"
    rsData.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsData.BOF And Not rsData.EOF Then
        If Not IsNull(rsData!Data) Then
            getData = Trim(rsData!Data)
        End If
    End If
    rsData.Close
    
End Function

Function SumaStock(idDeposito As Integer, idCorte As Integer, idAnimal As Integer, Cantidad As Integer, Kilogramos As String, Optional nTropa As String, Optional Reparto As String, Optional Precio As String)
    
    Set rsStock = New ADODB.Recordset
    SQL = "SELECT * FROM stock WHERE iddeposito = '" & idDeposito & "' AND idcorte = '" & idCorte & "' AND idanimal = '" & idAnimal & "'"
    
    If nTropa <> "" Then
        SQL = SQL & " AND ntropa = '" & nTropa & "'"
    End If
    
    If Reparto <> "" Then
        SQL = SQL & " AND reparto = '" & Reparto & "'"
    End If
    
    If Precio = "" Then
        Precio = "0,00"
    End If
    
    rsStock.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsStock.BOF And Not rsStock.EOF Then
        rsStock!Cantidad = CDec(rsStock!Cantidad) + Cantidad
        rsStock!Kilogramos = Format(CDec(rsStock!Kilogramos) + CDec(Kilogramos), "0.00")
        rsStock.Update
    Else
        rsStock.AddNew
        rsStock!idDeposito = idDeposito
        rsStock!idCorte = idCorte
        rsStock!idAnimal = idAnimal
        rsStock!Cantidad = Cantidad
        rsStock!Kilogramos = Format(CDec(Kilogramos), "0.00")
        rsStock!nTropa = nTropa
        rsStock!Reparto = Reparto
        rsStock!Precio = Precio
        rsStock.Update
    End If
    rsStock.Close
    
End Function

Function RestaStock(idDeposito As Integer, idCorte As Integer, idAnimal As Integer, Cantidad As Integer, Kilogramos As String, Optional nTropa As String, Optional Reparto As String)
    
    Set rsStock = New ADODB.Recordset
    SQL = "SELECT * FROM stock WHERE iddeposito = '" & idDeposito & "' AND idcorte = '" & idCorte & "' AND idanimal = '" & idAnimal & "'"
    
    If nTropa <> "" Then
        SQL = SQL & " AND ntropa = '" & nTropa & "'"
    End If
    
    If Reparto <> "" Then
        SQL = SQL & " AND reparto = '" & Reparto & "'"
    End If
    
    rsStock.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsStock.BOF And Not rsStock.EOF Then
        rsStock!Cantidad = CDec(rsStock!Cantidad) - Cantidad
        rsStock!Kilogramos = Format(CDec(rsStock!Kilogramos) - CDec(Kilogramos), "0.00")
        rsStock.Update
        
        'Si el Stock queda en 0, elimina el registro
        If rsStock!Cantidad = 0 Then
            rsStock.Delete
        End If
        
    End If
    rsStock.Close
    
End Function

Function BorraReg(table As String)
    
    'Borra todos los registros de una tabla (table)
    Set rsDelTable = New ADODB.Recordset
    SQL = "DELETE FROM " & table & ";"
    rsDelTable.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
End Function

Public Sub CargaCombo(rsTable As String, rsField As String, rsOrder As String, Cbo As ComboBox, _
                        Optional Todos As Boolean, Optional Condicion As String)
    
    'Carga los combos
    Set rsCargaCombo = New ADODB.Recordset
    SQL = "SELECT id, " & rsField & " AS field FROM " & rsTable & " WHERE eliminado = '0'"
    If Condicion <> "" Then
        SQL = SQL & " AND " & Condicion
    End If
    SQL = SQL & " ORDER BY " & rsOrder
    
    rsCargaCombo.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Cbo.Clear
    
    If Todos = True Then
        Cbo.AddItem "Todos"
        Cbo.ItemData(Cbo.NewIndex) = 0
    End If
    
    Do While Not rsCargaCombo.EOF
        Cbo.AddItem rsCargaCombo!field
        Cbo.ItemData(Cbo.NewIndex) = rsCargaCombo!id
        rsCargaCombo.MoveNext
    Loop
    
    rsCargaCombo.Close
    
End Sub

Public Sub PasarFoco()
    
    'If KeyCode = 13 Then
       ' envía la pulsación de tecla Tab y pasa el foco _
         a la siguiente caja de texto
       SendKeys "{TAB}"
    'End If
    
End Sub

Public Sub CerrarVentanas()

    On Error GoTo Errores
    'Unload ActiveForm
    
    Exit Sub
Errores:
    If Err.Number = 91 Then
        Exit Sub
    End If
    
End Sub

Public Sub initForm(Frm As Form)
    
    Frm.Top = (zMain.ScaleHeight / 2) - (Frm.Height / 2)
    Frm.Left = (zMain.ScaleWidth / 2) - (Frm.Width / 2)
    
End Sub

Public Sub MayusBox(Frm As Form)
    
    ' recorre todos los controles que hay en el formulario
    For Each Control In Frm.Controls
        'verifica que el control es de tipo TextBox o Combobox
        If TypeOf Control Is TextBox Then
            Control.Text = UCase(Control.Text)
        End If
    Next
    
End Sub

Public Sub LimpiarBox(Frm As Form)
    
    ' recorre todos los controles que hay en el formulario
    For Each Control In Frm.Controls
        'verifica que el control es de tipo TextBox o Combobox
        If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
            '... Si es un Textbox o Combobox, entonces lo limpia
            Control.Text = ""
        End If
    Next
    
End Sub

Public Sub BlanquearBox(Frm As Form)
    
    ' recorre todos los controles que hay en el formulario
    Dim Control As Control
    For Each Control In Frm.Controls
        'verifica que el control es de tipo TextBox o Combobox
        If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
            '... Si es un Textbox o Combobox, entonces lo limpia
            Control.BackColor = vbWhite
        End If
    Next
    
End Sub

Sub IvaVenta()
    
    'zMain.pBar.Value = 0
    'zMain.sBar.Height = 0
    'zMain.pBar.Height = 255
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT id, iva, ivaventas FROM articulos"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    zMain.pBar.Max = Recordset.RecordCount
    
    Do While Not Recordset.EOF
        
        If Recordset!ivaventas = "01" Then
            Recordset!IVA = "21"
        ElseIf Recordset!ivaventas = "03" Then
            Recordset!IVA = "10,5"
        ElseIf Recordset!ivaventas = "05" Then
            Recordset!IVA = "27"
        ElseIf Recordset!ivaventas = "11" Then
            Recordset!IVA = "00"
        End If
        Recordset.Update
        Recordset.MoveNext
        'zMain.pBar.Value = zMain.pBar.Value + 1
        
    Loop
    Recordset.Close
    
   ' zMain.pBar.Value = 0
   ' zMain.pBar.Height = 0
   ' zMain.sBar.Height = 255
    MsgBox "LISTO!"
    
End Sub

Function CalculaDescuentos(Lista As String, sDesc As String)
    
    'Si incluye más de un descuento, sDesc debe tener el formato: 'd1+d2+d3'
    
    'Si encuentra '++' sale
    If InStr(1, sDesc, "++") Then Exit Function
    
    'Si el último caracter es un '+', lo omite
    If Right(sDesc, 1) = "+" Then
        sDesc = Left(sDesc, Len(sDesc) - 1)
    End If
    
    'Divide los descuentos en un array
    d = Split(sDesc, "+")
    p = Lista
    
    'Calcula los descuentos
    For i = 0 To UBound(d)
        p = CDec(p) - ((CDec(p) * CDec(d(i))) / 100)
    Next i
    
    'Retorna el precio
    CalculaDescuentos = Format(p, "0.000")
    
End Function

Function CalculaRecargos(Lista As String, sPorc As String)
    
    'Si incluye más de un recargo, sPorc debe tener el formato: 'd1+d2+d3'
    
    'Si encuentra '++' sale
    If InStr(1, sPorc, "++") Then Exit Function
    
    'Si el último caracter es un '+', lo omite
    If Right(sPorc, 1) = "+" Then
        sPorc = Left(sPorc, Len(sPorc) - 1)
    End If
    
    'Divide los descuentos en un array
    d = Split(sPorc, "+")
    p = Lista
    
    'Calcula los descuentos
    For i = 0 To UBound(d)
        p = CDec(p) + ((CDec(p) * CDec(d(i))) / 100)
    Next i
    
    'Retorna el precio
    CalculaRecargos = Format(p, "0.000")
    
End Function

Function existeArchivo(strFile As String) As Boolean
    
    Dim strName As String
    strName = Dir$(strFile)
    existeArchivo = Not (strName = "")
    
End Function

'Función que devuelve el número aleatorio
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Aleatorio(Minimo As Long, Maximo As Long) As Long
    Randomize ' inicializar la semilla
    Aleatorio = CLng((Minimo - Maximo) * Rnd + Maximo)
End Function

Public Sub cambioImporte()

    Dim conexion As ADODB.Connection
    Dim rsExcel As ADODB.Recordset
    Dim cantidada As Integer
    
    'Cantidad = 0
    
    hoja = "Hoja1$"

    'Abre el Excel
    Set conexion = New ADODB.Connection
    conexion.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source= C:\CAMBIO.xls" & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
    
    Set rsExcel = New ADODB.Recordset
    With rsExcel
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
    SQL = "SELECT * FROM [" & hoja & "]"
    rsExcel.Open SQL, conexion, , , adCmdText
    
    Do While Not rsExcel.EOF
        Set rsActualiza = New ADODB.Recordset
        SQL = "UPDATE comidas SET precio = '" & Format(rsExcel!Precio, "0.00") & "' WHERE id = " & rsExcel!id
        rsActualiza.Open SQL, Data, adOpenKeyset, adLockOptimistic
        'Cantidad = Cantidad + 1
        rsExcel.MoveNext
    Loop
    
    rsExcel.Close
    
    'MsgBox ("Se modificaron: " & Cantidad)
        

End Sub


Public Sub ImportarExcel()
    
    Dim cantidadempleado As Integer
    
    Dim id As Double
    
    Dim Existe As Boolean
    
    Dim conexion As ADODB.Connection
    Dim rsExcel As ADODB.Recordset
    
    id = 0
    
    hoja = "Hoja1$"
    
    'Muestra la barra
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    zMain.pBar.Height = 255
    
    'Abre el Excel
    Set conexion = New ADODB.Connection
    conexion.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source= C:\Importar.xls" & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
    
    Set rsExcel = New ADODB.Recordset
    With rsExcel
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
    SQL = "SELECT * FROM [" & hoja & "]"
        rsExcel.Open SQL, conexion, , , adCmdText
    zMain.pBar.Max = rsExcel.RecordCount
    
    Do While Not rsExcel.EOF
        If rsExcel!nombre <> "" Then
                
            'Verificamos que exista el empleado
            Set rsEmpleado = New ADODB.Recordset
            SQL = "SELECT * FROM empleados WHERE numerodocumento = " & rsExcel!DNI & " AND eliminado = 0"
            rsEmpleado.Open SQL, Data, adOpenKeyset, adLockOptimistic
            If rsEmpleado.BOF And rsEmpleado.EOF Then
                cantidadempleado = cantidadempleado + 1
                rsEmpleado.AddNew
                rsEmpleado!nrolegajo = rsExcel!legajo
                rsEmpleado!nombre = rsExcel!nombre
                rsEmpleado!numerodocumento = rsExcel!DNI
                rsEmpleado!idTipo = rsExcel!tipo
                Existe = False
            Else
                Existe = True
                rsEmpleado!nrolegajo = rsExcel!legajo
                rsEmpleado!numerodocumento = rsExcel!DNI
                rsEmpleado!idTipo = rsExcel!tipo
                
            End If
            
            rsEmpleado.Update
            rsEmpleado.Close
            
        End If
        zMain.pBar.Value = zMain.pBar.Value + 1
        rsExcel.MoveNext
    Loop
    rsExcel.Close
    
    'Oculta la barra
    zMain.pBar.Value = 0
    zMain.pBar.Height = 0
    zMain.pBar.Height = 255
    
    Call MsgBox("Se han agregado " & cantidadempleado & " empleados nuevos ", vbInformation, App.Title)
    
    
End Sub

Public Sub HookForm(Obj As Object)
    PrevProc = SetWindowLong(Obj.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookForm(Obj As Object)
    SetWindowLong Obj.hwnd, GWL_WNDPROC, PrevProc
End Sub

' Procedimiento qie intercepta los mensajes de windows, en este caso para _
  interceptar el uso del Scroll del mouse
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    
    If uMsg = WM_MOUSEWHEEL Then
        If wParam < 0 Then
            ' envia mediante SendMessage el comando para mover el Scroll hacia abajo
            SendMessage hwnd, WM_VSCROLL, ByVal 1, ByVal 0
        Else
            ' Mueve el scroll hacia arriba
            SendMessage hwnd, WM_VSCROLL, ByVal 0, ByVal 0
        End If
    End If
End Function

Function Dia16Anterior(ByVal fecha As Date) As Date
    ' Si el día de la fecha proporcionada es mayor que 16
    ' Devuelve el día 16 del mismo mes
    ' De lo contrario, devuelve el día 16 del mes anterior
    If Day(fecha) > 16 Then
        Dia16Anterior = DateSerial(Year(fecha), Month(fecha), 16)
    Else
        ' Si estamos en enero y el día es menor o igual a 16
        ' Devuelve el 16 de diciembre del año anterior
        If Month(fecha) = 1 Then
            Dia16Anterior = DateSerial(Year(fecha) - 1, 12, 16)
        Else
            Dia16Anterior = DateSerial(Year(fecha), Month(fecha) - 1, 16)
        End If
    End If
End Function
