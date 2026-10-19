VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form informeDetalleXPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Totales Por Persona"
   ClientHeight    =   2520
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
   ScaleHeight     =   2520
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtNroDocumento 
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
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
      Left            =   3480
      Picture         =   "informeDetalleXPersona.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Salir "
      Top             =   1680
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
      Format          =   136642561
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
      Format          =   136642561
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
      Picture         =   "informeDetalleXPersona.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Exportar "
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Documento"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblHasta 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblFechaDesde 
      Caption         =   "Desde"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "informeDetalleXPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()

    Dim empresa As String
    Dim idturno As Integer
    Dim desde As Date
    Dim hasta As Date

    desde = dtpDesde.Value
    hasta = dtpHasta.Value
    
    Exportar_Excel_ComidasDetallado "C:\Sistema\Listado.xls", TxtNroDocumento.Text, desde, hasta

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

Private Function Exportar_Excel_ComidasDetallado(sOutputPath As String, nroDocumento As String, desde As Date, hasta As Date) As Boolean
    
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim filaInicio  As Long

    ' Crear Excel
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    o_Hoja.Name = "Detalle Comidas"
    
    ' Consulta SQL
    Dim SQL As String
    Set rsconsulta = New ADODB.Recordset

    SQL = "SELECT c.fecha, c.numero, e.nombre AS empleado, d.codigo, d.nombre AS articulo, " & _
          "d.cantidad, d.precio, d.total " & _
          "FROM comidas AS c " & _
          "INNER JOIN empleados AS e ON e.id = c.idempleado " & _
          "INNER JOIN comidasd AS d ON d.idcomida = c.id " & _
          "WHERE e.numerodocumento = '" & Trim(nroDocumento) & "' " & _
          "AND date(c.fecha) BETWEEN '" & Format(desde, "yyyy-MM-dd") & "' AND '" & Format(hasta, "yyyy-MM-dd") & "' " & _
          "AND c.eliminado = 0 AND e.eliminado = 0 " & _
          "ORDER BY c.fecha, d.codigo"
    
    rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic

    ' Encabezados
    Fila = 1
    o_Hoja.Cells(Fila, 1).Value = "Fecha"
    o_Hoja.Cells(Fila, 2).Value = "Número"
    o_Hoja.Cells(Fila, 3).Value = "Empleado"
    o_Hoja.Cells(Fila, 4).Value = "Código"
    o_Hoja.Cells(Fila, 5).Value = "Artículo"
    o_Hoja.Cells(Fila, 6).Value = "Cantidad"
    o_Hoja.Cells(Fila, 7).Value = "Precio"
    o_Hoja.Cells(Fila, 8).Value = "Total"
    
    o_Excel.Range("1:1").Font.Bold = True

    ' Cargar contenido
    Fila = 2
    filaInicio = Fila ' guardamos dónde empiezan los datos
    Do While Not rsconsulta.EOF
        
        o_Hoja.Cells(Fila, 1).Value = rsconsulta!fecha
        o_Hoja.Cells(Fila, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        o_Hoja.Cells(Fila, 2).Value = rsconsulta!numero
        o_Hoja.Cells(Fila, 3).Value = rsconsulta!empleado
        o_Hoja.Cells(Fila, 4).Value = rsconsulta!codigo
        o_Hoja.Cells(Fila, 5).Value = rsconsulta!articulo
        o_Hoja.Cells(Fila, 6).Value = rsconsulta!Cantidad
        o_Hoja.Cells(Fila, 7).Value = CDbl(rsconsulta!Precio)
        o_Hoja.Cells(Fila, 8).Value = CDbl(rsconsulta!total)

        Fila = Fila + 1
        rsconsulta.MoveNext
    Loop
    
    rsconsulta.Close
    
    If Fila > filaInicio Then
        o_Hoja.Cells(Fila, 7).Value = "TOTAL:"
        o_Hoja.Cells(Fila, 7).Font.Bold = True
        
        ' Fórmula de suma desde filaInicio hasta la fila anterior
        o_Hoja.Cells(Fila, 8).Formula = "=SUM(H" & filaInicio & ":H" & (Fila - 1) & ")"
        o_Hoja.Cells(Fila, 8).Font.Bold = True
        o_Hoja.Cells(Fila, 8).NumberFormat = "#,##0.00"
    End If
    
    
    ' Ajustar columnas
    o_Hoja.Columns("A:H").EntireColumn.AutoFit

    ' Borrar hojas restantes
    On Error Resume Next
    o_Excel.Worksheets("Hoja1").Delete
    o_Excel.Worksheets("Hoja2").Delete
    o_Excel.Worksheets("Hoja3").Delete
    On Error GoTo 0

    ' Mostrar Excel
    o_Excel.Visible = True
    
    Exportar_Excel_ComidasDetallado = True
    
    ' Liberar objetos
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)

End Function

