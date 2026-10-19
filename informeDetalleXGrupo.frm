VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form informeDetalleXGrupo
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle Por Grupo"
   ClientHeight    =   3105
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
   ScaleHeight     =   3105
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CboEmpresa
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
      ItemData        =   "informeDetalleXGrupo.frx":0000
      Left            =   120
      List            =   "informeDetalleXGrupo.frx":004C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
   End
   Begin VB.ComboBox cboTipos
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
      ItemData        =   "informeDetalleXGrupo.frx":0161
      Left            =   120
      List            =   "informeDetalleXGrupo.frx":01AD
      Style           =   2  'Dropdown List
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
      Picture         =   "informeDetalleXGrupo.frx":02C2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Salir "
      Top             =   2280
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
      Format          =   36503553
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
      Format          =   36503553
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
      Picture         =   "informeDetalleXGrupo.frx":084C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Exportar "
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1
      Caption         =   "Empresa"
      BeginProperty Font
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label2
      Caption         =   "Tipo Empleado"
      BeginProperty Font
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblHasta
      Caption         =   "Hasta"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblFechaDesde
      Caption         =   "Desde"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "informeDetalleXGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()

    Dim desde As Date
    Dim hasta As Date
    Dim idTipo As Integer
    Dim idEmpresa As Integer

    desde = dtpDesde.Value
    hasta = dtpHasta.Value

    If desde > hasta Then
        Call MsgBox("La fecha desde no puede ser mayor a la fecha hasta.", vbExclamation, "Fechas")
        Exit Sub
    End If

    idTipo = cboTipos.ItemData(cboTipos.ListIndex)
    idEmpresa = CboEmpresa.ItemData(CboEmpresa.ListIndex)

    Screen.MousePointer = vbHourglass
    Exportar_Excel_DetalleXGrupo desde, hasta, idTipo, idEmpresa
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    initForm Me
    If Day(Now) < 16 Then
        dtpDesde.Value = DateSerial(Year(Now), Month(Now), 1)
        dtpHasta.Value = DateSerial(Year(Now), Month(Now), 15)
    Else
        dtpDesde.Value = DateSerial(Year(Now), Month(Now), 16)
        dtpHasta.Value = DateSerial(Year(Now), Month(Now) + 1, 0)
    End If

    CargaCombo "tipos", "nombre", "id", cboTipos, True
    CargaCombo "empresas", "nombre", "id", CboEmpresa, True
    cboTipos.ListIndex = 0
    CboEmpresa.ListIndex = 0

End Sub

'---------------------------------------------------------------------------------------
' \\ -- Exporta a Excel el detalle de comidas de cada empleado del grupo (empresa/tipo),
' \\    un bloque por empleado con sus comidas día por día y un subtotal.
'---------------------------------------------------------------------------------------

Private Function Exportar_Excel_DetalleXGrupo(desde As Date, hasta As Date, idTipo As Integer, idEmpresa As Integer) As Boolean

    Dim o_Excel         As Object
    Dim o_Libro         As Object
    Dim o_Hoja          As Object
    Dim rsconsulta      As ADODB.Recordset
    Dim SQL             As String
    Dim Fila            As Long
    Dim idActual        As Long
    Dim cantEmpleado    As Double
    Dim totalEmpleado   As Double
    Dim cantGeneral     As Double
    Dim totalGeneral    As Double
    Dim empleados       As Long

    ' Consulta: detalle de cada ticket de los empleados del grupo
    SQL = "SELECT e.id AS idemp, e.nombre AS empleado, e.nrolegajo, e.numerodocumento, " & _
          "em.nombre AS empresa, t.nombre AS tipo, " & _
          "c.fecha, c.numero, d.codigo, d.nombre AS articulo, d.cantidad, d.precio, d.total " & _
          "FROM comidas AS c " & _
          "INNER JOIN empleados AS e ON e.id = c.idempleado " & _
          "INNER JOIN comidasd AS d ON d.idcomida = c.id " & _
          "LEFT JOIN empresas AS em ON em.id = e.idempresa " & _
          "LEFT JOIN tipos AS t ON t.id = e.idtipo " & _
          "WHERE date(c.fecha) BETWEEN '" & Format(desde, "yyyy-MM-dd") & "' AND '" & Format(hasta, "yyyy-MM-dd") & "' " & _
          "AND c.eliminado = 0 AND c.idempleado <> 9999 "
    If idTipo <> 0 Then
        SQL = SQL & "AND e.idtipo = " & idTipo & " "
    End If
    If idEmpresa <> 0 Then
        SQL = SQL & "AND e.idempresa = " & idEmpresa & " "
    End If
    ' Orden: Empresa --> Tipo de empleado --> Nombre
    SQL = SQL & "ORDER BY em.nombre, t.nombre, e.nombre, e.id, c.fecha, d.codigo"

    Set rsconsulta = New ADODB.Recordset
    rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic

    If rsconsulta.BOF And rsconsulta.EOF Then
        rsconsulta.Close
        Call MsgBox("No hay comidas para el grupo y las fechas seleccionadas.", vbExclamation, "No hay datos")
        Exit Function
    End If

    ' Crear Excel
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets(1)
    o_Hoja.Name = "Detalle Por Grupo"

    zMain.pBar.Value = 0
    zMain.pBar.Max = rsconsulta.RecordCount
    zMain.pBar.Height = 255

    ' Título
    o_Hoja.Cells(1, 1).Value = "Detalle de comidas por empleado"
    o_Hoja.Cells(1, 1).Font.Bold = True
    o_Hoja.Cells(1, 1).Font.Size = 14
    o_Hoja.Cells(2, 1).Value = "Período: " & Format(desde, "dd/mm/yyyy") & " al " & Format(hasta, "dd/mm/yyyy")
    o_Hoja.Cells(3, 1).Value = "Empresa: " & CboEmpresa.Text & "   -   Tipo: " & cboTipos.Text

    Fila = 5
    idActual = -1

    Do While Not rsconsulta.EOF

        ' Cambio de empleado: cierra el bloque anterior y abre uno nuevo
        If rsconsulta!idemp <> idActual Then

            If idActual <> -1 Then
                EscribirSubtotal o_Hoja, Fila, cantEmpleado, totalEmpleado
                Fila = Fila + 2
            End If

            idActual = rsconsulta!idemp
            cantEmpleado = 0
            totalEmpleado = 0
            empleados = empleados + 1

            o_Hoja.Cells(Fila, 1).Value = rsconsulta!empleado & "  -  Legajo: " & rsconsulta!nrolegajo & _
                                          "  -  DNI: " & rsconsulta!numerodocumento & _
                                          "  -  " & rsconsulta!empresa & "  -  " & rsconsulta!tipo
            o_Hoja.Range(o_Hoja.Cells(Fila, 1), o_Hoja.Cells(Fila, 8)).Font.Bold = True
            o_Hoja.Range(o_Hoja.Cells(Fila, 1), o_Hoja.Cells(Fila, 8)).Interior.Color = RGB(166, 166, 166)
            Fila = Fila + 1

            o_Hoja.Cells(Fila, 1).Value = "Fecha"
            o_Hoja.Cells(Fila, 2).Value = "Hora"
            o_Hoja.Cells(Fila, 3).Value = "Ticket"
            o_Hoja.Cells(Fila, 4).Value = "Código"
            o_Hoja.Cells(Fila, 5).Value = "Artículo"
            o_Hoja.Cells(Fila, 6).Value = "Cantidad"
            o_Hoja.Cells(Fila, 7).Value = "Precio"
            o_Hoja.Cells(Fila, 8).Value = "Total"
            o_Hoja.Range(o_Hoja.Cells(Fila, 1), o_Hoja.Cells(Fila, 8)).Font.Bold = True
            Fila = Fila + 1

        End If

        o_Hoja.Cells(Fila, 1).Value = DateValue(rsconsulta!fecha)
        o_Hoja.Cells(Fila, 1).NumberFormat = "dd/mm/yyyy"
        o_Hoja.Cells(Fila, 2).Value = TimeValue(rsconsulta!fecha)
        o_Hoja.Cells(Fila, 2).NumberFormat = "hh:mm"
        o_Hoja.Cells(Fila, 3).Value = rsconsulta!numero
        o_Hoja.Cells(Fila, 4).Value = rsconsulta!codigo
        o_Hoja.Cells(Fila, 5).Value = rsconsulta!articulo
        o_Hoja.Cells(Fila, 6).Value = CDbl(rsconsulta!Cantidad)
        o_Hoja.Cells(Fila, 7).Value = CDbl(rsconsulta!Precio)
        o_Hoja.Cells(Fila, 7).NumberFormat = "#,##0.00"
        o_Hoja.Cells(Fila, 8).Value = CDbl(rsconsulta!total)
        o_Hoja.Cells(Fila, 8).NumberFormat = "#,##0.00"

        cantEmpleado = cantEmpleado + CDbl(rsconsulta!Cantidad)
        totalEmpleado = totalEmpleado + CDbl(rsconsulta!total)
        cantGeneral = cantGeneral + CDbl(rsconsulta!Cantidad)
        totalGeneral = totalGeneral + CDbl(rsconsulta!total)

        Fila = Fila + 1
        zMain.pBar.Value = zMain.pBar.Value + 1
        rsconsulta.MoveNext

    Loop

    rsconsulta.Close

    ' Subtotal del último empleado y total general
    EscribirSubtotal o_Hoja, Fila, cantEmpleado, totalEmpleado
    Fila = Fila + 2

    o_Hoja.Cells(Fila, 5).Value = "TOTAL GENERAL (" & empleados & " empleados):"
    o_Hoja.Cells(Fila, 6).Value = cantGeneral
    o_Hoja.Cells(Fila, 8).Value = totalGeneral
    o_Hoja.Cells(Fila, 8).NumberFormat = "#,##0.00"
    o_Hoja.Range(o_Hoja.Cells(Fila, 1), o_Hoja.Cells(Fila, 8)).Font.Bold = True

    ' Ajustar columnas (la A no, porque tiene los títulos largos)
    o_Hoja.Columns("A").ColumnWidth = 12
    o_Hoja.Columns("B:H").EntireColumn.AutoFit

    zMain.pBar.Value = 0
    zMain.pBar.Height = 0

    ' Mostrar Excel
    o_Excel.Visible = True

    Exportar_Excel_DetalleXGrupo = True

    ' Liberar objetos
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)

End Function

Private Sub EscribirSubtotal(o_Hoja As Object, Fila As Long, Cantidad As Double, total As Double)

    o_Hoja.Cells(Fila, 5).Value = "Subtotal:"
    o_Hoja.Cells(Fila, 6).Value = Cantidad
    o_Hoja.Cells(Fila, 8).Value = total
    o_Hoja.Cells(Fila, 8).NumberFormat = "#,##0.00"
    o_Hoja.Range(o_Hoja.Cells(Fila, 5), o_Hoja.Cells(Fila, 8)).Font.Bold = True

End Sub
