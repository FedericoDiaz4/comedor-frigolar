VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Ingreso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso"
   ClientHeight    =   14565
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   9720
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
   MinButton       =   0   'False
   ScaleHeight     =   14565
   ScaleWidth      =   9720
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      Height          =   615
      Left            =   6120
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   615
      Left            =   6120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdZeta 
      Caption         =   "&Cierre Z"
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
      Picture         =   "Ingreso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Guardar"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtDocumento 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   10095
      Left            =   0
      TabIndex        =   4
      Top             =   4320
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
   Begin VB.Image Image 
      Height          =   2350
      Left            =   720
      Picture         =   "Ingreso.frx":058A
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2300
   End
   Begin VB.Label lblOpciones 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese la opcion deseada"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1170
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese documento y presione ENTER"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   210
      Width           =   4095
   End
End
Attribute VB_Name = "Ingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inicio As Boolean

Private Sub Zeta(blanco As Boolean)

    Dim fechadesde As Date
    Dim Cantidad As Double
    Dim total As Double
    
    'Cargo cantidades y totales 0 en el flex
    For Fila = 1 To Flex.Rows - 1
        Flex.TextMatrix(Fila, 4) = 0
        Flex.TextMatrix(Fila, 5) = Format(0, "$ 0.00")
    Next
    
    Cantidad = 0
    total = 0
    
    Set rsconsulta = New ADODB.Recordset
    SQL = "SELECT * FROM zeta ORDER BY id DESC LIMIT 1"
    rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If rsconsulta.EOF And rsconsulta.BOF Then
        fechadesde = "06-04-2020 00:00:00"
        Set rsZeta = New ADODB.Recordset
        If blanco Then
            SQL = "SELECT id, precio FROM comidas WHERE idempleado <> 9999"
        Else
            SQL = "SELECT id, precio FROM comidas WHERE idempleado = 9999"
        End If
        rsZeta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Else
        fechadesde = rsconsulta!fecha
        Set rsZeta = New ADODB.Recordset
        If blanco Then
            SQL = "SELECT id FROM comidas WHERE idempleado <> 9999 AND fecha > '" & Format(fechadesde, "yyyy-mm-dd hh:mm:ss") & "' "
        Else
            SQL = "SELECT id FROM comidas WHERE idempleado = 9999 AND fecha > '" & Format(fechadesde, "yyyy-mm-dd hh:mm:ss") & "' "
        End If
        rsZeta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    End If
    
    If Not rsZeta.EOF And Not rsZeta.BOF Then
        Do While Not rsZeta.EOF
            Cantidad = Cantidad + 1
            CargoDetalle rsZeta!id
            rsZeta.MoveNext
        Loop
    Else
        Call MsgBox("No existen tickets desde el último cierre Z.", vbExclamation, "No hay datos")
        Exit Sub
    End If
    
    For Fila = 1 To Flex.Rows - 1
        total = total + CDbl(Flex.TextMatrix(Fila, 5))
    Next
    
    rsZeta.Close
    rsconsulta.Close
    
    imprimeZeta fechadesde, Cantidad, total

End Sub

Private Sub imprimeZeta(fecha As Date, Cantidad As Double, total As Double)

    Printer.FontName = "Consolas"
    Printer.FontBold = True
    Printer.FontSize = 16
    Printer.Print Centrar("Different Food", 15)
    
    Printer.FontName = "Consolas"
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.Print Centrar("Cierre de Caja", 18)
    
    Printer.FontBold = False
    Printer.FontName = "Calibri"
    Printer.Print ""
    
    Printer.FontName = "Courier New"
    Printer.FontBold = True
    Printer.FontSize = 8
    Printer.Print "Desde: " & Format(fecha, "dd/mm/yyyy hh:mm:ss")
    
    Printer.FontName = "Courier New"
    Printer.FontBold = True
    Printer.FontSize = 8
    Printer.Print "Hasta: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Printer.FontBold = False
    Printer.FontName = "Calibri"
    Printer.Print ""
    
    Printer.FontName = "Courier New"
    Printer.FontBold = True
    Printer.FontSize = 10
    Printer.Print "Cantidad: " & Cantidad
    
    Printer.FontName = "Courier New"
    Printer.FontBold = True
    Printer.FontSize = 10
    Printer.Print "Total: " & Format(total, "$ 0.00")
    
    Printer.FontBold = False
    Printer.FontName = "Calibri"
    Printer.Print ""
    
    Printer.FontBold = True
    Printer.FontName = "Courier New"
    Printer.FontSize = 8
    Printer.Print "Detalle: "
    
    For Fila = 1 To Flex.Rows - 1
        Printer.FontName = "Courier New"
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.Print DetalleZ(Flex.TextMatrix(Fila, 2), Flex.TextMatrix(Fila, 4))
    Next

    Printer.EndDoc

End Sub

Private Sub cmdZeta_Click()

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
    
    Set rsconsulta = New ADODB.Recordset
    SQL = "SELECT * FROM zeta"
    rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    rsconsulta.AddNew
    rsconsulta!fecha = Now
    rsconsulta.Update
    rsconsulta.Close
    
    MsgBox ("Listo!")

End Sub

Private Sub Command_Click()

    If existeArchivo("C:\diferencia.txt") Then
        Kill "C:\diferencia.txt"
    End If
    
    Dim totalDetalle As Double

    Open "C:\diferencia.txt" For Output As #1
    
    Set rsComida = New ADODB.Recordset
    SQL = "SELECT id, idempleado, precio FROM comidas"
    rsComida.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not rsComida.EOF
        Set rsDetalle = New ADODB.Recordset
        SQL1 = "SELECT * FROM comidasd WHERE idcomida = " & rsComida!id
        rsDetalle.Open SQL1, Data, adOpenKeyset, adLockOptimistic
        totalDetalle = 0
        Do While Not rsDetalle.EOF
            totalDetalle = totalDetalle + rsDetalle!total
            rsDetalle.MoveNext
        Loop
        If totalDetalle <> rsComida!Precio Then
            Print #1, rsComida!id & " " & rsComida!idEmpleado & " " & rsComida!Precio & " " & totalDetalle
        End If
        rsComida.MoveNext
    Loop
    
    Close #1
    
    MsgBox ("Listo")

End Sub

Private Sub Command1_Click()

    cambioImporte

End Sub

Private Sub Command2_Click()
    
    Flex.Clear
    Flex.Rows = 2
    Flex.Cols = 5
    
    Flex.FormatString = "Id|Codigo|Nombre|Cantidad|Total"
    Flex.ColWidth(0) = 1000
    Flex.ColWidth(1) = 1100
    Flex.ColWidth(2) = 2000
    Flex.ColWidth(3) = 1000
    Flex.ColWidth(4) = 1000
    
    Set rsComida = New ADODB.Recordset
    SQL = "SELECT * FROM comidas WHERE idempleado <> 9999 AND fecha > '2020-04-08 05:54:53'"
    rsComida.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Do While Not rsComida.EOF
        Set rsComidad = New ADODB.Recordset
        SQL = "SELECT * FROM comidasd WHERE idcomida = " & rsComida!id
        rsComidad.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        Flex.TextMatrix(Flex.Rows - 1, 0) = rsComida!id
        Flex.TextMatrix(Flex.Rows - 1, 1) = rsComidad!codigo
        Flex.TextMatrix(Flex.Rows - 1, 2) = rsComidad!nombre
        Flex.TextMatrix(Flex.Rows - 1, 3) = rsComidad!Cantidad
        Flex.TextMatrix(Flex.Rows - 1, 4) = rsComidad!total
        Flex.Rows = Flex.Rows + 1
        rsComida.MoveNext
    Loop
    
    Flex.Rows = Flex.Rows - 1

    rsComidad.Close
    rsComida.Close

    Exportar_Excel1 "C:\qwe\excel.xls", Flex
    
    
    MsgBox ("listo")

End Sub

Private Sub Form_Load()

    Call ConectarDB
    backup
    Ingreso.Width = 4425
    Ingreso.Height = 4950
    Ingreso.Top = (Screen.Height / 2) - (Me.Height / 2)
    Ingreso.Left = (Screen.Width / 2) - (Me.Width / 2)

End Sub

Public Sub txtDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtDocumento.Text = "9999" Then
            zMain.Show
            Unload Me
            Exit Sub
        End If
        
        If txtDocumento.Text = "" Then
            Exit Sub
        End If
        
        If txtDocumento.Text = "99" Then
            Ticket.consFinal = True
            Ticket.idEmpleado = "9999"
            Ticket.Show
            Exit Sub
        End If
        
        Set rsconsulta = New ADODB.Recordset
        SQL = "SELECT * FROM empleados WHERE numerodocumento =" & txtDocumento.Text
        rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        If rsconsulta.BOF And rsconsulta.EOF Then
            Call MsgBox("El documento ingresado no existe en la base de datos.", vbCritical, "No existe documento")
            txtDocumento.Text = ""
            txtDocumento.SetFocus
            Exit Sub
        End If
        
        Ticket.idEmpleado = rsconsulta!id
        Ticket.consFinal = False
        Ticket.Show
    Else
        SoloEnteros txtDocumento, KeyAscii
    End If
        
End Sub

Sub Recargar()

    Unload Me
    Ingreso.Show

End Sub

Sub backup()

    Dim Database As String
    Dim Server As String
    Dim File As String
    
    If inicio = False Then
        Database = Data.Properties(0)
        Server = Split(Data.Properties(52), " ")(0)
        File = "C:\SISTEMA\BACKUP\" & Database & Format(Now, "yyyymmddhhnnss") & ".sql"
        
        'Ejecuta el comando MySQLDump para generar el backup
        Set comando1 = CreateObject("WSCript.shell")
        comando1.Run "cmd /K C: & CD C:\SISTEMA\BACKUP & mysqldump.exe -u root -h " & Server & " --port=3306 --password=cmsis00 " & Database & " --opt -c > " & File & " & exit"
        Set comando1 = Nothing
    End If

    inicio = True

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
