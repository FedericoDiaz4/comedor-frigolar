VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Ticket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Empleados"
   ClientHeight    =   10275
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   12300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   12300
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNumero 
      Enabled         =   0   'False
      Height          =   330
      Left            =   5520
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame frmTotales 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   9720
      TabIndex        =   13
      Top             =   8760
      Width           =   2415
      Begin VB.TextBox txtCantidad 
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtTotal 
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Total"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frmAnterior 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   9480
      Width           =   2895
      Begin VB.TextBox txtCantAnterior 
         Height          =   330
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtImporteAnterior 
         Height          =   330
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label 
         Caption         =   "Cant Anterior"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Anterior"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox picBotones 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   10080
      ScaleHeight     =   615
      ScaleWidth      =   2055
      TabIndex        =   4
      Top             =   9600
      Width           =   2055
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
         Left            =   120
         Picture         =   "RetiroList.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Guardar"
         Top             =   0
         Width           =   855
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
         Left            =   1080
         Picture         =   "RetiroList.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   " Salir "
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Frame frmBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtBuscar 
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
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   250
         Width           =   3495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexTicket 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   3413
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexMenu 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9551
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
   Begin VB.Label Label10 
      Caption         =   "Numero de Ticket"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   12240
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   12240
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label lblPersona 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   480
      Width           =   8055
   End
End
Attribute VB_Name = "Ticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public consFinal As Boolean
Public idEmpleado As String
Dim idArt As String
Dim nombreArt As String
Dim codigoArt As String
Public cantidadArt As String
Dim precioArt As String
Dim cantAnterior As Integer
Dim importeAnterior As Double
Dim id As Double

Private Sub cboBuscar_Click()
    
    On Error Resume Next
    txtBuscar.Text = ""
    txtBuscar.SetFocus
    
End Sub

Private Sub cmdGuardar_Click()

    If idEmpleado <> 9999 Then
        txtCantAnterior.Text = Val(txtCantAnterior.Text) + Val(txtCantidad.Text)
        txtImporteAnterior.Text = Val(txtImporteAnterior.Text) + Val(txtTotal.Text)
    End If

    If FlexTicket.TextMatrix(1, 0) = "" Then
        Call MsgBox("No puede guardar un ticket vacio.", vbCritical Or vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Set rsGuardar = New ADODB.Recordset
    SQL = "SELECT * FROM comidas WHERE id = 0"
    rsGuardar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    rsGuardar.AddNew
    rsGuardar!fecha = Format(Now, "yyyy-mm-dd hh:mm:ss")
    rsGuardar!numero = txtNumero.Text
    rsGuardar!idEmpleado = idEmpleado
    rsGuardar!Precio = Format(txtTotal.Text, "0.00")
    rsGuardar!eliminado = 0
    
    rsGuardar.Update
    id = rsGuardar!id
    rsGuardar.Close
    
    'Actualizo detalle de ticket con ID obtenido
    Set rsDetalle = New ADODB.Recordset
    SQL = "UPDATE comidasd SET idcomida = " & id & " WHERE idcomida = 0"
    rsDetalle.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    imprimeTicket
    Unload Me
    Ingreso.Show
    Ingreso.Recargar

End Sub

Private Sub imprimeTicket()

    Printer.FontName = "Consolas"
    Printer.FontBold = True
    Printer.FontSize = 16
    Printer.Print Centrar("Bs As Cheff", 11)
    
    Set rsTicket = New ADODB.Recordset
    SQL = "SELECT date(t.fecha) fecha, time(t.fecha) hora, t.precio FROM comidas as t WHERE id = " & id
    rsTicket.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Printer.FontSize = 8
    Printer.FontName = "Courier New"
    Printer.Print Centrar(Format(rsTicket!hora, "dd/mm/yyyy HH:MM:SS"), 19)
    
    Printer.FontBold = False
    Printer.FontName = "Calibri"
    Printer.Print ""
    
    rsTicket.Close

    Printer.FontSize = 8
    Printer.FontName = "Courier New"
    Printer.FontBold = True
    Printer.Print "Ticket: " & txtNumero.Text
    
    Printer.FontBold = False
    Printer.FontName = "Calibri"
    Printer.Print ""
    
    Printer.FontSize = 8
    Printer.FontBold = True
    Printer.FontName = "Courier New"
    Printer.Print "Legajo: "
    
    Printer.FontSize = 7
    Printer.FontName = "Courier New"
    If idEmpleado = 9999 Then
        Printer.Print "(9999 - Consumidor Final)"
    Else
        Printer.Print Split(lblPersona.Caption, "-")(0)
    End If

    Printer.FontSize = 8
    Printer.FontName = "Courier New"
    Printer.Print ""
    
    Printer.FontSize = 7
    Printer.FontBold = True
    Printer.FontName = "Courier New"
    Printer.Print "Producto           Cant    Total"
    
    For Fila = 1 To FlexTicket.Rows - 1
        Printer.FontSize = 7
        Printer.FontName = "Courier New"
        Printer.Print Detalle(FlexTicket.TextMatrix(Fila, 4), FlexTicket.TextMatrix(Fila, 6), FlexTicket.TextMatrix(Fila, 7))
    Next
    
    Printer.FontSize = 7
    Printer.FontName = "Courier New"
    Printer.Print "Totales:             " & Format(txtCantidad.Text, "0") & " " & Format(txtTotal.Text, "0.00")
    
    Printer.FontSize = 8
    Printer.FontName = "Courier New"
    Printer.Print ""
    
    If idEmpleado <> 9999 Then
        Printer.FontSize = 7
        Printer.FontName = "Courier New"
        Printer.Print "Cantidad en Periodo:    " & Format(txtCantAnterior.Text, "0")
        Printer.Print "Total en Periodo:       " & Format(txtImporteAnterior.Text, "0.00")
    End If
    
    Printer.EndDoc

End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    Ingreso.Show
    Ingreso.Recargar
    
End Sub

Private Sub FlexMenu_DblClick()

    If FlexMenu.TextMatrix(FlexMenu.Row, 0) = "" Then
        Exit Sub
    End If

    Cantidad.Show vbModal

    If cantidadArt <> 0 Then
        idArt = FlexMenu.TextMatrix(FlexMenu.Row, 0)
        codigoArt = FlexMenu.TextMatrix(FlexMenu.Row, 1)
        nombreArt = FlexMenu.TextMatrix(FlexMenu.Row, 2)
        precioArt = FlexMenu.TextMatrix(FlexMenu.Row, 3)
    Else
        txtBuscar.SetFocus
        Exit Sub
    End If
    
    Add
    txtBuscar.Text = ""

End Sub

Private Sub FlexMenu_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        FlexMenu_DblClick
    Else
        txtBuscar.SetFocus
    End If

End Sub

Private Sub FlexTicket_DblClick()

    If FlexTicket.TextMatrix(FlexTicket.Row, 0) = "" Then
        Exit Sub
    End If
    
    Select Case MsgBox("¿Desea eliminar la comida seleccionada?", vbYesNo Or vbInformation Or vbDefaultButton1, "Desea Eliminar")
        Case vbYes
            If FlexTicket.TextMatrix(FlexTicket.Row, 6) > 1 Then
                Set rsDelete = New ADODB.Recordset
                SQL = "SELECT id, cantidad, precio, total FROM comidasd WHERE id = " & FlexTicket.TextMatrix(FlexTicket.Row, 0)
                rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
                
                rsDelete!Cantidad = CInt(rsDelete!Cantidad) - 1
                rsDelete!total = Format(CDbl(rsDelete!total) - CInt(rsDelete!Precio), "0.00")
                
                rsDelete.Update
                rsDelete.Close
            Else
                Set rsDelete = New ADODB.Recordset
                SQL = "DELETE FROM comidasd WHERE id = " & FlexTicket.TextMatrix(FlexTicket.Row, 0)
                rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
            End If
    
    End Select

    cargarFlexTicket
    cargoTxt

End Sub

Private Sub Form_Load()
    
    txtTotal.Text = Format(0, "0.00")
    txtCantidad.Text = 0
    cantAnterior = 0
    importeAnterior = 0
    
    'Borro detalle de tickets sin ticket asignado
    Set rsDelete = New ADODB.Recordset
    SQL = "DELETE FROM comidasd WHERE idcomida = 0"
    rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Cargo Numero de Ticket
    Set rsNumero = New ADODB.Recordset
    SQL = "SELECT numero FROM comidas ORDER BY numero DESC LIMIT 1"
    rsNumero.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsNumero.BOF And Not rsNumero.EOF Then
        txtNumero.Text = Format(rsNumero!numero + 1, "00000000")
    Else
        txtNumero.Text = Format("1", "00000000")
    End If
    rsNumero.Close
    
    If Day(Now) < 16 Then
        desde = Format(DateSerial(Year(Now), Month(Now), 1), "yyyy-mm-dd")
        hasta = Format(DateSerial(Year(Now), Month(Now), 15), "yyyy-mm-dd")
    Else
        desde = Format(DateSerial(Year(Now), Month(Now), 16), "yyyy-mm-dd")
        hasta = Format(DateSerial(Year(Now), Month(Now) + 1, 0), "yyyy-mm-dd")
    End If
    
    'Cargo Datos de la persona
    If idEmpleado <> 9999 Then
        Set rsPersona = New ADODB.Recordset
        SQL = "SELECT * FROM empleados WHERE id = " & idEmpleado
        rsPersona.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        lblPersona.Caption = rsPersona!nombre & " - " & rsPersona!numerodocumento & " - " & rsPersona!nrolegajo
        rsPersona.Close
        
    End If
    'Cargo cantidad e importe acumulado en el mes.
    If idEmpleado <> 9999 Then
        Set Recordset = New ADODB.Recordset
        SQL = "SELECT * FROM comidas WHERE idempleado = " & idEmpleado & " "
        SQL = SQL & "AND fecha BETWEEN '" & desde & "' AND '" & hasta & "' "
        Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not Recordset.EOF And Not Recordset.BOF Then
            Do While Not Recordset.EOF
                idcomida = Recordset!id
                Set rsContador = New ADODB.Recordset
                SQL = "SELECT cantidad, total FROM comidasd WHERE idcomida = " & idcomida & " "
                rsContador.Open SQL, Data, adOpenKeyset, adLockOptimistic
                If Not rsContador.EOF And Not rsContador.BOF Then
                    Do While Not rsContador.EOF
                        cantAnterior = cantAnterior + rsContador!Cantidad
                        importeAnterior = importeAnterior + rsContador!total
                        rsContador.MoveNext
                    Loop
                End If
                Recordset.MoveNext
            Loop
        End If
        
        txtCantAnterior.Text = cantAnterior
        txtImporteAnterior.Text = Format(importeAnterior, "0.00")
    End If
    
    initForm Me
    CargarMenu
    
End Sub

Sub ordenaFlexMenu()
    
    FlexMenu.FormatString = "id|Código|Nombre|Precio"
    FlexMenu.ColWidth(0) = 0
    FlexMenu.ColWidth(1) = 1000
    FlexMenu.ColWidth(2) = 4000
    FlexMenu.ColWidth(3) = 1250
    
End Sub

Sub CargarMenu()
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT id, codigo, nombre, precio FROM menus WHERE eliminado = 0 "
    If txtBuscar.Text <> "" Then
        SQL = SQL & "AND (nombre LIKE '%" & txtBuscar.Text & "%' OR codigo LIKE '" & txtBuscar.Text & "%') "
    End If
    SQL = SQL & "ORDER BY nombre"
    Clipboard.Clear
    Clipboard.SetText SQL
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Not Recordset.BOF And Not Recordset.EOF Then
        Set FlexMenu.DataSource = Recordset
    Else
        FlexMenu.Clear
        FlexMenu.Rows = 2
    End If
    Recordset.Close
    
    ordenaFlexMenu
    
End Sub

Private Sub Form_Resize()
    
    If Me.ScaleHeight < 2000 Or Me.ScaleWidth < 2000 Then
        Exit Sub
    End If

    Const Margen = 120

    FlexMenu.Width = Me.ScaleWidth - 2 * Margen
    FlexMenu.Height = FlexMenu.Height + 1000
    
    Line1.Y1 = FlexMenu.Top + FlexMenu.Height + 100
    Line1.Y2 = FlexMenu.Top + FlexMenu.Height + 100
    Line1.X2 = Me.ScaleWidth - Margen
    
    Line2.Y1 = Line1.Y1 + 100
    Line2.Y2 = Line1.Y2 + 100
    Line2.X2 = Me.ScaleWidth - Margen
    
    FlexTicket.Width = Me.ScaleWidth - 2 * Margen
    FlexTicket.Top = Line2.Y1 + 100

    frmTotales.Top = FlexTicket.Top + FlexTicket.Height + 50
    frmTotales.Left = Me.ScaleWidth - frmTotales.Width - Margen

    picBotones.Left = Me.ScaleWidth - picBotones.Width - Margen
    picBotones.Top = frmTotales.Height + frmTotales.Top + 50

    frmAnterior.Top = FlexTicket.Top + FlexTicket.Height + 600

End Sub

Private Sub txtBuscar_Change()
    
    CargarMenu
    
End Sub

Private Sub Add()

    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM comidasd WHERE idArt = '" & idArt & "' AND idcomida = 0"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Recordset.EOF And Recordset.BOF Then
        Recordset.AddNew
        
        Recordset!idcomida = 0
        Recordset!idArt = idArt
        Recordset!codigo = codigoArt
        Recordset!nombre = nombreArt
        Recordset!Precio = Format(precioArt, "0.00")
        Recordset!Cantidad = cantidadArt
        Recordset!total = Format(precioArt * cantidadArt, "0.00")
    Else
        Recordset!Cantidad = CInt(Recordset!Cantidad) + cantidadArt
        Recordset!total = Format(precioArt * CInt(Recordset!Cantidad), "0.00")
    End If
    Recordset.Update
    Recordset.Close
    
    cargarFlexTicket
    
End Sub

Sub cargarFlexTicket()

    Set Recordset = New ADODB.Recordset
    SQL = "SELECT id, idcomida, idart, codigo, nombre, precio, cantidad, total FROM comidasd WHERE idcomida = 0"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Recordset.BOF And Recordset.EOF Then
        FlexTicket.Clear
        FlexTicket.Rows = 2
    Else
        Set FlexTicket.DataSource = Recordset
    End If
    Recordset.Close

    ordenaFlexTicket
    cargoTxt
    
    txtBuscar.SetFocus

End Sub

Private Sub cargoTxt()

    Dim Cantidad As Integer
    Dim total As Double
    
    Cantidad = 0
    total = 0

    If FlexTicket.TextMatrix(1, 0) = "" Then
        Exit Sub
    End If
    
    For Fila = 1 To FlexTicket.Rows - 1
        Cantidad = Cantidad + FlexTicket.TextMatrix(Fila, 6)
        total = total + FlexTicket.TextMatrix(Fila, 7)
    Next
    
    txtCantidad.Text = Format(Cantidad, "0.00")
    txtTotal.Text = Format(total, "$ 0.00")
    
End Sub

Sub ordenaFlexTicket()

    FlexTicket.FormatString = "id|idcomida|idart|Código|Nombre|Precio|Cantidad|Total"
    FlexTicket.ColWidth(0) = 0
    FlexTicket.ColWidth(1) = 0
    FlexTicket.ColWidth(2) = 0
    FlexTicket.ColWidth(3) = 1000
    FlexTicket.ColWidth(4) = 4000
    FlexTicket.ColWidth(5) = 1250
    FlexTicket.ColWidth(6) = 1000
    FlexTicket.ColWidth(7) = 1250

End Sub

' Instala el hook en el MSFlexGrid
Private Sub FlexTicket_GotFocus()
    HookForm FlexTicket
End Sub

' elimina el hook
Private Sub FlexTicket_LostFocus()
    UnHookForm FlexTicket
End Sub

' Instala el hook en el MSFlexGrid
Private Sub FlexMenu_GotFocus()
    HookForm FlexMenu
End Sub

' elimina el hook
Private Sub FlexMenu_LostFocus()
    UnHookForm FlexMenu
End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 40 Then
        FlexMenu.SetFocus
    End If

End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        FlexMenu.SetFocus
        FlexMenu_DblClick
    Else
        SoloEnteros txtBuscar, KeyAscii
    End If

End Sub
