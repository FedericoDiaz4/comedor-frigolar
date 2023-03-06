VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compra"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
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
   ScaleHeight     =   6135
   ScaleWidth      =   10455
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
      Left            =   9480
      Picture         =   "Compra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   " Salir "
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtNumero 
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
      Left            =   4200
      TabIndex        =   11
      Top             =   1080
      Width           =   1935
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
      Height          =   345
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   6000
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
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   1920
   End
   Begin VB.ComboBox cboLetra 
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "Compra.frx":058A
      Left            =   2160
      List            =   "Compra.frx":0597
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox cboTipo 
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "Compra.frx":05A4
      Left            =   120
      List            =   "Compra.frx":05B1
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   8520
      Picture         =   "Compra.frx":05D9
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   " Guardar"
      Top             =   5400
      Width           =   855
   End
   Begin VB.ComboBox cboProveedor 
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "Compra.frx":0B63
      Left            =   4200
      List            =   "Compra.frx":0B65
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   3975
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
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1920
   End
   Begin VB.TextBox txtImporte 
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
      Left            =   4200
      TabIndex        =   21
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtPrecio 
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
      Left            =   2160
      TabIndex        =   20
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtCantidad 
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
      TabIndex        =   19
      Top             =   2640
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   2895
      Left            =   120
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12640511
      ForeColorFixed  =   -2147483640
      GridColorFixed  =   -2147483630
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
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   " Seleccionar Fecha"
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   62849025
      CurrentDate     =   40544
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   8400
      TabIndex        =   22
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtBruto 
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
         Enabled         =   0   'False
         Height          =   345
         Left            =   0
         TabIndex        =   24
         Text            =   "0,00"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtRecargo 
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
         Left            =   0
         TabIndex        =   26
         Text            =   "0,00"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtDescuento 
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
         Left            =   0
         TabIndex        =   28
         Text            =   "0,00"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtIIBB 
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
         Left            =   0
         TabIndex        =   30
         Text            =   "0,00"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtOtros 
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
         Left            =   0
         TabIndex        =   32
         Text            =   "0,00"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtIVA 
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
         Left            =   0
         TabIndex        =   34
         Text            =   "0,00"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox txtTotal 
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
         Enabled         =   0   'False
         Height          =   345
         Left            =   0
         TabIndex        =   36
         Text            =   "0,00"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Bruto"
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Recargo"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Descuento"
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Otros Impuestos"
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "IIBB"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "IVA"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Total"
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
         Left            =   0
         TabIndex        =   35
         Top             =   4320
         Width           =   1095
      End
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   8280
      X2              =   8280
      Y1              =   0
      Y2              =   6120
   End
   Begin VB.Label Label12 
      Caption         =   "Código Proveedor"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Letra"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Artículo"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Tipo Comprobante"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   0
      X2              =   8280
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Pto Venta - Número"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Código"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Importe"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Precio"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "Compra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ID As Integer
Dim idArt As Single

Private Sub cboLetra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If

End Sub

Private Sub cboProveedor_Click()
    
    If cboProveedor.ListIndex <> -1 Then
        txtCodProveedor.Text = getData(cboProveedor.ItemData(cboProveedor.ListIndex), "codigo", "proveedores")
    End If
    
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If

End Sub

Private Sub txtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = vbKeyF1 Then
        ArticulosList.Entrada = "COMPRA"
        ArticulosList.Height = 5520
        ArticulosList.Show
    End If
    
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cboProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmdGuardar_Click()
    
    If txtBruto.Text = "0,00" Then Exit Sub
    
    'Guardar
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM compras"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Recordset.AddNew
    
    Set rsIndice = New ADODB.Recordset
    SQL = "SELECT compra FROM indices"
    rsIndice.Open SQL, Data, adOpenKeyset, adLockOptimistic
    ID = CInt(rsIndice!Compra) + 1
    rsIndice!Compra = ID
    rsIndice.Update
    rsIndice.Close
    
    Recordset!ID = ID
    Recordset!Tipo = cboTipo.Text
    Recordset!letra = cboLetra.Text
    Recordset!Numero = txtNumero.Text
    Recordset!idProveedor = cboProveedor.ItemData(cboProveedor.ListIndex)
    Recordset!fecha = DTFecha.Value
    Recordset!bruto = Format(txtBruto.Text, "0.00")
    Recordset!recargo = Format(txtRecargo.Text, "0.00")
    Recordset!descuento = Format(txtDescuento.Text, "0.00")
    Recordset!iibb = Format(txtIIBB.Text, "0.00")
    Recordset!otros = Format(txtOtros.Text, "0.00")
    Recordset!IVA = Format(txtIVA.Text, "0.00")
    Recordset!Total = Format(txtTotal.Text, "0.00")
    Recordset!Saldo = Format(txtTotal.Text, "0.00")
    Recordset!Estado = "CTACTE"
    Recordset!DateTime = DTFecha.Value & " " & Time
    Recordset.Update
    Recordset.Close
    
    'Establece el n de entrada
    Set rsEntradaD = New ADODB.Recordset
    SQL = "UPDATE comprasd SET idfac = '" & ID & "' WHERE idfac = '0'"
    rsEntradaD.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Flex_Click()
    
    Dim i As Integer
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then
        Exit Sub
    End If
    
    i = Flex.TextMatrix(Flex.Row, 0)
    
    Select Case MsgBox("¿Desea eliminar el item " & Flex.TextMatrix(Flex.Row, 3) & " de la compra?", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)
        Case vbNo: Exit Sub
    End Select
    
    Set rsDelete = New ADODB.Recordset
    SQL = "DELETE FROM comprasd WHERE id = '" & i & "'"
    rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    DTFecha.Value = Date
    
    CargaCombo "proveedores", "nombre", "nombre", cboProveedor
    
    'Borra los datos cargados en una compras no guardada
    Set rsDel = New ADODB.Recordset
    SQL = "DELETE FROM comprasd WHERE idfac = '0'"
    rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Sub Cargar()
    
    txtBruto.Text = "0,00"
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM comprasd WHERE idfac = 0 ORDER BY id"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        Set Flex.DataSource = Recordset
        Do While Not Recordset.EOF
            txtBruto.Text = Format(CDec(txtBruto.Text) + CDec(Recordset!importe), "0.00")
            Recordset.MoveNext
        Loop
    Else
        Flex.Clear
        Flex.Rows = 2
        txtBruto.Text = "0,00"
    End If
    Recordset.Close
    
    OrdenaFlex
    
End Sub

Sub OrdenaFlex()
    
    Flex.FormatString = "id|idfac|idart|Artículo|Cantidad|Precio|Importe"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 0
    Flex.ColWidth(2) = 0
    Flex.ColWidth(3) = 3800
    Flex.ColWidth(4) = 1300
    Flex.ColWidth(5) = 1300
    Flex.ColWidth(6) = 1300
    Flex.ColWidth(7) = 1300
    Flex.ColAlignment(3) = 1
    
End Sub

Sub Add()
    
    If txtCodArticulo.Text = "" Then
        txtCodArticulo.SetFocus
        Exit Sub
    ElseIf txtCantidad.Text = "" Then
        txtCantidad.SetFocus
        Exit Sub
    ElseIf txtPrecio.Text = "" Then
        txtPrecio.SetFocus
        Exit Sub
    ElseIf txtImporte.Text = "" Then
        txtImporte.SetFocus
        Exit Sub
    End If
    
    'Guardar
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM comprasd"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Recordset.AddNew
    Recordset!idFac = 0
    Recordset!idArt = idArt
    Recordset!art = txtArticulo.Text
    Recordset!Cantidad = txtCantidad.Text
    Recordset!precio = Format(txtPrecio.Text, "0.00")
    Recordset!importe = Format(txtImporte.Text, "0.00")
    Recordset.Update
    Recordset.Close
    
    'Mostrar
    Cargar
    
    'Vaciar Box
    txtCodArticulo.Text = ""
    txtArticulo.Text = ""
    txtCantidad.Text = ""
    txtPrecio.Text = ""
    txtImporte.Text = ""
    txtCodArticulo.SetFocus
    
End Sub

Sub CalcTotal()
    
    If txtBruto.Text = "" Or txtBruto.Text = "0,00" Then Exit Sub
    If txtDescuento.Text = "" Then txtDescuento.Text = "0,00"
    If txtRecargo.Text = "" Then txtRecargo.Text = "0,00"
    If txtIIBB.Text = "" Then txtIIBB.Text = "0,00"
    If txtOtros.Text = "" Then txtOtros.Text = "0,00"
    If txtIVA.Text = "" Then txtIVA.Text = "0,00"
    
    txtTotal.Text = Format(CDec(txtBruto.Text) - CDec(txtDescuento.Text) + CDec(txtRecargo.Text) + CDec(txtIIBB) + CDec(txtOtros) + CDec(txtIVA), "0.00")
    
End Sub

Private Sub txtBruto_Change()
    
    CalcTotal
    
End Sub

Private Sub txtCantidad_Change()
    
    If txtPrecio.Text = "" Or txtCantidad.Text = "" Then Exit Sub
    
    txtImporte.Text = Format(CDec(txtCantidad.Text) * CDec(txtPrecio.Text), "0.00")
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtCantidad, KeyAscii
    End If
    
End Sub

Private Sub txtCodArticulo_Change()
    
    If cboProveedor.ListIndex = -1 Then Exit Sub
    If Len(txtCodArticulo.Text) <> 6 Then Exit Sub
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM articulos WHERE codigo = '" & txtCodArticulo.Text & "' AND eliminado <> 1 LIMIT 1"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        idArt = Recordset!ID
        txtArticulo.Text = Recordset!nombre
        Recordset.Close
        
        Set Recordset = New ADODB.Recordset
        SQL = "SELECT * FROM articulosc WHERE idart = " & idArt & " AND idpro = " & cboProveedor.ItemData(cboProveedor.ListIndex)
        Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not Recordset.BOF And Not Recordset.EOF Then
            txtPrecio.Text = Format(Recordset!precio, "0.00")
        End If
        Recordset.Close
    Else
        txtArticulo.Text = ""
    End If
    
End Sub

Private Sub txtCodArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        ArticulosList.Entrada = "COMPRA"
        ArticulosList.Height = 5520
        ArticulosList.Show
    End If
    
End Sub

Private Sub txtCodArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtCodProveedor_Change()
    
    If txtCodProveedor.Text = "" Then
        cboProveedor.ListIndex = -1
        Exit Sub
    End If
    
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
    
    txtCodProveedor.SelStart = Len(txtCodProveedor.Text)
    
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDescuento_Change()
    
    CalcTotal
    
End Sub

Private Sub txtDescuento_GotFocus()
    
    txtDescuento.SelStart = 0
    txtDescuento.SelLength = Len(txtDescuento.Text)
    
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtDescuento, KeyAscii
    End If
    
End Sub

Private Sub txtIIBB_Change()
    
    CalcTotal

End Sub

Private Sub txtIIBB_GotFocus()
    
    txtIIBB.SelStart = 0
    txtIIBB.SelLength = Len(txtIIBB.Text)
    
End Sub

Private Sub txtIIBB_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtIIBB, KeyAscii
    End If
    
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Add
        KeyAscii = 0
    Else
        CambiaPunto txtImporte, KeyAscii
    End If
    
End Sub

Private Sub txtIVA_Change()
    
    CalcTotal

End Sub

Private Sub txtIVA_GotFocus()
    
    txtIVA.SelStart = 0
    txtIVA.SelLength = Len(txtIVA.Text)
    
End Sub

Private Sub txtIVA_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtIVA, KeyAscii
    End If
    
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtNumero_LostFocus()
    
    If txtNumero = "" Then
        Exit Sub
    End If
    
    i = InStr(1, txtNumero.Text, "-")
    
    If i = 0 Then
        txtNumero.Text = Format(txtNumero, "00000000")
    Else
        txtNumero.Text = Format(Mid(txtNumero, 1, i - 1), "0000") & "-" & Format(Mid(txtNumero, i + 1, Len(txtNumero)), "00000000")
    End If
    
End Sub

Private Sub txtOtros_Change()
    
    CalcTotal

End Sub

Private Sub txtOtros_GotFocus()
    
    txtOtros.SelStart = 0
    txtOtros.SelLength = Len(txtOtros.Text)
    
End Sub

Private Sub txtOtros_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtOtros, KeyAscii
    End If

End Sub

Private Sub txtPrecio_Change()
    
    If txtPrecio.Text = "" Or txtCantidad.Text = "" Then Exit Sub
    
    txtImporte.Text = Format(CDec(txtCantidad.Text) * CDec(txtPrecio.Text), "0.00")
    
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtPrecio, KeyAscii
    End If
    
End Sub

Private Sub txtBruto_GotFocus()
    
    txtBruto.SelLength = Len(txtBruto.Text)
    
End Sub

Private Sub txtBruto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtBruto, KeyAscii
    End If
    
End Sub

Private Sub txtRecargo_Change()
    
    CalcTotal
    
End Sub

Private Sub txtRecargo_GotFocus()
    
    txtRecargo.SelStart = 0
    txtRecargo.SelLength = Len(txtRecargo.Text)
    
End Sub

Private Sub txtRecargo_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtRecargo, KeyAscii
    End If
    
End Sub
