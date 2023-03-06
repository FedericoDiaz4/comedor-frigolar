VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Recibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   8.25
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
   ScaleHeight     =   7095
   ScaleWidth      =   11160
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   10200
      Picture         =   "Recibos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   " Salir "
      Top             =   6360
      Width           =   855
   End
   Begin VB.ListBox cboNomCli 
      Height          =   1635
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox txtNomCli 
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
      Left            =   4440
      TabIndex        =   6
      Top             =   360
      Width           =   5040
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   9240
      Picture         =   "Recibos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   " Guardar"
      Top             =   6360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   " Imputaciones "
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   10935
      Begin VB.TextBox txtSaldoFac 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdOkImp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10200
         Picture         =   "Recibos.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2620
         TabIndex        =   18
         Top             =   505
         Width           =   250
      End
      Begin VB.TextBox txtTipoFac 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTotalFac 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtAbonoFac 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8760
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtFechaFac 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTotalImp 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   360
         Left            =   9600
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtDescFac 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5880
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTotalDes 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   360
         Left            =   6000
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "0,00"
         Top             =   2160
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexImp 
         Height          =   1215
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   870
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2143
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
      Begin VB.TextBox txtNumFac 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Saldo"
         Height          =   255
         Left            =   7320
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   8880
         TabIndex        =   28
         Top             =   2205
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Nº"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
         Height          =   255
         Left            =   4440
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Abonó"
         Height          =   255
         Left            =   8760
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Descuento %"
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descuentos:"
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
         Left            =   4770
         TabIndex        =   26
         Top             =   2205
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Pagos "
      Height          =   2655
      Left            =   120
      TabIndex        =   30
      Top             =   3600
      Width           =   10935
      Begin VB.CommandButton cmdOkPag 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10200
         Picture         =   "Recibos.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtDetalle 
         Height          =   315
         Left            =   1560
         TabIndex        =   36
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtImportePago 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8760
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTotalPag 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   9600
         TabIndex        =   42
         Text            =   "0,00"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox cboTipo 
         Height          =   345
         ItemData        =   "Recibos.frx":1628
         Left            =   120
         List            =   "Recibos.frx":163B
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   480
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtFechaPago 
         Height          =   315
         Left            =   7320
         TabIndex        =   37
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexPag 
         Height          =   1215
         Left            =   120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   870
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2143
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   9000
         TabIndex        =   41
         Top             =   2205
         Width           =   585
      End
      Begin VB.Label Label14 
         Caption         =   "Importe"
         Height          =   255
         Left            =   8760
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   7320
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Detalle"
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtCodCli 
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
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   345
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   95223809
      CurrentDate     =   39589
   End
   Begin VB.Label lblNrecibo 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00000000"
      Height          =   345
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblRecibo 
      Caption         =   "Recibo nº"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Recibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As String) As Long
Private Const LB_SELECTSTRING = &H18C

Public idRecibo As Integer
Public Modificar As Boolean
Public idCheque As Integer
Public idFac As Integer
Dim DescPesos As Single

Sub cargaFlexPag()
    
    Set rsPag = New ADODB.Recordset
    SQL = "SELECT * FROM recibosp WHERE id = '" & idRecibo & "'"
    rsPag.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsPag.BOF And Not rsPag.EOF Then
        Set MSHFlexPag.DataSource = rsPag
    Else
        MSHFlexPag.Clear
        MSHFlexPag.Rows = 2
    End If
    rsPag.Close
    
    OrdenaFlexPag
    
End Sub

Sub OrdenaFlexImp()
    
    MSHFlexImp.FormatString = "id|idimp|idfactura|Tipo|Nº Factura|Fecha|Total|descporc|Descuento|Saldo|Abonado"
    MSHFlexImp.ColWidth(0) = 0
    MSHFlexImp.ColWidth(1) = 0
    MSHFlexImp.ColWidth(2) = 0
    MSHFlexImp.ColWidth(3) = 1440
    MSHFlexImp.ColWidth(4) = 1440
    MSHFlexImp.ColWidth(5) = 1440
    MSHFlexImp.ColWidth(6) = 1440
    MSHFlexImp.ColWidth(7) = 0
    MSHFlexImp.ColWidth(8) = 1440
    MSHFlexImp.ColWidth(9) = 1440
    MSHFlexImp.ColWidth(10) = 1440
    
End Sub

Sub OrdenaFlexPag()
    
    MSHFlexPag.FormatString = "id|idpag|Tipo|Detalle|Fecha|Importe|"
    MSHFlexPag.ColWidth(0) = 0
    MSHFlexPag.ColWidth(1) = 0
    MSHFlexPag.ColWidth(2) = 1440
    MSHFlexPag.ColWidth(3) = 5760
    MSHFlexPag.ColWidth(4) = 1440
    MSHFlexPag.ColWidth(5) = 1440
    MSHFlexPag.ColWidth(6) = 0
    
End Sub

Sub calculaTotal()
    
    Dim totalImp, totalPag As Single
    
    totalImp = 0
    totalPag = 0
    totalDes = 0
    
    'Total Imputaciones:
    Set rsTotImp = New ADODB.Recordset
    SQL = "SELECT id, descpesos, abonado FROM recibosi WHERE id = " & idRecibo
    rsTotImp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsTotImp.BOF And Not rsTotImp.EOF Then
        Do While Not rsTotImp.EOF
            totalDes = totalDes + CDec(Format(rsTotImp!DescPesos, "0.00"))
            totalImp = totalImp + CDec(Format(rsTotImp!Abonado, "0.00"))
            rsTotImp.MoveNext
        Loop
    End If
    txtTotalDes.Text = Format(totalDes, "0.00")
    txtTotalImp.Text = Format(totalImp, "0.00")
    
    'Total recibos:
    Set rsTotPag = New ADODB.Recordset
    SQL = "SELECT id, importe FROM recibosp WHERE id = " & idRecibo
    rsTotPag.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsTotPag.BOF And Not rsTotPag.EOF Then
        Do While Not rsTotPag.EOF
            totalPag = totalPag + CDec(Format(rsTotPag!importe, "0.00"))
            rsTotPag.MoveNext
        Loop
    End If
    txtTotalPag.Text = Format(totalPag, "0.00")
    
End Sub

Private Sub cboNomCli_Click()
    
    If cboNomCli.ListIndex = -1 Then
        Exit Sub
    End If
    
    Set rsPro = New ADODB.Recordset
    SQL = "SELECT codigo, nombre FROM clientes WHERE id = " & cboNomCli.ItemData(cboNomCli.ListIndex)
    rsPro.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsPro.BOF And Not rsPro.EOF Then
        txtCodCli.Text = rsPro!codigo
        txtNomCli.Text = rsPro!nombre
    Else
        txtCodCli.Text = ""
        txtNomCli.Text = ""
    End If
    rsPro.Close
    
    cboNomCli.Visible = False
    
End Sub

Private Sub cboTipo_GotFocus()
    
    txtFechaPago.Text = DTFecha.Value
    
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub cboTipo_LostFocus()
    
    If cboTipo.ListIndex <> 1 Then
        Exit Sub
    End If
    
    'Ingreso al Libro de Cheques
    Select Case MsgBox("¿Desea Ingresar el cheque al Libro de Cheques?", vbYesNo Or vbQuestion Or vbDefaultButton1, "Cheques - Fusca")
    Case vbYes
    
        Cheques.Entrada = "RECIBOS"
        Cheques.Nuevo = True
        Cheques.idCli = cboNomCli.ItemData(cboNomCli.ListIndex)
        Cheques.idRecibo = idRecibo
        Cheques.txtCliente.Text = cboNomCli.Text
        Cheques.DTIngreso.Value = DTFecha.Value
        Cheques.frmEgreso.Enabled = False
        Me.Hide
        Cheques.Show
        
    End Select
    
End Sub

Sub cmdAdd_Click()
    
    If cboNomCli.ListIndex = -1 Then
        Call MsgBox("Debe seleccionar un Cliente", vbExclamation, "Atención")
        cboNomCli.SetFocus
        Exit Sub
    End If
    
    VentaC.idCliente = cboNomCli.ItemData(cboNomCli.ListIndex)
    VentaC.lblCliente.Caption = cboNomCli.Text
    VentaC.idRecibo = idRecibo
    'Me.Hide
    VentaC.Show
    VentaC.Flex.SetFocus
    
End Sub

Private Sub cmdGuardar_Click()
    
    'Chequea que no queden campos vacíos
    If cboNomCli.ListIndex = -1 Then
        Call MsgBox("Debe seleccionar un Cliente.", vbExclamation, "Atención")
        cboNomCli.SetFocus
        Exit Sub
    End If
    
    'Abona las facturas cargadas
    Set rsImp = New ADODB.Recordset
    SQL = "SELECT * FROM recibosi where id = '" & idRecibo & "'"
    rsImp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not rsImp.EOF
        Set rsFac = New ADODB.Recordset
        SQL = "SELECT * FROM ventas WHERE id = '" & rsImp!idFactura & "'"
        rsFac.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not rsFac.BOF And Not rsFac.EOF Then
            rsFac!Saldo = Format(CDec(rsImp!Saldo), "0.00")
            If CDec(rsImp!Saldo) = 0 Then
                rsFac!Estado = "PAGA"
            End If
            rsFac.Update
        End If
        rsFac.Close
        rsImp.MoveNext
    Loop
    rsImp.Close
    
    'Guarda informacion en la base
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM recibosdr WHERE id = " & idRecibo
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Modificar = False Then
        Recordset.AddNew
        Recordset!ID = idRecibo
        Recordset!fecha = Format(DTFecha.Value, "dd/mm/yyyy")
        Recordset!idCliente = cboNomCli.ItemData(cboNomCli.ListIndex)
    End If
    Recordset!totalImp = Format(txtTotalImp.Text, "0.00")
    Recordset!totalPag = Format(txtTotalPag.Text, "0.00")
    Recordset!totalDes = Format(txtTotalDes.Text, "0.00")
    Recordset!Total = Format(CDec(txtTotalPag.Text) + CDec(txtTotalDes), "0.00")
    Recordset!DateTime = DTFecha.Value & " " & Time
    Recordset.Update
    Recordset.Close
    
    If Modificar = False Then
        'Aumenta +1 el nº de recibo
        Set dbOrden = New ADODB.Recordset
        SQL = "SELECT recibo FROM indices"
        dbOrden.Open SQL, Data, adOpenKeyset, adLockOptimistic
        dbOrden!recibo = idRecibo
        dbOrden.Update
        dbOrden.Close
    End If
    
    Unload Me
    
End Sub

Sub cmdOkImp_Click()
    
    If txtAbonoFac.Text = "" Then
        cmdAdd.SetFocus
        Exit Sub
    End If
    
    If txtNumFac.Text = "" Then
        cmdAdd.SetFocus
        Exit Sub
    End If
    
    If txtDescFac.Text = "" Then txtDescFac.Text = "0"
    
    'Si el abono es superior al saldo, escribe el saldo y sale
    If CDec(txtAbonoFac.Text) > CDec(txtSaldoFac.Text) Then
        txtAbonoFac.Text = txtTotalFac.Text
        txtAbonoFac.SetFocus
        Exit Sub
    End If
    
    'Guarda Imputación en la base
    Set rsImp = New ADODB.Recordset
    SQL = "SELECT * FROM recibosi where id = '" & idRecibo & "' AND idfactura = '" & txtTipoFac.Text & "'"
    rsImp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If rsImp.BOF And rsImp.EOF Then
        rsImp.AddNew
    End If
    rsImp!ID = idRecibo
    rsImp!idImp = i
    rsImp!idFactura = idFac
    rsImp!tfactura = txtTipoFac.Text
    rsImp!nfactura = txtNumFac.Text
    rsImp!fecha = txtFechaFac.Text
    rsImp!Total = Format(txtTotalFac.Text, "0.00")
    rsImp!descporc = Format(txtDescFac.Text, "0.00")
    rsImp!DescPesos = Format(DescPesos, "0.00")
    rsImp!Saldo = Format(CDec(txtSaldoFac.Text) - CDec(txtAbonoFac.Text), "0.00")
    rsImp!Abonado = Format(txtAbonoFac.Text, "0.00")
    rsImp.Update
    rsImp.Close
    
    idFac = 0
    txtTipoFac.Text = ""
    txtNumFac.Text = ""
    txtFechaFac.Text = ""
    txtTotalFac.Text = ""
    txtDescFac.Text = ""
    txtSaldoFac.Text = ""
    txtAbonoFac.Text = ""
    txtAbonoFac.Enabled = True
    
    cargaFlexImp
    calculaTotal
    
    cmdAdd.SetFocus
    
End Sub

Sub cargaFlexImp()
    
    Set rsImp = New ADODB.Recordset
    SQL = "SELECT * FROM recibosi WHERE id = '" & idRecibo & "'"
    rsImp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsImp.BOF And Not rsImp.EOF Then
        Set MSHFlexImp.DataSource = rsImp
    Else
        MSHFlexImp.Clear
        MSHFlexImp.Rows = 2
    End If
    rsImp.Close
    
    OrdenaFlexImp
    
End Sub

Public Sub cmdOkPag_Click()
    
    If txtImportePago.Text = "" Then
        txtImportePago.SetFocus
        Exit Sub
    End If
    
    If CDec(txtImportePago.Text) = 0 Then
        txtImportePago.SetFocus
        Exit Sub
    End If
    
    If cboTipo.ListIndex = -1 Then
        cboTipo.SetFocus
        Exit Sub
    End If
    
    'Guarda el pago en la base
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM recibosp where id = '" & idRecibo & "'"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Recordset.AddNew
    Recordset!ID = idRecibo
    Recordset!Tipo = cboTipo.Text
    Recordset!detalle = txtDetalle.Text
    Recordset!fecha = txtFechaPago.Text
    Recordset!importe = Format(txtImportePago.Text, "0.00")
    Recordset!idCheque = idCheque
    Recordset.Update
    Recordset.Close
    
    'Limpia los campos
    cboTipo.ListIndex = -1
    txtDetalle.Text = ""
    txtFechaPago.Text = DTFecha.Value
    txtImportePago.Text = ""
    idCheque = 0
    
    cargaFlexPag
    calculaTotal
    
    cboTipo.SetFocus
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub DTFecha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub Form_Load()
    
    'Posiciona el form
    Me.Top = 0
    Me.Left = (zMain.ScaleWidth / 2) - (Me.Width / 2)
    
    'Carga la Fecha de hoy
    DTFecha.Value = Format(Date, "dd/mm/yyyy")
    
    If Modificar = False Then
        
        'Busca el n de recibo en la tabla Orden
        Set Recordset = New ADODB.Recordset
        SQL = "SELECT * FROM indices"
        Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
        idRecibo = Recordset!recibo
        Recordset.Close
        idRecibo = idRecibo + 1
        
        'Borra las imputaciones, pagos, y cheques guardados en un recibo no guardado
        Set rsBorraImp = New ADODB.Recordset
        SQL = "DELETE FROM recibosi where id = '" & idRecibo & "'"
        rsBorraImp.Open SQL, Data, adOpenKeyset, adLockOptimistic
        Set rsBorraPag = New ADODB.Recordset
        SQL = "DELETE FROM recibosp where id = '" & idRecibo & "'"
        rsBorraPag.Open SQL, Data, adOpenKeyset, adLockOptimistic
        Set rsBorraChq = New ADODB.Recordset
        SQL = "DELETE FROM cheques where idrecibo = '" & idRecibo & "'"
        rsBorraChq.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
    End If
    
    'Escribe el número de Factura
    lblNrecibo = Format(idRecibo, "00000000")
    
    'Carga los clientes en los combos
    Set rsCli = New ADODB.Recordset
    SQL = "SELECT id, nombre FROM clientes WHERE eliminado = '0' ORDER BY nombre"
    rsCli.Open SQL, Data, adOpenKeyset, adLockOptimistic
    cboNomCli.Clear
    Do While Not rsCli.EOF
        cboNomCli.AddItem rsCli!nombre
        cboNomCli.ItemData(cboNomCli.NewIndex) = rsCli!ID
        rsCli.MoveNext
    Loop
    rsCli.Close
    
    
    cargaFlexImp
    cargaFlexPag
    calculaTotal
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    idRecibo = 0
    Modificar = False
    
End Sub

Private Sub MSHFlexImp_Click()
    
    Dim idImp As Integer
    
    idImp = MSHFlexImp.TextMatrix(MSHFlexImp.Row, 1)
    
    Select Case MsgBox("¿Está seguro que desea eliminar la imputación?", vbYesNo Or vbQuestion Or vbDefaultButton1, App.Title)
        Case vbNo: Exit Sub
    End Select
    
    'Elimina el registro de la base
    Set rsBorraImp = New ADODB.Recordset
    SQL = "DELETE FROM recibosi WHERE idimp = '" & idImp & "'"
    rsBorraImp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    cargaFlexImp
    calculaTotal
    
    cmdAdd.SetFocus
    
End Sub

Private Sub MSHFlexPag_Click()
    
    Dim idPag As Integer
    
    idPag = MSHFlexPag.TextMatrix(MSHFlexPag.Row, 1)
    
    Select Case MsgBox("¿Está seguro que desea eliminar el pago?", vbYesNo Or vbQuestion Or vbDefaultButton1, App.Title)
        Case vbNo: Exit Sub
    End Select
    
    'Elimina el registro de la base
    Set rsBorraPag = New ADODB.Recordset
    SQL = "DELETE FROM recibosp WHERE idpago = '" & idPag & "'"
    rsBorraPag.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    cargaFlexPag
    calculaTotal
    
    cboTipo.SetFocus
    
End Sub

Private Sub txtAbonoFac_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdOkImp_Click
    End If
    CambiaPunto txtAbonoFac, KeyAscii
    
End Sub

Private Sub txtCodCli_Change()
    
    Set rsPro = New ADODB.Recordset
    SQL = "SELECT * FROM clientes WHERE codigo = '" & txtCodCli.Text & "' AND eliminado <> 1"
    rsPro.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsPro.BOF And Not rsPro.EOF Then
        If rsPro!nombre <> "" Then
            cboNomCli.Text = rsPro!nombre
        End If
        rsPro.Close
    Else
        cboNomCli.ListIndex = -1
    End If
    
End Sub

Private Sub txtCodCli_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtCPago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDescFac_Change()
    
    If txtDescFac.Text = "" Or txtDescFac.Text = "0" Or txtDescFac.Text = "-" Or txtTotalFac.Text = "" Then
        DescPesos = 0
        txtSaldoFac.Text = txtTotalFac.Text
        txtAbonoFac.Text = txtTotalFac.Text
        txtAbonoFac.Enabled = True
        Exit Sub
    End If
    
    DescPesos = (CDec(txtTotalFac.Text) * CDec(txtDescFac.Text)) / 100
    
    txtSaldoFac.Text = Format(CDec(txtTotalFac.Text) - DescPesos, "0.00")
    txtAbonoFac.Text = txtSaldoFac.Text
    txtAbonoFac.Enabled = False
    
End Sub

Private Sub txtDescFac_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    Else
        CambiaPunto txtDescFac, KeyAscii
    End If
    
End Sub

Private Sub txtDetalle_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If

End Sub

Private Sub txtFechaPago_GotFocus()
    
    txtFechaPago.SelStart = 0
    txtFechaPago.SelLength = Len(txtFechaPago.Text) + 1
    
End Sub

Private Sub txtFechaPago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtImportePago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    CambiaPunto txtImportePago, KeyAscii
    
End Sub

Private Sub txtNHoja_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtNomCli_Change()
    
    SendMessage cboNomCli.hWnd, LB_SELECTSTRING, _
    -1, txtNomCli.Text
    
End Sub

Private Sub txtNomCli_GotFocus()
    
    cboNomCli.Visible = True
    
End Sub

Private Sub txtNomCli_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cboNomCli_Click
        PasarFoco
    End If
    
End Sub

Private Sub txtNumFac_GotFocus()
    
    cmdAdd_Click
    
End Sub

Private Sub txtTipoFac_GotFocus()
    
    cmdAdd_Click
    
End Sub

