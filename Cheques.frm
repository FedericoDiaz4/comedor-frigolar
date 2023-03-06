VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Cheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheques"
   ClientHeight    =   5175
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
   ScaleHeight     =   5175
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
      Picture         =   "Cheques.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   " Salir "
      Top             =   4440
      Width           =   855
   End
   Begin VB.Frame frmEgreso 
      Caption         =   "Datos de Egreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   10215
      Begin VB.TextBox txtProveedor 
         BackColor       =   &H8000000F&
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
         Left            =   2640
         TabIndex        =   27
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox txtPago 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   7680
         TabIndex        =   28
         Top             =   480
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTEgreso 
         Height          =   345
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   89980929
         CurrentDate     =   36526
      End
      Begin VB.Label Label26 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Fecha de Egreso"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label24 
         Caption         =   "Nº Orden de Pago"
         Height          =   255
         Left            =   7680
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frmCheques 
      Caption         =   "Datos del Cheque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   10215
      Begin VB.TextBox txtObs 
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
         Left            =   7680
         TabIndex        =   21
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtClering 
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
         Left            =   5160
         TabIndex        =   20
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtCUIT 
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
         Left            =   2640
         TabIndex        =   19
         Top             =   1200
         Width           =   2415
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
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtNCheque 
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
         Left            =   7680
         TabIndex        =   13
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtBanco 
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
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   4935
      End
      Begin MSComCtl2.DTPicker DTEmision 
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   89980929
         CurrentDate     =   36526
      End
      Begin VB.Label Label23 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   7680
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Clering (hs)"
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "CUIT (Firmante)"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Importe"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Nº Cheque"
         Height          =   255
         Left            =   7680
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Banco"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Fecha de Emisión"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame frmIngreso 
      Caption         =   "Datos de Ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10215
      Begin VB.TextBox txtRecibo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   7680
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtCliente 
         BackColor       =   &H8000000F&
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
         Left            =   2640
         TabIndex        =   5
         Top             =   480
         Width           =   4935
      End
      Begin MSComCtl2.DTPicker DTIngreso 
         Height          =   345
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   89980929
         CurrentDate     =   36526
      End
      Begin VB.Label Label22 
         Caption         =   "Nº Recibo"
         Height          =   255
         Left            =   7680
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha de Ingreso"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
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
      Left            =   8520
      Picture         =   "Cheques.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   " Guardar"
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "Cheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ID As Integer
Public idCli As Integer
Public idPro As Integer
Public idRecibo As Integer
Public idPago As Integer
Public Nuevo As Boolean
Public Entrada As String

Private Sub cmdGuardar_Click()
    
    'Comprueba que los campos obligatorios no estén vacíos
    If DTEmision.Value = "01/01/2000" Then
        Call MsgBox("Debe completar la fecha de emisión del cheque.     ", vbExclamation, App.Title)
        DTEmision.SetFocus
        Exit Sub
    End If
    
    If txtImporte.Text = "" Or CDec(txtImporte.Text) = 0 Then
        Call MsgBox("Debe ingresar el importe del cheque.     ", vbExclamation, App.Title)
        txtImporte.SetFocus
        Exit Sub
    End If
    
    
    'Guarda el Cheque
    Set rsGuardar = New ADODB.Recordset
    SQL = "SELECT * FROM cheques WHERE id = " & ID
    rsGuardar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Nuevo = True Then
        rsGuardar.AddNew
    End If
    
    rsGuardar!ingreso = DTIngreso.Value
    rsGuardar!idCliente = idCli
    rsGuardar!idRecibo = idRecibo
    
    rsGuardar!emision = DTEmision.Value
    rsGuardar!banco = txtBanco.Text
    rsGuardar!ncheque = txtNCheque.Text
    rsGuardar!importe = txtImporte.Text
    rsGuardar!cuit = txtCUIT.Text
    rsGuardar!clering = txtClering.Text
    rsGuardar!obs = txtObs.Text
    
    rsGuardar!egreso = DTEgreso.Value
    rsGuardar!idProveedor = idPro
    rsGuardar!idPago = idPago
    
    rsGuardar!eliminado = 0
    rsGuardar.Update
    
    rsGuardar.Close
    
    If Entrada = "RECIBOS" Then
        
        Recibos.txtDetalle.Text = txtBanco.Text & "  #" & txtNCheque.Text
        Recibos.txtFechaPago.Text = Format(DTIngreso.Value, "dd/mm/yyyy")
        Recibos.txtImportePago.Text = txtImporte.Text
        Recibos.Show
        Recibos.cmdOkPag_Click
        
    End If
    
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    
    DTIngreso.Value = Date
    txtCliente.Text = getData(idCli, "nombre", "clientes")
    txtProveedor.Text = getData(idPro, "nombre", "proveedores")
    
    txtRecibo.Text = idRecibo
    txtPago.Text = idPago
    
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    
    CambiaPunto txtImporte, KeyAscii
    
End Sub
