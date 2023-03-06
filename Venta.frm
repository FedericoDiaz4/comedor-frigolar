VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Venta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venta"
   ClientHeight    =   7695
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Venta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11415
   Begin VB.Timer Timer 
      Left            =   10800
      Top             =   2760
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9360
      Picture         =   "Venta.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   " Imprimir Comprobante"
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Frame frmCAE 
      Caption         =   " Autorización AFIP "
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   9360
      TabIndex        =   8
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdObtenerCAE 
         Caption         =   "Obtener CAE"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   " Obtener CAE"
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtCAE 
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
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1680
      End
      Begin MSComCtl2.DTPicker DTVencimientoCAE 
         Height          =   360
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   " Seleccionar Fecha"
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Unicode"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   89522177
         CurrentDate     =   42005
      End
      Begin VB.Label Label22 
         Caption         =   "Vencimiento CAE"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "CAE"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   9360
      Picture         =   "Venta.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   " Guardar Comprobante"
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtTotalTributos 
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
      Left            =   9360
      TabIndex        =   66
      Text            =   "0,00"
      Top             =   4800
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
      Left            =   9360
      TabIndex        =   68
      Text            =   "0,00"
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox txtTotalIVA 
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
      Left            =   9360
      TabIndex        =   64
      Text            =   "0,00"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtTotalNeto 
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
      Left            =   9360
      TabIndex        =   62
      Text            =   "0,00"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtNumComprobante 
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
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      TabIndex        =   6
      Text            =   "00000000"
      Top             =   360
      Width           =   1680
   End
   Begin VB.ComboBox cboPtoVenta 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "Venta.frx":109E
      Left            =   3840
      List            =   "Venta.frx":10A5
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cboTipoComprobante 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "Venta.frx":10AC
      Left            =   240
      List            =   "Venta.frx":10AE
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTFechaComprobante 
      Height          =   360
      Left            =   7440
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   " Seleccionar Fecha"
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
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
      Format          =   89522177
      CurrentDate     =   40544
   End
   Begin VB.Frame frmCli 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   9135
      Begin VB.TextBox txtEmail 
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
         Left            =   5520
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1320
         Width           =   3480
      End
      Begin VB.TextBox txtNumeroDocumento 
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
         Left            =   7320
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   1680
      End
      Begin VB.ComboBox cboTipoDocumento 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "Venta.frx":10B0
         Left            =   5520
         List            =   "Venta.frx":10B2
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtDireccion 
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5280
      End
      Begin VB.ComboBox cboCliente 
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "Venta.frx":10B4
         Left            =   1920
         List            =   "Venta.frx":10B6
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtCodCliente 
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
         TabIndex        =   18
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label Label20 
         Caption         =   "Email"
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Documento"
         Height          =   255
         Left            =   5520
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   3240
      Width           =   8895
      Begin VB.ComboBox cboIVA 
         ForeColor       =   &H00000000&
         Height          =   312
         ItemData        =   "Venta.frx":10B8
         Left            =   5400
         List            =   "Venta.frx":10CB
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   960
         Width           =   1695
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
         Height          =   360
         Left            =   0
         TabIndex        =   38
         Top             =   960
         Width           =   1680
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
         Enabled         =   0   'False
         Height          =   360
         Left            =   7200
         TabIndex        =   43
         Top             =   960
         Width           =   1680
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
         Height          =   360
         Left            =   0
         TabIndex        =   31
         Top             =   240
         Width           =   1680
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
         Height          =   360
         Left            =   1800
         TabIndex        =   32
         Top             =   240
         Width           =   7080
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
         Height          =   2775
         Left            =   0
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4895
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12640511
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   -2147483648
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
      Begin VB.TextBox txtPrecioU 
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
         Height          =   360
         Left            =   1800
         TabIndex        =   39
         Top             =   960
         Width           =   1680
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
         Height          =   360
         Left            =   3600
         TabIndex        =   41
         Text            =   "0"
         Top             =   960
         Width           =   1680
      End
      Begin VB.TextBox txtNeto 
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
         Height          =   360
         Left            =   3600
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label14 
         Caption         =   "IVA"
         Height          =   255
         Left            =   5400
         TabIndex        =   36
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Descuento"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Precio Unitario"
         Height          =   255
         Left            =   1800
         TabIndex        =   34
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Importe"
         Height          =   255
         Left            =   7200
         TabIndex        =   37
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Código"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   0
         Width           =   1935
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   4815
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2760
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Artículos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "IVA"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributos"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   1
      Left            =   240
      TabIndex        =   45
      Top             =   3240
      Width           =   8895
      Begin VB.Frame frmDetalleNeto 
         Height          =   2415
         Left            =   6720
         TabIndex        =   46
         Top             =   -120
         Width           =   2175
         Begin VB.TextBox txtNetoExento 
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
            Left            =   120
            TabIndex        =   52
            Text            =   "0,00"
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txtNetoNoGravado 
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
            Left            =   120
            TabIndex        =   50
            Text            =   "0,00"
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtNetoGravado 
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
            Left            =   120
            TabIndex        =   48
            Text            =   "0,00"
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label26 
            Caption         =   "Neto Exento"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label27 
            Caption         =   "Neto No Gravado"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "Neto Gravado"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1935
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexIVA 
         Height          =   4200
         Left            =   0
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7408
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12640511
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   -2147483648
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
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   2
      Left            =   240
      TabIndex        =   54
      Top             =   3240
      Width           =   8895
      Begin VB.TextBox txtDescripcionTributo 
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
         Left            =   3600
         TabIndex        =   59
         Top             =   240
         Width           =   3480
      End
      Begin VB.TextBox txtAlicuotaTributo 
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
         Height          =   360
         Left            =   7200
         TabIndex        =   60
         Top             =   240
         Width           =   1680
      End
      Begin VB.ComboBox cboTributos 
         ForeColor       =   &H00000000&
         Height          =   312
         ItemData        =   "Venta.frx":1108
         Left            =   0
         List            =   "Venta.frx":111B
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   240
         Width           =   3495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexTributos 
         Height          =   3495
         Left            =   0
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6165
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12640511
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   -2147483648
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
      Begin VB.Label Label25 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   3600
         TabIndex        =   56
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label24 
         Caption         =   "Alícuota"
         Height          =   255
         Left            =   7200
         TabIndex        =   57
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label23 
         Caption         =   "Tributo"
         Height          =   255
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Total Tributos"
      Height          =   255
      Left            =   9360
      TabIndex        =   65
      Top             =   4560
      Width           =   1935
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
      Left            =   9360
      TabIndex        =   67
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "Total IVA"
      Height          =   255
      Left            =   9360
      TabIndex        =   63
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Total Neto"
      Height          =   255
      Left            =   9360
      TabIndex        =   27
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Pto. Venta"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Comprobante"
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Tipo Comprobante"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modificando As Boolean
Public facturandoRemito As Boolean
Public idRemitoAFacturar As Integer
Public ID As String
Dim idArt As Single
Dim idArticuloSeleccionado As Integer
Dim modificandoArticulo As Boolean
Dim tabla As String
Dim tablad As String
Dim intentos As Integer

Private Sub cboCliente_Click()
    
    If cboCliente.ListIndex <> -1 Then
        txtCodCliente.Text = getData(cboCliente.ItemData(cboCliente.ListIndex), "codigo", "clientes")
    End If
    
End Sub

Private Sub cboIVA_Click()
    
    CalcImporte
    
End Sub

Private Sub cboIVA_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Add
        KeyAscii = 0
    End If
    
End Sub

Private Sub cboPtoVenta_Click()
    
    Set rsCargaNum = New ADODB.Recordset
    SQL = "SELECT ultimocomprobante FROM indices WHERE idtipocomprobante = " & cboTipoComprobante.ItemData(cboTipoComprobante.ListIndex) & " AND puntoventa = " & CInt(cboPtoVenta.Text)
    rsCargaNum.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsCargaNum.BOF And Not rsCargaNum.EOF Then
        txtNumComprobante.Text = Format(CInt(rsCargaNum!ultimocomprobante) + 1, "00000000")
    End If
    rsCargaNum.Close
    
End Sub

Private Sub cboTipoComprobante_Click()
    
    If cboTipoComprobante.Text = "" Then Exit Sub
    
    cboPtoVenta.Clear
    
    Set rsCargaPV = New ADODB.Recordset
    SQL = "SELECT puntoventa FROM indices WHERE idtipocomprobante = " & cboTipoComprobante.ItemData(cboTipoComprobante.ListIndex)
    rsCargaPV.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsCargaPV.BOF And Not rsCargaPV.EOF Then
        Do While Not rsCargaPV.EOF
            
            cboPtoVenta.AddItem rsCargaPV!puntoventa
            rsCargaPV.MoveNext
            
        Loop
        
        cboPtoVenta.ListIndex = 0
    End If
    rsCargaPV.Close
    
End Sub

Private Sub cboTipoComprobante_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtCodCliente.SetFocus
        KeyAscii = 0
    End If
    
End Sub

Private Sub cboTipoDocumento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If

End Sub

Private Sub cboTributos_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If

End Sub

Private Sub cmdImprimir_Click()
    
    'Vacio la tabla temporal
    Set rsDel = New ADODB.Recordset
    SQL = "DELETE FROM temp"
    rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Abro la tabla temporal
    Set rsTemp = New ADODB.Recordset
    SQL = "SELECT * FROM temp"
    rsTemp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    For i = 1 To Flex.Rows - 1
        rsTemp.AddNew
        rsTemp!col1 = Flex.TextMatrix(i, 4)
        rsTemp!col2 = Flex.TextMatrix(i, 3)
        rsTemp!col3 = Flex.TextMatrix(i, 5)
        rsTemp!col4 = Flex.TextMatrix(i, 10)
        rsTemp.Update
    Next i
    
    Set drComprobante.DataSource = rsTemp
    
    drComprobante.Sections("ReportHeader").Controls("lblTipo").Caption = Left(cboTipoComprobante.Text, Len(cboTipoComprobante.Text) - 1)
    drComprobante.Sections("ReportHeader").Controls("lblNumero").Caption = Format(cboPtoVenta.Text, "0000") & " - " & Format(txtNumComprobante.Text, "00000000")
    drComprobante.Sections("ReportHeader").Controls("lblFecha").Caption = "Fecha: " & Format(DTFechaComprobante.Value, "dd/mm/yyyy")
    
    drComprobante.Sections("ReportHeader").Controls("lblNombre").Caption = cboCliente.Text
    drComprobante.Sections("ReportHeader").Controls("lblDireccion").Caption = txtDireccion.Text
    drComprobante.Sections("ReportHeader").Controls("lblLocalidad").Caption = LocalidadCli
    drComprobante.Sections("ReportHeader").Controls("lblTipoIVA").Caption = "IVA: " & getData(idTipoResponsable, "nombre", "tiporesponsable")
    drComprobante.Sections("ReportHeader").Controls("lblCUIT").Caption = "CUIT: " & txtNumeroDocumento.Text
    
    If esComprobanteA Then
        drComprobante.Sections("ReportHeader").Controls("lblLetra").Caption = "[ A ]"
        drComprobante.Sections("PageFooter").Controls("Etiqueta21").Visible = True
        drComprobante.Sections("PageFooter").Controls("Etiqueta22").Visible = True
        drComprobante.Sections("PageFooter").Controls("Etiqueta24").Visible = True
        drComprobante.Sections("PageFooter").Controls("Etiqueta25").Visible = True
        drComprobante.Sections("PageFooter").Controls("lblSubtotal1").Caption = Format(txtTotalNeto.Text, "0.00")
        drComprobante.Sections("PageFooter").Controls("lblImpuestos").Caption = Format(txtTotalTributos.Text, "0.00")
        drComprobante.Sections("PageFooter").Controls("lblSubtotal2").Caption = Format(txtTotalNeto.Text, "0.00")
        drComprobante.Sections("PageFooter").Controls("lblTotalIVA").Caption = Format(txtTotalIVA.Text, "0.00")
    Else
        drComprobante.Sections("ReportHeader").Controls("lblLetra").Caption = "[ B ]"
        drComprobante.Sections("PageFooter").Controls("Etiqueta21").Visible = False
        drComprobante.Sections("PageFooter").Controls("Etiqueta22").Visible = False
        drComprobante.Sections("PageFooter").Controls("Etiqueta24").Visible = False
        drComprobante.Sections("PageFooter").Controls("Etiqueta25").Visible = False
        drComprobante.Sections("PageFooter").Controls("lblSubtotal1").Caption = ""
        drComprobante.Sections("PageFooter").Controls("lblImpuestos").Caption = ""
        drComprobante.Sections("PageFooter").Controls("lblSubtotal2").Caption = ""
        drComprobante.Sections("PageFooter").Controls("lblTotalIVA").Caption = ""
    End If
    drComprobante.Sections("PageFooter").Controls("lblTotal").Caption = Format(txtTotal.Text, "0.00")
    
    drComprobante.Sections("PageFooter").Controls("lblFechaVto").Caption = "Fecha Vto: " & Format(DTVencimientoCAE.Value, "dd/mm/yyyy")
    drComprobante.Sections("PageFooter").Controls("lblCAE").Caption = "C.A.E. Nº: " & txtCAE.Text
    drComprobante.Show
    
End Sub

Private Sub cmdObtenerCAE_Click()
    
    If existeArchivo("C:\FE\resul.txt") Then
        Kill "C:\FE\resul.txt"
    End If
    
    Open "C:\FE\comp.xml" For Output As #1
    
    Print #1, "<?xml version=""1.0"" encoding=""iso-8859-1""?>"
    Print #1, "<Comprobante>"
    Print #1, "    <Empresa>prueba</Empresa>"
    Print #1, "    <CbteTipo>" & cboTipoComprobante.ItemData(cboTipoComprobante.ListIndex) & "</CbteTipo>"
    Print #1, "    <PtoVta>" & cboPtoVenta.Text & "</PtoVta>"
    Print #1, "    <CbteNum>" & CInt(txtNumComprobante.Text) & "</CbteNum>"
    Print #1, "    <DocTipo>" & cboTipoDocumento.ItemData(cboTipoDocumento.ListIndex) & "</DocTipo>"
    Print #1, "    <DocNro>" & txtNumeroDocumento.Text & "</DocNro>"
    Print #1, "    <CbteFch>" & Format(DTFechaComprobante.Value, "dd/mm/yyyy") & "</CbteFch>"
    Print #1, "    <ImpTotal>" & Format(txtTotal.Text, "0.00") & "</ImpTotal>"
    Print #1, "    <ImpTotConc>" & Format(txtNetoNoGravado.Text, "0.00") & "</ImpTotConc>"
    Print #1, "    <ImpNeto>" & Format(txtNetoGravado.Text, "0.00") & "</ImpNeto>"
    Print #1, "    <ImpOpEx>" & Format(txtNetoExento.Text, "0.00") & "</ImpOpEx>"
    Print #1, "    <ImpTrib>" & Format(txtTotalTributos.Text, "0.00") & "</ImpTrib>"
    Print #1, "    <ImpIVA>" & Format(txtTotalIVA.Text, "0.00") & "</ImpIVA>"
    
    If txtTotalIVA.Text <> "0,00" Then
        Print #1, "    <Iva>"
        For i = 1 To FlexIVA.Rows - 1
            Print #1, "        <AlicIva>"
            Print #1, "            <Id>" & FlexIVA.TextMatrix(i, 0) & "</Id>"
            Print #1, "            <BaseImp>" & FlexIVA.TextMatrix(i, 2) & "</BaseImp>"
            Print #1, "            <Importe>" & FlexIVA.TextMatrix(i, 3) & "</Importe>"
            Print #1, "        </AlicIva>"
        Next i
        Print #1, "    </Iva>"
    End If
    
    If txtTotalTributos.Text <> "0,00" Then
        Print #1, "    <Tributos>"
        For i = 1 To FlexTributos.Rows - 1
            Print #1, "        <Tributo>"
            Print #1, "            <Id>" & FlexTributos.TextMatrix(i, 1) & "</Id>"
            Print #1, "            <Desc>" & FlexTributos.TextMatrix(i, 3) & "</Desc>"
            Print #1, "            <BaseImp>" & Format(txtNetoGravado.Text, "0.00") & "</BaseImp>"
            Print #1, "            <Alic>" & FlexTributos.TextMatrix(i, 4) & "</Alic>"
            Print #1, "            <Importe>" & FlexTributos.TextMatrix(i, 5) & "</Importe>"
            Print #1, "        </Tributo>"
        Next i
        Print #1, "    </Tributos>"
    End If
    
    Print #1, "</Comprobante>"
    Close #1
    
    Timer.Interval = 1000
    
End Sub

Private Sub FlexTributos_Click()
    
    Dim idTributoSeleccionado As Integer
    
    If FlexTributos.TextMatrix(FlexTributos.Row, 0) = "" Then Exit Sub
    
    idTributoSeleccionado = FlexTributos.TextMatrix(FlexTributos.Row, 0)
    Select Case MsgBox("¿DESEA ELIMINAR EL TRIBUTO SELECCIONADO?", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
        Case vbNo: Exit Sub
    End Select
    
    Set rsDel = New ADODB.Recordset
    SQL = "DELETE FROM ventastrib WHERE id = " & idTributoSeleccionado
    rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ID = 0
    Modificando = False
    modificandoArticulo = False
    Timer.Interval = 0
    
End Sub

Private Sub TabStrip_Click()
    
    'Cuando se pulsa en el tabstrip...
    Dim i&
    
    i = TabStrip.SelectedItem.Index
    'Mostrar el contenedor que corresponda
    Frame(i - 1).ZOrder
    
End Sub

Private Sub Timer_Timer()
    
    If Not existeArchivo("C:\FE\resul.txt") Then
        Exit Sub
    End If
    
    Timer.Interval = 0
    
    Dim nFile As Integer
    Dim strLinea As String
    
    nFile = FreeFile
    
    ' Abre el archivo de texto
    Open "C:\FE\resul.txt" For Input As #nFile
    
    If EOF(nFile) Then
        Timer.Interval = 1000
        Close #nFile
        Exit Sub
    End If
    
    'Lee una línea
    Line Input #nFile, strLinea
    
    If strLinea = "A" Then
        Line Input #nFile, strLinea
        txtCAE.Text = strLinea
        
        Line Input #nFile, strLinea
        DTVencimientoCAE.Value = Mid(strLinea, 7, 2) & "/" & Mid(strLinea, 5, 2) & "/" & Mid(strLinea, 1, 4)
        
        cmdObtenerCAE.Enabled = False
        cmdImprimir.Enabled = True
    Else
        Dim Mensaje As String
        Mensaje = "COMPROBANTE RECHAZADO " & vbCrLf
        Do
            Line Input #nFile, strLinea
            Mensaje = Mensaje & strLinea & vbCrLf
        Loop While Not EOF(nFile)
        
        Call MsgBox(Mensaje, vbCritical Or vbDefaultButton1, App.Title)
        
    End If
    
    Close #nFile
    
End Sub

Private Sub txtAlicuotaTributo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        AddTributo
        KeyAscii = 0
    Else
        CambiaPunto txtAlicuotaTributo, KeyAscii
    End If

End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmdGuardar_Click()
    
    If txtTotalNeto.Text = "0,00" Then Exit Sub
    If cboTipoComprobante.ListIndex = -1 Then Exit Sub
    If cboPtoVenta.ListIndex = -1 Then Exit Sub
    
    Select Case MsgBox("¿DESEA GUARDAR EL COMPROBANTE?", vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
        Case vbNo: Exit Sub
    End Select
    
    'Guardar
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM ventas where id = " & ID
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Not Modificando Then
        
        Recordset.AddNew
        Recordset!idtipocomprobante = cboTipoComprobante.ItemData(cboTipoComprobante.ListIndex)
        Recordset!ptoventa = cboPtoVenta.Text
        Recordset!numerocomprobante = Format(txtNumComprobante.Text, "00000000")
        
    End If
    
    Recordset!idCliente = cboCliente.ItemData(cboCliente.ListIndex)
    Recordset!fecha = Format(DTFechaComprobante.Value, "dd/mm/yyyy")
    Recordset!netogravado = Format(txtNetoGravado.Text, "0.00")
    Recordset!netonogravado = Format(txtNetoNoGravado.Text, "0.00")
    Recordset!netoexento = Format(txtNetoExento.Text, "0.00")
    Recordset!totalneto = Format(txtTotalNeto.Text, "0.00")
    Recordset!totaliva = Format(txtTotalIVA.Text, "0.00")
    Recordset!totaltributos = Format(txtTotalTributos.Text, "0.00")
    Recordset!Total = Format(txtTotal.Text, "0.00")
    Recordset!Estado = "CTACTE"
    Recordset!Saldo = Format(txtTotal.Text, "0.00")
    Recordset!cae = txtCAE.Text
    Recordset!fechavtocae = Format(DTVencimientoCAE.Value, "dd/mm/yyyy")
    Recordset!DateTime = DTFechaComprobante.Value & " " & Time
    Recordset.Update
    ID = Recordset!ID
    Recordset.Close
    
    'Actualiza las tablas relacionadas
    Set rsArt = New ADODB.Recordset
    SQL = "UPDATE ventasd SET idventa = " & ID & " WHERE idventa = 0;"
    rsArt.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Set rsTrib = New ADODB.Recordset
    SQL = "UPDATE ventastrib SET idventa = " & ID & " WHERE idventa = 0;"
    rsTrib.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Actualiza la tabla de indices
    If Not Modificando Then
        Set rsIndex = New ADODB.Recordset
        SQL = "UPDATE indices SET ultimocomprobante = " & CInt(txtNumComprobante.Text) & " WHERE idtipocomprobante = " & cboTipoComprobante.ItemData(cboTipoComprobante.ListIndex) & " AND puntoventa = " & cboPtoVenta.Text & ";"
        rsIndex.Open SQL, Data, adOpenKeyset, adLockOptimistic
    End If
    
    If txtCAE.Text <> "" Then
        ID = 0
        Unload Me
    Else
        Modificando = True
        Cargar
    End If
    
End Sub

Private Sub Flex_Click()
    
    modificandoArticulo = True
    idArticuloSeleccionado = Flex.TextMatrix(Flex.Row, 0)
    If Flex.TextMatrix(Flex.Row, 3) = "" Then
        txtCodArticulo.Text = "000000"
    Else
        txtCodArticulo.Text = Flex.TextMatrix(Flex.Row, 3)
    End If
    txtArticulo.Text = Flex.TextMatrix(Flex.Row, 4)
    txtPrecioU.Text = Flex.TextMatrix(Flex.Row, 6)
    txtCantidad.Text = Flex.TextMatrix(Flex.Row, 5)
    txtDescuento.Text = Flex.TextMatrix(Flex.Row, 7)
    cboIVA.Text = Flex.TextMatrix(Flex.Row, 9)
    txtCantidad.SetFocus
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    DTFechaComprobante.Value = Date
    
    CargaCombo "tipocomprobante", "nombre", "id", cboTipoComprobante
    CargaCombo "tipodocumento", "nombre", "id", cboTipoDocumento
    CargaCombo "clientes", "nombre", "nombre", cboCliente
    CargaCombo "tipoiva", "nombre", "id", cboIVA
    CargaCombo "tributos", "nombre", "id", cboTributos
    tabla = "ventas"
    tablad = "ventasd"
    
    cboTipoComprobante.ListIndex = 0
    
    Frame(0).ZOrder
    
    'Borra comprobantes temporales no guardados
    Set rsDel = New ADODB.Recordset
    SQL = "DELETE FROM ventasd WHERE idventa = 0"
    rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Set rsDel = New ADODB.Recordset
    SQL = "DELETE FROM ventastrib WHERE idventa = 0"
    rsDel.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Cargar
    
End Sub

Sub Cargar()
    
    txtNetoNoGravado.Text = "0,00"
    txtNetoExento.Text = "0,00"
    txtNetoGravado.Text = "0,00"
    txtTotalNeto.Text = "0,00"
    txtTotalIVA.Text = "0,00"
    txtTotalTributos.Text = "0,00"
    txtTotal.Text = "0,00"
    
    If Modificando Then
        If txtCAE.Text = "" Then
            Frame(0).Enabled = True
            Frame(1).Enabled = True
            Frame(2).Enabled = True
            cmdObtenerCAE.Enabled = True
            cmdImprimir.Enabled = False
            cmdGuardar.Enabled = True
        Else
            Frame(0).Enabled = False
            Frame(1).Enabled = False
            Frame(2).Enabled = False
            cmdObtenerCAE.Enabled = False
            cmdImprimir.Enabled = True
            cmdGuardar.Enabled = False
        End If
    Else
        Frame(0).Enabled = True
        Frame(1).Enabled = True
        Frame(2).Enabled = True
        cmdObtenerCAE.Enabled = False
        cmdImprimir.Enabled = False
        cmdGuardar.Enabled = True
    End If
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT v.id, v.idventa, v.idart, a.codigo, v.art, v.cantidad, v.precio, v.descuento, v.descuentoimp, tipoiva.nombre, v.ivaimp, v.importe "
    SQL = SQL & "FROM ventasd AS v Inner Join tipoiva ON v.idtipoiva = tipoiva.id "
    SQL = SQL & "Left Join articulos AS a ON v.idart = a.id WHERE v.idventa = " & ID & " ORDER BY v.id ASC"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Recordset.BOF And Recordset.EOF Then
        Flex.Clear
        Flex.Rows = 2
    Else
        Set Flex.DataSource = Recordset
    End If
    Recordset.Close
    
    OrdenaFlex
    
    Dim idTipoIVA As Integer
    Dim i As Integer
    
    idTipoIVA = 0
    FlexIVA.Clear
    FlexIVA.Rows = 2
    FlexIVA.Cols = 4
    OrdenaFlexIVA
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT v.baseimp, v.idtipoiva, tipoiva.nombre, tipoiva.alicuota, v.ivaimp, v.importe, v.total "
    SQL = SQL & "FROM ventasd AS v Inner Join tipoiva ON v.idtipoiva = tipoiva.id "
    SQL = SQL & "WHERE v.idventa = " & ID & " ORDER BY v.idtipoiva ASC"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not Recordset.EOF
    
        'Calcula las alicuotas
        If idTipoIVA <> Recordset!idTipoIVA Then
            idTipoIVA = Recordset!idTipoIVA
            i = FlexIVA.Rows
            
            If FlexIVA.TextMatrix(i - 1, 0) <> "" Then
                FlexIVA.Rows = i + 1
                i = i + 1
            End If
            
            FlexIVA.TextMatrix(i - 1, 0) = Recordset!idTipoIVA
            FlexIVA.TextMatrix(i - 1, 1) = Recordset!nombre
            FlexIVA.TextMatrix(i - 1, 2) = Recordset!baseimp
            FlexIVA.TextMatrix(i - 1, 3) = Recordset!ivaimp
        Else
            FlexIVA.TextMatrix(i - 1, 2) = Format(CDec(FlexIVA.TextMatrix(i - 1, 2)) + CDec(Recordset!baseimp), "0.00")
            FlexIVA.TextMatrix(i - 1, 3) = Format(CDec(FlexIVA.TextMatrix(i - 1, 3)) + CDec(Recordset!ivaimp), "0.00")
        End If
        
'        If Recordset!idTipoIVA = 3 Then
'            'No Gravado
'            txtNetoNoGravado.Text = CDec(txtNetoNoGravado.Text) + CDec(Recordset!baseimp)
'        ElseIf Recordset!idTipoIVA = 2 Then
'            'Exento
'            txtNetoExento.Text = CDec(txtNetoExento.Text) + CDec(Recordset!baseimp)
'        Else
            'Gravado
            txtNetoGravado.Text = CDec(txtNetoGravado.Text) + CDec(Recordset!baseimp)
'        End If
        
        'Calcula los totales
        txtTotalNeto.Text = CDec(txtTotalNeto.Text) + CDec(Recordset!baseimp)
        txtTotalIVA.Text = CDec(txtTotalIVA.Text) + CDec(Recordset!ivaimp)
        txtTotal.Text = CDec(txtTotal.Text) + CDec(Recordset!Total)
        Recordset.MoveNext
    Loop
    Recordset.Close
    
    ''Calcula los tributos
    FlexTributos.Clear
    FlexTributos.Rows = 2
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT ventastrib.id, ventastrib.idtributo, t.nombre, ventastrib.descripcion, ventastrib.alicuota, ventastrib.importe FROM ventastrib Inner Join tributos AS t ON t.id = ventastrib.idtributo "
    SQL = SQL & "WHERE ventastrib.idventa = " & ID
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        Do While Not Recordset.EOF
            Recordset!importe = Format(CDec(txtTotalNeto.Text) * CDec(Recordset!alicuota) / 100, "0.00")
            txtTotalTributos.Text = CDec(txtTotalTributos.Text) + CDec(Recordset!importe)
            Recordset.Update
            Recordset.MoveNext
        Loop
        Set FlexTributos.DataSource = Recordset
        txtTotal.Text = CDec(txtTotal.Text) + CDec(txtTotalTributos.Text)
    End If
    Recordset.Close
    OrdenaFlexTributos
    
    txtNetoNoGravado.Text = Format(txtNetoNoGravado.Text, "0.00")
    txtNetoExento.Text = Format(txtNetoExento.Text, "0.00")
    txtNetoGravado.Text = Format(txtNetoGravado.Text, "0.00")
    txtTotalNeto.Text = Format(txtTotalNeto.Text, "0.00")
    txtTotalIVA.Text = Format(txtTotalIVA.Text, "0.00")
    txtTotalTributos.Text = Format(txtTotalTributos.Text, "0.00")
    txtTotal.Text = Format(txtTotal.Text, "0.00")
    
End Sub

Sub OrdenaFlex()
    
    Flex.FormatString = "id|idventa|idart|Código|Artículo|Cantidad|Precio|Desc %|DescImp|IVA %|IvaImp|Importe|"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 0
    Flex.ColWidth(2) = 0
    Flex.ColWidth(3) = 950  'Cod
    Flex.ColWidth(4) = 3050 'Art
    Flex.ColWidth(5) = 950  'Canti
    Flex.ColWidth(6) = 950  'Precio
    Flex.ColWidth(7) = 850  'Desc
    Flex.ColWidth(8) = 0    'DescImp
    Flex.ColWidth(9) = 850  'IVA
    Flex.ColWidth(10) = 0   'IvaImp
    Flex.ColWidth(11) = 950 'Importe
    Flex.ColWidth(12) = 0
    Flex.ColAlignment(3) = 1
    
End Sub

Sub OrdenaFlexIVA()
    
    FlexIVA.FormatString = "id|Alícuota IVA|Base Imponible|Importe"
    FlexIVA.ColWidth(0) = 0
    FlexIVA.ColWidth(1) = 2080
    FlexIVA.ColWidth(2) = 2080
    FlexIVA.ColWidth(3) = 2080
    
End Sub

Sub OrdenaFlexTributos()
    
    FlexTributos.FormatString = "id|idtrib|Tributo|Descripción|Alícuota|Importe"
    FlexTributos.ColWidth(0) = 0
    FlexTributos.ColWidth(1) = 0
    FlexTributos.ColWidth(2) = 2400
    FlexTributos.ColWidth(3) = 3150
    FlexTributos.ColWidth(4) = 1500
    FlexTributos.ColWidth(5) = 1500
    
End Sub

Sub Add()
    
    Dim SubTotal As Single
    
    If txtCodArticulo.Text = "" Then
        txtCodArticulo.SetFocus
        Exit Sub
    ElseIf txtCantidad.Text = "" Then
        txtCantidad.SetFocus
        Exit Sub
    ElseIf txtPrecioU.Text = "" Then
        txtPrecioU.SetFocus
        Exit Sub
    ElseIf txtImporte.Text = "" Then
        txtImporte.SetFocus
        Exit Sub
    ElseIf txtDescuento.Text = "" Then
        txtDescuento.Text = "0"
    ElseIf cboIVA.ListIndex = -1 Then
        cboIVA.Text = "21%"
    End If
    
    'Si la cantidad es cero, elimina el item
    If txtCantidad.Text = "0" Then
        
        Select Case MsgBox("¿Desea eliminar el item " & Flex.TextMatrix(Flex.Row, 3) & "?", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)
            Case vbNo: Exit Sub
        End Select
        
        modificandoArticulo = False
        
        Set rsDelete = New ADODB.Recordset
        SQL = "DELETE FROM " & tablad & " WHERE id = '" & idArticuloSeleccionado & "'"
        rsDelete.Open SQL, Data, adOpenKeyset, adLockOptimistic
        
        Cargar
        
        'Vaciar Box
        txtCodArticulo.Text = ""
        txtArticulo.Text = ""
        txtCantidad.Text = ""
        txtPrecioU.Text = ""
        txtDescuento.Text = "0"
        cboIVA.ListIndex = 2
        txtImporte.Text = ""
        txtCodArticulo.SetFocus
        
        Exit Sub
    End If
    
    SubTotal = CDec(txtPrecioU.Text) * CDec(txtCantidad.Text)
    
    'Guardar
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM " & tablad & " where id = " & idArticuloSeleccionado
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If modificandoArticulo Then
        modificandoArticulo = False
    Else
        Recordset.AddNew
        Recordset!idventa = ID
    End If
    If txtCodArticulo.Text = "" Then
        Recordset!idArt = 0
    Else
        Recordset!idArt = idArt
    End If
    Recordset!Art = txtArticulo.Text
    Recordset!Precio = Format(txtPrecioU.Text, "0.00")
    Recordset!Cantidad = txtCantidad.Text
    Recordset!descuento = txtDescuento.Text
    Recordset!descuentoimp = Format((SubTotal * CDec(txtDescuento.Text)) / 100, "0.00")
    Recordset!baseimp = Format(txtNeto.Text, "0.00")
    Recordset!idTipoIVA = cboIVA.ItemData(cboIVA.ListIndex)
    
    If cboIVA.ItemData(cboIVA.ListIndex) = 4 Then
        Recordset!ivaimp = Format(CDec(txtNeto.Text) * 0.105, "0.00")
    ElseIf cboIVA.ItemData(cboIVA.ListIndex) = 5 Then
        Recordset!ivaimp = Format(CDec(txtNeto.Text) * 0.21, "0.00")
    Else
        Recordset!ivaimp = "0,00"
    End If
    Recordset!importe = Format(txtImporte.Text, "0.00")
    Recordset!Total = Format(CDec(Recordset!baseimp) + CDec(Recordset!ivaimp), "0.00")
    Recordset.Update
    Recordset.Close
    
    'Mostrar
    Cargar
    
    'Vaciar Box
    txtCodArticulo.Text = ""
    txtArticulo.Text = ""
    txtCantidad.Text = ""
    txtDescuento.Text = "0"
    txtPrecioU.Text = ""
    txtNeto.Text = ""
    cboIVA.Text = "21%"
    txtImporte.Text = ""
    txtCodArticulo.SetFocus
    
End Sub

Sub AddTributo()
    
    If txtAlicuotaTributo.Text = "" Then
        txtAlicuotaTributo.SetFocus
        Exit Sub
    End If
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT * FROM ventastrib WHERE idventa = " & ID & " AND idtributo = " & cboTributos.ItemData(cboTipoComprobante.ListIndex)
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Recordset.AddNew
    Recordset!idtributo = cboTributos.ItemData(cboTipoComprobante.ListIndex)
    Recordset!idventa = ID
    Recordset!descripcion = txtDescripcionTributo.Text
    Recordset!alicuota = txtAlicuotaTributo.Text
    Recordset.Update
    Recordset.Close
    
    txtDescripcionTributo.Text = ""
    txtAlicuotaTributo.Text = ""
    Cargar
    
End Sub


Sub CalcImporte()
    
    If txtPrecioU.Text <> "" And txtCantidad.Text <> "" And txtDescuento.Text <> "" And cboIVA.ListIndex <> -1 Then
        
        Dim SubT As Single
        Dim alicuotaIVA As Single
        alicuotaIVA = getData(cboIVA.ItemData(cboIVA.ListIndex), "alicuota", "tipoiva")
        
        SubT = CDec(txtPrecioU.Text) * CDec(txtCantidad.Text)
        SubT = SubT * (1 - CDec(txtDescuento.Text) / 100)
        
        If esComprobanteA Then
            txtNeto.Text = Format(SubT, "0.00")
            txtImporte.Text = Format(SubT, "0.00")
        Else
            SubT = SubT / (1 + alicuotaIVA / 100)
            txtNeto.Text = Format(SubT, "0.00")
            txtImporte.Text = Format(SubT * (1 + alicuotaIVA / 100), "0.00")
        End If
        
    Else
        txtImporte.Text = "0,00"
    End If
    
End Sub

Private Sub txtCantidad_Change()
    
    CalcImporte
    
End Sub

Private Sub txtCantidad_GotFocus()
    
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad)
    
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
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT id, nombre, idtipoiva, precio FROM articulos WHERE codigo = '" & txtCodArticulo.Text & "' AND eliminado = 0 LIMIT 1"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Recordset.BOF And Recordset.EOF Then
        txtArticulo.Text = ""
        txtPrecioU.Text = "0,00"
        txtArticulo.Enabled = True
        Exit Sub
    End If
    
    Dim alicuotaIVA As Double
    
    idArt = Recordset!ID
    txtArticulo.Text = Recordset!nombre
    txtArticulo.Enabled = False
    cboIVA.Text = getData(Recordset!idTipoIVA, "nombre", "tipoiva")
    alicuotaIVA = getData(cboIVA.ItemData(cboIVA.ListIndex), "alicuota", "tipoiva")
    
    If esComprobanteA Then
        txtPrecioU.Text = Recordset!Precio
    Else
        txtPrecioU.Text = Format(CDec(Recordset!Precio) * (1 + alicuotaIVA / 100), "0.00")
    End If
    
    Recordset.Close
    
End Sub

Private Function esComprobanteA() As Boolean
    
    esComprobanteA = Right(cboTipoComprobante.Text, 1) = "A"
    
End Function

Private Sub txtCodArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        If txtArticulo.Text = "" And txtCodArticulo.Text <> "0" Then
            ArticulosList.Entrada = "VENTA"
            ArticulosList.WindowState = 0
            ArticulosList.BorderStyle = 1
            ArticulosList.Height = 5520
            ArticulosList.txtBuscar.Text = txtCodArticulo.Text
            ArticulosList.txtBuscar.SelStart = Len(txtCodArticulo.Text)
            ArticulosList.Show
            ArticulosList.Cargar
            ArticulosList.txtBuscar.SetFocus
        Else
            PasarFoco
            KeyAscii = 0
        End If
        
    End If
    
End Sub

Private Sub txtCodCliente_Change()
    
    If txtCodCliente.Text = "" Then
        cboCliente.ListIndex = -1
        txtDireccion.Text = ""
        cboTipoDocumento.ListIndex = 0
        txtNumeroDocumento.Text = ""
        txtEmail.Text = ""
        Exit Sub
    End If
    
    Set rsCli = New ADODB.Recordset
    SQL = "SELECT id, nombre, direccion, localidad, idtipodocumento, numerodocumento, email FROM clientes WHERE codigo = '" & txtCodCliente.Text & "' AND eliminado = 0"
    rsCli.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsCli.BOF And Not rsCli.EOF Then
        If rsCli!nombre <> "" Then
            cboCliente.Text = rsCli!nombre
        End If
        If rsCli!direccion <> "" Then
            txtDireccion.Text = rsCli!direccion & ", " & rsCli!localidad
        End If
        If rsCli!idTipoDocumento <> "0" Then
            cboTipoDocumento.Text = getData(rsCli!idTipoDocumento, "nombre", "tipodocumento")
        End If
        If rsCli!numerodocumento <> "" Then
            txtNumeroDocumento.Text = rsCli!numerodocumento
        End If
        If rsCli!email <> "" Then
            txtEmail.Text = rsCli!email
        End If
        rsCli.Close
    Else
        cboCliente.ListIndex = -1
        txtDireccion.Text = ""
        cboTipoDocumento.ListIndex = 0
        txtNumeroDocumento.Text = ""
        txtEmail.Text = ""
    End If
    
    txtCodCliente.SelStart = Len(txtCodCliente.Text)
    
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDescripcionTributo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDescuento_Change()
    
    CalcImporte
    
End Sub

Private Sub txtDescuento_GotFocus()
    
    txtDescuento.SelStart = 0
    txtDescuento.SelLength = Len(txtDescuento.Text)
    
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Add
        KeyAscii = 0
    Else
        CambiaPunto txtDescuento, KeyAscii, "-"
    End If
    
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If

End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If

End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
    Else
        CambiaPunto txtImporte, KeyAscii
    End If
    
End Sub

Private Sub txtNumeroDocumento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If

End Sub

Private Sub txtPrecioU_GotFocus()
    
    txtPrecioU.SelStart = 0
    txtPrecioU.SelLength = Len(txtPrecioU)
    
End Sub

Private Sub txtTotalIVA_GotFocus()
    
    txtTotalIVA.SelStart = 0
    txtTotalIVA.SelLength = Len(txtTotalIVA.Text)
    
End Sub

Private Sub txtTotalIVA_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtTotalIVA, KeyAscii
    End If
    
End Sub

Private Sub txtPrecioU_Change()
    
    CalcImporte
    
End Sub

Private Sub txtPrecioU_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    ElseIf KeyAscii = 32 Then
        If chkEspecial.Value = 0 Then
            chkEspecial.Value = 1
        Else
            chkEspecial.Value = 0
        End If
        KeyAscii = 0
    Else
        CambiaPunto txtPrecioU, KeyAscii
    End If
    
End Sub

Private Sub txtTotalNeto_GotFocus()
    
    txtTotalNeto.SelLength = Len(txtTotalNeto.Text)
    
End Sub

Private Sub txtTotalNeto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtTotalNeto, KeyAscii
    End If
    
End Sub
