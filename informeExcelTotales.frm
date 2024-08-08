VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form informeExcelTotales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe Excel"
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
      ItemData        =   "informeExcelTotales.frx":0000
      Left            =   120
      List            =   "informeExcelTotales.frx":004C
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
      ItemData        =   "informeExcelTotales.frx":0161
      Left            =   120
      List            =   "informeExcelTotales.frx":01AD
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
      Picture         =   "informeExcelTotales.frx":02C2
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
      Picture         =   "informeExcelTotales.frx":084C
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
Attribute VB_Name = "informeExcelTotales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()

    Dim empresa As String
    Dim idturno As Integer
    Dim desde As Date
    Dim hasta As Date
    Dim idTipo As Integer
    Dim idEmpresa As Integer

    desde = dtpDesde.Value
    hasta = dtpHasta.Value
    idTipo = cboTipos.ItemData(cboTipos.ListIndex)
    idEmpresa = CboEmpresa.ItemData(CboEmpresa.ListIndex)
    
    Exportar_Excel_Totales "C:\Sistema\Listado.xls", desde, hasta, idTipo, idEmpresa

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Me.Width = 4590
    'Me.Height = 2955
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
    initForm Me
    
    
End Sub
