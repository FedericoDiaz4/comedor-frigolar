VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TicketList 
   Caption         =   "Listado de Tickets"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   9855
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9855
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboPersona 
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
      ItemData        =   "TicketList.frx":0000
      Left            =   4440
      List            =   "TicketList.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Picture         =   "TicketList.frx":0030
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Salir "
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picBotones 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   8880
      ScaleHeight     =   615
      ScaleWidth      =   855
      TabIndex        =   6
      Top             =   5040
      Width           =   855
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
         Left            =   0
         Picture         =   "TicketList.frx":05BA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Salir "
         Top             =   0
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7435
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
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   240
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
      Format          =   53542913
      CurrentDate     =   43717
   End
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
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
      Format          =   117571585
      CurrentDate     =   43717
   End
   Begin VB.Label lblFechaDesde 
      Caption         =   "Desde"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblHasta 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "TicketList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOk_Click()

    Cargar

End Sub

Private Sub Form_Load()
    
    initForm Me
    dtDesde.Value = DateSerial(Year(Date), Month(Date) + 0, 1)
    dtHasta.Value = DateSerial(Year(Date), Month(Date) + 1, 0)
    cboPersona.ListIndex = 0
    Cargar
    
End Sub

Sub ordenaFlex()
    
    Flex.FormatString = "id|Fecha y Hora|Empleado|Importe"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 2500
    Flex.ColWidth(2) = 4000
    Flex.ColWidth(3) = 1500
    
End Sub

Sub Cargar()

    Flex.Clear
    Flex.Rows = 2
    
    Set Recordset = New ADODB.Recordset
    SQL = "SELECT c.id, c.fecha, e.nombre, c.precio FROM comidas as c "
    SQL = SQL & "INNER JOIN empleados as e on e.id = c.idempleado "
    If cboPersona.ListIndex = 1 Then
        SQL = SQL & "WHERE idempleado <> 9999 "
    ElseIf cboPersona.ListIndex = 2 Then
        SQL = SQL & "WHERE idempleado = 9999 "
    End If
    If cboPersona.ListIndex = 0 Then
        SQL = SQL & "WHERE date(c.fecha) BETWEEN '" & Format(dtDesde, "yyyy-MM-dd") & "' AND '" & Format(dtHasta, "yyyy-MM-dd") & "' "
    Else
        SQL = SQL & "AND date(c.fecha) BETWEEN '" & Format(dtDesde, "yyyy-MM-dd") & "' AND '" & Format(dtHasta, "yyyy-MM-dd") & "' "
    End If
    SQL = SQL & " ORDER BY c.fecha"
    Recordset.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not Recordset.BOF And Not Recordset.EOF Then
        Set Flex.DataSource = Recordset
    Else
        Flex.Clear
        Flex.Rows = 2
    End If
    Recordset.Close
    
    ordenaFlex
    
End Sub

Private Sub Form_Resize()
    
    If Me.ScaleHeight < 2000 Or Me.ScaleWidth < 2000 Then
        Exit Sub
    End If

    Const Margen = 120

    Flex.Width = Me.ScaleWidth - 2 * Margen
    Flex.Height = Me.ScaleHeight - picBotones.Height - dtHasta.Height - 4 * Margen

    picBotones.Left = Me.ScaleWidth - picBotones.Width - Margen
    picBotones.Top = Flex.Height + dtHasta.Height + Margen * 3

End Sub


' Instala el hook en el MSFlexGrid
Private Sub Flex_GotFocus()
    HookForm Flex
End Sub

' elimina el hook
Private Sub Flex_LostFocus()
    UnHookForm Flex
End Sub
