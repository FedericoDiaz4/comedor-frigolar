VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ProveedorPredeterminado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Provedor Predeterminado"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   9.75
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
   ScaleHeight     =   1680
   ScaleWidth      =   9840
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   9
      Top             =   960
      Width           =   1815
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
         Left            =   0
         Picture         =   "ProveedorPredeterminado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   " Establecer el proveedor seleccionado como predeterminado."
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
         Left            =   960
         Picture         =   "ProveedorPredeterminado.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Salir "
         Top             =   0
         Width           =   855
      End
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
      TabIndex        =   5
      Top             =   360
      Width           =   1680
   End
   Begin VB.ComboBox cboProveedor 
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
      Height          =   360
      ItemData        =   "ProveedorPredeterminado.frx":0B14
      Left            =   1920
      List            =   "ProveedorPredeterminado.frx":0B16
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox txtDesde 
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
      Left            =   6240
      TabIndex        =   7
      Top             =   360
      Width           =   1680
   End
   Begin VB.TextBox txtHasta 
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
      Left            =   8040
      TabIndex        =   8
      Top             =   360
      Width           =   1680
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label12 
      Caption         =   "Código Proveedor"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Proveedor"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Artículo Desde"
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
      Left            =   6240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Artículo Hasta"
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
      Left            =   8040
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "ProveedorPredeterminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboProveedor_Click()
    
    If cboProveedor.ListIndex <> -1 Then
        txtCodProveedor.Text = getData(cboProveedor.ItemData(cboProveedor.ListIndex), "codigo", "proveedores")
    End If
    
End Sub

Private Sub cboProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmdGuardar_Click()
    
    If cboProveedor.ListIndex = -1 Then
        Call MsgBox("DEBE SELECCIONAR UN PROVEEDOR     ", vbExclamation Or vbDefaultButton1, App.Title)
        Exit Sub
    End If
    
    Dim idProveedor As Single
    idProveedor = cboProveedor.ItemData(cboProveedor.ListIndex)
    
    Dim Cantidad As Integer
    Set rsCantidad = New ADODB.Recordset
    SQL = "SELECT COUNT(c.idart) AS total FROM articulosc AS c Inner Join articulos AS a ON c.idart = a.id "
    SQL = SQL & "WHERE c.idpro = " & idProveedor & " "
    If txtDesde.Text <> "" And txtHasta.Text <> "" Then
        SQL = SQL & "AND a.codigo BETWEEN '" & txtDesde.Text & "' AND '" & txtHasta.Text & "' "
    End If
    SQL = SQL & "ORDER BY a.codigo;"
    rsCantidad.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsCantidad.BOF And Not rsCantidad.EOF Then
        Cantidad = rsCantidad!Total
    End If
    rsCantidad.Close
    
    Select Case MsgBox("ESTÁ POR MODIFICAR EL PROVEEDOR PREDETERMINADO DE " & Cantidad & " ARTÍCULOS    " _
                       & vbCrLf & "LA OPERACIÓN ES IRREVERSIBLE, ¿DESEA CONTINUAR?" _
                       , vbYesNo Or vbQuestion Or vbDefaultButton2, App.Title)
    
        Case vbNo: Exit Sub
        
    End Select
    
    Set rsArtC = New ADODB.Recordset
    SQL = "SELECT c.idart FROM articulosc AS c Inner Join articulos AS a ON c.idart = a.id "
    SQL = SQL & "WHERE c.idpro = " & idProveedor & " "
    If txtDesde.Text <> "" And txtHasta.Text <> "" Then
        SQL = SQL & "AND a.codigo BETWEEN '" & txtDesde.Text & "' AND '" & txtHasta.Text & "' "
    End If
    SQL = SQL & "ORDER BY a.codigo;"
    rsArtC.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    pBar.Max = Cantidad
    
    Do While Not rsArtC.EOF
        establecerProveedorPredeterminado idProveedor, rsArtC!idArt
        pBar.Value = pBar.Value + 1
        rsArtC.MoveNext
    Loop
    
    pBar.Value = 0
    rsArtC.Close
    
    Call MsgBox("PROCESO FINALIZADO    ", vbInformation, App.Title)
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    
    'Carga los Combos
    CargaCombo "proveedores", "nombre", "nombre", cboProveedor
    
End Sub

Private Sub txtCodProveedor_Change()
    
    VerificarConexion
    
    Set rsPro = New ADODB.Recordset
    SQL = "SELECT nombre FROM proveedores WHERE codigo = '" & txtCodProveedor.Text & "' AND eliminado <> 1"
    rsPro.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsPro.BOF And Not rsPro.EOF Then
        If rsPro!nombre <> "" Then
            cboProveedor.Text = rsPro!nombre
        End If
        rsPro.Close
    Else
        cboProveedor.ListIndex = -1
    End If
    
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDesde_LostFocus()
    
    If txtDesde.Text <> "" Then
        txtDesde.Text = Format(txtDesde.Text, "000000")
    End If
    
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtHasta_LostFocus()
    
    If txtHasta.Text <> "" Then
        txtHasta.Text = Format(txtHasta.Text, "000000")
    End If
    
End Sub
