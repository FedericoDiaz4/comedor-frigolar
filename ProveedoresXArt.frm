VERSION 5.00
Begin VB.Form ProveedoresXArt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores Por Artículo"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
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
   ScaleHeight     =   855
   ScaleWidth      =   7455
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
      Left            =   6480
      Picture         =   "ProveedoresXArt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Salir "
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   615
      Left            =   5520
      Picture         =   "ProveedoresXArt.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Listar "
      Top             =   120
      Width           =   870
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
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Text            =   "000000"
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
      Height          =   360
      Left            =   1920
      TabIndex        =   4
      Text            =   "000000"
      Top             =   360
      Width           =   1680
   End
   Begin VB.ComboBox cboRubro 
      Height          =   360
      Left            =   3720
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Rubro"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "ProveedoresXArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboRubro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If

End Sub

Private Sub cmdListar_Click()
    
    If Not CargarProvedoresXArticulo(txtDesde.Text, txtHasta.Text, cboRubro.Text) Then
        Exit Sub
    End If
    
    SQL = "SHAPE {"
    SQL = SQL & "SELECT col1,col2,col3,col4,col5,col5,col6,CAST(REPLACE(col7,',','.') AS DECIMAL(10,3)) col7,col8,col9,col10,col11,col12,col13,col14 FROM temp_proxart_" & nTmp & ""
    SQL = SQL & " ORDER BY col1, col7}  AS temp COMPUTE temp, ANY(temp.'col1') AS col1, ANY(temp.'col2') AS col2 BY 'col1'"
    
    DataEnvironment1.Commands("temp_group").CommandType = adCmdText
    DataEnvironment1.Commands("temp_group").CommandText = SQL
    
    If DataEnvironment1.rstemp_group.State = adStateOpen Then
        DataEnvironment1.rstemp_group.Close
    End If
    DataEnvironment1.rstemp_group.Open SQL
    
    'Muestra el DataReport
    drProXArt.Sections("ReportHeader").Controls("lblDescripcion").Caption = "DESDE: " & txtDesde.Text & " - HASTA: " & txtHasta.Text
    drProXArt.Show
    
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    CargaCombo "rubros", "nombre", "nombre", cboRubro
    
End Sub

Private Sub txtDesde_GotFocus()
    
    txtDesde.SelStart = 0
    txtDesde.SelLength = Len(txtDesde.Text)
    
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

Private Sub txtHasta_GotFocus()
    
    txtHasta.SelStart = 0
    txtHasta.SelLength = Len(txtHasta.Text)

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
