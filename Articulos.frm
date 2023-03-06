VERSION 5.00
Begin VB.Form Articulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
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
   ScaleHeight     =   1815
   ScaleWidth      =   8055
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
      Left            =   2760
      TabIndex        =   7
      Text            =   "0,00"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox cboIVA 
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtNombre 
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
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   5175
   End
   Begin VB.TextBox txtCodigo 
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
      TabIndex        =   2
      Top             =   360
      Width           =   2535
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
      Left            =   6120
      Picture         =   "Articulos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Guardar"
      Top             =   1080
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
      Left            =   7080
      Picture         =   "Articulos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Salir "
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Precio"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Descripción*"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Código*"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "IVA %"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Nuevo As Boolean
Public ID As Single

Private Sub cboIVA_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
    
    Call MsgBox("Codigo ya existe", vbCritical, "Codigo Repetido")
    
End Sub

Private Sub cmdGuardar_Click()
    
    If txtCodigo.Text = "" Then
        Call MsgBox("El CAMPO CÓDIGO ES OBLIGATORIO", vbExclamation, App.Title)
        txtCodigo.SetFocus
        Exit Sub
    End If
    
    If txtNombre.Text = "" Then
        Call MsgBox("El CAMPO DESCRIPCIÓN ES OBLIGATORIO", vbExclamation, App.Title)
        txtNombre.SetFocus
        Exit Sub
    End If
    
    Set rsGuardar = New ADODB.Recordset
    SQL = "SELECT * FROM articulos WHERE id = " & ID
    rsGuardar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Nuevo = True Then
        
        'Verificar que no se carguen dos artículos con el mismo código
        Set rsValidar = New ADODB.Recordset
        SQL = "SELECT codigo from articulos WHERE codigo = '" & txtCodigo.Text & "' AND eliminado = 0 LIMIT 1"
        rsValidar.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not rsValidar.BOF And Not rsValidar.EOF Then
            Call MsgBox("YA EXISTE UN ARTÍCULO CON EL MISMO CÓDIGO    ", vbExclamation, App.Title)
            Exit Sub
        End If
        rsValidar.Close
        
        rsGuardar.AddNew
        
    End If
    
    rsGuardar!codigo = txtCodigo.Text
    rsGuardar!nombre = txtNombre.Text
    rsGuardar!idTipoIVA = cboIVA.ItemData(cboIVA.ListIndex)
    rsGuardar!Precio = Format(txtPrecio.Text, "0.00")
    rsGuardar.Update
    rsGuardar.Close
        
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    
    VerificarConexion
    
    'Carga el combo IVA
    CargaCombo "tipoiva", "nombre", "id", cboIVA
    cboIVA.ListIndex = 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ArticulosList.Show
    
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        CambiaPunto txtPrecio, KeyAscii
    End If
    
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub
