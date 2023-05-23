VERSION 5.00
Begin VB.Form Tipos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   10815
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
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2535
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
      Left            =   9720
      Picture         =   "Tipos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Salir "
      Top             =   840
      Width           =   855
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
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   7815
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
      Left            =   8760
      Picture         =   "Tipos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Guardar"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Codigo"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre*"
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
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Nuevo As Boolean
Public id As Integer

Private Sub cmdGuardar_Click()
        
    If txtCodigo.Text = "" Then
        Call MsgBox("El CAMPO CODIGO ES OBLIGATORIO", vbExclamation, App.Title)
        txtNumero.SetFocus
        Exit Sub
    End If
    
    If txtNombre.Text = "" Then
        Call MsgBox("El CAMPO NOMBRE ES OBLIGATORIO", vbExclamation, App.Title)
        txtNombre.SetFocus
        Exit Sub
    End If
    
    Set rsGuardar = New ADODB.Recordset
    SQL = "SELECT * FROM Tipos WHERE id = " & id
    rsGuardar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Nuevo Then
        
        'Verificar que no se carguen dos Tipos con el mismo código
        Set rsValidar = New ADODB.Recordset
        SQL = "SELECT id from Tipos WHERE codigo = '" & txtCodigo.Text & "' LIMIT 1"
        rsValidar.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not rsValidar.BOF And Not rsValidar.EOF Then
            Call MsgBox("CODIGO EN USO", vbExclamation, App.Title)
            rsGuardar.Close
            Exit Sub
        End If
        rsValidar.Close
        
        rsGuardar.AddNew
        
    End If
    
    rsGuardar!codigo = UCase(txtCodigo.Text)
    rsGuardar!nombre = txtNombre.Text
    rsGuardar.Update
    rsGuardar.Close
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    initForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    TipoList.Show
    
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        SoloEnteros txtCodigo, KeyAscii
    End If

End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

