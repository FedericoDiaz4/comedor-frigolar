VERSION 5.00
Begin VB.Form Sectores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sectores"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
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
   ScaleHeight     =   1575
   ScaleWidth      =   5535
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
      Left            =   4560
      Picture         =   "Sectores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5295
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
      Left            =   3600
      Picture         =   "Sectores.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Guardar"
      Top             =   840
      Width           =   855
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Sectores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Nuevo As Boolean
Public id As Integer

Private Sub cmdGuardar_Click()
        
    If txtNombre.Text = "" Then
        Call MsgBox("El CAMPO NOMBRE ES OBLIGATORIO", vbExclamation, App.Title)
        txtNombre.SetFocus
        Exit Sub
    End If
    
    Set rsGuardar = New ADODB.Recordset
    SQL = "SELECT * FROM sectores WHERE id = " & id
    rsGuardar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Nuevo Then
        
        'Verificar que no se carguen dos empresas con el mismo código
        Set rsValidar = New ADODB.Recordset
        SQL = "SELECT id from sectores WHERE nombre = '" & txtNombre.Text & "' AND eliminado = 0 LIMIT 1"
        rsValidar.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not rsValidar.BOF And Not rsValidar.EOF Then
            Call MsgBox("YA EXISTE UN SECTOR CON EL MISMO NOMBRE   ", vbExclamation, App.Title)
            rsGuardar.Close
            Exit Sub
        End If
        rsValidar.Close
        
        rsGuardar.AddNew
        
    End If
    
    rsGuardar!nombre = txtNombre.Text
    rsGuardar!eliminado = 0
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
    
    SectoresList.Show
    
End Sub



Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub
