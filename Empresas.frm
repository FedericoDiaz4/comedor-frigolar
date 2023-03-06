VERSION 5.00
Begin VB.Form Menus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menus"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
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
   ScaleHeight     =   1605
   ScaleWidth      =   5910
   Begin VB.TextBox txtPrecio 
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
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
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
      TabIndex        =   6
      Top             =   360
      Width           =   1335
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
      Left            =   4920
      Picture         =   "Empresas.frx":0000
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
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   4215
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
      Left            =   3960
      Picture         =   "Empresas.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Guardar"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Precio"
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
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      TabIndex        =   4
      Top             =   120
      Width           =   1095
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
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "MENUS"
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
    
    If txtCodigo.Text = "" Then
        Call MsgBox("El CAMPO CODIGO ES OBLIGATORIO", vbExclamation, App.Title)
        txtNombre.SetFocus
        Exit Sub
    End If
    
    If txtPrecio.Text = "" Then
        Call MsgBox("El CAMPO PRECIO ES OBLIGATORIO", vbExclamation, App.Title)
        txtNombre.SetFocus
        Exit Sub
    End If
    
    Set rsGuardar = New ADODB.Recordset
    SQL = "SELECT * FROM menus WHERE id = " & id
    rsGuardar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Nuevo Then
        
        'Verificar que no se carguen dos menus con el mismo código
        Set rsValidar = New ADODB.Recordset
        SQL = "SELECT id from menus WHERE nombre = '" & txtNombre.Text & "' AND eliminado = 0 LIMIT 1"
        rsValidar.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not rsValidar.BOF And Not rsValidar.EOF Then
            Call MsgBox("YA EXISTE UNA MENU CON EL MISMO NOMBRE   ", vbExclamation, App.Title)
            rsGuardar.Close
            Exit Sub
        End If
        rsValidar.Close
        
        rsGuardar.AddNew
        
    End If
    
    rsGuardar!nombre = txtNombre.Text
    rsGuardar!codigo = txtCodigo.Text
    rsGuardar!Precio = Format(txtPrecio.Text, "0.00")
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
    
    MenuList.Show
    
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

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        SoloEnteros txtPrecio, KeyAscii
    End If
    
End Sub
