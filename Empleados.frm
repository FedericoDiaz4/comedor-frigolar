VERSION 5.00
Begin VB.Form Empleados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleados"
   ClientHeight    =   2355
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
   ScaleHeight     =   2355
   ScaleWidth      =   10815
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
      ItemData        =   "Empleados.frx":0000
      Left            =   8040
      List            =   "Empleados.frx":004C
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1080
      Width           =   2535
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
      ItemData        =   "Empleados.frx":0161
      Left            =   4800
      List            =   "Empleados.frx":01AD
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtCuil 
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
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtNroLegajo 
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
      Picture         =   "Empleados.frx":02C2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Salir "
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtNumero 
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
      Picture         =   "Empleados.frx":084C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Guardar"
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Left            =   8040
      TabIndex        =   12
      Top             =   840
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
      Left            =   4800
      TabIndex        =   11
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Cuil"
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
      TabIndex        =   5
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label13 
      Caption         =   "Nro Legajo"
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
   Begin VB.Label Label5 
      Caption         =   "Numero Documento"
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
      Top             =   840
      Width           =   2535
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
Attribute VB_Name = "Empleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Nuevo As Boolean
Public id As Integer

Private Sub cmdGuardar_Click()
        
    If txtNumero.Text = "" Then
        Call MsgBox("El CAMPO DOCUMENTO ES OBLIGATORIO", vbExclamation, App.Title)
        txtNumero.SetFocus
        Exit Sub
    End If
    
    If txtNombre.Text = "" Then
        Call MsgBox("El CAMPO DESCRIPCIÓN ES OBLIGATORIO", vbExclamation, App.Title)
        txtNombre.SetFocus
        Exit Sub
    End If
    
    Set rsGuardar = New ADODB.Recordset
    SQL = "SELECT * FROM empleados WHERE id = " & id
    rsGuardar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Nuevo Then
        
        'Verificar que no se carguen dos empleados con el mismo código
        Set rsValidar = New ADODB.Recordset
        SQL = "SELECT id from empleados WHERE numerodocumento = '" & txtNumero.Text & "' AND eliminado = 0 LIMIT 1"
        rsValidar.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not rsValidar.BOF And Not rsValidar.EOF Then
            Call MsgBox("YA EXISTE UN EMPLEADO CON EL MISMO DNI   ", vbExclamation, App.Title)
            rsGuardar.Close
            Exit Sub
        End If
        rsValidar.Close
        
        rsGuardar.AddNew
        
    End If
    
    rsGuardar!nombre = UCase(txtNombre.Text)
    rsGuardar!numerodocumento = txtNumero.Text
    rsGuardar!nrolegajo = txtNroLegajo.Text
    rsGuardar!Cuil = txtCuil.Text
    rsGuardar!idTipo = cboTipos.ItemData(cboTipos.ListIndex)
    rsGuardar!idEmpresa = CboEmpresa.ItemData(CboEmpresa.ListIndex)
    rsGuardar!eliminado = 0
    rsGuardar.Update
    rsGuardar.Close
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    CargaCombo "tipos", "nombre", "id", cboTipos
    CargaCombo "empresas", "nombre", "id", CboEmpresa
    initForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    EmpleadosList.Show
    
End Sub

Private Sub txtCuil_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        SoloEnteros txtCuil, KeyAscii
    End If

End Sub

Private Sub txtNroLegajo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        SoloEnteros txtNroLegajo, KeyAscii
    End If

End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    Else
        SoloEnteros txtNumero, KeyAscii
    End If

End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

Private Sub cboTipos_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
        KeyAscii = 0
    End If
    
End Sub

