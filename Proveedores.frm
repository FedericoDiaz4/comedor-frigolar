VERSION 5.00
Begin VB.Form Proveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
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
   ScaleHeight     =   6615
   ScaleWidth      =   8055
   Begin VB.TextBox txtFPago 
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
      TabIndex        =   38
      Top             =   5400
      Width           =   5175
   End
   Begin VB.TextBox txtCelular 
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
      TabIndex        =   17
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtFax 
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
      TabIndex        =   18
      Top             =   2520
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
      Left            =   7080
      Picture         =   "Proveedores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   " Salir "
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox txtCotizacionEuro 
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
      Left            =   5400
      TabIndex        =   39
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox txtCPago 
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
      Left            =   5400
      TabIndex        =   31
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox txtObs 
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
      TabIndex        =   34
      Top             =   4680
      Width           =   5175
   End
   Begin VB.TextBox txtEmail2 
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
      TabIndex        =   30
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox txtEmail1 
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
      TabIndex        =   29
      Top             =   3960
      Width           =   2535
   End
   Begin VB.ComboBox cboIVA 
      Height          =   360
      ItemData        =   "Proveedores.frx":058A
      Left            =   5400
      List            =   "Proveedores.frx":058C
      TabIndex        =   25
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtCP 
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
      Left            =   5400
      TabIndex        =   19
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtContacto 
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
      Left            =   5400
      TabIndex        =   13
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtTelefono2 
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
      TabIndex        =   12
      Top             =   1800
      Width           =   2535
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
   Begin VB.TextBox txtCotizacionDolar 
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
      Left            =   5400
      TabIndex        =   35
      Top             =   4680
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
      Picture         =   "Proveedores.frx":058E
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   " Guardar"
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox txtTelefono1 
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
      TabIndex        =   11
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtLocalidad 
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
      TabIndex        =   23
      Top             =   3240
      Width           =   2535
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
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5175
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
   Begin VB.TextBox txtCUIT 
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
      Left            =   5400
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox cboProvincia 
      Height          =   360
      ItemData        =   "Proveedores.frx":0B18
      Left            =   2760
      List            =   "Proveedores.frx":0B64
      TabIndex        =   24
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label20 
      Caption         =   "Celular"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "Fax"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "Cotización Euro (€)"
      DataField       =   " "
      Height          =   255
      Left            =   5400
      TabIndex        =   37
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "Forma Pago"
      DataField       =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Cotización Dólar (U$S)"
      DataField       =   " "
      Height          =   255
      Left            =   5400
      TabIndex        =   33
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Observaciones"
      DataField       =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Email 2"
      DataField       =   " "
      Height          =   255
      Left            =   2760
      TabIndex        =   27
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Email 1"
      DataField       =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Tipo IVA"
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Código Postal"
      DataField       =   " "
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Contacto"
      DataField       =   " "
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Teléfono 2"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Código*"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Cond. Pago"
      DataField       =   " "
      Height          =   255
      Left            =   5400
      TabIndex        =   28
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Teléfono 1"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Localidad"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre*"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "CUIT"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   3000
      Width           =   1335
   End
End
Attribute VB_Name = "Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Nuevo As Boolean
Public ID As Integer
Public idProvincia As Integer

Private Sub cboIVA_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub cboProvincia_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub cmdGuardar_Click()
    
    Set rsGuardar = New ADODB.Recordset
    SQL = "SELECT * FROM proveedores WHERE id = " & ID
    rsGuardar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Nuevo = True Then
        
        'Verificar que no se carguen dos clientes con el mismo CUIT
        Set rsValidar = New ADODB.Recordset
        SQL = "SELECT id from proveedores WHERE cuit = '" & txtCUIT.Text & "' AND eliminado = 0 LIMIT 1"
        rsValidar.Open SQL, Data, adOpenKeyset, adLockOptimistic
        If Not rsValidar.BOF And Not rsValidar.EOF And txtCUIT.Text <> "" And txtCUIT.Text <> "  -        -" Then
            Call MsgBox("YA EXISTE UN PROVEEDOR CON EL MISMO CUIT   ", vbExclamation, App.Title)
            rsGuardar.Close
            Exit Sub
        End If
        rsValidar.Close
    
        rsGuardar.AddNew
    End If
    
    rsGuardar!ID = ID
    rsGuardar!codigo = Format(txtCodigo.Text, "0000")
    rsGuardar!nombre = txtNombre.Text
    rsGuardar!domicilio = txtDireccion.Text
    rsGuardar!cuit = txtCUIT.Text
    rsGuardar!telefono = txtTelefono1.Text
    rsGuardar!tel1 = txtTelefono2.Text
    rsGuardar!celular = txtCelular.Text
    rsGuardar!fax = txtFax.Text
    rsGuardar!contacto = txtContacto.Text
    rsGuardar!provincia = cboProvincia.Text
    rsGuardar!localidad = txtLocalidad.Text
    rsGuardar!cp = txtCP.Text
    rsGuardar!tipoiva = cboIVA.Text
    rsGuardar!email = txtEmail1.Text
    rsGuardar!email2 = txtEmail2.Text
    rsGuardar!cotizaciondolar = txtCotizacionDolar.Text
    rsGuardar!cotizacioneuro = txtCotizacionEuro.Text
    rsGuardar!tel2 = txtObs.Text
    rsGuardar!cpago = txtCPago.Text
    rsGuardar!fpago = txtFPago.Text
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
    
    'Carga Provincias
    Set rsProvincia = New ADODB.Recordset
    SQL = "SELECT DISTINCT(provincia) FROM proveedores ORDER BY provincia"
    rsProvincia.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not rsProvincia.EOF
        cboProvincia.AddItem rsProvincia!provincia
        rsProvincia.MoveNext
    Loop
    rsProvincia.Close
    
    'Carga Tipos de IVA
    CargaCombo "tipoiva", "descripcion", "descripcion", cboIVA
    
    If Nuevo = True Then
        Set rsIndice = New ADODB.Recordset
        SQL = "SELECT codigo FROM proveedores WHERE eliminado = 0 ORDER BY codigo DESC LIMIT 1"
        rsIndice.Open SQL, Data, adOpenKeyset, adLockOptimistic
        txtCodigo = CInt(rsIndice!codigo) + 1
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ProveedoresList.Show
    
End Sub

Private Sub txtCelular_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtContacto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If

End Sub

Private Sub txtCotizacionDolar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    Else
        CambiaPunto txtCotizacionDolar, KeyAscii
    End If
    
End Sub

Private Sub txtCotizacionEuro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    Else
        CambiaPunto txtCotizacionEuro, KeyAscii
    End If
    
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtCPago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtCUIT_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtEmail1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtEmail2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If

End Sub

Private Sub txtFPago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtTelefono1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If
    
End Sub

Private Sub txtTelefono2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        PasarFoco
    End If

End Sub
