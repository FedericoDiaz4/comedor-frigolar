VERSION 5.00
Begin VB.Form Cantidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cantidad"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2820
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
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2820
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
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
      Left            =   1800
      Picture         =   "Cantidad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Salir "
      Top             =   840
      Width           =   855
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
      Left            =   840
      Picture         =   "Cantidad.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Guardar"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Cantidad"
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
End
Attribute VB_Name = "Cantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGuardar_Click()

    If txtCantidad.Text <> "" Then
        Ticket.cantidadArt = CInt(txtCantidad.Text)
    Else
        Ticket.cantidadArt = 0
    End If
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()

    Ticket.cantidadArt = 0
    Unload Me

End Sub

Private Sub Form_Load()
    
    initForm Me
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdGuardar_Click
    Else
        SoloEnteros txtCantidad, KeyAscii
    End If

End Sub
