VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form InformeTxt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe txt"
   ClientHeight    =   1560
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
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
   ScaleHeight     =   1560
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
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
      Left            =   3480
      Picture         =   "Detalle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Salir "
      Top             =   840
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   360
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
      Format          =   117964801
      CurrentDate     =   43717
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
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
      Format          =   117964801
      CurrentDate     =   43717
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "E&xportar"
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
      Left            =   2520
      Picture         =   "Detalle.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Exportar "
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblHasta 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblFechaDesde 
      Caption         =   "Desde"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "InformeTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()

    Dim empresa As String
    Dim idturno As Integer
    Dim desde As Date
    Dim hasta As Date

    desde = dtpDesde.Value
    hasta = dtpHasta.Value
    
    If existeArchivo("C:\sistema\expo.txt") Then
        Kill "C:\sistema\expo.txt"
    End If
    
    Open "C:\sistema\expo.txt" For Output As #1
    
    Set rsconsulta = New ADODB.Recordset
    SQL = "SELECT c.id, c.fecha, c.numero, c.precio, e.nrolegajo, e.nombre FROM comidas AS c "
    SQL = SQL & "INNER JOIN empleados AS e ON e.id = c.idempleado "
    SQL = SQL & "WHERE date(fecha) BETWEEN '" & Format(dtpDesde.Value, "yyyy-mm-dd") & "' AND '" & Format(dtpHasta.Value, "yyyy-mm-dd") & "' "
    SQL = SQL & "AND e.nombre <> 'CONSUMIDOR FINAL' "
    rsconsulta.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    If Not rsconsulta.EOF And Not rsconsulta.BOF Then
        Do While Not rsconsulta.EOF
                Precio = rsconsulta!Precio * 100
                Print #1, Format(rsconsulta!numero, "00000000") & Format(rsconsulta!nrolegajo, "000000") & Format(rsconsulta!fecha, "ddmmyyyy") & Format(rsconsulta!fecha, "HHMMSS") & Format(Precio, "00000000")
                rsconsulta.MoveNext
        Loop
        Close #1
    End If
    
    MsgBox ("Finalizado!")
    
    'Exportar_Excel "C:\Listado.xls", empresa, idturno, desde, hasta

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.Width = 4590
    Me.Height = 2955
    initForm Me
    
    If Day(Now) < 16 Then
        dtpDesde.Value = DateSerial(Year(Now), Month(Now), 1)
        dtpHasta.Value = DateSerial(Year(Now), Month(Now), 15)
    Else
        dtpDesde.Value = DateSerial(Year(Now), Month(Now), 16)
        dtpHasta.Value = DateSerial(Year(Now), Month(Now) + 1, 0)
    End If
    
    
End Sub
