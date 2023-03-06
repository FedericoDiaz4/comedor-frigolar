VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ArticulosPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos por Proveedor"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13095
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
   ScaleHeight     =   5400
   ScaleWidth      =   13095
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Text            =   "CODIGO"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   4920
      TabIndex        =   5
      Text            =   "IVA"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   4
      Text            =   "DESCUENTOS"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   3
      Text            =   "PRECIO"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   12120
      Picture         =   "ArticulosPro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Salir "
      Top             =   4680
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   15658734
      ForeColor       =   0
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   12640511
      ForeColorFixed  =   0
      ForeColorSel    =   16777215
      BackColorBkg    =   8421504
      BackColorUnpopulated=   12632256
      GridColor       =   0
      GridColorFixed  =   0
      GridColorUnpopulated=   8421504
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.Label lblProveedor 
      Caption         =   "PROVEEDOR"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "ArticulosPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idProveedor As Integer
Dim rsCargar As ADODB.Recordset

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Sub Flex_Click()
    
    'Posiciona los TextBox
    For i = 3 To 6
        txtEdit(i).Visible = False
        txtEdit(i).Top = Flex.Top + Flex.CellTop
        txtEdit(i).Left = Flex.ColPos(i) + 175
        txtEdit(i).Width = Flex.ColWidth(i) - 60
        txtEdit(i).Text = Flex.TextMatrix(Flex.Row, i)
    Next i
    
    'Le da el foco al correspondiente
    If 3 <= Flex.Col And Flex.Col <= 6 Then
        txtEdit(Flex.Col).Visible = True
        txtEdit(Flex.Col).SetFocus
    End If
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    Cargar
    
End Sub

Sub OrdenaFlex()
    
    Flex.FormatString = "id|Código|Artículo|Código|Precio|Descuentos|IVA|Total"
    
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 1000
    Flex.ColWidth(2) = 5200
    Flex.ColWidth(3) = 1400
    Flex.ColWidth(4) = 1200
    Flex.ColWidth(5) = 1300
    Flex.ColWidth(6) = 1200
    Flex.ColWidth(7) = 1200
    
    Flex.ColAlignment(3) = 2
    
End Sub

Sub Cargar()
    
    'Muestra el resultado en el flex
    Set rsCargar = New ADODB.Recordset
    SQL = "SELECT c.id, a.codigo, a.nombre, c.codigo, c.precio, c.descuento, c.iva, c.total FROM articulosc AS c Inner Join articulos AS a ON c.idart = a.id AND C.idpro = " & idProveedor & " "
    SQL = SQL & "ORDER BY c.codigo"
    rsCargar.Open SQL, Data, adOpenKeyset, adLockOptimistic
    If Not rsCargar.BOF And Not rsCargar.EOF Then
        Set Flex.DataSource = rsCargar
        
        'Color de celdas
        For i = 1 To Flex.Rows - 1
            For a = 3 To 6
                Flex.Row = i
                Flex.Col = a
                Flex.CellBackColor = vbWhite
            Next a
        Next i
        
    Else
        Flex.Clear
        Flex.Rows = 2
    End If
    rsCargar.Close
    OrdenaFlex
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error Resume Next
    rsCargar.Close
    
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    
    txtEdit(Index).SelStart = 0
    txtEdit(Index).SelLength = Len(txtEdit(Index).Text)
    
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        Case vbKeyDown
            ActualizarReg Index
            If Flex.Row < Flex.Rows - 1 Then
                Flex.Row = Flex.Row + 1
                Flex.Col = Index
                Flex_Click
            End If
            
        Case vbKeyUp
            ActualizarReg Index
            If Flex.Row > 1 Then
                Flex.Row = Flex.Row - 1
                Flex.Col = Index
                Flex_Click
            End If
        
        Case vbKeyLeft
            ActualizarReg Index
            If Flex.Col = 3 And Flex.Row > 1 Then
                Flex.Row = Flex.Row - 1
                Flex.Col = 6
            Else
                Flex.Col = Flex.Col - 1
            End If
            Flex_Click
        
        Case vbKeyRight, 13
            ActualizarReg Index
            If Flex.Col = 6 And Flex.Row < Flex.Rows - 1 Then
                Flex.Row = Flex.Row + 1
                Flex.Col = 3
            Else
                Flex.Col = Flex.Col + 1
            End If
            Flex_Click
            
    End Select
    
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index <> 3 Then
        
        CambiaPunto txtEdit(Index), KeyAscii, "+"
        
    End If
    
End Sub

Sub ActualizarReg(Index As Integer)
    
    Dim idArtC As Long
    
    'Guarda el ID
    idArtC = Flex.TextMatrix(Flex.Row, 0)
    
    'Sólo actualiza si se modifica el TextBox
    If txtEdit(Index).Text = Flex.TextMatrix(Flex.Row, Flex.Col) Then
        Exit Sub
    End If
    
    'Calcula el total
    If Index <> 3 Then
        If txtEdit(4).Text <> "" And txtEdit(5).Text <> "" And txtEdit(6).Text <> "" Then
            Tot = CalculaDescuentos(txtEdit(4).Text, txtEdit(5).Text)
            Tot = Format(Tot + ((Tot * CDec(txtEdit(6).Text)) / 100), "0.00")
        Else
            Tot = "0,00"
        End If
    Else
        Tot = Flex.TextMatrix(Flex.Row, 7)
    End If
    
    'Actualiza la tabla de la base
    Set rsUpd = New ADODB.Recordset
    SQL = "UPDATE articulosc SET total = '" & Tot & "', "
    
    Select Case Index
    Case 3: SQL = SQL & "codigo = '" & txtEdit(Index).Text & "' "
    Case 4: SQL = SQL & "precio = '" & txtEdit(Index).Text & "' "
    Case 5: SQL = SQL & "descuento = '" & txtEdit(Index).Text & "' "
    Case 6: SQL = SQL & "iva = '" & txtEdit(Index).Text & "' "
    End Select
    
    SQL = SQL & "WHERE id = " & idArtC & ";"
    rsUpd.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Actualiza la información del flex
    Flex.TextMatrix(Flex.Row, Flex.Col) = txtEdit(Index).Text
    Flex.TextMatrix(Flex.Row, 7) = Tot
    
End Sub
