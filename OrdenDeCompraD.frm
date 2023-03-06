VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form OrdenDeCompraD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Compra"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
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
   ScaleHeight     =   4935
   ScaleWidth      =   8895
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
      Left            =   7920
      Picture         =   "OrdenDeCompraD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Salir "
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "&Finalizar"
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
      Left            =   6960
      Picture         =   "OrdenDeCompraD.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Finalizar "
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtCantidad 
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
      Height          =   330
      Left            =   7080
      TabIndex        =   2
      Top             =   840
      Width           =   1680
   End
   Begin VB.TextBox txtArticulo 
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
      Enabled         =   0   'False
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6840
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   2880
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5080
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   3
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
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label lblOrden 
      Alignment       =   1  'Right Justify
      Caption         =   "Nº Orden: 1234"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line 
      X1              =   0
      X2              =   8880
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label9 
      Caption         =   "Artículo"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   7080
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblTitulo 
      Caption         =   "CMsis Informática"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "OrdenDeCompraD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IdOC As Single
Dim idArt As Single

Private Sub cmdFinalizar_Click()
    
    Set rsOC = New ADODB.Recordset
    SQL = "UPDATE ocompras SET estado = 'FINALIZADO' WHERE id = " & IdOC
    rsOC.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    Set rsOCD = New ADODB.Recordset
    SQL = "UPDATE ocomprasd SET estado = 'FINALIZADO' WHERE idorden = " & IdOC
    rsOCD.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
    'Agrega a faltantes los artículos que no se hayan entregado
    Set rsOCD = New ADODB.Recordset
    SQL = "SELECT * FROM ocomprasd WHERE idorden = " & IdOC
    rsOCD.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Do While Not rsOCD.BOF And Not rsOCD.EOF
        
        For i = 1 To Flex.Rows - 1
            
            If (Flex.TextMatrix(i, 2) = rsOCD!idArt) And (Flex.TextMatrix(i, 4) < rsOCD!Cantidad) Then
                
                Cantidad = rsOCD!Cantidad - Flex.TextMatrix(i, 4)
                
                'Lo agrega a la tabla articulosfaltantes
                Set rsFaltante = New ADODB.Recordset
                SQL = "SELECT * FROM articulosfaltantes WHERE idart = " & rsOCD!idArt
                rsFaltante.Open SQL, Data, adOpenKeyset, adLockOptimistic
                
                If rsFaltante.BOF And rsFaltante.EOF Then
                    rsFaltante.AddNew
                    rsFaltante!idArt = rsOCD!idArt
                    rsFaltante!idPro = getProveedorRecomendado(rsOCD!idArt) 'le asigna el proveedor recomendado
                End If
                
                rsFaltante!Cantidad = Cantidad
                rsFaltante.Update
                rsFaltante.Close
                
            End If
        Next
        
        'Aca deberia actualizar la fecha de los articulos comprados

        rsOCD.MoveNext
    Loop
    rsOCD.Close
    
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Flex_Click()
    
    If Flex.TextMatrix(Flex.Row, 0) = "" Then Exit Sub
    idArt = Flex.TextMatrix(Flex.Row, 2)
    txtArticulo.Text = Flex.TextMatrix(Flex.Row, 3)
    txtCantidad.Text = Flex.TextMatrix(Flex.Row, 4)
    txtCantidad.SetFocus
    
End Sub

Private Sub Form_Load()
    
    initForm Me
    
End Sub

Sub Cargar(ID As Single)
    
    IdOC = ID
    
    Set rsOC = New ADODB.Recordset
    SQL = "SELECT id, idpro FROM ocompras WHERE id = " & ID & ";"
    rsOC.Open SQL, Data, adOpenKeyset, adLockOptimistic
    lblOrden.Caption = "Nº Orden " & Format(ID, "0000")
    lblTitulo.Caption = getData(rsOC!idPro, "nombre", "proveedores")
    rsOC.Close
    
    Set rsOCD = New ADODB.Recordset
    SQL = "SELECT * FROM ocomprasd WHERE idorden = " & ID & " ORDER BY id;"
    rsOCD.Open SQL, Data, adOpenKeyset, adLockOptimistic
    Set Flex.DataSource = rsOCD
    OrdenaFlex
    rsOCD.Close
    
End Sub

Private Sub OrdenaFlex()
    
    Flex.FormatString = "id|idorden|idart|Artículo|Cantidad|Estado"
    Flex.ColWidth(0) = 0
    Flex.ColWidth(1) = 0
    Flex.ColWidth(2) = 0
    Flex.ColWidth(3) = 6800
    Flex.ColWidth(4) = 1300
    Flex.ColWidth(5) = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    OrdenDeCompra.Show
    
End Sub

Private Sub txtCantidad_GotFocus()
    
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Add
    Else
        CambiaPunto txtCantidad, KeyAscii
    End If
    
End Sub

Sub Add()
    
    If txtCantidad.Text = "" Then txtCantidad.Text = "0"
    For i = 1 To Flex.Rows - 1
        If idArt = Flex.TextMatrix(i, 2) Then
            Flex.TextMatrix(i, 4) = txtCantidad.Text
        End If
    Next
    
    idArt = 0
    txtArticulo.Text = ""
    txtCantidad.Text = ""
    Flex.SetFocus
    
End Sub
