VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.MDIForm zMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SISTEMA DE VENTAS"
   ClientHeight    =   6390
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9480
   Icon            =   "zMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar pBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu mParametros 
      Caption         =   " &PARÁMETROS"
      Begin VB.Menu mpEmpleados 
         Caption         =   "&EMPLEADOS"
      End
      Begin VB.Menu mpEmpresas 
         Caption         =   "&MENUS"
      End
      Begin VB.Menu mpTipos 
         Caption         =   "&TIPOS"
      End
      Begin VB.Menu mpSector 
         Caption         =   "&SECTORES"
         Visible         =   0   'False
      End
      Begin VB.Menu mpl2 
         Caption         =   "-"
      End
      Begin VB.Menu mpImportar 
         Caption         =   "&IMPORTAR"
      End
   End
   Begin VB.Menu mpControl 
      Caption         =   "&CONTROL COMENSALES"
   End
   Begin VB.Menu mcControlTickets 
      Caption         =   "&CONTROL TICKETS"
   End
   Begin VB.Menu mCompras 
      Caption         =   "&COMPRAS"
      Visible         =   0   'False
      Begin VB.Menu mcAFaltantes 
         Caption         =   "&ARTICULOS FALTANTES"
      End
      Begin VB.Menu mcOCompra 
         Caption         =   "&ORDENES DE COMPRA"
      End
      Begin VB.Menu mcLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mcCompras 
         Caption         =   "&COMPRAS        "
      End
      Begin VB.Menu mcPagos 
         Caption         =   "&PAGOS"
      End
      Begin VB.Menu mcCtaCte 
         Caption         =   "C&UENTA CORRIENTE"
      End
   End
   Begin VB.Menu mVentas 
      Caption         =   "&VENTAS"
      Visible         =   0   'False
      Begin VB.Menu mvListado 
         Caption         =   "&VENTAS"
      End
      Begin VB.Menu mvRecibos 
         Caption         =   "&COBRANZA"
         Visible         =   0   'False
      End
      Begin VB.Menu mvCtaCte 
         Caption         =   "C&UENTA CORRIENTE"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mlInformes 
      Caption         =   "&INFORME"
      Begin VB.Menu miInformeTXT 
         Caption         =   "INFORME &TXT"
      End
      Begin VB.Menu miInformeExcel 
         Caption         =   "INFORME &EXCEL DETALLE"
      End
      Begin VB.Menu miInformeExcelTotal 
         Caption         =   "&INFORME &EXCEL TOTALES"
      End
      Begin VB.Menu miInformeZeta 
         Caption         =   "INFORME ZETA"
      End
      Begin VB.Menu miInformeTotalesXDNI 
         Caption         =   "INFORME TOTALES X DNI"
      End
   End
   Begin VB.Menu mListados 
      Caption         =   "LISTADOS"
      Visible         =   0   'False
      Begin VB.Menu mlTotalesPersona 
         Caption         =   "TOTALES X PERSONA"
      End
      Begin VB.Menu mlTotalesDia 
         Caption         =   "TOTALES X DIA"
      End
   End
   Begin VB.Menu mSalir 
      Caption         =   " &SALIR"
   End
End
Attribute VB_Name = "zMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim code As String

Private Sub mcControlTickets_Click()

    TicketList.Show

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Elimina la tabla temporal
    Set rsTmp = New ADODB.Recordset
    SQL = "DROP TABLE IF EXISTS `temp_proxart_" & nTmp & "`;"
    rsTmp.Open SQL, Data, adOpenKeyset, adLockOptimistic
    
End Sub


Private Sub miInformeExcel_Click()

    InformeExcel.Show

End Sub

Private Sub miInformeExcelTotal_Click()

    informeExcelTotales.Show

End Sub

Private Sub miInformeTotalesXDNI_Click()

    informeTotalesXDNI.Show
    
End Sub

Private Sub miInformeTXT_Click()

    InformeTxt.Show

End Sub

Private Sub miInformeZeta_Click()

    InformeZeta.Show

End Sub

Private Sub mpControl_Click()
    
    Ingreso.Show
    Unload Me
    
End Sub

Private Sub mpEmpleados_Click()

    EmpleadosList.Show

End Sub

Private Sub mpEmpresas_Click()

    MenuList.Show

End Sub


Private Sub mpImportar_Click()

    ImportarExcel

End Sub

Private Sub mpSector_Click()

    SectoresList.Show

End Sub

Private Sub mpTipos_Click()

    TipoList.Show

End Sub

Private Sub mSalir_Click()
    
    Select Case MsgBox("¿DESEA SALIR DEL SISTEMA?         ", vbYesNo Or vbQuestion Or vbDefaultButton1, "Salir")
        Case vbYes: End
    End Select
    
End Sub

