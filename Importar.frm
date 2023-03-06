VERSION 5.00
Begin VB.Form Importar 
   Caption         =   "Importar"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   3150
   Begin VB.ComboBox cboImportar 
      Height          =   315
      ItemData        =   "Importar.frx":0000
      Left            =   120
      List            =   "Importar.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "&Importar"
      Height          =   615
      Left            =   2160
      Picture         =   "Importar.frx":0023
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " Exportar "
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblImportar 
      Caption         =   "Tabla a Importar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Importar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


