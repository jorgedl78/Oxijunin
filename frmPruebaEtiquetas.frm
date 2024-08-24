VERSION 5.00
Begin VB.Form frmPruebaEtiquetas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1234567898765"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblPrecio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   1
      Left            =   3600
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblPrecio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   4680
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Barras"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7440
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Barras"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   7320
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblPrecio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   8280
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   2
      Left            =   7200
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7440
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmPruebaEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Me.PrintForm
End Sub
