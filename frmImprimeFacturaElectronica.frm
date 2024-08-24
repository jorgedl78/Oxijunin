VERSION 5.00
Begin VB.Form frmImprimeFacturaElectronica 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   16005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   16005
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line20 
      X1              =   120
      X2              =   11040
      Y1              =   12840
      Y2              =   12840
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TUBOS DEVUELTOS"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   126
      Top             =   11520
      Width           =   1935
   End
   Begin VB.Label lblDetalleDevueltos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      TabIndex        =   125
      Top             =   11880
      Width           =   10695
   End
   Begin VB.Label lblIva105 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Iva"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   124
      Top             =   14400
      Width           =   855
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "IVA 10.5%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   123
      Top             =   13920
      Width           =   1095
   End
   Begin VB.Line Line19 
      X1              =   7440
      X2              =   7440
      Y1              =   13800
      Y2              =   14280
   End
   Begin VB.Line Line18 
      X1              =   7440
      X2              =   7440
      Y1              =   14280
      Y2              =   14760
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   7080
      TabIndex        =   122
      Top             =   9120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   8760
      TabIndex        =   121
      Top             =   9120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   1800
      TabIndex        =   120
      Top             =   9120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   240
      TabIndex        =   119
      Top             =   9120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   7080
      TabIndex        =   118
      Top             =   8880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   8760
      TabIndex        =   117
      Top             =   8880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   1800
      TabIndex        =   116
      Top             =   8880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   240
      TabIndex        =   115
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   7080
      TabIndex        =   114
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   8760
      TabIndex        =   113
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   1800
      TabIndex        =   112
      Top             =   8640
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   111
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   7080
      TabIndex        =   110
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   8760
      TabIndex        =   109
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   1800
      TabIndex        =   108
      Top             =   8400
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   107
      Top             =   8400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   7080
      TabIndex        =   106
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   8760
      TabIndex        =   105
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   1800
      TabIndex        =   104
      Top             =   8160
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   103
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   7080
      TabIndex        =   102
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   8760
      TabIndex        =   101
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   1800
      TabIndex        =   100
      Top             =   7920
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   99
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   7080
      TabIndex        =   98
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   8760
      TabIndex        =   97
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   1800
      TabIndex        =   96
      Top             =   7680
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   95
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   7080
      TabIndex        =   94
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   8760
      TabIndex        =   93
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   1800
      TabIndex        =   92
      Top             =   7440
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   91
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   7080
      TabIndex        =   90
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   8760
      TabIndex        =   89
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1800
      TabIndex        =   88
      Top             =   7200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   87
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblfacturaOriginal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblfacturaOriginal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   86
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblDetalleTubos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      TabIndex        =   85
      Top             =   10080
      Width           =   10695
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TUBOS VENDIDOS"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   84
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Line Line17 
      X1              =   120
      X2              =   11040
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Label lblCae 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCae"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   83
      Top             =   15120
      Width           =   10695
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9840
      TabIndex        =   82
      Top             =   13920
      Width           =   615
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "IVA 21%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   81
      Top             =   13920
      Width           =   855
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   80
      Top             =   13920
      Width           =   855
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Impuestos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   79
      Top             =   13920
      Width           =   855
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Neto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   78
      Top             =   13920
      Width           =   495
   End
   Begin VB.Line Line16 
      X1              =   8880
      X2              =   8880
      Y1              =   14280
      Y2              =   14760
   End
   Begin VB.Line Line15 
      X1              =   5880
      X2              =   5880
      Y1              =   14280
      Y2              =   14760
   End
   Begin VB.Line Line14 
      X1              =   4080
      X2              =   4080
      Y1              =   14280
      Y2              =   14760
   End
   Begin VB.Line Line13 
      X1              =   2160
      X2              =   2160
      Y1              =   14280
      Y2              =   14760
   End
   Begin VB.Line Line12 
      X1              =   120
      X2              =   11040
      Y1              =   14760
      Y2              =   14760
   End
   Begin VB.Line Line11 
      X1              =   8880
      X2              =   8880
      Y1              =   13800
      Y2              =   14280
   End
   Begin VB.Line Line10 
      X1              =   5880
      X2              =   5880
      Y1              =   13800
      Y2              =   14280
   End
   Begin VB.Line Line9 
      X1              =   4080
      X2              =   4080
      Y1              =   13800
      Y2              =   14280
   End
   Begin VB.Line Line8 
      X1              =   2160
      X2              =   2160
      Y1              =   13800
      Y2              =   14280
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   11040
      Y1              =   14280
      Y2              =   14280
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   11040
      Y1              =   13800
      Y2              =   13800
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   11040
      Y1              =   11400
      Y2              =   11400
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9480
      TabIndex        =   77
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8160
      TabIndex        =   76
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   75
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   74
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblTipoComprobante 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblTipoComprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6600
      TabIndex        =   59
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Distribuidor Autorizado: AIR LIQUIDE ARGENTINA S.A."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   68
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "de NUÑER ELIO RICARDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   67
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OXIJUNIN"
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
      Left            =   360
      TabIndex        =   66
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lblLetra 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblLetra"
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
      Height          =   615
      Left            =   5520
      TabIndex        =   60
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCategoria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCategoria"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      TabIndex        =   73
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblEmail"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   72
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label lblLocalidadEmpresa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblLocalidadEmpresa"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   71
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label lblTelefonos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblTelefonos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   70
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblDomicilioEmpresa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDomicilioEmpresa"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   69
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lblInicioActividades 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblInicioActividades"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   65
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label lblIngresosBrutos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblIngresosBrutos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   64
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label lblCuitEmpresa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCuitEmpresa"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   63
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "IVA RESPONSABLE INSCRIPTO"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4200
      TabIndex        =   62
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   11040
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   11040
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11040
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11040
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblNumero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   61
      Top             =   720
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   5280
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblFecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7560
      TabIndex        =   58
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblNombre"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   57
      Top             =   2400
      Width           =   4695
   End
   Begin VB.Label lblDomicilio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDomicilio"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   56
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Label lblLocalidad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblLocalidad"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   55
      Top             =   3120
      Width           =   4935
   End
   Begin VB.Label lblCuit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCuit"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      TabIndex        =   54
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Label lblCondicion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCondicion"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label lblNeto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Neto"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   52
      Top             =   14400
      Width           =   975
   End
   Begin VB.Label lblImpuesto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Impuesto"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   51
      Top             =   14400
      Width           =   855
   End
   Begin VB.Label lblSubtotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Subtotal"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   50
      Top             =   14400
      Width           =   855
   End
   Begin VB.Label lblIva 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Iva"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   49
      Top             =   14400
      Width           =   855
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9360
      TabIndex        =   48
      Top             =   14400
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   47
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   46
      Top             =   4560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   45
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   44
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   43
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   42
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   41
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   40
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   39
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   38
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   37
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   36
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   35
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   34
      Top             =   4800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   33
      Top             =   5040
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   32
      Top             =   5280
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   30
      Top             =   5760
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   29
      Top             =   6000
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   28
      Top             =   6240
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   27
      Top             =   6480
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   26
      Top             =   6720
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   1800
      TabIndex        =   25
      Top             =   6960
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8760
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   8760
      TabIndex        =   23
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   8760
      TabIndex        =   22
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   8760
      TabIndex        =   21
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8760
      TabIndex        =   20
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   8760
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   8760
      TabIndex        =   18
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   8760
      TabIndex        =   17
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   8760
      TabIndex        =   16
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   8760
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblVencimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblVencimiento"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   12960
      Width           =   10695
   End
   Begin VB.Label lblRemito 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblRemito"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   7080
      TabIndex        =   11
      Top             =   6960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   7080
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   7080
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   7080
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   7080
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7080
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   15600
      Left            =   120
      Top             =   -600
      Width           =   10935
   End
End
Attribute VB_Name = "frmImprimeFacturaElectronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cn.Open
    Set rs = cn.Execute("SELECT Ventas.Fecha, Ventas.Comprobante, Clientes.Nombre, Clientes.Domicilio, Clientes.Localidad, Categorias.Categoria, Clientes.NumeroDocumento, Ventas.Neto, Ventas.Impuestos, Ventas.IVA, Ventas.IVA2, Ventas.Total, Ventas.Tipo , Ventas.Condicion, Ventas.idVenta, Ventas.CAE, Ventas.VencimientoCae, Ventas.Puesto, Ventas.Numero, IsNull(Ventas.Comentario,'') as Comentario FROM Clientes INNER JOIN Ventas ON Clientes.idCliente = Ventas.idCliente INNER JOIN Categorias ON Clientes.idCategoria = Categorias.idCategoria WHERE (Ventas.idVenta = " & idComprobante & ")")
    lblTipoComprobante = rs!Comprobante
    lblLetra = rs!Tipo
    lblNumero = Format(rs!Puesto, "0000") & "-" & Format(rs!numero, "00000000")
    lblFecha = "Fecha: " & Format(rs!Fecha, "dd/mm/yyyy")
    lblNombre = "Razón Social: " & rs!Nombre
    lblDomicilio = "Domicilio:" & rs!Domicilio
    lblLocalidad = "Localidad: " & Format(rs!Localidad, "")
    lblCuit = "C.U.I.T.: " & rs!NumeroDocumento
    lblNeto = Format(rs!Neto, "0.00")
    lblImpuesto = Format(rs!Impuestos, "0.00")
    lblSubtotal = Format(rs!Neto + rs!Impuestos, "0.00")
    lblIva = Format(rs!Iva, "0.00")
    lblIva105 = Format(rs!Iva2, "0.00")
    lblTotal = Format(rs!Total, "0.00")
    lblComentario = rs!Comentario
    If lblLetra = "B" Then
        lblNeto = Format(rs!Neto + rs!Iva + rs!Iva2, "0.00")
        lblIva = Format(0, "0.00")
        lblIva105 = Format(0, "0.00")
        lblSubtotal = Format(rs!Neto + rs!Iva + rs!Iva2, "0.00")
    End If
    lblCategoria = "Categoría: " & rs!Categoria
    lblCondicion = "Condicion: " & rs!Condicion
    
    lblCae = "CAE: " & rs!CAE & "               Vencimiento: " & rs!VencimientoCAE
    Set rs = cn.Execute("SELECT Cuit, IngresosBrutos,InicioActividades,Domicilio,Localidad,CP,Telefonos,Email FROM Parametros")
    lblDomicilioEmpresa = rs!Domicilio
    lblLocalidadEmpresa = rs!Localidad & "  CP: " & rs!CP
    lblTelefonos = rs!Telefonos
    lblEmail = "e_mail: " & rs!Email
    lblCuitEmpresa = "C.U.I.T.: " & Format(rs!Cuit, "00-00000000-0")
    lblIngresosBrutos = "Ingresos Brutos: " & rs!IngresosBrutos
    lblInicioActividades = "Inicio Actividades: " & rs!InicioActividades
    
    

    Set rs = cn.Execute("SELECT DetalleVenta.Cantidad, Articulos.Descripcion, DetalleVenta.PrecioTotal FROM Articulos INNER JOIN DetalleVenta ON Articulos.idArticulo = DetalleVenta.idArticulo WHERE DetalleVenta.idVenta = " & idComprobante)
    x = 0
    While rs.EOF = False
        lblCantidad(x).Visible = True
        lblDescripcion(x).Visible = True
        lblUnitario(x).Visible = True
        lblImporte(x).Visible = True
        lblCantidad(x) = rs!Cantidad
        lblDescripcion(x) = rs!Descripcion
        lblUnitario(x) = Format(rs!PrecioTotal / rs!Cantidad, "0.00")
        lblImporte(x) = Format(rs!PrecioTotal, "0.00")
        rs.MoveNext
        x = x + 1
    Wend
    
    'armo el detalle de los tubos vendidos
    Set rs = cn.Execute("SELECT Articulos.Descripcion, Tubos.Numero, Tubos.Capacidad FROM Articulos INNER JOIN Tubos ON Articulos.idArticulo = Tubos.idArticulo INNER JOIN DetalleTubosVendidos ON Tubos.idTubo = DetalleTubosVendidos.idTubo WHERE DetalleTubosVendidos.idVenta = " & idComprobante & " ORDER BY Articulos.Descripcion")
    x = 0
    While rs.EOF = False
        lblDetalleTubos = lblDetalleTubos & rs!Descripcion & "-" & rs!numero & "(" & Format(rs!Capacidad, "0.00") & ")          "
        rs.MoveNext
    Wend

    'armo el detalle de los tubos devueltos
    Set rs = cn.Execute("SELECT Articulos.Descripcion, Tubos.Numero, Tubos.Capacidad FROM Articulos INNER JOIN Tubos ON Articulos.idArticulo = Tubos.idArticulo INNER JOIN DetalleTubosDevueltos ON Tubos.idTubo = DetalleTubosDevueltos.idTubo WHERE DetalleTubosDevueltos.idVenta = " & idComprobante & " ORDER BY Articulos.Descripcion")
    x = 0
    While rs.EOF = False
        lblDetalleDevueltos = lblDetalleDevueltos & rs!Descripcion & "-" & rs!numero & "(" & Format(rs!Capacidad, "0.00") & ")          "
        rs.MoveNext
    Wend


    cn.Close

End Sub

