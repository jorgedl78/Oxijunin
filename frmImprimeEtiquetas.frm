VERSION 5.00
Begin VB.Form frmImprimeEtiquetas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   16005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16005
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3600
      TabIndex        =   93
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   3960
      TabIndex        =   92
      Top             =   16200
      Width           =   2415
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   7560
      TabIndex        =   91
      Top             =   16200
      Width           =   2415
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   360
      TabIndex        =   90
      Top             =   16200
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   7680
      TabIndex        =   89
      Top             =   12240
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
      Index           =   17
      Left            =   7560
      TabIndex        =   88
      Top             =   13440
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
      Index           =   17
      Left            =   8520
      TabIndex        =   87
      Top             =   12840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   17
      Left            =   7440
      Top             =   12120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   7680
      TabIndex        =   86
      Top             =   12600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   7560
      TabIndex        =   85
      Top             =   13800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   4080
      TabIndex        =   84
      Top             =   12600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   16
      Left            =   3840
      Top             =   12120
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
      Index           =   16
      Left            =   4920
      TabIndex        =   83
      Top             =   12840
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
      Index           =   16
      Left            =   3960
      TabIndex        =   82
      Top             =   13440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   4080
      TabIndex        =   81
      Top             =   12240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   3960
      TabIndex        =   80
      Top             =   13800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   480
      TabIndex        =   79
      Top             =   12240
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
      Index           =   15
      Left            =   360
      TabIndex        =   78
      Top             =   13440
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
      Index           =   15
      Left            =   1320
      TabIndex        =   77
      Top             =   12840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   15
      Left            =   240
      Top             =   12120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   480
      TabIndex        =   76
      Top             =   12600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   75
      Top             =   13800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   7680
      TabIndex        =   74
      Top             =   9840
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
      Index           =   14
      Left            =   7560
      TabIndex        =   73
      Top             =   11040
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
      Index           =   14
      Left            =   8520
      TabIndex        =   72
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   14
      Left            =   7440
      Top             =   9720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   7680
      TabIndex        =   71
      Top             =   10200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   7560
      TabIndex        =   70
      Top             =   11400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   69
      Top             =   10200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   13
      Left            =   3840
      Top             =   9720
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
      Index           =   13
      Left            =   4920
      TabIndex        =   68
      Top             =   10440
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
      Index           =   13
      Left            =   3960
      TabIndex        =   67
      Top             =   11040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   66
      Top             =   9840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   65
      Top             =   11400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   480
      TabIndex        =   64
      Top             =   9840
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
      Index           =   12
      Left            =   360
      TabIndex        =   63
      Top             =   11040
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
      Index           =   12
      Left            =   1320
      TabIndex        =   62
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   12
      Left            =   240
      Top             =   9720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   480
      TabIndex        =   61
      Top             =   10200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   60
      Top             =   11400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   7680
      TabIndex        =   59
      Top             =   7440
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
      Index           =   11
      Left            =   7560
      TabIndex        =   58
      Top             =   8640
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
      Index           =   11
      Left            =   8520
      TabIndex        =   57
      Top             =   8040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   11
      Left            =   7440
      Top             =   7320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   7680
      TabIndex        =   56
      Top             =   7800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   7560
      TabIndex        =   55
      Top             =   9000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   54
      Top             =   7800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   10
      Left            =   3840
      Top             =   7320
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
      Index           =   10
      Left            =   4920
      TabIndex        =   53
      Top             =   8040
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
      Index           =   10
      Left            =   3960
      TabIndex        =   52
      Top             =   8640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   51
      Top             =   7440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3960
      TabIndex        =   50
      Top             =   9000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   49
      Top             =   7440
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
      Index           =   9
      Left            =   360
      TabIndex        =   48
      Top             =   8640
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
      Index           =   9
      Left            =   1320
      TabIndex        =   47
      Top             =   8040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   9
      Left            =   240
      Top             =   7320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   46
      Top             =   7800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   45
      Top             =   9000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   7680
      TabIndex        =   44
      Top             =   5040
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
      Index           =   8
      Left            =   7560
      TabIndex        =   43
      Top             =   6240
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
      Index           =   8
      Left            =   8520
      TabIndex        =   42
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   8
      Left            =   7440
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   7680
      TabIndex        =   41
      Top             =   5400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   7560
      TabIndex        =   40
      Top             =   6600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   4080
      TabIndex        =   39
      Top             =   5400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   7
      Left            =   3840
      Top             =   4920
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
      Index           =   7
      Left            =   4920
      TabIndex        =   38
      Top             =   5640
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
      Index           =   7
      Left            =   3960
      TabIndex        =   37
      Top             =   6240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   4080
      TabIndex        =   36
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3960
      TabIndex        =   35
      Top             =   6600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   34
      Top             =   5040
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
      Index           =   6
      Left            =   360
      TabIndex        =   33
      Top             =   6240
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
      Index           =   6
      Left            =   1320
      TabIndex        =   32
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   6
      Left            =   240
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   31
      Top             =   5400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   7680
      TabIndex        =   29
      Top             =   2640
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
      Index           =   5
      Left            =   7560
      TabIndex        =   28
      Top             =   3840
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
      Index           =   5
      Left            =   8520
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   5
      Left            =   7440
      Top             =   2520
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   7680
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   24
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   4
      Left            =   3840
      Top             =   2520
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
      Index           =   4
      Left            =   4920
      TabIndex        =   23
      Top             =   3240
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
      Index           =   4
      Left            =   3960
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   19
      Top             =   2640
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
      Index           =   3
      Left            =   360
      TabIndex        =   18
      Top             =   3840
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
      Index           =   3
      Left            =   1320
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   3
      Left            =   240
      Top             =   2520
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblNumeroBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   2
      Left            =   7440
      Top             =   120
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
      Index           =   2
      Left            =   8520
      TabIndex        =   12
      Top             =   840
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
      Index           =   2
      Left            =   7560
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   10
      Top             =   240
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
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   240
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
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   1440
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
      Index           =   1
      Left            =   4920
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   1
      Left            =   3840
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   5
      Top             =   600
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
      Left            =   3960
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descripcion"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Index           =   0
      Left            =   240
      Top             =   120
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
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Codigo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "frmImprimeEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
