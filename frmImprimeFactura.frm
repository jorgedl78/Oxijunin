VERSION 5.00
Begin VB.Form frmImprimeFactura 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11250
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6960
      TabIndex        =   11
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   63
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   62
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7440
      TabIndex        =   61
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   7440
      TabIndex        =   60
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   7440
      TabIndex        =   59
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   58
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   7440
      TabIndex        =   57
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   7440
      TabIndex        =   56
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   7440
      TabIndex        =   55
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   7440
      TabIndex        =   54
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblUnitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   7440
      TabIndex        =   53
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblRemito 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblRemito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8880
      TabIndex        =   52
      Top             =   3675
      Width           =   1095
   End
   Begin VB.Label lblComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblComentario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2040
      TabIndex        =   51
      Top             =   7680
      Width           =   5415
   End
   Begin VB.Label lblVencimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblVencimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9600
      TabIndex        =   50
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   9000
      TabIndex        =   49
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   9000
      TabIndex        =   48
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   9000
      TabIndex        =   47
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   9000
      TabIndex        =   46
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   9000
      TabIndex        =   45
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   9000
      TabIndex        =   44
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   9000
      TabIndex        =   43
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   9000
      TabIndex        =   42
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   9000
      TabIndex        =   41
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   9000
      TabIndex        =   40
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   1920
      TabIndex        =   39
      Top             =   7080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   1920
      TabIndex        =   38
      Top             =   6840
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   1920
      TabIndex        =   37
      Top             =   6600
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   1920
      TabIndex        =   36
      Top             =   6360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   35
      Top             =   6120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   34
      Top             =   5880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   33
      Top             =   5640
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   32
      Top             =   5400
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   31
      Top             =   5160
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   30
      Top             =   4920
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   29
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   28
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   26
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   23
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblImporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   9000
      TabIndex        =   19
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDescripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   18
      Top             =   4680
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   16
      Top             =   10800
      Width           =   975
   End
   Begin VB.Label lblIva 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Iva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9360
      TabIndex        =   15
      Top             =   10440
      Width           =   855
   End
   Begin VB.Label lblSubtotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9360
      TabIndex        =   14
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label lblImpuesto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Impuesto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9360
      TabIndex        =   13
      Top             =   9480
      Width           =   855
   End
   Begin VB.Label lblNeto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Neto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   12
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label lblCuentaCorriente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblContado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblResponsableInscripto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblCuit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblCuit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   3645
      Width           =   1575
   End
   Begin VB.Label lblLocalidad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblLocalidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label lblDomicilio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDomicilio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label lblNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblNombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2880
      Width           =   8895
   End
   Begin VB.Label lblAnio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblAnio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10440
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblMes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblMes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblDia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblDia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblTipoComprobante 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lblTipoComprobante"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmImprimeFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    cn.Open
    Set rs = cn.Execute("SELECT Ventas.Fecha, Ventas.Comprobante, Clientes.Nombre, Clientes.Domicilio, Clientes.Localidad, Categorias.Categoria, Clientes.NumeroDocumento, Ventas.Neto, Ventas.Impuestos, Ventas.IVA, Ventas.Total, Ventas.Tipo , Ventas.Condicion, Ventas.idVenta FROM Clientes INNER JOIN Ventas ON Clientes.idCliente = Ventas.idCliente INNER JOIN Categorias ON Clientes.idCategoria = Categorias.idCategoria WHERE (Ventas.idVenta = " & idComprobante & ")")
    lblTipoComprobante = Comprobante
    lblDia = Format(rs!Fecha, "dd")
    lblMes = Format(rs!Fecha, "mm")
    lblAnio = Format(rs!Fecha, "yy")
    lblNombre = rs!Nombre
    lblDomicilio = rs!Domicilio
    lblLocalidad = Format(rs!Localidad, "")
    lblCuit = rs!NumeroDocumento
    lblNeto = Format(rs!Neto, "0.00")
    lblImpuesto = Format(rs!Impuestos, "0.00")
    lblSubtotal = Format(rs!Neto + rs!Impuestos, "0.00")
    lblIva = Format(rs!Iva, "0.00")
    lblTotal = Format(rs!Total, "0.00")
    If rs!Categoria = "Responsable Inscripto" Then lblResponsableInscripto.Visible = True
    If rs!Condicion = "CONTADO" Then
        lblContado.Visible = True
    Else
        lblCuentaCorriente.Visible = True
    End If
    

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
    cn.Close
End Sub

