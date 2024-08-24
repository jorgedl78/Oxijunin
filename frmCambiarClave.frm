VERSION 5.00
Begin VB.Form frmCambiarClave 
   Caption         =   "Cambiar Clave"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   1560
      Picture         =   "frmCambiarClave.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   3240
      Picture         =   "frmCambiarClave.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtConfirmarClave 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtNuevaClave 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   5520
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   5520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Confirmar Clave:"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nueva Clave:"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmCambiarClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If txtNuevaClave <> txtConfirmarClave Then MsgBox ("Las claves no coinciden"): Exit Sub
    cn.Open
    cn.Execute ("CambiarClaveUsuario '" & txtNuevaClave & "'," & idUsuario)
    cn.Close
    Unload Me
End Sub

Private Sub Form_Load()
    lblUsuario = Usuario
End Sub
