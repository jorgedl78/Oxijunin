VERSION 5.00
Begin VB.Form frmPedirPrecio 
   BackColor       =   &H00404040&
   Caption         =   "Precio de Venta"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescripcion 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   975
      Left            =   4440
      Picture         =   "frmPedirPrecio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtPrecio 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio: $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "frmPedirPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAceptar_Click()
   If Val(frmPedirPrecio.txtPrecio) > 9999999.99 Then
    MsgBox ("El precio es incorrecto"): Exit Sub
   End If
   txtPrecio = Replace(txtPrecio, ".", ",")
   With frmFacturador
   If txtDescripcion <> "" Then
        .grDetalle.TextMatrix(.grDetalle.Rows - 1, 1) = txtDescripcion
   End If
   End With
   Me.Hide
End Sub

Private Sub Form_Activate()
   'txtPrecio = ""
   'txtDescripcion = ""
   'With frmFacturador
   'If .grDetalle.TextMatrix(.grDetalle.Rows - 1, 1) = "VARIOS" Then
   '     Label2.Visible = True
   '     txtDescripcion.Visible = True
   '     txtDescripcion.SetFocus
   'ElseIf .grDetalle.TextMatrix(.grDetalle.Rows - 1, 1) = "FIAMBRERIA" Or .grDetalle.TextMatrix(.grDetalle.Rows - 1, 1) = "PANADERIA" Or .grDetalle.TextMatrix(.grDetalle.Rows - 1, 1) = "KIOSCO" Or .grDetalle.TextMatrix(.grDetalle.Rows - 1, 1) = "ESPECIAS" Or .grDetalle.TextMatrix(.grDetalle.Rows - 1, 1) = "PASTAS CASERAS" Or .grDetalle.TextMatrix(.grDetalle.Rows - 1, 1) = "PESCADERIA" Then
   '     Label2.Visible = False
   '     txtDescripcion.Visible = False
   'Else
   '     Label2.Visible = False
   '     txtDescripcion.Visible = False
   'End If
   'End With
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPrecio.SetFocus
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then cmdAceptar_Click
End Sub
