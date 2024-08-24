VERSION 5.00
Begin VB.Form frmCajaApertura 
   Caption         =   "Apertura de Caja"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   3240
      Picture         =   "frmCajaApertura.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   855
      Left            =   720
      Picture         =   "frmCajaApertura.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtSaldoApertura 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtCredito 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtDebito 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtEfectivo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblFecha 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Apertura de Caja"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblUsuario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1440
      TabIndex        =   11
      Top             =   840
      Width           =   930
   End
   Begin VB.Label Label6 
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4920
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4800
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label5 
      Caption         =   "Saldo de apertura:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Crédito:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Débito:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Efectivo:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "frmCajaApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbrir_Click()
    Dim rs As New ADODB.Recordset
    Set rs = cn.Execute("AbrirCaja " & idUsuario & ",'" & Format(Date, "yyyy/mm/dd") & "'," & Replace(txtEfectivo, ",", ".") & "," & Replace(txtDebito, ",", ".") & "," & Replace(txtCredito, ",", "."))
    frmFacturador.lblCaja = rs!NuevoID
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    frmFacturador.lblCaja = 0
    Unload Me
End Sub

Private Sub Form_Load()
    lblFecha = Date
End Sub

