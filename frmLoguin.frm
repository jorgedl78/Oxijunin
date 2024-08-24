VERSION 5.00
Begin VB.Form frmLoguin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loguin"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Height          =   735
      Left            =   1800
      Picture         =   "frmLoguin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtClave 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.ComboBox cbUsuario 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Clave:"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "frmLoguin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const GWL_EXSTYLE As Long = (-20)
Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private m_bActivated As Boolean

Private Sub cbUsuario_Click()
    txtClave.SetFocus
End Sub

Private Sub cmdAceptar_Click()
    If cbUsuario.ListIndex < 0 Then
       MsgBox ("Debe elejir un usuario")
       Exit Sub
    End If
    idUsuario = cbUsuario.ItemData(cbUsuario.ListIndex)
    Usuario = cbUsuario.Text
    Dim rs As New ADODB.Recordset
    cn.Open
    Set rs = cn.Execute("VerClaveUsuario " & idUsuario)
    Clave = rs!Clave
    rs.Close
    Set rs = Nothing
    cn.Close
    If txtClave <> Clave Then MsgBox ("Clave incorrecta"): txtClave.SelStart = 0: txtClave.SelLength = Len(txtClave): txtClave.SetFocus: Exit Sub
    frmMenuPrincipal.Show 1
    End
End Sub


Private Sub Form_Activate()
If Not m_bActivated Then
    m_bActivated = True
    Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
    Call ShowWindow(hwnd, SW_HIDE)
    Call ShowWindow(hwnd, SW_SHOW)
End If
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cn.Open
    Set rs = cn.Execute("PedirUsuarios")
    Do While rs.EOF = False
        cbUsuario.AddItem rs!Nombre
        cbUsuario.ItemData(cbUsuario.NewIndex) = rs!idUsuario
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub

Private Sub Text2_Change()

End Sub


Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAceptar_Click
End Sub
