VERSION 5.00
Begin VB.Form frmMovimientosDeCaja 
   Caption         =   "Movimientos de Caja"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComentario 
      Height          =   375
      Left            =   360
      MaxLength       =   30
      TabIndex        =   15
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox txtCredito 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   14
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtDebito 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   1920
      Picture         =   "frmMovimientosDeCaja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   4440
      Picture         =   "frmMovimientosDeCaja.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtEfectivo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox cmbTipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   6
      Text            =   "Tipo de Movimiento"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Comentario"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Crédito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Débito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   11
      Top             =   2760
      Width           =   990
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7200
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Efectivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
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
      Left            =   6000
      TabIndex        =   5
      Top             =   960
      Width           =   1215
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
      TabIndex        =   4
      Top             =   720
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
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7200
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label7 
      Caption         =   "Caja Nº:"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblCajaNro 
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
      TabIndex        =   1
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "Movimientos de Caja"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmMovimientosDeCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If cmbTipo.ListIndex < 0 Then
        MsgBox ("Debe elejir un tipo de movimiento")
        Exit Sub
    End If
    cn.Open
    cn.Execute ("AgregarDetalleCaja " & lblCajaNro & ",'" & Format(Date, "yyyy/mm/dd") & "'," & Replace((Val(txtEfectivo) * (-1)), ",", ".") & "," & Replace((Val(txtDebito) * (-1)), ",", ".") & "," & Replace((Val(txtCredito) * (-1)), ",", ".") & "," & cmbTipo.ItemData(cmbTipo.ListIndex) & ",'" & txtComentario & "'")
    cn.Close
    Unload Me
End Sub

Private Sub Form_Load()
    lblFecha = Date
    lblUsuario = frmFacturador.lblCajero
    lblCajaNro = frmFacturador.lblCaja
    cn.Open
    Dim rs As New ADODB.Recordset
    Set rs = cn.Execute("MostrarTiposDetalle")
    Do While rs.EOF = False
        cmbTipo.AddItem rs!Tipo
        cmbTipo.ItemData(cmbTipo.NewIndex) = rs!idTipoDetalle
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCredito_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDebito_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
