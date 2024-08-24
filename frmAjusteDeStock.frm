VERSION 5.00
Begin VB.Form frmAjusteDeStock 
   Caption         =   "Ajuste de Stock"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8745
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Ajuste"
      Height          =   2535
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   7935
      Begin VB.Frame frMovimiento 
         Caption         =   "Movimiento"
         Height          =   1575
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   4575
         Begin VB.TextBox txtMotivo 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   19
            Top             =   960
            Width           =   3255
         End
         Begin VB.OptionButton optSuma 
            Caption         =   "Suma"
            Height          =   255
            Left            =   1440
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optResta 
            Caption         =   "Resta"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Motivo:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.TextBox txtNuevo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtAjuste 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtStock 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   2640
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Stock:"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ajuste:"
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Stock Actual:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   960
      End
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtCodBarras 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   960
      Width           =   7215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   3240
      Picture         =   "frmAjusteDeStock.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   4920
      Picture         =   "frmAjusteDeStock.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   6240
      TabIndex        =   8
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Barras:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmAjusteDeStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If Val(txtAjuste) = 0 Then MsgBox ("Debe especificar la cantidad a ajustar"): txtAjuste.SetFocus: Exit Sub
    If IsNumeric(txtAjuste) = False Then MsgBox ("La cantidad no es correcta"): txtAjuste.SetFocus: Exit Sub
    If txtMotivo = "" Then MsgBox ("Debe especificar un motivo"): Exit Sub
    Respuesta = MsgBox("¿Confirma el ajuste?", vbYesNo, "Confirmar")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    If optResta.Value = True Then
        MOvimiento = "Resta"
        cn.Execute ("UPDATE Articulos set Stock=Stock - " & txtAjuste & " WHERE idArticulo=" & txtCodigo)
    End If
    If optSuma.Value = True Then
        MOvimiento = "Suma"
        cn.Execute ("UPDATE Articulos set Stock=Stock + " & txtAjuste & " WHERE idArticulo=" & txtCodigo)
    End If
    cn.Execute ("INSERT INTO Ajustes_Stock(Cantidad,Movimiento,Fecha,Motivo,idArticulo, StockActual, StockAjustado) VALUES(" & txtAjuste & ",'" & MOvimiento & "','" & Format(Date, "yyyy/mm/dd") & "','" & txtMotivo & "'," & txtCodigo & "," & txtStock & "," & txtNuevo & ")")

    cn.Close
    Saltar = 0
    Unload Me

End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cn.Open
    Set rs = cn.Execute("VerArticulo " & idArticulo)
    txtCodigo = rs!idArticulo
    txtCodBarras = rs!CodBar
    txtDescripcion = rs!Descripcion
    txtStock = rs!Stock
    cn.Close
    'optResta.Value = True
End Sub

Private Sub optResta_Click()
    txtAjuste.BackColor = &HC0C0FF
    txtAjuste.SetFocus
End Sub

Private Sub optSuma_Click()
    txtAjuste.BackColor = &HC0FFC0
    txtAjuste.SetFocus
End Sub

Private Sub txtAjuste_Change()
    If optResta.Value = True Then
        txtNuevo = Val(txtStock) - Val(txtAjuste)
    Else
        txtNuevo = Val(txtStock) + Val(txtAjuste)
    End If
End Sub

Private Sub txtAjuste_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
