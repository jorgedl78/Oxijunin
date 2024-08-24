VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReciboProveedores 
   Caption         =   "Pago a Proveedores"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumero 
      Height          =   375
      Left            =   8400
      TabIndex        =   19
      Top             =   240
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   91750401
      CurrentDate     =   42390
   End
   Begin VB.Frame Frame4 
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   10215
      Begin VB.TextBox txtDetalle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   1200
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   8295
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "(Hasta 300 caracteres y/o 18 lineas)"
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   1920
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Detalle:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   810
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   10215
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   855
         Left            =   5760
         Picture         =   "frmReciboProveedores.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Confirmar"
         Height          =   855
         Left            =   3360
         Picture         =   "frmReciboProveedores.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   5655
      Begin VB.TextBox txtImporte 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe a pagarr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10215
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   5175
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00FFFFFF&
         Height          =   855
         Left            =   6840
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo:"
         Height          =   255
         Left            =   6240
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblSaldo 
         Alignment       =   2  'Center
         Caption         =   "lblSaldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Número:"
      Height          =   255
      Left            =   7560
      TabIndex        =   18
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   240
      Width           =   450
   End
   Begin VB.Label lblRecibo 
      Alignment       =   1  'Right Justify
      Caption         =   "lblRecibo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmReciboProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NuevoNumero As Integer

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If txtNumero = "" Then MsgBox ("Debe especificar en número de comprobante"): Exit Sub
    If txtImporte = "" Then MsgBox ("Debe especificar importe"): Exit Sub
    If (IsNumeric(txtImporte) = False) Then MsgBox ("El importe no es válido"): Exit Sub
    If txtImporte <= 0 Then MsgBox ("Debe especificar un importe a pagar"): Exit Sub
    Respuesta = MsgBox("¿Esta seguro de ingresar el pago?", vbYesNo, "Guardar")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    cn.Execute ("INSERT INTO RecibosProveedor(Fecha,Numero,Importe,idProveedor, Detalle) VALUES ('" & dtFecha.Value & "'," & txtNumero & "," & Replace(txtImporte, ",", ".") & "," & idProveedor & ",'" & txtDetalle & "')")
    cn.Execute ("AgregarCuentaCorrienteProveedor '" & Format(dtFecha, "yyyy/mm/dd") & "','Recibo " & Format(txtNumero, "00000000") & "',0," & Replace(txtImporte, ",", ".") & "," & idProveedor & ",'Rec'," & txtNumero)
    
    'Set rs = cn.Execute("SELECT max(idRecibo) as UltimoRecibo FROM Recibos")
    'idRecibo = rs!UltimoRecibo
    'Set ImprimeRecibo.DataSource = rs

    'ImprimeRecibo.WindowState = 2
    
    'ImprimeRecibo.Show 1
    
    cn.Close
    Unload Me
End Sub

Private Sub Form_Activate()
    txtNumero.SetFocus
End Sub

Private Sub Form_Load()
    Dim rs As Recordset
    cn.Open
    'Set rs = cn.Execute("SELECT Recibo + 1 AS NuevoR FROM Parametros")
    'NuevoNumero = rs!NuevoR
    lblRecibo = "Recibo Nº: " & Format(NuevoNumero, "00000000")
    Set rs = cn.Execute("SELECT IsNull(sum(Debe) - sum(Haber),0) as saldo FROM CuentaCorrienteProveedor  where idProveedor=" & idProveedor)
    txtImporte.Text = Format(rs!saldo, "0.0000")
    lblSaldo = Format(rs!saldo, "0.0000")
    cn.Close
    dtFecha.Value = Date
End Sub


Private Sub txtImporte_Change()
    'Text1.Text = Format(txtImporte.Text, "#,##0.00")
End Sub

Private Sub txtImporte_GotFocus()
    txtImporte.SelStart = 0
    txtImporte.SelLength = Len(txtImporte)
    
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789,." & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    separador = 0
    For i = 1 To Len(txtImporte) + 1
        If (Mid(txtImporte, i, 1) = "," Or Mid(txtImporte, i, 1) = ".") Then
            separador = separador + 1
        End If
    Next i
    If separador > 0 And KeyAscii <> 8 And (KeyAscii = 46 Or KeyAscii = 44) Then KeyAscii = 0
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(13) & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If

End Sub
