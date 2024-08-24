VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecibo 
   Caption         =   "Recibo"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46268417
      CurrentDate     =   42390
   End
   Begin VB.Frame Frame4 
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   8655
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
         Width           =   7215
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
      Left            =   0
      TabIndex        =   8
      Top             =   6000
      Width           =   8895
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   855
         Left            =   4920
         Picture         =   "frmRecibo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Confirmar"
         Height          =   855
         Left            =   2520
         Picture         =   "frmRecibo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   1440
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
         Caption         =   "Importe a cobrar:"
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
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8895
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
         Width           =   1695
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
         Width           =   1455
      End
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
      Left            =   5280
      TabIndex        =   11
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NuevoNumero As Integer

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If txtImporte = "" Then MsgBox ("Debe especificar importe"): Exit Sub
    If (IsNumeric(txtImporte) = False) Then MsgBox ("El importe no es válido"): Exit Sub
    If txtImporte <= 0 Then MsgBox ("Debe especificar un importe a cobrar"): Exit Sub
    Respuesta = MsgBox("¿Esta seguro de ingresar el pago?", vbYesNo, "Guardar")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    cn.Execute ("INSERT INTO Recibos(Fecha,Numero,Importe,idCliente, Detalle) VALUES ('" & dtFecha.Value & "'," & NuevoNumero & "," & Replace(txtImporte, ",", ".") & "," & idCliente & ",'" & txtDetalle & "')")
    cn.Execute ("AgregarCuentaCorriente '" & Format(dtFecha, "yyyy/mm/dd") & "','Recibo " & Format(NuevoNumero, "00000000") & "',0," & Replace(txtImporte, ",", ".") & "," & idCliente & ",'Rec'," & NuevoNumero)
    cn.Execute ("UPDATE Parametros SET Recibo=Recibo + 1")
    Set rs = cn.Execute("SELECT max(idRecibo) as UltimoRecibo FROM Recibos")
    idRecibo = rs!UltimoRecibo
    Set ImprimeRecibo.DataSource = rs

    ImprimeRecibo.WindowState = 2
    
    ImprimeRecibo.Show 1
    
    cn.Close
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As Recordset
    cn.Open
    Set rs = cn.Execute("SELECT Recibo + 1 AS NuevoR FROM Parametros")
    NuevoNumero = rs!NuevoR
    lblRecibo = "Recibo Nº: " & Format(NuevoNumero, "00000000")
    Set rs = cn.Execute("SELECT IsNull(sum(Debe) - sum(Haber),0) as saldo FROM CuentaCorriente  where idCliente=" & idCliente)
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
    For I = 1 To Len(txtImporte) + 1
        If (Mid(txtImporte, I, 1) = "," Or Mid(txtImporte, I, 1) = ".") Then
            separador = separador + 1
        End If
    Next I
    If separador > 0 And KeyAscii <> 8 And (KeyAscii = 46 Or KeyAscii = 44) Then KeyAscii = 0
End Sub
