VERSION 5.00
Begin VB.Form frmCajaCierre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Caja"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalidas 
      Caption         =   "Salidas"
      Height          =   855
      Left            =   3120
      Picture         =   "frmCajaCierre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCierreX 
      Caption         =   "Cierre X"
      Height          =   855
      Left            =   1800
      Picture         =   "frmCajaCierre.frx":0534
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCierreZ 
      Caption         =   "Cierre Z"
      Height          =   855
      Left            =   360
      Picture         =   "frmCajaCierre.frx":0649
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtFinalCredito 
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
      Left            =   6600
      TabIndex        =   24
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtFinalDebito 
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
      Left            =   4440
      TabIndex        =   23
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtFinalEfectivo 
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
      Left            =   2280
      TabIndex        =   22
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtSumaCredito 
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
      Left            =   6600
      TabIndex        =   21
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtSumaDebito 
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
      Left            =   4440
      TabIndex        =   20
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtSumaEfectivo 
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
      Left            =   2280
      TabIndex        =   19
      Top             =   2640
      Width           =   1695
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
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
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
      Left            =   4440
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
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
      Left            =   6600
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtSaldoCierre 
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
      Left            =   4800
      TabIndex        =   2
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   855
      Left            =   5880
      Picture         =   "frmCajaCierre.frx":0733
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   7200
      Picture         =   "frmCajaCierre.frx":0C95
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderStyle     =   6  'Inside Solid
      X1              =   2160
      X2              =   8400
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8520
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Final"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1320
      TabIndex        =   18
      Top             =   3720
      Width           =   750
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   17
      Top             =   2760
      Width           =   1650
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Inicial"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   16
      Top             =   2160
      Width           =   1050
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
      Left            =   1320
      TabIndex        =   15
      Top             =   960
      Width           =   930
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
      Left            =   0
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Efectivo"
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
      Left            =   2520
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Débito"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Crédito"
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
      Left            =   6840
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Saldo total de Cierre:"
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
      Left            =   1320
      TabIndex        =   10
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8520
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8640
      Y1              =   1440
      Y2              =   1440
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
      Left            =   0
      TabIndex        =   9
      Top             =   600
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
      Left            =   1320
      TabIndex        =   8
      Top             =   600
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "Cierre de Caja"
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
      Left            =   2760
      TabIndex        =   7
      Top             =   0
      Width           =   2895
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
      Left            =   7320
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmCajaCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Respuesta = MsgBox("¿Está seguro de cerrar la caja?", vbYesNo, "Atención!")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    cn.Execute ("CerrarCaja " & Replace(lblCajaNro, ",", ".") & "," & Replace(txtFinalEfectivo, ",", ".") & "," & Replace(txtFinalDebito, ",", ".") & "," & Replace(txtFinalCredito, ",", "."))
    
    'muestro reporte de caja
    Set rs = cn.Execute("VerCaja " & Val(lblCajaNro))
    With ReporteCaja.Sections("Sección4")
    .Controls("lblNroCaja").Caption = Val(lblCajaNro)
    .Controls("lblApertura").Caption = rs!Apertura
    .Controls("lblCierre").Caption = rs!Cierre
    .Controls("lblCajero").Caption = rs!Nombre
    .Controls("lblEfectivo").Caption = Format(rs!EfectivoInicial, "#.00")
    .Controls("lblDebito").Caption = Format(rs!DebitoInicial, "#.00")
    .Controls("lblCredito").Caption = Format(rs!CreditoInicial, "#.00")
    End With
    
    With ReporteCaja.Sections("Sección5")
    .Controls("lblEfectivoFinal").Caption = Format(rs!EfectivoFinal, "#.00")
    .Controls("lblDebitoFinal").Caption = Format(rs!DebitoFinal, "#.00")
    .Controls("lblCreditoFinal").Caption = Format(rs!CreditoFinal, "#.00")
    End With
        
        
    Set rs = cn.Execute("VerDetalleCaja " & Val(lblCajaNro))
    Set ReporteCaja.DataSource = rs
    ReporteCaja.WindowState = 2
    
    ReporteCaja.Show 1
    
    
    cn.Close
    CerroCaja = 1
    Unload Me
End Sub

Private Sub cmdCierreX_Click()
On Error GoTo impresora_apag
Procesar:
    frmFacturador.HASAR1.ReporteX
    Exit Sub
impresora_apag:
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If
End Sub

Private Sub cmdCierreZ_Click()
On Error GoTo impresora_apag
Procesar:
    frmFacturador.HASAR1.ReporteZ
    Exit Sub
impresora_apag:
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If
End Sub

Private Sub cmdSalidas_Click()
    frmMovimientosDeCaja.Show 1
    CargarCierre
End Sub

Private Sub Form_Load()
    lblFecha = Date
    lblUsuario = frmFacturador.lblCajero
    lblCajaNro = frmFacturador.lblCaja
    CargarCierre
End Sub

Sub CargarCierre()
    cn.Open
    Dim rs As New ADODB.Recordset
    Set rs = cn.Execute("TraerCajaAbierta " & Val(lblCajaNro))
    If rs.EOF = False Then
        txtEfectivo = Format(rs!EfectivoInicial, "0.00")
        txtDebito = Format(rs!DebitoInicial, "0.00")
        txtCredito = Format(rs!CreditoInicial, "0.00")
        txtSumaEfectivo = Format(rs!SumaEfectivo, "0.00")
        txtSumaDebito = Format(rs!SumaDebito, "0.00")
        txtSumaCredito = Format(rs!SumaCredito, "0.00")
        txtFinalEfectivo = Format(rs!EfectivoInicial + rs!SumaEfectivo, "0.00")
        txtFinalDebito = Format(rs!DebitoInicial + rs!SumaDebito, "0.00")
        txtFinalCredito = Format(rs!CreditoInicial + rs!SumaCredito, "0.00")
        txtSaldoCierre = Format(Val(txtFinalEfectivo) + Val(txtFinalDebito) + Val(txtFinalCredito), "0.00")
    End If
    cn.Close
End Sub
