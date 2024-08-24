VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCuentaCorrienteProveedor 
   Caption         =   "Cuenta Corriente de Proveedor"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin MSComCtl2.DTPicker CalendarHasta 
         Height          =   375
         Left            =   5160
         TabIndex        =   5
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   91029505
         CurrentDate     =   42391
      End
      Begin MSComCtl2.DTPicker CalendarDesde 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   91029505
         CurrentDate     =   42391
      End
      Begin VB.CommandButton cmdImprimirDetalle 
         Caption         =   "Imprimir"
         Height          =   855
         Left            =   3240
         Picture         =   "frmCuentaCorrienteProveedor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
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
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
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
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCuentaCorrienteProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimirDetalle_Click()
    Dim rs As New Recordset
    cn.Open
    Set rs = cn.Execute("SELECT IsNull(sum(Debe) - sum(Haber),0) as inicial FROM CuentaCorrienteProveedor  where idProveedor=" & idProveedor & " and Fecha < '" & CalendarDesde.Value & "'")
    With CuentaCorrienteProveedor.Sections("Sección4")
        .Controls("lblCliente").Caption = frmProveedores.grclientes.TextMatrix(frmProveedores.grclientes.Row, 1)
        .Controls("lblDesde").Caption = CalendarDesde.Value
        .Controls("lblHasta").Caption = CalendarHasta.Value
        .Controls("lblSaldoInicial").Caption = Format(rs!inicial, "0.00")
    End With
    
    With CuentaCorrienteProveedor.Sections("Sección5")
        Set rs = cn.Execute("SELECT IsNull(sum(Debe) - sum(Haber),0) as final FROM CuentaCorrienteProveedor  where idProveedor=" & idProveedor & " and Fecha <= '" & CalendarHasta.Value & "'")
        .Controls("lblSaldoFinal").Caption = Format(rs!final, "0.00")
    End With
        
    Set rs = cn.Execute("ConsultaCtaCteProveedor '" & CalendarDesde.Value & "','" & CalendarHasta.Value & "'," & idProveedor)
    Set CuentaCorrienteProveedor.DataSource = rs
    CuentaCorrienteProveedor.WindowState = 2
    
    CuentaCorrienteProveedor.Show 1
    cn.Close

End Sub

Private Sub Form_Load()
    CalendarDesde.Value = Date
    CalendarHasta.Value = Date
End Sub
