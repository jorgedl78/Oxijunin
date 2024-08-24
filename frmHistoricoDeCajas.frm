VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmHistoricoDeCajas 
   Caption         =   "Histórico de Cajas"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimirCaja 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   4920
      Picture         =   "frmHistoricoDeCajas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grCaja 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      GridColorFixed  =   255
      TextStyleFixed  =   3
      HighLight       =   2
      SelectionMode   =   1
      GridLineWidthFixed=   1
      FontWidthFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmHistoricoDeCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdImprimirCaja_Click()
    'muestro reporte de caja
    cajaSeleccionada = grCaja.TextMatrix(grCaja.Row, 0)
    cn.Open
    Set rs = cn.Execute("VerCaja " & cajaSeleccionada)
    With ReporteCaja.Sections("Sección4")
    .Controls("lblNroCaja").Caption = cajaSeleccionada
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
        
        
    Set rs = cn.Execute("VerDetalleCaja " & cajaSeleccionada)
    Set ReporteCaja.DataSource = rs
    ReporteCaja.WindowState = 2
    
    ReporteCaja.Show 1
       
    cn.Close

End Sub

Private Sub Form_Load()
    grCaja.Cols = 10
    grCaja.ColWidth(0) = 500
    grCaja.ColWidth(1) = 1000
    grCaja.ColWidth(2) = 2000
    grCaja.ColWidth(3) = 800
    grCaja.ColWidth(4) = 800
    grCaja.ColWidth(5) = 800
    grCaja.ColWidth(6) = 800
    grCaja.ColWidth(7) = 800
    grCaja.ColWidth(8) = 800
    grCaja.ColWidth(9) = 1700
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("VerListadoCaja")
    With grCaja
    .Rows = 1
    .TextArray(0) = "Nº"
    .TextArray(1) = "Fecha"
    .TextArray(2) = "Cajero"
    .TextArray(3) = "Efectivo"
    .TextArray(4) = "Débito"
    .TextArray(5) = "Crédito"
    .TextArray(6) = "Efectivo"
    .TextArray(7) = "Débito"
    .TextArray(8) = "Crédito"
    .TextArray(9) = "Cierre"
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idCaja
        .TextMatrix(.Rows - 1, 1) = rs!Apertura
        .TextMatrix(.Rows - 1, 2) = rs!Nombre
        .TextMatrix(.Rows - 1, 3) = Format(rs!EfectivoInicial, "0.00")
        .TextMatrix(.Rows - 1, 4) = Format(rs!DebitoInicial, "0.00")
        .TextMatrix(.Rows - 1, 5) = Format(rs!CreditoInicial, "0.00")
        .TextMatrix(.Rows - 1, 6) = Format(rs!EfectivoFinal, "0.00")
        .TextMatrix(.Rows - 1, 7) = Format(rs!DebitoFinal, "0.00")
        .TextMatrix(.Rows - 1, 8) = Format(rs!CreditoFinal, "0.00")
        .TextMatrix(.Rows - 1, 9) = rs!Cierre
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub
