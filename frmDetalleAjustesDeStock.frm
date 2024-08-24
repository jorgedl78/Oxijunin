VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDetalleAjustesDeStock 
   Caption         =   "Detalle de justes de Stock"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      Height          =   855
      Left            =   4320
      Picture         =   "frmDetalleAjustesDeStock.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grDetalleAjuste 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10398
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
   Begin VB.Label lblDescripcion 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmDetalleAjustesDeStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    grDetalleAjuste.Cols = 6
    grDetalleAjuste.ColWidth(0) = 1000
    grDetalleAjuste.ColWidth(1) = 800
    grDetalleAjuste.ColWidth(2) = 1000
    grDetalleAjuste.ColWidth(3) = 800
    grDetalleAjuste.ColWidth(4) = 800
    grDetalleAjuste.ColWidth(5) = 4000
    
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT Fecha, StockActual, Cantidad, StockAjustado, Movimiento, Motivo FROM Ajustes_Stock WHERE (idArticulo = " & idArticulo & ") ORDER BY idAjusteStock DESC")
    With grDetalleAjuste
    .Rows = 1
    .TextArray(0) = "Fecha"
    .TextArray(1) = "Stock"
    .TextArray(2) = "Movimiento"
    .TextArray(3) = "Ajuste"
    .TextArray(4) = "Nuevo"
    .TextArray(5) = "Motivo"
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!Fecha
        .TextMatrix(.Rows - 1, 1) = rs!StockActual
        .TextMatrix(.Rows - 1, 2) = rs!MOvimiento
        .TextMatrix(.Rows - 1, 3) = rs!Cantidad
        .TextMatrix(.Rows - 1, 4) = rs!StockAjustado
        .TextMatrix(.Rows - 1, 5) = rs!Motivo
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With
    rs.Close
    Set rs = Nothing
    cn.Close

End Sub
