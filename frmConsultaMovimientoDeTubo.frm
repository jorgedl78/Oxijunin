VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConsultaMovimientoDeTubo 
   Caption         =   "Movimientos de Tubo"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   3120
      Picture         =   "frmConsultaMovimientoDeTubo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Movimientos"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grMovimientos 
         Height          =   6375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   11245
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
End
Attribute VB_Name = "frmConsultaMovimientoDeTubo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With grMovimientos
        .Cols = 2
        .ColWidth(0) = 1200 'fecha
        .ColWidth(1) = 5000 'detalle

    
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("Select Fecha, Detalle FROM MovimientosTubos WHERE idTubo = " & idTubo)
    'lblEncontrados = rs.RecordCount
   ' Set grClientes.DataSource = rs

    .Rows = 1
    .TextArray(0) = "Fecha"
    .TextArray(1) = "Detalle"
   
    
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!Fecha
        .TextMatrix(.Rows - 1, 1) = rs!Detalle
        rs.MoveNext
        
        .FixedRows = 1
    Loop
    End With
    rs.Close
    Set rs = Nothing
    cn.Close



End Sub
