VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVerTubosEnCliente 
   Caption         =   "Tubos en Cliente"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Movimientos"
      Height          =   6855
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   11055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grTubos 
         Height          =   6375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   5040
      Picture         =   "frmVerTubosEnCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   1095
   End
End
Attribute VB_Name = "frmVerTubosEnCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With grTubos
        .Cols = 8
        .ColWidth(0) = 0
        .ColWidth(1) = 800
        .ColWidth(2) = 1400
        .ColWidth(3) = 900
        .ColWidth(4) = 400
        .ColWidth(5) = 1000
        .ColWidth(6) = 2500
        .ColWidth(7) = 3500
        
        .ColAlignment(1) = 1
        

    
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT Tubos.idTubo, Tubos.Numero, Articulos.Descripcion, Tubos.Capacidad, UnidadesMedidas.Unidad, Isnull(Tubos.UltimoMovimiento,'') as UltimoMovimiento, IsNull(Tubos.DetalleUltimo,'') as DetalleUltimo, Clientes.Nombre as Propietario,  Tubos.ClienteActual FROM Tubos INNER JOIN UnidadesMedidas ON Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida INNER JOIN Articulos ON Tubos.idArticulo = Articulos.idArticulo INNER JOIN Clientes ON Tubos.Propietario = Clientes.idCliente Where Tubos.ClienteActual = " & idCliente & " order by Tubos.UltimoMovimiento")
    'lblEncontrados = rs.RecordCount
   ' Set grClientes.DataSource = rs

    .Rows = 1
    .TextArray(0) = "idTubo"
    .TextArray(1) = "Número"
    .TextArray(2) = "Descripción"
    .TextArray(3) = "Capacidad"
    .TextArray(4) = ""
    .TextArray(5) = "Fecha"
    .TextArray(6) = "Detalle"
    .TextArray(7) = "Propietario"
   
    
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idTubo
        .TextMatrix(.Rows - 1, 1) = rs!numero
        .TextMatrix(.Rows - 1, 2) = rs!Descripcion
        .TextMatrix(.Rows - 1, 3) = rs!Capacidad
        .TextMatrix(.Rows - 1, 4) = rs!Unidad
        .TextMatrix(.Rows - 1, 5) = rs!UltimoMovimiento
        .TextMatrix(.Rows - 1, 6) = rs!DetalleUltimo
        .TextMatrix(.Rows - 1, 7) = rs!Propietario
        rs.MoveNext
        
        .FixedRows = 1
    Loop
    End With
    rs.Close
    Set rs = Nothing
    cn.Close

End Sub
