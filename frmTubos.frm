VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTubos 
   Caption         =   "Tubos"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   13035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Ocultar"
      Height          =   1095
      Left            =   8880
      Picture         =   "frmTubos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "Esta acciòn inactiva al tubo y ya no queda visible"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdVerMovimientos 
      Caption         =   "Movimientos"
      Height          =   1095
      Left            =   3240
      Picture         =   "frmTubos.frx":0115
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemitosTubos 
      Caption         =   "Remitos de Tubos"
      Height          =   1095
      Left            =   4920
      Picture         =   "frmTubos.frx":09DF
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   1095
      Left            =   11040
      Picture         =   "frmTubos.frx":0D40
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtBusca 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   1095
      Left            =   7440
      Picture         =   "frmTubos.frx":160A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grTubos 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9763
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
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblEncontrados 
      Caption         =   "lblEncontrados"
      Height          =   255
      Left            =   9360
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Total de Tubos:"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   240
      Width           =   1335
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
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   9015
   End
End
Attribute VB_Name = "frmTubos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    Estado = "Agregando"
    frmFichaTubo.Show 1
End Sub

Private Sub cmdQuitar_Click()
    If grTubos.Rows <= 1 Then Exit Sub
    Respuesta = MsgBox("¿Confirma la ocultación del tubo " & grTubos.TextMatrix(grTubos.Row, 1) & "?", vbYesNo, "Atencion!")
    If Respuesta = vbNo Then Exit Sub
    
    cn.Open
    Set rs = cn.Execute("UPDATE Tubos SET Inactivo = 1 WHERE idTubo = " & grTubos.TextMatrix(grTubos.Row, 0))
    cn.Close
    BuscarTubos
End Sub

Private Sub cmdRemitosTubos_Click()
    frmRemitosTubos.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdVerMovimientos_Click()
    idTubo = grTubos.TextMatrix(grTubos.Row, 0)
    frmConsultaMovimientoDeTubo.Show 1
End Sub

Private Sub Form_Load()
    grTubos.Cols = 9
    grTubos.ColWidth(0) = 0
    grTubos.ColWidth(1) = 1400
    grTubos.ColWidth(2) = 2000
    grTubos.ColWidth(3) = 800
    grTubos.ColWidth(4) = 800
    grTubos.ColWidth(5) = 3000
    grTubos.ColWidth(6) = 1500
    grTubos.ColWidth(7) = 0
    grTubos.ColWidth(8) = 5000
    lblEncontrados = 0
    lblDescripcion = ""
End Sub



Private Sub grTubos_DblClick()
    EditarTubo
End Sub

Private Sub grTubos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then EditarTubo
End Sub

Private Sub grTubos_RowColChange()
    lblDescripcion = grTubos.TextMatrix(grTubos.Row, 1)
End Sub


Private Sub txtBusca_Change()
    'BuscarArticulos
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then cmdSalir_Click
    If KeyAscii <> 13 Then Exit Sub
    BuscarTubos
End Sub
Sub EditarTubo()
    If grTubos.Rows <= 1 Then Exit Sub
    idTubo = grTubos.TextMatrix(grTubos.Row, 0)
    TuboSeleccionado = grTubos.Row
    Saltar = 1
    Estado = "Modificando"
    frmFichaTubo.Show 1
    If Saltar = 0 Then
        BuscarTubos
        grTubos.Row = TuboSeleccionado
        grTubos_RowColChange
    End If
End Sub

Sub BuscarTubos()
    cn.Open
    Dim rs As ADODB.Recordset

    Set rs = cn.Execute("SELECT Tubos.idTubo, Tubos.Numero, Articulos.Descripcion, Tubos.Capacidad, UnidadesMedidas.Unidad, EstadoTubos.Movimiento, EstadoTubos.Estado, clientes.idCliente, clientes.Nombre FROM Articulos INNER JOIN Tubos ON Articulos.idArticulo = Tubos.idArticulo INNER JOIN UnidadesMedidas ON Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida INNER JOIN EstadoTubos ON Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos inner join Clientes on clientes.idCliente=tubos.ClienteActual  where Inactivo = 0 AND Numero like '%" & txtBusca & "%' order by Numero")
    lblEncontrados = rs.RecordCount
   ' Set grArticulos.DataSource = rs
    With grTubos
    .Rows = 1
    .TextArray(0) = "idTubo"
    .TextArray(1) = "Numero"
    .TextArray(2) = "Descripción"
    .TextArray(3) = "Capacidad"
    .TextArray(4) = "Unidad"
    .TextArray(5) = "Movimiento "
    .TextArray(6) = "Estado"
    .TextArray(7) = "idCliente"
    .TextArray(8) = "Cliente"
    
    .ColWidth(0) = 0
    
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idTubo
        .TextMatrix(.Rows - 1, 1) = rs!numero
        .TextMatrix(.Rows - 1, 2) = rs!Descripcion
        .TextMatrix(.Rows - 1, 3) = Format(rs!Capacidad, "0.00")
        .TextMatrix(.Rows - 1, 4) = rs!Unidad
        .TextMatrix(.Rows - 1, 5) = rs!Movimiento
        .TextMatrix(.Rows - 1, 6) = rs!Estado
        .TextMatrix(.Rows - 1, 7) = rs!idCliente
        .TextMatrix(.Rows - 1, 8) = rs!Nombre

        
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With
    If rs.RecordCount > 0 Then
        'grArticulos.SetFocus
        'grArticulos_RowColChange
    Else
        txtBusca.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub

