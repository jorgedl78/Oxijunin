VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmArticulos 
   Caption         =   "Artículos"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado Por"
      Height          =   615
      Left            =   5160
      TabIndex        =   13
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton optDescripcion 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optCodigo 
         Caption         =   "Código"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdUltimo 
      Caption         =   "Ultimo Ingresado"
      Height          =   1095
      Left            =   2280
      Picture         =   "frmArticulos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdDetalleAjustes 
      Caption         =   "Detalle de Ajustes"
      Height          =   1095
      Left            =   6480
      Picture         =   "frmArticulos.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdAjuste 
      Caption         =   "Ajustar Stock"
      Height          =   1095
      Left            =   5040
      Picture         =   "frmArticulos.frx":12CC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrecios 
      Caption         =   "Precios"
      Height          =   1095
      Left            =   3600
      Picture         =   "frmArticulos.frx":162D
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton frmEtiquetas 
      Caption         =   "Etiquetas"
      Enabled         =   0   'False
      Height          =   975
      Left            =   360
      Picture         =   "frmArticulos.frx":7E7F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   1095
      Left            =   8160
      Picture         =   "frmArticulos.frx":81E0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtBusca 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grArticulos 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   10575
      _ExtentX        =   18653
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   1095
      Left            =   9480
      Picture         =   "frmArticulos.frx":8AAA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1215
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
      TabIndex        =   6
      Top             =   6480
      Width           =   9015
   End
   Begin VB.Label Label2 
      Caption         =   "Total de Articulos:"
      Height          =   255
      Left            =   8040
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblEncontrados 
      Caption         =   "lblEncontrados"
      Height          =   255
      Left            =   9480
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    Estado = "Agregando"
    frmFichaArticulo.Show 1
End Sub

Private Sub cmdAjuste_Click()
    If grArticulos.Rows <= 1 Then Exit Sub
    idArticulo = grArticulos.TextMatrix(grArticulos.Row, 0)
    ArticuloSeleccionado = grArticulos.Row
    frmAjusteDeStock.Show 1
End Sub

Private Sub cmdDetalleAjustes_Click()
    If grArticulos.Rows <= 1 Then Exit Sub
    idArticulo = grArticulos.TextMatrix(grArticulos.Row, 0)
    frmDetalleAjustesDeStock.lblDescripcion = grArticulos.TextMatrix(grArticulos.Row, 2)
    ArticuloSeleccionado = grArticulos.Row
    frmDetalleAjustesDeStock.Show 1
End Sub

Private Sub cmdPrecios_Click()
    frmActualizarPrecios.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub cmdUltimo_Click()
  cn.Open
  Set rs = cn.Execute("select CodBar, Descripcion from Articulos where CodBar=(SELECT max(cast (CodBar as numeric)) from Articulos)")
  MsgBox ("Mayor código utilizado: " & rs!CodBar & " - " & rs!Descripcion)
  cn.Close
End Sub

Private Sub Form_Load()
    grArticulos.Cols = 8
    grArticulos.ColWidth(0) = 900
    grArticulos.ColWidth(1) = 1400
    grArticulos.ColWidth(2) = 3500
    grArticulos.ColWidth(3) = 800
    grArticulos.ColWidth(4) = 800
    grArticulos.ColWidth(5) = 600
    grArticulos.ColWidth(6) = 1200
    grArticulos.ColWidth(7) = 1200
    lblEncontrados = 0
    lblDescripcion = ""
End Sub

Private Sub frmEtiquetas_Click()
    frmArticulosAEtiquetar.Show 1
End Sub

Private Sub grArticulos_DblClick()
    EditarArticulo
End Sub

Private Sub grArticulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then EditarArticulo
End Sub

Private Sub grArticulos_RowColChange()
    lblDescripcion = grArticulos.TextMatrix(grArticulos.Row, 2)
End Sub

Private Sub optCodigo_Click()
    BuscarArticulos
End Sub


Private Sub optDescripcion_Click()
    BuscarArticulos
End Sub

Private Sub txtBusca_Change()
    'BuscarArticulos
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then cmdSalir_Click
    If KeyAscii <> 13 Then Exit Sub
    BuscarArticulos
End Sub
Sub EditarArticulo()
    If grArticulos.Rows <= 1 Then Exit Sub
    'MsgBox (grArticulos.TextMatrix(grArticulos.Row, 0))
    idArticulo = grArticulos.TextMatrix(grArticulos.Row, 0)
    ArticuloSeleccionado = grArticulos.Row
    Saltar = 1
    Estado = "Modificando"
    frmFichaArticulo.Show 1
    If Saltar = 0 Then
        BuscarArticulos
        grArticulos.Row = ArticuloSeleccionado
        grArticulos_RowColChange
    End If
End Sub

Sub BuscarArticulos()
    cn.Open
    Dim rs As ADODB.Recordset
    ordenaPor = "Descripcion"
    If optCodigo.Value = True Then
        ordenaPor = "idArticulo"
    End If

    Set rs = cn.Execute("ABMArticulos '" & txtBusca & "'," & ordenaPor)
    lblEncontrados = rs.RecordCount
   ' Set grArticulos.DataSource = rs
    With grArticulos
    .Rows = 1
    .TextArray(0) = "Codigo"
    .TextArray(1) = "Codigo Barras"
    .TextArray(2) = "Descripción"
    .TextArray(3) = "Precio"
    .TextArray(4) = "Costo"
    .TextArray(5) = "Stock"
    .TextArray(6) = "Rubro"
    .TextArray(7) = "Marca"
    
    .ColWidth(0) = 0
    
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idArticulo
        .TextMatrix(.Rows - 1, 1) = rs!CodBar
        .TextMatrix(.Rows - 1, 2) = rs!Descripcion
        .TextMatrix(.Rows - 1, 3) = Format(rs!Venta, "0.00")
        .TextMatrix(.Rows - 1, 4) = Format(rs!Costo, "0.00")
        .TextMatrix(.Rows - 1, 5) = Val(rs!Stock)
        .TextMatrix(.Rows - 1, 6) = rs!Rubro
        .TextMatrix(.Rows - 1, 7) = rs!Marca
        
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
