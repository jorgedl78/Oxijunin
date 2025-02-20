VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBuscarTubos 
   BackColor       =   &H00404040&
   Caption         =   "Tubos"
   ClientHeight    =   7605
   ClientLeft      =   12465
   ClientTop       =   2190
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBusca 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   6000
      Picture         =   "frmBuscarTubos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdElejir 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   4080
      Picture         =   "frmBuscarTubos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grTubos 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8916
      _Version        =   393216
      BackColor       =   -2147483633
      ForeColor       =   0
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   4210752
      GridColor       =   0
      GridColorFixed  =   0
      WordWrap        =   -1  'True
      TextStyle       =   3
      TextStyleFixed  =   4
      FocusRect       =   2
      HighLight       =   0
      FillStyle       =   1
      ScrollBars      =   2
      MergeCells      =   2
      AllowUserResizing=   2
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontWidthFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   3
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Encontrados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label txtEncontrados 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   5175
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   8175
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Width           =   6015
   End
   Begin VB.Label lblPrecio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   5880
      Width           =   1695
   End
End
Attribute VB_Name = "frmBuscarTubos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With grTubos
        .Cols = 7
        .ColWidth(0) = 0  'idTubo
        .ColWidth(1) = 0 'idArticulo
        .ColWidth(2) = 1500 'Numero de tubo
        .ColWidth(3) = 2500 'Gas
        .ColWidth(4) = 800 'capacidad
        .ColWidth(5) = 800 'Unidad de medida
        .ColWidth(6) = 2000 'Precio
    End With
End Sub

Private Sub grTubos_DblClick()
    cmdElejir_Click
End Sub
Private Sub cmdElejir_Click()
    If grTubos.Rows <= 1 Then Exit Sub
    If buscarTubosPara = "Vender" Then
        With frmFacturador.grDetalleTubos
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = grTubos.TextMatrix(grTubos.Row, 0)
            .TextMatrix(.Rows - 1, 1) = grTubos.TextMatrix(grTubos.Row, 1)
            .TextMatrix(.Rows - 1, 2) = grTubos.TextMatrix(grTubos.Row, 2)
            .TextMatrix(.Rows - 1, 3) = grTubos.TextMatrix(grTubos.Row, 3)
            .TextMatrix(.Rows - 1, 4) = grTubos.TextMatrix(grTubos.Row, 4)
            .TextMatrix(.Rows - 1, 5) = grTubos.TextMatrix(grTubos.Row, 5)
            .TextMatrix(.Rows - 1, 6) = grTubos.TextMatrix(grTubos.Row, 4) * grTubos.TextMatrix(grTubos.Row, 6)
        End With
        frmFacturador.txtCantidad.Text = grTubos.TextMatrix(grTubos.Row, 4)
        frmFacturador.txtBarras.Text = grTubos.TextMatrix(grTubos.Row, 1)
        Unload Me
        frmFacturador.CargarDetalle
    End If
    If buscarTubosPara = "Devolver" Then
        With frmFacturador.grTubosDevueltos
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = grTubos.TextMatrix(grTubos.Row, 0)
            .TextMatrix(.Rows - 1, 1) = grTubos.TextMatrix(grTubos.Row, 1)
            .TextMatrix(.Rows - 1, 2) = grTubos.TextMatrix(grTubos.Row, 2)
            .TextMatrix(.Rows - 1, 3) = grTubos.TextMatrix(grTubos.Row, 3)
            .TextMatrix(.Rows - 1, 4) = grTubos.TextMatrix(grTubos.Row, 4)
            .TextMatrix(.Rows - 1, 5) = grTubos.TextMatrix(grTubos.Row, 5)
            '.TextMatrix(.Rows - 1, 6) = grTubos.TextMatrix(grTubos.Row, 4) * grTubos.TextMatrix(grTubos.Row, 6)
        End With
        frmFacturador.txtCantidad.Text = grTubos.TextMatrix(grTubos.Row, 4)
        frmFacturador.txtBarras.Text = grTubos.TextMatrix(grTubos.Row, 1)
        Unload Me
        'frmFacturador.CargarDetalle
    End If
    
    If buscarTubosPara = "Remito" Then
        With frmRemitosTubos.grTubos
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = grTubos.TextMatrix(grTubos.Row, 0)
            .TextMatrix(.Rows - 1, 1) = grTubos.TextMatrix(grTubos.Row, 1)
            .TextMatrix(.Rows - 1, 2) = grTubos.TextMatrix(grTubos.Row, 2)
            .TextMatrix(.Rows - 1, 3) = grTubos.TextMatrix(grTubos.Row, 3)
            .TextMatrix(.Rows - 1, 4) = grTubos.TextMatrix(grTubos.Row, 4)
            .TextMatrix(.Rows - 1, 5) = grTubos.TextMatrix(grTubos.Row, 5)
            .TextMatrix(.Rows - 1, 6) = grTubos.TextMatrix(grTubos.Row, 6)
            End With
        Unload Me
    End If
    
End Sub

Private Sub grTubos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdElejir_Click
    End If
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then cmdCancelar_Click
    If KeyAscii <> 13 Then Exit Sub
    If buscarTubosPara = "Vender" Then BuscarTubosParaVender
    If buscarTubosPara = "Remito" Then BuscarTubosParaRemito
    If buscarTubosPara = "Devolver" Then BuscarTubosParaDevolver
End Sub

Private Sub BuscarTubosParaVender()
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT     Tubos.idTubo, Articulos.idArticulo, Tubos.Numero, Articulos.Descripcion, Tubos.Capacidad, UnidadesMedidas.Unidad, Articulos.Venta " _
                        & " FROM Articulos INNER JOIN " _
                        & " Tubos ON Articulos.idArticulo = Tubos.idArticulo INNER JOIN " _
                        & " EstadoTubos ON Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos AND Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos AND " _
                        & " Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos INNER JOIN " _
                        & " UnidadesMedidas ON Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND " _
                        & " Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND " _
                        & " Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida " _
                        & " WHERE (Tubos.Numero like '" & txtBusca & "%') AND ((Tubos.idEstadoTubos = 4) OR (Tubos.idEstadoTubos = 13)) AND Inactivo = 0 " _
                        & " ORDER BY Tubos.Numero")
    
    txtEncontrados = rs.RecordCount
   ' Set grArticulos.DataSource = rs
    With grTubos
    .Rows = 1
    .TextArray(2) = "Numero"
    .TextArray(3) = "Gas"
    .TextArray(4) = "Capacidad"
    .TextArray(5) = "Unidad"
    .TextArray(6) = "Precio"
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idTubo
        .TextMatrix(.Rows - 1, 1) = rs!idArticulo
        .TextMatrix(.Rows - 1, 2) = rs!numero
        .TextMatrix(.Rows - 1, 3) = rs!Descripcion
        .TextMatrix(.Rows - 1, 4) = Format(rs!Capacidad, "0.00")
        .TextMatrix(.Rows - 1, 5) = rs!Unidad
        .TextMatrix(.Rows - 1, 6) = Format(rs!Venta, "0.00")
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With

    If rs.RecordCount > 0 Then
        grTubos.SetFocus
        grTubos.Col = 2
        'grTubos_RowColChange
    Else
        txtBusca.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub
Private Sub BuscarTubosParaDevolver()
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT     Tubos.idTubo, Articulos.idArticulo, Tubos.Numero, Articulos.Descripcion, Tubos.Capacidad, UnidadesMedidas.Unidad, Articulos.Venta " _
                        & " FROM Articulos INNER JOIN " _
                        & " Tubos ON Articulos.idArticulo = Tubos.idArticulo INNER JOIN " _
                        & " EstadoTubos ON Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos AND Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos AND " _
                        & " Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos INNER JOIN " _
                        & " UnidadesMedidas ON Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND " _
                        & " Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND " _
                        & " Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida " _
                        & " WHERE (Tubos.Numero like '" & txtBusca & "%') AND (Tubos.idEstadoTubos = 1) AND (Tubos.ClienteActual = " & idCliente & ") " _
                        & " ORDER BY Tubos.Numero")
    
    txtEncontrados = rs.RecordCount
   ' Set grArticulos.DataSource = rs
    With grTubos
    .Rows = 1
    .TextArray(2) = "Numero"
    .TextArray(3) = "Gas"
    .TextArray(4) = "Capacidad"
    .TextArray(5) = "Unidad"
    .TextArray(6) = "Precio"
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idTubo
        .TextMatrix(.Rows - 1, 1) = rs!idArticulo
        .TextMatrix(.Rows - 1, 2) = rs!numero
        .TextMatrix(.Rows - 1, 3) = rs!Descripcion
        .TextMatrix(.Rows - 1, 4) = Format(rs!Capacidad, "0.00")
        .TextMatrix(.Rows - 1, 5) = rs!Unidad
        .TextMatrix(.Rows - 1, 6) = Format(rs!Venta, "0.00")
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With

    If rs.RecordCount > 0 Then
        grTubos.SetFocus
        grTubos.Col = 2
        'grTubos_RowColChange
    Else
        txtBusca.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub

Private Sub BuscarTubosParaRemito()
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT     Tubos.idTubo, Articulos.idArticulo, Tubos.Numero, Articulos.Descripcion, Tubos.Capacidad, UnidadesMedidas.Unidad, Articulos.Venta, EstadoTubos.Estado " _
                        & " FROM Articulos INNER JOIN " _
                        & " Tubos ON Articulos.idArticulo = Tubos.idArticulo INNER JOIN " _
                        & " EstadoTubos ON Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos AND Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos AND " _
                        & " Tubos.idEstadoTubos = EstadoTubos.idEstadoTubos INNER JOIN " _
                        & " UnidadesMedidas ON Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND " _
                        & " Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida AND " _
                        & " Tubos.idUnidadMedida = UnidadesMedidas.idUnidadMedida " _
                        & " WHERE (Tubos.Numero like '" & txtBusca & "%') " _
                        & " ORDER BY Tubos.Numero")
    
    txtEncontrados = rs.RecordCount
   ' Set grArticulos.DataSource = rs
    With grTubos
    .Rows = 1
    .TextArray(2) = "Numero"
    .TextArray(3) = "Gas"
    .TextArray(4) = "Capacidad"
    .TextArray(5) = "Unidad"
    .TextArray(6) = "Estado"
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idTubo
        .TextMatrix(.Rows - 1, 1) = rs!idArticulo
        .TextMatrix(.Rows - 1, 2) = rs!numero
        .TextMatrix(.Rows - 1, 3) = rs!Descripcion
        .TextMatrix(.Rows - 1, 4) = Format(rs!Capacidad, "0.00")
        .TextMatrix(.Rows - 1, 5) = rs!Unidad
        .TextMatrix(.Rows - 1, 6) = Format(rs!Estado, "0.00")
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With

    If rs.RecordCount > 0 Then
        grTubos.SetFocus
        grTubos.Col = 2
        'grTubos_RowColChange
    Else
        txtBusca.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub

