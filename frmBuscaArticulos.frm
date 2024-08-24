VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBuscaArticulos 
   BackColor       =   &H00404040&
   ClientHeight    =   7710
   ClientLeft      =   5850
   ClientTop       =   330
   ClientWidth     =   10500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdElejir 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   4920
      Picture         =   "frmBuscaArticulos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   7920
      Picture         =   "frmBuscaArticulos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   855
   End
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grArticulos 
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9128
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
      FontWidthFixed  =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   3
      _Band(0).TextStyleHeader=   0
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
      Left            =   8520
      TabIndex        =   8
      Top             =   6000
      Width           =   1695
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
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Width           =   7695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   10095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   2895
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
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   975
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
      Left            =   2040
      TabIndex        =   3
      Top             =   6840
      Width           =   1455
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
      Left            =   360
      TabIndex        =   2
      Top             =   6840
      Width           =   1575
   End
End
Attribute VB_Name = "frmBuscaArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdElejir_Click()
    If grArticulos.Rows <= 1 Then Exit Sub
    frmFacturador.txtBarras = grArticulos.TextMatrix(grArticulos.Row, 1)
    Unload Me
End Sub

Private Sub Form_Load()
    grArticulos.Cols = 4
    grArticulos.ColWidth(0) = 900
    grArticulos.ColWidth(1) = 1800
    grArticulos.ColWidth(2) = 5900
    grArticulos.ColWidth(3) = 1000
End Sub

Private Sub grArticulos_DblClick()
    cmdElejir_Click
End Sub

Private Sub grArticulos_KeyPress(KeyAscii As Integer)
    cmdElejir_Click
End Sub

Private Sub grArticulos_RowColChange()
    lblDescripcion = grArticulos.TextMatrix(grArticulos.Row, 2)
    lblPrecio = Format(grArticulos.TextMatrix(grArticulos.Row, 3), "$0.00")
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then cmdCancelar_Click
    If KeyAscii <> 13 Then Exit Sub
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("BuscaArticulos '" & txtBusca & "'")
    txtEncontrados = rs.RecordCount
   ' Set grArticulos.DataSource = rs
    With grArticulos
    .Rows = 1
    .TextArray(0) = "Codigo"
    .TextArray(1) = "Codigo Barras"
    .TextArray(2) = "Descripción"
    .TextArray(3) = "Precio"
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idArticulo
        .TextMatrix(.Rows - 1, 1) = rs!CodBar
        .TextMatrix(.Rows - 1, 2) = rs!Descripcion
        .TextMatrix(.Rows - 1, 3) = Format(rs!Venta, "0.00")
        rs.MoveNext
        .FixedRows = 1
    Loop
    End With
    If rs.RecordCount > 0 Then
        grArticulos.SetFocus
        grArticulos.Col = 2
        grArticulos_RowColChange
    Else
        txtBusca.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub

