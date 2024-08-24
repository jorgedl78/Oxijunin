VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmArticulosAEtiquetar 
   Caption         =   "Articulos a Etiquetar"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Probar fuentes"
      Height          =   855
      Left            =   3240
      TabIndex        =   5
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   6120
      Picture         =   "frmArticulosAEtiquetar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   4680
      Picture         =   "frmArticulosAEtiquetar.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grArticulos 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10186
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
      Caption         =   "Etiquetas a imprimir"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label lblEncontrados 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6240
      Width           =   735
   End
End
Attribute VB_Name = "frmArticulosAEtiquetar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Total = 0
    With frmPruebaEtiquetas
    .lblCodigo(Total).Visible = True
    .lblDescripcion(Total).Visible = True
    .lblPrecio(Total).Visible = True
    .lblBarras(Total).Visible = True
    .lblNumeroBarras(Total).Visible = True
    .Shape1(Total).Visible = True

    .lblCodigo(Total) = grArticulos.TextMatrix(Total + 1, 0)
    .lblDescripcion(Total) = grArticulos.TextMatrix(Total + 1, 2)
    .lblPrecio(Total) = "$ " & grArticulos.TextMatrix(Total + 1, 3)
    .lblBarras(Total) = "*" & Trim(grArticulos.TextMatrix(Total + 1, 1)) & "*"
    .lblNumeroBarras(Total) = grArticulos.TextMatrix(Total + 1, 1)
    Total = Total + 1
    
    .lblCodigo(Total).Visible = True
    .lblDescripcion(Total).Visible = True
    .lblPrecio(Total).Visible = True
    .lblBarras(Total).Visible = True
    .lblNumeroBarras(Total).Visible = True
    .Shape1(Total).Visible = True

    .lblCodigo(Total) = grArticulos.TextMatrix(Total + 1, 0)
    .lblDescripcion(Total) = grArticulos.TextMatrix(Total + 1, 2)
    .lblPrecio(Total) = "$ " & grArticulos.TextMatrix(Total + 1, 3)
    .lblBarras(Total) = "*" & Trim(grArticulos.TextMatrix(Total + 1, 1)) & "*"
    .lblNumeroBarras(Total) = grArticulos.TextMatrix(Total + 1, 1)
    Total = Total + 1
        
    .lblCodigo(Total).Visible = True
    .lblDescripcion(Total).Visible = True
    .lblPrecio(Total).Visible = True
    .lblBarras(Total).Visible = True
    .lblNumeroBarras(Total).Visible = True
    .Shape1(Total).Visible = True

    .lblCodigo(Total) = grArticulos.TextMatrix(Total + 1, 0)
    .lblDescripcion(Total) = grArticulos.TextMatrix(Total + 1, 2)
    .lblPrecio(Total) = "$ " & grArticulos.TextMatrix(Total + 1, 3)
    .lblBarras(Total) = Trim(grArticulos.TextMatrix(Total + 1, 1))
    .lblNumeroBarras(Total) = grArticulos.TextMatrix(Total + 1, 1)
    Total = Total + 1
        End With
    frmPruebaEtiquetas.Show 1

End Sub

Private Sub cmdImprimir_Click()
    Total = 0
    With frmImprimeEtiquetas
    For I = 1 To grArticulos.Rows - 1
        .lblCodigo(Total).Visible = True
        .lblDescripcion(Total).Visible = True
        .lblPrecio(Total).Visible = True
        .lblBarras(Total).Visible = True
        .lblNumeroBarras(Total).Visible = True
        .Shape1(Total).Visible = True

        .lblCodigo(Total) = grArticulos.TextMatrix(I, 0)
        .lblDescripcion(Total) = grArticulos.TextMatrix(I, 2)
        .lblPrecio(Total) = "$ " & grArticulos.TextMatrix(I, 3)
        .lblBarras(Total) = "*" & Trim(grArticulos.TextMatrix(I, 1)) & "*"
        .lblNumeroBarras(Total) = grArticulos.TextMatrix(I, 1)
        Total = Total + 1
        If Total = 18 Then
            .PrintForm
            Total = 0
        End If
    Next I
    If Total < 18 Then .PrintForm
        '.lblCodigo(2) = grArticulos.TextMatrix(2, 0)
        '.lblDescripcion(2) = grArticulos.TextMatrix(2, 2)
        '.lblPrecio(2) = "$ " & grArticulos.TextMatrix(2, 3)
        '.lblBarras(2) = grArticulos.TextMatrix(2, 1)
        '.lblNumeroBarras(2) = grArticulos.TextMatrix(2, 1)
        '.PrintForm
        '.Show 1
    End With
    cn.Open
    cn.Execute ("UPDATE Articulos set Etiquetar=0")
    cn.Close
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    grArticulos.Cols = 4
    grArticulos.ColWidth(0) = 900
    grArticulos.ColWidth(1) = 1400
    grArticulos.ColWidth(2) = 3500
    grArticulos.ColWidth(3) = 800
    lblEncontrados = 0
    
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("ArticulosAEtiquetar")
    lblEncontrados = rs.RecordCount
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
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub
