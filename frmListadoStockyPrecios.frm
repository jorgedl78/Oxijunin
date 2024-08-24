VERSION 5.00
Begin VB.Form frmListadoStockyPrecios 
   Caption         =   "Listado de Stock y Precios"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   2055
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   6015
      Begin VB.ComboBox cmRubro 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   3735
      End
      Begin VB.ComboBox cmMarca 
         Height          =   315
         ItemData        =   "frmListadoStockyPrecios.frx":0000
         Left            =   960
         List            =   "frmListadoStockyPrecios.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Rubro:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdImprimirDetalle 
      Caption         =   "Imprimir Detalle de Totales"
      Height          =   1095
      Left            =   2400
      Picture         =   "frmListadoStockyPrecios.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "frmListadoStockyPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimirDetalle_Click()

    idDesdeRubro = 0
    idHastRubro = 0
    idDesdearca = 0
    idHastaMarca = 0
    If cmRubro.ItemData(cmRubro.ListIndex) = 0 Then
        idDesdeRubro = 0: idHastaRubro = 99999
    Else
        idDesdeRubro = cmRubro.ItemData(cmRubro.ListIndex): idHastaRubro = cmRubro.ItemData(cmRubro.ListIndex)
    End If
    
    If cmMarca.ItemData(cmMarca.ListIndex) = 0 Then
        idDesdeMarca = 0: idHastaMarca = 99999
    Else
        idDesdeMarca = cmMarca.ItemData(cmMarca.ListIndex): idHastaMarca = cmMarca.ItemData(cmMarca.ListIndex)
    End If

    
    cn.Open
    With ListadoStockYPrecios.Sections("Sección4")
        .Controls("lblRubro").Caption = cmRubro.Text
        .Controls("lblMarca").Caption = cmMarca.Text
    End With
       
        
    Set rs = cn.Execute("SELECT Articulos.CodBar, Marcas.Marca, Rubros.Rubro, Articulos.Descripcion, Articulos.Venta, Articulos.VentaRevendedor, Articulos.Stock FROM Articulos INNER JOIN Marcas ON Articulos.idMarca = Marcas.idMarca INNER JOIN Rubros ON Articulos.idRubro = Rubros.idRubro where articulos.idMarca between " & idDesdeMarca & " and " & idHastaMarca & " and Articulos.idRubro between " & idDesdeRubro & " and " & idHastaRubro & " ORDER BY Marcas.Marca, Rubros.Rubro, cast(Articulos.CodBar as numeric)")
    Set ListadoStockYPrecios.DataSource = rs
    ListadoStockYPrecios.WindowState = 2
    
    ListadoStockYPrecios.Show 1
    
    cn.Close

End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cn.Open
    Set rs = cn.Execute("VerRubros")
    cmRubro.AddItem ("Todos")
    cmRubro.ItemData(cmRubro.NewIndex) = 0
    Do While rs.EOF = False
        cmRubro.AddItem (rs!Rubro)
        cmRubro.ItemData(cmRubro.NewIndex) = rs!IdRubro
        rs.MoveNext
    Loop
    cmRubro.ListIndex = 0
    
    Set rs = Nothing
    Set rs = cn.Execute("VerMarcas")
    cmMarca.Clear
    cmMarca.AddItem ("Todas")
    cmMarca.ItemData(cmMarca.NewIndex) = 0
    
    Do While rs.EOF = False
        cmMarca.AddItem (rs!Marca)
        cmMarca.ItemData(cmMarca.NewIndex) = rs!idMarca
        rs.MoveNext
    Loop
    cmMarca.ListIndex = 0
    rs.Close
    Set rs = Nothing
    cn.Close
End Sub
