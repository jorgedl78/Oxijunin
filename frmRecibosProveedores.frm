VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecibosProveedores 
   Caption         =   "Recibos de Proveedores"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   3120
      Picture         =   "frmRecibosProveedores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Height          =   855
      Left            =   1920
      Picture         =   "frmRecibosProveedores.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   5880
      Picture         =   "frmRecibosProveedores.frx":09DF
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   855
         Left            =   7080
         Picture         =   "frmRecibosProveedores.frx":12A9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker CalendarHasta 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   83034113
         CurrentDate     =   42391
      End
      Begin MSComCtl2.DTPicker CalendarDesde 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   83099649
         CurrentDate     =   42391
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resultados"
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   10335
      Begin VB.TextBox txtDetalle 
         Height          =   5535
         Left            =   4440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   5775
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grRecibos 
         Height          =   5535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
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
   End
End
Attribute VB_Name = "frmRecibosProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnular_Click()
    If grRecibos.TextMatrix(grRecibos.Row, 4) = "Anulado" Then MsgBox ("El Recibo ya se encuentra anulado"): Exit Sub
   
    Respuesta = MsgBox("¿Confirma la anulación del Recibo " & grRecibos.TextMatrix(grRecibos.Row, 2) & "?", vbYesNo, "Atencion!")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    cn.Execute ("UPDATE RecibosProveedor SET Anulado='Anulado' WHERE idReciboProveedor=" & grRecibos.TextMatrix(grRecibos.Row, 0))
    'MsgBox ("AgregarCuentaCorriente '" & Format(Date, "yyyy/mm/dd") & "','Anulación Recibo " & grRecibos.TextMatrix(grRecibos.Row, 2) & "'," & grRecibos.TextMatrix(grRecibos.Row, 3) & ",0.00," & idCliente & ",'Anu'," & grRecibos.TextMatrix(grRecibos.Row, 0))
    cn.Execute ("AgregarCuentaCorrienteProveedor '" & Format(Date, "yyyy/mm/dd") & "','Anulación Recibo " & grRecibos.TextMatrix(grRecibos.Row, 2) & "'," & Replace(grRecibos.TextMatrix(grRecibos.Row, 3), ",", ".") & ",0.00," & idProveedor & ",'Anu'," & grRecibos.TextMatrix(grRecibos.Row, 0))
    cn.Close
    lineaActual = grRecibos.Row
    Consultar
    grRecibos.Row = lineaActual
End Sub

Private Sub cmdConsultar_Click()
    Consultar
End Sub

Private Sub cmdImprimir_Click()
    If grRecibos.TextMatrix(grRecibos.Row, 4) = "Anulado" Then MsgBox ("No se puede reimprimir ya que el recibo se encuentra anulado"): Exit Sub
    idRecibo = grRecibos.TextMatrix(grRecibos.Row, 0)
    cn.Open
    Set rs = cn.Execute("SELECT max(idRecibo) as UltimoRecibo FROM Recibos")
    Set ImprimeRecibo.DataSource = rs

    ImprimeRecibo.WindowState = 2
    ImprimeRecibo.Show 1
    cn.Close
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CalendarDesde.Value = Date
    CalendarHasta.Value = Date
    
    With grRecibos
        .Cols = 6
        .ColWidth(0) = 0
        .ColWidth(1) = 1000
        .ColWidth(2) = 700
        .ColWidth(3) = 1000
        .ColWidth(4) = 800
        .ColWidth(5) = 0
        .Rows = 1
        .TextArray(0) = "idRecibo"
        .TextArray(1) = "Fecha"
        .TextArray(2) = "Número"
        .TextArray(3) = "Importe"
        .TextArray(4) = "Detalle"
        .TextArray(5) = "Anulado"
    End With
End Sub

Private Sub Consultar()
    cn.Open
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT idReciboProveedor, Fecha, Numero, Importe, IsNull(Anulado,'') as Anulado, Detalle  from RecibosProveedor WHERE idProveedor BETWEEN " & idProveedor & " and " & idProveedor & " AND Fecha BETWEEN '" & CalendarDesde.Value & "' AND '" & CalendarHasta.Value & "'")

    With grRecibos
    .Rows = 1
    .TextArray(0) = "idRecibo"
    .TextArray(1) = "Fecha"
    .TextArray(2) = "Número"
    .TextArray(3) = "Importe"
    .TextArray(4) = "Anulado"
    .TextArray(5) = "Detalle"
    
    Do While rs.EOF = False
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rs!idReciboProveedor
        .TextMatrix(.Rows - 1, 1) = rs!Fecha
        .TextMatrix(.Rows - 1, 2) = rs!numero
        .TextMatrix(.Rows - 1, 3) = Format(rs!Importe, "0.00")
        .TextMatrix(.Rows - 1, 4) = rs!Anulado
        .TextMatrix(.Rows - 1, 5) = rs!Detalle
        rs.MoveNext
        
        .FixedRows = 1
    Loop
    End With
    If grRecibos.Rows > 1 Then grRecibos_RowColChange
    rs.Close
    Set rs = Nothing
    cn.Close


End Sub


Private Sub grRecibos_RowColChange()
    txtDetalle.Text = grRecibos.TextMatrix(grRecibos.Row, 5)
End Sub

