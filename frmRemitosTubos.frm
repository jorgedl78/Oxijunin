VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRemitosTubos 
   Caption         =   "Remito de Tubos"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Detalle de Tubos"
      Height          =   5175
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   9855
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   855
         Left            =   8520
         MaskColor       =   &H00000000&
         Picture         =   "frmRemitosTubos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   855
         Left            =   8520
         MaskColor       =   &H00000000&
         Picture         =   "frmRemitosTubos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2040
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grTubos 
         Height          =   4575
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8070
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
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   3600
      Picture         =   "frmRemitosTubos.frx":09DF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   6000
      Picture         =   "frmRemitosTubos.frx":12A9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cabecera del Remito"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   9855
      Begin VB.ComboBox cmMovimientos 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker dateFecha 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   42366
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   9
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   7200
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Fecha 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Movimiento:"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblPuesto 
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblNumero 
         Caption         =   "00000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRemitosTubos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregar_Click()
    buscarTubosPara = "Remito"
    frmBuscarTubos.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If cmMovimientos.ListIndex < 0 Then MsgBox ("Debe seleccionar el Movimiento del remito"): Exit Sub
    Respuesta = MsgBox("¿Esta seguro de Confirmar el Remito?", vbYesNo, "Guardar")
    If Respuesta = vbNo Then Exit Sub
    cn.Open
    
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 2 Then EnDestino = 1
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 3 Then EnDestino = 2
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 4 Then EnDestino = 1
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 12 Then EnDestino = 916
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 13 Then EnDestino = 1
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 14 Then EnDestino = 909
    If cmMovimientos.ItemData(cmMovimientos.ListIndex) = 15 Then EnDestino = 1
    cn.Execute ("insert into RemitosTubos(Fecha,Puesto,Numero,idEstadoTubos) values('" & Format(dateFecha, "dd/mm/yyyy") & "'," & lblPuesto & "," & lblNumero & "," & cmMovimientos.ItemData(cmMovimientos.ListIndex) & ")")
    Set rs = cn.Execute("SELECT MAX(idRemitoTubos) AS Nuevoid FROM RemitosTubos")
    NuevoID = rs!NuevoID
    idComprobante = NuevoID
    
    With grTubos
    For I = 1 To grTubos.Rows - 1
        cn.Execute ("INSERT INTO DetalleRemitoTubos(idTubo, idRemitoTubos) VALUES (" & .TextMatrix(I, 0) & " , " & NuevoID & ")")
        cn.Execute ("UPDATE Tubos set idEstadoTubos=" & cmMovimientos.ItemData(cmMovimientos.ListIndex) & ", ClienteActual= " & EnDestino & ", UltimoMovimiento= '" & Format(dateFecha, "dd/mm/yyyy") & "', DetalleUltimo='Remito " & Format(lblPuesto, "0000") & "-" & Format(lblNumero, "00000000") & "' where idTubo=" & .TextMatrix(I, 0))
        cn.Execute ("INSERT INTO MovimientosTubos (Fecha, Detalle, idTubo) VALUES ('" & Format(dateFecha, "dd/mm/yyyy") & "', '" & cmMovimientos.Text & " Remito " & Format(lblPuesto, "0000") & "-" & Format(lblNumero, "00000000") & "'," & .TextMatrix(I, 0) & ")")
    Next I
    End With

    idRemitoTubo = NuevoID
    RemitoTubos.WindowState = 2
    RemitoTubos.Show 1
    
    'actualizo numero de remito utilizado
    cn.Execute ("UPDATE Parametros set NumeroRemito=NumeroRemito + 1")
    
    cn.Close

    Unload Me

End Sub

Private Sub cmdQuitar_Click()
    If grTubos.Rows = 1 Or grTubos.Row = 0 Then Exit Sub
    Respuesta = MsgBox("¿Está seguro de quitar el Tubo?", vbYesNo, "Borrar")
    If Respuesta = vbNo Then Exit Sub
    If grTubos.Rows > 1 Then
        grTubos.RemoveItem (grTubos.Row)
    Else
        'grTubos.Rows = 0
    End If

End Sub



Private Sub Form_Load()
    dateFecha.Value = Date
    Dim rs As New ADODB.Recordset
    cn.Open
    Set rs = cn.Execute("SELECT idEstadoTubos,Movimiento FROM EstadoTubos WHERE (idEstadoTubos between 3 and 4) or (idEstadoTubos between 12 and 15)")
    Do While rs.EOF = False
        cmMovimientos.AddItem (rs!Movimiento)
        cmMovimientos.ItemData(cmMovimientos.NewIndex) = rs!idEstadoTubos
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set rs = cn.Execute("SELECT PuestoRemito, (NumeroRemito + 1) as Numero from Parametros")
    lblPuesto = Format(rs!PuestoRemito, "0000")
    lblNumero = Format(rs!numero, "00000000")
    cn.Close
    
    'grTubos.Cols = 9
    'grTubos.ColWidth(0) = 0
    'grTubos.ColWidth(1) = 1400
    'grTubos.ColWidth(2) = 2000
    'grTubos.ColWidth(3) = 800
    'grTubos.ColWidth(4) = 800
    'grTubos.ColWidth(5) = 3000
    'grTubos.ColWidth(6) = 1500
    'grTubos.ColWidth(7) = 0
    'grTubos.ColWidth(8) = 5000
    
    'inicializo grilla de tubos
    grTubos.Cols = 7
    grTubos.ColWidth(0) = 0  'idTubo
    grTubos.ColWidth(1) = 0 'idArticulo
    grTubos.ColWidth(2) = 1500 'tubo
    grTubos.ColWidth(3) = 2500 'Gas
    grTubos.ColWidth(4) = 1000 'capacidad
    grTubos.ColWidth(5) = 1500 'Precio
    grTubos.ColWidth(6) = 1500 'Estado
    grTubos.TextArray(0) = "idTubo"
    grTubos.TextArray(1) = "idArticulo"
    grTubos.TextArray(2) = "Número"
    grTubos.TextArray(3) = "Gas"
    grTubos.TextArray(4) = "Capacidad"
    grTubos.TextArray(5) = "Precio"
    grTubos.TextArray(6) = "Estado"
    
End Sub
