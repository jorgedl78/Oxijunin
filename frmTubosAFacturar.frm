VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTubosAFacturar 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   855
      Left            =   3600
      Picture         =   "frmTubosAFacturar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   855
      Left            =   9000
      Picture         =   "frmTubosAFacturar.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   5040
      Picture         =   "frmTubosAFacturar.frx":09DF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   855
      Left            =   9000
      Picture         =   "frmTubosAFacturar.frx":12A9
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grDetalleTubos 
      Height          =   1695
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   0
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      GridColor       =   14737632
      GridColorFixed  =   14737632
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   8295
   End
End
Attribute VB_Name = "frmTubosAFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub grDetalle_Click()

End Sub

Private Sub cmdAgregar_Click()
    frmBuscarTubos.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    grDetalleTubos.Cols = 7
    grDetalleTubos.ColWidth(0) = 0  'idTubo
    grDetalleTubos.ColWidth(1) = 1500 'tubo
    grDetalleTubos.ColWidth(2) = 0 'idArticulo
    grDetalleTubos.ColWidth(3) = 2500 'Gas
    grDetalleTubos.ColWidth(4) = 1000 'capacidad
    grDetalleTubos.ColWidth(5) = 1500 'precio
    grDetalleTubos.ColWidth(6) = 1500 'total
End Sub

