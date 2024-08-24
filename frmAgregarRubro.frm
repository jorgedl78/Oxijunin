VERSION 5.00
Begin VB.Form frmAgregarRubro 
   Caption         =   "Agregar Rubro"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   720
      Picture         =   "frmAgregarRubro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   2400
      Picture         =   "frmAgregarRubro.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtRubro 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Nuevo Rubro:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmAgregarRubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If txtRubro = "" Then MsgBox ("Debe especificar una descripcion"): Exit Sub
    cn.Open
    cn.Execute ("INSERT INTO Rubros(Rubro) VALUES ('" & txtRubro & "')")
    cn.Close
    frmFichaArticulo.CargarRubros
    For I = 0 To frmFichaArticulo.cmRubro.ListCount - 1
        If frmFichaArticulo.cmRubro.List(I) = txtRubro.Text Then frmFichaArticulo.cmRubro.ListIndex = I
    Next I
    Unload Me
End Sub

