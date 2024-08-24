VERSION 5.00
Begin VB.Form frmAgregarMarca 
   Caption         =   "Agregar Marca"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   840
      Picture         =   "frmAgregarMarca.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   2520
      Picture         =   "frmAgregarMarca.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtTipo 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Nuevo Tipo:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmAgregarMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If txtTipo = "" Then MsgBox ("Debe especificar una descripcion"): Exit Sub
    cn.Open
    cn.Execute ("INSERT INTO Marcas(Marca) VALUES ('" & txtTipo & "')")
    cn.Close
    frmFichaArticulo.CargarMarcas
    For I = 0 To frmFichaArticulo.cmMarca.ListCount - 1
        If frmFichaArticulo.cmMarca.List(I) = txtTipo Then frmFichaArticulo.cmMarca.ListIndex = I
    Next I
    Unload Me
End Sub

