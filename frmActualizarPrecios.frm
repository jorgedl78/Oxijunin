VERSION 5.00
Begin VB.Form frmActualizarPrecios 
   Caption         =   "Actualización de Precios"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPorcentaje 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   1680
      Picture         =   "frmActualizarPrecios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   855
      Left            =   3480
      Picture         =   "frmActualizarPrecios.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox cmRubro 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.ComboBox cmMarca 
      Height          =   315
      ItemData        =   "frmActualizarPrecios.frx":1194
      Left            =   1320
      List            =   "frmActualizarPrecios.frx":119B
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Porcentaje;"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Rubro:"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Marca:"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmActualizarPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If IsNumeric(txtPorcentaje) = False Then MsgBox ("El porcentaje de actualización no es válido"): Exit Sub
    Dim rs As New ADODB.Recordset
    cn.Open
    If cmRubro.ItemData(cmRubro.ListIndex) = 0 Then
        desderubro = 0: hastarubro = 99999
    Else
        desderubro = cmRubro.ItemData(cmRubro.ListIndex): hastarubro = cmRubro.ItemData(cmRubro.ListIndex)
    End If
    If cmMarca.ItemData(cmMarca.ListIndex) = 0 Then
        desdemarca = 0: hastamarca = 99999
    Else
        desdemarca = cmMarca.ItemData(cmMarca.ListIndex): hastamarca = cmMarca.ItemData(cmMarca.ListIndex)
    End If

    Set rs = cn.Execute("ContarArticulosPorRubroyMarca " & desderubro & "," & hastarubro & "," & desdemarca & "," & hastamarca)
    
    Respuesta = MsgBox("¿Esta seguro de actualizar " & rs!Total & " artículos?", vbYesNo, "Atención")
    If Respuesta = vbYes Then
        cn.Execute ("ActualizarPrecios " & Replace(Val(txtPorcentaje.Text), ",", ".") & "," & desderubro & "," & hastarubro & "," & desdemarca & "," & hastamarca)
        cn.Close
    Else
        cn.Close
    End If

End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cn.Open
    cmRubro.AddItem ("Todos")
    cmRubro.ItemData(cmRubro.NewIndex) = 0
    Set rs = cn.Execute("VerRubros")
    Do While rs.EOF = False
        cmRubro.AddItem (rs!Rubro)
        cmRubro.ItemData(cmRubro.NewIndex) = rs!IdRubro
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    cmRubro.ListIndex = 0
    Set rs = cn.Execute("VerMarcas")
    cmMarca.Clear
    cmMarca.AddItem ("Todos")
    cmMarca.ItemData(cmMarca.NewIndex) = 0
    Do While rs.EOF = False
        cmMarca.AddItem (rs!Marca)
        cmMarca.ItemData(cmMarca.NewIndex) = rs!idMarca
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    cmMarca.ListIndex = I
    cn.Close
End Sub

Private Sub Text1_Change()

End Sub

