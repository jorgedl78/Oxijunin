Attribute VB_Name = "Module1"
Public cn As ADODB.Connection
Public cnDataShape As ADODB.Connection
Public idUsuario As Integer
Public Usuario As String
Public idArticulo As Double
Public idCliente As Integer
Public idProveedor As Integer
Public idTubo As Double
Public idRecibo As Integer
Public CerroCaja As Integer
Public PidiendoPrecio As Integer
Public Saltar As Integer
Public Estado As String
Public EligiendoCliente As Integer
Public portscan As Integer
Public portfiscal As Integer
Public idUsuarioPermiso As Integer
Public idComprobante As Integer
Public CAE As String
Public VencimientoCAE As String
Public cbte_nro As Integer
Public punto_vta As Integer
Public buscarTubosPara As String
Public idRemitoTubo As Double
Public facturaOriginal As String
Public fechaDesde As Date
Public fechaHasta As Date
Public condicionComprobante As String
Public codigo_QR64 As String

'Para usar archivos ini
Declare Function GetPrivateProfileInt Lib "KERNEL32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'funciones para que el formulario se vea en la barra de tareas
'Public Const WS_EX_APPWINDOW As Long = &H40000
'Public Const GWL_EXSTYLE As Long = (-20)
'Public Const SW_HIDE As Long = 0
'Public Const SW_SHOW As Long = 5
'Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Public m_bActivated As Boolean


Sub Main()
    'chequear configuracion regional
    'SeparadorDecimal = Format(0.1, "#. #")
    'SeparadorDecimal = IIf(InStr(SeparadorDecimal, ","), ",", ".")
    'If SeparadorDecimal = "," Then
    '   MsgBox ("La configuracion regional no es la recomendada" & Chr(13) & "Debe configurar el punto para separador decimal y la coma para separador de miles"): Exit Sub
    'End If
    
    Dim i As Integer
    Dim Est As String
    On Error GoTo noInicia
    Est = String$(50, " ")
    i = GetPrivateProfileString("Config", "srv", "", Est, Len(Est), "./config.ini")
    srv = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    i = GetPrivateProfileString("Config", "db", "", Est, Len(Est), "./config.ini")
    db = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    i = GetPrivateProfileString("Config", "us", "", Est, Len(Est), "./config.ini")
    us = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    i = GetPrivateProfileString("Config", "pw", "", Est, Len(Est), "./config.ini")
    pw = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    i = GetPrivateProfileString("Config", "portscan", "", Est, Len(Est), "./config.ini")
    portscan = Mid(Est, 1, Len(Trim(Est)) - 1)
    Est = String$(50, " ")
    i = GetPrivateProfileString("Config", "portfiscal", "", Est, Len(Est), "./config.ini")
    portfiscal = Mid(Est, 1, Len(Trim(Est)) - 1)

    
    'para escribir ini
    'Dim I As Integer
    'Dim Est As String
    'Est = "Ejemplo - Apartado"
    'I = WritePrivateProfileString("Ejemplo", "Nombre", Est, "Ejemplo.ini")
    
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    
    'esta cadena es para conectar a sqlserver2000
    'cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Pwd=soloitenet;Initial Catalog=Ejemplo;Data Source=JDL\EXPRESS"
    'esta cadena es para sqlserver2008Express
    cn.ConnectionString = "Provider=SQLNCLI10;Data Source=" & srv & ";Persist Security info=True;Initial Catalog=" & db & ";User ID=" & us & ";Password=" & pw
    
   
    frmLoguin.Show
    Exit Sub
    
noInicia:
    MsgBox ("Error de configuración" & Chr(13) & "No se puede iniciar la aplicación")
    Exit Sub
End Sub

Public Function EnLetras(numero As String) As String
    Dim b, paso As Integer
    Dim expresion, entero, deci, flag As String
       
    flag = "N"
    For paso = 1 To Len(numero)
        If Mid(numero, paso, 1) = "." Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero
            Else
                deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero
            End If
        End If
    Next paso
   
    If Len(deci) = 1 Then
        deci = deci & "0"
    End If
   
    flag = "N"
    If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
        For paso = Len(entero) To 1 Step -1
            b = Len(entero) - (paso - 1)
            Select Case paso
            Case 3, 6, 9
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then
                            expresion = expresion & "cien "
                        Else
                            expresion = expresion & "ciento "
                        End If
                    Case "2"
                        expresion = expresion & "doscientos "
                    Case "3"
                        expresion = expresion & "trescientos "
                    Case "4"
                        expresion = expresion & "cuatrocientos "
                    Case "5"
                        expresion = expresion & "quinientos "
                    Case "6"
                        expresion = expresion & "seiscientos "
                    Case "7"
                        expresion = expresion & "setecientos "
                    Case "8"
                        expresion = expresion & "ochocientos "
                    Case "9"
                        expresion = expresion & "novecientos "
                End Select
               
            Case 2, 5, 8
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" Then
                            flag = "S"
                            expresion = expresion & "diez "
                        End If
                        If Mid(entero, b + 1, 1) = "1" Then
                            flag = "S"
                            expresion = expresion & "once "
                        End If
                        If Mid(entero, b + 1, 1) = "2" Then
                            flag = "S"
                            expresion = expresion & "doce "
                        End If
                        If Mid(entero, b + 1, 1) = "3" Then
                            flag = "S"
                            expresion = expresion & "trece "
                        End If
                        If Mid(entero, b + 1, 1) = "4" Then
                            flag = "S"
                            expresion = expresion & "catorce "
                        End If
                        If Mid(entero, b + 1, 1) = "5" Then
                            flag = "S"
                            expresion = expresion & "quince "
                        End If
                        If Mid(entero, b + 1, 1) > "5" Then
                            flag = "N"
                            expresion = expresion & "dieci"
                        End If
               
                    Case "2"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "veinte "
                            flag = "S"
                        Else
                            expresion = expresion & "veinti"
                            flag = "N"
                        End If
                   
                    Case "3"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "treinta "
                            flag = "S"
                        Else
                            expresion = expresion & "treinta y "
                            flag = "N"
                        End If
               
                    Case "4"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cuarenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cuarenta y "
                            flag = "N"
                        End If
               
                    Case "5"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cincuenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cincuenta y "
                            flag = "N"
                        End If
               
                    Case "6"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "sesenta "
                            flag = "S"
                        Else
                            expresion = expresion & "sesenta y "
                            flag = "N"
                        End If
               
                    Case "7"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "setenta "
                            flag = "S"
                        Else
                            expresion = expresion & "setenta y "
                            flag = "N"
                        End If
               
                    Case "8"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "ochenta "
                            flag = "S"
                        Else
                            expresion = expresion & "ochenta y "
                            flag = "N"
                        End If
               
                    Case "9"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "noventa "
                            flag = "S"
                        Else
                            expresion = expresion & "noventa y "
                            flag = "N"
                        End If
                End Select
               
            Case 1, 4, 7
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "uno "
                            Else
                                expresion = expresion & "un "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            expresion = expresion & "dos "
                        End If
                    Case "3"
                        If flag = "N" Then
                            expresion = expresion & "tres "
                        End If
                    Case "4"
                        If flag = "N" Then
                            expresion = expresion & "cuatro "
                        End If
                    Case "5"
                        If flag = "N" Then
                            expresion = expresion & "cinco "
                        End If
                    Case "6"
                        If flag = "N" Then
                            expresion = expresion & "seis "
                        End If
                    Case "7"
                        If flag = "N" Then
                            expresion = expresion & "siete "
                        End If
                    Case "8"
                        If flag = "N" Then
                            expresion = expresion & "ocho "
                        End If
                    Case "9"
                        If flag = "N" Then
                            expresion = expresion & "nueve "
                        End If
                End Select
            End Select
            If paso = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 6) Then
                    expresion = expresion & "mil "
                End If
            End If
            If paso = 7 Then
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "millón "
                Else
                    expresion = expresion & "millones "
                End If
            End If
        Next paso
       
        If deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion & "con " & deci ' & "/100"
            Else
                EnLetras = expresion & "con " & deci ' & "/100"
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion
            Else
                EnLetras = expresion
            End If
        End If
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        EnLetras = ""
    End If
End Function

Public Function EncodeBase64(ByVal strData As String) As Byte()
    Dim objStream As Object
    Dim objNode As Object
    Dim objXML As Object
    Dim bArray() As Byte
 
    Set objStream = CreateObject("ADODB.Stream")
 
    With objStream
        .Type = 2
        .Open
        .Charset = "unicode"
        .WriteText strData
        .Flush
        .Position = 0
        .Type = 1
        .Read (2)
        bArray = .Read
        .Close
    End With
 
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
 
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = bArray
    EnecodeBase64 = objNode.Text
    codigo_QR64 = EnecodeBase64
    
    Set objStream = Nothing
    Set objNode = Nothing
    Set objXML = Nothing
 
End Function

