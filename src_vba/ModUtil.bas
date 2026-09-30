Attribute VB_Name = "ModUtil"
Option Explicit

' =====================================================================
'  ModUtil - utilidades compartidas: formatos numericos, tiempos,
'  acceso a hojas y textos.
' =====================================================================

Public Const HOJA_SIM As String = "Simulador"
Public Const HOJA_LOG As String = "Log"
Public Const HOJA_CPU As String = "_CPU"
Public Const HOJA_RAM As String = "_RAM"
Public Const HOJA_MICRO As String = "_Micro"
Public Const HOJA_RUTAS As String = "_Rutas"
Public Const HOJA_ARISTAS As String = "_Aristas"
Public Const HOJA_EJEMPLOS As String = "_Ejemplos"

Public Const DEPURAR As Boolean = False

' =====================================================================
'  TEXTOS
' =====================================================================
Public Function t(ByVal s As String) As String
    s = Replace(s, "~a", ChrW(225))
    s = Replace(s, "~e", ChrW(233))
    s = Replace(s, "~i", ChrW(237))
    s = Replace(s, "~o", ChrW(243))
    s = Replace(s, "~u", ChrW(250))
    s = Replace(s, "~n", ChrW(241))
    s = Replace(s, "~N", ChrW(209))
    s = Replace(s, "~?", ChrW(191))
    s = Replace(s, "~!", ChrW(161))
    s = Replace(s, "<-", ChrW(8592))
    s = Replace(s, "->", ChrW(8594))
    t = s
End Function

' =====================================================================
'  FORMATOS
' =====================================================================
Public Function Hex2(ByVal v As Long) As String
    Hex2 = Right$("0" & Hex$(v And &HFF), 2)
End Function

Public Function HexH(ByVal v As Long) As String
    HexH = Hex2(v) & "h"
End Function

Public Function Hex0x(ByVal v As Long) As String
    Hex0x = "0x" & Hex2(v)
End Function

Public Function Bin8(ByVal v As Long) As String
    Dim i As Long, s As String
    v = v And &HFF
    For i = 7 To 0 Step -1
        If (v And (2 ^ i)) <> 0 Then s = s & "1" Else s = s & "0"
    Next i
    Bin8 = s
End Function

Public Function ConSigno(ByVal v As Long) As Long
    v = v And &HFF
    If v >= 128 Then ConSigno = v - 256 Else ConSigno = v
End Function

Public Function LeerNumero(ByVal s As String, ByRef v As Long) As Boolean
    Dim i As Long, c As String, base As Long, cuerpo As String, acum As Double
    s = UCase$(Trim$(s))
    LeerNumero = False
    If Len(s) = 0 Then Exit Function

    If Left$(s, 2) = "0X" Then
        base = 16: cuerpo = Mid$(s, 3)
    ElseIf Right$(s, 1) = "H" Then
        base = 16: cuerpo = Left$(s, Len(s) - 1)
    ElseIf Right$(s, 1) = "B" And Len(s) > 1 And Replace(Replace(Left$(s, Len(s) - 1), "0", ""), "1", "") = "" Then
        base = 2: cuerpo = Left$(s, Len(s) - 1)
    Else
        base = 10: cuerpo = s
    End If
    If Len(cuerpo) = 0 Then Exit Function

    acum = 0
    For i = 1 To Len(cuerpo)
        c = Mid$(cuerpo, i, 1)
        Select Case base
            Case 2
                If c <> "0" And c <> "1" Then Exit Function
                acum = acum * 2 + Val(c)
            Case 10
                If c < "0" Or c > "9" Then Exit Function
                acum = acum * 10 + Val(c)
            Case 16
                If InStr("0123456789ABCDEF", c) = 0 Then Exit Function
                acum = acum * 16 + (InStr("0123456789ABCDEF", c) - 1)
        End Select
        If acum > 65535 Then Exit Function
    Next i
    v = CLng(acum)
    LeerNumero = True
End Function

' =====================================================================
'  HOJAS
' =====================================================================
Public Function Hoja(ByVal nombre As String) As Worksheet
    Set Hoja = ThisWorkbook.Worksheets(nombre)
End Function

Public Function Sim() As Worksheet
    Set Sim = ThisWorkbook.Worksheets(HOJA_SIM)
End Function

Public Function Forma(ByVal nombre As String) As Shape
    Dim s As Shape
    On Error Resume Next
    Set s = Sim().Shapes(nombre)
    On Error GoTo 0
    Set Forma = s
End Function

Public Sub FijarTexto(ByVal nombre As String, ByVal t As String)
    Dim s As Shape
    Set s = Forma(nombre)
    If s Is Nothing Then Exit Sub
    s.TextFrame2.TextRange.Text = t
End Sub

Public Sub EscribirTexto(ByVal c As Range, ByVal texto As String)
    c.Value = "'" & texto
    On Error Resume Next
    c.Errors.Item(xlNumberAsText).Ignore = True
    On Error GoTo 0
End Sub

Public Sub Estado(ByVal msg As String, Optional ByVal esError As Boolean = False)
    With Sim().Range("ESTADO")
        .Value = msg
        If esError Then .Font.Color = RGB(192, 0, 0) Else .Font.Color = RGB(31, 56, 100)
    End With
End Sub

' =====================================================================
'  TIEMPO
' =====================================================================
Public Sub Espera(ByVal segundos As Double)
    Dim t0 As Double
    If segundos <= 0 Then DoEvents: Exit Sub
    t0 = Timer
    Do
        DoEvents
        If Timer < t0 Then Exit Do
    Loop While (Timer - t0) < segundos
End Sub

Public Sub Traza(ByVal m As String)
    If Not DEPURAR Then Exit Sub
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open ThisWorkbook.Path & "\traza_vba.log" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & "  " & m
    Close #f
    On Error GoTo 0
End Sub
