Attribute VB_Name = "ModEditor"
Option Explicit

' =====================================================================
'  ModEditor - presentacion del editor de ensamblador.
'
'  Colorea cada linea como lo haria un editor de codigo, usando el
'  formato por caracteres de Excel (Range.Characters). El color no es
'  decorativo: dice que ha entendido el ensamblador de cada palabra.
'    comentario   gris claro en cursiva (no se ensambla)
'    etiqueta     morado en negrita
'    mnemonica    color segun el grupo de la instruccion
'    registro     magenta
'    [direccion]  verde azulado
'    numero       ambar
'    directiva    gris oscuro en negrita (ORG, DB)
' =====================================================================

Private Const COL_TXT As Long = 2105376     ' 202020
Private Const COL_COMENT As Long = 10066329    ' 999999
Private Const COL_ETIQ As Long = 10498160    ' 7030A0 morado
Private Const COL_TRANS As Long = 12611584    ' 0070C0 azul
Private Const COL_MEMORIA As Long = 11825152    ' 0070B4 azul memoria
Private Const COL_ARIT As Long = 4227072     ' 008040 verde
Private Const COL_LOGICA As Long = 9868800     ' 009696 verde azulado
Private Const COL_SALTO As Long = 1137349     ' C55A11 naranja
Private Const COL_CTRL As Long = 192         ' C00000 rojo
Private Const COL_DIREC As Long = 5921370     ' 5A5A5A gris
Private Const COL_REG As Long = 5243035     ' 9B0050 magenta
Private Const COL_NUM As Long = 28351       ' BF6E00 ambar

Private Const N_FILAS As Long = 40

' =====================================================================
'  COLOREADO
' =====================================================================
Public Sub EditorColorearTodo()
    Dim i As Long
    Application.ScreenUpdating = False
    For i = 1 To N_FILAS
        EditorColorear i
    Next i
    Application.ScreenUpdating = True
End Sub

Public Sub EditorColorear(ByVal fila As Long)
    Dim c As Range, s As String, codigo As String
    Dim pComent As Long, pEtq As Long, ini As Long

    If fila < 1 Or fila > N_FILAS Then Exit Sub
    Set c = Sim().Range("EDITOR_COD").Cells(fila, 1)
    s = CStr(c.Value)

    With c.Font
        .color = COL_TXT
        .Italic = False
        .Bold = False
    End With
    If Len(Trim$(s)) = 0 Then Exit Sub

    ' --- comentario: desde el ; hasta el final ---
    pComent = InStr(s, ";")
    If pComent > 0 Then
        With c.Characters(pComent, Len(s) - pComent + 1).Font
            .color = COL_COMENT
            .Italic = True
        End With
        codigo = Left$(s, pComent - 1)
    Else
        codigo = s
    End If
    If Len(Trim$(codigo)) = 0 Then Exit Sub

    ' --- etiqueta: lo que va antes de los dos puntos ---
    pEtq = InStr(codigo, ":")
    If pEtq > 0 Then
        With c.Characters(1, pEtq).Font
            .color = COL_ETIQ
            .Bold = True
        End With
        ini = pEtq + 1
    Else
        ini = 1
    End If

    ColorearTokens c, codigo, ini
End Sub

' Recorre las palabras del codigo y pinta cada una segun lo que sea
Private Sub ColorearTokens(ByVal c As Range, ByVal codigo As String, ByVal desde As Long)
    Dim i As Long, ini As Long, tok As String, primera As Boolean, col As Long

    primera = True
    i = desde
    Do While i <= Len(codigo)
        If EsSeparador(Mid$(codigo, i, 1)) Then
            i = i + 1
        Else
            ini = i
            Do While i <= Len(codigo)
                If EsSeparador(Mid$(codigo, i, 1)) Then Exit Do
                i = i + 1
            Loop
            tok = Mid$(codigo, ini, i - ini)
            col = ColorDeToken(tok, primera)
            If col >= 0 Then
                c.Characters(ini, Len(tok)).Font.color = col
                If primera Then c.Characters(ini, Len(tok)).Font.Bold = True
            End If
            primera = False
        End If
    Loop
End Sub

Private Function EsSeparador(ByVal ch As String) As Boolean
    EsSeparador = (ch = " " Or ch = vbTab Or ch = ",")
End Function

' -1 = dejar el color base
Private Function ColorDeToken(ByVal tok As String, ByVal esPrimera As Boolean) As Long
    Dim u As String, v As Long
    u = UCase$(Trim$(tok))
    If Len(u) = 0 Then ColorDeToken = -1: Exit Function

    If esPrimera Then
        If u = "ORG" Or u = "DB" Then ColorDeToken = COL_DIREC: Exit Function
        ColorDeToken = ColorDeMnemonica(u)
        Exit Function
    End If

    If u = "AX" Or u = "BX" Then ColorDeToken = COL_REG: Exit Function
    If Left$(u, 1) = "[" Then ColorDeToken = COL_MEMORIA: Exit Function
    If LeerNumero(Replace(u, "-", ""), v) Then ColorDeToken = COL_NUM: Exit Function
    ColorDeToken = COL_ETIQ           ' referencia a una etiqueta (destino de salto)
End Function

Private Function ColorDeMnemonica(ByVal mn As String) As Long
    ' alias de x86 que el ensamblador acepta
    If mn = "JE" Or mn = "JNE" Then ColorDeMnemonica = COL_SALTO: Exit Function
    If Not IsaEsMnemonica(mn) Then ColorDeMnemonica = COL_CTRL: Exit Function   ' desconocida: en rojo

    Select Case IsaGrupoDeMnemonica(mn)
        Case "TRANS": ColorDeMnemonica = COL_TRANS
        Case "MEMR", "MEMW": ColorDeMnemonica = COL_MEMORIA
        Case "ALU", "CMP", "UNARIA"
            If mn = "AND" Or mn = "OR" Or mn = "XOR" Or mn = "NOT" Then
                ColorDeMnemonica = COL_LOGICA
            Else
                ColorDeMnemonica = COL_ARIT
            End If
        Case "SALTO": ColorDeMnemonica = COL_SALTO
        Case Else: ColorDeMnemonica = COL_CTRL
    End Select
End Function

' =====================================================================
'  RESALTADO DE LA LINEA EN EJECUCION
'  Solo cambia el fondo: si tocara la negrita se perderia el coloreado.
' =====================================================================
Public Sub EditorResaltar(ByVal fila As Long)
    Dim ws As Worksheet, anterior As Long
    Set ws = Sim()
    anterior = CpuNum(R_ULTLINEA)
    Application.EnableEvents = False
    If anterior > 0 Then
        Intersect(ws.Range("EDITOR_DIR").Cells(anterior, 1).EntireRow, ws.Range("EDITOR_TODO")).Interior.color = RGB(255, 255, 255)
        ws.Range("EDITOR_DIR").Cells(anterior, 1).Interior.color = RGB(247, 247, 247)
        ws.Range("EDITOR_HEX").Cells(anterior, 1).Interior.color = RGB(247, 247, 247)
    End If
    If fila > 0 Then
        Intersect(ws.Range("EDITOR_DIR").Cells(fila, 1).EntireRow, ws.Range("EDITOR_TODO")).Interior.color = RGB(255, 230, 153)
    End If
    Application.EnableEvents = True
    CpuSet R_ULTLINEA, fila
End Sub
