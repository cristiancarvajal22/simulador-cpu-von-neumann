Attribute VB_Name = "ModMemoria"
Option Explicit

' =====================================================================
'  ModMemoria - Memoria principal: 256 posiciones de 8 bits (00h-FFh).
'
'  Almacen: hoja oculta _RAM, fila = direccion + 1
'     A  valor (0..255)
'     B  fila del editor de la instruccion/dato que ocupa el byte (0 = libre)
'     C  tipo de byte: I = inicio de instruccion | O = operando | D = dato
'
'  Primitivas del enunciado: Read(address) y Write(address, value).
'  (En VBA "Write" es una palabra reservada, por eso se llaman
'  MemRead y MemWrite.)
'
'  Segmentacion logica (solo visual):
'     00h-7Fh  Segmento de CODIGO     80h-FFh  Segmento de DATOS
' =====================================================================

Public Const SEG_DATOS As Long = &H80

' colores de la matriz
Private Const C_COD As Long = 16312540       ' DCE8F8 azul: segmento de codigo
Private Const C_DAT As Long = 15135453       ' DDF2E6 verde menta: segmento de datos
Private Const C_LEER As Long = 10086143      ' 99E6FF -> amarillo claro (BGR de FFE699)
Private Const C_ESCR As Long = 11389944      ' ADCBF8 -> salmon (BGR de F8CBAD)
Private Const C_BASE As Long = 16447220      ' F4F6FA gris azulado: sin acceso todavia

' =====================================================================
'  PRIMITIVAS
' =====================================================================
Public Function MemRead(ByVal direccion As Long) As Long
    direccion = direccion And &HFF
    MemRead = CLng(Val(CStr(ThisWorkbook.Worksheets(HOJA_RAM).Cells(direccion + 1, 1).Value))) And &HFF
End Function

Public Sub MemWrite(ByVal direccion As Long, ByVal valor As Long)
    direccion = direccion And &HFF
    ThisWorkbook.Worksheets(HOJA_RAM).Cells(direccion + 1, 1).Value = valor And &HFF
    MemPintar direccion
    If Abs(direccion - CpuNum(R_MAR)) <= 4 Then MemVentana CpuNum(R_MAR)
End Sub

' Metadatos del byte (los fija el ensamblador)
Public Sub MemMarcarOrigen(ByVal direccion As Long, ByVal filaEditor As Long, ByVal tipo As String)
    With ThisWorkbook.Worksheets(HOJA_RAM)
        .Cells(direccion + 1, 2).Value = filaEditor
        .Cells(direccion + 1, 3).Value = tipo
    End With
End Sub

Public Function MemFilaEditor(ByVal direccion As Long) As Long
    MemFilaEditor = CLng(Val(CStr(ThisWorkbook.Worksheets(HOJA_RAM).Cells((direccion And &HFF) + 1, 2).Value)))
End Function

Public Function MemTipo(ByVal direccion As Long) As String
    MemTipo = CStr(ThisWorkbook.Worksheets(HOJA_RAM).Cells((direccion And &HFF) + 1, 3).Value)
End Function

' Borra toda la memoria (valores y metadatos)
Public Sub MemReset()
    With ThisWorkbook.Worksheets(HOJA_RAM)
        .Range("A1:A256").Value = 0
        .Range("B1:C256").ClearContents
    End With
    CpuSet R_ULTMEM, -1
    CpuSet R_ULTPC, -1
    MemPintarTodo
End Sub

' =====================================================================
'  MATRIZ 16 x 16 EN LA HOJA
' =====================================================================
Public Function MemCelda(ByVal direccion As Long) As Range
    direccion = direccion And &HFF
    Set MemCelda = Sim().Range("MEM_GRID").Cells((direccion \ 16) + 1, (direccion Mod 16) + 1)
End Function

Public Sub MemPintar(ByVal direccion As Long)
    Dim c As Range, v As Long
    Application.EnableEvents = False
    Set c = MemCelda(direccion)
    v = MemRead(direccion)
    EscribirTexto c, Hex2(v)

    ' fondo: segmento, o resaltado si es la ultima celda accedida
    If direccion = CpuNum(R_ULTMEM) Then
        If CStr(CpuGet(R_ULTMEMT)) = "W" Then c.Interior.color = C_ESCR Else c.Interior.color = C_LEER
    ElseIf direccion < SEG_DATOS Then
        c.Interior.color = C_COD
    Else
        c.Interior.color = C_DAT
    End If

    ' texto: los ceros en gris, lo ocupado en negro
    If v = 0 And MemTipo(direccion) = "" Then
        c.Font.color = RGB(166, 166, 166): c.Font.Bold = False
    Else
        c.Font.color = RGB(0, 0, 0): c.Font.Bold = (MemTipo(direccion) = "I")
    End If

    ' marco azul grueso en la celda a la que apunta el PC
    If direccion = CpuNum(R_ULTPC) Then
        c.BorderAround xlContinuous, xlThick, , RGB(31, 78, 121)
    Else
        c.BorderAround xlContinuous, xlThin, , RGB(191, 191, 191)
    End If

    ' linea divisoria entre el segmento de codigo (00h-7Fh) y el de datos (80h-FFh)
    If direccion \ 16 = 7 Then
        With c.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: .color = RGB(31, 56, 100): End With
    ElseIf direccion \ 16 = 8 Then
        With c.Borders(xlEdgeTop): .LineStyle = xlContinuous: .Weight = xlMedium: .color = RGB(31, 56, 100): End With
    End If
    Application.EnableEvents = True
End Sub

Public Sub MemPintarTodo()
    Dim d As Long
    Application.ScreenUpdating = False
    For d = 0 To 255
        MemPintar d
    Next d
    Application.ScreenUpdating = True
End Sub

' Resalta la celda leida ("R") o escrita ("W") y quita el resalte anterior
Public Sub MemMarcarAcceso(ByVal direccion As Long, ByVal tipo As String)
    Dim anterior As Long
    anterior = CpuNum(R_ULTMEM)
    CpuSet R_ULTMEM, direccion And &HFF
    CpuSet R_ULTMEMT, tipo
    If anterior >= 0 And anterior <> direccion Then MemPintar anterior
    MemPintar direccion
End Sub

' Mueve el marco del PC en la matriz
Public Sub MemMarcarPC(ByVal direccion As Long)
    Dim anterior As Long
    anterior = CpuNum(R_ULTPC)
    CpuSet R_ULTPC, direccion And &HFF
    If anterior >= 0 And anterior <> direccion Then MemPintar anterior
    MemPintar direccion
End Sub

' =====================================================================
'  PANEL "PRIMITIVAS DE MEMORIA" DEL DIAGRAMA
'  Muestra la ultima llamada a Read(address) o Write(address, value):
'  amarillo = lectura, salmon = escritura (los colores de la matriz).
' =====================================================================
Public Sub PrimitivaMostrar(ByVal texto As String, ByVal tipo As String)
    Dim s As Shape
    Set s = Forma("PRIM_ULT")
    If s Is Nothing Then Exit Sub
    If Len(texto) = 0 Then texto = ChrW(8211)
    s.TextFrame2.TextRange.Text = texto
    Select Case tipo
        Case "R": s.Fill.ForeColor.RGB = C_LEER
        Case "W": s.Fill.ForeColor.RGB = C_ESCR
        Case Else: s.Fill.ForeColor.RGB = C_BASE
    End Select
End Sub

' =====================================================================
'  VENTANA DE MEMORIA DEL DIAGRAMA (8 filas; la 4a es la celda del MAR)
' =====================================================================
Public Sub MemVentana(ByVal centro As Long)
    Dim i As Long, d As Long
    For i = 1 To 8
        d = centro + i - 4
        If d < 0 Or d > 255 Then
            FijarTexto "MEMW_D" & i, ""
            FijarTexto "MEMW_H" & i, ""
            FijarTexto "MEMW_B" & i, ""
        Else
            FijarTexto "MEMW_D" & i, HexH(d)
            FijarTexto "MEMW_H" & i, Hex2(MemRead(d))
            FijarTexto "MEMW_B" & i, Bin8(MemRead(d))
        End If
    Next i
End Sub

' Resalta (o no) la fila del MAR en la ventana del diagrama
Public Sub MemVentanaResaltar(ByVal tipo As String)
    Dim n As Variant, s As Shape, color As Long
    Select Case tipo
        Case "R": color = RGB(255, 230, 153)
        Case "W": color = RGB(248, 203, 173)
        Case Else: color = RGB(255, 242, 204)
    End Select
    For Each n In Array("MEMW_D4", "MEMW_H4", "MEMW_B4")
        Set s = Forma(CStr(n))
        If Not s Is Nothing Then s.Fill.ForeColor.RGB = color
    Next n
End Sub

' =====================================================================
'  INSPECTOR (celda seleccionada de la matriz)
' =====================================================================
Public Sub MemInspeccionar(ByVal direccion As Long)
    Dim v As Long, ws As Worksheet, desc As String, tipo As String, op As Long
    Set ws = Sim()
    direccion = direccion And &HFF
    v = MemRead(direccion)
    tipo = MemTipo(direccion)

    Select Case tipo
        Case "I"
            desc = t("Instrucci~on: ") & IsaTexto(v, MemRead(direccion + 1))
        Case "O"
            op = MemRead(direccion - 1)
            desc = t("Operando de ") & IsaTexto(op, v)
        Case "D"
            desc = "Dato (DB)"
        Case Else
            If IsaExiste(v) And direccion < SEG_DATOS And v <> 0 Then
                desc = t("Como instrucci~on ser~ia: ") & IsaTexto(v, MemRead(direccion + 1))
            ElseIf v = 0 Then
                desc = t("Libre (00h = HLT si se ejecutara)")
            Else
                desc = "Dato"
            End If
    End Select

    Application.EnableEvents = False
    ws.Range("INSP_DIR").Value = "'" & HexH(direccion) & "  (" & direccion & ")  " & _
                                 IIf(direccion < SEG_DATOS, t("c~odigo"), "datos")
    ws.Range("INSP_HEX").Value = "'" & HexH(v)
    ws.Range("INSP_BIN").Value = "'" & Bin8(v) & "b"
    ws.Range("INSP_DEC").Value = "'" & v & t("   (con signo: ") & ConSigno(v) & ")"
    ws.Range("INSP_MNEM").Value = desc
    Application.EnableEvents = True
End Sub

' Lo que el usuario escribe en una celda de la matriz:
'   hexadecimal por defecto (0A, 3F, 0x3F, 3Fh) | #123 decimal | 00101010b binario
'   (el decimal va con prefijo # porque un sufijo "d" chocaria con el digito hex D)
Public Function MemLeerEntrada(ByVal s As String, ByRef v As Long) As Boolean
    Dim u As String
    u = UCase$(Trim$(s))
    If Len(u) = 0 Then v = 0: MemLeerEntrada = True: Exit Function
    If Left$(u, 1) = "#" Then
        MemLeerEntrada = LeerNumero(Mid$(u, 2), v)
    ElseIf Right$(u, 1) = "B" And Len(u) >= 3 Then
        MemLeerEntrada = LeerNumero(u, v)
    ElseIf Right$(u, 1) = "H" Or Left$(u, 2) = "0X" Then
        MemLeerEntrada = LeerNumero(u, v)
    Else
        MemLeerEntrada = LeerNumero(u & "H", v)
    End If
    If MemLeerEntrada And v > 255 Then MemLeerEntrada = False
End Function

' Direccion de una celda de la matriz, o -1 si la celda no es de la matriz
Public Function MemDireccionDeCelda(ByVal c As Range) As Long
    Dim g As Range, f As Long, k As Long
    Set g = Sim().Range("MEM_GRID")
    If Intersect(c, g) Is Nothing Then MemDireccionDeCelda = -1: Exit Function
    f = c.Row - g.Row
    k = c.Column - g.Column
    MemDireccionDeCelda = f * 16 + k
End Function
