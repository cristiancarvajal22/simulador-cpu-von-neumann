Attribute VB_Name = "ModAnimacion"
Option Explicit

' =====================================================================
'  ModAnimacion - todo lo que se ve moverse en el diagrama.
'
'  Cableado = grafo de aristas (formas E_*). Una RUTA es una secuencia
'  de aristas (hoja oculta _Rutas). Encender una ruta:
'    - pinta sus aristas de naranja SIN cambiar el grosor,
'    - las trae al frente (quedan por encima de las grises),
'    - deja una sola punta de flecha, en el destino.
'  Al apagarla, cada arista recupera su color y sus puntas de reposo
'  (hoja oculta _Aristas).
'
'  Velocidad: el dato viaja a velocidad CONSTANTE (puntos/segundo), asi
'  un bus largo y uno corto se recorren al mismo ritmo. El control
'  deslizante (1..10) escala esa velocidad; 10 es el maximo.
' =====================================================================

Private Const COL_ON As Long = 3243501       ' ED7D31 naranja
Private Const COL_FLASH As Long = 49407      ' FFC000 ambar
Private Const COL_CHIP As Long = 6240798     ' 1E3A5F azul marino (valor actual)
Private Const VEL_MAX As Double = 900        ' puntos por segundo con el control al maximo
Private Const T_MIN_TRAMO As Double = 0.12   ' duracion minima de un recorrido (s) al maximo

' =====================================================================
'  VELOCIDAD
' =====================================================================
' Factor 0.15 (nivel 1, muy lento) .. 1.0 (nivel 10, maximo)
Public Function Factor() As Double
    Dim v As Long
    v = CpuNum(R_VELOC)
    If v < 1 Then v = 1
    If v > 10 Then v = 10
    Factor = 0.15 + 0.85 * (v - 1) / 9
End Function

Public Function Animando() As Boolean
    Dim a As Variant
    a = CpuGet(R_ANIMAR)
    If VarType(a) = vbBoolean Then Animando = a Else Animando = (CStr(a) <> "False" And CStr(a) <> "0" And CStr(a) <> "")
End Function

' Pausa escalada por la velocidad
Public Sub Pausa(ByVal segundosAlMaximo As Double)
    Espera segundosAlMaximo / Factor()
End Sub

' =====================================================================
'  RUTAS
' =====================================================================
Private Function FilaRuta(ByVal id As String) As Long
    Dim ws As Worksheet, f As Variant
    Set ws = Hoja(HOJA_RUTAS)
    f = Application.Match(id, ws.Range("A:A"), 0)
    If IsError(f) Then FilaRuta = 0 Else FilaRuta = CLng(f)
End Function

Public Sub EncenderRuta(ByVal id As String)
    Dim ws As Worksheet, f As Long, partes As Variant, k As Long
    Dim nom As String, dirn As Long, s As Shape, p As Long

    Set ws = Hoja(HOJA_RUTAS)
    f = FilaRuta(id)
    If f = 0 Then Traza "Ruta inexistente: " & id: Exit Sub
    partes = Split(CStr(ws.Cells(f, 2).Value), ";")
    For k = LBound(partes) To UBound(partes)
        p = InStr(partes(k), ":")
        nom = Left$(partes(k), p - 1)
        dirn = CLng(Mid$(partes(k), p + 1))
        Set s = Forma(nom)
        If Not s Is Nothing Then
            s.Line.ForeColor.RGB = COL_ON
            s.Line.DashStyle = msoLineSolid
            s.Line.BeginArrowheadStyle = msoArrowheadNone
            s.Line.EndArrowheadStyle = msoArrowheadNone
            If k = UBound(partes) Then
                If dirn > 0 Then s.Line.EndArrowheadStyle = msoArrowheadTriangle _
                            Else s.Line.BeginArrowheadStyle = msoArrowheadTriangle
            End If
            s.ZOrder msoBringToFront
        End If
    Next k
End Sub

Public Sub ApagarRutas()
    Dim ws As Worksheet, r As Long, ultima As Long, s As Shape
    Set ws = Hoja(HOJA_ARISTAS)
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To ultima
        Set s = Forma(CStr(ws.Cells(r, 1).Value))
        If Not s Is Nothing Then
            If s.Line.ForeColor.RGB = COL_ON Then
                s.Line.ForeColor.RGB = CLng(ws.Cells(r, 4).Value)
                s.Line.BeginArrowheadStyle = IIf(ws.Cells(r, 2).Value = 1, msoArrowheadTriangle, msoArrowheadNone)
                s.Line.EndArrowheadStyle = IIf(ws.Cells(r, 3).Value = 1, msoArrowheadTriangle, msoArrowheadNone)
                If ws.Cells(r, 5).Value = 1 Then s.Line.DashStyle = msoLineDash
            End If
        End If
    Next r
End Sub

' Enciende la ruta y, si la animacion esta activa, hace viajar el dato
Public Sub Transferir(ByVal id As String, ByVal texto As String)
    EncenderRuta id
    If Animando() Then
        Pausa 0.05
        Viajar id, texto
    End If
End Sub

' Mueve DATO_VIAJERO por la polilinea de la ruta a velocidad constante
Public Sub Viajar(ByVal id As String, ByVal texto As String)
    Dim ws As Worksheet, f As Long, n As Long, k As Long
    Dim xs() As Double, ys() As Double, seg() As Double, total As Double
    Dim v As Shape, dur As Double, t0 As Double, u As Double, d As Double, acum As Double, q As Double

    ' sin valor no hay nada que mostrar: se ilumina la ruta y ya
    If Len(texto) = 0 Then Exit Sub

    Set ws = Hoja(HOJA_RUTAS)
    f = FilaRuta(id)
    If f = 0 Then Exit Sub
    n = CLng(ws.Cells(f, 3).Value)
    If n < 2 Then Exit Sub
    ReDim xs(0 To n - 1): ReDim ys(0 To n - 1): ReDim seg(0 To n - 2)
    For k = 0 To n - 1
        xs(k) = ws.Cells(f, 4 + 2 * k).Value
        ys(k) = ws.Cells(f, 5 + 2 * k).Value
    Next k
    For k = 0 To n - 2
        seg(k) = Sqr((xs(k + 1) - xs(k)) ^ 2 + (ys(k + 1) - ys(k)) ^ 2)
        total = total + seg(k)
    Next k
    If total <= 0 Then Exit Sub

    Set v = Forma("DATO_VIAJERO")
    If v Is Nothing Then Exit Sub
    v.TextFrame2.TextRange.Text = texto
    v.Visible = msoTrue
    v.ZOrder msoBringToFront

    dur = total / VEL_MAX
    If dur < T_MIN_TRAMO Then dur = T_MIN_TRAMO
    dur = dur / Factor()

    t0 = Timer
    Do
        u = (Timer - t0) / dur
        If u > 1 Or u < 0 Then u = 1
        d = u * total: acum = 0
        For k = 0 To n - 2
            If acum + seg(k) >= d Or k = n - 2 Then
                If seg(k) > 0 Then q = (d - acum) / seg(k) Else q = 0
                If q > 1 Then q = 1
                v.Left = xs(k) + (xs(k + 1) - xs(k)) * q - v.Width / 2
                v.Top = ys(k) + (ys(k + 1) - ys(k)) * q - v.Height / 2
                Exit For
            End If
            acum = acum + seg(k)
        Next k
        DoEvents
    Loop While u < 1
    v.Visible = msoFalse
End Sub

' =====================================================================
'  CHIPS Y DESTELLOS
' =====================================================================
Public Sub Destello(ByVal nombreChip As String)
    Dim s As Shape
    If Not Animando() Then Exit Sub
    Set s = Forma(nombreChip)
    If s Is Nothing Then Exit Sub
    s.Fill.ForeColor.RGB = COL_FLASH
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Pausa 0.12
    s.Fill.ForeColor.RGB = COL_CHIP
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
End Sub

' Nombre del chip que muestra un registro
Public Function ChipDe(ByVal reg As String) As String
    Select Case UCase$(reg)
        Case "IROP", "IRARG": ChipDe = "VAL_IR"
        Case "ALU": ChipDe = "VAL_COP"
        Case Else: ChipDe = "VAL_" & UCase$(reg)
    End Select
End Function

' =====================================================================
'  PANELES
' =====================================================================
Public Sub FaseMostrar(ByVal fase As String)
    Dim nombres As Variant, colores As Variant, k As Long, s As Shape
    nombres = Array("FETCH", "DECODE", "EXECUTE", "STORE")
    colores = Array(RGB(91, 155, 213), RGB(255, 192, 0), RGB(142, 124, 195), RGB(112, 173, 71))   ' colores de la diapositiva "Ciclos maquina"
    For k = 0 To 3
        Set s = Forma("FASE_" & nombres(k))
        If Not s Is Nothing Then
            If nombres(k) = fase Then
                s.Fill.ForeColor.RGB = colores(k)
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Else
                s.Fill.ForeColor.RGB = RGB(231, 230, 230)
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(127, 127, 127)
            End If
        End If
    Next k
    FijarTexto "SEC_TXT", fase
End Sub

' Rellena la lista del microprograma con el contenido de _Micro
Public Sub MicroListaMostrar()
    Dim ws As Worksheet, i As Long, n As Long, s As Shape
    Set ws = Hoja(HOJA_MICRO)
    n = CpuNum(R_NMICRO)
    For i = 1 To 13
        Set s = Forma("MICROP_" & Format$(i, "00"))
        If Not s Is Nothing Then
            If i <= n Then
                s.Visible = msoTrue
                s.TextFrame2.TextRange.Text = ws.Cells(i + 1, 2).Value & "   " & ws.Cells(i + 1, 3).Value
            Else
                s.TextFrame2.TextRange.Text = ""
                s.Visible = msoFalse
            End If
        End If
    Next i
End Sub

Public Sub MicroListaResaltar(ByVal actual As Long)
    Dim i As Long, s As Shape
    For i = 1 To 13
        Set s = Forma("MICROP_" & Format$(i, "00"))
        If Not s Is Nothing Then
            If i = actual Then
                s.Fill.ForeColor.RGB = COL_ON
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            ElseIf i < actual Then
                s.Fill.ForeColor.RGB = RGB(242, 242, 242)
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(128, 128, 128)
            Else
                s.Fill.ForeColor.RGB = RGB(255, 255, 255)
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End If
        End If
    Next i
End Sub

Public Sub MicroTitulo(ByVal t As String)
    FijarTexto "LBL_MICROPROG", t
End Sub

Public Sub DecodMostrar(ByVal op As Long)
    Dim i As TInstr
    If IsaExiste(op) Then
        i = IsaInfo(op)
        FijarTexto "DECOD_TXT", "Opcode " & HexH(op) & " -> " & IsaFormato(op) & " | " & _
                   i.Bytes & IIf(i.Bytes = 1, " byte", " bytes") & " | modo " & IsaModo(op) & _
                   IIf(Len(i.Flags) > 0, " | flags: " & Replace(Replace(Replace(Replace(i.Flags, "Z", "ZF "), "C", "CF "), "S", "SF "), "O", "OF "), "")
        FijarTexto "VAL_DEC", IsaFormato(op)
        ChipAlu IIf(i.Grupo = "ALU" Or i.Grupo = "CMP" Or i.Grupo = "UNARIA", IsaSimbolo(op), "")
    Else
        FijarTexto "DECOD_TXT", "Opcode " & HexH(op) & t(" no pertenece a la ISA: instrucci~on inv~alida")
        FijarTexto "VAL_DEC", "??"
    End Sub
End Sub

Public Sub DecodLimpiar()
    FijarTexto "DECOD_TXT", t("Esperando una instrucci~on en el IR.")
    FijarTexto "VAL_DEC", "-"
    ChipAlu ""
End Sub

' Chip de la ALU: muestra la operacion; gris cuando la instruccion no la usa
Public Sub ChipAlu(ByVal simbolo As String)
    Dim s As Shape
    Set s = Forma("VAL_COP")
    If s Is Nothing Then Exit Sub
    s.TextFrame2.TextRange.Text = simbolo
    If Len(simbolo) = 0 Then
        s.Fill.ForeColor.RGB = RGB(217, 217, 217)
        s.Line.ForeColor.RGB = RGB(166, 166, 166)
    Else
        s.Fill.ForeColor.RGB = COL_CHIP
        s.Line.ForeColor.RGB = RGB(18, 38, 63)
    End If
End Sub

Public Sub Narrar(ByVal texto As String)
    FijarTexto "NARRA_TXT", texto
End Sub

Public Sub RelojMostrar()
    If CpuNum(R_HALT) = 1 Then
        FijarTexto "VAL_RELOJ", "HLT"
    Else
        FijarTexto "VAL_RELOJ", "T " & CpuNum(R_CICLOS)
    End If
    FijarTexto "LBL_CONTADOR", t("Instrucciones: ") & CpuNum(R_INSTR) & "    Ciclos T: " & CpuNum(R_CICLOS)
End Sub

' El resaltado de la linea en ejecucion vive ahora en ModEditor, junto al
' coloreado de sintaxis (tocar la negrita aqui lo destruiria).
