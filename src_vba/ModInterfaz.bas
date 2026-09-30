Attribute VB_Name = "ModInterfaz"
Option Explicit

' =====================================================================
'  ModInterfaz - controles de la hoja Simulador.
'
'    LOAD        ensambla el programa del editor y lo carga en memoria
'    STEP        avanza UNA micro-operacion (un pulso de reloj)
'    STEP INSTR  avanza hasta terminar la instruccion en curso
'    RUN         ejecucion continua; la velocidad la marca el deslizador
'    PAUSE       detiene RUN al terminar la micro-operacion en curso
'    RESET       registros y PC a cero (la memoria se conserva)
'
'  Todos los puntos de entrada atrapan sus errores: un error sin manejar
'  abriria el depurador y dejaria la hoja bloqueada.
' =====================================================================

Private Const MAX_MICRO_RUN As Long = 20000

' =====================================================================
'  BOTONES
' =====================================================================
Public Sub Btn_Load()
    On Error GoTo fallo
    If Ocupado() Then Exit Sub
    Application.ScreenUpdating = False
    CpuReset
    LogLimpiar
    Ensamblar
    VistaInicial
    Application.ScreenUpdating = True
    Exit Sub
fallo:
    ErrorInterno "LOAD"
End Sub

Public Sub Btn_Step()
    On Error GoTo fallo
    If Ocupado() Then Exit Sub
    MicroPaso
    Exit Sub
fallo:
    ErrorInterno "STEP"
End Sub

Public Sub Btn_StepInstr()
    On Error GoTo fallo
    If Ocupado() Then Exit Sub
    CpuSet R_CORRIENDO, 1
    CpuSet R_PAUSA, 0
    Do
        If Not MicroPaso() Then Exit Do
        If EnFrontera() Then Exit Do
        If CpuNum(R_PAUSA) = 1 Then Exit Do
        EsperaEntreMicros
    Loop
    CpuSet R_CORRIENDO, 0
    Exit Sub
fallo:
    CpuSet R_CORRIENDO, 0
    ErrorInterno "STEP INSTR"
End Sub

Public Sub Btn_Run()
    Dim n As Long
    On Error GoTo fallo
    If CpuNum(R_CORRIENDO) = 1 Then Exit Sub
    CpuSet R_CORRIENDO, 1
    CpuSet R_PAUSA, 0
    Estado t("Ejecutando... (PAUSE para detener)")
    Do
        If Not MicroPaso() Then Exit Do
        n = n + 1
        If CpuNum(R_HALT) = 1 Then Exit Do
        If n >= MAX_MICRO_RUN Then
            Estado t("RUN detenido tras ") & n & t(" ciclos: ~?el programa tiene un bucle infinito?"), True
            Exit Do
        End If
        EsperaEntreMicros
        If CpuNum(R_PAUSA) = 1 Then
            Estado t("En pausa. STEP para avanzar paso a paso, RUN para continuar.")
            Exit Do
        End If
    Loop
    CpuSet R_CORRIENDO, 0
    Exit Sub
fallo:
    CpuSet R_CORRIENDO, 0
    ErrorInterno "RUN"
End Sub

Public Sub Btn_Pause()
    CpuSet R_PAUSA, 1
End Sub

Public Sub Btn_Reset()
    On Error GoTo fallo
    If Ocupado() Then Exit Sub
    Application.ScreenUpdating = False
    CpuReset
    LogLimpiar
    VistaInicial
    Application.ScreenUpdating = True
    Estado t("RESET: registros y PC a cero. La memoria conserva su contenido (LOAD recarga el programa).")
    Exit Sub
fallo:
    ErrorInterno "RESET"
End Sub

Public Sub Btn_Ejemplo()
    On Error GoTo fallo
    If Ocupado() Then Exit Sub
    EjemploCargar CStr(Sim().Range("EJEMPLO_SEL").Value)
    Btn_Load
    Exit Sub
fallo:
    ErrorInterno "Cargar ejemplo"
End Sub

' Al mover el deslizador de velocidad
Public Sub Btn_Velocidad()
    FijarTexto "LBL_VELOC", "Velocidad " & CpuNum(R_VELOC) & "/10"
End Sub

' =====================================================================
'  AUXILIARES
' =====================================================================
Private Sub EsperaEntreMicros()
    If Animando() Then Pausa 0.25 Else Pausa 0.04
End Sub

' Si RUN esta en marcha, se pide la pausa y se ignora el boton pulsado
Private Function Ocupado() As Boolean
    If CpuNum(R_CORRIENDO) = 1 Then
        CpuSet R_PAUSA, 1
        Estado t("La CPU estaba en marcha: se ha pausado. Vuelve a pulsar el bot~on.")
        Ocupado = True
    End If
End Function

Private Sub ErrorInterno(ByVal donde As String)
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Estado "Error interno en " & donde & ": " & Err.Description, True
End Sub

' Deja el diagrama en el estado de "listo para empezar"
Public Sub VistaInicial()
    ApagarRutas
    FaseMostrar ""
    FijarTexto "SEC_TXT", "READY"
    Hoja(HOJA_MICRO).Range("A2:F41").ClearContents
    CpuSet R_NMICRO, 0
    CpuSet R_MICRO, 0
    MicroListaMostrar
    MicroTitulo t("Microprogram (pulsa STEP)")
    DecodLimpiar
    CpuSet R_ULTMEM, -1
    MemPintarTodo
    MemMarcarPC 0
    MemVentanaResaltar ""
    PrimitivaMostrar "", ""
    EditorResaltar MemFilaEditor(0)
    CpuPintarTodo
    RelojMostrar
    Btn_Velocidad
    MemInspeccionar 0
    Narrar t("Listo. STEP ejecuta una micro-operaci~on; STEP INSTR, una instrucci~on completa; RUN, el programa entero. ") & _
           t("La l~inea amarilla del editor es la instrucci~on que va a ejecutarse.")
End Sub

' Llamado por el generador al construir el libro
Public Function Ui_Inicializar() As String
    On Error GoTo fallo
    TeoriaEscribirISA
    EjemploCargar CStr(Sim().Range("EJEMPLO_SEL").Value)
    CpuReset
    LogLimpiar
    Ui_Inicializar = Ensamblar()
    VistaInicial
    Exit Function
fallo:
    Ui_Inicializar = "ERR " & Err.Number & ": " & Err.Description
End Function

' Tabla de la ISA en la hoja Teoria, generada desde ModISA (una sola fuente)
Public Sub TeoriaEscribirISA()
    Dim ws As Worksheet, f As Long, ops As Variant, k As Long, i As TInstr, Flags As String
    Set ws = ThisWorkbook.Worksheets("Teoria")
    f = ws.Range("ISA_INICIO").Row
    ws.Range(ws.Cells(f, 1), ws.Cells(f + 80, 6)).ClearContents
    ws.Cells(f, 1).Value = "Opcode"
    ws.Cells(f, 2).Value = t("Instrucci~on")
    ws.Cells(f, 3).Value = "Bytes"
    ws.Cells(f, 4).Value = t("Operaci~on (RTL)")
    ws.Cells(f, 5).Value = "Flags"
    ws.Cells(f, 6).Value = "Modo"
    ws.Range(ws.Cells(f, 1), ws.Cells(f, 6)).Font.Bold = True
    ws.Range(ws.Cells(f, 1), ws.Cells(f, 6)).Interior.color = RGB(31, 56, 100)
    ws.Range(ws.Cells(f, 1), ws.Cells(f, 6)).Font.color = RGB(255, 255, 255)
    ops = IsaListaOpcodes()
    For k = LBound(ops) To UBound(ops)
        i = IsaInfo(ops(k))
        Flags = ""
        If InStr(i.Flags, "Z") > 0 Then Flags = Flags & "ZF "
        If InStr(i.Flags, "C") > 0 Then Flags = Flags & "CF "
        If InStr(i.Flags, "S") > 0 Then Flags = Flags & "SF "
        If InStr(i.Flags, "O") > 0 Then Flags = Flags & "OF "
        If Flags = "" Then Flags = "-"
        ws.Cells(f + 1 + k, 1).Value = "'" & HexH(ops(k))
        ws.Cells(f + 1 + k, 2).Value = IsaFormato(ops(k))
        ws.Cells(f + 1 + k, 3).Value = i.Bytes
        ws.Cells(f + 1 + k, 4).Value = t(i.desc)
        ws.Cells(f + 1 + k, 5).Value = Trim$(Flags)
        ws.Cells(f + 1 + k, 6).Value = IsaModo(ops(k))
    Next k
End Sub

' =====================================================================
'  API DE PRUEBAS (la usan los scripts de verificacion, sin interfaz)
' =====================================================================
Public Function Test_Cargar(ByVal ejemplo As String) As String
    On Error GoTo fallo
    EjemploCargar ejemplo
    CpuReset
    LogLimpiar
    Test_Cargar = Ensamblar()
    VistaInicial
    Exit Function
fallo:
    Test_Cargar = "ERR " & Err.Description
End Function

Public Function Test_Correr(ByVal maxMicro As Long) As String
    Dim n As Long, animaba As Variant
    On Error GoTo fallo
    animaba = CpuGet(R_ANIMAR)
    CpuSet R_ANIMAR, False
    Do While n < maxMicro
        If Not MicroPaso() Then Exit Do
        n = n + 1
        If CpuNum(R_HALT) = 1 Then Exit Do
    Loop
    CpuSet R_ANIMAR, animaba
    Test_Correr = IIf(CpuNum(R_HALT) = 1, "HALT", "CORRIENDO") & " micro=" & n & _
                  " instr=" & CpuNum(R_INSTR) & " PC=" & HexH(CpuNum(R_PC)) & _
                  " AX=" & CpuNum(R_AX) & " BX=" & CpuNum(R_BX) & " " & TextoFlags()
    Exit Function
fallo:
    CpuSet R_ANIMAR, animaba
    Test_Correr = "ERR " & Err.Number & " (" & Err.Source & "): " & Err.Description
End Function

Public Function Test_Mem(ByVal d As Long) As Long
    Test_Mem = MemRead(d)
End Function

' Ensambla una sola linea en el editor (para probar errores del ensamblador)
Public Function Test_Linea(ByVal linea As String) As String
    Dim cod As Range
    On Error GoTo fallo
    Set cod = Sim().Range("EDITOR_COD")
    Application.EnableEvents = False
    cod.ClearContents
    cod.Cells(1, 1).Value = "'" & linea
    Application.EnableEvents = True
    Test_Linea = Ensamblar() & " | bytes: " & CStr(Sim().Range("EDITOR_HEX").Cells(1, 1).Value)
    Exit Function
fallo:
    Application.EnableEvents = True
    Test_Linea = "ERR " & Err.Description
End Function

' =====================================================================
'  DOCUMENTACION (la usa build\Exportar-Doc.ps1 para el README)
' =====================================================================
' Tabla de la ISA en Markdown, desde ModISA (una sola fuente)
Public Function Doc_IsaMarkdown() As String
    Dim ops As Variant, k As Long, i As TInstr, s As String, fl As String
    s = "| Opcode | Instrucci" & ChrW(243) & "n | Bytes | Operaci" & ChrW(243) & "n | Flags | Modo |" & vbLf & _
        "|:---:|---|:---:|---|:---:|---|" & vbLf
    ops = IsaListaOpcodes()
    For k = LBound(ops) To UBound(ops)
        i = IsaInfo(ops(k))
        fl = ""
        If InStr(i.Flags, "Z") > 0 Then fl = fl & "ZF "
        If InStr(i.Flags, "C") > 0 Then fl = fl & "CF "
        If InStr(i.Flags, "S") > 0 Then fl = fl & "SF "
        If InStr(i.Flags, "O") > 0 Then fl = fl & "OF "
        If fl = "" Then fl = "-"
        s = s & "| `" & HexH(ops(k)) & "` | `" & IsaFormato(ops(k)) & "` | " & i.Bytes & " | " & _
            t(i.desc) & " | " & Trim$(fl) & " | " & IsaModo(ops(k)) & " |" & vbLf
    Next k
    Doc_IsaMarkdown = s
End Function

' Traza instruccion a instruccion de un ejemplo, en Markdown
Public Function Doc_TrazaMarkdown(ByVal ejemplo As String, ByVal maxInstr As Long) As String
    Dim s As String, n As Long, pcIni As Long, c0 As Long, animaba As Variant, texto As String
    Dim op As Long, arg As Long, mem As String, g As String
    On Error GoTo fallo
    Test_Cargar ejemplo
    animaba = CpuGet(R_ANIMAR)
    CpuSet R_ANIMAR, False
    s = "| # | Dir | Instrucci" & ChrW(243) & "n | Ciclos | PC | AX | BX | ZF | CF | SF | OF | Efecto |" & vbLf & _
        "|---:|:---:|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|---|" & vbLf
    Do
        pcIni = CpuNum(R_PC)
        c0 = CpuNum(R_CICLOS)
        Do
            If Not MicroPaso() Then Exit Do
            If EnFrontera() Then Exit Do
        Loop
        n = n + 1
        op = CpuNum(R_IROP): arg = CpuNum(R_IRARG)
        texto = IsaTexto(op, arg)
        g = IsaInfo(op).Grupo
        Select Case g
            Case "MEMW": mem = "M[" & HexH(arg) & "] = " & HexH(MemRead(arg))
            Case "MEMR": mem = IsaInfo(op).Op1 & " = M[" & HexH(arg) & "]"
            Case "SALTO"
                If IsaCondicion(op) = "" Then
                    mem = "salta a " & HexH(arg)
                ElseIf CpuNum(R_SALTO) = 1 Then
                    mem = IsaCondicion(op) & ": salta a " & HexH(arg)
                Else
                    mem = "no salta"
                End If
            Case "CTRL": mem = IIf(IsaInfo(op).Mnem = "HLT", "fin", "-")
            Case "CMP": mem = "solo flags"
            Case Else: mem = "-"
        End Select
        s = s & "| " & n & " | `" & HexH(pcIni) & "` | `" & texto & "` | " & (CpuNum(R_CICLOS) - c0) & _
            " | " & HexH(CpuNum(R_PC)) & " | " & CpuNum(R_AX) & " | " & CpuNum(R_BX) & " | " & _
            CpuNum(R_ZF) & " | " & CpuNum(R_CF) & " | " & CpuNum(R_SF) & " | " & CpuNum(R_OF) & " | " & mem & " |" & vbLf
        If CpuNum(R_HALT) = 1 Or n >= maxInstr Then Exit Do
    Loop
    CpuSet R_ANIMAR, animaba
    Doc_TrazaMarkdown = s & vbLf & "Total: " & n & " instrucciones, " & CpuNum(R_CICLOS) & " ciclos de reloj."
    Exit Function
fallo:
    CpuSet R_ANIMAR, animaba
    Doc_TrazaMarkdown = "ERR " & Err.Description
End Function

' Ejecuta una operacion de la ALU aislada y devuelve resultado + flags
Public Function Test_Alu(ByVal mn As String, ByVal a As Long, ByVal b As Long) As String
    Dim r As Long
    On Error GoTo fallo
    CpuSet R_ZF, 0: CpuSet R_CF, 0: CpuSet R_SF, 0: CpuSet R_OF, 0
    r = AluOperar(mn, a, b)
    Test_Alu = r & " " & TextoFlags()
    Exit Function
fallo:
    Test_Alu = "ERR " & Err.Description
End Function
