Attribute VB_Name = "ModSecuenciador"
Option Explicit

' =====================================================================
'  ModSecuenciador - la Unidad de Control.
'
'  Cada instruccion se descompone en micro-operaciones agrupadas en las
'  4 fases del ciclo de instruccion:
'
'    FETCH    F1 MAR <- PC      F2 MDR <- M[MAR]   F3 IR <- MDR   F4 PC <- PC+1
'    DECODE   D1 el Decodificador interpreta el opcode
'             (si la instruccion ocupa 2 bytes, prepara el operando:
'              D2 MAR <- PC  D3 MDR <- M[MAR]  D4 IR.op <- MDR  D5 PC <- PC+1)
'    EXECUTE  la ALU opera, se accede a memoria o se evalua el salto
'    STORE    write-back: resultado al registro destino o a memoria
'
'  El microprograma de la instruccion en curso se escribe en la hoja
'  oculta _Micro (A fase | B codigo | C texto RTL | D accion | E arg1 | F arg2)
'  y se ejecuta una fila por cada pulso de reloj (STEP).
'
'  Acciones:  XFER src dst | MEMREAD | MEMWRITE | INCPC | DECODE
'             ALU op | CHK cond | JMPIF | HALT | NOTA
' =====================================================================

Private Const N_MAX_MICRO As Long = 40

' =====================================================================
'  UN PULSO DE RELOJ = UNA MICRO-OPERACION
'  Devuelve False si la CPU no puede avanzar (HLT o sin programa).
' =====================================================================
Public Function MicroPaso() As Boolean
    If CpuNum(R_HALT) = 1 Then
        Estado t("CPU detenida por HLT. Pulsa RESET (o LOAD para recargar el programa)."), False
        MicroPaso = False: Exit Function
    End If
    If CpuNum(R_PROGOK) <> 1 Then
        Estado t("No hay programa en memoria: escr~ibelo y pulsa LOAD."), True
        MicroPaso = False: Exit Function
    End If

    If CpuNum(R_MICRO) < 1 Or CpuNum(R_MICRO) > CpuNum(R_NMICRO) Then IniciarInstruccion

    MicroEjecutar CpuNum(R_MICRO)
    CpuSet R_CICLOS, CpuNum(R_CICLOS) + 1
    CpuSet R_MICRO, CpuNum(R_MICRO) + 1

    If CpuNum(R_MICRO) > CpuNum(R_NMICRO) Then
        CpuSet R_INSTR, CpuNum(R_INSTR) + 1
        If CpuNum(R_HALT) = 1 Then CpuSet R_FASE, "HALT"
    End If
    RelojMostrar
    MicroPaso = True
End Function

' True si la proxima micro-operacion es la primera de una instruccion
Public Function EnFrontera() As Boolean
    EnFrontera = (CpuNum(R_MICRO) < 1 Or CpuNum(R_MICRO) > CpuNum(R_NMICRO))
End Function

' =====================================================================
'  CONSTRUCCION DEL MICROPROGRAMA
' =====================================================================
Private Sub LimpiarMicro()
    Hoja(HOJA_MICRO).Range("A2:F" & (N_MAX_MICRO + 1)).ClearContents
    CpuSet R_NMICRO, 0
End Sub

Private Sub AM(ByVal fase As String, ByVal codigo As String, ByVal texto As String, _
               ByVal accion As String, Optional ByVal a1 As String = "", Optional ByVal a2 As String = "")
    Dim f As Long
    f = CpuNum(R_NMICRO) + 2
    With Hoja(HOJA_MICRO)
        .Cells(f, 1).Value = fase
        .Cells(f, 2).Value = codigo
        .Cells(f, 3).Value = t(texto)
        .Cells(f, 4).Value = accion
        .Cells(f, 5).Value = a1
        .Cells(f, 6).Value = a2
    End With
    CpuSet R_NMICRO, CpuNum(R_NMICRO) + 1
End Sub

Private Sub IniciarInstruccion()
    Dim pc As Long, fila As Long
    pc = CpuNum(R_PC)
    CpuSet R_PCINSTR, pc
    LimpiarMicro
    AM "FETCH", "F1", "MAR <- PC", "XFER", "PC", "MAR"
    AM "FETCH", "F2", "MDR <- M[MAR]", "MEMREAD"
    AM "FETCH", "F3", "IR <- MDR", "XFER", "MDR", "IROP"
    AM "FETCH", "F4", "PC <- PC + 1", "INCPC"
    AM "DECODE", "D1", "Decodificar el opcode del IR", "DECODE"
    CpuSet R_MICRO, 1

    fila = MemFilaEditor(pc)
    EditorResaltar fila
    MemMarcarPC pc
    DecodLimpiar
    MicroTitulo t("Microprogram: instrucci~on en ") & HexH(pc)
    MicroListaMostrar
End Sub

' Se llama al ejecutar D1: ya se conoce el opcode, se completa el resto
Private Sub CompletarMicroprograma(ByVal op As Long)
    Dim i As TInstr, d As String, src As String, cond As String
    i = IsaInfo(op)

    If i.Bytes = 2 Then
        AM "DECODE", "D2", "MAR <- PC", "XFER", "PC", "MAR"
        AM "DECODE", "D3", "MDR <- M[MAR]", "MEMREAD"
        AM "DECODE", "D4", "IR.op <- MDR", "XFER", "MDR", "IRARG"
        AM "DECODE", "D5", "PC <- PC + 1", "INCPC"
    End If

    If i.Op2 = "IMM" Then src = "IRARG" Else src = i.Op2
    d = i.Op1

    Select Case i.Grupo
        Case "CTRL"
            If i.Mnem = "HLT" Then
                ' HLT para el reloj en EXECUTE: no hay fase STORE detras
                AM "EXECUTE", "E1", "Detener el reloj", "HALT"
            Else
                AM "EXECUTE", "E1", "NOP: no hay operaci~on", "NOTA"
                AM "STORE", "S1", "Sin escritura", "NOTA"
            End If

        Case "TRANS"
            AM "EXECUTE", "E1", "MOV no usa la ALU: los flags no cambian", "NOTA"
            AM "STORE", "S1", d & " <- " & nombre(src), "XFER", src, d

        Case "MEMR"            ' LOAD reg, [dir]
            AM "EXECUTE", "E1", "MAR <- IR.op", "XFER", "IRARG", "MAR"
            AM "EXECUTE", "E2", "MDR <- M[MAR]", "MEMREAD"
            AM "STORE", "S1", d & " <- MDR", "XFER", "MDR", d

        Case "MEMW"            ' STORE [dir], reg
            AM "EXECUTE", "E1", "MAR <- IR.op", "XFER", "IRARG", "MAR"
            AM "EXECUTE", "E2", "MDR <- " & i.Op2, "XFER", i.Op2, "MDR"
            AM "STORE", "S1", "M[MAR] <- MDR", "MEMWRITE"

        Case "ALU", "CMP"
            AM "EXECUTE", "E1", "TMP1 <- " & d, "XFER", d, "REN1"
            AM "EXECUTE", "E2", "TMP2 <- " & nombre(src), "XFER", src, "REN2"
            AM "EXECUTE", "E3", "ALU: TMP1 " & IsaSimbolo(op) & " TMP2, flags", "ALU", i.Mnem
            If i.Grupo = "CMP" Then
                AM "STORE", "S1", "CMP descarta el resultado: solo quedan los flags", "NOTA"
            Else
                AM "STORE", "S1", d & " <- ALU", "XFER", "ALU", d
            End If

        Case "UNARIA"          ' INC, DEC, NOT
            AM "EXECUTE", "E1", "TMP1 <- " & d, "XFER", d, "REN1"
            If i.Mnem = "NOT" Then
                AM "EXECUTE", "E2", "ALU: NOT TMP1 (sin flags)", "ALU", "NOT"
            Else
                AM "EXECUTE", "E2", "TMP2 <- 1", "XFER", "UNO", "REN2"
                AM "EXECUTE", "E3", "ALU: TMP1 " & IsaSimbolo(op) & ", flags (CF intacto)", "ALU", i.Mnem
            End If
            AM "STORE", "S1", d & " <- ALU", "XFER", "ALU", d

        Case "SALTO"
            cond = IsaCondicion(op)
            If Len(cond) = 0 Then
                AM "EXECUTE", "E1", "PC <- IR.op (salto incondicional)", "XFER", "IRARG", "PC"
            Else
                AM "EXECUTE", "E1", "Sequencer consulta " & cond, "CHK", cond
                AM "EXECUTE", "E2", "Si se cumple: PC <- IR.op", "JMPIF"
            End If
            AM "STORE", "S1", "Sin escritura (el PC ya apunta a la siguiente)", "NOTA"
    End Select
End Sub

Private Function nombre(ByVal reg As String) As String
    Select Case reg
        Case "IRARG": nombre = "IR.op"
        Case "IROP": nombre = "IR"
        Case "UNO": nombre = "1"
        Case "REN1": nombre = "TMP1"
        Case "REN2": nombre = "TMP2"
        Case Else: nombre = reg
    End Select
End Function

' Componente del bus/diagrama asociado a un registro (para elegir la ruta)
Private Function Comp(ByVal reg As String) As String
    Select Case reg
        Case "IROP", "IRARG": Comp = "IR"
        Case Else: Comp = reg
    End Select
End Function

' =====================================================================
'  EJECUCION DE UNA MICRO-OPERACION
' =====================================================================
Private Sub MicroEjecutar(ByVal k As Long)
    Dim ws As Worksheet, fase As String, codigo As String, texto As String
    Dim accion As String, a1 As String, a2 As String
    Dim v As Long, d As Long, r As Long, op As Long, detalle As String, narr As String

    Set ws = Hoja(HOJA_MICRO)
    fase = ws.Cells(k + 1, 1).Value
    codigo = ws.Cells(k + 1, 2).Value
    texto = ws.Cells(k + 1, 3).Value
    accion = ws.Cells(k + 1, 4).Value
    a1 = ws.Cells(k + 1, 5).Value
    a2 = ws.Cells(k + 1, 6).Value

    CpuSet R_FASE, fase
    FaseMostrar fase
    MicroListaResaltar k
    ApagarRutas
    MemVentanaResaltar ""

    Select Case accion
        Case "XFER"
            If a1 = "ALU" Then v = CpuNum(R_ALU) Else v = RegLeer(a1)
            Narrar fase & " | " & texto & ": " & ExplicarXfer(codigo, a1, a2, v)
            Transferir Comp(a1) & ">" & Comp(a2), HexH(v)
            If a2 = "IROP" Then
                CpuSet R_IROP, v: RegPintar "IR"
            ElseIf a2 = "IRARG" Then
                CpuSet R_IRARG, v: RegPintar "IR"
            Else
                RegEscribir a2, v
            End If
            If a2 = "PC" Then MemMarcarPC v
            Destello ChipDe(a2)
            detalle = nombre(a2) & "=" & Hex0x(v)

        Case "MEMREAD"
            d = CpuNum(R_MAR)
            v = MemRead(d)
            PrimitivaMostrar "Read(" & HexH(d) & t(") -> ") & HexH(v), "R"
            Narrar fase & " | " & texto & ": Read(" & HexH(d) & t("). El Selector localiza la celda ") & HexH(d) & _
                   t(" y su contenido (") & HexH(v) & t(") llega al MDR por el bus de datos.")
            Transferir "MAR>SEL", HexH(d)
            RegEscribir "SEL", d
            Transferir "SEL>MEM", HexH(d)
            MemMarcarAcceso d, "R"
            MemVentanaResaltar "R"
            Transferir "MEM>MDR", HexH(v)
            RegEscribir "MDR", v
            Destello "VAL_MDR"
            detalle = "MAR=" & Hex0x(d) & ", MDR=" & Hex0x(v) & "  Read(" & Hex0x(d) & ")"

        Case "MEMWRITE"
            d = CpuNum(R_MAR)
            v = CpuNum(R_MDR)
            Narrar fase & " | " & texto & ": Write(" & HexH(d) & ", " & HexH(v) & t("). El dato del MDR se escribe en la celda ") & _
                   HexH(d) & "."
            Transferir "MAR>SEL", HexH(d)
            RegEscribir "SEL", d
            Transferir "SEL>MEM", HexH(d)
            Transferir "MDR>MEM", HexH(v)
            MemWrite d, v
            PrimitivaMostrar "Write(" & HexH(d) & ", " & HexH(v) & ")", "W"
            MemMarcarAcceso d, "W"
            MemVentanaResaltar "W"
            detalle = "RAM[" & Hex0x(d) & "] <- " & Hex0x(v) & "  Write(" & Hex0x(d) & ", " & Hex0x(v) & ")"

        Case "INCPC"
            v = (CpuNum(R_PC) + 1) And &HFF
            Narrar fase & " | " & texto & t(": el PC se incrementa y apunta a ") & HexH(v) & "."
            Transferir "PC+1", HexH(v)
            RegEscribir "PC", v
            MemMarcarPC v
            Destello "VAL_PC"
            detalle = "PC=" & Hex0x(v)

        Case "DECODE"
            op = CpuNum(R_IROP)
            Transferir "IR>DEC", HexH(op)
            DecodMostrar op
            If Not IsaExiste(op) Then
                Narrar fase & " | Opcode " & HexH(op) & t(" desconocido: la CPU se detiene.")
                CpuSet R_HALT, 1
                Estado t("Opcode inv~alido ") & HexH(op) & " en " & HexH(CpuNum(R_PCINSTR)) & t(". ~?Se est~a ejecutando un dato?"), True
                detalle = "opcode invalido " & Hex0x(op)
            Else
                ' el Decodificador comunica al Secuenciador QUE operacion
                ' hay que secuenciar: el dato que viaja es la mnemonica
                Transferir "DEC>SEC", IsaInfo(op).Mnem
                CompletarMicroprograma op
                MicroListaMostrar
                MicroListaResaltar k
                MicroTitulo "Microprogram: " & IsaFormato(op)
                Narrar fase & " | " & t("El Decoder interpreta el opcode ") & HexH(op) & ": " & _
                       IsaFormato(op) & " (" & IsaInfo(op).Bytes & t(" byte(s), modo ") & IsaModo(op) & _
                       t(") y avisa al Sequencer, que prepara las control signals.")
                detalle = "IR=" & Hex0x(op) & " -> " & IsaFormato(op)
            End If

        Case "ALU"
            EncenderRuta "REN1>ALU"
            If a1 <> "NOT" Then EncenderRuta "REN2>ALU"      ' NOT es unaria: solo usa REN1
            Pausa 0.1
            v = CpuNum(R_REN1)
            r = AluOperar(a1, v, CpuNum(R_REN2))
            ChipAlu IsaSimbolo(CpuNum(R_IROP))
            Transferir "ALU>FLG", TextoFlagsCorto()
            Narrar fase & " | " & texto & ": " & ExplicarAlu(a1, v, CpuNum(R_REN2), r)
            detalle = "ALU=" & Hex0x(r) & "  " & TextoFlags()

        Case "CHK"
            CpuSet R_SALTO, IIf(CondicionCumple(a1), 1, 0)
            Transferir "FLG>SEC", a1
            Narrar fase & " | " & t("El Sequencer consulta el flag ") & Left$(a1, 2) & "=" & _
                   IIf(FlagLeer(Left$(a1, 2)), 1, 0) & t(": la condici~on ") & a1 & _
                   IIf(CpuNum(R_SALTO) = 1, t(" SE CUMPLE, se tomar~a el salto."), t(" NO se cumple, no se salta."))
            detalle = a1 & " -> " & IIf(CpuNum(R_SALTO) = 1, "salta", "no salta")

        Case "JMPIF"
            If CpuNum(R_SALTO) = 1 Then
                v = CpuNum(R_IRARG)
                Narrar fase & " | " & t("Salto tomado: PC <- IR.op = ") & HexH(v) & "."
                Transferir "IR>PC", HexH(v)
                RegEscribir "PC", v
                MemMarcarPC v
                Destello "VAL_PC"
                detalle = "PC=" & Hex0x(v) & " (salto)"
            Else
                Narrar fase & " | " & t("No se salta: el PC sigue en ") & HexH(CpuNum(R_PC)) & "."
                detalle = "sin salto"
            End If

        Case "HALT"
            CpuSet R_HALT, 1
            Narrar fase & " | " & t("HLT: el reloj se detiene. Fin del programa.")
            Estado t("Programa terminado (HLT) tras ") & (CpuNum(R_INSTR) + 1) & t(" instrucciones y ") & _
                   (CpuNum(R_CICLOS) + 1) & " ciclos."
            detalle = "reloj detenido"

        Case "NOTA"
            Narrar fase & " | " & texto & "."
            detalle = "-"
    End Select

    LogAgregar fase, codigo, texto, detalle
End Sub

Private Function TextoFlagsCorto() As String
    TextoFlagsCorto = CpuNum(R_ZF) & CpuNum(R_CF) & CpuNum(R_SF) & CpuNum(R_OF)
End Function

Private Function ExplicarXfer(ByVal cod As String, ByVal src As String, ByVal dst As String, ByVal v As Long) As String
    Select Case True
        Case cod = "F1"
            ExplicarXfer = t("la direcci~on de la siguiente instrucci~on (") & HexH(v) & t(") sale por el address bus.")
        Case cod = "F3"
            ExplicarXfer = t("el opcode (") & HexH(v) & t(") pasa del MDR al IR (Instruction Register).")
        Case cod = "D2"
            ExplicarXfer = t("la instrucci~on ocupa 2 bytes, as~i que se busca tambi~en su operando en ") & HexH(v) & "."
        Case cod = "D4"
            ExplicarXfer = t("el operando (") & HexH(v) & t(") se guarda en la segunda mitad del IR.")
        Case dst = "MAR"
            ExplicarXfer = t("la direcci~on del operando (") & HexH(v) & t(") pasa al MAR.")
        Case src = "IRARG" And (dst = "AX" Or dst = "BX")
            ExplicarXfer = t("el dato inmediato ") & HexH(v) & " (" & v & t(") del IR se copia en ") & dst & _
                           t(" por el internal data bus, sin pasar por la ALU.")
        Case (src = "AX" Or src = "BX") And (dst = "AX" Or dst = "BX")
            ExplicarXfer = t("el contenido de ") & src & " (" & HexH(v) & t(") se copia en ") & dst & "."
        Case dst = "REN1" Or dst = "REN2"
            ExplicarXfer = t("el valor ") & HexH(v) & " (" & v & t(") entra en ") & nombre(dst) & t(" de la ALU.")
        Case src = "ALU"
            ExplicarXfer = t("write-back: el resultado ") & HexH(v) & " (" & v & t(") se guarda en ") & dst & "."
        Case src = "MDR"
            ExplicarXfer = t("write-back: el dato le~ido (") & HexH(v) & t(") se guarda en ") & dst & "."
        Case dst = "MDR"
            ExplicarXfer = t("el dato de ") & src & " (" & HexH(v) & t(") pasa al MDR para escribirse en memoria.")
        Case dst = "PC"
            ExplicarXfer = t("salto incondicional: el PC pasa a ") & HexH(v) & "."
        Case Else
            ExplicarXfer = t("el valor ") & HexH(v) & t(" pasa de ") & nombre(src) & " a " & dst & "."
    End Select
End Function

Private Function ExplicarAlu(ByVal mn As String, ByVal a As Long, ByVal b As Long, ByVal r As Long) As String
    Dim s As String
    Select Case mn
        Case "NOT": s = "NOT " & HexH(a) & " = " & HexH(r)
        Case "INC": s = HexH(a) & " + 1 = " & HexH(r)
        Case "DEC": s = HexH(a) & " - 1 = " & HexH(r)
        Case "ADD": s = HexH(a) & " + " & HexH(b) & " = " & HexH(r) & " (" & a & " + " & b & " = " & r & IIf(a + b > 255, t(", con acarreo"), "") & ")"
        Case "SUB", "CMP": s = HexH(a) & " - " & HexH(b) & " = " & HexH(r) & " (" & a & " - " & b & t(" = ") & ConSigno(r) & ")"
        Case Else: s = HexH(a) & " " & mn & " " & HexH(b) & " = " & HexH(r)
    End Select
    ExplicarAlu = s & ". Flags: " & TextoFlags() & "."
End Function

' =====================================================================
'  LOG DE MICRO-OPERACIONES (hoja Log)
'  A # | B instr | C fase | D micro | E operacion | F detalle |
'  G..L PC IR MAR MDR AX BX | M..P ZF CF SF OF | Q linea formateada
' =====================================================================
Public Sub LogAgregar(ByVal fase As String, ByVal codigo As String, ByVal texto As String, ByVal detalle As String)
    Dim ws As Worksheet, n As Long, f As Long, linea As String, fila(1 To 17) As Variant
    Set ws = Hoja(HOJA_LOG)
    n = CpuNum(R_LOGN) + 1
    CpuSet R_LOGN, n
    f = n + 1
    linea = "[" & Format$(n, "0000") & "] " & fase & " " & codigo & "  " & texto & "   " & detalle
    fila(1) = n: fila(2) = CpuNum(R_INSTR) + 1: fila(3) = fase: fila(4) = codigo
    fila(5) = texto: fila(6) = detalle
    ' con el prefijo 0x (el formato del enunciado) Excel ya no lo toma por
    ' un numero guardado como texto
    fila(7) = Hex0x(CpuNum(R_PC)): fila(8) = "'" & TextoIR()
    fila(9) = Hex0x(CpuNum(R_MAR)): fila(10) = Hex0x(CpuNum(R_MDR))
    fila(11) = Hex0x(CpuNum(R_AX)): fila(12) = Hex0x(CpuNum(R_BX))
    fila(13) = CpuNum(R_ZF): fila(14) = CpuNum(R_CF): fila(15) = CpuNum(R_SF): fila(16) = CpuNum(R_OF)
    fila(17) = linea
    ws.Range(ws.Cells(f, 1), ws.Cells(f, 17)).Value = fila
    ' el IR ("10 00") si parece un numero en configuraciones regionales que
    ' separan los miles con espacio: se marca el aviso como ignorado
    On Error Resume Next
    ws.Cells(f, 8).Errors.Item(xlNumberAsText).Ignore = True
    On Error GoTo 0
End Sub

Public Sub LogLimpiar()
    With Hoja(HOJA_LOG)
        .Range("A2:Q" & .Rows.Count).ClearContents
    End With
    CpuSet R_LOGN, 0
End Sub
