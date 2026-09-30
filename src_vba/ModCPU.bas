Attribute VB_Name = "ModCPU"
Option Explicit

' =====================================================================
'  ModCPU - registros, flags y ALU.
'
'  El estado vive en la hoja oculta _CPU (columna B, una fila por
'  registro). Asi sobrevive aunque el proyecto VBA se reinicie, y el
'  control deslizante y la casilla "Animar" pueden enlazarse a celdas.
'
'  Todos los registros son de 8 bits (0..255).
' =====================================================================

Public Const R_PC As Long = 1
Public Const R_IROP As Long = 2        ' IR: byte de opcode
Public Const R_IRARG As Long = 3       ' IR: byte de operando (inmediato o direccion)
Public Const R_MAR As Long = 4
Public Const R_MDR As Long = 5
Public Const R_AX As Long = 6
Public Const R_BX As Long = 7
Public Const R_REN1 As Long = 8        ' registros de entrada de la ALU (temporales)
Public Const R_REN2 As Long = 9
Public Const R_ALU As Long = 10        ' latch de salida de la ALU
Public Const R_ZF As Long = 11
Public Const R_CF As Long = 12
Public Const R_SF As Long = 13
Public Const R_OF As Long = 14
Public Const R_FASE As Long = 15       ' FETCH | DECODE | EXECUTE | STORE | HALT | READY
Public Const R_MICRO As Long = 16      ' proxima micro-operacion a ejecutar (fila en _Micro)
Public Const R_NMICRO As Long = 17     ' micro-operaciones de la instruccion en curso
Public Const R_CICLOS As Long = 18     ' micro-operaciones ejecutadas (ciclos de reloj T)
Public Const R_INSTR As Long = 19      ' instrucciones completadas
Public Const R_HALT As Long = 20
Public Const R_CORRIENDO As Long = 21
Public Const R_PAUSA As Long = 22
Public Const R_VELOC As Long = 23      ' enlazada al control deslizante (1..10)
Public Const R_ANIMAR As Long = 24     ' enlazada a la casilla "Animar datos"
Public Const R_SALTO As Long = 25      ' 1 si el salto condicional se toma
Public Const R_PROGOK As Long = 26     ' 1 si hay un programa ensamblado en memoria
Public Const R_LOGN As Long = 27
Public Const R_PCINSTR As Long = 28    ' direccion donde empezo la instruccion en curso
Public Const R_ULTMEM As Long = 29     ' ultima celda accedida (-1 ninguna)
Public Const R_ULTMEMT As Long = 30    ' "R" lectura | "W" escritura
Public Const R_ULTPC As Long = 31      ' celda marcada como PC en la matriz
Public Const R_ULTLINEA As Long = 32   ' fila del editor resaltada
Public Const R_SEL As Long = 33        ' direccion en el Selector

' =====================================================================
'  ESTADO
' =====================================================================
Public Function CpuGet(ByVal r As Long) As Variant
    CpuGet = ThisWorkbook.Worksheets(HOJA_CPU).Cells(r, 2).Value
End Function

Public Function CpuNum(ByVal r As Long) As Long
    CpuNum = CLng(Val(CStr(ThisWorkbook.Worksheets(HOJA_CPU).Cells(r, 2).Value)))
End Function

Public Sub CpuSet(ByVal r As Long, ByVal v As Variant)
    ThisWorkbook.Worksheets(HOJA_CPU).Cells(r, 2).Value = v
End Sub

' Fila de estado de un registro visible por su nombre
Public Function FilaReg(ByVal nombre As String) As Long
    Select Case UCase$(nombre)
        Case "PC": FilaReg = R_PC
        Case "IROP": FilaReg = R_IROP
        Case "IRARG", "IR": FilaReg = R_IRARG
        Case "MAR": FilaReg = R_MAR
        Case "MDR": FilaReg = R_MDR
        Case "AX": FilaReg = R_AX
        Case "BX": FilaReg = R_BX
        Case "REN1": FilaReg = R_REN1
        Case "REN2": FilaReg = R_REN2
        Case "ALU": FilaReg = R_ALU
        Case "SEL": FilaReg = R_SEL
        Case "UNO": FilaReg = 0
        Case Else: FilaReg = -1
    End Select
End Function

Public Function RegLeer(ByVal nombre As String) As Long
    Dim f As Long
    f = FilaReg(nombre)
    If f = 0 Then RegLeer = 1: Exit Function        ' la constante "1" de INC/DEC
    If f < 0 Then Err.Raise vbObjectError + 1, "RegLeer", "Registro desconocido: " & nombre
    RegLeer = CpuNum(f) And &HFF
End Function

' Escribe un registro (8 bits) y refresca su chip en el diagrama
Public Sub RegEscribir(ByVal nombre As String, ByVal v As Long)
    Dim f As Long
    v = v And &HFF
    f = FilaReg(nombre)
    If f <= 0 Then Err.Raise vbObjectError + 2, "RegEscribir", "Registro no escribible: " & nombre
    CpuSet f, v
    RegPintar nombre
    If UCase$(nombre) = "MAR" Then MemVentana v
End Sub

' Texto del chip de un registro
Public Function RegTexto(ByVal nombre As String) As String
    Dim v As Long
    Select Case UCase$(nombre)
        Case "AX", "BX"
            v = RegLeer(nombre)
            RegTexto = HexH(v) & " (" & v & ")"
        Case "IROP", "IRARG", "IR"
            RegTexto = TextoIR()
        Case Else
            RegTexto = HexH(RegLeer(nombre))
    End Select
End Function

' IR = opcode + operando. Si la instruccion es de 1 byte, el operando
' no se usa y se muestra "--".
Public Function TextoIR() As String
    Dim op As Long
    op = CpuNum(R_IROP) And &HFF
    If IsaExiste(op) Then
        If IsaInfo(op).Bytes = 1 Then TextoIR = Hex2(op) & " --": Exit Function
    End If
    TextoIR = Hex2(op) & " " & Hex2(CpuNum(R_IRARG))
End Function

Public Sub RegPintar(ByVal nombre As String)
    Select Case UCase$(nombre)
        Case "IROP", "IRARG", "IR": FijarTexto "VAL_IR", TextoIR()
        Case "ALU", "UNO": ' sin chip propio
        Case Else: FijarTexto "VAL_" & UCase$(nombre), RegTexto(nombre)
    End Select
End Sub

' =====================================================================
'  FLAGS
' =====================================================================
Public Sub FlagEscribir(ByVal nombre As String, ByVal b As Boolean)
    Dim r As Long, s As Shape
    Select Case nombre
        Case "ZF": r = R_ZF
        Case "CF": r = R_CF
        Case "SF": r = R_SF
        Case "OF": r = R_OF
    End Select
    CpuSet r, IIf(b, 1, 0)
    Set s = Forma("FLAG_" & nombre)
    If s Is Nothing Then Exit Sub
    s.TextFrame2.TextRange.Text = IIf(b, "1", "0")
    If b Then
        s.Fill.ForeColor.RGB = RGB(30, 58, 95)
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        s.Fill.ForeColor.RGB = RGB(255, 255, 255)
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End If
End Sub

Public Function FlagLeer(ByVal nombre As String) As Boolean
    Select Case nombre
        Case "ZF": FlagLeer = (CpuNum(R_ZF) = 1)
        Case "CF": FlagLeer = (CpuNum(R_CF) = 1)
        Case "SF": FlagLeer = (CpuNum(R_SF) = 1)
        Case "OF": FlagLeer = (CpuNum(R_OF) = 1)
    End Select
End Function

Public Function TextoFlags() As String
    TextoFlags = "ZF=" & CpuNum(R_ZF) & " CF=" & CpuNum(R_CF) & _
                 " SF=" & CpuNum(R_SF) & " OF=" & CpuNum(R_OF)
End Function

' Evalua la condicion de un salto ("ZF=1", "CF=0"...). "" = incondicional.
Public Function CondicionCumple(ByVal cond As String) As Boolean
    If Len(cond) = 0 Then CondicionCumple = True: Exit Function
    CondicionCumple = (FlagLeer(Left$(cond, 2)) = (Right$(cond, 1) = "1"))
End Function

' =====================================================================
'  ALU
'  Calcula a <op> b en 8 bits y actualiza SOLO los flags que la
'  instruccion altera (semantica x86):
'    ADD        CF = acarreo (resultado > 255)
'    SUB / CMP  CF = prestamo (a < b)
'    OF         desbordamiento con signo (complemento a 2)
'    INC / DEC  no tocan CF
'    AND OR XOR CF = 0, OF = 0
'    NOT        no altera ningun flag
' =====================================================================
Public Function AluOperar(ByVal mn As String, ByVal a As Long, ByVal b As Long) As Long
    Dim r As Long, bruto As Long
    a = a And &HFF: b = b And &HFF

    Select Case mn
        Case "ADD"
            bruto = a + b
            r = bruto And &HFF
            FlagEscribir "CF", (bruto > 255)
            FlagEscribir "OF", (((a Xor r) And (b Xor r) And &H80) <> 0)
        Case "SUB", "CMP"
            bruto = a - b
            r = (bruto + 256) And &HFF
            FlagEscribir "CF", (a < b)
            FlagEscribir "OF", (((a Xor b) And (a Xor r) And &H80) <> 0)
        Case "INC"
            r = (a + 1) And &HFF
            FlagEscribir "OF", (a = &H7F)
        Case "DEC"
            r = (a + 255) And &HFF
            FlagEscribir "OF", (a = &H80)
        Case "AND": r = a And b: FlagEscribir "CF", False: FlagEscribir "OF", False
        Case "OR":  r = a Or b:  FlagEscribir "CF", False: FlagEscribir "OF", False
        Case "XOR": r = a Xor b: FlagEscribir "CF", False: FlagEscribir "OF", False
        Case "NOT"
            r = (Not a) And &HFF
            CpuSet R_ALU, r
            AluOperar = r
            Exit Function                  ' NOT no altera flags
        Case Else
            Err.Raise vbObjectError + 3, "AluOperar", "Operacion ALU desconocida: " & mn
    End Select

    FlagEscribir "ZF", (r = 0)
    FlagEscribir "SF", ((r And &H80) <> 0)
    CpuSet R_ALU, r
    AluOperar = r
End Function

' =====================================================================
'  REINICIO Y REPINTADO
' =====================================================================
' RESET: registros y PC a cero (la memoria no se toca).
Public Sub CpuReset()
    Dim r As Long
    For r = R_PC To R_OF
        CpuSet r, 0
    Next r
    CpuSet R_FASE, "READY"
    CpuSet R_MICRO, 0
    CpuSet R_NMICRO, 0
    CpuSet R_CICLOS, 0
    CpuSet R_INSTR, 0
    CpuSet R_HALT, 0
    CpuSet R_SALTO, 0
    CpuSet R_PCINSTR, 0
    CpuSet R_SEL, 0
    CpuPintarTodo
End Sub

Public Sub CpuPintarTodo()
    Dim n As Variant
    For Each n In Array("PC", "IR", "MAR", "MDR", "AX", "BX", "REN1", "REN2", "SEL")
        RegPintar CStr(n)
    Next n
    FlagEscribir "ZF", FlagLeer("ZF")
    FlagEscribir "CF", FlagLeer("CF")
    FlagEscribir "SF", FlagLeer("SF")
    FlagEscribir "OF", FlagLeer("OF")
    MemVentana CpuNum(R_MAR)
End Sub
