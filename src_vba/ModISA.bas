Attribute VB_Name = "ModISA"
Option Explicit

' =====================================================================
'  Arquitectura del Conjunto de Instrucciones (ISA)
'
'  Aquí defino el repertorio de instrucciones que el Decodificador de
'  la Unidad de Control es capaz de interpretar.
'  Mapeamos cada código de operación (opcode) de 8 bits a su
'  correspondiente microprograma. Por ejemplo, se definen operaciones
'  de carga (usando MAR y MDR), aritméticas (que usan la ALU), y saltos
'  (que modifican el PC directamente si se cumplen las condiciones).
' =====================================================================

Public Type TInstr
    Opcode As Long
    Mnem As String
    Op1 As String
    Op2 As String
    Bytes As Long
    Flags As String      ' letras de los flags que altera: Z C S O
    Grupo As String      ' CTRL | TRANS | MEMR | MEMW | ALU | CMP | UNARIA | SALTO
    desc As String       ' semantica en notacion RTL (con marcadores de T())
End Type

Private mTabla(0 To 255) As TInstr
Private mDefinido(0 To 255) As Boolean
Private mListo As Boolean

' /////////////////////////////////////////////////////////////////
'  DEFINICION DE LA TABLA
' /////////////////////////////////////////////////////////////////
Private Sub Def(ByVal op As Long, ByVal mn As String, ByVal o1 As String, ByVal o2 As String, _
                ByVal nb As Long, ByVal fl As String, ByVal gr As String, ByVal ds As String)
    With mTabla(op)
        .Opcode = op: .Mnem = mn: .Op1 = o1: .Op2 = o2
        .Bytes = nb: .Flags = fl: .Grupo = gr: .desc = ds
    End With
    mDefinido(op) = True
End Sub

' Las 4 variantes de una operacion de dos operandos (x0..x3)
Private Sub Def4(ByVal base As Long, ByVal mn As String, ByVal fl As String, _
                 ByVal gr As String, ByVal simb As String)
    Def base + 0, mn, "AX", "IMM", 2, fl, gr, "AX <- AX " & simb & " imm"
    Def base + 1, mn, "BX", "IMM", 2, fl, gr, "BX <- BX " & simb & " imm"
    Def base + 2, mn, "AX", "BX", 1, fl, gr, "AX <- AX " & simb & " BX"
    Def base + 3, mn, "BX", "AX", 1, fl, gr, "BX <- BX " & simb & " AX"
End Sub

Public Sub IsaInit()
    If mListo Then Exit Sub

    Def &H0, "HLT", "", "", 1, "", "CTRL", "Detiene el reloj (fin del programa)"
    Def &H1, "NOP", "", "", 1, "", "CTRL", "No opera"

    Def &H10, "MOV", "AX", "IMM", 2, "", "TRANS", "AX <- imm"
    Def &H11, "MOV", "BX", "IMM", 2, "", "TRANS", "BX <- imm"
    Def &H12, "MOV", "AX", "BX", 1, "", "TRANS", "AX <- BX"
    Def &H13, "MOV", "BX", "AX", 1, "", "TRANS", "BX <- AX"

    Def &H20, "LOAD", "AX", "MEM", 2, "", "MEMR", "AX <- M[dir]"
    Def &H21, "LOAD", "BX", "MEM", 2, "", "MEMR", "BX <- M[dir]"
    Def &H22, "STORE", "MEM", "AX", 2, "", "MEMW", "M[dir] <- AX"
    Def &H23, "STORE", "MEM", "BX", 2, "", "MEMW", "M[dir] <- BX"

    Def4 &H30, "ADD", "ZCSO", "ALU", "+"
    Def4 &H40, "SUB", "ZCSO", "ALU", "-"
    Def4 &H50, "CMP", "ZCSO", "CMP", "-"
    ' CMP no guarda el resultado: se reescribe su semantica
    mTabla(&H50).desc = "flags <- AX - imm": mTabla(&H51).desc = "flags <- BX - imm"
    mTabla(&H52).desc = "flags <- AX - BX":  mTabla(&H53).desc = "flags <- BX - AX"

    Def &H60, "INC", "AX", "", 1, "ZSO", "UNARIA", "AX <- AX + 1"
    Def &H61, "INC", "BX", "", 1, "ZSO", "UNARIA", "BX <- BX + 1"
    Def &H62, "DEC", "AX", "", 1, "ZSO", "UNARIA", "AX <- AX - 1"
    Def &H63, "DEC", "BX", "", 1, "ZSO", "UNARIA", "BX <- BX - 1"
    Def &H64, "NOT", "AX", "", 1, "", "UNARIA", "AX <- NOT AX"
    Def &H65, "NOT", "BX", "", 1, "", "UNARIA", "BX <- NOT BX"

    Def4 &H70, "AND", "ZCSO", "ALU", "AND"
    Def4 &H80, "OR", "ZCSO", "ALU", "OR"
    Def4 &H90, "XOR", "ZCSO", "ALU", "XOR"

    Def &HA0, "JMP", "DIR", "", 2, "", "SALTO", "PC <- dir"
    Def &HA1, "JZ", "DIR", "", 2, "", "SALTO", "si ZF=1: PC <- dir"
    Def &HA2, "JNZ", "DIR", "", 2, "", "SALTO", "si ZF=0: PC <- dir"
    Def &HA3, "JC", "DIR", "", 2, "", "SALTO", "si CF=1: PC <- dir"
    Def &HA4, "JNC", "DIR", "", 2, "", "SALTO", "si CF=0: PC <- dir"
    Def &HA5, "JS", "DIR", "", 2, "", "SALTO", "si SF=1: PC <- dir"
    Def &HA6, "JNS", "DIR", "", 2, "", "SALTO", "si SF=0: PC <- dir"

    mListo = True
End Sub

' /////////////////////////////////////////////////////////////////
'  CONSULTAS
' /////////////////////////////////////////////////////////////////
Public Function IsaExiste(ByVal op As Long) As Boolean
    IsaInit
    If op < 0 Or op > 255 Then IsaExiste = False Else IsaExiste = mDefinido(op)
End Function

Public Function IsaInfo(ByVal op As Long) As TInstr
    IsaInit
    IsaInfo = mTabla(op And &HFF)
End Function

' Opcode para una mnemonica y tipos de operando; -1 si no existe
Public Function IsaBuscar(ByVal mn As String, ByVal t1 As String, ByVal t2 As String) As Long
    Dim op As Long
    IsaInit
    For op = 0 To 255
        If mDefinido(op) Then
            With mTabla(op)
                If .Mnem = mn And .Op1 = t1 And .Op2 = t2 Then IsaBuscar = op: Exit Function
            End With
        End If
    Next op
    IsaBuscar = -1
End Function

' Grupo al que pertenece una mnemonica (lo usa el coloreado del editor)
Public Function IsaGrupoDeMnemonica(ByVal mn As String) As String
    Dim op As Long
    IsaInit
    For op = 0 To 255
        If mDefinido(op) Then
            If mTabla(op).Mnem = mn Then IsaGrupoDeMnemonica = mTabla(op).Grupo: Exit Function
        End If
    Next op
    IsaGrupoDeMnemonica = ""
End Function

Public Function IsaEsMnemonica(ByVal mn As String) As Boolean
    Dim op As Long
    IsaInit
    For op = 0 To 255
        If mDefinido(op) Then
            If mTabla(op).Mnem = mn Then IsaEsMnemonica = True: Exit Function
        End If
    Next op
    IsaEsMnemonica = False
End Function

' Texto de un operando para el desensamblado
Private Function TextoOperando(ByVal tipo As String, ByVal arg As Long) As String
    Select Case tipo
        Case "AX", "BX": TextoOperando = tipo
        Case "IMM": TextoOperando = HexH(arg)
        Case "MEM": TextoOperando = "[" & HexH(arg) & "]"
        Case "DIR": TextoOperando = HexH(arg)
        Case Else: TextoOperando = ""
    End Select
End Function

' Desensamblado: "ADD AX, 05h"  "LOAD BX, [80h]"  "JZ 0Ah"
Public Function IsaTexto(ByVal op As Long, ByVal arg As Long) As String
    Dim s As String
    If Not IsaExiste(op) Then IsaTexto = "DB " & HexH(op): Exit Function
    With mTabla(op)
        s = .Mnem
        If Len(.Op1) > 0 Then s = s & " " & TextoOperando(.Op1, arg)
        If Len(.Op2) > 0 Then s = s & ", " & TextoOperando(.Op2, arg)
    End With
    IsaTexto = s
End Function

' Formato generico: "ADD AX, imm8"
Public Function IsaFormato(ByVal op As Long) As String
    Dim s As String
    If Not IsaExiste(op) Then IsaFormato = "?": Exit Function
    With mTabla(op)
        s = .Mnem
        If Len(.Op1) > 0 Then s = s & " " & NombreTipo(.Op1)
        If Len(.Op2) > 0 Then s = s & ", " & NombreTipo(.Op2)
    End With
    IsaFormato = s
End Function

Private Function NombreTipo(ByVal tipo As String) As String
    Select Case tipo
        Case "IMM": NombreTipo = "imm8"
        Case "MEM": NombreTipo = "[dir]"
        Case "DIR": NombreTipo = "dir"
        Case Else: NombreTipo = tipo
    End Select
End Function

' Modo de direccionamiento, para el Decodificador
Public Function IsaModo(ByVal op As Long) As String
    If Not IsaExiste(op) Then IsaModo = "?": Exit Function
    With mTabla(op)
        If .Op1 = "MEM" Or .Op2 = "MEM" Then
            IsaModo = "directo (memoria)"
        ElseIf .Op1 = "DIR" Then
            IsaModo = "salto directo"
        ElseIf .Op2 = "IMM" Then
            IsaModo = "inmediato"
        ElseIf Len(.Op1) > 0 Then
            IsaModo = "registro"
        Else
            IsaModo = t("impl~icito")
        End If
    End With
End Function

' Condicion de un salto: "" (incondicional), "ZF=1", "CF=0"...
Public Function IsaCondicion(ByVal op As Long) As String
    Select Case op
        Case &HA1: IsaCondicion = "ZF=1"
        Case &HA2: IsaCondicion = "ZF=0"
        Case &HA3: IsaCondicion = "CF=1"
        Case &HA4: IsaCondicion = "CF=0"
        Case &HA5: IsaCondicion = "SF=1"
        Case &HA6: IsaCondicion = "SF=0"
        Case Else: IsaCondicion = ""
    End Select
End Function

' Simbolo que se muestra dentro de la ALU (C.OP)
Public Function IsaSimbolo(ByVal op As Long) As String
    IsaInit
    Select Case mTabla(op And &HFF).Mnem
        Case "ADD": IsaSimbolo = "+"
        Case "SUB", "CMP": IsaSimbolo = "-"
        Case "INC": IsaSimbolo = "+1"
        Case "DEC": IsaSimbolo = "-1"
        Case Else: IsaSimbolo = mTabla(op And &HFF).Mnem
    End Select
End Function

' Número de opcodes disponibles en el Instruction Set
Public Function IsaListaOpcodes() As Variant
    Dim op As Long, n As Long, r() As Long
    IsaInit
    ReDim r(0 To 255)
    For op = 0 To 255
        If mDefinido(op) Then r(n) = op: n = n + 1
    Next op
    ReDim Preserve r(0 To n - 1)
    IsaListaOpcodes = r
End Function
