Attribute VB_Name = "ModEnsamblador"
Option Explicit

' =====================================================================
'  ModEnsamblador - traduce el programa del editor a bytes en memoria.
'
'  Ensamblador de DOS PASADAS:
'    1a pasada: calcula la direccion y el tamano de cada linea y anota
'               las etiquetas (las de mas adelante aun no tienen valor).
'    2a pasada: resuelve etiquetas y numeros y escribe los bytes.
'
'  Sintaxis de una linea:   [etiqueta:] [instruccion | directiva] [; comentario]
'    ORG n          fija la direccion donde se sigue ensamblando
'    DB a, b, ...   reserva bytes con esos valores (datos)
'  Numeros:  decimal 12 | hex 0Ch o 0x0C | binario 1010b | negativo -1
'  Alias x86 aceptados:  MOV reg,[dir] = LOAD   MOV [dir],reg = STORE
'                        JE = JZ   JNE = JNZ
' =====================================================================

Private Const N_LINEAS As Long = 40

Private mLoc(1 To N_LINEAS) As Long       ' direccion de la linea
Private mTam(1 To N_LINEAS) As Long       ' bytes que ocupa
Private mTipo(1 To N_LINEAS) As String    ' "" vacia | I instruccion | D datos | ORG
Private mOp(1 To N_LINEAS) As Long        ' opcode (si es instruccion)
Private mArg(1 To N_LINEAS) As String     ' texto del operando que va al 2o byte
Private mDB(1 To N_LINEAS) As String      ' lista de valores de un DB
Private mError As String
Private mErrorFila As Long

' =====================================================================
'  PUNTO DE ENTRADA
'  Devuelve "OK ..." o "ERR ..." (el texto tambien va a la barra de estado)
' =====================================================================
Public Function Ensamblar() As String
    Dim ws As Worksheet, cod As Range, etiquetas As Object
    Dim i As Long, total As Long, maxDir As Long, minDato As Long

    Set ws = Sim()
    Set cod = ws.Range("EDITOR_COD")
    Set etiquetas = CreateObject("Scripting.Dictionary")
    etiquetas.CompareMode = vbTextCompare
    mError = "": mErrorFila = 0

    Application.EnableEvents = False
    cod.Interior.color = RGB(255, 255, 255)
    ws.Range("EDITOR_DIR").ClearContents
    ws.Range("EDITOR_HEX").ClearContents
    Application.EnableEvents = True

    If Not PrimeraPasada(cod, etiquetas) Then GoTo fallo
    MemReset
    If Not SegundaPasada(cod, etiquetas, total, maxDir, minDato) Then GoTo fallo

    CpuSet R_PROGOK, 1
    EditorColorearTodo
    Ensamblar = "OK " & total & " bytes"
    Estado t("Programa cargado: ") & total & t(" bytes. C~odigo en 00h-") & HexH(maxDir) & _
           IIf(minDato <= 255, ", datos desde " & HexH(minDato), "") & t(". Pulsa STEP o RUN.")
    Exit Function

fallo:
    CpuSet R_PROGOK, 0
    EditorColorearTodo
    If mErrorFila > 0 Then cod.Cells(mErrorFila, 1).Interior.color = RGB(255, 199, 206)
    Ensamblar = "ERR linea " & mErrorFila & ": " & mError
    Estado t("Error en la l~inea ") & mErrorFila & ": " & mError, True
End Function

' =====================================================================
'  PRIMERA PASADA
' =====================================================================
Private Function PrimeraPasada(ByVal cod As Range, ByVal etiquetas As Object) As Boolean
    Dim i As Long, loc As Long, linea As String, etq As String, resto As String
    Dim mn As String, ops() As String, t1 As String, t2 As String, op As Long, v As Long

    loc = 0
    For i = 1 To N_LINEAS
        mLoc(i) = loc: mTam(i) = 0: mTipo(i) = "": mOp(i) = -1: mArg(i) = "": mDB(i) = ""
        linea = QuitarComentario(CStr(cod.Cells(i, 1).Value))
        If Len(linea) = 0 Then GoTo siguiente

        SepararEtiqueta linea, etq, resto
        If Len(etq) > 0 Then
            If Not EsIdentificador(etq) Then Falla i, t("Etiqueta inv~alida: ") & etq: Exit Function
            If etiquetas.Exists(etq) Then Falla i, "Etiqueta repetida: " & etq: Exit Function
            If EsRegistro(etq) Or IsaEsMnemonica(UCase$(etq)) Then Falla i, t("Nombre reservado como etiqueta: ") & etq: Exit Function
        End If

        SepararMnemonica resto, mn, ops
        If Len(mn) = 0 Then                                ' solo etiqueta
            If Len(etq) > 0 Then etiquetas(etq) = loc
            GoTo siguiente
        End If

        Select Case mn
            Case "ORG"
                If UBound(ops) <> 0 Then Falla i, "ORG necesita una direccion": Exit Function
                If Not LeerNumero(ops(0), v) Or v > 255 Then Falla i, t("Direcci~on de ORG inv~alida: ") & ops(0): Exit Function
                loc = v: mLoc(i) = loc: mTipo(i) = "ORG"
                If Len(etq) > 0 Then etiquetas(etq) = loc

            Case "DB"
                If UBound(ops) < 0 Then Falla i, "DB necesita al menos un valor": Exit Function
                If Len(etq) > 0 Then etiquetas(etq) = loc
                mTipo(i) = "D": mTam(i) = UBound(ops) + 1: mDB(i) = Join(ops, ",")
                loc = loc + mTam(i)

            Case Else
                If Len(etq) > 0 Then etiquetas(etq) = loc
                If Not ResolverOpcode(i, mn, ops, op) Then Exit Function
                mTipo(i) = "I": mOp(i) = op: mTam(i) = IsaInfo(op).Bytes
                loc = loc + mTam(i)
        End Select

        If loc > 256 Then Falla i, t("El programa no cabe en 256 bytes"): Exit Function
siguiente:
    Next i
    PrimeraPasada = True
End Function

' Determina el opcode segun mnemonica y tipos de operando (y guarda el
' texto del operando que ira en el segundo byte)
Private Function ResolverOpcode(ByVal i As Long, ByVal mn As String, ops() As String, ByRef op As Long) As Boolean
    Dim t1 As String, t2 As String, n As Long

    n = UBound(ops) + 1
    If n > 2 Then Falla i, "Demasiados operandos": Exit Function
    If n >= 1 Then t1 = TipoOperando(ops(0))
    If n >= 2 Then t2 = TipoOperando(ops(1))

    ' alias x86
    If mn = "JE" Then mn = "JZ"
    If mn = "JNE" Then mn = "JNZ"
    If mn = "MOV" And t2 = "MEM" Then mn = "LOAD"
    If mn = "MOV" And t1 = "MEM" Then mn = "STORE"
    ' en los saltos el operando es una direccion de destino, no un dato
    If Left$(mn, 1) = "J" And t1 = "IMM" Then t1 = "DIR"

    If Not IsaEsMnemonica(mn) Then Falla i, t("Instrucci~on desconocida: ") & mn: Exit Function
    op = IsaBuscar(mn, t1, t2)
    If op < 0 Then
        Falla i, t("Combinaci~on no soportada: ") & mn & " " & DescribirTipos(t1, t2)
        Exit Function
    End If

    ' el operando que viaja en el 2o byte (inmediato, [dir] o destino)
    If IsaInfo(op).Bytes = 2 Then
        If t1 = "IMM" Or t1 = "MEM" Or t1 = "DIR" Then mArg(i) = ops(0) Else mArg(i) = ops(1)
    End If
    ResolverOpcode = True
End Function

' =====================================================================
'  SEGUNDA PASADA
' =====================================================================
Private Function SegundaPasada(ByVal cod As Range, ByVal etiquetas As Object, _
                               ByRef total As Long, ByRef maxDir As Long, ByRef minDato As Long) As Boolean
    Dim ws As Worksheet, i As Long, k As Long, v As Long, bytesTxt As String
    Dim Ocupado(0 To 255) As Boolean, vals() As String, d As Long

    Set ws = Sim()
    total = 0: maxDir = 0: minDato = 999
    Application.EnableEvents = False

    For i = 1 To N_LINEAS
        Select Case mTipo(i)
            Case "I"
                bytesTxt = Hex2(mOp(i))
                If Not Reservar(i, mLoc(i), Ocupado) Then GoTo fin
                MemWrite mLoc(i), mOp(i)
                MemMarcarOrigen mLoc(i), i, "I"
                If mTam(i) = 2 Then
                    If Not valor(i, mArg(i), etiquetas, v) Then GoTo fin
                    If Not Reservar(i, mLoc(i) + 1, Ocupado) Then GoTo fin
                    MemWrite mLoc(i) + 1, v
                    MemMarcarOrigen mLoc(i) + 1, i, "O"
                    bytesTxt = bytesTxt & " " & Hex2(v)
                End If
                If mLoc(i) + mTam(i) - 1 > maxDir Then maxDir = mLoc(i) + mTam(i) - 1
                EscribirTexto ws.Range("EDITOR_DIR").Cells(i, 1), Hex2(mLoc(i))
                EscribirTexto ws.Range("EDITOR_HEX").Cells(i, 1), bytesTxt
                total = total + mTam(i)

            Case "D"
                vals = Split(mDB(i), ",")
                bytesTxt = ""
                For k = 0 To UBound(vals)
                    d = mLoc(i) + k
                    If Not valor(i, vals(k), etiquetas, v) Then GoTo fin
                    If Not Reservar(i, d, Ocupado) Then GoTo fin
                    MemWrite d, v
                    MemMarcarOrigen d, i, "D"
                    bytesTxt = bytesTxt & IIf(k > 0, " ", "") & Hex2(v)
                Next k
                If mLoc(i) < minDato Then minDato = mLoc(i)
                EscribirTexto ws.Range("EDITOR_DIR").Cells(i, 1), Hex2(mLoc(i))
                EscribirTexto ws.Range("EDITOR_HEX").Cells(i, 1), bytesTxt
                total = total + mTam(i)

            Case "ORG"
                EscribirTexto ws.Range("EDITOR_DIR").Cells(i, 1), Hex2(mLoc(i))
        End Select
    Next i
    SegundaPasada = True
fin:
    Application.EnableEvents = True
End Function

Private Function Reservar(ByVal i As Long, ByVal d As Long, Ocupado() As Boolean) As Boolean
    If d > 255 Then Falla i, t("Se sale de la memoria (m~as all~a de FFh)"): Exit Function
    If Ocupado(d) Then Falla i, t("Solapamiento: la direcci~on ") & HexH(d) & t(" ya est~a ocupada"): Exit Function
    Ocupado(d) = True
    Reservar = True
End Function

' Valor numerico de un operando: numero, etiqueta o [numero/etiqueta]
Private Function valor(ByVal i As Long, ByVal s As String, ByVal etiquetas As Object, ByRef v As Long) As Boolean
    Dim neg As Boolean
    s = Trim$(s)
    If Left$(s, 1) = "[" And Right$(s, 1) = "]" Then s = Trim$(Mid$(s, 2, Len(s) - 2))
    If Left$(s, 1) = "-" Then neg = True: s = Trim$(Mid$(s, 2))

    If LeerNumero(s, v) Then
        ' ok
    ElseIf etiquetas.Exists(s) Then
        v = etiquetas(s)
    ElseIf EsIdentificador(s) Then
        Falla i, "Etiqueta no definida: " & s: Exit Function
    Else
        Falla i, t("Valor inv~alido: ") & s: Exit Function
    End If

    If neg Then
        If v > 128 Then Falla i, t("Negativo fuera de rango (m~inimo -128): -") & v: Exit Function
        v = (256 - v) And &HFF
    ElseIf v > 255 Then
        Falla i, t("No cabe en 8 bits (m~aximo 255): ") & s: Exit Function
    End If
    valor = True
End Function

' =====================================================================
'  AYUDANTES DE PARSEO
' =====================================================================
Private Function QuitarComentario(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, ";")
    If p > 0 Then s = Left$(s, p - 1)
    QuitarComentario = Trim$(Replace(s, vbTab, " "))
End Function

Private Sub SepararEtiqueta(ByVal linea As String, ByRef etq As String, ByRef resto As String)
    Dim p As Long
    p = InStr(linea, ":")
    If p > 0 And InStr(Left$(linea, p), "[") = 0 Then
        etq = Trim$(Left$(linea, p - 1))
        resto = Trim$(Mid$(linea, p + 1))
    Else
        etq = "": resto = linea
    End If
End Sub

Private Sub SepararMnemonica(ByVal s As String, ByRef mn As String, ByRef ops() As String)
    Dim p As Long, lista As String, k As Long
    s = Trim$(s)
    If Len(s) = 0 Then mn = "": ops = Split(""): Exit Sub
    p = InStr(s, " ")
    If p = 0 Then
        mn = UCase$(s): lista = ""
    Else
        mn = UCase$(Left$(s, p - 1)): lista = Trim$(Mid$(s, p + 1))
    End If
    If Len(lista) = 0 Then
        ops = Split("")
    Else
        ops = Split(lista, ",")
        For k = 0 To UBound(ops)
            ops(k) = Trim$(ops(k))
        Next k
    End If
End Sub

Private Function TipoOperando(ByVal s As String) As String
    s = UCase$(Trim$(s))
    If s = "AX" Or s = "BX" Then
        TipoOperando = s
    ElseIf Left$(s, 1) = "[" And Right$(s, 1) = "]" Then
        TipoOperando = "MEM"
    ElseIf Len(s) = 0 Then
        TipoOperando = ""
    Else
        TipoOperando = "IMM"          ' numero o etiqueta
    End If
End Function

Private Function DescribirTipos(ByVal t1 As String, ByVal t2 As String) As String
    Dim s As String
    s = NombreT(t1)
    If Len(t2) > 0 Then s = s & ", " & NombreT(t2)
    DescribirTipos = s
End Function

' OJO: el parametro NO puede llamarse "t": VBA no distingue mayusculas y
' taparia la funcion T() de ModUtil.
Private Function NombreT(ByVal tipo As String) As String
    Select Case tipo
        Case "IMM": NombreT = "inmediato"
        Case "MEM": NombreT = "[memoria]"
        Case "DIR": NombreT = t("direcci~on")
        Case "": NombreT = "(nada)"
        Case Else: NombreT = tipo
    End Select
End Function

Private Function EsRegistro(ByVal s As String) As Boolean
    EsRegistro = (UCase$(s) = "AX" Or UCase$(s) = "BX")
End Function

Private Function EsIdentificador(ByVal s As String) As Boolean
    Dim i As Long, c As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = UCase$(Mid$(s, i, 1))
        If Not ((c >= "A" And c <= "Z") Or c = "_" Or (i > 1 And c >= "0" And c <= "9")) Then Exit Function
    Next i
    EsIdentificador = True
End Function

Private Sub Falla(ByVal fila As Long, ByVal msg As String)
    mErrorFila = fila
    mError = msg
End Sub

' =====================================================================
'  EJEMPLOS  (hoja oculta _Ejemplos: A nombre | B linea)
' =====================================================================
Public Sub EjemploCargar(ByVal nombre As String)
    Dim ws As Worksheet, r As Long, ultima As Long, k As Long, cod As Range
    Set ws = Hoja(HOJA_EJEMPLOS)
    Set cod = Sim().Range("EDITOR_COD")
    Application.EnableEvents = False
    cod.ClearContents
    cod.Interior.color = RGB(255, 255, 255)
    Sim().Range("EDITOR_DIR").ClearContents
    Sim().Range("EDITOR_HEX").ClearContents
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    k = 0
    For r = 2 To ultima
        If CStr(ws.Cells(r, 1).Value) = nombre Then
            k = k + 1
            If k > N_LINEAS Then Exit For
            cod.Cells(k, 1).Value = "'" & CStr(ws.Cells(r, 2).Value)
        End If
    Next r
    Application.EnableEvents = True
    EditorColorearTodo
End Sub
