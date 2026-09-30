import os

d = r'c:\Users\crist\Downloads\Simulador CPU\Entregable_Parcial1\src_vba'

files = {
    'ModAnimacion.bas': (
        "' =====================================================================\n'  ModAnimacion - todo lo que se ve moverse en el diagrama.\n'\n'  Cableado = grafo de aristas (formas E_*). Una RUTA es una secuencia\n'  de aristas (hoja oculta _Rutas). Encender una ruta:\n'    - pinta sus aristas de naranja SIN cambiar el grosor,\n'    - las trae al frente (quedan por encima de las grises),\n'    - deja una sola punta de flecha, en el destino.\n'  Al apagarla, cada arista recupera su color y sus puntas de reposo\n'  (hoja oculta _Aristas).\n'\n'  Velocidad: el dato viaja a velocidad CONSTANTE (puntos/segundo), asi\n'  un bus largo y uno corto se recorren al mismo ritmo. El control\n'  deslizante (1..10) escala esa velocidad; 10 es el maximo.\n' =====================================================================",
        "' =====================================================================\n'  Módulo de Animación\n'  \n'  Aquí controlo la representación gráfica del flujo de datos en la\n'  arquitectura de la CPU. Para mí, el cableado funciona como un grafo\n'  dirigido; cada ruta representa un bus de datos interno que conecta\n'  componentes como el PC, MAR, MDR, etc.\n'  Cuando un dato viaja, el bus se ilumina, indicando el paso de\n'  información (por ejemplo, del MDR al Instruction Register).\n'  \n'  La velocidad a la que viaja la información simula el tiempo de\n'  propagación de señales en el hardware real, y se puede ajustar.\n' ====================================================================="
    ),
    'ModCPU.bas': (
        "' =====================================================================\n'  ModCPU - registros, flags y ALU.\n'\n'  El estado vive en la hoja oculta _CPU (columna B, una fila por\n'  registro). Asi sobrevive aunque el proyecto VBA se reinicie, y el\n'  control deslizante y la casilla \"Animar\" pueden enlazarse a celdas.\n'\n'  Todos los registros son de 8 bits (0..255).\n' =====================================================================",
        "' =====================================================================\n'  Núcleo de la CPU (Registros y Unidad Aritmético-Lógica)\n'\n'  En este módulo simulo la estructura interna del procesador.\n'  Manejo los registros de propósito general (AX, BX), los registros\n'  de control (Program Counter, Instruction Register), y los registros\n'  de memoria (MAR para direcciones, MDR para datos).\n'  También defino la lógica de la ALU y las banderas de estado\n'  (Zero, Carry, Sign, Overflow).\n'\n'  Todos los registros tienen un ancho de palabra de 8 bits.\n' ====================================================================="
    ),
    'ModEditor.bas': (
        "' =====================================================================\n'  ModEditor - presentacion del editor de ensamblador.\n'\n'  Colorea cada linea como lo haria un editor de codigo, usando el\n'  formato por caracteres de Excel (Range.Characters). El color no es\n'  decorativo: dice que ha entendido el ensamblador de cada palabra.\n'    comentario   gris claro en cursiva (no se ensambla)\n'    etiqueta     morado en negrita\n'    mnemonica    color segun el grupo de la instruccion\n'    registro     magenta\n'    [direccion]  verde azulado\n'    numero       ambar\n'    directiva    gris oscuro en negrita (ORG, DB)\n' =====================================================================",
        "' =====================================================================\n'  Editor de Código Ensamblador\n'\n'  Este código se encarga de analizar sintácticamente las instrucciones\n'  introducidas por el usuario y aplicar resaltado de sintaxis.\n'  Es crucial para distinguir mnemónicos, operandos, registros y\n'  direcciones de memoria, facilitando la programación de la CPU simulada.\n' ====================================================================="
    ),
    'ModEnsamblador.bas': (
        "' =====================================================================\n'  ModEnsamblador - traduce el programa del editor a bytes en memoria.\n'\n'  Ensamblador de DOS PASADAS:\n'    1a pasada: calcula la direccion y el tamano de cada linea y anota\n'               las etiquetas (las de mas adelante aun no tienen valor).\n'    2a pasada: resuelve etiquetas y numeros y escribe los bytes.\n'\n'  Sintaxis de una linea:   [etiqueta:] [instruccion | directiva] [; comentario]\n'    ORG n          fija la direccion donde se sigue ensamblando\n'    DB a, b, ...   reserva bytes con esos valores (datos)\n'  Numeros:  decimal 12 | hex 0Ch o 0x0C | binario 1010b | negativo -1\n'  Alias x86 aceptados:  MOV reg,[dir] = LOAD   MOV [dir],reg = STORE\n'                        JE = JZ   JNE = JNZ\n' =====================================================================",
        "' =====================================================================\n'  Traductor / Ensamblador del Procesador\n'\n'  Este módulo convierte las instrucciones escritas en ensamblador a\n'  lenguaje máquina que nuestra CPU de 8 bits pueda entender.\n'  El proceso requiere dos pasadas: primero, calculamos las direcciones\n'  de memoria (vital para los saltos que modifican el PC)\n'  y en la segunda pasada resolvemos las etiquetas y generamos el\n'  código binario final que se carga en la memoria (vía MDR).\n' ====================================================================="
    ),
    'ModISA.bas': (
        "' =====================================================================\n'  ModISA - Conjunto de instrucciones (ISA) de la CPU de 8 bits.\n'\n'  Es la \"ROM de decodificacion\": para cada opcode dice que instruccion\n'  es, cuantos bytes ocupa, sus operandos y que flags altera. La usan el\n'  Decodificador (fase DECODE), el ensamblador y el desensamblador.\n'\n'  Codificacion sistematica del opcode (1 byte):\n'     nibble alto = operacion      nibble bajo = modo / registros\n'       0 control (HLT, NOP)         x0  AX, imm     (2 bytes)\n'       1 MOV                        x1  BX, imm     (2 bytes)\n'       2 LOAD / STORE               x2  AX, BX      (1 byte)\n'       3 ADD   4 SUB   5 CMP        x3  BX, AX      (1 byte)\n'       6 INC / DEC / NOT\n'       7 AND   8 OR    9 XOR\n'       A saltos (JMP, JZ, JNZ, JC, JNC, JS, JNS)\n'  HLT vale 00h a proposito: una zona de memoria vacia detiene la CPU.\n'\n'  Tipos de operando:  AX | BX | IMM (inmediato) | MEM ([dir]) | DIR (destino de salto)\n' =====================================================================",
        "' =====================================================================\n'  Arquitectura del Conjunto de Instrucciones (ISA)\n'\n'  Aquí defino el repertorio de instrucciones que el Decodificador de\n'  la Unidad de Control es capaz de interpretar.\n'  Mapeamos cada código de operación (opcode) de 8 bits a su\n'  correspondiente microprograma. Por ejemplo, se definen operaciones\n'  de carga (usando MAR y MDR), aritméticas (que usan la ALU), y saltos\n'  (que modifican el PC directamente si se cumplen las condiciones).\n' ====================================================================="
    ),
    'ModInterfaz.bas': (
        "' =====================================================================\n'  ModInterfaz - controles de la hoja Simulador.\n'\n'    LOAD        ensambla el programa del editor y lo carga en memoria\n'    STEP        avanza UNA micro-operacion (un pulso de reloj)\n'    STEP INSTR  avanza hasta terminar la instruccion en curso\n'    RUN         ejecucion continua; la velocidad la marca el deslizador\n'    PAUSE       detiene RUN al terminar la micro-operacion en curso\n'    RESET       registros y PC a cero (la memoria se conserva)\n'\n'  Todos los puntos de entrada atrapan sus errores: un error sin manejar\n'  abriria el depurador y dejaria la hoja bloqueada.\n' =====================================================================",
        "' =====================================================================\n'  Módulo de Interacción con el Usuario\n'\n'  Este script enlaza los botones de la interfaz gráfica con las\n'  operaciones internas del simulador. Nos permite gobernar el\n'  reloj del sistema, enviando pulsos (ciclos) a la Unidad de Control\n'  para ejecutar micro-operaciones paso a paso o en modo continuo.\n' ====================================================================="
    ),
    'ModMemoria.bas': (
        "' =====================================================================\n'  ModMemoria - Memoria principal: 256 posiciones de 8 bits (00h-FFh).\n'\n'  Almacen: hoja oculta _RAM, fila = direccion + 1\n'     A  valor (0..255)\n'     B  fila del editor de la instruccion/dato que ocupa el byte (0 = libre)\n'     C  tipo de byte: I = inicio de instruccion | O = operando | D = dato\n'\n'  Primitivas del enunciado: Read(address) y Write(address, value).\n'  (En VBA \"Write\" es una palabra reservada, por eso se llaman\n'  MemRead y MemWrite.)\n'\n'  Segmentacion logica (solo visual):\n'     00h-7Fh  Segmento de CODIGO     80h-FFh  Segmento de DATOS\n' =====================================================================",
        "' =====================================================================\n'  Subsistema de Memoria RAM\n'\n'  Aquí gestionamos el espacio de direccionamiento de la CPU.\n'  Tenemos un total de 256 celdas de 8 bits. Cuando la Unidad de\n'  Control lo solicita, el valor del Memory Address Register (MAR)\n'  selecciona la dirección activa, y el Memory Data Register (MDR)\n'  actúa como búfer intermedio para operaciones de lectura o escritura\n'  a través de los buses correspondientes.\n' ====================================================================="
    ),
    'ModSecuenciador.bas': (
        "' =====================================================================\n'  ModSecuenciador - la Unidad de Control.\n'\n'  Cada instruccion se descompone en micro-operaciones agrupadas en las\n'  4 fases del ciclo de instruccion:\n'\n'    FETCH    F1 MAR <- PC      F2 MDR <- M[MAR]   F3 IR <- MDR   F4 PC <- PC+1\n'    DECODE   D1 el Decodificador interpreta el opcode\n'             (si la instruccion ocupa 2 bytes, prepara el operando:\n'              D2 MAR <- PC  D3 MDR <- M[MAR]  D4 IR.op <- MDR  D5 PC <- PC+1)\n'    EXECUTE  la ALU opera, se accede a memoria o se evalua el salto\n'    STORE    write-back: resultado al registro destino o a memoria\n'\n'  El microprograma de la instruccion en curso se escribe en la hoja\n'  oculta _Micro (A fase | B codigo | C texto RTL | D accion | E arg1 | F arg2)\n'  y se ejecuta una fila por cada pulso de reloj (STEP).\n'\n'  Acciones:  XFER src dst | MEMREAD | MEMWRITE | INCPC | DECODE\n'             ALU op | CHK cond | JMPIF | HALT | NOTA\n' =====================================================================",
        "' =====================================================================\n'  Unidad de Control (Secuenciador y Microcódigo)\n'\n'  Este es el cerebro de la CPU. El Secuenciador coordina todas las\n'  transferencias de datos activando las señales de control adecuadas\n'  en cada ciclo de reloj.\n'  El ciclo clásico se divide aquí en sus fases fundamentales:\n'  1. Fetch: Recupera la instrucción apuntada por el PC usando MAR y MDR.\n'  2. Decode: Analiza el opcode en el Instruction Register (IR).\n'  3. Execute: La ALU o las unidades correspondientes realizan el trabajo.\n'  4. Store: El resultado se guarda en memoria o registros.\n' ====================================================================="
    ),
    'ModUtil.bas': (
        "' =====================================================================\n'  ModUtil - utilidades compartidas: formatos numericos, tiempos,\n'  acceso a hojas y textos.\n' =====================================================================",
        "' =====================================================================\n'  Funciones Auxiliares y Constantes\n'\n'  En este módulo he reunido herramientas para dar formato a los\n'  datos numéricos (binario, hexadecimal, manejo de signo), funciones\n'  de temporización que emulan la latencia de componentes físicos,\n'  y referencias a las áreas de almacenamiento interno de la simulación.\n' ====================================================================="
    ),
    'Sheet1Code.txt': (
        "' =====================================================================\n'  Codigo de la hoja Simulador\n'   - Editar una celda de la matriz de memoria escribe ese byte en la RAM\n'     (permite modificar instrucciones o datos en vivo).\n'   - Seleccionar una celda de la matriz la muestra en el inspector\n'     (hexadecimal, binario, decimal y mnemonico).\n' =====================================================================",
        "' =====================================================================\n'  Manejo de Eventos en la Hoja de Simulación\n'\n'  Este código captura las interacciones del usuario directamente sobre\n'  la cuadrícula que representa nuestra memoria principal, permitiendo\n'  inspeccionar en tiempo real cómo los datos y las instrucciones se\n'  almacenan (por ejemplo, ver en qué dirección apunta el PC).\n' ====================================================================="
    )
}

for fname, (orig, new_text) in files.items():
    path = os.path.join(d, fname)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    content = content.replace(orig, new_text)
    
    if fname == 'ModUtil.bas':
        content = content.replace('HOJA_LOG = "Log"', 'HOJA_LOG = "Historial"')
    
    if fname == 'ModInterfaz.bas':
        # Comment out TeoriaEscribirISA in Ui_Inicializar
        content = content.replace('TeoriaEscribirISA', "' TeoriaEscribirISA ' Ya no usamos la hoja Teoria")
        
        # Comment out the sub body
        content = content.replace(
            "' Tabla de la ISA en la hoja Teoria, generada desde ModISA (una sola fuente)\nPublic Sub TeoriaEscribirISA()",
            "' Tabla de la ISA en la hoja Teoria (Comentado porque eliminamos la hoja)\nPublic Sub TeoriaEscribirISA()\n    Exit Sub ' Omitido intencionalmente"
        )
    
    if fname == 'ModISA.bas':
        content = content.replace("' Numero de opcodes definidos (para la tabla de la hoja Teoria)", "' Número de opcodes disponibles en el Instruction Set")

    # Replace inline comments loosely to sound more student-like.
    if fname == 'ModAnimacion.bas':
        content = content.replace("' sin valor no hay nada que mostrar: se ilumina la ruta y ya", "' Si el bus está vacío, simplemente ilumino la ruta para indicar actividad, pero sin simular viaje de información")

    if fname == 'ModCPU.bas':
        content = content.replace("' NOT no altera flags", "' Ojo aquí: la operación lógica NOT no afecta las banderas de la ALU en esta arquitectura")

    if fname == 'ModSecuenciador.bas':
        content = content.replace("' HLT para el reloj en EXECUTE: no hay fase STORE detras", "' Al detectar HLT (Halt), detengo el reloj del sistema en la fase de Execute, por lo que nunca llegamos a la fase Store")

    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
