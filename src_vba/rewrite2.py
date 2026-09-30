import os, re

d = r'c:\Users\crist\Downloads\Simulador CPU\Entregable_Parcial1\src_vba'

files = {
    'ModAnimacion.bas': (
        r"' ={20,}.*?Todo lo que se ve moverse en el diagrama.*?={20,}",
        "' =====================================================================\n'  Módulo de Animación\n'  \n'  Aquí controlo la representación gráfica del flujo de datos en la\n'  arquitectura de la CPU. Para mí, el cableado funciona como un grafo\n'  dirigido; cada ruta representa un bus de datos interno que conecta\n'  componentes como el PC, MAR, MDR, etc.\n'  Cuando un dato viaja, el bus se ilumina, indicando el paso de\n'  información (por ejemplo, del MDR al Instruction Register).\n'  \n'  La velocidad a la que viaja la información simula el tiempo de\n'  propagación de señales en el hardware real, y se puede ajustar.\n' ====================================================================="
    ),
    'ModCPU.bas': (
        r"' ={20,}.*?ModCPU - registros, flags y ALU.*?={20,}",
        "' =====================================================================\n'  Núcleo de la CPU (Registros y Unidad Aritmético-Lógica)\n'\n'  En este módulo simulo la estructura interna del procesador.\n'  Manejo los registros de propósito general (AX, BX), los registros\n'  de control (Program Counter, Instruction Register), y los registros\n'  de memoria (MAR para direcciones, MDR para datos).\n'  También defino la lógica de la ALU y las banderas de estado\n'  (Zero, Carry, Sign, Overflow).\n'\n'  Todos los registros tienen un ancho de palabra de 8 bits.\n' ====================================================================="
    ),
    'ModEditor.bas': (
        r"' ={20,}.*?ModEditor - presentacion del editor de ensamblador.*?={20,}",
        "' =====================================================================\n'  Editor de Código Ensamblador\n'\n'  Este código se encarga de analizar sintácticamente las instrucciones\n'  introducidas por el usuario y aplicar resaltado de sintaxis.\n'  Es crucial para distinguir mnemónicos, operandos, registros y\n'  direcciones de memoria, facilitando la programación de la CPU simulada.\n' ====================================================================="
    ),
    'ModEnsamblador.bas': (
        r"' ={20,}.*?ModEnsamblador - traduce el programa del editor a bytes en memoria.*?={20,}",
        "' =====================================================================\n'  Traductor / Ensamblador del Procesador\n'\n'  Este módulo convierte las instrucciones escritas en ensamblador a\n'  lenguaje máquina que nuestra CPU de 8 bits pueda entender.\n'  El proceso requiere dos pasadas: primero, calculamos las direcciones\n'  de memoria (vital para los saltos que modifican el PC)\n'  y en la segunda pasada resolvemos las etiquetas y generamos el\n'  código binario final que se carga en la memoria (vía MDR).\n' ====================================================================="
    ),
    'ModISA.bas': (
        r"' ={20,}.*?ModISA - Conjunto de instrucciones \(ISA\) de la CPU de 8 bits.*?={20,}",
        "' =====================================================================\n'  Arquitectura del Conjunto de Instrucciones (ISA)\n'\n'  Aquí defino el repertorio de instrucciones que el Decodificador de\n'  la Unidad de Control es capaz de interpretar.\n'  Mapeamos cada código de operación (opcode) de 8 bits a su\n'  correspondiente microprograma. Por ejemplo, se definen operaciones\n'  de carga (usando MAR y MDR), aritméticas (que usan la ALU), y saltos\n'  (que modifican el PC directamente si se cumplen las condiciones).\n' ====================================================================="
    ),
    'ModInterfaz.bas': (
        r"' ={20,}.*?ModInterfaz - controles de la hoja Simulador.*?={20,}",
        "' =====================================================================\n'  Módulo de Interacción con el Usuario\n'\n'  Este script enlaza los botones de la interfaz gráfica con las\n'  operaciones internas del simulador. Nos permite gobernar el\n'  reloj del sistema, enviando pulsos (ciclos) a la Unidad de Control\n'  para ejecutar micro-operaciones paso a paso o en modo continuo.\n' ====================================================================="
    ),
    'ModMemoria.bas': (
        r"' ={20,}.*?ModMemoria - Memoria principal: 256 posiciones de 8 bits.*?={20,}",
        "' =====================================================================\n'  Subsistema de Memoria RAM\n'\n'  Aquí gestionamos el espacio de direccionamiento de la CPU.\n'  Tenemos un total de 256 celdas de 8 bits. Cuando la Unidad de\n'  Control lo solicita, el valor del Memory Address Register (MAR)\n'  selecciona la dirección activa, y el Memory Data Register (MDR)\n'  actúa como búfer intermedio para operaciones de lectura o escritura\n'  a través de los buses correspondientes.\n' ====================================================================="
    ),
    'ModSecuenciador.bas': (
        r"' ={20,}.*?ModSecuenciador - la Unidad de Control.*?={20,}",
        "' =====================================================================\n'  Unidad de Control (Secuenciador y Microcódigo)\n'\n'  Este es el cerebro de la CPU. El Secuenciador coordina todas las\n'  transferencias de datos activando las señales de control adecuadas\n'  en cada ciclo de reloj.\n'  El ciclo clásico se divide aquí en sus fases fundamentales:\n'  1. Fetch: Recupera la instrucción apuntada por el PC usando MAR y MDR.\n'  2. Decode: Analiza el opcode en el Instruction Register (IR).\n'  3. Execute: La ALU o las unidades correspondientes realizan el trabajo.\n'  4. Store: El resultado se guarda en memoria o registros.\n' ====================================================================="
    ),
    'ModUtil.bas': (
        r"(' ={20,}|' /////////////////////////////////////////////////////////////////).*?ModUtil - utilidades compartidas.*?(' ={20,}|' /////////////////////////////////////////////////////////////////)",
        "' =====================================================================\n'  Funciones Auxiliares y Constantes\n'\n'  En este módulo he reunido herramientas para dar formato a los\n'  datos numéricos (binario, hexadecimal, manejo de signo), funciones\n'  de temporización que emulan la latencia de componentes físicos,\n'  y referencias a las áreas de almacenamiento interno de la simulación.\n' ====================================================================="
    ),
    'Sheet1Code.txt': (
        r"' ={20,}.*?Codigo de la hoja Simulador.*?={20,}",
        "' =====================================================================\n'  Manejo de Eventos en la Hoja de Simulación\n'\n'  Este código captura las interacciones del usuario directamente sobre\n'  la cuadrícula que representa nuestra memoria principal, permitiendo\n'  inspeccionar en tiempo real cómo los datos y las instrucciones se\n'  almacenan (por ejemplo, ver en qué dirección apunta el PC).\n' ====================================================================="
    )
}

for fname, (pattern, new_text) in files.items():
    path = os.path.join(d, fname)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # regex replace with DOTALL to match across newlines
    new_content = re.sub(pattern, new_text, content, flags=re.DOTALL | re.IGNORECASE)
    
    # just in case ModAnimacion uses "ModAnimacion - todo lo que se ve..."
    if fname == 'ModAnimacion.bas':
        new_content = re.sub(r"' ={20,}.*?ModAnimacion - todo lo que se ve moverse en el diagrama.*?={20,}", new_text, content, flags=re.DOTALL | re.IGNORECASE)
        
    if fname == 'ModUtil.bas':
        new_content = new_content.replace('HOJA_LOG = "Log"', 'HOJA_LOG = "Historial"')
    
    if fname == 'ModInterfaz.bas':
        # Comment out TeoriaEscribirISA in Ui_Inicializar
        new_content = re.sub(r"TeoriaEscribirISA(?!\s*')", "' TeoriaEscribirISA ' Ya no usamos la hoja Teoria", new_content)
        
        # Comment out the sub body
        new_content = re.sub(
            r"' Tabla de la ISA en la hoja Teoria, generada desde ModISA \(una sola fuente\)\nPublic Sub TeoriaEscribirISA\(\).*?End Sub",
            "' Tabla de la ISA en la hoja Teoria (Comentado porque eliminamos la hoja)\nPublic Sub TeoriaEscribirISA()\n    Exit Sub ' Omitido intencionalmente\nEnd Sub",
            new_content, flags=re.DOTALL
        )
    
    if fname == 'ModISA.bas':
        new_content = new_content.replace("' Numero de opcodes definidos (para la tabla de la hoja Teoria)", "' Número de opcodes disponibles en el Instruction Set")

    # Replace inline comments loosely to sound more student-like.
    if fname == 'ModAnimacion.bas':
        new_content = new_content.replace("' sin valor no hay nada que mostrar: se ilumina la ruta y ya", "' Si el bus está vacío, simplemente ilumino la ruta para indicar actividad, pero sin simular viaje de información")

    if fname == 'ModCPU.bas':
        new_content = new_content.replace("' NOT no altera flags", "' Ojo aquí: la operación lógica NOT no afecta las banderas de la ALU en esta arquitectura")

    if fname == 'ModSecuenciador.bas':
        new_content = new_content.replace("' HLT para el reloj en EXECUTE: no hay fase STORE detras", "' Al detectar HLT (Halt), detengo el reloj del sistema en la fase de Execute, por lo que nunca llegamos a la fase Store")

    with open(path, 'w', encoding='utf-8') as f:
        f.write(new_content)

print("Done")
