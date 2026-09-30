import os
import re

d = r'c:\Users\crist\Downloads\Simulador CPU\Entregable_Parcial1\src_vba'

files_map = {
    'ModAnimacion.bas': (
        "' =====================================================================\n'  Módulo de Animación\n'  \n'  Aquí controlo la representación gráfica del flujo de datos en la\n'  arquitectura de la CPU. Para mí, el cableado funciona como un grafo\n'  dirigido; cada ruta representa un bus de datos interno que conecta\n'  componentes como el PC, MAR, MDR, etc.\n'  Cuando un dato viaja, el bus se ilumina, indicando el paso de\n'  información (por ejemplo, del MDR al Instruction Register).\n'  \n'  La velocidad a la que viaja la información simula el tiempo de\n'  propagación de señales en el hardware real, y se puede ajustar.\n' ====================================================================="
    ),
    'ModCPU.bas': (
        "' =====================================================================\n'  Núcleo de la CPU (Registros y Unidad Aritmético-Lógica)\n'\n'  En este módulo simulo la estructura interna del procesador.\n'  Manejo los registros de propósito general (AX, BX), los registros\n'  de control (Program Counter, Instruction Register), y los registros\n'  de memoria (MAR para direcciones, MDR para datos).\n'  También defino la lógica de la ALU y las banderas de estado\n'  (Zero, Carry, Sign, Overflow).\n'\n'  Todos los registros tienen un ancho de palabra de 8 bits.\n' ====================================================================="
    ),
    'ModEditor.bas': (
        "' =====================================================================\n'  Editor de Código Ensamblador\n'\n'  Este código se encarga de analizar sintácticamente las instrucciones\n'  introducidas por el usuario y aplicar resaltado de sintaxis.\n'  Es crucial para distinguir mnemónicos, operandos, registros y\n'  direcciones de memoria, facilitando la programación de la CPU simulada.\n' ====================================================================="
    ),
    'ModEnsamblador.bas': (
        "' =====================================================================\n'  Traductor / Ensamblador del Procesador\n'\n'  Este módulo convierte las instrucciones escritas en ensamblador a\n'  lenguaje máquina que nuestra CPU de 8 bits pueda entender.\n'  El proceso requiere dos pasadas: primero, calculamos las direcciones\n'  de memoria (vital para los saltos que modifican el PC)\n'  y en la segunda pasada resolvemos las etiquetas y generamos el\n'  código binario final que se carga en la memoria (vía MDR).\n' ====================================================================="
    ),
    'ModISA.bas': (
        "' =====================================================================\n'  Arquitectura del Conjunto de Instrucciones (ISA)\n'\n'  Aquí defino el repertorio de instrucciones que el Decodificador de\n'  la Unidad de Control es capaz de interpretar.\n'  Mapeamos cada código de operación (opcode) de 8 bits a su\n'  correspondiente microprograma. Por ejemplo, se definen operaciones\n'  de carga (usando MAR y MDR), aritméticas (que usan la ALU), y saltos\n'  (que modifican el PC directamente si se cumplen las condiciones).\n' ====================================================================="
    ),
    'ModInterfaz.bas': (
        "' =====================================================================\n'  Módulo de Interacción con el Usuario\n'\n'  Este script enlaza los botones de la interfaz gráfica con las\n'  operaciones internas del simulador. Nos permite gobernar el\n'  reloj del sistema, enviando pulsos (ciclos) a la Unidad de Control\n'  para ejecutar micro-operaciones paso a paso o en modo continuo.\n' ====================================================================="
    ),
    'ModMemoria.bas': (
        "' =====================================================================\n'  Subsistema de Memoria RAM\n'\n'  Aquí gestionamos el espacio de direccionamiento de la CPU.\n'  Tenemos un total de 256 celdas de 8 bits. Cuando la Unidad de\n'  Control lo solicita, el valor del Memory Address Register (MAR)\n'  selecciona la dirección activa, y el Memory Data Register (MDR)\n'  actúa como búfer intermedio para operaciones de lectura o escritura\n'  a través de los buses correspondientes.\n' ====================================================================="
    ),
    'ModSecuenciador.bas': (
        "' =====================================================================\n'  Unidad de Control (Secuenciador y Microcódigo)\n'\n'  Este es el cerebro de la CPU. El Secuenciador coordina todas las\n'  transferencias de datos activando las señales de control adecuadas\n'  en cada ciclo de reloj.\n'  El ciclo clásico se divide aquí en sus fases fundamentales:\n'  1. Fetch: Recupera la instrucción apuntada por el PC usando MAR y MDR.\n'  2. Decode: Analiza el opcode en el Instruction Register (IR).\n'  3. Execute: La ALU o las unidades correspondientes realizan el trabajo.\n'  4. Store: El resultado se guarda en memoria o registros.\n' ====================================================================="
    ),
    'ModUtil.bas': (
        "' =====================================================================\n'  Funciones Auxiliares y Constantes\n'\n'  En este módulo he reunido herramientas para dar formato a los\n'  datos numéricos (binario, hexadecimal, manejo de signo), funciones\n'  de temporización que emulan la latencia de componentes físicos,\n'  y referencias a las áreas de almacenamiento interno de la simulación.\n' ====================================================================="
    ),
    'Sheet1Code.txt': (
        "' =====================================================================\n'  Manejo de Eventos en la Hoja de Simulación\n'\n'  Este código captura las interacciones del usuario directamente sobre\n'  la cuadrícula que representa nuestra memoria principal, permitiendo\n'  inspeccionar en tiempo real cómo los datos y las instrucciones se\n'  almacenan (por ejemplo, ver en qué dirección apunta el PC).\n' ====================================================================="
    )
}

for fname, new_text in files_map.items():
    path = os.path.join(d, fname)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Remove misplaced comments at the very beginning of ModCPU.bas or others
    content = re.sub(r"^' (//+|==+)\n.*?\n' (//+|==+)\n", "", content, flags=re.DOTALL)
    
    # We want to replace the FIRST big comment block AFTER Option Explicit (or near top)
    # The pattern is: matches starting with ' ==== or ' ///// and ending with ' ==== or ' /////
    pattern = r"' (//+|==+).*?\n' (//+|==+)\n"
    content = re.sub(pattern, new_text + "\n", content, count=1, flags=re.DOTALL)
    
    if fname == 'ModUtil.bas':
        content = content.replace('HOJA_LOG = "Log"', 'HOJA_LOG = "Historial"')
        content = content.replace('HOJA_LOG = "Historial"\n', '') # safety cleanup of duplicate if we replaced over and over
        # Let's do a reliable replacement:
        content = re.sub(r'Public Const HOJA_LOG As String = ".*?"', 'Public Const HOJA_LOG As String = "Historial"', content)

    if fname == 'ModInterfaz.bas':
        # Comment out TeoriaEscribirISA in Ui_Inicializar
        content = re.sub(r"TeoriaEscribirISA(?!\s*')", "' TeoriaEscribirISA ' Ya no usamos la hoja Teoria", content)
        
        # Comment out the sub body
        content = re.sub(
            r"' Tabla de la ISA en la hoja Teoria.*?\nPublic Sub TeoriaEscribirISA\(\).*?End Sub",
            "' Tabla de la ISA en la hoja Teoria (Comentado porque eliminamos la hoja)\nPublic Sub TeoriaEscribirISA()\n    Exit Sub ' Omitido intencionalmente\nEnd Sub",
            content, flags=re.DOTALL
        )
    
    if fname == 'ModISA.bas':
        content = content.replace("' Numero de opcodes definidos (para la tabla de la hoja Teoria)", "' Número de opcodes disponibles en el Instruction Set")

    # Write back
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
print("Done final cleanup")
