# 🛡️ Guía de Preparación para la Defensa Oral (Cheat Sheet)

Esta guía contiene los 3 puntos clave técnicos sobre la implementación en Google Apps Script para responder con seguridad durante la evaluación en vivo.

## 1. Modificación de retardo (Animación) y forzado del C.P.

**Si el docente pide:** *"Modifica el código para que la animación vaya más rápido/lento, o fuerza un salto del Program Counter a la dirección 05h."*

*   **Para alterar el retardo:**
    Debes ir al archivo `Code.gs` a la función `ejecutar()`. Allí se utiliza `Utilities.sleep(leerPausa_())`. La función `leerPausa_()` lee el valor directamente de la celda `AQ4` en la hoja `Diagrama` (limitado entre 100 y 2000 ms). Para forzar un retardo de 1 segundo en el código, reemplaza `leerPausa_()` por `1000`: `Utilities.sleep(1000)`.
*   **Para forzar el C.P.:**
    El estado completo de la CPU se guarda en formato JSON en las propiedades del documento (`PropertiesService`). En la función `paso()` o durante el ciclo, si quieres forzar un salto, puedes inyectar `s.r.PC = 5` en el objeto de estado `s` antes de que se llame a `props.setProperty(STATE_KEY, JSON.stringify(s))`. También puedes editar la memoria directamente para poner una instrucción `JMP 05h` (opcode `30 05` en nuestra ISA) y dejar que el ciclo la ejecute.

## 2. Estructura matemática para actualizar las Banderas (ZF, CF, SF)

**Si el docente pide:** *"Explica cómo calculaste y actualizaste las banderas lógicas en tu ALU."*

En tu función de la ALU (`function alu(op, a, b)` en `Code.gs`), las banderas se actualizan tras cada operación matemática (ejemplo para suma en 8 bits):

*   **ZF (Zero Flag):** Se evalúa matemáticamente con un comparador booleano casteado a número: `ZF: +(result === 0)`.
*   **CF (Carry Flag):** Tras sumar `n = a + b`, verificamos si `n > 255`. Si es mayor, hubo desbordamiento y se activa la bandera con `CF = +(n > 255)`. El resultado se trunca a 8 bits con `n & 255`. Para restas, la condición es `CF = +(n < 0)`.
*   **SF (Sign Flag):** En complemento a 2, el bit más significativo (bit 7) indica el signo. Se evalúa aplicando una máscara bit a bit AND de 128 (10000000 en binario): `SF: +((result & 128) !== 0)`.

## 3. Flujo de datos desde la RAM hasta el RIM (Fase Fetch)

**Si el docente pide:** *"Muéstrame en el código cómo fluye exactamente el dato desde la memoria hasta el registro de datos (RIM) durante el Fetch."*

El flujo en la función `fetch(s)` y el motor de encolado de micro-operaciones funciona así:

1.  **Activación del MAR (RDM):** El motor encola la transferencia del C.P. al MAR usando `copy(s, 'PC', 'MAR', 'FETCH')`.
2.  **Lectura (Selector a Memoria):** Se invoca una primitiva de memoria `read(s, a)` que retorna el byte de la matriz interna de RAM: `return s.ram[byte(a)]`. Esta transferencia se encola para el RIM con `copy(s, 'RAM', 'MDR', 'FETCH')`.
3.  **Llegada al MDR (RIM):** La función `step(s)` procesa esa cola. Cuando procesa el paso `RAM` a `MDR`, lee el valor invocando a `value(s, 'RAM')` (que extrae `s.ram[s.r.MAR]`) y lo asigna a `s.r.MDR`. A nivel visual, `animar_(d, e)` calcula la trayectoria (`cablesDe_('RAM', 'MDR')`) y mueve un cursor visual recorriendo el bus desde la celda correspondiente de memoria hasta el registro `RIM`, animado temporalmente con llamadas a `Utilities.sleep(tiempo)`.
