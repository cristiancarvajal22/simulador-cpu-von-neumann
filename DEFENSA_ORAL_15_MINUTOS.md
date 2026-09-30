# 🛡️ Guía de Preparación para la Defensa Oral (Cheat Sheet)

Esta guía contiene los 3 puntos clave técnicos sobre la implementación en Excel (VBA) para responder con seguridad durante la evaluación en vivo.

## 1. Modificación de retardo (Animación) y forzado del C.P.

**Si el docente pide:** *"Modifica el código para que la animación vaya más rápido/lento, o fuerza un salto del Program Counter a la dirección 05h."*

*   **Para alterar el retardo (Sleep):**
    Debes ir al archivo VBA, módulo `ModAnimacion`, a la función `Pausa(ByVal segundosAlMaximo As Double)`. 
    Allí se utiliza `Espera segundosAlMaximo / Factor()`. Para hacerlo estático o cambiarlo, puedes modificar la constante `T_MIN_TRAMO` o directamente reemplazar la línea con `Application.Wait (Now + TimeValue("0:00:01"))` para forzar un segundo exacto.
*   **Para forzar el C.P.:**
    En el código (por ejemplo en `ModSecuenciador.bas` durante el Fetch), el C.P. se actualiza con `CpuSet R_PC, nuevo_valor`. Si te piden que el programa salte a la dirección `05h`, puedes inyectar `CpuSet R_PC, 5` justo antes de que termine la fase de Decode, o simplemente editar la celda de valor del C.P. en la hoja si el simulador permite interacción bidireccional.

## 2. Estructura matemática para actualizar las Banderas (ZF, CF, SF)

**Si el docente pide:** *"Explica cómo calculaste y actualizaste las banderas lógicas en tu ALU."*

En tu módulo de la ALU (Arithmetic Logic Unit), las banderas se actualizan tras cada operación matemática de la siguiente forma (ejemplo para suma en 8 bits):

*   **ZF (Zero Flag):** Se evalúa matemáticamente como `Si Resultado = 0 Entonces ZF = 1, sino ZF = 0`.
*   **CF (Carry Flag):** Dado que estamos limitados a 8 bits (0-255), tras sumar operando 1 + operando 2, verificamos si `(Op1 + Op2) > 255`. Si es mayor, hubo desbordamiento y `CF = 1`. Internamente se hace un `Resultado AND &HFF` para truncarlo a 8 bits.
*   **SF (Sign Flag):** En complemento a 2, el bit más significativo (el bit 7) indica el signo. Se evalúa con una máscara bit a bit: `Si (Resultado AND &H80) <> 0 Entonces SF = 1` (es decir, el bit 7 está en 1).

## 3. Flujo de datos desde la RAM hasta el RIM (Fase Fetch)

**Si el docente pide:** *"Muéstrame en el código cómo fluye exactamente el dato desde la memoria hasta el registro de datos (RIM) durante el Fetch."*

El flujo interno programado en VBA funciona en tres tiempos durante la **Fase de Búsqueda**:

1.  **Activación del MAR (RDM):** El valor del C.P. se asigna a la variable que representa el MAR: `MAR = PC`. Luego se invoca `EncenderRuta("CP_MAR")` para pintar el bus en la UI.
2.  **Lectura (Selector a Memoria):** Se lee el valor del arreglo/matriz bidimensional de la hoja que representa la memoria usando la dirección contenida en MAR. En VBA esto es algo como `DatoLeido = HojaMemoria.Cells(FilaOffset + MAR, ColumnaBase).Value`.
3.  **Llegada al MDR (RIM):** Ese valor leído se asigna al registro de datos: `MDR = DatoLeido`. Visualmente, la macro llama a `EncenderRuta("MEM_MDR")` y actualiza la celda/Shape del RIM con el nuevo dato en formato hexadecimal.
