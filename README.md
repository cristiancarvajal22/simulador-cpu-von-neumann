# Simulador de CPU 8-bits - Entregable Final

Proyecto correspondiente al Parcial 1 de Arquitectura de Computadoras. Este repositorio contiene el entregable final de un simulador de una CPU de 8 bits implementado íntegramente en Excel VBA, con un rediseño personalizado de la interfaz.

## Arquitectura (Arquitectura de Von Neumann)

`mermaid
graph TD
    subgraph CPU
        PC[Program Counter] --> MAR
        IR[Instruction Register] --> ControlUnit[Control Unit / Decoder]
        ControlUnit --> Sequencer[Sequencer]
        AX[Accumulator AX] --> ALU
        BX[Register BX] --> ALU
        ALU --> Flags[Status Flags: ZF, CF, SF, OF]
        ALU --> AX
        ALU --> BX
    end
    
    subgraph Memoria
        MAR[Memory Address Register] --> RAM[RAM 256x8 bits]
        RAM <--> MDR[Memory Data Register]
    end
    
    MDR --> IR
    MDR <--> AX
    MDR <--> BX
    
    %% Data buses
    MDR -.->|Internal Data Bus| CPU
`

## Tabla ISA (Instruction Set Architecture)

| Mnemónico | Hex | Descripción |
|---|---|---|
| HLT | 00 | Halt - Detiene la ejecución |
| MOV AX, imm | 10 | Mueve un valor inmediato a AX |
| MOV BX, imm | 11 | Mueve un valor inmediato a BX |
| LOAD AX, [dir] | 20 | Carga en AX el valor de memoria |
| STORE [dir], AX | 22 | Guarda AX en memoria |
| ADD AX, BX | 32 | Suma BX a AX (AX = AX + BX) |
| SUB AX, BX | 42 | Resta BX de AX (AX = AX - BX) |
| CMP AX, BX | 52 | Compara AX y BX (actualiza Flags) |
| INC AX | 60 | Incrementa AX en 1 |
| DEC BX | 61 | Decrementa BX en 1 |
| AND AX, BX | 72 | Operación lógica AND |
| OR AX, BX | 82 | Operación lógica OR |
| XOR AX, BX | 92 | Operación lógica XOR |
| NOT AX | 6A | Inversión de bits de AX |
| JMP dir | A0 | Salto incondicional |
| JZ dir | A1 | Salto si Zero (ZF=1) |
| JC dir | A2 | Salto si Carry (CF=1) |

## Características del Simulador
- **Diseño Inmersivo y Limpio**: UI completamente personalizada, estructurada y sin textos redundantes (Dark Mode / Terminal Style).
- 256 bytes de memoria RAM (00h - FFh).
- Ciclo de instrucción de 4 fases: Fetch, Decode, Execute, Store.
- Interfaz gráfica animada que muestra el flujo de datos y el encendido de buses.
- Log detallado tipo terminal mostrando las micro-operaciones.

## 📦 Estructura del Repositorio y Código Fuente
Para facilitar la revisión del código sin necesidad de abrir el binario de Excel, todos los módulos VBA del simulador han sido extraídos a la carpeta src_vba.

*(Nota de entrega: Todo el código fuente de los módulos y la reestructuración de la interfaz fue finalizado y subido al repositorio como entregable definitivo).*
