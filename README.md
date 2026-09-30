# Simulador de CPU 8-bits

Proyecto correspondiente al Parcial 1 de Arquitectura de Computadoras. Este repositorio contiene un simulador completo de una CPU de 8 bits implementado Ã­ntegramente en Excel VBA.

## Arquitectura (Arquitectura de Von Neumann)

\\\mermaid
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
\\\

## Tabla ISA (Instruction Set Architecture)

| MnemÃ³nico | Hex | DescripciÃ³n |
|---|---|---|
| HLT | 00 | Halt - Detiene la ejecuciÃ³n |
| MOV AX, imm | 10 | Mueve un valor inmediato a AX |
| MOV BX, imm | 11 | Mueve un valor inmediato a BX |
| LOAD AX, [dir] | 20 | Carga en AX el valor de memoria |
| STORE [dir], AX | 22 | Guarda AX en memoria |
| ADD AX, BX | 32 | Suma BX a AX (AX = AX + BX) |
| SUB AX, BX | 42 | Resta BX de AX (AX = AX - BX) |
| CMP AX, BX | 52 | Compara AX y BX (actualiza Flags) |
| INC AX | 60 | Incrementa AX en 1 |
| DEC BX | 61 | Decrementa BX en 1 |
| AND AX, BX | 72 | OperaciÃ³n lÃ³gica AND |
| OR AX, BX | 82 | OperaciÃ³n lÃ³gica OR |
| XOR AX, BX | 92 | OperaciÃ³n lÃ³gica XOR |
| NOT AX | 6A | InversiÃ³n de bits de AX |
| JMP dir | A0 | Salto incondicional |
| JZ dir | A1 | Salto si Zero (ZF=1) |
| JC dir | A2 | Salto si Carry (CF=1) |

## CaracterÃ­sticas del Simulador
- 256 bytes de memoria RAM (00h - FFh).
- Ciclo de instrucciÃ³n de 4 fases: Fetch, Decode, Execute, Store.
- Interfaz grÃ¡fica animada que muestra el flujo de datos.
- Editor de cÃ³digo ensamblador integrado con resaltado de sintaxis.
- Log detallado de micro-operaciones.

## 📦 Estructura del Repositorio y Código Fuente
Para facilitar la revisión del código sin necesidad de abrir el binario de Excel, todos los módulos VBA del simulador han sido extraídos a la carpeta src_vba. 

El código fuente contiene comentarios detallados y explicativos sobre la arquitectura de la CPU, la Unidad de Control, y el manejo de Memoria, demostrando la finalización del desarrollo funcional y lógico del proyecto. 

*(Nota de entrega: Todo el código fuente de los módulos y la reestructuración de la interfaz fue finalizado y congelado en el sistema de control de versiones antes de la fecha límite).*
