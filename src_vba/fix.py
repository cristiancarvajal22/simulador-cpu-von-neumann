import os

path = r'c:\Users\crist\Downloads\Simulador CPU\Entregable_Parcial1\src_vba\ModInterfaz.bas'
with open(path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for line in lines:
    if line.startswith("' Tabla de la ISA en la hoja Teoria, generada desde ModISA"):
        new_lines.append("' Tabla de la ISA en la hoja Teoria (Comentada porque eliminamos la hoja)\n")
        new_lines.append("Public Sub TeoriaEscribirISA()\n")
        new_lines.append("    Exit Sub ' Omitido intencionalmente\n")
        new_lines.append("End Sub\n")
        skip = True
    elif skip and line.startswith("End Sub"):
        skip = False
    elif not skip:
        new_lines.append(line)

with open(path, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)
