/** @OnlyCurrentDoc */
/* =====================================================================
 *  SIMULADOR DE CPU 8 bits (von Neumann) + MEMORIA · Google Sheets
 *  SIS131 – Parcial 1
 *  Nomenclatura: C.P.=PC · R.I.=IR · RDM=MAR · RIM=MDR · AC=salida de C.OP.
 *  Secciones: 1 Motor · 2 Ejemplos · 3 Configuración · 4 Acciones (botones)
 *             5 Pintado · 6 Animación · 7 Controles · 8 Preparar diagrama
 * ===================================================================== */

/* ================== 1. MOTOR DE LA CPU (independiente de la hoja) ================== */
var Cpu = (() => {
  const ISA = {};
  function def(code, op, dst, src, size) {
    ISA[code] = {code, op, dst, src, size,
      mnemonic: op === 'HLT' ? 'HLT' : op + ' ' + (dst || '') + (src ? ',' + src : '')};
  }
  // Variantes con AX
  def(1,'MOV','AX','imm',2); def(2,'MOV','BX','imm',2);
  def(3,'MOV','AX','BX',1);  def(4,'MOV','BX','AX',1);
  def(5,'LOAD','AX','[dir]',2); def(6,'LOAD','BX','[dir]',2);
  def(7,'STORE','[dir]','AX',2); def(8,'STORE','[dir]','BX',2);
  [['ADD',16],['SUB',18],['CMP',24],['AND',32],['OR',34],['XOR',36]].forEach(([op, n]) => {
    def(n, op, 'AX', 'imm', 2);
    def(n + 1, op, 'AX', 'BX', 1);
  });
  def(20,'INC','AX','',1); def(21,'INC','BX','',1);
  def(22,'DEC','AX','',1); def(23,'DEC','BX','',1);
  def(38,'NOT','AX','',1);
  def(48,'JMP','dir','',2); def(49,'JZ','dir','',2); def(50,'JNZ','dir','',2);
  def(255,'HLT','','',1);
  // Variantes con BX
  [['ADD',64],['SUB',66],['CMP',68],['AND',70],['OR',72],['XOR',74]].forEach(([op, n]) => {
    def(n, op, 'BX', 'imm', 2);
    def(n + 1, op, 'BX', 'AX', 1);
  });
  def(76,'NOT','BX','',1);

  const hex = n => n.toString(16).toUpperCase().padStart(2, '0');
  const bin = n => n.toString(2).padStart(8, '0');

  function byte(n) {
    if (!Number.isInteger(n) || n < 0 || n > 255) throw Error('Se requiere un byte entre 0 y 255.');
    return n;
  }
  // Primitivas de memoria: Read(address) / Write(address, value)
  function read(s, a) { return s.ram[byte(a)]; }
  function write(s, a, v) { s.ram[byte(a)] = byte(v); }
  function parseHex(v) {
    const t = String(v).trim();
    if (!/^[\da-f]{1,2}$/i.test(t)) throw Error('Byte hexadecimal inválido: ' + t);
    return parseInt(t, 16);
  }

  // Programa demo: multiplicación 5 x 3 por sumas sucesivas (resultado en RAM[82h] = 0Fh)
  function demo() {
    const ram = Array(256).fill(0);
    [5,129,24,0,49,25,6,128,5,130,17,7,130,5,129,22,7,129,24,0,50,8,48,25,0,255]
      .forEach((v, i) => ram[i] = v);
    ram[128] = 5; ram[129] = 3;
    return ram;
  }

  function create(ram) {
    if (!Array.isArray(ram) || ram.length !== 256) throw Error('La RAM debe contener 256 bytes.');
    return reset({ram: ram.map(byte)});
  }
  function reset(s) {
    s.r = Object.fromEntries(['PC','IR','OP','MAR','MDR','AX','BX','AC','REN1','REN2'].map(k => [k, 0]));
    s.flags = {ZF: 0, CF: 0, SF: 0};
    s.queue = []; s.index = 0; s.steps = 0;
    s.phase = 'LISTO'; s.instruction = 'Sin decodificar';
    s.instructionAddress = 0; s.hasOperand = false;
    s.halted = false; s.error = ''; s.last = null;
    return s;
  }

  // ALU: devuelve resultado y banderas. ZF = resultado 0 · SF = bit 7 · CF = acarreo/préstamo sin signo
  function alu(op, a, b) {
    byte(a); byte(b);
    let n = 0, c = 0;
    switch (op) {
      case 'ADD': n = a + b; c = +(n > 255); break;
      case 'SUB': case 'CMP': n = a - b; c = +(n < 0); break;
      case 'INC': n = a + 1; c = +(n > 255); break;
      case 'DEC': n = a - 1; c = +(n < 0); break;
      case 'AND': n = a & b; break;
      case 'OR': n = a | b; break;
      case 'XOR': n = a ^ b; break;
      case 'NOT': n = (~a) & 255; break;
      default: throw Error('Operación ALU desconocida.');
    }
    const result = n & 255;
    return {result, flags: {ZF: +(result === 0), CF: c, SF: +((result & 128) !== 0)}};
  }

  // Cola de micro-operaciones: cada elemento es UNA transferencia entre componentes
  function add(s, kind, src, dst, phase, label) {
    s.queue.push({kind, src, dst, phase, label: label || dst + ' ← ' + src});
  }
  function copy(s, src, dst, phase) { add(s, 'copy', src, dst, phase); }

  // FETCH: MAR←PC, MDR←RAM[MAR], IR←MDR, PC←PC+1  (luego DECODE)
  function fetch(s) {
    s.queue = []; s.index = 0;
    s.instructionAddress = s.r.PC;
    s.instruction = 'Sin decodificar'; s.hasOperand = false;
    copy(s, 'PC', 'MAR', 'FETCH');
    copy(s, 'RAM', 'MDR', 'FETCH');
    copy(s, 'MDR', 'IR', 'FETCH');
    add(s, 'inc', 'PC', 'PC', 'FETCH', 'PC ← PC + 1');
    add(s, 'decode', 'IR', 'DEC', 'DECODE', 'Interpretar IR y preparar microórdenes');
  }

  // DECODE: trae el operando (si lo hay) y encola las fases EXECUTE y STORE
  function decode(s) {
    const d = ISA[s.r.IR];
    s.r.OP = 0; s.hasOperand = false;
    if (!d) {
      s.halted = true;
      s.error = 'Opcode ' + hex(s.r.IR) + 'h inválido en ' + hex(s.instructionAddress) + 'h. Pulsa RESET.';
      return;
    }
    s.instruction = d.mnemonic;
    if (d.size === 2) {
      copy(s, 'PC', 'MAR', 'DECODE');
      copy(s, 'RAM', 'MDR', 'DECODE');
      copy(s, 'MDR', 'OP', 'DECODE');
      add(s, 'inc', 'PC', 'PC', 'DECODE', 'PC ← PC + 1 (operando)');
    }
    switch (d.op) {
      case 'MOV':
        copy(s, d.src === 'imm' ? 'OP' : d.src, 'AC', 'EXECUTE');
        copy(s, 'AC', d.dst, 'STORE');
        break;
      case 'LOAD':
        copy(s, 'OP', 'MAR', 'EXECUTE');
        copy(s, 'RAM', 'MDR', 'EXECUTE');
        copy(s, 'MDR', d.dst, 'STORE');
        break;
      case 'STORE':
        copy(s, 'OP', 'MAR', 'EXECUTE');
        copy(s, d.src, 'MDR', 'EXECUTE');
        copy(s, 'MDR', 'RAM', 'STORE');
        break;
      case 'JMP': case 'JZ': case 'JNZ':
        add(s, 'jump', 'OP', 'PC', 'EXECUTE', 'Evaluar ' + d.op);
        add(s, 'none', 'SEQ', 'SEQ', 'STORE', 'Salto completado; sin escritura de datos');
        break;
      case 'HLT':
        add(s, 'none', 'DEC', 'SEQ', 'EXECUTE', 'HLT ordena detener el reloj');
        add(s, 'halt', 'SEQ', 'CLOCK', 'STORE', 'HLT: procesador detenido');
        break;
      default: // ADD, SUB, AND, OR, XOR, CMP, INC, DEC, NOT
        copy(s, d.dst, 'REN1', 'EXECUTE');
        copy(s, ['INC','DEC'].includes(d.op) ? 'ONE' : d.op === 'NOT' ? 'ZERO' : d.src === 'imm' ? 'OP' : d.src, 'REN2', 'EXECUTE');
        add(s, 'alu', 'ALU', 'AC', 'EXECUTE', 'AC ← ' + d.op + '(REN1, REN2); actualizar banderas');
        if (d.op === 'CMP') add(s, 'none', 'FLAGS', 'FLAGS', 'STORE', 'CMP conserva AX y BX');
        else copy(s, 'AC', d.dst, 'STORE');
    }
  }

  function value(s, k) {
    return k === 'RAM' ? read(s, s.r.MAR) : k === 'ONE' ? 1 : k === 'ZERO' ? 0 : (s.r[k] ?? 0);
  }

  // Ejecuta UNA micro-operación y devuelve el evento (para animar y registrar)
  function step(s) {
    if (s.halted) return null;
    if (s.index >= s.queue.length) fetch(s);
    const q = s.queue[s.index];
    const before = {r: {...s.r}, flags: {...s.flags}, instruction: s.instruction};
    let v = value(s, q.src), from = q.src, to = q.dst, label = q.label;
    const address = s.r.MAR;
    s.phase = q.phase;
    switch (q.kind) {
      case 'copy':
        if (q.dst === 'RAM') write(s, s.r.MAR, v); else s.r[q.dst] = v;
        if (q.dst === 'OP') s.hasOperand = true;
        break;
      case 'inc':
        s.r.PC = (s.r.PC + 1) & 255; v = s.r.PC;
        break;
      case 'decode':
        decode(s);
        label = s.error || 'IR = ' + hex(s.r.IR) + 'h → ' + s.instruction;
        break;
      case 'alu': {
        const r = alu(ISA[s.r.IR].op, s.r.REN1, s.r.REN2);
        s.r.AC = r.result; s.flags = r.flags; v = r.result;
        break;
      }
      case 'jump': {
        const op = ISA[s.r.IR].op;
        const taken = op === 'JMP' || (op === 'JZ' && s.flags.ZF === 1) || (op === 'JNZ' && s.flags.ZF === 0);
        if (taken) s.r.PC = s.r.OP; else { from = 'FLAGS'; to = 'SEQ'; }
        v = taken ? s.r.OP : s.flags.ZF;
        label = op + (taken ? ' tomado' : ' no tomado') + '; PC = ' + hex(s.r.PC) + 'h';
        break;
      }
      case 'halt':
        s.halted = true;
        break;
    }
    s.index++; s.steps++;
    const event = {step: s.steps, phase: s.phase, from, to, value: v, label, address, before};
    s.last = {...event}; delete s.last.before;
    return event;
  }

  // ---- Utilidades de presentación (assembly legible) ----
  function formato(d, operand) {
    const h = n => hex(n) + 'h';
    switch (d.op) {
      case 'HLT': return 'HLT';
      case 'INC': case 'DEC': case 'NOT': return d.op + ' ' + d.dst;
      case 'JMP': case 'JZ': case 'JNZ': return d.op + ' ' + h(operand);
      case 'LOAD': return 'LOAD ' + d.dst + ',[' + h(operand) + ']';
      case 'STORE': return 'STORE [' + h(operand) + '],' + d.src;
      default: return d.op + ' ' + d.dst + ',' + (d.src === 'imm' ? h(operand) : d.src);
    }
  }
  function texto(s) {
    const d = ISA[s.r.IR];
    if (!d || s.instruction === 'Sin decodificar') return 'Sin decodificar';
    return s.hasOperand ? formato(d, s.r.OP) : d.mnemonic;
  }
  // Desensambla el segmento de código (00h–7Fh) hasta el primer HLT
  function desensamblar(ram) {
    const out = Array(256).fill('');
    let a = 0;
    while (a < 128) {
      const d = ISA[ram[a]];
      if (!d) break;
      out[a] = formato(d, d.size === 2 ? ram[(a + 1) & 255] : 0);
      if (d.size === 2) out[(a + 1) & 255] = '   ↳ operando';
      if (d.op === 'HLT') break;
      a += d.size;
    }
    return out;
  }
  function notas(ram) {
    const asm = desensamblar(ram);
    return Array.from({length: 16}, (_, r) => Array.from({length: 16}, (_, c) => {
      const a = r * 16 + c, v = ram[a];
      return 'Dir ' + hex(a) + 'h · ' + (a < 128 ? 'CÓDIGO' : 'DATOS') + '\nHex ' + hex(v) + 'h\nBin ' + bin(v) +
        '\nDec ' + v + (asm[a] ? '\nAsm ' + asm[a] : '');
    }));
  }
  return {ISA, hex, bin, byte, parseHex, read, write, create, reset, demo, alu, step, formato, texto, desensamblar, notas};
})();
if (typeof module !== 'undefined') module.exports = Cpu;

/* ================== 2. EJEMPLOS: un programa pequeño por instrucción ================== */
var Ejemplos = (() => {
  function crear(code, a, b, inmediato) {
    const d = Cpu.ISA[code];
    if (!d) throw Error('Selecciona una instrucción válida.');
    [a, b, inmediato].forEach(Cpu.byte);
    const ram = Array(256).fill(0);
    ram[128] = a; ram[129] = b;
    const bytes = [5, 128, 6, 129];                         // LOAD AX,[80h] · LOAD BX,[81h]
    const notas = {0: 'LOAD AX,[80h] · leer dato A', 2: 'LOAD BX,[81h] · leer dato B'};
    if (['JMP', 'JZ', 'JNZ'].includes(d.op)) {
      notas[4] = 'CMP AX,BX · preparar condición'; bytes.push(25);
      notas[5] = d.op + ' 09h · saltar a almacenamiento'; bytes.push(code, 9);
      notas[7] = 'MOV AX,EEh · solo si no se toma el salto'; bytes.push(1, 238);
    } else {
      notas[bytes.length] = d.mnemonic;
      bytes.push(code);
      if (d.size === 2) bytes.push(d.op === 'LOAD' ? 128 : d.op === 'STORE' ? 132 : inmediato);
    }
    if (d.op !== 'HLT') {
      notas[bytes.length] = 'STORE [82h],AX · conservar AX final'; bytes.push(7, 130);
      notas[bytes.length] = 'STORE [83h],BX · conservar BX final'; bytes.push(8, 131);
      notas[bytes.length] = 'HLT'; bytes.push(255);
    }
    bytes.forEach((v, i) => ram[i] = v);
    notas[128] = 'Dato A'; notas[129] = 'Dato B'; notas[130] = 'AX final'; notas[131] = 'BX final'; notas[132] = 'Destino del ejemplo STORE';
    return {ram, notas};
  }
  return {crear};
})();

/* ================== 3. CONFIGURACIÓN DE LA HOJA ================== */
const HOJA = {DIAGRAMA: 'Diagrama', PROGRAMA: 'Programa', MEMORIA: 'Memoria', REGISTRO: 'Registro', OPERACION: 'Operación', ESTADO: 'Estado'};
const STATE_KEY = 'cpu8.native.v1', RUN_KEY = 'cpu8.run';
const ANCHO_COL = 20, ALTO_FILA = 16;                       // píxeles por celda del diagrama
const REGIONES = {
  PC: 'U20:W22', IR: 'Q16:W18', OP: 'Q20:S22', MAR: 'V32:AD34',
  MDR: 'AK32:AS34', AX: 'AD21:AF23', BX: 'AH21:AJ23', AC: 'AD10:AJ12',
  REN1: 'AP20:AS22', REN2: 'AU20:AX22', FLAGS: 'AD16:AJ18',
  DEC: 'Q10:W12', SEQ: 'I10:N15', CLOCK: 'C10:F12', ALU: 'AQ12:AU14'
};
const REG_VALOR = ['PC', 'IR', 'OP', 'MAR', 'MDR', 'AX', 'BX', 'AC', 'REN1', 'REN2'];
const LABEL_RANGO = {REN1: 'AP19:AQ19', REN2: 'AW19:AX19', MAR: 'X31:AD31', MDR: 'AN31:AS31'};  // dejan libres los cables
const C = {fondo: '#808080', uc: '#D8EDF3', alu: '#ECF1DC', mem: '#FFEAD4', valor: '#C65A49', activo: '#FFD166',
           busDir: '#2B2B2B', busDat: '#F5F5F5', texto: '#17212B'};
const VENTANA_FILA0 = 39;                                   // primera fila de la ventana de memoria del diagrama

// Buses dibujados con celdas. Bus de Direcciones (oscuro) y Bus de Datos (claro)
const BUS_DIR = ['Q23:W23', 'W24:W31', 'W35:W36', 'AC39:AE39'];
const BUS_DAT = ['X17:Y17', 'Z17:Z25', 'Z26:AV26', 'AC11:AC25', 'AE24:AE25', 'AI24:AI25', 'AR23:AR25', 'AV23:AV26',
                 'AL27:AL31', 'AR35:AR46', 'AK12:AP12', 'AK17:AQ17', 'AR15:AR19', 'AU15:AU19', 'T13:T14', 'T21'];
// Recorrido de cada transferencia (en orden origen → destino). La inversa se deduce sola.
const CABLES = {
  'PC>MAR': ['U23:W23', 'W24:W31'],
  'OP>MAR': ['Q23:W23', 'W24:W31'],
  'RAM>MDR': ['AR35:AR46'],
  'MDR>IR': ['AL27:AL31', 'Z26:AL26', 'Z17:Z25', 'X17:Y17'],
  'MDR>OP': ['AL27:AL31', 'Z26:AL26', 'Z17:Z25', 'X17:Y17'],
  'MDR>AX': ['AL27:AL31', 'AE26:AL26', 'AE24:AE25'],
  'MDR>BX': ['AL27:AL31', 'AI26:AL26', 'AI24:AI25'],
  'MDR>AC': ['AL27:AL31', 'AC26:AL26', 'AC11:AC25'],
  'OP>AC': ['X17:Y17', 'Z17:Z25', 'Z26:AC26', 'AC11:AC25'],
  'AX>AC': ['AE24:AE25', 'AC26:AE26', 'AC11:AC25'],
  'BX>AC': ['AI24:AI25', 'AC26:AI26', 'AC11:AC25'],
  'AX>REN1': ['AE24:AE25', 'AE26:AR26', 'AR23:AR25'],
  'BX>REN1': ['AI24:AI25', 'AI26:AR26', 'AR23:AR25'],
  'AX>REN2': ['AE24:AE25', 'AE26:AV26', 'AV23:AV25'],
  'BX>REN2': ['AI24:AI25', 'AI26:AV26', 'AV23:AV25'],
  'OP>REN2': ['X17:Y17', 'Z17:Z25', 'Z26:AV26', 'AV23:AV25'],
  'ONE>REN2': ['AV23:AV26'],
  'ZERO>REN2': ['AV23:AV26'],
  'REN1>ALU': ['AR15:AR19'],
  'REN2>ALU': ['AU15:AU19'],
  'ALU>AC': ['AK12:AP12', 'AK17:AQ17'],
  'IR>DEC': ['T13:T14'],
  'OP>PC': ['T21']
};
const SELECCION = ['W35:W36', 'AC39:AE39'];                 // MAR → Selector → Memoria (solo resaltado)

// Hoja oculta "Estado": el diagrama lee de ella con fórmulas, así cada paso escribe UNA sola vez.
const KEYS = ['PC', 'IR', 'OP', 'MAR', 'MDR', 'AX', 'BX', 'AC', 'REN1', 'REN2', 'FLAGS', 'DEC', 'SEQ', 'CLOCK', 'ALU',
  'ENTRADAS', 'FASE', 'INSTR', 'LISTA', 'MENSAJE', 'ESTADO', 'SEL']
  .concat(Array.from({length: 8}, (_, i) => 'MA' + i), Array.from({length: 8}, (_, i) => 'MV' + i));
const FILA = {}; KEYS.forEach((k, i) => FILA[k] = i + 1);
const DESTINO = (() => {
  const d = {};
  Object.keys(REGIONES).forEach(k => d[k] = REGIONES[k].split(':')[0]);
  Object.assign(d, {ENTRADAS: 'B25', FASE: 'B27', MENSAJE: 'B30', INSTR: 'B35', LISTA: 'B39', ESTADO: 'B50', SEL: 'V38'});
  for (let i = 0; i < 8; i++) { d['MA' + i] = 'AF' + (39 + i); d['MV' + i] = 'AI' + (39 + i); }
  return d;
})();

/* ---- Utilidades de coordenadas (sin llamadas a la API) ---- */
function col_(letras) { let n = 0; for (const ch of letras) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function letras_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function rango_(a1) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a1);
  const c1 = col_(m[1]), r1 = +m[2];
  return {c1, r1, c2: m[3] ? col_(m[3]) : c1, r2: m[4] ? +m[4] : r1};
}
function centroPx_(a1) {
  const g = rango_(a1);
  return [((g.c1 - 1) + g.c2) / 2 * ANCHO_COL, ((g.r1 - 1) + g.r2) / 2 * ALTO_FILA];
}
function inicioVentana_(mar) { return Math.min(248, Math.max(0, mar - 3)); }

/* ================== 4. ACCIONES (botones y menú) ================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Simulador CPU')
    .addItem('Elegir operación y datos', 'elegirOperacion').addItem('Cargar operación elegida', 'cargarOperacion')
    .addItem('Cargar demo: multiplicación', 'cargarDemo').addSeparator()
    .addItem('PASO · una microoperación', 'paso').addItem('EJECUTAR', 'ejecutar')
    .addItem('PAUSAR', 'pausar').addItem('RESET · conservar RAM', 'reiniciar')
    .addItem('CARGAR · bytes de Programa', 'cargarPrograma').addSeparator()
    .addItem('Leer memoria', 'leerMemoria').addItem('Escribir memoria', 'escribirMemoria')
    .addSeparator().addItem('Preparar diagrama', 'prepararDiagrama').addItem('Actualizar vista', 'actualizarVista').addToUi();
}
function conBloqueo_(callback, espera) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(espera || 30000);
  try { return callback(PropertiesService.getDocumentProperties()); } finally { lock.releaseLock(); }
}
function leerPrograma_() {
  return SpreadsheetApp.getActive().getSheetByName(HOJA.PROGRAMA).getRange('B7:B262').getDisplayValues().map((row, i) => {
    try { return Cpu.parseHex(row[0]); } catch (error) { throw Error('Programa!B' + (i + 7) + ': ' + error.message); }
  });
}
function sesion_(props) {
  const raw = props.getProperty(STATE_KEY);
  return raw ? JSON.parse(raw) : Cpu.create(leerPrograma_());
}
function leerPausa_() {
  const v = Number(SpreadsheetApp.getActive().getSheetByName(HOJA.DIAGRAMA).getRange('AQ4').getValue());
  return Math.max(100, Math.min(2000, Number.isFinite(v) ? v : 500));
}

// PASO: una micro-operación
function paso() {
  conBloqueo_(props => {
    if (props.getProperty(RUN_KEY)) throw Error('Pulsa PAUSAR antes de avanzar manualmente.');
    const s = sesion_(props), e = Cpu.step(s);
    pintar_(s, e, false);
    props.setProperty(STATE_KEY, JSON.stringify(s));
  });
}
// EJECUTAR: micro-operaciones seguidas con pausa ajustable (Utilities.sleep)
function ejecutar() {
  const token = Utilities.getUuid();
  conBloqueo_(props => {
    if (props.getProperty(RUN_KEY)) throw Error('Ya está en ejecución.');
    props.setProperty(RUN_KEY, token);
  });
  const inicio = Date.now();
  try {
    while (Date.now() - inicio < 240000) {
      const terminado = conBloqueo_(props => {
        if (props.getProperty(RUN_KEY) !== token) return true;
        const s = sesion_(props), e = Cpu.step(s);
        pintar_(s, e, false);
        props.setProperty(STATE_KEY, JSON.stringify(s));
        return s.halted;
      });
      if (terminado) return;
      Utilities.sleep(leerPausa_());                        // [DEFENSA-1] retardo entre pasos
    }
    SpreadsheetApp.getActive().toast('Pausa automática tras 4 minutos. Pulsa EJECUTAR para continuar.');
  } finally {
    conBloqueo_(props => { if (props.getProperty(RUN_KEY) === token) props.deleteProperty(RUN_KEY); });
  }
}
function pausar() {
  conBloqueo_(props => props.deleteProperty(RUN_KEY));
  SpreadsheetApp.getActive().toast('Pausado al terminar el paso actual.');
}
function reiniciar() { sustituir_(false); }
function cargarPrograma() { sustituir_(true); }
function sustituir_(cargar) {
  conBloqueo_(props => {
    const s = cargar ? Cpu.create(leerPrograma_()) : Cpu.reset(sesion_(props));
    props.deleteProperty(RUN_KEY);
    pintar_(s, null, true);
    props.setProperty(STATE_KEY, JSON.stringify(s));
  });
}
function leerMemoria() {
  conBloqueo_(props => {
    const sh = SpreadsheetApp.getActive().getSheetByName(HOJA.MEMORIA);
    const a = Cpu.parseHex(sh.getRange('C25').getDisplayValue()), v = Cpu.read(sesion_(props), a);
    sh.getRange('F25').setNumberFormat('@').setValue(Cpu.hex(v));
    sh.getRange('I25').setValue(v);
    sh.getRange('L25').setNumberFormat('@').setValue(Cpu.bin(v));
    sh.getRange('O25').setValue(Cpu.ISA[v] ? Cpu.ISA[v].mnemonic : '(dato)');
  });
}
function escribirMemoria() {
  conBloqueo_(props => {
    if (props.getProperty(RUN_KEY)) throw Error('Pausa antes de editar la RAM.');
    const sh = SpreadsheetApp.getActive().getSheetByName(HOJA.MEMORIA);
    const a = Cpu.parseHex(sh.getRange('C25').getDisplayValue()), v = Cpu.parseHex(sh.getRange('F25').getDisplayValue());
    const s = sesion_(props);
    Cpu.write(s, a, v);
    pintar_(s, {step: s.steps, phase: 'EDICIÓN', from: 'USUARIO', to: 'RAM', value: v, address: a,
                label: 'RAM[' + Cpu.hex(a) + 'h] ← ' + Cpu.hex(v) + 'h'}, false);
    props.setProperty(STATE_KEY, JSON.stringify(s));
  });
}
function elegirOperacion() { SpreadsheetApp.getActive().getSheetByName(HOJA.OPERACION).activate(); }
function cargarOperacion() {
  conBloqueo_(props => {
    const libro = SpreadsheetApp.getActive(), sh = libro.getSheetByName(HOJA.OPERACION);
    const code = Cpu.parseHex(sh.getRange('C5').getDisplayValue().split(' · ')[0]);
    const valores = sh.getRange('C7:C9').getValues().flat().map(v => Cpu.byte(Number(v)));
    instalarPrograma_(props, Ejemplos.crear(code, ...valores));
    libro.getSheetByName(HOJA.DIAGRAMA).activate();
  });
}
function cargarDemo() {
  conBloqueo_(props => {
    instalarPrograma_(props, {ram: Cpu.demo(), notas: {128: 'Multiplicando: 5', 129: 'Contador: 3', 130: 'Resultado'}});
    SpreadsheetApp.getActive().getSheetByName(HOJA.DIAGRAMA).activate();
  });
}
function instalarPrograma_(props, programa) {
  const libro = SpreadsheetApp.getActive();
  libro.getSheetByName(HOJA.PROGRAMA).getRange('B7:C262').setNumberFormat('@')
    .setValues(programa.ram.map((v, i) => [Cpu.hex(v), programa.notas[i] || '']));
  const s = Cpu.create(programa.ram);
  props.deleteProperty(RUN_KEY);
  pintar_(s, null, true);
  props.setProperty(STATE_KEY, JSON.stringify(s));
}
// Figuras de la hoja que aún tienen una función asignada (solo muestran ayuda)
function mostrarALU() { SpreadsheetApp.getActive().toast('La ALU opera con REN1 y REN2; deposita el resultado en AC y actualiza ZF, CF y SF.'); }
function mostrarBuses() { SpreadsheetApp.getActive().toast('Bus de Direcciones (oscuro) y Bus de Datos (claro).'); }
function mostrarTransferencia() { SpreadsheetApp.getActive().toast('El cuadrado rojo recorre la conexión de la microoperación actual.'); }

/* ================== 5. PINTADO DEL DIAGRAMA (pocas llamadas a la API) ================== */
function asegurarEstado_() {
  const libro = SpreadsheetApp.getActive();
  let est = libro.getSheetByName(HOJA.ESTADO);
  if (!est) {
    const previa = libro.getActiveSheet();
    est = libro.insertSheet(HOJA.ESTADO);
    est.getRange(1, 1, KEYS.length, 1).setValues(KEYS.map(k => [k]));
    est.getRange(1, 2, KEYS.length, 1).setNumberFormat('@');
    libro.setActiveSheet(previa);
    est.hideSheet();
  }
  return est;
}
// El diagrama muestra fórmulas =Estado!Bn (se crean una vez)
function enlazar_(sh) {
  KEYS.forEach(k => sh.getRange(DESTINO[k]).setFormula('=' + HOJA.ESTADO + '!B' + FILA[k]));
}
function construirVista_(s, e) {
  const hex = Cpu.hex, v = {};
  Object.keys(s.r).forEach(k => v[k] = hex(s.r[k]) + 'h');
  v.FLAGS = 'ZF ' + s.flags.ZF + '   CF ' + s.flags.CF + '   SF ' + s.flags.SF;
  v.DEC = s.instruction === 'Sin decodificar' ? '—' : s.instruction.split(' ')[0];
  v.SEQ = s.steps + '\n' + s.phase;
  v.CLOCK = s.halted ? '■' : (s.steps % 2 ? '●' : '○');
  const d = Cpu.ISA[s.r.IR], op = d && d.op;
  const hayALU = s.instruction !== 'Sin decodificar' && ['ADD', 'SUB', 'INC', 'DEC', 'AND', 'OR', 'XOR', 'NOT', 'CMP'].includes(op);
  v.ALU = hayALU ? op + '\n' + s.r.REN1 + (op === 'NOT' ? '' : ' , ' + s.r.REN2) + '\nAC = ' + s.r.AC : 'Sin cálculo';
  v.ENTRADAS = 'Entradas ALU: REN1 = ' + s.r.REN1 + '   REN2 = ' + s.r.REN2 + '       AC = ' + s.r.AC + '   (valores decimales)';
  v.FASE = s.error ? 'ERROR' : s.halted ? 'DETENIDO · HLT' : s.phase;
  v.MENSAJE = s.error || (e ? e.label : 'Pulsa PASO para comenzar.');
  v.INSTR = Cpu.texto(s) + (s.hasOperand ? '\nOperando: ' + hex(s.r.OP) + 'h' : '');
  v.LISTA = s.queue.map((q, i) => (i === s.index - 1 ? '▶ ' : '   ') + (i + 1) + '. ' + q.label)
    .slice(Math.max(0, s.index - 4), s.index + 2).join('\n');
  v.ESTADO = e ? 'Paso ' + e.step + ': ' + e.from + ' → ' + e.to + '     Valor: ' + hex(e.value) + 'h / ' + e.value + ' / ' + Cpu.bin(e.value)
               : 'PASO realiza una microoperación. RESET conserva RAM; CARGAR restaura Programa.';
  v.SEL = hex(s.r.MAR) + 'h';
  const ini = inicioVentana_(s.r.MAR);
  for (let i = 0; i < 8; i++) {                              // ventana de memoria: "03", "0A", ... siempre con dos dígitos
    v['MA' + i] = hex(ini + i);
    v['MV' + i] = hex(s.ram[ini + i]) + 'h    (' + s.ram[ini + i] + ')';
  }
  return v;
}
function volcarEstado_(v) {
  asegurarEstado_().getRange(1, 2, KEYS.length, 1).setValues(KEYS.map(k => [v[k] === undefined ? '' : String(v[k])]));
}
function cablesDe_(from, to) {
  return CABLES[from + '>' + to] || (CABLES[to + '>' + from] || []).slice().reverse();
}
function colorear_(d, s, e) {
  d.getRangeList(REG_VALOR.map(k => REGIONES[k])).setBackground(C.valor).setFontColor('#FFFFFF');
  d.getRangeList(['DEC', 'SEQ', 'CLOCK'].map(k => REGIONES[k])).setBackground(C.uc).setFontColor(C.texto);
  d.getRangeList(['FLAGS', 'ALU'].map(k => REGIONES[k])).setBackground(C.alu).setFontColor(C.texto);
  d.getRangeList(BUS_DIR).setBackground(C.busDir);
  d.getRangeList(BUS_DAT).setBackground(C.busDat);
  d.getRange('AF39:AQ46').setBackground(C.mem);
  const fila = VENTANA_FILA0 + (s.r.MAR - inicioVentana_(s.r.MAR));
  d.getRange('AF' + fila + ':AQ' + fila).setBackground(C.activo);
  if (!e) return;
  const cables = cablesDe_(e.from, e.to).concat((e.from === 'RAM' || e.to === 'RAM') ? SELECCION : []);
  if (cables.length) d.getRangeList(cables).setBackground(C.activo);
  const activos = [e.from, e.to].filter(k => REGIONES[k]).map(k => REGIONES[k]);
  if (activos.length) d.getRangeList(activos).setBackground(C.activo).setFontColor(C.texto);
}
function actualizarMemoria_(libro, s, cambiaRam) {
  const mem = libro.getSheetByName(HOJA.MEMORIA), prog = libro.getSheetByName(HOJA.PROGRAMA);
  const matriz = mem.getRange('C6:R21');
  if (cambiaRam) {
    matriz.setNumberFormat('@').setValues(Array.from({length: 16}, (_, r) => s.ram.slice(r * 16, r * 16 + 16).map(Cpu.hex)));
    matriz.setNotes(Cpu.notas(s.ram));                       // hex · binario · decimal · mnemónico por dirección
    prog.getRange('D7:D262').setValues(Cpu.desensamblar(s.ram).map(t => [t]));   // assembly línea por línea
  }
  matriz.setBackgrounds(Array.from({length: 16}, (_, r) => Array.from({length: 16}, (_, c) =>
    r * 16 + c === s.r.MAR ? C.activo : r < 8 ? C.uc : C.alu)));                 // código 00h–7Fh / datos 80h–FFh
  mem.getRange('U6:X15').setValues(Object.entries(s.r).map(([k, v]) => [k, Cpu.hex(v), v, Cpu.bin(v)]));
  // Línea de assembly que se está ejecutando
  prog.getRange('B7:D262').setBackground(null);
  const d = Cpu.ISA[s.r.IR];
  if (s.steps > 0) {
    const n = (s.phase === 'FETCH' || !d) ? 1 : d.size;
    prog.getRange(7 + s.instructionAddress, 2, Math.min(n, 256 - s.instructionAddress), 3).setBackground(C.activo);
  }
}
function registrar_(libro, s, e, limpiar) {
  const log = libro.getSheetByName(HOJA.REGISTRO), h = Cpu.hex;
  if (limpiar && log.getLastRow() >= 5) log.getRange(5, 1, log.getLastRow() - 4, 10).clearContent();
  if (!e) return;
  const texto = '[Paso ' + String(e.step).padStart(2, '0') + '] ' + e.phase + ': ' + e.label +
    ' | MAR=0x' + h(s.r.MAR) + ', MDR=0x' + h(s.r.MDR) + ', IR=' + Cpu.texto(s);
  log.appendRow([e.step, e.phase, texto, h(e.value) + 'h', ...['PC', 'IR', 'MAR', 'MDR', 'AX', 'BX'].map(k => h(s.r[k]) + 'h')]);
}
function estacionarFicha_(d) {
  const ficha = d.getDrawings().find(x => x.getOnAction() === 'mostrarTransferencia');
  if (ficha) ficha.setPosition(51, 49, 0, 0);
}
function pintar_(s, e, limpiarLog) {
  const libro = SpreadsheetApp.getActive(), d = libro.getSheetByName(HOJA.DIAGRAMA);
  asegurarEstado_();
  if (!d.getRange(DESTINO.PC).getFormula()) enlazar_(d);     // primera vez: vincular diagrama y hoja Estado
  if (e && e.before) animar_(d, e);                          // el dato viaja ANTES de mostrar el resultado
  volcarEstado_(construirVista_(s, e));
  colorear_(d, s, e);
  actualizarMemoria_(libro, s, !e || e.to === 'RAM' || e.phase === 'EDICIÓN');
  registrar_(libro, s, e, limpiarLog);
  if (!e) estacionarFicha_(d);
  SpreadsheetApp.flush();
}

/* ================== 6. ANIMACIÓN: cuadrado rojo sobre los buses ================== */
function transferenciasVisuales_(e) {
  if (e.from !== 'ALU') return [e];
  const op = Cpu.ISA[e.before.r.IR].op;
  const rutas = [{from: 'REN1', to: 'ALU', value: e.before.r.REN1, label: 'Entrada 1 de ' + op + ' llega a la ALU'}];
  if (op !== 'NOT') rutas.push({from: 'REN2', to: 'ALU', value: e.before.r.REN2, label: 'Entrada 2 de ' + op + ' llega a la ALU'});
  rutas.push(e);
  return rutas;
}
function busNombre_(t) {
  if (t.to === 'MAR') return 'BUS DE DIRECCIONES';
  if (t.from === 'RAM' || t.to === 'RAM') return 'BUS DE DATOS (memoria)';
  return cablesDe_(t.from, t.to).length ? 'BUS DE DATOS' : 'TRANSFERENCIA INTERNA';
}
function puntoDe_(k, e) {
  if (k === 'RAM') {
    const fila = VENTANA_FILA0 + (e.address - inicioVentana_(e.address));
    return centroPx_('AF' + fila + ':AQ' + fila);
  }
  if (k === 'ONE' || k === 'ZERO') return centroPx_('AN27:AX28');
  return REGIONES[k] ? centroPx_(REGIONES[k]) : null;
}
function animar_(d, e) {
  const pausa = leerPausa_();
  const ficha = pausa >= 250 ? d.getDrawings().find(x => x.getOnAction() === 'mostrarTransferencia') : null;  // <250 ms: modo rápido, sin ficha
  const est = asegurarEstado_();
  for (const t of transferenciasVisuales_(e)) {
    const cables = cablesDe_(t.from, t.to).concat((t.from === 'RAM' || t.to === 'RAM') ? SELECCION : []);
    const activos = [t.from, t.to].filter(k => REGIONES[k]).map(k => REGIONES[k]);
    est.getRange(FILA.MENSAJE, 2, 2, 1).setValues([[t.label],
      ['EN TRANSFERENCIA · ' + busNombre_(t) + ' · ' + t.from + ' → ' + t.to + '     ' + Cpu.hex(t.value) + 'h = ' + t.value + ' = ' + Cpu.bin(t.value)]]);
    if (cables.length) d.getRangeList(cables).setBackground(C.activo);
    if (activos.length) d.getRangeList(activos).setBackground(C.activo).setFontColor(C.texto);
    if (!ficha) continue;
    // Recorrido: origen → centro de cada tramo de cable → destino
    const puntos = [puntoDe_(t.from, e)].concat(cablesDe_(t.from, t.to).map(centroPx_), [puntoDe_(t.to, e)]).filter(Boolean);
    const tiempo = Math.max(30, Math.min(140, Math.round(pausa * 0.6 / Math.max(1, puntos.length - 1))));
    for (let i = 0; i < puntos.length; i++) {
      const x = puntos[i][0] - 10, y = puntos[i][1] - 10;
      ficha.setPosition(Math.floor(y / ALTO_FILA) + 1, Math.floor(x / ANCHO_COL) + 1, Math.max(0, Math.round(x % ANCHO_COL)), Math.max(0, Math.round(y % ALTO_FILA)));
      SpreadsheetApp.flush(); Utilities.sleep(tiempo);
    }
  }
}
function actualizarVista() {
  const d = SpreadsheetApp.getActive().getSheetByName(HOJA.DIAGRAMA);
  asegurarEstado_(); enlazar_(d);
  const drawings = d.getDrawings(), ficha = drawings.find(x => x.getOnAction() === 'mostrarTransferencia');
  if (ficha) ficha.setWidth(20).setHeight(20).setZIndex(Math.max(...drawings.map(x => x.getZIndex())) + 1);
  conBloqueo_(props => { props.deleteProperty(RUN_KEY); pintar_(sesion_(props), null, false); });
}

/* ================== 7. HOJA "OPERACIÓN" (lista desplegable de instrucciones) ================== */
function prepararControles() {
  const libro = SpreadsheetApp.getActive(); libro.setSpreadsheetTimeZone('America/La_Paz');
  const sh = libro.getSheetByName(HOJA.OPERACION) || libro.insertSheet(HOJA.OPERACION);
  sh.setHiddenGridlines(true).setColumnWidths(1, 1, 22).setColumnWidths(2, 1, 220).setColumnWidths(3, 1, 245).setColumnWidths(4, 3, 100);
  sh.getRange('A1:F23').setFontFamily('Arial').setFontSize(12).setBackground('#F3F6FA').setVerticalAlignment('middle');
  sh.setRowHeights(1, 23, 30);
  const titulo = (r, t) => sh.getRange(r).merge().setValue(t).setWrap(true);
  titulo('B2:F3', 'ELIGE UNA INSTRUCCIÓN · DATOS DE 8 BITS');
  sh.getRange('B2:F3').setBackground('#16324F').setFontColor('#FFFFFF').setFontSize(18);
  sh.getRange('B5').setValue('Instrucción'); sh.getRange('C5:F5').merge().setValue('11 · ADD AX,BX');
  sh.getRange('C5').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(Cpu.ISA).map(d => Cpu.hex(d.code) + ' · ' + d.mnemonic), true).setAllowInvalid(false).build());
  sh.getRange('B7:C9').setValues([['Dato A (RAM 80h → AX)', 5], ['Dato B (RAM 81h → BX)', 3], ['Inmediato (si dice imm)', 3]]);
  sh.getRange('C7:C9').setNumberFormat('0').setBackground('#FFF0C2')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 255).setAllowInvalid(false).build());
  titulo('B11:F12', '1. Elige una instrucción y escribe enteros de 0 a 255.\n2. Menú Simulador CPU → Cargar operación elegida.\n3. Usa PASO o EJECUTAR en el diagrama.');
  titulo('B14:F16', 'ADD suma; SUB resta; AND/OR/XOR operan bit a bit. CMP compara sin modificar AX/BX. INC, DEC y NOT necesitan un solo dato. REN1 y REN2 son las entradas internas de la ALU; AC recoge su resultado.');
  titulo('B18:F20', 'El programa lee primero A y B desde RAM. En las variantes imm, el segundo operando es el inmediato. Al terminar, AX queda también en RAM[82h] y BX en RAM[83h]. LOAD usa 80h; STORE escribe en 84h.');
  titulo('B22:F23', 'JMP/JZ/JNZ incluyen CMP AX,BX y un salto de prueba: si no se toma, AX recibe EEh. HLT se detiene sin almacenamiento final. Para un bucle: menú → Cargar demo: multiplicación.');
  const mem = libro.getSheetByName(HOJA.MEMORIA);
  mem.getRange('U5:X5').setValues([['Registro', 'HEX', 'DEC', 'BIN']]).setFontWeight('bold');
  mem.getRange('V6:V15').setNumberFormat('@'); mem.getRange('X6:X15').setNumberFormat('@');
  mem.setColumnWidths(21, 3, 75).setColumnWidth(24, 105);
  onOpen();
  conBloqueo_(props => { props.deleteProperty(RUN_KEY); pintar_(sesion_(props), null, false); });
  sh.activate();
}

/* ================== 8. PREPARAR DIAGRAMA (layout, buses y vínculos) ================== */
function prepararDiagrama() {
  const libro = SpreadsheetApp.getActive(), sh = libro.getSheetByName(HOJA.DIAGRAMA);
  asegurarEstado_();
  if (sh.getMaxColumns() < 50) sh.insertColumnsAfter(sh.getMaxColumns(), 50 - sh.getMaxColumns());
  if (sh.getMaxRows() < 52) sh.insertRowsAfter(sh.getMaxRows(), 52 - sh.getMaxRows());
  sh.getRange('A1:AX52').breakApart().clear();
  sh.setHiddenGridlines(true).setColumnWidths(1, 50, ANCHO_COL).setRowHeights(1, 52, ALTO_FILA);
  sh.getRange('A1:AX52').setBackground(C.fondo).setFontFamily('Arial').setFontSize(10).setFontColor(C.texto).setVerticalAlignment('middle');
  const texto = (rango, valor, tam, color) => sh.getRange(rango).merge().setValue(valor).setFontSize(tam || 10).setFontColor(color || C.texto).setWrap(true);
  const caja = (rango, color) => sh.getRange(rango).setBackground(color)
    .setBorder(true, true, true, true, false, false, '#111111', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  texto('B1:AX2', 'SIMULADOR CPU · CICLO DE INSTRUCCIÓN', 16, '#FFFFFF');
  texto('B4:AL5', '', 11, '#FFFFFF');                       // franja de botones (figuras)
  texto('AN3:AX3', 'Pausa entre pasos (100–2000 ms)', 9, '#FFFFFF');
  texto('AQ4:AX5', 500, 11);
  sh.getRange('AQ4:AX5').setBackground('#FFFFFF')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(100, 2000).setAllowInvalid(false).build());
  caja('B7:X24', C.uc); caja('AB7:AX24', C.alu); caja('T29:AT48', C.mem);
  texto('B6:X6', 'Unidad de Control', 11, '#FFFFFF');
  texto('AB6:AX6', 'Unidad Aritmético Lógica', 11, '#FFFFFF');
  texto('Y28:AJ28', 'Memoria Central', 11, '#FFFFFF');

  const etiquetas = {CLOCK: 'Reloj', SEQ: 'Secuenciador', DEC: 'Decodificador', IR: 'R.I. · opcode', OP: 'Operando', PC: 'C.P.',
    AC: 'Acumulador', FLAGS: 'R. Estado', AX: 'AX', BX: 'BX', REN1: 'REN 1', REN2: 'REN 2', MAR: 'RDM / MAR', MDR: 'RIM / MDR', ALU: 'C. OP.'};
  Object.keys(REGIONES).forEach(k => {
    const g = rango_(REGIONES[k]);
    sh.getRange(REGIONES[k]).merge().setHorizontalAlignment('center')
      .setFontSize(['FLAGS', 'DEC', 'SEQ'].includes(k) ? 10 : 14).setWrap(true);
    if (k !== 'ALU') caja(REGIONES[k], ['FLAGS', 'DEC', 'SEQ', 'CLOCK'].includes(k) ? (g.c1 < 25 ? C.uc : C.alu) : C.valor);
    const lab = LABEL_RANGO[k] || letras_(g.c1) + (g.r1 - 1) + ':' + letras_(g.c2) + (g.r1 - 1);
    sh.getRange(lab).merge().setValue(etiquetas[k]).setFontSize(10);
  });
  texto('I17:N19', '↓ ↓ ↓ ↓ ↓ ↓\nMicroórdenes', 10);
  texto('V37:AB37', 'Selector', 10); texto('V38:AB40', '', 10); caja('V38:AB40', C.mem);
  texto('AN27:AX28', 'Constante interna: 1 / 0', 9); sh.getRange('AN27:AX28').setBackground(C.alu);
  texto('AF37:AQ38', 'Memoria · dirección / contenido', 10);
  for (let i = 0; i < 8; i++) {
    sh.getRange(39 + i, 32, 1, 3).merge().setHorizontalAlignment('center');
    sh.getRange(39 + i, 35, 1, 9).merge();
    caja('AF' + (39 + i) + ':AQ' + (39 + i), C.mem);
  }
  sh.getRange('AF39:AH46').setNumberFormat('@');
  texto('B25:R26', 'Entradas ALU', 12, '#FFFFFF');
  ['B27:R28', 'B30:R33', 'B35:R37', 'B39:R48', 'B50:AX51'].forEach(r => texto(r, '', r === 'B27:R28' ? 14 : 10, '#FFFFFF'));
  sh.getRange('B39:R48').setVerticalAlignment('top');

  // Buses: se dibujan con celdas (las figuras de bus no pueden crearse por código)
  sh.getRangeList(BUS_DIR).setBackground(C.busDir);
  sh.getRangeList(BUS_DAT).setBackground(C.busDat);
  sh.getRange('W27').setValue('BD').setFontSize(7).setFontColor('#FFFFFF').setHorizontalAlignment('center');
  sh.getRange('AM26').setValue('BDt').setFontSize(7).setFontColor(C.texto).setHorizontalAlignment('center');
  texto('B52:X52', 'BD = Bus de Direcciones (oscuro) · BDt = Bus de Datos (claro) · el cuadrado rojo recorre el bus activo', 8, '#FFFFFF');

  // Hoja Memoria: encabezados, leyenda de segmentos y mnemónico
  const mem = libro.getSheetByName(HOJA.MEMORIA);
  mem.getRange('B24').setValue('Dirección HEX'); mem.getRange('F24').setValue('Dato HEX');
  mem.getRange('I24').setValue('Decimal'); mem.getRange('L24').setValue('Binario'); mem.getRange('O24').setValue('Mnemónico');
  mem.getRange('C25').setNumberFormat('@').setValue('80'); mem.getRange('F25').setNumberFormat('@').setValue('05');
  mem.getRange('C5:R5').setNumberFormat('@').setValues([Array.from({length: 16}, (_, i) => i.toString(16).toUpperCase())]).setFontWeight('bold').setHorizontalAlignment('center');
  mem.getRange('B6:B21').setNumberFormat('@').setValues(Array.from({length: 16}, (_, i) => [i.toString(16).toUpperCase() + '0h'])).setFontWeight('bold');
  mem.getRange('C22:R22').breakApart();
  mem.getRange('C22:J22').merge().setValue('Segmento de CÓDIGO 00h–7Fh').setBackground(C.uc).setHorizontalAlignment('center').setFontWeight('bold');
  mem.getRange('K22:R22').merge().setValue('Segmento de DATOS 80h–FFh').setBackground(C.alu).setHorizontalAlignment('center').setFontWeight('bold');
  mem.getRange('B27:R28').breakApart().merge().setValue('Menú Simulador CPU → Leer / Escribir memoria. Pausa antes de escribir.').setWrap(true);
  const prog = libro.getSheetByName(HOJA.PROGRAMA);
  if (prog) prog.getRange('D6').setValue('Assembly (desensamblado)').setFontWeight('bold');

  enlazar_(sh);
  onOpen();
  conBloqueo_(props => { props.deleteProperty(RUN_KEY); pintar_(sesion_(props), null, false); });
  sh.activate();
}
