// Revisa el código de Apps Script sin ejecutarlo:
//  1. que cada archivo .gs tenga sintaxis válida;
//  2. que no haya dos funciones con el mismo nombre (en Apps Script todos
//     los archivos comparten el mismo espacio y una pisaría a la otra);
//  3. que cada opción del menú y cada trigger apunte a una función que existe.
// Uso: node tools/verificar.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const dir = path.join(__dirname, '..', 'src');
const archivos = fs.readdirSync(dir).filter(f => f.endsWith('.gs'));
const definidas = new Map();
const errores = [];
let codigo = '';

for (const f of archivos) {
  const src = fs.readFileSync(path.join(dir, f), 'utf8');
  try {
    new vm.Script(src, { filename: f });
  } catch (e) {
    errores.push(`${f}: error de sintaxis: ${e.message}`);
  }
  for (const m of src.matchAll(/^function\s+([A-Za-z_$][\w$]*)\s*\(/gm)) {
    if (definidas.has(m[1])) errores.push(`función repetida: ${m[1]} (${definidas.get(m[1])} y ${f})`);
    else definidas.set(m[1], f);
  }
  codigo += src + '\n';
}

const referidas = new Set();
for (const m of codigo.matchAll(/\.addItem\(\s*'[^']*'\s*,\s*'([^']+)'\s*\)/g)) referidas.add(m[1]);
for (const m of codigo.matchAll(/newTrigger\(\s*"([^"]+)"\s*\)/g)) referidas.add(m[1]);
for (const nombre of referidas) {
  if (!definidas.has(nombre)) errores.push(`el menú o un trigger llama a "${nombre}", que no existe`);
}

if (errores.length) {
  console.error(errores.join('\n'));
  process.exit(1);
}
console.log(`OK: ${archivos.length} archivos, ${definidas.size} funciones, ${referidas.size} entradas de menú y triggers verificadas.`);
