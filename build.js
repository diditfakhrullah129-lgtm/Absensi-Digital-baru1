const fs = require('fs');
const url = process.argv[2];
if (!url) { console.error('Pakai: node build.js https://script.google.com/macros/s/AKfycbxbVXvyDyGgWoRZPC9YkZpd4w1A_5eSrRVZtZmg0NZzfTvYbncg7ZAPIy1K7PfPdsai/exec'); process.exit(1); }

let html = fs.readFileSync('index.html', 'utf8');
html = html.replace(/<\?!=\s*include\('([^']+)'\)\s*\?>/g, (m, n) =>
  fs.existsSync(n + '.html') ? fs.readFileSync(n + '.html', 'utf8') : '<!-- ' + n + '.html tidak ditemukan -->');

const shim = fs.readFileSync('gas-shim.js', 'utf8');
html = html.replace('<head>', () =>
  '<head>\n<script>window.GAS_API_URL=' + JSON.stringify(url) + ';</script>\n<script>' + shim + '</script>');

fs.mkdirSync('dist', { recursive: true });
fs.writeFileSync('dist/index.html', html);
console.log('Selesai: dist/index.html');
