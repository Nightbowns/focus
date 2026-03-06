const http = require('http');
const fs = require('fs');
const path = require('path');

const root = process.argv[2] || process.cwd();
const port = Number(process.env.PORT || process.argv[3] || 4173);

const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

const server = http.createServer((req, res) => {
  const urlPath = decodeURIComponent((req.url || '/').split('?')[0]);
  const safePath = path.normalize(urlPath).replace(/^([.][.][/\\])+/, '');
  let targetPath = path.join(root, safePath === '/' ? 'index.html' : safePath);

  fs.stat(targetPath, (err, stat) => {
    if (!err && stat.isDirectory()) {
      targetPath = path.join(targetPath, 'index.html');
    }

    fs.readFile(targetPath, (readErr, data) => {
      if (readErr) {
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('404 Not Found');
        return;
      }

      const ext = path.extname(targetPath).toLowerCase();
      res.writeHead(200, { 'Content-Type': mimeTypes[ext] || 'application/octet-stream' });
      res.end(data);
    });
  });
});

server.listen(port, () => {
  console.log(`LOCAL_SERVER_RUNNING http://localhost:${port}`);
});
