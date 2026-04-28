const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');

const PORT = 3001;

// Use office-addin-dev-certs which are already trusted by Office
const devCertsDir = path.join(os.homedir(), '.office-addin-dev-certs');

// Find the cert files
let keyFile  = path.join(devCertsDir, 'localhost.key');
let certFile = path.join(devCertsDir, 'localhost.crt');
let caFile   = path.join(devCertsDir, 'ca.crt');

// Fallback to installer certs folder
if (!fs.existsSync(keyFile)) {
  keyFile  = path.join(__dirname, 'certs', 'server.key');
  certFile = path.join(__dirname, 'certs', 'server.cert');
  caFile   = null;
}

if (!fs.existsSync(keyFile)) {
  console.error('ERROR: No certificates found.');
  console.error('Expected at: ' + devCertsDir);
  console.error('Please run: npx office-addin-dev-certs install --machine');
  process.exit(1);
}

console.log('Using certificates from:', path.dirname(keyFile));

const options = {
  key:  fs.readFileSync(keyFile),
  cert: fs.readFileSync(certFile),
};
if (caFile && fs.existsSync(caFile)) {
  options.ca = fs.readFileSync(caFile);
}

const mimeTypes = {
  '.html': 'text/html',
  '.js':   'application/javascript',
  '.css':  'text/css',
  '.png':  'image/png',
  '.xml':  'application/xml',
  '.json': 'application/json'
};

const server = https.createServer(options, (req, res) => {
  let filePath = path.join(__dirname, 'app', req.url === '/' ? 'taskpane.html' : req.url);

  filePath = filePath.split('?')[0];

  const ext = path.extname(filePath);
  const contentType = mimeTypes[ext] || 'text/plain';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found: ' + req.url);
      return;
    }
    res.writeHead(200, {
      'Content-Type': contentType,
      'Access-Control-Allow-Origin': '*'
    });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log('Electrical Node Designer server running at https://localhost:' + PORT);
  console.log('Press Ctrl+C to stop');
});
