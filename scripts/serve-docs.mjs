import { createReadStream } from 'node:fs';
import { stat } from 'node:fs/promises';
import { createServer } from 'node:http';
import { extname, resolve, sep } from 'node:path';

const arguments_ = process.argv.slice(2);
const option = (name, fallback) => {
  const index = arguments_.indexOf(name);
  return index < 0 ? fallback : arguments_[index + 1];
};
const host = option('--host', '127.0.0.1');
const port = Number(option('--port', '4175'));
const publicDirectory = resolve(import.meta.dirname, '../website/build');
const basePath = '/tego-sheet';
const contentTypes = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.ico': 'image/x-icon',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

async function resolveFile(requestUrl) {
  const url = new URL(requestUrl ?? '/', `http://${host}:${port}`);
  let pathname = decodeURIComponent(url.pathname);
  if (pathname === basePath) pathname = '/';
  else if (pathname.startsWith(`${basePath}/`)) pathname = pathname.slice(basePath.length);
  else return undefined;

  const relativePath = pathname.replace(/^\/+/, '');
  const candidates = [
    relativePath,
    relativePath === '' || relativePath.endsWith('/')
      ? `${relativePath}index.html`
      : `${relativePath}.html`,
  ];
  for (const candidate of candidates) {
    const filePath = resolve(publicDirectory, candidate);
    if (filePath !== publicDirectory && !filePath.startsWith(`${publicDirectory}${sep}`)) continue;
    try {
      if ((await stat(filePath)).isFile()) return filePath;
    } catch {
      // Try the next clean-URL candidate.
    }
  }
  return undefined;
}

const server = createServer(async (request, response) => {
  const filePath = await resolveFile(request.url);
  if (filePath === undefined) {
    response.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' });
    response.end('Not found');
    return;
  }
  response.writeHead(200, {
    'cache-control': 'no-store',
    'content-type': contentTypes[extname(filePath)] ?? 'application/octet-stream',
  });
  createReadStream(filePath).pipe(response);
});

server.listen(port, host, () => {
  process.stdout.write(`Serving documentation at http://${host}:${port}${basePath}/\n`);
});
