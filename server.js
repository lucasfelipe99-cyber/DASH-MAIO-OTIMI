const http = require('http');
const fs = require('fs');
const path = require('path');

const port = process.env.PORT || 8787;
const projectDir = __dirname;
const dataDir = path.join(projectDir, 'data');
const metadataPath = path.join(dataDir, 'metadata.json');
const allowedBaseExtensions = new Set(['.xlsx', '.xls', '.csv']);
const monthKeys = new Set(Array.from({ length: 12 }, (_, index) => String(index + 1)));

const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel',
  '.csv': 'text/csv; charset=utf-8',
  '.json': 'application/json; charset=utf-8'
};

fs.mkdirSync(dataDir, { recursive: true });

function sendFile(response, filePath, contentType) {
  fs.readFile(filePath, (error, content) => {
    if (error) {
      response.writeHead(error.code === 'ENOENT' ? 404 : 500, {
        'Content-Type': 'text/plain; charset=utf-8'
      });
      response.end(error.code === 'ENOENT' ? 'Arquivo nao encontrado.' : 'Erro ao ler arquivo.');
      return;
    }

    response.writeHead(200, {
      'Content-Type': contentType,
      'Cache-Control': 'no-store'
    });
    response.end(content);
  });
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store'
  });
  response.end(JSON.stringify(payload));
}

function sendText(response, statusCode, text) {
  response.writeHead(statusCode, {
    'Content-Type': 'text/plain; charset=utf-8',
    'Cache-Control': 'no-store'
  });
  response.end(text);
}

function resolveProjectFile(urlPath) {
  const decodedPath = decodeURIComponent(urlPath.split('?')[0]);
  const relativePath = decodedPath === '/' ? 'index.html' : decodedPath.replace(/^\/+/, '');
  const resolvedPath = path.resolve(projectDir, relativePath);

  if (!resolvedPath.startsWith(projectDir)) {
    return null;
  }

  return resolvedPath;
}

function normalizeMonth(value) {
  const month = String(value || '').trim();

  return monthKeys.has(month) ? month : '';
}

function getMonthFromUrl(url) {
  try {
    const parsedUrl = new URL(url, 'http://localhost');

    return normalizeMonth(parsedUrl.searchParams.get('month'));
  } catch (error) {
    return '';
  }
}

function readMetadata() {
  if (!fs.existsSync(metadataPath)) {
    return { months: {} };
  }

  try {
    const metadata = JSON.parse(fs.readFileSync(metadataPath, 'utf8'));

    if (metadata.months) {
      return metadata;
    }

    if (metadata.storedName) {
      const legacyMonth = metadata.updatedAt
        ? String(new Date(metadata.updatedAt).getMonth() + 1)
        : String(new Date().getMonth() + 1);

      return { months: { [legacyMonth]: metadata } };
    }

    return { months: {} };
  } catch (error) {
    return { months: {} };
  }
}

function getPublishedMetadata(month) {
  const metadata = readMetadata();
  const monthMetadata = metadata.months[month] || {};
  const filePath = path.join(dataDir, monthMetadata.storedName || '');

  if (!monthMetadata.storedName || !fs.existsSync(filePath)) {
    return { exists: false, month };
  }

  return {
    exists: true,
    month,
    fileName: monthMetadata.fileName,
    storedName: monthMetadata.storedName,
    rowsName: monthMetadata.rowsName,
    updatedAt: monthMetadata.updatedAt,
    rowsUpdatedAt: monthMetadata.rowsUpdatedAt,
    size: monthMetadata.size,
    url: '/data/' + encodeURIComponent(monthMetadata.storedName),
    rowsUrl: monthMetadata.rowsName && fs.existsSync(path.join(dataDir, monthMetadata.rowsName))
      ? '/data/' + encodeURIComponent(monthMetadata.rowsName)
      : ''
  };
}

function getAllPublishedMetadata() {
  const metadata = readMetadata();
  const months = {};

  monthKeys.forEach((month) => {
    const monthMetadata = getPublishedMetadata(month);

    if (monthMetadata.exists) {
      months[month] = monthMetadata;
    }
  });

  return {
    exists: Object.keys(months).length > 0,
    months
  };
}

function deletePreviousBases(month) {
  fs.readdirSync(dataDir).forEach((name) => {
    const pattern = new RegExp('^current-(base|rows)-' + month + '(?:-[0-9]+)?\\.(xlsx|xls|csv|json)$', 'i');

    if (pattern.test(name)) {
      fs.unlinkSync(path.join(dataDir, name));
    }
  });
}

function writeMonthMetadata(month, nextMetadata) {
  const metadata = readMetadata();

  metadata.months[month] = nextMetadata;
  fs.writeFileSync(metadataPath, JSON.stringify(metadata, null, 2));
}

function getLegacyPublishedMetadata() {
  if (!fs.existsSync(metadataPath)) {
    return { exists: false };
  }

  try {
    const metadata = JSON.parse(fs.readFileSync(metadataPath, 'utf8'));
    const filePath = path.join(dataDir, metadata.storedName || '');

    if (!metadata.storedName || !fs.existsSync(filePath)) {
      return { exists: false };
    }

    return {
      exists: true,
      fileName: metadata.fileName,
      storedName: metadata.storedName,
      updatedAt: metadata.updatedAt,
      size: metadata.size,
      url: '/data/' + encodeURIComponent(metadata.storedName)
    };
  } catch (error) {
    return { exists: false };
  }
}

function collectRequestBody(request, callback) {
  const chunks = [];

  request.on('data', (chunk) => {
    chunks.push(chunk);
  });

  request.on('end', () => {
    callback(null, Buffer.concat(chunks));
  });

  request.on('error', (error) => {
    callback(error);
  });
}

function handleBaseUpload(request, response) {
  const configuredPassword = process.env.ADMIN_PASSWORD;
  const providedPassword = request.headers['x-admin-password'] || '';
  const month = normalizeMonth(request.headers['x-base-month']);

  if (!configuredPassword) {
    sendText(response, 500, 'ADMIN_PASSWORD nao configurada no servidor.');
    return;
  }

  if (providedPassword !== configuredPassword) {
    sendText(response, 401, 'Senha invalida.');
    return;
  }

  if (!month) {
    sendText(response, 400, 'Mes invalido para publicacao.');
    return;
  }

  const originalName = decodeURIComponent(request.headers['x-file-name'] || '');
  const extension = path.extname(originalName).toLowerCase();

  if (!allowedBaseExtensions.has(extension)) {
    sendText(response, 400, 'Formato nao aceito. Envie .xlsx, .xls ou .csv.');
    return;
  }

  collectRequestBody(request, (error, body) => {
    if (error || !body || body.length === 0) {
      sendText(response, 400, 'Arquivo vazio ou invalido.');
      return;
    }

    try {
      const storedName = 'current-base-' + month + extension;
      const storedPath = path.join(dataDir, storedName);
      const updatedAt = new Date().toISOString();
      const metadata = {
        fileName: path.basename(originalName),
        storedName,
        updatedAt,
        size: body.length
      };

      deletePreviousBases(month);
      fs.writeFileSync(storedPath, body);
      writeMonthMetadata(month, metadata);

      sendJson(response, 200, getPublishedMetadata(month));
    } catch (writeError) {
      console.error('Erro ao publicar base:', writeError);
      sendText(response, 500, 'Erro ao salvar a base no servidor: ' + writeError.message);
    }
  });
}

function handleRowsUpload(request, response) {
  const configuredPassword = process.env.ADMIN_PASSWORD;
  const providedPassword = request.headers['x-admin-password'] || '';
  const month = normalizeMonth(request.headers['x-base-month']);

  if (!configuredPassword) {
    sendText(response, 500, 'ADMIN_PASSWORD nao configurada no servidor.');
    return;
  }

  if (providedPassword !== configuredPassword) {
    sendText(response, 401, 'Senha invalida.');
    return;
  }

  if (!month) {
    sendText(response, 400, 'Mes invalido para publicacao.');
    return;
  }

  collectRequestBody(request, (error, body) => {
    if (error || !body || body.length === 0) {
      sendText(response, 400, 'Dados processados vazios ou invalidos.');
      return;
    }

    try {
      const metadata = readMetadata();
      const monthMetadata = metadata.months[month];

      if (!monthMetadata || !monthMetadata.storedName) {
        sendText(response, 400, 'Publique a base do mes antes de salvar os dados processados.');
        return;
      }

      const rowsName = 'current-rows-' + month + '.json';
      const rowsPath = path.join(dataDir, rowsName);

      fs.writeFileSync(rowsPath, body);
      monthMetadata.rowsName = rowsName;
      monthMetadata.rowsUpdatedAt = new Date().toISOString();
      metadata.months[month] = monthMetadata;
      fs.writeFileSync(metadataPath, JSON.stringify(metadata, null, 2));

      sendJson(response, 200, getPublishedMetadata(month));
    } catch (writeError) {
      console.error('Erro ao salvar dados processados:', writeError);
      sendText(response, 500, 'Erro ao congelar os dados do mes: ' + writeError.message);
    }
  });
}

const server = http.createServer((request, response) => {
  if (request.method === 'GET' && request.url.startsWith('/api/latest-base')) {
    const month = getMonthFromUrl(request.url);

    sendJson(response, 200, month ? getPublishedMetadata(month) : getAllPublishedMetadata());
    return;
  }

  if (request.method === 'POST' && request.url.startsWith('/api/upload-base')) {
    handleBaseUpload(request, response);
    return;
  }

  if (request.method === 'POST' && request.url.startsWith('/api/upload-rows')) {
    handleRowsUpload(request, response);
    return;
  }

  const filePath = resolveProjectFile(request.url);

  if (!filePath) {
    sendText(response, 403, 'Acesso negado.');
    return;
  }

  const extension = path.extname(filePath).toLowerCase();
  sendFile(response, filePath, mimeTypes[extension] || 'application/octet-stream');
});

server.listen(port, '0.0.0.0', () => {
  console.log(`Dashboard rodando na porta ${port}`);
});
