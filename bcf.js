// BCF 2.1 export — turn merged Studio markups into a portable .bcfzip that
// ACC, Newforma Konekt, and other openBIM tools import natively. Topics +
// comments (no snapshots in v1). Zero-dependency store-mode ZIP writer, so the
// deploy picks up no new packages.
const crypto = require('crypto');

const CRC_TABLE = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c;
  }
  return t;
})();
function crc32(buf) {
  let c = 0 ^ -1;
  for (let i = 0; i < buf.length; i++) c = (c >>> 8) ^ CRC_TABLE[(c ^ buf[i]) & 0xFF];
  return (c ^ -1) >>> 0;
}

// entries: [{ name, data(Buffer) }] -> a valid uncompressed ZIP Buffer.
function zipStore(entries) {
  const now = new Date();
  const dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (Math.floor(now.getSeconds() / 2))) & 0xFFFF;
  const dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const e of entries) {
    const nameBuf = Buffer.from(e.name, 'utf8');
    const data = e.data;
    const crc = crc32(data);
    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);           // version needed
    local.writeUInt16LE(0x0800, 6);       // flags: UTF-8 filename
    local.writeUInt16LE(0, 8);            // method: store
    local.writeUInt16LE(dosTime, 10);
    local.writeUInt16LE(dosDate, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18); // compressed size
    local.writeUInt32LE(data.length, 22); // uncompressed size
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);           // extra len
    nameBuf.copy(local, 30);
    locals.push(local, data);

    const central = Buffer.alloc(46 + nameBuf.length);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);         // version made by
    central.writeUInt16LE(20, 6);         // version needed
    central.writeUInt16LE(0x0800, 8);     // flags: UTF-8
    central.writeUInt16LE(0, 10);         // method
    central.writeUInt16LE(dosTime, 12);
    central.writeUInt16LE(dosDate, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);         // extra len
    central.writeUInt16LE(0, 32);         // comment len
    central.writeUInt16LE(0, 34);         // disk number
    central.writeUInt16LE(0, 36);         // internal attrs
    central.writeUInt32LE(0, 38);         // external attrs
    central.writeUInt32LE(offset, 42);    // local header offset
    nameBuf.copy(central, 46);
    centrals.push(central);
    offset += local.length + data.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralBuf.length, 12); // size of central directory
  end.writeUInt32LE(offset, 16);          // central dir offset = end of locals
  end.writeUInt16LE(0, 20);               // ZIP comment length
  return Buffer.concat([...locals, centralBuf, end]);
}

function xmlEsc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
function bcfGuid() { return crypto.randomUUID(); }

// Map a Bluebeam markup status to a BCF TopicStatus while keeping the raw value.
function bcfStatus(raw) {
  const s = String(raw || '').toLowerCase();
  if (/(closed|resolved|verified|complete|accept|approved)/.test(s)) return 'Closed';
  if (/(progress|review|rework|reopen)/.test(s)) return 'In Progress';
  return 'Open';
}

function buildBcfBuffer(session, file, markups) {
  const nowIso = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const sid = (session && session.Id) || '';
  const docName = (file && file.name) || '';
  const entries = [];
  entries.push({
    name: 'bcf.version',
    data: Buffer.from(
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Version VersionId="2.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n' +
      '  <DetailedVersion>2.1</DetailedVersion>\n' +
      '</Version>\n', 'utf8')
  });
  (markups || []).forEach((m, i) => {
    const topicGuid = bcfGuid();
    const rawStatus = m.status || '';
    const title = [m.subject || m.type || 'Markup',
      rawStatus ? '(' + rawStatus.replace(/_/g, ' ') + ')' : '',
      'Sheet ' + (m.pageNumber || 1)].filter(Boolean).join(' ');
    const author = m.author || 'Unknown';
    const descLines = [];
    if (m.contents) descLines.push(m.contents);
    descLines.push('Sheet: ' + (m.pageNumber || 1));
    if (m.type) descLines.push('Markup type: ' + m.type);
    if (rawStatus) descLines.push('Studio status: ' + rawStatus);
    if (Array.isArray(m.rect)) descLines.push('PDF rect: [' + m.rect.join(', ') + ']');
    descLines.push('Source: Bluebeam Studio Session ' + sid + (docName ? ' / ' + docName : ''));
    const description = descLines.join('\n');
    const labels = rawStatus ? '    <Labels>' + xmlEsc(rawStatus) + '</Labels>\n' : '';
    const commentXml = m.contents
      ? '  <Comment Guid="' + bcfGuid() + '">\n' +
        '    <Date>' + nowIso + '</Date>\n' +
        '    <Author>' + xmlEsc(author) + '</Author>\n' +
        '    <Comment>' + xmlEsc(m.contents) + '</Comment>\n' +
        '  </Comment>\n'
      : '';
    const markupXml =
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Markup>\n' +
      '  <Topic Guid="' + topicGuid + '" TopicType="Issue" TopicStatus="' + bcfStatus(rawStatus) + '">\n' +
      '    <Title>' + xmlEsc(title) + '</Title>\n' +
      '    <Index>' + i + '</Index>\n' +
      labels +
      '    <CreationDate>' + nowIso + '</CreationDate>\n' +
      '    <CreationAuthor>' + xmlEsc(author) + '</CreationAuthor>\n' +
      '    <Description>' + xmlEsc(description) + '</Description>\n' +
      '  </Topic>\n' +
      commentXml +
      '</Markup>\n';
    entries.push({ name: topicGuid + '/markup.bcf', data: Buffer.from(markupXml, 'utf8') });
  });
  return zipStore(entries);
}

module.exports = { buildBcfBuffer, zipStore, crc32, bcfStatus };
