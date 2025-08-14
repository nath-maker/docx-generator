import JSZip from 'jszip';

export default async function handler(req, res) {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  try {
    const { data } = req.body;
    console.log('Received data:', JSON.stringify(data, null, 2));

    // Create a new ZIP file
    const zip = new JSZip();

    // Add _rels folder
    const relsFolder = zip.folder('_rels');
    relsFolder.file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

    // Add word folder
    const wordFolder = zip.folder('word');
    
    // Add _rels in word folder
    const wordRelsFolder = wordFolder.folder('_rels');
    wordRelsFolder.file('document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`);

    // Create the main document content with proper escaping
    function escapeXml(text) {
      if (!text) return '';
      return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    }

    // Add document.xml with your content
    const documentContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>${escapeXml(data.title || 'Document')}</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Authors: ${escapeXml(data.authors || '')}</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Year: ${data.year || ''}</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Citation: ${escapeXml(data.citation || '')}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Synthesis</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>${escapeXml(data.synthesis || '')}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Critical Analysis</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>${escapeXml(data.critical || '')}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Prospector Notes</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>${escapeXml(data.prospector || '')}</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    
    wordFolder.file('document.xml', documentContent);

    // Add [Content_Types].xml
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

    // Generate the DOCX file
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });

    // Sanitize filename for Content-Disposition header
    const safeFilename = (data.filename || 'document.docx')
      .replace(/[^a-zA-Z0-9.-_]/g, '_');

    // Set response headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}"`);
    
    // Send the file
    res.status(200).send(buffer);

  } catch (error) {
    console.error('Error generating DOCX:', error);
    console.error('Error stack:', error.stack);
    res.status(500).json({ error: 'Failed to generate document', details: error.message });
  }
}
