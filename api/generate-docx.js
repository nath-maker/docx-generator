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

    // Create a simple HTML document that Word can open
    const htmlContent = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:w="urn:schemas-microsoft-com:office:word"
            xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8">
        <title>${data.title || 'Document'}</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
          @page { margin: 1in; }
          body { font-family: Arial, sans-serif; line-height: 1.6; }
          h1 { color: #333; font-size: 24pt; }
          h2 { color: #666; font-size: 18pt; margin-top: 20pt; }
          p { margin: 10pt 0; }
        </style>
      </head>
      <body>
        <h1>${data.title || 'Document'}</h1>
        <p><strong>Authors:</strong> ${data.authors || ''}</p>
        <p><strong>Year:</strong> ${data.year || ''}</p>
        <p><strong>Citation:</strong> ${data.citation || ''}</p>
        
        <h2>Synthesis</h2>
        <p>${data.synthesis || ''}</p>
        
        <h2>Critical Analysis</h2>
        <p>${data.critical || ''}</p>
        
        <h2>Prospector Notes</h2>
        <p>${data.prospector || ''}</p>
      </body>
      </html>
    `;

    // Sanitize filename
    const safeFilename = (data.filename || 'document.doc')
      .replace(/[^a-zA-Z0-9.-_]/g, '_')
      .replace(/\.docx?$/i, '') + '.doc';

    // Set headers for Word document
    res.setHeader('Content-Type', 'application/msword');
    res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}"`);
    
    // Send the HTML as a .doc file (Word can open HTML with .doc extension)
    res.status(200).send(htmlContent);

  } catch (error) {
    console.error('Error generating document:', error);
    res.status(500).json({ error: 'Failed to generate document', details: error.message });
  }
}
