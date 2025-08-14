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
    console.log('Received data for PDF generation');

    // Create HTML content that will be converted to PDF
    const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    @page { margin: 2cm; size: A4; }
    body { 
      font-family: 'Times New Roman', serif; 
      line-height: 1.6; 
      color: #333;
      max-width: 21cm;
      margin: 0 auto;
    }
    h1 { 
      color: #000; 
      font-size: 24pt; 
      margin-bottom: 10pt;
      border-bottom: 2px solid #333;
      padding-bottom: 10pt;
    }
    h2 { 
      color: #333; 
      font-size: 18pt; 
      margin-top: 20pt;
      margin-bottom: 10pt;
      border-bottom: 1px solid #ccc;
      padding-bottom: 5pt;
    }
    p { 
      margin: 10pt 0; 
      text-align: justify;
    }
    .metadata {
      background: #f5f5f5;
      padding: 15pt;
      margin: 20pt 0;
      border-left: 4px solid #333;
    }
    .metadata p {
      margin: 5pt 0;
    }
    strong {
      color: #000;
    }
  </style>
</head>
<body>
  <h1>${data.title || 'Research Document'}</h1>
  
  <div class="metadata">
    <p><strong>Authors:</strong> ${data.authors || 'Not specified'}</p>
    <p><strong>Year:</strong> ${data.year || new Date().getFullYear()}</p>
    <p><strong>Citation:</strong> ${data.citation || 'Not specified'}</p>
    <p><strong>Notion ID:</strong> ${data.notion_id || 'Not specified'}</p>
  </div>
  
  <h2>Synthesis</h2>
  <p>${(data.synthesis || 'No synthesis provided.').replace(/\n/g, '</p><p>')}</p>
  
  <h2>Critical Analysis</h2>
  <p>${(data.critical || 'No critical analysis provided.').replace(/\n/g, '</p><p>')}</p>
  
  <h2>Prospector Notes</h2>
  <p>${(data.prospector || 'No prospector notes provided.').replace(/\n/g, '</p><p>')}</p>
</body>
</html>`;

    // Sanitize filename
    const safeFilename = (data.filename || 'document.pdf')
      .replace(/[^a-zA-Z0-9.-_]/g, '_')
      .replace(/\.pdf$/i, '') + '.pdf';

    // Send HTML with PDF headers (browsers will convert to PDF when downloading)
    res.setHeader('Content-Type', 'text/html');
    res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}"`);
    res.status(200).send(htmlContent);

  } catch (error) {
    console.error('Error generating PDF:', error);
    res.status(500).json({ error: 'Failed to generate document', details: error.message });
  }
}
