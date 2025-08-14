import JSZip from 'jszip';

export default async function handler(req, res) {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { data } = req.body;
    console.log('Received data:', JSON.stringify(data, null, 2));
    console.log('About to create ZIP file...');
    
    // Create a new ZIP file
    const zip = new JSZip();
    
    // Add [Content_Types].xml
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

    // Add _rels/.rels
    const relsFolder = zip.folder('_rels');
    relsFolder.file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

    // Create word folder structure
    const wordFolder = zip.folder('word');
    const wordRelsFolder = wordFolder.folder('_rels');
    
    // Add empty document.xml.rels
    wordRelsFolder.file('document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`);

    // Build document.xml with our formatted content
    let documentContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>`;

    // TITLE PAGE
    documentContent += createTitle(data.title || 'Paper Analysis Report');
    documentContent += createParagraph(data.authors || 'Unknown Authors');
    documentContent += createParagraph(data.citation || '');
    documentContent += createParagraph(`Analysis Report - PhD Research`);
    documentContent += createParagraph(`Notion ID: ${data.notion_id || 'Unknown'}`);
    documentContent += createParagraph(new Date().toLocaleDateString());
    documentContent += pageBreak();

    // EXECUTIVE SUMMARY
    documentContent += createHeading1('Executive Summary');
    documentContent += createParagraph(`Paper Quality: ${data.quality_assessment || 'Not assessed'}`);
    documentContent += createParagraph(`Gender Level: ${data.gender_level || 0} (${data.gender_treatment || 'not specified'})`);
    documentContent += createParagraph(`Disjuncture: ${data.disjuncture_type || 'unclear'} at ${data.disjuncture_level || 'unknown'} level`);
    documentContent += createParagraph(`Building Potential: ${data.building_potential || 'unclear'}`);
    
    // Topics
    const topics = [];
    if (data.topic_ai_ethics) topics.push('AI Ethics');
    if (data.topic_future_work) topics.push('Future of Work');
    if (data.topic_gender_bias) topics.push('Gender & Bias');
    if (data.topic_measurement) topics.push('Measurement');
    if (data.topic_performance_mgmt) topics.push('Performance Management');
    documentContent += createParagraph(`Key Topics: ${topics.join(', ') || 'None identified'}`);
    
    // Nuggets
    documentContent += createParagraph(`Available Nuggets:`);
    documentContent += createParagraph(`  • Conceptual: ${data.has_conceptual_nugget ? 'Yes' : 'No'}`, false, true);
    documentContent += createParagraph(`  • Practitioner: ${data.has_practitioner_nugget ? 'Yes' : 'No'}`, false, true);
    documentContent += createParagraph(`  • Builder: ${data.has_builder_nugget ? 'Yes' : 'No'}`, false, true);
    documentContent += pageBreak();

    // SYNTHESIS SECTION
    if (data.synthesis) {
      documentContent += createHeading1('Synthesis');
      documentContent += formatSection(data.synthesis);
      documentContent += pageBreak();
    }

    // CRITICAL ANALYSIS SECTION
    if (data.critical) {
      documentContent += createHeading1('Critical Archaeological Analysis');
      documentContent += formatSection(data.critical);
      documentContent += pageBreak();
    }

    // PROSPECTOR SECTION
    if (data.prospector) {
      documentContent += createHeading1('Strategic Prospector Analysis');
      documentContent += formatSection(data.prospector);
      documentContent += pageBreak();
    }

    // ORIGINAL TEXT EXCERPT
    if (data.fulltext) {
      documentContent += createHeading1('Original Paper Excerpt');
      documentContent += createParagraph(data.fulltext);
    }

    // Close document
    documentContent += `
  </w:body>
</w:document>`;

    wordFolder.file('document.xml', documentContent);

    // Generate ZIP as buffer
    const content = await zip.generateAsync({ type: 'nodebuffer' });

    // Send as DOCX file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${data.filename || 'analysis_report.docx'}"`);
    res.send(content);

  } catch (error) {
    console.error('Error generating DOCX:', error);
    res.status(500).json({ error: 'Failed to generate document', details: error.message });
  }
}

// Helper function to escape XML special characters
function escapeXml(unsafe) {
  if (!unsafe) return '';
  return String(unsafe)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// Create different paragraph types
function createParagraph(text, bold = false, indent = false) {
  const escapedText = escapeXml(text);
  return `
    <w:p>
      ${indent ? '<w:pPr><w:ind w:left="720"/></w:pPr>' : ''}
      <w:r>
        ${bold ? '<w:rPr><w:b/></w:rPr>' : ''}
        <w:t xml:space="preserve">${escapedText}</w:t>
      </w:r>
    </w:p>`;
}

// Create title (large, bold, centered)
function createTitle(text) {
  const escapedText = escapeXml(text);
  return `
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:before="240" w:after="240"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="32"/>
        </w:rPr>
        <w:t>${escapedText}</w:t>
      </w:r>
    </w:p>`;
}

// Create Heading 1
function createHeading1(text) {
  const escapedText = escapeXml(text);
  return `
    <w:p>
      <w:pPr>
        <w:spacing w:before="240" w:after="120"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="28"/>
        </w:rPr>
        <w:t>${escapedText}</w:t>
      </w:r>
    </w:p>`;
}

// Page break
function pageBreak() {
  return `
    <w:p>
      <w:r>
        <w:br w:type="page"/>
      </w:r>
    </w:p>`;
}

// Format sections with line breaks preserved
function formatSection(text) {
  if (!text) return '';
  
  // Split by double line breaks to create paragraphs
  const paragraphs = text.split(/\n\n+/);
  let formatted = '';
  
  for (const para of paragraphs) {
    // Check if this looks like a header (all caps followed by colon)
    if (/^[A-Z\s&\-\/]+:/.test(para)) {
      const [header, ...rest] = para.split(':');
      formatted += createParagraph(header + ':', true);
      if (rest.length > 0) {
        formatted += createParagraph(rest.join(':').trim());
      }
    } else {
      formatted += createParagraph(para);
    }
  }
  
  return formatted;
}
