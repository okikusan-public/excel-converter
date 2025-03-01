import * as fs from 'fs';
import * as path from 'path';

const outputDir = path.join(__dirname, '..', 'output');
const mappingFilePath = path.join(outputDir, 'filename_mapping.json');

// Read the mapping file
fs.readFile(mappingFilePath, 'utf8', (err, data) => {
  if (err) {
    console.error('Error reading mapping file:', err);
    return;
  }

  try {
    const mappings = JSON.parse(data);

    // Iterate through the mappings and rename files
    for (const generatedName in mappings) {
      if (mappings.hasOwnProperty(generatedName)) {
        const originalName = mappings[generatedName].originalName;
        const oldPath = path.join(outputDir, generatedName);
        const newPath = path.join(outputDir, `${originalName}.pdf`);

        // Check if the file exists before renaming
        if (fs.existsSync(oldPath)) {
          fs.rename(oldPath, newPath, (renameErr) => {
            if (renameErr) {
              console.error(`Error renaming file ${generatedName}:`, renameErr);
            } else {
              console.log(`Renamed ${generatedName} to ${originalName}.pdf`);
            }
          });
        } else {
          console.warn(`File not found, skipping rename: ${generatedName}`);
        }
      }
    }
  } catch (parseError) {
    console.error('Error parsing mapping file:', parseError);
  }
});
