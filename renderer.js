const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const admZip = require('adm-zip');
const outputPath = "D:/Savvas/ElectronConversionTool/output.html";
const selectFileBtn = document.getElementById('selectFileBtn');
const selectedFilePathElement = document.getElementById('selectedFilePath');
const showAlertBtn = document.getElementById('showAlertBtn');

selectFileBtn.addEventListener('click', () => {
  ipcRenderer.send('open-file-dialog');
});

ipcRenderer.on('selected-file', (event, filePath) => {
  // Update the selected file path element
  selectedFilePathElement.textContent = filePath;
});

showAlertBtn.addEventListener('click', async () => {
  let selectedFilePath = selectedFilePathElement.textContent;
  if (selectedFilePath) {
    const result = await convertDocxToHtml(selectedFilePath);
    console.log('Converted HTML:', result);

    const jsonOutput = generateJsonFromHtml(result, selectedFilePath);
    console.log('Generated JSON:', jsonOutput);

    const jsonFilePath = "D:/Savvas/ElectronConversionTool/output.json";
    fs.writeFileSync(jsonFilePath, JSON.stringify(jsonOutput));

    alert('JSON file generated.');

  } else {
    alert('No file selected.');
  }
});

ipcRenderer.on('send-selected-file', async (event, filePath) => {
  try {
    // Convert .docx file to HTML
    const result = await convertDocxToHtml(filePath);
    console.log('Converted HTML:', result);

    const jsonOutput = generateJsonFromHtml(result, filePath);
    console.log('Generated JSON:', jsonOutput);

    const jsonFilePath = "D:/Savvas/ElectronConversionTool/output.json";
    fs.writeFileSync(jsonFilePath, JSON.stringify(jsonOutput)); // Save the JSON to the output file
    // Do something with the generated JSON
    // You can display it, process it further, etc.

    // var bodyTag = jsonOutput.children[1]
  } catch (error) {
    console.error('Error converting to HTML:', error);
  }
});

const convertDocxToHtml = (filePath) => {
  return new Promise((resolve, reject) => {
    mammoth.convertToHtml({ path: filePath })
      .then((result) => {
        const html = result.value; // The generated HTML
        resolve(html);
      })
      .catch((error) => {
        console.error("An error occurred:", error);
        reject(error);
      });
  });
};

const generateJsonFromHtml = (html, docxFilePath) => {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  const imagesDir = path.dirname(docxFilePath) + "/images";
  if (!fs.existsSync(imagesDir)) {
    fs.mkdirSync(imagesDir);
  }

  const jsonOutput = traverseDOM(doc.documentElement, imagesDir);
  return jsonOutput;
};

const traverseDOM = (element, imagesDir) => {
  const jsonNode = {
    tag: element.tagName.toLowerCase(),
    attributes: {},
    children: [],
  };

  for (const attr of element.attributes) {
    jsonNode.attributes[attr.name] = attr.value;
  }

  for (const childNode of element.childNodes) {
    if (childNode.nodeType === Node.ELEMENT_NODE) {
      const childJsonNode = traverseDOM(childNode, imagesDir);
      jsonNode.children.push(childJsonNode);
    } else if (childNode.nodeType === Node.TEXT_NODE) {
      const text = childNode.textContent.trim();
      if (text.length > 0) {
        jsonNode.children.push(text);
      }
    }
  }

  if (jsonNode.tag === 'img') {
    const base64Data = jsonNode.attributes.src.replace(/^data:image\/png;base64,/, "");
    const imageFileName = 'image_' + Date.now() + '.png';
    const imagePath = path.join(imagesDir, imageFileName);
    fs.writeFileSync(imagePath, base64Data, 'base64');
    jsonNode.attributes.src = 'images/' + imageFileName;
  }

  return jsonNode;
};
