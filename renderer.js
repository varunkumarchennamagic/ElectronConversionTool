const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
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

let imageCounter = 1; // Initialize the counter for image naming

showAlertBtn.addEventListener('click', async () => {
  let selectedFilePath = selectedFilePathElement.textContent;
  if (selectedFilePath) {
    const result = await convertDocxToHtml(selectedFilePath);
    console.log('Converted HTML:', result);

    const { doc, jsonOutput } = generateJsonFromHtml(result, selectedFilePath);
    console.log('Generated JSON:', jsonOutput);

    const outputFolderPath = path.dirname(selectedFilePath);
    const jsonFilePath = path.join(outputFolderPath, "output.json");
    fs.writeFileSync(jsonFilePath, JSON.stringify(jsonOutput));

    const htmlFilePath = path.join(outputFolderPath, "output.html");
    fs.writeFileSync(htmlFilePath, result); // Write the HTML to output.html

    // Save images in the 'images' folder
    const imagesFolderPath = path.join(outputFolderPath, "images");
    fs.mkdirSync(imagesFolderPath, { recursive: true });

    // Find images in the HTML
    const imageTags = doc.getElementsByTagName("img");
    for (const imageTag of imageTags) {
      const base64Data = imageTag.getAttribute("src");
      const imageExtension = base64Data.substring(base64Data.indexOf("/") + 1, base64Data.indexOf(";base64"));
      const imageName = `image${imageCounter}.${imageExtension}`;
      const imagePath = path.join(imagesFolderPath, imageName);
      fs.writeFileSync(imagePath, base64Data.replace(/^data:image\/(png|jpeg|jpg);base64,/, ""), "base64");
      imageCounter++;
      imageTag.setAttribute("src", `images/${imageName}`);
    }

    alert('HTML and JSON files generated.');

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
  return { doc, jsonOutput }; // Return both the doc and the jsonOutput
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
    const imageFileName = 'image_' + Date.now() + '.png';
    jsonNode.attributes.src = 'images/' + imageFileName;
  }

  return jsonNode;
};
