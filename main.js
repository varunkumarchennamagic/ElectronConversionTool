const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');

let mainWindow;
let selectedFilePath = null;
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true, // Enable Node.js integration
      contextIsolation: false // Disable context isolation
    //   preload: path.join(__dirname, 'preload.js')
    }
  });
  mainWindow.webContents.openDevTools()
  mainWindow.loadFile('index.html');

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.on('ready', createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow();
  }
});

ipcMain.on('open-file-dialog', (event) => {
  dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'Word Documents', extensions: ['docx'] }
    ]
  }).then(result => {
    const filePaths = result.filePaths;
    if (filePaths && filePaths.length > 0) {
      selectedFilePath = filePaths[0];
      console.log(filePaths[0]);
      event.reply('selected-file', selectedFilePath);
    }
  });
});

ipcMain.on('get-selected-file', (event) => {
    event.returnValue = selectedFilePath;
});
