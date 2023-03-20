const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const XLSX = require('xlsx');
const path = require('path');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
        }
    });

    mainWindow.loadFile('index.html');

    mainWindow.on('closed', function () {
        mainWindow = null;
    });
}

function readXLSX(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    console.log(data);
}

app.whenReady().then(() => {
    createWindow();

    ipcMain.on('open-file-dialog', (event) => {
        dialog.showOpenDialog({
            properties: ['openFile'],
            filters: [
                { name: 'XLSX Files', extensions: ['xlsx'] }
            ]
        }).then(result => {
            if (!result.canceled && result.filePaths.length > 0) {
                event.reply('selected-file', result.filePaths[0]);
            }
        });
    });
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});
