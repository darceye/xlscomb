const electron = require('electron')
// Module to control application life.
const app = electron.app
// Module to create native browser window.
const BrowserWindow = electron.BrowserWindow

const path = require('path')
const url = require('url')

// 菜单栏
const Menu = electron.Menu

let menuTemplate = []


menuTemplate = [
//   {
//     label: '文稿',
//     submenu: [{
//         label: '打开',
//         accelerator: 'CmdOrCtrl + O',
//         role: 'open',
//         click: function(item, focusedWindow) {
//             dialog.showOpenDialog({
//                 properties: ['openFile', 'openDirectory']
//             }, function(files) {
//                 if (files) {
//                     mainWindow.send('selected-directory', files)
//                 }
//             })
//         }
//     }, {
//         label: '新建',
//         accelerator: 'CmdOrCtrl + N',
//         role: 'new',
//         click: function(item, focusedWindow){
//             mainWindow.send('new-doc')
//         }
//     }, {
//         label: '历史文稿',
//         role: 'history',
//         click: function(item, focusedWindow) {
//             mainWindow.send('goto-guide')
//         }
//     }, {
//         type: 'separator',
//     }, {
//         label: '退出',
//         role: 'quit',
//     }]
// },  {
//     label: '编辑',
//     submenu: [
//       {
//         label: '撤销',
//         role: 'undo'
//       },
//       {
//         label: '重做',
//         role: 'redo'
//       },
//       {
//         type: 'separator'
//       },
//       {
//         label: '剪切',
//         role: 'cut'
//       },
//       {
//         label: '复制',
//         role: 'copy'
//       },
//       {
//         label: '粘贴',
//         role: 'paste'
//       },
//       {
//         label: '全选',
//         role: 'selectall'
//       }
//     ]
//   },
//   {
//     label: '查看',
//     submenu: [
//       {
//         label: '正常大小',
//         role: 'resetzoom'
//       },
//       {
//         label: '放大视图',
//         role: 'zoomin'
//       },
//       {
//         label: '缩小视图',
//         role: 'zoomout'
//       },
//       {
//         type: 'separator'
//       },
//       {
//         label: '全屏',
//         role: 'togglefullscreen'
//       }
//     ]
//   },
  {
    label: '帮助',
    role: 'help',
    submenu: [{
        label: '调试',
        role: 'toggledevtools',
        click: function(item, focusedWindow) {
            mainWindow.webContents.openDevTools()
        }
    },{
        label: '关于',
        click: function(item, focusedWindow) {
            if (focusedWindow) {
                const options = {
                    type: 'info',
                    title: '关于Excel合并器',
                    buttons: ['好'],
                    message: `Excel合并器由 蜻蜓的企鹅 友情开发`,
                    detail: '版本: 1.0\nCopyright ©Jacques Yang 2017 '
                }
                electron.dialog.showMessageBox(focusedWindow, options, function() {})
            }
        }
    },{
        label: '企鹅站点',
        click () { require('electron').shell.openExternal('http://idarc.cn') }
    }]
  }
]

// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
let mainWindow

function createWindow () {

  const menu = Menu.buildFromTemplate(menuTemplate)
  Menu.setApplicationMenu(menu)
  // Create the browser window.
  mainWindow = new BrowserWindow({width: 800, height: 600,useContentSize: true, fullscreenable: true})
  mainWindow.maximize()

  // and load the index.html of the app.
  mainWindow.loadURL(url.format({
    pathname: path.join(__dirname, 'index.html'),
    protocol: 'file:',
    slashes: true
  }))

  // Open the DevTools.
  mainWindow.webContents.openDevTools()

  // Emitted when the window is closed.
  mainWindow.on('closed', function () {
    // Dereference the window object, usually you would store windows
    // in an array if your app supports multi windows, this is the time
    // when you should delete the corresponding element.
    mainWindow = null
  })
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', createWindow)

// Quit when all windows are closed.
app.on('window-all-closed', function () {
  // On OS X it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', function () {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (mainWindow === null) {
    createWindow()
  }
})

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.
