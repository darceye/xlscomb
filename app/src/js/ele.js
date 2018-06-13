const remote = require('electron').remote

// for electron op.
function openFiles(){
  var files = remote.dialog.showOpenDialog({
    properties:[
      'openFile', 'multiSelections'],
    filters:[{name: 'Excel文件', extensions: ['xlsx', 'xls']}],
    })

  return files
}

function openOutputFile(){
  var files = remote.dialog.showSaveDialog({
    filters:[{name: 'Excel文件', extensions: ['xlsx']}],
    title: '合并后的文件'
    })

  return files
}
