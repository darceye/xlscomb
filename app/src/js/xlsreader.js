const XLSX = require('xlsx')
const utils = XLSX.utils
const path = require('path')
const ipc = require('electron').ipcRenderer
/*

          filesData: this.filesData,
          filePassword: this.filePassword,
          sheetName: this.sheetName,
          outputData: this.outputData,
          outputFile: this.outputFile,
          changeProgress: function(newProgress),
          changeInfo: function(newInfo),

*/

var gOutputs = []
var gFiles = []
var gTotalFiles = 0, gDoneFiles = 0
var gCurrentFile = ''
var gP
addEventListener('readDone', function(e){
    console.log('readDone')
    readFileAsync(gP, gFiles.pop())
})

function combine(p){

    var outputs = readFiles(p)
    writeOutputs(p, outputs)
}



function writeOutputs(p, outputs){
    var wb = createWb(p, outputs)
    try{
      XLSX.writeFile(wb, p.outputFile)

    }catch(e){
        console.log('Write Error!', e)
        p.changeError(e.toString())
        p.changeInfo(e.toString())
        p.finished()
        return
    }
    console.log('wb writed to ' + p.outputFile)
    p.changeInfo('文件已写入 ' + p.outputFile)
    p.changeProgress(100)
    p.finished()

}


function combineAsync(p){
    // create gFiles Queue.
    gFiles = []
    var totalFiles = p.filesData.length
    for(var i = 0; i < p.filesData.length; i+=1){
        gFiles.push(p.filesData[i].name)
    }

    // prepare gOutputs and others
    gOutputs = []
    gTotalFiles = totalFiles
    gDoneFiles = 0
    gCurrentFile = ''
    gP = p


    // start processing files 
    // readFileAsync.on('done', readFileAsync(p, gFiles.pop()) )
    readFileAsync(p, gFiles.pop())
}

function processAsync(p){

    var process = parseInt(gDoneFiles/gTotalFiles*100)
    console.log('Processing:', gCurrentFile)
    p.changeProgress(process)
    p.changeInfo('正在处理:'+ gCurrentFile) 
}

function readFileAsync(p, filename){
    if(typeof(filename) !== 'undefined'){
        gCurrentFile = filename
        processAsync(p)
        setTimeout( ()=>{
            var wb
            try{
                wb = XLSX.readFile(filename, {password: p.filePassword})
            }catch(e){
                console.error('ERROR! Wrong Password', e)
                p.changeInfo(e.toString())
                p.finished()
            }
            var output = getCells(wb.Sheets[p.sheetName], p.outputData, p)
            output.unshift({
                v: path.basename(filename)
            })        
            gOutputs.push(output)
            gDoneFiles+=1;
            dispatchEvent(new Event('readDone'))

        },500)

    }else{
        // no more files to process
        writeOutputs(p, gOutputs)
    }
           
}

function readFiles(p){
    // output a 2-dimension array.
    var outputs = []
    var totalFiles = p.filesData.length
    for(var i = 0; i < p.filesData.length; i+=1){
        var filename = p.filesData[i].name

        var process = parseInt(i/totalFiles*100)
        console.log('Processing:', filename)
        p.changeProgress(process)
        p.changeInfo('正在处理:'+ filename)   
        vm.info = 'Processing:'+ filename

        var wb
        try{
            wb = XLSX.readFile(filename, {password: this.filePassword})
        }catch(e){
            console.error('ERROR! Wrong Password', e)
            p.changeInfo(e.toString())
            continue
        }
        console.log('Readed: ', filename)
        p.changeInfo('已处理: '+ filename)
        var output = getCells(wb.Sheets[p.sheetName], p.outputData)
        output.unshift({
            v: path.basename(filename)
        })        
        outputs.push(output)
    }
    return outputs
}

function getCells(sheet, data, p){
    if (typeof(sheet) === 'undefined'){
        console.log('sheet is not exist')
        p.changeError(p.sheetName + '工作表不存在! ')
        return []
    }
    var output = []
    for(var i = 0; i < data.length; i+=1){
        var cellname = data[i].cell
        var cell
        if(typeof(sheet[cellname]) !== 'undefined'){
            cell = {
                v: sheet[cellname].v,
                t: sheet[cellname].t
            }
        }else{
            cell = {
                v: '',
                t: 's'
            }
        }
        output.push(cell)
    }
    return output
}

function createWb(p, outputs){
    outputs.unshift(createTitle(p))
    var range = utils.encode_range({
        s: {c: 0, r: 0},
        e: {c: p.outputData.length + 1, r: p.filesData.length + 1}
    })
    var wb = {
        SheetNames: ['Sheet1'],
        Sheets: {
            Sheet1: {
                '!ref': range
            }
        }
    }
    var sheet = wb.Sheets.Sheet1

    for(var r = 0; r < outputs.length; r+=1){
        var out = outputs[r]
        for(var c = 0; c < out.length; c+=1){
            var cell = utils.encode_cell({c: c,r: r})
            sheet[cell] = out[c]
        }
    }

    return wb
}

function createTitle(p){
    var title = [{v:'文件名'}]
    for(var i = 0; i < p.outputData.length; i++){
        title.push({
            v: p.outputData[i].name
        })
    }
    return title
}

