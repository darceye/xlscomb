'use strict'

const NO_OUTPUT_FILE = '未选择任何文件'
const DEFAULT_OUTPUT_FILENAME = 'C:\\output.xlsx'
const DEC_FILE_FOLDER = '__decrypted'
const DEFAULT_SHEET = '分析表'
const DEFAULT_PASSWORD = ''
var defaultOutputData = [
          {
              name: '单次平均运行时间（分钟）',
              cell: 'I4',
          },
          {
              name: '定额单次',
              cell: 'L4'
          }
      ]

const fs = require('fs')
const shell = require('electron').shell
const iconv = require('iconv-lite')

const DOCS_PATH = __dirname + '\\..\\docs\\'

var settings = fs.readFileSync(DOCS_PATH + 'settings.json', {
    encoding: 'utf-8'
  })

if(settings === ''){
  settings = {
    sheetName: DEFAULT_SHEET,
    outputFile: DEFAULT_OUTPUT_FILENAME,
    outputData: defaultOutputData,
    filePassword: DEFAULT_PASSWORD
  }
}else{
  settings = JSON.parse(settings)
}

var vm = new Vue({
    el:'#app',
    data:{
      jsheets: [], // created by xlsx
      nxwb: [], // created by nx
      encFiles:[],
      files: [],
      filesTitles: [{
          title: '文件名', key: 'name'
        }, {
          title: '操作', key: 'action', width: 80, align: 'center',
          render: (h, params) => {
              return h('div', [
                  h('Button', {
                      props: {
                          type: 'error',
                          size: 'small'
                      },
                      on: {
                          click: () => {
                              this.vm.removeFilesItem(params.index)
                          }
                      }
                  }, '删除')
              ]);
          }

      }],
      encFilesTitles: [{
          title: '文件名', key: 'name'
        }, {
          title: '操作', key: 'action', width: 80, align: 'center',
          render: (h, params) => {
              return h('div', [
                  h('Button', {
                      props: {
                          type: 'error',
                          size: 'small'
                      },
                      on: {
                          click: () => {
                              this.vm.removeEncFilesItem(params.index)
                          }
                      }
                  }, '删除')
              ]);
          }

      }],
      filesData: [],
      encFilesData: [],
      filePassword: settings.filePassword,
      passwordType: 'password',
      sheetName: settings.sheetName,
      newOutput: {
        name: "",
        cell: ""
      },
      outputFile: settings.outputFile,
      isAddOutputData: false,
      outputTitles: [
                    {
                        title: '列名',
                        key: 'name'
                    },
                    {
                        title: '单元格',
                        key: 'cell'
                    },
                    {
                        title: '操作',
                        key: 'action',
                        width: 80,
                        align: 'center',
                        render: (h, params) => {
                            return h('div', [
                                h('Button', {
                                    props: {
                                        type: 'error',
                                        size: 'small'
                                    },
                                    on: {
                                        click: () => {
                                            this.vm.removeOutputTableItem(params.index)
                                        }
                                    }
                                }, '删除')
                            ]);
                        }
                    }
                ],
      outputData: settings.outputData,
      error: '',
      info: '',
      isCombining: false,
      isShowCombineProgress: false,
      combineProgress: 0

    },
    methods: {
      decrypt: function(){
        console.log('正在打开解密xlsm')
        this.saveEncFilelist()
        this.savePassword()
        shell.openItem(path.join(DOCS_PATH , 'decrypter.xlsm'))
      },
      loadDecrypted: function(){
        var decfiles = fs.readdirSync(path.join(DOCS_PATH, DEC_FILE_FOLDER))
        console.log(decfiles)
        for(var i = 0; i < decfiles.length; i+=1){
          var ext = path.extname(decfiles[i]).toLowerCase()
          if(ext === '.xls' || ext === '.xlsx' || ext === '.xlsm'){
            this.filesData.push({
              name: path.join(DOCS_PATH, DEC_FILE_FOLDER,decfiles[i])
            })
          }
        }

        this.filesData = _.uniqBy(this.filesData, 'name')
      },
      openEncFiles: function (){
        var files = openFiles();
        this.encFiles.splice.apply(this.encFiles, [0, this.encFiles.length].concat(files))
        for(var i = 0; i < this.encFiles.length; i+=1){
          this.encFilesData.push({
            name: this.encFiles[i]
          })
        }
        this.encFilesData = _.uniqBy(this.encFilesData, 'name')
        console.log(files)
      },
      openFiles: function (){
        var files = openFiles();
        this.files.splice.apply(this.files, [0, this.files.length].concat(files))
        for(var i = 0; i < this.files.length; i+=1){
          this.filesData.push({
            name: this.files[i]
          })
        }
        this.filesData = _.uniqBy(this.filesData, 'name')
        console.log(files)
      },
      deleteFile: function(i){
        this.files.splice(i,1)
      },
      saveEncFilelist(){
        var str = ''
        for(var i = 0; i < this.encFilesData.length; i+=1){
          str += this.encFilesData[i].name
          str += '\r\n'
        }
        fs.writeFileSync(path.join(DOCS_PATH, 'filelist.xccfg'), iconv.encode(str, 'gbk'))
        
      },
      savePassword(){
        fs.writeFileSync(path.join(DOCS_PATH, 'password.xccfg'), this.filePassword + '\r\n')
        
      },
      togglePasswordView: function(){
        if(this.passwordType === 'password'){
          this.passwordType = 'text'
        }else{
          this.passwordType = 'password'
        }
      },

      openOutputFile: function (){
        var file = openOutputFile();
        this.outputFile = file
        console.log(file);
      },
      showFile:function(){
        this.nxwb = nxParse(this.files[0])
        this.jsheets = xlsxParse(this.files[0])
      },
      removeEncFilesItem: function (index) {
          this.encFilesData.splice(index, 1);
      },
      removeFilesItem: function (index) {
          this.filesData.splice(index, 1);
      },
      saveSettings:function () {

        fs.writeFileSync(path.join(__dirname,'settings.json'), JSON.stringify({
          sheetName: this.sheetName,
          outputFile: this.outputFile,
          outputData: this.outputData,
          filePassword: this.filePassword
        }))
      },
      addOutputData: function () {
        if (this.newOutput.name === '' || this.newOutput.cell === ''){
          this.showAddOutputError()

          return
        }
        this.outputData.push({
          name: this.newOutput.name,
          cell: this.newOutput.cell
        })
      },
      removeOutputTableItem: function (index) {
          this.outputData.splice(index, 1);
      },
      showAddOutput: function () {
        this.isAddOutputData = true
      },

      showAddOutputError: function () {
        this.$Message.error('输出列名称和单元格引用都必须填写');
      },

      combine: function () {
        var that = this
        this.isCombining = true
        this.info = ''
        this.error = ''
        combineAsync({
          filesData: this.filesData,
          filePassword: this.filePassword,
          sheetName: this.sheetName,
          outputData: this.outputData,
          outputFile: this.outputFile,
          changeProgress: (newProgress)=>{
            console.log('new: ' + newProgress, 'old: ',that.combineProgress)
            that.combineProgress = newProgress
          },
          changeInfo: (newInfo)=>{
            that.info = newInfo            
          },
          changeError: (newError)=>{
            that.error += newError
          },
          finished: ()=>{
            that.isCombining = false
          }
        })
      },
      viewCombined: function(){
        shell.openItem(this.outputFile)

      }
    }
})

