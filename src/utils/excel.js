import { saveAs } from 'file-saver'
import xlsxStyle from 'xlsx-style'
import xlsx from 'xlsx'
// https://www.npmjs.com/package/xlsx-style

function dateNum (v, date1904) {
  if (date1904) v += 1462
  var epoch = Date.parse(v)
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000)
}

function sheetFromArrayOfArrays (data, opts) {
  var ws = {}
  var range = {
    s: {
      c: 10000000,
      r: 10000000
    },
    e: {
      c: 0,
      r: 0
    }
  }
  for (var R = 0; R !== data.length; ++R) {
    for (var C = 0; C !== data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R
      if (range.s.c > C) range.s.c = C
      if (range.e.r < R) range.e.r = R
      if (range.e.c < C) range.e.c = C
      var cell = {
        v: data[R][C]
      }
      if (cell.v == null) continue
      var cellRef = xlsxStyle.utils.encode_cell({
        c: C,
        r: R
      })

      if (typeof cell.v === 'number') cell.t = 'n'
      else if (typeof cell.v === 'boolean') cell.t = 'b'
      else if (cell.v instanceof Date) {
        cell.t = 'n'
        cell.z = xlsxStyle.SSF._table[14]
        cell.v = dateNum(cell.v)
      } else cell.t = 's'

      ws[cellRef] = cell
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = xlsxStyle.utils.encode_range(range)
  return ws
}

function Workbook () {
  if (!(this instanceof Workbook)) return new Workbook()
  this.SheetNames = []
  this.Sheets = {}
}

function s2ab (s) {
  var buf = new ArrayBuffer(s.length)
  var view = new Uint8Array(buf)
  for (var i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff
  return buf
}

function autoWidthFunc (ws, data) {
  const colWidth = data.map(row =>
    row.map(val => {
      if (val === null) {
        return { wch: 10 }
      } else {
        return { wch: val.toString().length * 2 }
      }
    })
  )
  const result = colWidth[0]
  for (let i = 1; i < colWidth.length; i++) {
    for (let j = 0; j < colWidth[i].length; j++) {
      if (result[j].wch < colWidth[i][j].wch) {
        result[j].wch = colWidth[i][j].wch
      }
    }
  }
  ws['!cols'] = result
}

function jsonToArray (key, jsonData) {
  return jsonData.map(v =>
    key.map(j => {
      return v[j]
    })
  )
}

export const getWb = ({
  key,
  list,
  title,
  autoWidth = true,
  multiHeader = [],
  merges = []
}) => {
  const data = jsonToArray(key, list)
  data.unshift(title)

  for (let i = multiHeader.length - 1; i > -1; i--) {
    data.unshift(multiHeader[i])
  }
  const wsName = 'SheetJS'
  const wb = new Workbook()
  const ws = sheetFromArrayOfArrays(data)

  if (merges.length > 0) {
    if (!ws['!merges']) ws['!merges'] = []
    merges.forEach(item => {
      ws['!merges'].push(xlsxStyle.utils.decode_range(item))
    })
  }

  if (merges.length > 0) {
    if (!ws['!merges']) ws['!merges'] = []
    merges.forEach(item => {
      ws['!merges'].push(xlsxStyle.utils.decode_range(item))
    })
  }

  if (autoWidth) {
    autoWidthFunc(ws, data)
  }
  wb.SheetNames.push(wsName)
  wb.Sheets[wsName] = ws

  return wb
}

export const writeExcel = ({
  wb,
  filename = 'excel-list',
  bookType = 'xlsx'
}) => {
  const wbOut = xlsxStyle.write(wb, {
    bookType: bookType,
    bookSST: false,
    type: 'binary'
  })
  saveAs(
    new Blob([s2ab(wbOut)], {
      type: 'application/octet-stream'
    }),
    `${filename}.${bookType}`
  )
}

/**
 * @method 生成Excel文件
 * @param params {}
      key Array ['name', 'address', 'age']
      title Array  ['姓名', '地址', '年龄']
      list Array 数据
      autoWidth = true 宽度自适应
      multiHeader = []  [['统计表', '', '']] 多表头
      merges = []  ['A1:C1'] 合并单元格
      wb Object wb
      filename = 'excel-list' 文件名
      bookType = 'xlsx' 文件类型

 * @example
      exportExcelWithStyle() {
        // https://www.npmjs.com/package/xlsx-style
        const list = JSON.parse(JSON.stringify(this.list))
        list.push({
          age: '1122',
          name: '合计',
          address: ''
        })

        const multiHeader = [['统计表', '', '']]
        const merges = ['A1:C1']
        const title = ['姓名', '地址', '年龄']
        const key = ['name', 'address', 'age']
        const params = {
          multiHeader,
          merges,
          title,
          key,
          list,
          autoWidth: true,
          filename: '清单'
        }
        const wb = getWb(params)
        const dataInfo = wb.Sheets[wb.SheetNames[0]]

        dataInfo.A1.s = {
          font: {
            name: '微软雅黑',
            sz: 14,
            color: {
              rgb: 'CC66FF'
            },
            bold: true,
            italic: false,
            underline: false
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center'
          },
          fill: {
            fgColor: {
              rgb: 'FFFFCC'
            }
          }
        }

        dataInfo.A13.s = {
          border: {
            top: {
              style: 'thin'
            },
            bottom: {
              style: 'thin'
            },
            left: {
              style: 'thin'
            },
            right: {
              style: ''
            }
          },
          font: {
            name: '微软雅黑',
            sz: 11,
            bold: true
          },
          alignment: {
            horizontal: '',
            vertical: ''
          },
          fill: {
            fgColor: {
              rgb: 'f0f0f0'
            }
          }
        }
        dataInfo.B13.s = {
          border: {
            top: {
              style: 'thin'
            },
            bottom: {
              style: 'thin'
            },
            left: {
              style: ''
            },
            right: {
              style: ''
            }
          },
          font: {
            name: '微软雅黑',
            sz: 11,
            bold: true
          },
          alignment: {
            horizontal: '',
            vertical: ''
          },
          fill: {
            fgColor: {
              rgb: 'f0f0f0'
            }
          }
        }
        dataInfo.C13.s = {
          border: {
            top: {
              style: 'thin'
            },
            bottom: {
              style: 'thin'
            },
            left: {
              style: ''
            },
            right: {
              style: 'thin'
            }
          },
          font: {
            name: '微软雅黑',
            sz: 11,
            bold: true
          },
          alignment: {
            horizontal: '',
            vertical: ''
          },
          fill: {
            fgColor: {
              rgb: 'f0f0f0'
            }
          }
        }
        params.wb = wb
        exportJsonToExcel(params)
    }
 */
export const exportJsonToExcel = ({
  key,
  title,
  list,
  autoWidth = true,
  multiHeader = [],
  merges = [],
  wb,
  filename = 'excel-list',
  bookType = 'xlsx'
}) => {
  wb = wb || getWb({ key, list, title, autoWidth, multiHeader, merges })
  writeExcel({ wb, filename, bookType })
}

/**
 * @method 读取Excel文件
 * @param params {}
      file 文件
      Sheet2JSONOpts Object 配置项
        raw true Use raw values (true) or formatted strings (false)
        range from WS Override Range (see table below) range: 0 从第0行开始
        header Control output format (see table below)
        dateNF FMT 14 Use specified date format in string output
        defval Use specified value in place of null or undefined
        blankrows ** Include blank lines in the output **
 * @example
      readExcel(files[0], { range: 0 }).then(res => {
        console.log(res)
        state.list = res
      })
  @return Promise
 */
export const readExcel = (file, Sheet2JSONOpts = {}) => {
  const fileReader = new FileReader()
  return new Promise((resolve, reject) => {
    fileReader.onload = ev => {
      try {
        const data = ev.target.result
        const workbook = xlsx.read(data, {
          type: 'array'
        })
        const sheetNames = workbook.SheetNames
        const res = []
        sheetNames.forEach(name => {
          const ws = xlsx.utils.sheet_to_json(
            workbook.Sheets[name],
            Sheet2JSONOpts
          ) // 生成json表格内容
          if (ws.length !== 0) res.push(ws)
        })
        // resolve(res.flat(Infinity))
        resolve(res)
      } catch (e) {
        reject(e)
        return false
      }
    }
    fileReader.readAsArrayBuffer(file)
  })
}
