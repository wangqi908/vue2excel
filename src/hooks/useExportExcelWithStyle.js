import { getWb, exportJsonToExcel } from '@/utils/excel'

export default function useExportExcelWithStyle (data) {
  // https://www.npmjs.com/package/xlsx-style
  const list = JSON.parse(JSON.stringify(data))
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
