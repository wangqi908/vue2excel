import { exportJsonToExcel } from '@/utils/excel'

export default function useExportExcel (data) {
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
  exportJsonToExcel(params)
}
