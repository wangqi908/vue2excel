<template>
  <button @click="exportExcel">导出</button>
  <button @click="exportExcelWithStyle">导出(带样式)</button>
</template>

<script>
import { useExportExcel, useExportExcelWithStyle } from '../hooks'
export default {
  name: 'ExportExcel',
  setup () {
    const { list } = require('./page.json')
    function exportExcel () {
      useExportExcel(list)
    }
    function exportExcelWithStyle () {
      useExportExcelWithStyle(list)
    }
    return {
      exportExcel,
      exportExcelWithStyle
    }
  }
}
</script>
