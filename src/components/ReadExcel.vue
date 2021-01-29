<template>
  <input type="file" @change="inpChange" />
  {{ list }}
</template>

<script>
import { reactive, toRefs } from 'vue'
import { readExcel } from '@/utils/excel'
export default {
  name: 'ReadExcel',
  setup () {
    const state = reactive({
      list: []
    })
    async function inpChange (e) {
      const { files } = e.target
      const list = await readExcel(files[0], { range: 2 })
      state.list = list
    }

    return {
      inpChange,
      ...toRefs(state)
    }
  }
}
</script>
