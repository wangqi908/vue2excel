<template>
  <div id="app">
    <my-download-excel />
    <!-- 请上传static目录下的测试文件.xlsx -->
    <my-read-excel @read-success="readSuccess" />
    <ul>
      <li v-for="(item, index) in list" :key="index">
        <span>工程名称---</span> <span>{{item.projectName}}</span>
        <span>工程地址---</span> <span>{{item.projectAddress}}</span>
        <span>X坐标---</span> <span>{{item.xCoordinate}}</span>
      </li>
    </ul>
  </div>
</template>

<script>
import { MyDownloadExcel, MyReadExcel } from "@/components";
export default {
  name: "App",
  components: { MyDownloadExcel, MyReadExcel },
  data() {
    return {
      list: []
    };
  },
  methods: {
    // 读取成功
    readSuccess(list) {
      this.list = list;
    }
  }
};
</script>
