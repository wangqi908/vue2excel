<template>
  <div class="upload_excel">
    <span>上传Excel</span>
    <input type="file" ref="upload" @change="readExcel" accept=".xls,.xlsx" class="outputlist_upload">
  </div>
</template>

<script>
import XLSX from "xlsx";
export default {
  methods: {
    readExcel(e) {
      const files = e.target.files;
      if (files.length <= 0) {
        return false;
      } else if (!/\.(xls|xlsx)$/.test(files[0].name.toLowerCase())) {
        this.$Message.error("上传格式不正确，请上传xls或者xlsx格式");
        return false;
      }

      const fileReader = new FileReader();
      fileReader.onload = ev => {
        try {
          const data = ev.target.result;
          const workbook = XLSX.read(data, {
            type: "binary"
          });
          const wsname = workbook.SheetNames[0]; //取第一张表
          const ws = XLSX.utils.sheet_to_json(workbook.Sheets[wsname]); //生成json表格内容
          let outputs = []; //清空接收数据
          ws.forEach((ele, index) => {
            let obj = {
              pipeNo: ele["管点编号"],
              geophysicalNo: ele["物探点号"],
              joinNo: ele["连接点号"],
              featurePoints: ele["特征点"],
              pipeType: ele["管点类型"],
              xCoordinate: ele["X坐标"],
              yCoordinate: ele["Y坐标"],
              groundElevation: ele["地面高程"],
              pipeElevation: ele["管道高程"],
              depth: ele["管径"],
              pipeDiameter: ele["材质"],
              pipeMetal: ele["埋深"],
              ownerUnit: ele["权属单位"],
              buriedType: ele["埋设方式"],
              projectName: ele["工程名称"],
              projectAddress: ele["工程地址"],
              equipmentMode: ele["设备型号"],
              equipmentCode: ele["设备识别"],
              designUnit: ele["设计单位"],
              measurementUnit: ele["测量单位"],
              buriedDate: ele["埋设日期"],
              constructionDate: ele["施工日期"],
              completionDate: ele["竣工日期"]
            };
            for (const key in obj) {
              const element = obj[key];
              if (element === undefined) {
                obj[key] = "";
              }
            }
            let isObjValueAllEmpty = Object.values(obj).every(v => v === ""); //判断对象的值全为空
            if (!isObjValueAllEmpty) outputs.push(obj);
          });
          console.log(outputs);
          this.$emit("read-success", outputs);
          this.$refs.upload.value = "";
        } catch (e) {
          console.log(e);
          return false;
        }
      };
      fileReader.readAsBinaryString(files[0]);
    }
  }
};
</script>

<style >
.upload_excel {
  position: relative;
  display: inline-block;
  background: #f5a623;
  border: 1px solid #f5a623;
  border-radius: 4px;
  padding: 8px 12px;
  overflow: hidden;
  color: #fff;
  font-size: 14px;
  margin: 10px;
}

input {
  position: absolute;
  font-size: 100px;
  right: 0;
  top: 0;
  opacity: 0;
}
.upload_excel:hover {
  background: #f7c16a;
  border-color: #f5cb87;
  color: #fff;
}
.upload_excel input:hover {
  cursor: pointer;
}
</style>