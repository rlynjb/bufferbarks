<template>
  <div>
    <p>To work on Microsoft Apps, we need the root path of the file</p>

    <el-select
      v-if="files"
      v-model="selectedFile"
      class="m-2" placeholder="Select" size="large">
      <el-option
        v-for="file in files.value" :key="file.id"
        :label="file.name"
        :value="file.id"
        @click="getFile(file.id)"
      />
    </el-select>

    <el-collapse v-if="worksheets" v-model="selectedSheet" accordion
      class="m-4">
      <el-collapse-item v-for="worksheet in worksheets.value" :key="worksheet.id"
        :title="worksheet.name" :name="worksheet.name">
        <div v-if="columns">
          <p>These fields are being pulled dynamically from an excel sheet table</p>
          <div v-for="field in columns.value" :key="field.id">
            <el-input
              v-model="columnValues[field.name]" :placeholder="'Please enter '+ field.name">
              <template #prepend>{{ field.name }}</template>
            </el-input>
          </div>

          <el-button @click="submit">Submit</el-button>
        </div>
      </el-collapse-item>
    </el-collapse>
  </div>
</template>

<script setup lang="ts">
import {
  ref,
  onMounted,
  watch,
  reactive,
  toRaw,
  Ref
} from 'vue';

import { useMsGraph } from '@/composition-api/useMsGraph';

const {
  getDriveFiles,
  getExcel,
  getTables,
  getColumns,
  postRow
} = useMsGraph();

// good read about ref vs reactive in vue3: https://www.danvega.dev/blog/2020/02/12/vue3-ref-vs-reactive/
// this is better: https://blog.deepgram.com/diving-into-vue-3-reactivity-api/
// https://chrysanthos.xyz/article/how-to-get-the-data-of-a-proxy-object-in-vue-js-3/
// something to consider: https://stackoverflow.com/questions/70339961/iterating-over-a-proxy-in-vue-composition-api
// and https://github.com/vuejs/rfcs/discussions/369
const files = ref();
const worksheets = ref();
const tables = ref();
const columns = ref();

const selectedFile = ref();
const selectedSheet = ref();

// good read for dynamically assigning properties to object
// ref: https://stackoverflow.com/questions/12710905/how-do-i-dynamically-assign-properties-to-an-object-in-typescript
interface ColumnValuesDictionary {
  [index: string]: string
}
const columnValues = ref<ColumnValuesDictionary>({});

watch(selectedSheet, (currentVal, oldVal) => {
  tables.value = [];
  columns.value = [];

  getTable(selectedFile.value, selectedSheet.value);
});

onMounted(async () => {
  files.value = await getDriveFiles();
});


async function getFile(id: string) {
  worksheets.value = await getExcel(id);
}

async function getTable(
  fileID: string,
  worksheetID: string
) {
  // getTable pass in fileID, worksheetID
  tables.value = await getTables(fileID, worksheetID);

  getFields(
    selectedFile.value,
    selectedSheet.value,
    tables.value.value[0].id
  );
}

async function getFields(
  fileID: string,
  worksheetID: string,
  tableID: string
) {
  // getColumns pass in fileID, worksheetID, tableID
  columns.value = await getColumns(fileID, worksheetID, tableID);

  // add empty object properties for field values
  columns.value.value.forEach((field: { name: string }) => {
    let label = field.name as string;
    columnValues.value[label] = '';
  });
}

async function submit() {
  const rawLabelObj: Array<any> = toRaw(columns.value.value);
  const labelsInOrder: Array<string> = rawLabelObj.map((c: { name:string }) => c.name);

  const rawValueObj: any = toRaw(columnValues.value);
  const valuesInOrder: Array<string> = labelsInOrder.map((v: string) => rawValueObj[v]);

  const payload: Array<string> = valuesInOrder;
  const rawTableID: string = toRaw(tables.value.value)[0].id;

  const res = await postRow(
    selectedFile.value,
    selectedSheet.value,
    rawTableID, // we'll assume there's always going to be only 1 table per sheet
    payload
  );
}
</script>

<style scoped>
ul {
  list-style: none;
}
</style>
