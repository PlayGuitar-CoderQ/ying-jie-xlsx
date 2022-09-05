import type { Ref } from 'vue';
import type { WorkSheet } from 'xlsx';

import { useNotification, UploadFileInfo } from 'naive-ui';
import { read, writeFile, utils, } from 'xlsx';
import { ref } from 'vue';

type TransformType = 0 | 1;

enum notifyEnum {
  "error",
  "success",
}

const NOTIFY_CONFIG = {
  [notifyEnum[1]]: {
    meta: "解析成功，即将弹出下载按钮",
    content: "成功 (success)",
    duration: 2500,
    keepAliveOnHover: false
  },
  [notifyEnum[0]]: {
    meta: "解析失败，请检查你的 xlsx 格式是否和以前一样, 如果还不行请联系波爷",
    content: "失败 (error)",
    duration: 2500,
    keepAliveOnHover: false
  }
}
const rangeAll = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
const rangeStartIndex = 3; // 扫描列范围开始索引
const rangeEndIndex = 20; // 扫描列范围结束索引
const newXlsxWriteRowIndex = 1; // 新创建 xlsx 文件写入的起始行索引
const newXlsxWriteColIndex = 3; // 新创建 xlsx 文件写入的起始列索引

/**
 * 获取上传表格指定索引的每行数据
 * @param sheet 
 * @param rowIndex 
 * @returns 
 */
const getRowData = (sheet: WorkSheet, rowIndex: number) => {
  const result: any[] = [];
  const range = rangeAll.slice(rangeStartIndex, rangeEndIndex);
  let index = 0;
  while (index <= range.length - 1) {
    let letter = range[index];
    let cellVal = sheet[`${letter}${rowIndex}`]
    let w = cellVal?.w ?? "";
    result.push(w);
    index++
  }
  return result;
}

/**
 * 获取适配写入新 xlsx 的转换后的数据
 * @param cacheRes 
 * @param sheet 
 * @returns 
 */
 const getTransFormResult = (cacheRes: any = {}, sheet: WorkSheet) => {
  const result: any[] = []; // 最终被转换的结果
  for (let key in cacheRes) {
    const rowIndex = cacheRes[key].i;
    const rowData = getRowData(sheet, rowIndex);
    result.push(rowData);
  }
  return result;
}

/**
 * 导出下载生成解析转换后数据的 xlsx
 * @param cacheRes 
 * @param sheet 
 * @returns 
 */
const exportGenerateNewXlsx = (cacheRes: any = {}, sheet: WorkSheet) => {
  // 生成逻辑
  const result = getTransFormResult(cacheRes, sheet);
  const newWorkBook = utils.book_new();
  const newWorkSheet = utils.sheet_add_aoa(newWorkBook, result, { origin: { r: newXlsxWriteRowIndex, c: newXlsxWriteColIndex } });
  utils.book_append_sheet(newWorkBook, newWorkSheet, "new.xlsx");
  return writeFile(newWorkBook, "new.xlsx", { type: 'binary' });
}



const useTransformXlsx = (): [(file: File | null | undefined) => void] => {
  const notification = useNotification();
  const transformType = ref<TransformType>(0); // 为了配合使用枚举使用 0 | 1 定义状态

  const notify = () => {
    const type = notifyEnum[transformType.value] as unknown as "error" | "success";
    const notifyConfig = NOTIFY_CONFIG[type];
    notification[type](notifyConfig);
  }

  // 转换逻辑核心
  const transformXlsx = (file: File | null | undefined) => {
    if (!file) {
      notify();
      return;
    };

    try {
      const readerFile = new FileReader();

      readerFile.onload = function (e: ProgressEvent<FileReader>) {
        let fileData = e.target?.result;
        let startRowIndex = 2; // 从当前行索引开始扫描
        let endRowIndex = 10000; // 从当前行索引结束扫描
        let _cycleIndex = 0; // 遍历循环指针
        const cacheObj: any = {}; // 暂存筛选出结果的行id、行占比的键值对 (包含了行的索引)

        const workbook = read(fileData, { type: 'array' }) as any;
        const sheet = workbook.Sheets["Sheet1"];

        // 扫描当前表格
        while (_cycleIndex < endRowIndex - 1) {
          const mainIdItem = sheet[`D${startRowIndex}`]; // mainId 的行数据
          const percentItem = sheet[`T${startRowIndex}`]; // 占比的行数据

          if (!mainIdItem) {
            startRowIndex += 1;
            _cycleIndex += 1;
            break;
          }
          // v: 是 xlsx 解析出每行的数据，文档：https://www.npmjs.com/package/xlsx#modifying-cell-values
          // 把对比结果放入暂存
          if (cacheObj.hasOwnProperty(mainIdItem.v)) {
            const npAndCpList = [cacheObj[mainIdItem.v].v, percentItem.v]; // np: 当前循环这行的占比，Cp: 暂存里面同 key 下值的占比
            const nrAndCrList = [cacheObj[mainIdItem.v].i, startRowIndex]; // nr: 当前循环的行索引，Cr: 暂存里面同 key 下值索引
            let choiceIndex = cacheObj[mainIdItem.v].v > percentItem.v ? 0 : 1;
            cacheObj[mainIdItem.v] = { v: npAndCpList[choiceIndex], i: nrAndCrList[choiceIndex] };
          } else {
            cacheObj[mainIdItem.v] = { v: percentItem.v, i: startRowIndex }
          }

          startRowIndex += 1;
          _cycleIndex += 1;
        }

        exportGenerateNewXlsx(cacheObj, sheet);
        transformType.value = 1;
        notify();
      }

      readerFile.readAsArrayBuffer(file);
    } catch (err: any) {
      throw new Error(err);
    }
  }

  return [transformXlsx]
}

export default useTransformXlsx;