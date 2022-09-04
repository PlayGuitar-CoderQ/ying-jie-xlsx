<script setup lang="ts">
import { NUpload, NUploadDragger, NIcon, NText, NP, UploadFileInfo } from 'naive-ui';
import { ArchiveOutline as ArchiveIcon } from '@vicons/ionicons5';
import useTransformXlsx from '@hooks/useTransformXlsx';

interface UploadChangeOption {
  file: UploadFileInfo;
  fileList: Array<UploadFileInfo>;
  event?: Event;
}

const uploadTip = '点击或者拖动文件到该区域上传';
const uploadDangerTip = '请不要上传敏感数据，比如你的银行卡号和密码，信用卡号有效期和安全码';
const joke = '要不然我就偷了你的 🫡';

const [transformXlsx] = useTransformXlsx();

const handleUploadFinish = (option: UploadChangeOption) => {
  transformXlsx(option.file.file);
}
</script>

<template>
  <div class="wrapper">
    <NUpload multiple directory-dnd :show-file-list=false @change="handleUploadFinish">
      <NUploadDragger>

        <div style="margin-bottom: 12px">
          <NIcon size="48" :depth="3">
            <ArchiveIcon />
          </NIcon>
        </div>

        <NText style="font-size: 16px">
          {{ uploadTip }}
        </NText>

        <NP depth="3" style="margin: 8px 0 0 0">
          {{ uploadDangerTip }}
        </NP>

        <NP depth="3" style="margin: 8px 0 0 0">
          {{ joke }}
        </NP>

      </NUploadDragger>
    </NUpload>
  </div>
</template>

<style scoped>
.wrapper {
  display: flex;
  justify-content: center;
  align-items: center;
  width: 100%;
  height: 97vh;
}
</style>
