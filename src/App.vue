<script setup lang="ts">
import Upload from './components/Upload.vue';
import { darkTheme, NConfigProvider, NNotificationProvider } from 'naive-ui';
</script>

<template>
  <NConfigProvider :theme="darkTheme">
    <NNotificationProvider>
      <Upload />
    </NNotificationProvider>
  </NConfigProvider>
</template>

<style scoped>
</style>
