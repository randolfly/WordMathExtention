<template>
  <div class="app">
    <div class="row">
      <div class="label">AsciiMath</div>
      <textarea v-model="ascii" placeholder="例如: sum_(i=1)^n i = (n(n+1))/2" />
    </div>

    <div class="btns">
      <button class="primary" :disabled="busy" @click="insertOrUpdate">
        插入/更新到 Word
      </button>
      <button :disabled="busy" @click="loadFromSelection">从选中公式加载</button>
      <button :disabled="busy" @click="refreshPreview">刷新预览</button>
    </div>

    <div class="row">
      <div class="label">预览（KaTeX）</div>
      <div class="preview" v-html="previewHtml"></div>
    </div>

    <div v-if="error" class="error">{{ error }}</div>
  </div>
</template>

<script setup lang="ts">
import { onMounted, ref, watch } from "vue";
import { officeReady } from "@/shared/officeReady";
import { asciiMathToLatex, latexToPreviewHtml } from "@/core/convert";
import { insertOrUpdateEquation, loadAsciiMathFromSelection } from "@/core/word";

const ascii = ref("");
const previewHtml = ref("");
const busy = ref(false);
const error = ref<string | null>(null);

const refreshPreview = () => {
  error.value = null;
  const tex = asciiMathToLatex(ascii.value);
  previewHtml.value = latexToPreviewHtml(tex);
};

const insertOrUpdate = async () => {
  busy.value = true;
  error.value = null;
  try {
    await insertOrUpdateEquation(ascii.value);
  } catch (e) {
    error.value = String(e);
  } finally {
    busy.value = false;
  }
};

const loadFromSelection = async () => {
  busy.value = true;
  error.value = null;
  try {
    const loaded = await loadAsciiMathFromSelection();
    if (loaded != null) {
      ascii.value = loaded;
      refreshPreview();
    } else {
      error.value = "未检测到选中位置的 WordMath 公式（需要先用本插件插入）。";
    }
  } catch (e) {
    error.value = String(e);
  } finally {
    busy.value = false;
  }
};

watch(ascii, () => {
  refreshPreview();
});

onMounted(async () => {
  await officeReady();
  refreshPreview();
});
</script>

