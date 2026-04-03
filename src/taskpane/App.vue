<template>
  <div class="app">
    <div class="row">
      <div class="label">AsciiMath</div>
      <textarea v-model="ascii" placeholder="例如: sum_(i=1)^n i = (n(n+1))/2" />
    </div>

    <div class="btns">
      <button class="primary" :disabled="busy" @click="insertOrUpdate">插入到 Word </button>
      <button :disabled="busy" @click="refreshPreview">刷新</button>
      <button :disabled="!ascii.trim() || busy" @click="copyLatexToClipboard">复制 LaTeX </button>
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
const copyLatexToClipboard = async () => {
  if (!ascii.value.trim()) return;

  busy.value = true;
  error.value = null;

  try {
    const latex = asciiMathToLatex(ascii.value);
    await navigator.clipboard.writeText(latex);
  } catch (e) {
    error.value = "复制失败：" + String(e);
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

