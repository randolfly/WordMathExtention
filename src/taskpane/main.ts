import { createApp } from "vue";
import App from "./App.vue";
import "./styles.css";

// 确保 Office.js 完全加载后再初始化 Vue 应用
Office.onReady(() => {
  createApp(App).mount("#app");
});