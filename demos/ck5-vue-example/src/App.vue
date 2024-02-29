<!-- App.vue -->
<template>
  <div id="app">
    <!-- 文件选择按钮 -->
    <input type="file" @change="onFileChange" accept=".docx" ref="fileInput" />
    <!-- CKEditor 组件 -->
    <ckeditor
      ref="ckeditor"
      :editor="editor"
      v-model="editorData"
      :config="editorConfig"
    ></ckeditor>
    <button type="button" @click="loadAndPreviewDocx">预览文件</button>
    <div ref="previewContainer" class="docx-preview-container"></div>
    <!-- 导出 -->
    <input type="button" value="导出为docx" @click="exportToDocx" />
    <p contenteditable=true>这是一个可以编辑的段落</p>
  </div>
</template>

<style scoped>
.docx-preview-container {
  width: 100%;
  height: 400px; /* 根据需要设置高度 */
  overflow-y: auto;
}
</style>

<script>
// ⚠️ NOTE: We don't use @ckeditor/ckeditor5-build-classic any more!
// Since we're building CKEditor&nbsp;5 from source, we use the source version of ClassicEditor.
import { ClassicEditor } from "@ckeditor/ckeditor5-editor-classic";

import { Essentials } from "@ckeditor/ckeditor5-essentials";
import { Bold, Italic } from "@ckeditor/ckeditor5-basic-styles";
import { Link } from "@ckeditor/ckeditor5-link";
import { Paragraph } from "@ckeditor/ckeditor5-paragraph";
import mammoth from "mammoth";
import htmlDocx from "html-docx-js/dist/html-docx";
import { saveAs } from "file-saver";
import { renderAsync } from "docx-preview";

export default {
  name: "app",
  data() {
    return {
      editor: ClassicEditor,
      editorData: "<p>Content of the editor.</p>",
      editorConfig: {
        plugins: [Essentials, Bold, Italic, Link, Paragraph],

        toolbar: {
          items: ["bold", "italic", "link", "undo", "redo"],
        },
      },
    };
  },
  methods: {
    async onFileChange(event) {
      const selectedFiles = event.target.files;
      if (selectedFiles && selectedFiles.length > 0) {
        await this.handleFileUpload(selectedFiles[0]);
      }
    },
    async handleFileUpload(file) {
      const reader = new FileReader();
      const that = this;
      reader.readAsArrayBuffer(file);
      reader.onload = async () => {
        try {
          const buffer = reader.result;
          console.log(buffer);
          // const tableStyleMap = [
          //   // 将Word表格转换为HTML表格
          //   "w:tblPrEx => table",
          //   // 保持表格内文字默认转换为单元格内的段落
          //   "w:tc => td",
          //   "w:tr => tr",
          // ];
          // const result = await mammoth.convertToHtml({
          //   arrayBuffer: buffer,
          //   // 可选配置项，比如保留样式等
          //   // styleMap: [
          //   //   "p[style-name='Title'] => h1",
          //   //   "p[style-name='Heading 1'] => h2",
          //   // ],'
          //   styleMap: tableStyleMap,
          // });
          // this.editorData = result.value;
          // console.log(result.messages); // 输出转换过程中的警告和错误信息
          const previewContainer = that.$refs.previewContainer;
          // 渲染docx文件
          renderAsync(buffer, previewContainer, {
            className: "docx",
            prefix: "myprefix-", // 可选，前缀用于生成的DOM元素ID
            inWrapper: true, // 可选，是否包裹在一个div中
          })
            .then(() => {})
            .catch((error) => {
              console.error("Failed to render docx:", error);
            });
        } catch (error) {
          console.error("Error converting file to HTML:", error);
        }
      };
      reader.onerror = function (error) {
        console.error("Error reading the file:", error);
      };
    },
    exportToDocx() {
      // 获取编辑器的HTML内容
      console.log(this.editorData);
      // 使用html-docx-js将HTML转换为DOCX格式的Blob对象
      const docx = htmlDocx.asBlob(this.editorData);
      // 使用file-saver保存Blob对象为DOCX文件
      saveAs(docx, "output.docx");
    },

    async loadAndPreviewDocx() {
      // 假设fileUrl是从某个地方获得的.docx文件路径或blob URL
      // var fileUrl =
      //   "https://cdn.perche.cc/%E5%A1%91%E5%A3%B3%E6%96%AD%E8%B7%AF%E5%99%A8%E6%B5%8B%E8%AF%95%E6%A8%A1%E6%9D%BF%20-%20%E5%8F%B0%E5%8C%BA%28_20230607095909%29%20-%20%E5%89%AF%E6%9C%AC.docxF";
      // 获取容器元素
      const previewContainer = this.$refs.previewContainer;
      // 渲染docx文件
      try {
        await renderAsync(fileUrl, previewContainer, {
          className: "docx",
          prefix: "myprefix-", // 可选，前缀用于生成的DOM元素ID
          inWrapper: true, // 可选，是否包裹在一个div中
          imageHandler: (imageElement, imageData) => {
            /* 自定义图片处理逻辑 */
          },
        });
      } catch (error) {
        console.error("Failed to render docx:", error);
      }
    },
  },
};
</script>
