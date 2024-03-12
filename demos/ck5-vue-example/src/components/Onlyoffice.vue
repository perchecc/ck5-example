<template>
  <DocumentEditor
    id="docEditor"
    documentServerUrl="http://documentserver/"
    :config="config"
    :events_onDocumentReady="onDocumentReady"
  />
</template>

<script>
import { DocumentEditor } from "@onlyoffice/document-editor-vue";

export default {
  name: "ExampleComponent",
  components: {
    DocumentEditor,
  },
  data() {
    return {
      config: {
        document: {
          fileType: "docx",
          key: "Khirz6zTPdfd7",
          title: "Example Document Title.docx",
          url: "https://example.com/url-to-example-document.docx",
        },
        documentType: "word",
        editorConfig: {
          callbackUrl: "https://example.com/url-to-callback.ashx",
        },
      },
    };
  },
  methods: {
    onDocumentReady() {
      console.log("Document is loaded");
    },
  },
};
</script>
