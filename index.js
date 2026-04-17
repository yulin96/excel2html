/**
 * Word to HTML 转换器
 * 使用 mammoth.js 将 .docx 文件转换为 HTML，
 * 然后通过 DOM 后处理来移除用户选择的格式。
 */
(function () {
  'use strict';

  // ===== DOM Elements =====
  const uploadArea = document.getElementById('uploadArea');
  const fileInput = document.getElementById('fileInput');
  const fileInfo = document.getElementById('fileInfo');
  const fileName = document.getElementById('fileName');
  const fileRemove = document.getElementById('fileRemove');
  const btnConvert = document.getElementById('btnConvert');
  const tabPreview = document.getElementById('tabPreview');
  const tabSource = document.getElementById('tabSource');
  const panePreview = document.getElementById('panePreview');
  const paneSource = document.getElementById('paneSource');
  const previewFrame = document.getElementById('previewFrame');
  const outputPlaceholder = document.getElementById('outputPlaceholder');
  const sourceCode = document.getElementById('sourceCode');
  const btnCopy = document.getElementById('btnCopy');
  const btnDownload = document.getElementById('btnDownload');
  const messagesBar = document.getElementById('messagesBar');
  const messagesCount = document.getElementById('messagesCount');
  const messagesList = document.getElementById('messagesList');
  const loadingOverlay = document.getElementById('loadingOverlay');
  const toast = document.getElementById('toast');
  const btnTheme = document.getElementById('btnTheme');

  // ===== Theme =====
  function initTheme() {
    const saved = localStorage.getItem('docx2html-theme');
    // 默认浅色，只有明确保存了 dark 才用暗色
    if (saved === 'dark') {
      document.documentElement.setAttribute('data-theme', 'dark');
    } else {
      document.documentElement.removeAttribute('data-theme');
    }
  }

  function toggleTheme() {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    if (isDark) {
      document.documentElement.removeAttribute('data-theme');
      localStorage.setItem('docx2html-theme', 'light');
    } else {
      document.documentElement.setAttribute('data-theme', 'dark');
      localStorage.setItem('docx2html-theme', 'dark');
    }
  }

  initTheme();
  btnTheme.addEventListener('click', toggleTheme);

  // ===== State =====
  let currentFile = null;
  let convertedHtml = '';

  // ===== Upload Handling =====
  uploadArea.addEventListener('click', () => fileInput.click());

  uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
  });

  uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
  });

  uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFile(files[0]);
    }
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
      handleFile(fileInput.files[0]);
    }
  });

  fileRemove.addEventListener('click', (e) => {
    e.stopPropagation();
    clearFile();
  });

  function handleFile(file) {
    if (!file.name.toLowerCase().endsWith('.docx')) {
      showToast('请选择 .docx 文件', 'error');
      return;
    }
    currentFile = file;
    fileName.textContent = file.name;
    fileInfo.style.display = 'flex';
    uploadArea.style.display = 'none';
    btnConvert.disabled = false;
  }

  function clearFile() {
    currentFile = null;
    fileInput.value = '';
    fileInfo.style.display = 'none';
    uploadArea.style.display = '';
    btnConvert.disabled = true;
    convertedHtml = '';
    previewFrame.style.display = 'none';
    previewFrame.innerHTML = '';
    outputPlaceholder.style.display = '';
    sourceCode.textContent = '';
    messagesBar.style.display = 'none';
  }

  // ===== Tab Switching =====
  tabPreview.addEventListener('click', () => switchTab('preview'));
  tabSource.addEventListener('click', () => switchTab('source'));

  function switchTab(tab) {
    tabPreview.classList.toggle('active', tab === 'preview');
    tabSource.classList.toggle('active', tab === 'source');
    panePreview.classList.toggle('active', tab === 'preview');
    paneSource.classList.toggle('active', tab === 'source');
  }

  // ===== Conversion =====
  btnConvert.addEventListener('click', doConvert);

  async function doConvert() {
    if (!currentFile) return;

    loadingOverlay.style.display = 'flex';
    btnConvert.disabled = true;

    try {
      const arrayBuffer = await readFileAsArrayBuffer(currentFile);

      // 构建 mammoth 选项
      const options = buildMammothOptions();

      const result = await mammoth.convertToHtml({ arrayBuffer }, options);

      // 后处理 HTML
      let html = result.value;
      html = postProcessHtml(html);

      convertedHtml = html;

      // 显示预览
      outputPlaceholder.style.display = 'none';
      previewFrame.style.display = '';
      previewFrame.innerHTML = html;

      // 显示源码（带简单高亮）
      sourceCode.textContent = formatHtml(html);

      // 显示消息
      showMessages(result.messages);

      showToast('转换完成！', 'success');
    } catch (err) {
      console.error(err);
      showToast('转换失败: ' + err.message, 'error');
    } finally {
      loadingOverlay.style.display = 'none';
      btnConvert.disabled = false;
    }
  }

  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  // ===== Build Mammoth Options =====
  function buildMammothOptions() {
    const opts = {};
    const styleMap = [];

    // mammoth 默认会把 b/i/u/s 映射到 strong/em 等
    // 如果用户选择移除，我们映射到空（不生成标签）
    const removeBold = document.getElementById('removeBold').checked;
    const removeItalic = document.getElementById('removeItalic').checked;
    const removeStrikethrough = document.getElementById('removeStrikethrough').checked;
    const removeUnderline = document.getElementById('removeUnderline').checked;

    if (removeBold) {
      styleMap.push('b => ');
    }
    if (removeItalic) {
      styleMap.push('i => ');
    }
    if (removeStrikethrough) {
      styleMap.push('strike => ');
    }
    if (removeUnderline) {
      styleMap.push('u => ');
    }

    if (styleMap.length > 0) {
      opts.styleMap = styleMap;
    }

    return opts;
  }

  // ===== Post-Process HTML =====
  function postProcessHtml(html) {
    // 使用 DOMParser 来操作 HTML
    const parser = new DOMParser();
    const doc = parser.parseFromString('<div id="root">' + html + '</div>', 'text/html');
    const root = doc.getElementById('root');

    const removeFonts = document.getElementById('removeFonts').checked;
    const removeFontSize = document.getElementById('removeFontSize').checked;
    const removeColor = document.getElementById('removeColor').checked;
    const removeBgColor = document.getElementById('removeBgColor').checked;
    const removeImages = document.getElementById('removeImages').checked;
    const removeLinks = document.getElementById('removeLinks').checked;
    const removeAllStyles = document.getElementById('removeAllStyles').checked;
    const removeEmptyP = document.getElementById('removeEmptyP').checked;
    const removeBold = document.getElementById('removeBold').checked;
    const removeItalic = document.getElementById('removeItalic').checked;
    const removeStrikethrough = document.getElementById('removeStrikethrough').checked;
    const removeUnderline = document.getElementById('removeUnderline').checked;

    // 移除图片
    if (removeImages) {
      root.querySelectorAll('img').forEach(el => el.remove());
    }

    // 移除链接（保留文字）
    if (removeLinks) {
      root.querySelectorAll('a').forEach(el => {
        const text = doc.createTextNode(el.textContent);
        el.parentNode.replaceChild(text, el);
      });
    }

    // 移除 bold 标签（如果 mammoth styleMap 没处理到的）
    if (removeBold) {
      unwrapTag(root, 'strong');
      unwrapTag(root, 'b');
    }

    // 移除 italic 标签
    if (removeItalic) {
      unwrapTag(root, 'em');
      unwrapTag(root, 'i');
    }

    // 移除 strikethrough 标签
    if (removeStrikethrough) {
      unwrapTag(root, 's');
      unwrapTag(root, 'del');
      unwrapTag(root, 'strike');
    }

    // 移除 underline
    if (removeUnderline) {
      unwrapTag(root, 'u');
      // 也移除样式中的 text-decoration
      root.querySelectorAll('[style]').forEach(el => {
        el.style.textDecoration = '';
      });
    }

    // 处理内联样式
    if (removeAllStyles) {
      root.querySelectorAll('[style]').forEach(el => {
        el.removeAttribute('style');
      });
    } else {
      // 选择性移除样式属性
      root.querySelectorAll('[style]').forEach(el => {
        if (removeFonts) {
          el.style.fontFamily = '';
        }
        if (removeFontSize) {
          el.style.fontSize = '';
        }
        if (removeColor) {
          el.style.color = '';
        }
        if (removeBgColor) {
          el.style.backgroundColor = '';
        }
        // 如果 style 变空了，移除 style 属性
        if (el.getAttribute('style') && el.getAttribute('style').trim() === '') {
          el.removeAttribute('style');
        }
      });
    }

    // 移除空段落
    if (removeEmptyP) {
      root.querySelectorAll('p').forEach(el => {
        if (el.textContent.trim() === '' && el.querySelectorAll('img, br').length === 0) {
          el.remove();
        }
      });
    }

    return root.innerHTML;
  }

  /**
   * 将元素解包：把标签移除但保留其子元素
   */
  function unwrapTag(root, tagName) {
    const elements = root.querySelectorAll(tagName);
    elements.forEach(el => {
      const parent = el.parentNode;
      while (el.firstChild) {
        parent.insertBefore(el.firstChild, el);
      }
      parent.removeChild(el);
    });
  }

  // ===== Format HTML (简单缩进) =====
  function formatHtml(html) {
    // 简单格式化
    let formatted = '';
    let indent = 0;
    const tab = '  ';

    // 将标签分离
    const tokens = html.replace(/>\s*</g, '>\n<').split('\n');

    tokens.forEach(token => {
      const trimmed = token.trim();
      if (!trimmed) return;

      // 闭合标签减少缩进
      if (trimmed.match(/^<\/\w/)) {
        indent = Math.max(0, indent - 1);
      }

      formatted += tab.repeat(indent) + trimmed + '\n';

      // 开始标签增加缩进（自闭合标签除外）
      if (
        trimmed.match(/^<\w[^>]*[^/]>$/) &&
        !trimmed.match(/^<(br|hr|img|input|meta|link)\b/i) &&
        !trimmed.match(/^<\//)
      ) {
        indent++;
      }
    });

    return formatted.trim();
  }

  // ===== Messages =====
  function showMessages(messages) {
    if (!messages || messages.length === 0) {
      messagesBar.style.display = 'none';
      return;
    }

    messagesBar.style.display = '';
    messagesCount.textContent = messages.length + ' 条消息';
    messagesList.innerHTML = '';

    messages.forEach(msg => {
      const div = document.createElement('div');
      div.className = 'message-item ' + (msg.type || 'warning');
      div.textContent = msg.message;
      messagesList.appendChild(div);
    });
  }

  // ===== Copy & Download =====
  btnCopy.addEventListener('click', () => {
    if (!convertedHtml) {
      showToast('没有可复制的内容', 'error');
      return;
    }
    navigator.clipboard.writeText(convertedHtml).then(() => {
      showToast('已复制到剪贴板', 'success');
    }).catch(() => {
      // 回退方案
      const textarea = document.createElement('textarea');
      textarea.value = convertedHtml;
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      showToast('已复制到剪贴板', 'success');
    });
  });

  btnDownload.addEventListener('click', () => {
    if (!convertedHtml) {
      showToast('没有可下载的内容', 'error');
      return;
    }

    const fullHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(currentFile ? currentFile.name.replace(/\.docx$/i, '') : 'document')}</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', sans-serif;
      line-height: 1.8;
      max-width: 900px;
      margin: 0 auto;
      padding: 40px 24px;
      color: #333;
    }
    table { border-collapse: collapse; width: 100%; margin: 1em 0; }
    th, td { border: 1px solid #ddd; padding: 8px 12px; }
    th { background: #f5f5f5; }
    img { max-width: 100%; height: auto; }
    h1 { font-size: 2em; margin: 0.8em 0 0.4em; }
    h2 { font-size: 1.5em; margin: 0.8em 0 0.4em; }
    h3 { font-size: 1.17em; margin: 0.8em 0 0.4em; }
    p { margin: 0.5em 0; }
    ul, ol { margin: 0.5em 0; padding-left: 2em; }
  </style>
</head>
<body>
${convertedHtml}
</body>
</html>`;

    const blob = new Blob([fullHtml], { type: 'text/html; charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = (currentFile ? currentFile.name.replace(/\.docx$/i, '') : 'document') + '.html';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('文件下载中…', 'success');
  });

  // ===== Helpers =====
  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  let toastTimer;
  function showToast(message, type = '') {
    clearTimeout(toastTimer);
    toast.textContent = message;
    toast.className = 'toast show' + (type ? ' ' + type : '');
    toastTimer = setTimeout(() => {
      toast.className = 'toast';
    }, 3000);
  }

  // ===== "移除所有内联样式" 联动 =====
  const removeAllStylesCheckbox = document.getElementById('removeAllStyles');
  const styleSubs = ['removeFonts', 'removeFontSize', 'removeColor', 'removeBgColor'];

  removeAllStylesCheckbox.addEventListener('change', () => {
    if (removeAllStylesCheckbox.checked) {
      styleSubs.forEach(id => {
        const el = document.getElementById(id);
        el.checked = true;
        el.closest('.option-item').style.opacity = '0.5';
        el.closest('.option-item').style.pointerEvents = 'none';
      });
    } else {
      styleSubs.forEach(id => {
        const el = document.getElementById(id);
        el.closest('.option-item').style.opacity = '';
        el.closest('.option-item').style.pointerEvents = '';
      });
    }
  });
})();
