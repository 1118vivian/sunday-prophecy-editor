(function () {
  const fields = [
    "group",
    "name",
    "week",
    "startDate",
    "endDate",
    "title",
    "scripture",
    "revelation",
    "application",
  ];

  const form = document.getElementById("prophecyForm");
  const fileNamePreview = document.getElementById("fileNamePreview");
  const preview = document.getElementById("documentPreview");
  const WORD_BODY_SIZE = 28;
  const A4_WIDTH_DXA = 11906;
  const A4_HEIGHT_DXA = 16838;
  const PAGE_HORIZONTAL_MARGIN_DXA = 720;
  const TABLE_WIDTH_DXA = A4_WIDTH_DXA - PAGE_HORIZONTAL_MARGIN_DXA * 2;
  const LABEL_WIDTH_DXA = 1100;
  const CONTENT_WIDTH_DXA = TABLE_WIDTH_DXA - LABEL_WIDTH_DXA;

  const data = () =>
    Object.fromEntries(fields.map((field) => [field, document.getElementById(field).value.trim()]));

  const escapeText = (value) =>
    value.replace(/[&<>"']/g, (match) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[match]);

  const dateText = (value) => {
    if (!value) return "";
    const [year, month, day] = value.split("-");
    return `${year}年${Number(month)}月${Number(day)}日`;
  };

  const fileName = (values, extension) => {
    const name = values.name || "未命名";
    const week = values.week ? `第${values.week}週` : "主日申言稿";
    return `${name}_${week}.${extension}`;
  };

  const syncPreview = () => {
    const values = data();
    const previewValues = {
      ...values,
      startDateText: dateText(values.startDate),
      endDateText: dateText(values.endDate),
    };

    document.querySelectorAll("[data-preview]").forEach((node) => {
      const key = node.getAttribute("data-preview");
      node.innerHTML = escapeText(previewValues[key] || "");
    });

    fileNamePreview.textContent = fileName(values, "docx");
  };

  const lineBreaks = (text) => {
    const docx = window.docx;
    const lines = text ? text.split(/\r?\n/) : [""];
    return lines.flatMap((line, index) => {
      const runs = [];
      if (index > 0) runs.push(new docx.TextRun({ break: 1 }));
      runs.push(new docx.TextRun({ text: line, font: "Microsoft JhengHei", size: WORD_BODY_SIZE }));
      return runs;
    });
  };

  const makeCell = (children, options = {}) => {
    const docx = window.docx;
    return new docx.TableCell({
      width: options.width ? { size: options.width, type: docx.WidthType.DXA } : undefined,
      shading: options.shading ? { fill: options.shading } : undefined,
      margins: { top: 120, bottom: 120, left: 140, right: 140 },
      verticalAlign: options.verticalAlign || docx.VerticalAlign.CENTER,
      children,
    });
  };

  const labelParagraph = (text) => {
    const docx = window.docx;
    return new docx.Paragraph({
      alignment: docx.AlignmentType.CENTER,
      children: [new docx.TextRun({ text, bold: true, font: "Microsoft JhengHei", size: WORD_BODY_SIZE })],
    });
  };

  const contentParagraph = (text) => {
    const docx = window.docx;
    return new docx.Paragraph({
      spacing: { line: 360 },
      children: lineBreaks(text),
    });
  };

  const sectionRow = (label, text, minHeight) => {
    const docx = window.docx;
    return new docx.TableRow({
      height: { value: minHeight, rule: docx.HeightRule.ATLEAST },
      children: [
        makeCell([labelParagraph(label)], { width: LABEL_WIDTH_DXA }),
        makeCell([contentParagraph(text)], { width: CONTENT_WIDTH_DXA }),
      ],
    });
  };

  const captureA4Preview = async (scale = 2) => {
    if (!window.html2canvas) throw new Error("版面產生套件尚未載入，請重新整理頁面後再試。");
    const exportPreview = preview.cloneNode(true);
    exportPreview.id = "documentPreviewExport";
    exportPreview.classList.add("pdf-export-page");
    document.body.appendChild(exportPreview);

    try {
      return await window.html2canvas(exportPreview, {
        scale,
        useCORS: true,
        backgroundColor: "#ffffff",
      });
    } finally {
      exportPreview.remove();
    }
  };

  const buildDocx = async () => {
    if (!window.docx) throw new Error("DOCX 套件尚未載入，請確認網路連線後重試。");
    if (!window.html2canvas) throw new Error("手機版 Word 版面套件尚未載入，請重新整理頁面後再試。");
    const docx = window.docx;
    const values = data();
    const canvas = await captureA4Preview(1);

    const doc = new docx.Document({
      sections: [
        {
          properties: {
            page: {
              size: { width: A4_WIDTH_DXA, height: A4_HEIGHT_DXA },
              margin: { top: 0, right: 0, bottom: 0, left: 0 },
            },
          },
          children: [
            new docx.Paragraph({
              spacing: { before: 0, after: 0 },
              children: [
                new docx.ImageRun({
                  data: canvas.toDataURL("image/png"),
                  transformation: { width: 794, height: 1123 },
                  altText: {
                    title: "主日申言稿",
                    description: "由主日申言稿編輯器產生的固定 A4 版面",
                    name: "主日申言稿",
                  },
                }),
              ],
            }),
          ],
        },
      ],
    });

    const blob = await docx.Packer.toBlob(doc);
    saveBlob(blob, fileName(values, "docx"));
  };

  const buildPdf = async () => {
    if (!window.html2canvas || !window.jspdf?.jsPDF) throw new Error("PDF 單頁產生套件尚未載入，請重新整理頁面後再試。");
    const values = data();
    const canvas = await captureA4Preview(2);
    const pdf = new window.jspdf.jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
    const pageWidth = 210;
    const pageHeight = 297;
    const margin = 6;
    const maxWidth = pageWidth - margin * 2;
    const maxHeight = pageHeight - margin * 2;
    const imageRatio = canvas.width / canvas.height;
    let imageWidth = maxWidth;
    let imageHeight = imageWidth / imageRatio;

    if (imageHeight > maxHeight) {
      imageHeight = maxHeight;
      imageWidth = imageHeight * imageRatio;
    }

    const x = (pageWidth - imageWidth) / 2;
    const y = (pageHeight - imageHeight) / 2;
    pdf.addImage(canvas.toDataURL("image/jpeg", 0.98), "JPEG", x, y, imageWidth, imageHeight);
    pdf.save(fileName(values, "pdf"));
  };

  const saveBlob = (blob, name) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const runAction = async (button, action) => {
    const original = button.textContent;
    button.disabled = true;
    button.textContent = "產生中...";
    try {
      await action();
    } catch (error) {
      alert(error.message || "產生檔案時發生錯誤。");
    } finally {
      button.disabled = false;
      button.textContent = original;
    }
  };

  form.addEventListener("input", syncPreview);
  document.getElementById("clearForm").addEventListener("click", () => {
    form.reset();
    syncPreview();
  });
  document.getElementById("downloadDocx").addEventListener("click", (event) => runAction(event.currentTarget, buildDocx));
  document.getElementById("downloadPdf").addEventListener("click", (event) => runAction(event.currentTarget, buildPdf));

  syncPreview();
})();
