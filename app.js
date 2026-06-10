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
      runs.push(new docx.TextRun({ text: line, font: "Microsoft JhengHei", size: 22 }));
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
      children: [new docx.TextRun({ text, bold: true, font: "Microsoft JhengHei", size: 24 })],
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
        makeCell([labelParagraph(label)], { width: 1100 }),
        makeCell([contentParagraph(text)]),
      ],
    });
  };

  const buildDocx = async () => {
    if (!window.docx) throw new Error("DOCX 套件尚未載入，請確認網路連線後重試。");
    const docx = window.docx;
    const values = data();

    const doc = new docx.Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 720, right: 720, bottom: 560, left: 720 },
            },
          },
          children: [
            new docx.Paragraph({
              alignment: docx.AlignmentType.CENTER,
              children: [new docx.TextRun({ text: "臺灣福音工作全時間訓練壯年成全班", bold: true, font: "Microsoft JhengHei", size: 28 })],
            }),
            new docx.Paragraph({
              alignment: docx.AlignmentType.CENTER,
              spacing: { after: 260 },
              children: [new docx.TextRun({ text: "主日申言稿", bold: true, font: "Microsoft JhengHei", size: 36 })],
            }),
            new docx.Paragraph({
              spacing: { after: 180 },
              children: [
                new docx.TextRun({ text: `第 ${values.group || "　"} 組   `, font: "Microsoft JhengHei", size: 22 }),
                new docx.TextRun({ text: `姓名 ${values.name || "　　　　　"}   `, font: "Microsoft JhengHei", size: 22 }),
                new docx.TextRun({ text: `第 ${values.week || "　"} 週   `, font: "Microsoft JhengHei", size: 22 }),
                new docx.TextRun({
                  text: `${dateText(values.startDate) || "20　年　月　日"} 至 ${dateText(values.endDate) || "20　年　月　日"}`,
                  font: "Microsoft JhengHei",
                  size: 22,
                }),
              ],
            }),
            new docx.Table({
              width: { size: 100, type: docx.WidthType.PERCENTAGE },
              rows: [
                sectionRow("題目", values.title, 560),
                sectionRow("經文", values.scripture, 1320),
                sectionRow("啟示", values.revelation, 3050),
                sectionRow("應用", values.application, 3050),
              ],
            }),
            new docx.Paragraph({ spacing: { before: 220 }, children: [new docx.TextRun({ text: "備註：1. 申言時注意：靈要剛強、思路清明、體態合宜。", font: "Microsoft JhengHei", size: 20 })] }),
            new docx.Paragraph({ children: [new docx.TextRun({ text: "　　　2. 申言稿內容以當週晨興聖言進度為範圍，以3分鐘為限。", font: "Microsoft JhengHei", size: 20 })] }),
            new docx.Paragraph({ children: [new docx.TextRun({ text: "　　　3. 若本週改寫其他作業，則不必填寫本表。", font: "Microsoft JhengHei", size: 20 })] }),
          ],
        },
      ],
    });

    const blob = await docx.Packer.toBlob(doc);
    saveBlob(blob, fileName(values, "docx"));
  };

  const buildPdf = async () => {
    if (!window.html2pdf) throw new Error("PDF 套件尚未載入，請確認網路連線後重試。");
    const values = data();
    await window.html2pdf()
      .set({
        margin: 8,
        filename: fileName(values, "pdf"),
        image: { type: "jpeg", quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, backgroundColor: "#fffdfa" },
        jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
        pagebreak: { mode: ["avoid-all"] },
      })
      .from(preview)
      .save();
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
