import * as fs from 'fs';
import AdmZip from 'adm-zip';
import * as xml2js from 'xml2js';
import * as path from 'path';

// 辅助函数：解析 XML 字符串
export async function parseXml(xmlString: string): Promise<any> {
  const parser = new xml2js.Parser({
    explicitArray: true, // 总是将节点值放入数组
    mergeAttrs: true, // 将属性合并到其元素上
    charkey: '_', // 文本内容将存储在 '_' 键中，如果元素同时具有属性和文本
  });
  return parser.parseStringPromise(xmlString);
}

// 辅助函数：从 <p:txBody> 结构中提取文本
function extractTextFromTxBody(txBody: any): string {
  const paragraphTexts: string[] = [];
  if (txBody && txBody['a:p']) {
    // txBody['a:p'] 是段落数组
    for (const p of txBody['a:p']) {
      let currentParagraphText = '';
      if (p['a:r']) {
        // p['a:r'] 是文本运行 (run) 数组
        for (const r of p['a:r']) {
          if (r['a:t'] && r['a:t'][0]) {
            // r['a:t'] 是包含文本元素的数组
            const textContent = r['a:t'][0];
            if (typeof textContent === 'string') {
              currentParagraphText += textContent;
            } else if (textContent && typeof textContent._ === 'string') {
              // 当元素同时拥有属性和文本时，文本在 _ 键中
              currentParagraphText += textContent._;
            }
          }
        }
      }
      if (p['a:br']) {
        // 处理换行符 <a:br>
        // 通常 <a:br> 意味着一个空行或者段落间的视觉分隔
        // 在演讲者注释中，我们可能希望将其视为空行
        if (currentParagraphText.trim() !== '') {
          paragraphTexts.push(currentParagraphText.trim());
          currentParagraphText = ''; // 重置，准备新段落
        }
        paragraphTexts.push(''); // 添加一个空字符串代表空行
      }
      if (currentParagraphText.trim() !== '') {
        paragraphTexts.push(currentParagraphText.trim());
      }
    }
  }
  return paragraphTexts.join('\n\n'); // 段落间用双换行符分隔，形成空行效果
}

/**
 * 从 PPTX 文件中提取演讲者注释。
 * @param pptxFilePath PPTX 文件的路径。
 * @param outputDir 输出目录，用于图片路径生成。
 * @returns 格式化后的 Markdown 字符串，包含图片标签和演讲者注释。
 */
export async function extractNotesFromPptx(
  pptxFilePath: string,
  outputDir: string
): Promise<string> {
  console.log(`正在从 ${pptxFilePath} 提取注释...`);
  if (!fs.existsSync(pptxFilePath)) {
    return `错误：PPTX 文件 "${pptxFilePath}" 未找到。`;
  }

  // 检查 images 目录中的图片文件
  const imagesDir = path.join(outputDir, 'images');
  const imageFiles = fs.existsSync(imagesDir)
    ? fs
        .readdirSync(imagesDir)
        .filter((file) => /\.(png|jpg|jpeg|gif)$/i.test(file))
        .sort()
    : [];

  let zip: AdmZip;
  try {
    zip = new AdmZip(pptxFilePath);
  } catch (e: any) {
    console.error(`AdmZip 错误: ${e.message}`);
    return `错误：无法读取 PPTX 文件 "${pptxFilePath}"。它可能已损坏或不是有效的 zip 文件。详细信息: ${e.message}`;
  }

  let outputMarkdown = '';

  // 1. 解析 presentation.xml 获取幻灯片顺序和 rId
  const presentationXmlEntry = zip.getEntry('ppt/presentation.xml');
  if (!presentationXmlEntry) {
    return '错误：无法在 PPTX 文件中找到 ppt/presentation.xml。';
  }
  const presentationXmlContent = zip.readAsText(presentationXmlEntry);
  const presentationDoc = await parseXml(presentationXmlContent);

  const sldIdLstNode =
    presentationDoc?.['p:presentation']?.['p:sldIdLst']?.[0]?.['p:sldId'];

  if (!sldIdLstNode || !Array.isArray(sldIdLstNode)) {
    return '错误：无法解析幻灯片列表 (p:sldIdLst) 从 presentation.xml。文件可能已损坏或格式不正确。';
  }

  // 2. 解析 presentation.xml.rels 将幻灯片 rId 映射到幻灯片文件路径
  const presentationRelsEntry = zip.getEntry('ppt/_rels/presentation.xml.rels');
  if (!presentationRelsEntry) {
    return '错误：无法在 PPTX 文件中找到 ppt/_rels/presentation.xml.rels。';
  }
  const presentationRelsContent = zip.readAsText(presentationRelsEntry);
  const presentationRelsDoc = await parseXml(presentationRelsContent);

  const slideRelsMap: { [rId: string]: string } = {};

  if (
    presentationRelsDoc.Relationships &&
    presentationRelsDoc.Relationships.Relationship
  ) {
    for (const rel of presentationRelsDoc.Relationships.Relationship) {
      const relType = rel.Type?.[0];
      const relId = rel.Id?.[0];
      const relTarget = rel.Target?.[0];

      if (
        relType ===
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' &&
        relId &&
        relTarget
      ) {
        slideRelsMap[relId] = relTarget;
      }
    }
  } else {
    console.warn(
      '警告：在 ppt/_rels/presentation.xml.rels 中未找到 Relationships 或 Relationship 元素。'
    );
  }

  let slideCounter = 0;

  for (const sldIdEntry of sldIdLstNode) {
    slideCounter++;
    const slideRidArray = sldIdEntry['r:id'];

    // 获取当前幻灯片的图片文件名
    const imageFileName =
      imageFiles[slideCounter - 1] ||
      `${String(slideCounter).padStart(3, '0')}.png`;
    const imagePath = `images/${imageFileName}`;

    if (
      !slideRidArray ||
      !Array.isArray(slideRidArray) ||
      slideRidArray.length === 0
    ) {
      console.warn(
        `警告：幻灯片 ${slideCounter} 的 sldIdEntry 缺少有效的 r:id 数组。内容: ${JSON.stringify(
          sldIdEntry
        )}`
      );
      outputMarkdown += `![](${imagePath})\n\n`;
      outputMarkdown += `(无法找到此幻灯片的注释信息 - r:id 缺失或无效)\n\n`;
      continue;
    }

    const actualRidString = slideRidArray[0];
    const slideTarget = slideRelsMap[actualRidString];

    if (!slideTarget) {
      console.warn(
        `警告：未找到 rId 为 ${actualRidString} (幻灯片 ${slideCounter}) 的幻灯片目标。在 slideRelsMap 中未找到。`
      );
      outputMarkdown += `![](${imagePath})\n\n`;
      outputMarkdown += `(无法找到此幻灯片的注释信息 - 缺少目标文件映射)\n\n`;
      continue;
    }

    const slideFileName = path.basename(slideTarget);
    const slideRelsPath = `ppt/slides/_rels/${slideFileName}.rels`;
    const slideSpecificRelsEntry = zip.getEntry(slideRelsPath);

    let notesText = '';

    if (slideSpecificRelsEntry) {
      const slideSpecificRelsContent = zip.readAsText(slideSpecificRelsEntry);
      const slideSpecificRelsDoc = await parseXml(slideSpecificRelsContent);

      let notesSlideTargetRelPath: string | null = null;
      if (
        slideSpecificRelsDoc.Relationships &&
        slideSpecificRelsDoc.Relationships.Relationship
      ) {
        for (const rel of slideSpecificRelsDoc.Relationships.Relationship) {
          const relType = rel.Type?.[0];
          const relTarget = rel.Target?.[0];
          if (
            relType ===
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide' &&
            relTarget
          ) {
            notesSlideTargetRelPath = relTarget;
            break;
          }
        }
      }

      if (notesSlideTargetRelPath) {
        const fullSlideXmlPath = `ppt/${slideTarget}`;
        const baseDirForNotesRel = path.posix.dirname(fullSlideXmlPath);
        const notesSlidePath = path.posix.join(
          baseDirForNotesRel,
          notesSlideTargetRelPath
        );

        const notesSlideEntry = zip.getEntry(notesSlidePath);

        if (notesSlideEntry) {
          const notesSlideXmlContent = zip.readAsText(notesSlideEntry);
          const notesSlideDoc = await parseXml(notesSlideXmlContent);
          const pNotesNode = notesSlideDoc['p:notes'];

          if (pNotesNode) {
            let txBodyNode: any = undefined;
            const cSldNode = pNotesNode['p:cSld']?.[0];

            if (cSldNode) {
              const spTreeNode = cSldNode['p:spTree']?.[0];
              if (spTreeNode && Array.isArray(spTreeNode['p:sp'])) {
                for (const shape of spTreeNode['p:sp']) {
                  const ph =
                    shape?.['p:nvSpPr']?.[0]?.['p:nvPr']?.[0]?.['p:ph']?.[0];
                  const shapeType = ph?.type?.[0];
                  const shapeIdx = ph?.idx?.[0];

                  if (shapeType === 'body') {
                    txBodyNode = shape['p:txBody']?.[0];
                    if (txBodyNode) {
                      break;
                    } else {
                      console.warn(
                        `警告：幻灯片 ${slideCounter} 形状标记为 body 但未找到 txBody。`
                      );
                    }
                  }
                  // Fallback for cases where type might not be explicitly 'body' but idx is typical for notes
                  if (!txBodyNode && shapeIdx === '1') {
                    const potentialTxBody = shape['p:txBody']?.[0];
                    if (potentialTxBody) {
                      txBodyNode = potentialTxBody;
                    }
                  }
                }

                // If specific placeholder not found, try to find any txBody if only one exists
                if (!txBodyNode) {
                  let foundTxBodyInAnyShape: any = undefined;
                  let countTxBodyShapes = 0;
                  for (const shape of spTreeNode['p:sp']) {
                    const currentShapeTxBody = shape['p:txBody']?.[0];
                    if (currentShapeTxBody) {
                      foundTxBodyInAnyShape = currentShapeTxBody;
                      countTxBodyShapes++;
                    }
                  }
                  if (countTxBodyShapes === 1) {
                    txBodyNode = foundTxBodyInAnyShape;
                  } else if (countTxBodyShapes > 1) {
                    console.warn(
                      `警告：幻灯片 ${slideCounter} 在形状中找到 ${countTxBodyShapes} 个 txBody 且无明确标记。将尝试使用第一个。`
                    );
                    // Attempt to use the first one found if multiple exist and no clear primary
                    for (const shape of spTreeNode['p:sp']) {
                      const currentShapeTxBody = shape['p:txBody']?.[0];
                      if (currentShapeTxBody) {
                        txBodyNode = currentShapeTxBody;
                        break;
                      }
                    }
                  }
                }
              }
            }

            if (txBodyNode) {
              notesText = extractTextFromTxBody(txBodyNode);
            } else {
              notesText = '(注释页中无内容区域)';
            }
          } else {
            notesText = '(无法解析注释页结构)';
          }
        } else {
          console.warn(
            `警告：幻灯片 ${slideCounter} 未找到注释幻灯片文件: ${notesSlidePath}`
          );
          notesText = '(无法加载注释幻灯片文件)';
        }
      } else {
        notesText = '(此幻灯片无链接的注释页)';
      }

      if (notesText.trim() === '(此幻灯片无链接的注释页)') {
        notesText = '';
      }
    } else {
      console.warn(
        `警告：幻灯片 ${slideCounter} (目标 ${slideTarget}) 没有找到关联的 .rels 文件 (${slideRelsPath})。假设没有注释。`
      );
    }

    outputMarkdown += `![](${imagePath})\n\n`;
    outputMarkdown += notesText + (notesText.length > 0 ? '\n\n' : '');
  }
  console.log('注释已成功提取。');
  return outputMarkdown.trim();
}
