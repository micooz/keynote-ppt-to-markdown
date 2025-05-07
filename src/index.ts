import { extractNotesFromPptx } from './utils';
import fs from 'fs';
import { exec } from 'child_process';
import util from 'util';
import path from 'path';
import AdmZip from 'adm-zip'; // For direct PPTX media extraction
import { parseXml } from './utils';

const execPromise = util.promisify(exec);

// This function is now primarily for Keynote files on macOS,
// or any file Keynote can open and export slides from, via AppleScript.
async function exportSlidesAsImagesViaAppleScript(
  inputFileAbsPath: string, // .key file or a file Keynote can process on macOS
  baseOutputDirectory: string
): Promise<void> {
  const appleScriptPath = path.join(
    __dirname,
    'export_slides_to_images.applescript'
  );
  const tempImageExportPath = path.join(
    baseOutputDirectory,
    'applescript_exported_images_temp'
  );
  const finalImageOutputPath = path.join(baseOutputDirectory, 'images');

  // 复制 AppleScript 文件到输出目录
  const tempAppleScriptPath = path.join(
    baseOutputDirectory,
    'export_slides_to_images.applescript'
  );
  fs.copyFileSync(appleScriptPath, tempAppleScriptPath);

  const command = `osascript "${tempAppleScriptPath}" "${inputFileAbsPath}" "${tempImageExportPath}"`;

  try {
    console.log(`正在通过 AppleScript 导出图片: ${command}`);
    const { stdout, stderr } = await execPromise(command);

    // Refined error/warning handling for AppleScript
    let scriptError = false;
    if (stderr && stderr.trim() !== '') {
      if (stdout.trim().toLowerCase().startsWith('成功：幻灯片已临时导出到')) {
        console.log(`AppleScript stderr (可能是警告或提示): ${stderr}`);
      } else {
        console.warn(`AppleScript stderr (可能包含错误信息): ${stderr}`);
        scriptError = true; // Assume error if stdout is not success
      }
    }
    if (stdout.startsWith('错误：') || stdout.startsWith('AppleScript 错误:')) {
      scriptError = true;
    }
    if (
      !stdout.trim().toLowerCase().startsWith('成功：幻灯片已临时导出到') &&
      scriptError
    ) {
      throw new Error(
        `AppleScript 未成功导出图片: ${stdout.trim() || stderr.trim()}`
      );
    }
    if (
      !stdout.trim().toLowerCase().startsWith('成功：幻灯片已临时导出到') &&
      !scriptError &&
      stderr.trim() === ''
    ) {
      // If stdout is not success, and there was no stderr, it's still an issue.
      throw new Error(`AppleScript 未成功导出图片 (stdout): ${stdout.trim()}`);
    }
    console.log(
      stdout.trim() || `AppleScript 图片导出已执行 (无显式成功消息，但无错误)。`
    );

    if (!fs.existsSync(finalImageOutputPath)) {
      fs.mkdirSync(finalImageOutputPath, { recursive: true });
    } else {
      const existingFiles = fs.readdirSync(finalImageOutputPath);
      for (const file of existingFiles) {
        fs.unlinkSync(path.join(finalImageOutputPath, file));
      }
    }

    if (!fs.existsSync(tempImageExportPath)) {
      console.warn(
        `警告：AppleScript 临时图片目录 ${tempImageExportPath} 未创建。`
      );
      // Attempt to create the 'images' directory anyway, in case notes extraction expects it
      if (!fs.existsSync(finalImageOutputPath)) {
        fs.mkdirSync(finalImageOutputPath, { recursive: true });
      }
      return; // No images to process
    }

    const exportedImages = fs
      .readdirSync(tempImageExportPath)
      .filter(
        (file) =>
          file.toLowerCase().endsWith('.jpg') ||
          file.toLowerCase().endsWith('.jpeg') ||
          file.toLowerCase().endsWith('.png')
      )
      .sort((a, b) => {
        const numA = parseInt(a.replace(/[^0-9]/g, ''), 10);
        const numB = parseInt(b.replace(/[^0-9]/g, ''), 10);
        if (!isNaN(numA) && !isNaN(numB)) {
          return numA - numB;
        }
        return a.localeCompare(b);
      });

    if (exportedImages.length === 0) {
      console.warn(
        `警告：在 ${tempImageExportPath} 中没有找到符合条件的图片文件。`
      );
    } else {
      for (let i = 0; i < exportedImages.length; i++) {
        const oldPath = path.join(tempImageExportPath, exportedImages[i]);
        const newFilename = `${String(i + 1).padStart(3, '0')}${path.extname(
          exportedImages[i]
        )}`;
        const newPath = path.join(finalImageOutputPath, newFilename);
        fs.renameSync(oldPath, newPath);
      }
      console.log(
        `图片已成功处理并保存到 ${finalImageOutputPath}，共 ${exportedImages.length} 张图片。`
      );
    }

    if (fs.existsSync(tempImageExportPath)) {
      fs.rmSync(tempImageExportPath, { recursive: true, force: true });
    }
    // 删除临时 AppleScript 文件
    if (fs.existsSync(tempAppleScriptPath)) {
      fs.unlinkSync(tempAppleScriptPath);
    }
  } catch (error) {
    console.error('通过 AppleScript 导出幻灯片为图片过程中发生错误:', error);
    if (fs.existsSync(tempImageExportPath)) {
      try {
        fs.rmSync(tempImageExportPath, { recursive: true, force: true });
      } catch (cleanupError) {
        // Log cleanup error but re-throw original
      }
    }
    // 确保在错误情况下也删除临时 AppleScript 文件
    if (fs.existsSync(tempAppleScriptPath)) {
      try {
        fs.unlinkSync(tempAppleScriptPath);
      } catch (cleanupError) {
        // Log cleanup error but re-throw original
      }
    }
    throw error;
  }
}

async function extractMediaFromPptxDirectly(
  pptxFilePath: string,
  baseOutputDirectory: string
): Promise<void> {
  const finalImageOutputPath = path.join(baseOutputDirectory, 'images');
  console.log(`正在从 PPTX 文件直接提取媒体: ${pptxFilePath}`);

  try {
    if (!fs.existsSync(finalImageOutputPath)) {
      fs.mkdirSync(finalImageOutputPath, { recursive: true });
    } else {
      // Clear previous images
      const existingFiles = fs.readdirSync(finalImageOutputPath);
      for (const file of existingFiles) {
        fs.unlinkSync(path.join(finalImageOutputPath, file));
      }
    }

    const zip = new AdmZip(pptxFilePath);

    // 1. 首先获取幻灯片列表
    const presentationXmlEntry = zip.getEntry('ppt/presentation.xml');
    if (!presentationXmlEntry) {
      throw new Error('无法在 PPTX 文件中找到 ppt/presentation.xml');
    }
    const presentationXmlContent = zip.readAsText(presentationXmlEntry);
    const presentationDoc = await parseXml(presentationXmlContent);

    const sldIdLstNode =
      presentationDoc?.['p:presentation']?.['p:sldIdLst']?.[0]?.['p:sldId'];
    if (!sldIdLstNode || !Array.isArray(sldIdLstNode)) {
      throw new Error('无法解析幻灯片列表 (p:sldIdLst) 从 presentation.xml');
    }

    // 2. 获取幻灯片关系映射
    const presentationRelsEntry = zip.getEntry(
      'ppt/_rels/presentation.xml.rels'
    );
    if (!presentationRelsEntry) {
      throw new Error('无法在 PPTX 文件中找到 ppt/_rels/presentation.xml.rels');
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
    }

    // 3. 遍历每个幻灯片并提取图片
    for (let i = 0; i < sldIdLstNode.length; i++) {
      const sldIdEntry = sldIdLstNode[i];
      const slideRidArray = sldIdEntry['r:id'];

      if (
        !slideRidArray ||
        !Array.isArray(slideRidArray) ||
        slideRidArray.length === 0
      ) {
        console.warn(
          `警告：幻灯片 ${i + 1} 的 sldIdEntry 缺少有效的 r:id 数组`
        );
        continue;
      }

      const actualRidString = slideRidArray[0];
      const slideTarget = slideRelsMap[actualRidString];

      if (!slideTarget) {
        console.warn(
          `警告：未找到 rId 为 ${actualRidString} (幻灯片 ${
            i + 1
          }) 的幻灯片目标`
        );
        continue;
      }

      // 4. 获取幻灯片关系文件
      const slideFileName = path.basename(slideTarget);
      const slideRelsPath = `ppt/slides/_rels/${slideFileName}.rels`;
      const slideSpecificRelsEntry = zip.getEntry(slideRelsPath);

      if (slideSpecificRelsEntry) {
        const slideSpecificRelsContent = zip.readAsText(slideSpecificRelsEntry);
        const slideSpecificRelsDoc = await parseXml(slideSpecificRelsContent);

        // 5. 查找幻灯片图片关系
        if (
          slideSpecificRelsDoc.Relationships &&
          slideSpecificRelsDoc.Relationships.Relationship
        ) {
          for (const rel of slideSpecificRelsDoc.Relationships.Relationship) {
            const relType = rel.Type?.[0];
            const relTarget = rel.Target?.[0];

            if (
              relType ===
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' &&
              relTarget
            ) {
              // 6. 提取图片
              const imagePath = `ppt/media/${path.basename(relTarget)}`;
              const imageEntry = zip.getEntry(imagePath);

              if (imageEntry) {
                const fileExtension = path.extname(imagePath);
                const newFilename = `${String(i + 1).padStart(
                  3,
                  '0'
                )}${fileExtension}`;
                const outputPath = path.join(finalImageOutputPath, newFilename);
                fs.writeFileSync(outputPath, imageEntry.getData());
              }
            }
          }
        }
      }
    }

    const extractedImages = fs.readdirSync(finalImageOutputPath);
    if (extractedImages.length > 0) {
      console.log(
        `媒体文件已成功提取并保存到 ${finalImageOutputPath}，共提取 ${extractedImages.length} 个图片文件。`
      );
    } else {
      console.log(`在 PPTX 中未找到可提取的幻灯片图片。`);
    }
  } catch (error) {
    console.error('从 PPTX 直接提取媒体文件过程中发生错误:', error);
    // Create an empty 'images' directory if it doesn't exist so notes extraction doesn't fail
    if (!fs.existsSync(finalImageOutputPath)) {
      try {
        fs.mkdirSync(finalImageOutputPath, { recursive: true });
      } catch (dirError) {
        /* ignore */
      }
    }
    throw error; // Re-throw to be caught by runCli
  }
}

async function convertKeynoteToPptx(
  keynoteFilePath: string,
  pptxFilePath: string
): Promise<string> {
  const appleScriptPath = path.join(
    __dirname,
    'convert_keynote_to_pptx.applescript'
  );

  // 复制 AppleScript 文件到输出目录
  const outputDir = path.dirname(pptxFilePath);
  const tempAppleScriptPath = path.join(
    outputDir,
    'convert_keynote_to_pptx.applescript'
  );
  fs.copyFileSync(appleScriptPath, tempAppleScriptPath);

  const command = `osascript "${tempAppleScriptPath}" "${keynoteFilePath}" "${pptxFilePath}"`;

  try {
    console.log(`正在执行 Keynote 到 PPTX 转换 AppleScript: ${command}`);
    const { stdout, stderr } = await execPromise(command);

    let scriptError = false;
    if (stderr && stderr.trim() !== '') {
      if (stdout.trim().toLowerCase().startsWith('成功将')) {
        console.log(`AppleScript stderr (可能是警告或提示): ${stderr}`);
      } else {
        console.warn(`AppleScript stderr (可能包含错误信息): ${stderr}`);
        scriptError = true;
      }
    }
    if (stdout.startsWith('错误：') || stdout.startsWith('AppleScript 错误:')) {
      scriptError = true;
    }
    if (!stdout.trim().toLowerCase().startsWith('成功将') && scriptError) {
      throw new Error(
        `AppleScript Keynote 到 PPTX 转换失败: ${
          stdout.trim() || stderr.trim()
        }`
      );
    }
    if (
      !stdout.trim().toLowerCase().startsWith('成功将') &&
      !scriptError &&
      stderr.trim() === ''
    ) {
      throw new Error(
        `AppleScript Keynote 到 PPTX 转换未成功 (stdout): ${stdout.trim()}`
      );
    }
    console.log(stdout.trim() || `AppleScript Keynote 到 PPTX 转换已执行。`);
    // 删除临时 AppleScript 文件
    if (fs.existsSync(tempAppleScriptPath)) {
      fs.unlinkSync(tempAppleScriptPath);
    }
    return stdout.trim();
  } catch (error) {
    console.error('Keynote 到 PPTX 转换过程中发生错误:', error);
    // 确保在错误情况下也删除临时 AppleScript 文件
    if (fs.existsSync(tempAppleScriptPath)) {
      try {
        fs.unlinkSync(tempAppleScriptPath);
      } catch (cleanupError) {
        // Log cleanup error but re-throw original
      }
    }
    throw error;
  }
}

export async function runCli() {
  const inputFilePath = process.argv[2];
  const outputDirInput = process.argv[3];

  if (!inputFilePath) {
    console.error(
      '错误：请提供 Keynote 或 PPTX 文件路径作为第一个命令行参数。'
    );
    console.log('用法: npx keynote-ppt-to-markdown <文件路径> [输出目录]');
    process.exit(1);
  }

  const inputFilePathAbs = path.resolve(inputFilePath);
  if (!fs.existsSync(inputFilePathAbs)) {
    console.error(`错误：输入文件 "${inputFilePathAbs}" 未找到。`);
    process.exit(1);
  }

  const baseName = path.basename(
    inputFilePathAbs,
    path.extname(inputFilePathAbs)
  );
  const fileExt = path.extname(inputFilePathAbs).toLowerCase();
  const isPptx = fileExt === '.pptx';
  const isKeynote = fileExt === '.key';

  if (!isPptx && !isKeynote) {
    console.error(
      '错误：输入文件必须是 Keynote (.key) 或 PowerPoint (.pptx) 文件。'
    );
    process.exit(1);
  }

  const outputDir = outputDirInput
    ? path.resolve(outputDirInput)
    : path.dirname(inputFilePathAbs);

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Ensure the 'images' directory exists before note extraction, even if image export fails
  const imagesOutputPath = path.join(outputDir, 'images');
  if (!fs.existsSync(imagesOutputPath)) {
    fs.mkdirSync(imagesOutputPath, { recursive: true });
  }

  let pptxFilePathForNotes = ''; // This will be the .pptx file used for note extraction

  try {
    if (isKeynote) {
      if (process.platform !== 'darwin') {
        console.error('错误：处理 Keynote (.key) 文件仅在 macOS 上受支持。');
        process.exit(1);
      }
      // Keynote file processing (macOS only)
      const generatedPptxPath = path.join(outputDir, `${baseName}.pptx`);
      console.log(`正在将 Keynote 转换为 PPTX...`);
      await convertKeynoteToPptx(inputFilePathAbs, generatedPptxPath);
      console.log(`Keynote 已成功转换为 PPTX: ${generatedPptxPath}`);
      pptxFilePathForNotes = generatedPptxPath;

      console.log(`正在从 Keynote 导出幻灯片为图片...`);
      await exportSlidesAsImagesViaAppleScript(inputFilePathAbs, outputDir);
    } else {
      // isPptx
      console.log(`检测到 PPTX 文件。`);
      pptxFilePathForNotes = inputFilePathAbs;

      if (process.platform === 'darwin') {
        console.log(
          `检测到 macOS 系统，将通过 AppleScript 从 PPTX 导出幻灯片整页图片...`
        );
        await exportSlidesAsImagesViaAppleScript(inputFilePathAbs, outputDir);
      } else {
        console.log(
          `检测到非 macOS 系统，将尝试从 PPTX 直接提取内嵌媒体文件...`
        );
        console.warn(
          `警告：此方法提取的是 PPTX 文件中内嵌的图片，可能并非完整的幻灯片图片。为了获得最佳的幻灯片图片导出效果 (整页导出)，请在 macOS 上运行此脚本。`
        );
        await extractMediaFromPptxDirectly(inputFilePathAbs, outputDir);
      }
    }

    console.log(`正在生成 Markdown 文件 (从 ${pptxFilePathForNotes})...`);
    const markdownContent = await extractNotesFromPptx(
      pptxFilePathForNotes,
      outputDir
    );
    fs.writeFileSync(path.join(outputDir, `${baseName}.md`), markdownContent);
    console.log(
      `Markdown 文件已生成: ${path.join(outputDir, `${baseName}.md`)}`
    );

    // Clean up generated PPTX for Keynote inputs
    if (
      isKeynote &&
      fs.existsSync(pptxFilePathForNotes) &&
      pptxFilePathForNotes !== inputFilePathAbs
    ) {
      fs.unlinkSync(pptxFilePathForNotes);
      console.log(`已清理临时生成的 PPTX 文件: ${pptxFilePathForNotes}`);
    }
  } catch (error) {
    console.error('处理过程中发生主要错误:', error);
    // Attempt to clean up generated PPTX even if there was an error during processing a .key file
    if (
      isKeynote &&
      pptxFilePathForNotes &&
      fs.existsSync(pptxFilePathForNotes) &&
      pptxFilePathForNotes !== inputFilePathAbs
    ) {
      try {
        fs.unlinkSync(pptxFilePathForNotes);
        console.log(
          `错误后已清理临时生成的 PPTX 文件: ${pptxFilePathForNotes}`
        );
      } catch (cleanupError) {
        console.error(`错误后清理临时 PPTX 文件失败: ${cleanupError}`);
      }
    }
    process.exit(1);
  }
}
