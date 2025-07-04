import { useState } from "react";
import JSZip from "jszip";
import "./styles.css";

interface ProcessedFile {
  name: string;
  status: "idle" | "processing" | "completed" | "error";
  error?: string;
  downloadUrl?: string;
  message?: string;
}

export default function ImageResizer() {
  const [files, setFiles] = useState<File[]>([]);
  const [width, setWidth] = useState<string>("3.8");
  const [height, setHeight] = useState<string>("2.14");
  const [unit, setUnit] = useState<string>("cm");
  const [processedFiles, setProcessedFiles] = useState<ProcessedFile[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const selectedFiles = Array.from(e.target.files);
      const validFiles = selectedFiles.filter(file => /\.(doc|docx)$/i.test(file.name));
      
      if (validFiles.length === 0) {
        alert('请选择 .doc 或 .docx 文件');
        return;
      }
      
      if (validFiles.length > 30) {
        alert('最多只能选择30个文件');
        return;
      }
      
      setFiles(validFiles);
      setProcessedFiles([]);
    }
  };

  const convertToPixels = (value: number, unit: string): number => {
    switch (unit) {
      case "cm":
        return Math.round(value * 37.8); // 1cm ≈ 37.8 pixels at 96 DPI
      case "mm":
        return Math.round(value * 3.78); // 1mm ≈ 3.78 pixels at 96 DPI
      case "in":
        return Math.round(value * 96); // 1 inch = 96 pixels at 96 DPI
      case "px":
        return Math.round(value);
      default:
        return Math.round(value * 37.8);
    }
  };

  const resizeImage = async (imageBuffer: ArrayBuffer, targetWidth: number, targetHeight: number, imagePath: string): Promise<ArrayBuffer> => {
    return new Promise((resolve, reject) => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      const img = new Image();
      
      // Determine the original format
      const extension = imagePath.toLowerCase().split('.').pop();
      let mimeType = 'image/jpeg';
      let quality = 0.95;
      
      switch (extension) {
        case 'png':
          mimeType = 'image/png';
          quality = 1.0; // PNG doesn't use quality, but set to max
          break;
        case 'jpg':
        case 'jpeg':
          mimeType = 'image/jpeg';
          quality = 0.95;
          break;
        case 'gif':
          mimeType = 'image/png'; // Convert GIF to PNG to preserve transparency
          quality = 1.0;
          break;
        case 'bmp':
          mimeType = 'image/png'; // Convert BMP to PNG
          quality = 1.0;
          break;
        default:
          mimeType = 'image/jpeg';
          quality = 0.95;
      }
      
      img.onload = () => {
        const originalWidth = img.naturalWidth;
        const originalHeight = img.naturalHeight;
        
        console.log(`Original image size: ${originalWidth}x${originalHeight}, Target: ${targetWidth}x${targetHeight}`);
        
        canvas.width = targetWidth;
        canvas.height = targetHeight;
        
        if (ctx) {
          // Clear the canvas
          ctx.clearRect(0, 0, targetWidth, targetHeight);
          
          // Set high quality rendering
          ctx.imageSmoothingEnabled = true;
          ctx.imageSmoothingQuality = 'high';
          
          // Draw the resized image
          ctx.drawImage(img, 0, 0, targetWidth, targetHeight);
          
          console.log(`Drawing image to canvas: ${targetWidth}x${targetHeight}, format: ${mimeType}`);
          
          canvas.toBlob((blob) => {
            if (blob) {
              console.log(`Converted to blob: ${blob.size} bytes, type: ${blob.type}`);
              blob.arrayBuffer().then((resizedBuffer) => {
                console.log(`Final buffer size: ${resizedBuffer.byteLength} bytes (original: ${imageBuffer.byteLength} bytes)`);
                resolve(resizedBuffer);
              }).catch(reject);
            } else {
              reject(new Error('Failed to create blob'));
            }
          }, mimeType, quality);
        } else {
          reject(new Error('Could not get canvas context'));
        }
      };
      
      img.onerror = (error) => {
        console.error('Image load error:', error);
        reject(new Error('Failed to load image'));
      };
      
             const blob = new Blob([imageBuffer]);
       const imageUrl = URL.createObjectURL(blob);
       console.log(`Loading image from blob URL, original size: ${imageBuffer.byteLength} bytes`);
       img.src = imageUrl;
       
       // Store the original onload function
       const originalOnload = img.onload;
       img.onload = (event) => {
         URL.revokeObjectURL(imageUrl);
         if (originalOnload) {
           originalOnload.call(img, event);
         }
       };
    });
  };

  const processDocFiles = async () => {
    if (files.length === 0) return;

    setIsProcessing(true);
    
    // Initialize processing status for all files
    const initialProcessedFiles = files.map(file => ({
      name: file.name,
      status: "processing" as const
    }));
    setProcessedFiles(initialProcessedFiles);

    const targetWidth = convertToPixels(parseFloat(width), unit);
    const targetHeight = convertToPixels(parseFloat(height), unit);

    console.log(`Target size: ${targetWidth}x${targetHeight} pixels`);

    // Process each file
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      
      try {
        // Check if it's a .doc file (binary format)
        if (file.name.toLowerCase().endsWith('.doc')) {
          setProcessedFiles(prev => prev.map((pf, index) => 
            index === i ? {
              ...pf,
              status: "error",
              error: "暂不支持 .doc 格式，请将文件另存为 .docx 格式后重试"
            } : pf
          ));
          continue;
        }

        // Read the DOCX file as ZIP
        const zip = new JSZip();
        const zipContent = await zip.loadAsync(file);

        console.log(`Processing ${file.name}...`);
        
        // Find all image files in the document
        const mediaFiles: string[] = [];
        
        zipContent.forEach((relativePath) => {
          if ((relativePath.startsWith('word/media/') || relativePath.startsWith('word/embeddings/')) && 
              (relativePath.toLowerCase().endsWith('.jpg') || relativePath.toLowerCase().endsWith('.jpeg') || 
               relativePath.toLowerCase().endsWith('.png') || relativePath.toLowerCase().endsWith('.gif') || 
               relativePath.toLowerCase().endsWith('.bmp') || relativePath.toLowerCase().endsWith('.tiff') ||
               relativePath.toLowerCase().endsWith('.svg') || relativePath.toLowerCase().endsWith('.emf') ||
               relativePath.toLowerCase().endsWith('.wmf'))) {
            mediaFiles.push(relativePath);
          }
        });

        console.log(`Found ${mediaFiles.length} image files in ${file.name}`);

        let processedImages = 0;
        let failedImages = 0;

        // Convert target dimensions to EMUs (English Metric Units) for Word XML
        const targetWidthEMUs = Math.round(targetWidth * 9525);
        const targetHeightEMUs = Math.round(targetHeight * 9525);

        if (mediaFiles.length > 0) {
          // Process each image
          for (const mediaPath of mediaFiles) {
            const imageFile = zipContent.file(mediaPath);
            if (imageFile) {
              try {
                const imageBuffer = await imageFile.async('arraybuffer');
                
                // Skip SVG, EMF, WMF files as they can't be resized with canvas
                if (mediaPath.toLowerCase().endsWith('.svg') || 
                    mediaPath.toLowerCase().endsWith('.emf') || 
                    mediaPath.toLowerCase().endsWith('.wmf')) {
                  console.log(`Skipping vector format: ${mediaPath}`);
                  continue;
                }
                
                const resizedImageBuffer = await resizeImage(imageBuffer, targetWidth, targetHeight, mediaPath);
                zipContent.file(mediaPath, resizedImageBuffer);
                processedImages++;
              } catch (error) {
                console.error(`Failed to resize image ${mediaPath}:`, error);
                failedImages++;
              }
            }
          }

          // Update Word document XML to modify image dimensions
          try {
            const documentXmlFile = zipContent.file('word/document.xml');
            if (documentXmlFile) {
              let documentXml = await documentXmlFile.async('string');
              
              documentXml = documentXml.replace(
                /<wp:extent\s+cx="\d+"\s+cy="\d+"/g,
                `<wp:extent cx="${targetWidthEMUs}" cy="${targetHeightEMUs}"`
              );
              
              documentXml = documentXml.replace(
                /<a:ext\s+cx="\d+"\s+cy="\d+"/g,
                `<a:ext cx="${targetWidthEMUs}" cy="${targetHeightEMUs}"`
              );
              
              documentXml = documentXml.replace(
                /<wp:inline[^>]*><wp:extent\s+cx="\d+"\s+cy="\d+"/g,
                (match) => {
                  return match.replace(
                    /cx="\d+"\s+cy="\d+"/,
                    `cx="${targetWidthEMUs}" cy="${targetHeightEMUs}"`
                  );
                }
              );
              
              zipContent.file('word/document.xml', documentXml);
            }
          } catch (error) {
            console.error('Failed to update document.xml:', error);
          }

          // Update headers/footers if they contain images
          const headerFooterFiles = ['word/header1.xml', 'word/header2.xml', 'word/header3.xml', 'word/header4.xml',
                                    'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml', 'word/footer4.xml'];
          
          for (const fileName of headerFooterFiles) {
            try {
              const xmlFile = zipContent.file(fileName);
              if (xmlFile) {
                let xmlContent = await xmlFile.async('string');
                
                xmlContent = xmlContent.replace(
                  /<wp:extent\s+cx="\d+"\s+cy="\d+"/g,
                  `<wp:extent cx="${targetWidthEMUs}" cy="${targetHeightEMUs}"`
                );
                
                xmlContent = xmlContent.replace(
                  /<a:ext\s+cx="\d+"\s+cy="\d+"/g,
                  `<a:ext cx="${targetWidthEMUs}" cy="${targetHeightEMUs}"`
                );
                
                zipContent.file(fileName, xmlContent);
              }
            } catch (error) {
              console.warn(`Failed to update ${fileName}:`, error);
            }
          }
        }

        // Generate the modified DOCX file
        const modifiedDocx = await zipContent.generateAsync({ type: "blob" });
        const downloadUrl = URL.createObjectURL(modifiedDocx);
        
        let resultMessage = '';
        if (processedImages === 0 && mediaFiles.length === 0) {
          resultMessage = '文档中未找到可处理的图片，返回原始文档。';
        } else if (processedImages === 0 && mediaFiles.length > 0) {
          resultMessage = `找到 ${mediaFiles.length} 个图片文件，但都无法处理（可能是矢量格式），返回原始文档。`;
        } else {
          resultMessage = `成功处理 ${processedImages} 个图片，调整为 ${width}×${height}${unit} 尺寸。`;
          if (failedImages > 0) {
            resultMessage += ` ${failedImages} 个图片处理失败，保持原始尺寸。`;
          }
        }
        
        // Update the specific file's status
        setProcessedFiles(prev => prev.map((pf, index) => 
          index === i ? {
            ...pf,
            status: "completed",
            downloadUrl,
            message: resultMessage
          } : pf
        ));

      } catch (error) {
        console.error(`Error processing file ${file.name}:`, error);
        setProcessedFiles(prev => prev.map((pf, index) => 
          index === i ? {
            ...pf,
            status: "error",
            error: error instanceof Error ? error.message : "处理文件时发生错误"
          } : pf
        ));
      }
    }

    setIsProcessing(false);
  };



  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (files.length === 0) {
      alert('请先选择文件');
      return;
    }
    if (!width || !height || parseFloat(width) <= 0 || parseFloat(height) <= 0) {
      alert('请输入有效的尺寸');
      return;
    }
    processDocFiles();
  };

  const downloadAll = async () => {
    const completedFiles = processedFiles.filter((pf, index) => 
      pf.status === "completed" && pf.downloadUrl
    );
    
    if (completedFiles.length === 0) return;
    
    // Add loading state for download all
    setIsProcessing(true);
    
    try {
      for (let i = 0; i < completedFiles.length; i++) {
        const processedFile = completedFiles[i];
        const originalIndex = processedFiles.findIndex(pf => pf === processedFile);
        
        if (processedFile.downloadUrl) {
          const fileName = files[originalIndex].name.replace(/\.docx$/i, `_resized_${width}x${height}${unit}.docx`);
          const link = document.createElement('a');
          link.href = processedFile.downloadUrl;
          link.download = fileName;
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          
          // Add delay between downloads to prevent browser blocking
          // Only add delay if there are more files to download
          if (i < completedFiles.length - 1) {
            await new Promise(resolve => setTimeout(resolve, 500)); // 500ms delay
          }
        }
      }
      
      alert(`成功开始下载 ${completedFiles.length} 个文件！`);
    } catch (error) {
      console.error('Download error:', error);
      alert('下载过程中出现错误，请重试');
    } finally {
      setIsProcessing(false);
    }
  };

  const renderStatusBadge = (status: string, error?: string) => {
    switch (status) {
      case "processing":
        return <span className="status-badge processing">处理中...</span>;
      case "completed":
        return <span className="status-badge completed">完成</span>;
      case "error":
        return (
          <span className="status-badge error" title={error}>
            错误
          </span>
        );
      default:
        return <span className="status-badge idle">待处理</span>;
    }
  };

  return (
    <div className="app">
      <div className="container">
        <h1>文档图片尺寸调整工具</h1>
        <p className="description">
          上传 Word 文档（.doc/.docx），最多可选择30个文件，调整文档中所有图片的尺寸。如果文档没有图片，将返回原始文档。
        </p>

        <form onSubmit={handleSubmit} className="form">
          <div className="form-group">
            <label htmlFor="file-input">选择文档文件 (.doc/.docx) - 最多30个文件:</label>
            <input
              id="file-input"
              type="file"
              accept=".doc,.docx"
              onChange={handleFileChange}
              className="file-input"
              multiple
            />
            {files.length > 0 && (
              <div className="selected-files">
                <p>已选择 {files.length} 个文件:</p>
                <ul>
                  {files.map((file, index) => (
                    <li key={index}>{file.name}</li>
                  ))}
                </ul>
              </div>
            )}
          </div>

          <div className="form-group">
            <label>目标图片尺寸:</label>
            <div className="size-inputs">
              <input
                type="number"
                value={width}
                onChange={(e) => setWidth(e.target.value)}
                placeholder="宽度"
                min="0.01"
                step="0.01"
                className="size-input"
              />
              <span>×</span>
              <input
                type="number"
                value={height}
                onChange={(e) => setHeight(e.target.value)}
                placeholder="高度"
                min="0.01"
                step="0.01"
                className="size-input"
              />
              <select
                value={unit}
                onChange={(e) => setUnit(e.target.value)}
                className="unit-select"
              >
                <option value="cm">厘米 (cm)</option>
                <option value="mm">毫米 (mm)</option>
                <option value="in">英寸 (in)</option>
                <option value="px">像素 (px)</option>
              </select>
            </div>
          </div>

          <button
            type="submit"
            className="submit-button"
            disabled={files.length === 0 || isProcessing}
          >
            {isProcessing ? "处理中..." : "开始处理"}
          </button>
        </form>

        {processedFiles.length > 0 && (
          <div className="results">
            <div className="results-header">
              <h2>处理结果</h2>
              {processedFiles.some(pf => pf.status === "completed") && (
                <button 
                  onClick={downloadAll}
                  className="download-all-button"
                  disabled={isProcessing}
                >
                  下载所有完成的文件
                </button>
              )}
            </div>
            {processedFiles.map((processedFile, index) => (
              <div key={index} className="file-result">
                <div className="file-info">
                  <span className="file-name">{processedFile.name}</span>
                  {renderStatusBadge(processedFile.status, processedFile.error)}
                </div>
                {processedFile.error && (
                  <div className="error-message">
                    错误: {processedFile.error}
                  </div>
                )}
                {processedFile.status === "completed" && (
                  <div className="success-message">
                    {processedFile.message || `如果文档中包含图片，所有图片已调整为 ${width}×${height}${unit} 尺寸。`}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        <div className="info-section">
          <h3>使用说明</h3>
          <ul>
            <li>支持 .doc 和 .docx 格式的 Word 文档，最多可选择30个文件</li>
            <li>会自动识别并调整文档中的所有图片尺寸</li>
            <li>如果文档中没有图片，将返回原始文档</li>
            <li>可以使用"下载所有完成的文件"按钮批量下载处理完成的文件</li>
            <li>处理后的文件名包含尺寸信息</li>
            <li>支持厘米、毫米、英寸和像素单位</li>
          </ul>
        </div>
      </div>
    </div>
  );
} 