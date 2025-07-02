import { useState, useEffect } from "react";
import * as mammoth from "mammoth";
import * as ExcelJS from "exceljs";
import "./styles.css";

// Add interface for per-file settings
interface FileSettings {
  region: string;
  plannedCheckDate: string;
  actualCheckDate: string;
  town: string;
}

export default function SafetyReportPage() {
  const [files, setFiles] = useState<File[]>([]);
  // Add state for per-file settings
  const [fileSettings, setFileSettings] = useState<Record<number, FileSettings>>({});
  const [responses, setResponses] = useState<
    Array<{
      fileName: string;
      annex1: string;
      annex2: string;
      status: "idle" | "processing" | "completed" | "error";
      error?: string;
    }>
  >([]);
  const [isLoading, setIsLoading] = useState(false);
  const [globalError, setGlobalError] = useState<string | null>(null);

  // State for user inputs
  const [region, setRegion] = useState("张家港");
  const [town, setTown] = useState("经开区");
  const [factoryMappingText, setFactoryMappingText] = useState("");
  const [checkDate, setCheckDate] = useState(
    new Date().toLocaleDateString("zh-CN")
  );

  const [copiedAnnex1, setCopiedAnnex1] = useState(false);
  const [copiedAnnex2, setCopiedAnnex2] = useState(false);

  // 定义表头常量
  const annex1Header =
    "区域\t镇/街道\t出租方名称\t承租方名称\t计划核查时间\t实际核查时间";
  const annex2Header =
    "核查机构名称\t地区\t厂中厂名称\t核查时间\t存在问题\t重大隐患数量\t一般隐患数量\t隐患总数量\t现场隐患\t管理隐患\t是否属于涉爆粉尘、金属熔融企业";

  // Helper function to get file settings with defaults
  const getFileSettings = (index: number): FileSettings => {
    return fileSettings[index] || {
      region: region,
      plannedCheckDate: checkDate,
      actualCheckDate: checkDate,
      town: town
    };
  };

  // Helper function to update file settings
  const updateFileSetting = (index: number, field: keyof FileSettings, value: string) => {
    setFileSettings(prev => ({
      ...prev,
      [index]: {
        ...getFileSettings(index),
        [field]: value
      }
    }));
  };

  // Update file settings when global values change
  useEffect(() => {
    if (files.length > 0) {
      setFileSettings(prev => {
        const updated: Record<number, FileSettings> = {};
        files.forEach((_, index) => {
          const currentSettings = prev[index];
          // Only update if the setting hasn't been manually changed
          // We consider it manually changed if it differs from the old global values
          updated[index] = {
            region: currentSettings?.region || region,
            plannedCheckDate: currentSettings?.plannedCheckDate || checkDate,
            actualCheckDate: currentSettings?.actualCheckDate || checkDate,
            town: currentSettings?.town || town
          };
        });
        return updated;
      });
    }
  }, [region, checkDate, town, files.length]);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const newFiles = Array.from(e.target.files).slice(0, 30);
      setFiles(newFiles);

             // Initialize file settings with current global values
       const newFileSettings: Record<number, FileSettings> = {};
       newFiles.forEach((_, index) => {
         newFileSettings[index] = {
           region: region,
           plannedCheckDate: checkDate,
           actualCheckDate: checkDate,
           town: town
         };
       });
      setFileSettings(newFileSettings);

      setResponses(
        newFiles.map((file) => ({
          fileName: file.name,
          annex1: "",
          annex2: "",
          status: "idle",
        }))
      );
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (files.length === 0) return;

    setIsLoading(true);
    setGlobalError(null);
    setResponses((prev) =>
      prev.map((resp) => ({
        ...resp,
        status: "processing",
        error: undefined,
      }))
    );

    try {
      await Promise.all(
        files.map(async (file, index) => {
          try {
            // 从DOCX中提取文本
            const extractedText = await extractTextFromDocx(file);

            // Get file-specific settings
            const currentFileSettings = getFileSettings(index);

            // 准备prompt
            const prompt = `请根据以下文档内容和安全检查信息，严格按照以下格式返回数据：

              第一部分（附件1）：
              ${annex1Header}具体数据行（用制表符分隔）
              

              第二部分（附件2）：
              ${annex2Header}具体数据行（用制表符分隔）
              

              要求：
              1. 附件1和附件2都必须用制表符(\t)分隔各列
              2. 附件1必须包含：区域、镇/街道、出租方名称、承租方名称、计划核查时间、实际核查时间
              3. 附件2必须包含：核查机构名称、地区、厂中厂名称、核查时间、存在问题、重大隐患数量、一般隐患数量、隐患总数量、现场隐患、管理隐患、是否属于涉爆粉尘金属熔融企业
              4. 核查机构名称固定为"常州市安平安全技术服务有限公司"
              5. 重大隐患数量默认为0
              6. 是否属于涉爆粉尘、金属熔融企业默认为"否"
              7. 存在问题需要按照 出租方：xxx 1、2、3、 承租方：xxx 1、2、3、4、 的格式来生成（不要加其他词）
              8. 一般隐患数量=隐患总数量，现场隐患=承租方问题数量，管理隐患=出租方问题数量
              

              当前信息：
              - 区域: ${currentFileSettings.region}
              - 镇/街道: ${currentFileSettings.town}
              - 预计核查/核查时间: ${currentFileSettings.plannedCheckDate}
              - 实际核查时间: ${currentFileSettings.actualCheckDate}
              - 厂中厂映射关系: ${factoryMappingText || "无"},假如在这个列表里找不到厂中厂，则使用出租方名称当作厂中厂名称

              文档内容：
              ${extractedText.substring(0, 40000)}

              在返回时，请在第一部分和第二部分之间添加一行，内容为四个大写字母：XXXX

              返回例子1：
              张家港	经开区	张家港市杨舍镇农联村股份经济合作社	张家港市海达兴纺机有限公司	2025/6/23	2025/6/23
              XXXX
              常州市安平安全技术服务有限公司	张家港	农联村村级租用厂房	2025.6.23	出租方：张家港市杨舍镇农联村股份经济合作社 1、8楼安全出口指示灯不亮 承租方：苏州凡赛特材料科技有限公司1、9楼安全出口指示灯不亮 2、消火栓箱未见点检记录 3、消火栓箱未张贴操作说明 4、注塑机安全风险告知牌未划分风险等级和未明确管理责任人员	0	15	15	14	1	否

              返回例子2：
              张家港	经开区	张家港市杨舍镇徐丰村股份经济合作社	张家港市创新线业有限公司	2025/6/23	2025/6/23"
              XXXX
              常州市安平安全技术服务有限公司	张家港	徐丰村村级租用厂房	2025.6.23	出租方：张家港市杨舍镇徐丰村股份经济合作社 1、出租方公告栏内各企业较大风险未更新 2、出租方公告栏内各企业安全风险四色图未更新 承租方;张家港市创新线业有限公司 1、货架未见限重标识 2、消火栓箱内放置灭火器 3、车间内通道堵塞 4、配电盒未张贴警示标识 5、电缆槽盒未跨接 6、绝缘胶垫未见检测合格标签 7、灭火器箱前堆放杂物 8、防腐剂放置点未见MSDS 9、较大风险公告栏未及时更新 10、清洁剂使用完放置在车间现场	0	12	12	10	2	否
              
              //所以你的回答只应该有像这样的三行，不要再有其他东西了
              `;

            // 准备API负载
            const payload = {
              model: "deepseek-chat",
              messages: [
                {
                  role: "system",
                  content:
                    "你是一个严格遵循指令的数据生成器，必须返回符合要求的文本格式，使用XXXX分隔两部分内容。",
                },
                {
                  role: "user",
                  content: prompt,
                },
              ],
              temperature: 0.1,
              max_tokens: 2000,
            };

            const apiResponse = await fetch(
              "https://api.deepseek.com/v1/chat/completions",
              {
                method: "POST",
                headers: {
                  "Content-Type": "application/json",
                  Authorization: `Bearer sk-dedd19d6c11846b8b7f2fc08e9be60de`,
                },
                body: JSON.stringify(payload),
              }
            );

            if (!apiResponse.ok) {
              const errorData = await apiResponse.json();
              throw new Error(
                `API请求失败: ${apiResponse.status} ${apiResponse.statusText} - ${JSON.stringify(errorData)}`
              );
            }

            const data = await apiResponse.json();
            const responseText = data.choices[0].message.content;

            // 使用"XXXX"分割响应内容
            const parts = responseText.split("XXXX");
            if (parts.length !== 2) {
              throw new Error(`响应格式错误: 未找到XXXX分隔符或找到多个分隔符`);
            }

            // 提取附件1和附件2内容
            const annex1 = parts[0].trim();
            const annex2 = parts[1].trim();

            // 验证内容格式
            if (!annex1 || !annex2) {
              throw new Error("响应内容不完整");
            }

            setResponses((prev) =>
              prev.map((resp, i) =>
                i === index
                  ? {
                      ...resp,
                      annex1,
                      annex2,
                      status: "completed",
                    }
                  : resp
              )
            );
          } catch (err) {
            setResponses((prev) =>
              prev.map((resp, i) =>
                i === index
                  ? {
                      ...resp,
                      status: "error",
                      error: err instanceof Error ? err.message : "处理失败",
                    }
                  : resp
              )
            );
          }
        })
      );
    } catch (err) {
      setGlobalError(err instanceof Error ? err.message : "发生未知错误");
      console.error("处理失败:", err);
    } finally {
      setIsLoading(false);
    }
  };

  // Extract text from DOCX file
  const extractTextFromDocx = async (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          if (event.target?.result) {
            const arrayBuffer = event.target.result as ArrayBuffer;
            const result = await mammoth.extractRawText({ arrayBuffer });
            resolve(result.value);
          } else {
            reject(new Error("文件读取失败"));
          }
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  };

  // Excel generation and download functions
  const downloadExcelFile = async () => {
    // Create rich text with bold landlord/tenant names and company names
    const createRichTextForIssues = (issuesText: string) => {
      if (!issuesText) return "";
      
      const richText: any[] = [];
      
      // Split by 出租方 and 承租方 to process each section
      const parts = issuesText.split(/(出租方：|承租方：)/);
      
      for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        
        if (part === "出租方：" || part === "承租方：") {
          // Add bold landlord/tenant label
          richText.push({
            text: part,
            font: { name: "宋体", size: 9, bold: true }
          });
        } else if (part.trim()) {
          // Process the content after landlord/tenant label
          let content = part.trim();
          
          // Extract company name (text before first numbered item)
          const companyMatch = content.match(/^([^1-9]*?)(\s*\d+、)/);
          let companyName = "";
          let issues = content;
          
          if (companyMatch) {
            companyName = companyMatch[1].trim();
            issues = content.substring(companyMatch[1].length).trim();
          } else {
            // If no numbered items found, treat entire content as company name
            companyName = content;
            issues = "";
          }
          
          // Add bold company name
          if (companyName) {
            richText.push({
              text: companyName,
              font: { name: "宋体", size: 9, bold: true }
            });
          }
          
          // Process numbered issues
          if (issues) {
            // Add line breaks before numbered items
            issues = issues.replace(/(\d+、)/g, '\n$1');
            
            // Split by lines to handle each line
            const lines = issues.split('\n');
            
            for (let j = 0; j < lines.length; j++) {
              const line = lines[j].trim();
              if (line) {
                richText.push({
                  text: '\n' + line,
                  font: { name: "宋体", size: 9 }
                });
              }
            }
          }
          
          // Add line break after section if not the last part
          if (i < parts.length - 1 && parts[i + 1] && (parts[i + 1] === "出租方：" || parts[i + 1] === "承租方：")) {
            richText.push({
              text: '\n',
              font: { name: "宋体", size: 9 }
            });
          }
        }
      }
      
      return { richText };
    };
    const workbook = new ExcelJS.Workbook();
    
    // Get completed responses
    const completedResponses = responses.filter(resp => resp.status === "completed");
    if (completedResponses.length === 0) return;

    // Create 附件1 worksheet
    const annex1Sheet = workbook.addWorksheet("附件1");
    
    // Add headers for 附件1
    const annex1Headers = ["区域", "镇/街道", "出租方名称", "承租方名称", "计划核查时间", "实际核查时间"];
    annex1Sheet.addRow(annex1Headers);
    
    // Style header row for 附件1
    const annex1HeaderRow = annex1Sheet.getRow(1);
    annex1HeaderRow.font = { name: "宋体", size: 9, bold: true, color: { argb: "FFFFFFFF" } };
    annex1HeaderRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2E86AB" }
    };
    annex1HeaderRow.alignment = { horizontal: "center", vertical: "middle" };
    
    // Add data rows for 附件1
    completedResponses.forEach((response) => {
      if (response.annex1) {
        // Split data directly since AI response contains only data rows (no headers)
        const lines = response.annex1.trim().split("\n");
        lines.forEach((line) => {
          if (line.trim()) {
            const row = line.split("\t");
            const dataRow = annex1Sheet.addRow(row);
            dataRow.alignment = { vertical: "middle", wrapText: true };
            // Set font for all cells in row
            dataRow.eachCell((cell: any) => {
              cell.font = { name: "宋体", size: 9 };
            });
          }
        });
      }
    });

    // Auto-size columns and add borders for 附件1
    annex1Headers.forEach((_, index) => {
      const column = annex1Sheet.getColumn(index + 1);
      column.width = 20;
      column.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    });

    // Add borders to all cells in 附件1
    annex1Sheet.eachRow((row: any, rowNumber: number) => {
      row.eachCell((cell: any) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
        if (rowNumber > 1) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: rowNumber % 2 === 0 ? "FFF8F9FA" : "FFFFFFFF" }
          };
          // Ensure all data cells have 宋体 font
          if (!cell.font || !cell.font.name) {
            cell.font = { name: "宋体", size: 9 };
          }
        }
      });
    });

    // Create 附件2 worksheet
    const annex2Sheet = workbook.addWorksheet("附件2");
    
    // Add headers for 附件2
    const annex2Headers = [
      "核查机构名称", "地区", "厂中厂名称", "核查时间", "存在问题", 
      "重大隐患数量", "一般隐患数量", "隐患总数量", "现场隐患", "管理隐患", 
      "是否属于涉爆粉尘、金属熔融企业"
    ];
    annex2Sheet.addRow(annex2Headers);
    
    // Style header row for 附件2
    const annex2HeaderRow = annex2Sheet.getRow(1);
    annex2HeaderRow.font = { name: "宋体", size: 9, bold: true, color: { argb: "FFFFFFFF" } };
    annex2HeaderRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF27AE60" }
    };
    annex2HeaderRow.alignment = { horizontal: "center", vertical: "middle" };
    
    // Add data rows for 附件2
    completedResponses.forEach((response) => {
      if (response.annex2) {
        // Split data directly since AI response contains only data rows (no headers)
        const lines = response.annex2.trim().split("\n");
        lines.forEach((line) => {
          if (line.trim()) {
            const row = line.split("\t");
            const dataRow = annex2Sheet.addRow(row);
            dataRow.alignment = { vertical: "middle", wrapText: true };
            
            // Set font for all cells in row
            dataRow.eachCell((cell: any) => {
              cell.font = { name: "宋体", size: 9 };
            });
            
            // Special formatting for the "存在问题" column (index 4)
            if (row[4]) {
              const issueCell = dataRow.getCell(5);
              issueCell.alignment = { vertical: "top", wrapText: true };
              
              // Create rich text with bold landlord/tenant names
              const issuesText = row[4];
              const richText = createRichTextForIssues(issuesText);
              issueCell.value = richText;
            }
          }
        });
      }
    });

    // Auto-size columns and add borders for 附件2
    annex2Headers.forEach((_, index) => {
      const column = annex2Sheet.getColumn(index + 1);
      if (index === 4) { // "存在问题" column
        column.width = 50;
      } else if (index < 4) {
        column.width = 20;
      } else {
        column.width = 15;
      }
    });

    // Add borders to all cells in 附件2
    annex2Sheet.eachRow((row: any, rowNumber: number) => {
      row.eachCell((cell: any) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
        if (rowNumber > 1) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: rowNumber % 2 === 0 ? "FFF8F9FA" : "FFFFFFFF" }
          };
          // Ensure all data cells have 宋体 font
          if (!cell.font || !cell.font.name) {
            cell.font = { name: "宋体", size: 9 };
          }
        }
      });
    });

    // Generate and download the file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { 
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
    });
    
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `安全检查报告_${new Date().toLocaleDateString("zh-CN").replace(/\//g, "-")}.xlsx`;
    link.click();
    window.URL.revokeObjectURL(url);
  };

  // Copy to clipboard functions
  const copyAnnex1ToClipboard = () => {
    const textToCopy = responses
      .filter((resp) => resp.annex1 && resp.status === "completed")
      .map((resp) => {
        // 如果内容不包含表头，则添加表头
        return resp.annex1;
      })
      .join("\n"); // 文件间用空行分隔

    if (!textToCopy) return;

    navigator.clipboard
      .writeText(textToCopy)
      .then(() => {
        setCopiedAnnex1(true);
        setTimeout(() => setCopiedAnnex1(false), 2000);
      })
      .catch((err) => {
        console.error("复制失败:", err);
        setGlobalError("复制附件1失败");
      });
  };

  const copyAnnex2ToClipboard = () => {
    const textToCopy = responses
      .filter((resp) => resp.annex2 && resp.status === "completed")
      .map((resp) => {
        // 如果内容不包含表头，则添加表头

        return resp.annex2;
      })
      .join("\n"); // 文件间用空行分隔

    if (!textToCopy) return;

    navigator.clipboard
      .writeText(textToCopy)
      .then(() => {
        setCopiedAnnex2(true);
        setTimeout(() => setCopiedAnnex2(false), 2000);
      })
      .catch((err) => {
        console.error("复制失败:", err);
        setGlobalError("复制附件2失败");
      });
  };

  // Parse tab-separated data for table display
  const parseTabData = (data: string) => {
    if (!data) return { headers: [], rows: [] };

    const lines = data.split("\n");
    const headers = lines[0]?.split("\t") || [];
    const rows = lines.slice(1).map((line) => line.split("\t"));

    return { headers, rows };
  };

  // Render status badge
  const renderStatusBadge = (status: string, error?: string) => {
    switch (status) {
      case "processing":
        return <span className="status-badge processing">处理中</span>;
      case "completed":
        return <span className="status-badge completed">完成</span>;
      case "error":
        return <span className="status-badge error">错误: {error}</span>;
      default:
        return <span className="status-badge">等待</span>;
    }
  };
  
  return (
    <div className="app">
      <div className="container">
        <h1>安全周报自动化系统</h1>

        <form onSubmit={handleSubmit}>
          <div className="input-section">
            <h2>基本信息填写</h2>

            <div className="input-group">
              <label>区域:</label>
              <input
                type="text"
                value={region}
                onChange={(e) => setRegion(e.target.value)}
                required
              />
            </div>

            <div className="input-group">
              <label>镇/街道:</label>
              <input
                type="text"
                value={town}
                onChange={(e) => setTown(e.target.value)}
                required
              />
            </div>

            <div className="input-group">
              <label>核查时间:</label>
              <input
                type="date"
                value={checkDate}
                onChange={(e) => setCheckDate(e.target.value)}
                required
              />
            </div>

            <div className="input-group">
              <label>厂中厂映射关系:</label>
              <textarea
                value={factoryMappingText}
                onChange={(e) => setFactoryMappingText(e.target.value)}
                placeholder="例如: 张家港市杨舍镇农联村股份经济合作社 → 农联村村级租用厂房"
                rows={3}
              />
            </div>
          </div>

          <div className="file-upload">
            <label htmlFor="file-upload">上传检查文档 (最多30个):</label>
            <input
              id="file-upload"
              type="file"
              accept=".docx,.doc"
              onChange={handleFileChange}
              disabled={isLoading}
              multiple
              required
            />
            {files.length > 0 && (
              <div className="file-list">
                <p>已选择文件 ({files.length}):</p>
                <div className="file-settings-container">
                  {files.map((file, index) => {
                    const settings = getFileSettings(index);
                    return (
                      <div key={index} className="file-item">
                        <div className="file-name">{file.name}</div>
                        <div className="file-settings">
                          <div className="setting-group">
                            <label>区域:</label>
                            <input
                              type="text"
                              value={settings.region}
                              onChange={(e) => updateFileSetting(index, 'region', e.target.value)}
                              disabled={isLoading}
                            />
                          </div>
                          <div className="setting-group">
                            <label>镇/街道:</label>
                            <input
                              type="text"
                              value={settings.town}
                              onChange={(e) => updateFileSetting(index, 'town', e.target.value)}
                              disabled={isLoading}
                            />
                          </div>
                          <div className="setting-group">
                            <label>计划核查时间:</label>
                            <input
                              type="date"
                              value={settings.plannedCheckDate}
                              onChange={(e) => updateFileSetting(index, 'plannedCheckDate', e.target.value)}
                              disabled={isLoading}
                            />
                          </div>
                          <div className="setting-group">
                            <label>实际核查时间:</label>
                            <input
                              type="date"
                              value={settings.actualCheckDate}
                              onChange={(e) => updateFileSetting(index, 'actualCheckDate', e.target.value)}
                              disabled={isLoading}
                            />
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>

          <button type="submit" disabled={files.length === 0 || isLoading}>
            {isLoading ? "处理中..." : "生成安全检查报告"}
          </button>
        </form>

        {globalError && (
          <div className="error">
            <h3>全局错误:</h3>
            <p>{globalError}</p>
          </div>
        )}

        {responses.length > 0 && (
          <div className="actions">
            <button
              className="excel-download-btn"
              onClick={downloadExcelFile}
              disabled={
                !responses.some((r) => r.status === "completed")
              }
            >
              📊 下载Excel文件
            </button>
            <button
              onClick={copyAnnex1ToClipboard}
              disabled={
                !responses.some((r) => r.annex1 && r.status === "completed")
              }
            >
              {copiedAnnex1 ? "已复制所有附件1!" : "复制所有附件1"}
            </button>
            <button
              onClick={copyAnnex2ToClipboard}
              disabled={
                !responses.some((r) => r.annex2 && r.status === "completed")
              }
            >
              {copiedAnnex2 ? "已复制所有附件2!" : "复制所有附件2"}
            </button>
          </div>
        )}

        <div className="results-container">
          {responses.map((response, index) => (
            <div key={index} className="file-result">
              <div className="file-header">
                <h3>
                  {response.fileName}
                  {renderStatusBadge(response.status, response.error)}
                </h3>
              </div>

              {response.status === "completed" && (
                <>
                  <div className="annex-section">
                    <div className="annex-header">
                      <h4>附件1: 基本信息</h4>
                    </div>
                    <div className="annex-content">
                      {parseTabData(response.annex1).headers.length > 0 ? (
                        <div className="table-container">
                          <table>
                            <thead>
                              <tr>
                                {parseTabData(response.annex1).headers.map(
                                  (header, i) => (
                                    <th key={i}>{header}</th>
                                  )
                                )}
                              </tr>
                            </thead>
                            <tbody>
                              {parseTabData(response.annex1).rows.map(
                                (row, rowIndex) => (
                                  <tr key={rowIndex}>
                                    {row.map((cell, cellIndex) => (
                                      <td key={cellIndex}>{cell}</td>
                                    ))}
                                  </tr>
                                )
                              )}
                            </tbody>
                          </table>
                        </div>
                      ) : (
                        <div className="raw-data">
                          <h5>原始数据:</h5>
                          <pre>{response.annex1}</pre>
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="annex-section">
                    <div className="annex-header">
                      <h4>附件2: 隐患详情</h4>
                    </div>
                    <div className="annex-content">
                      {parseTabData(response.annex2).headers.length > 0 ? (
                        <div className="table-container">
                          <table>
                            <thead>
                              <tr>
                                {parseTabData(response.annex2).headers.map(
                                  (header, i) => (
                                    <th key={i}>{header}</th>
                                  )
                                )}
                              </tr>
                            </thead>
                            <tbody>
                              {parseTabData(response.annex2).rows.map(
                                (row, rowIndex) => (
                                  <tr key={rowIndex}>
                                    {row.map((cell, cellIndex) => (
                                      <td
                                        key={cellIndex}
                                        className={
                                          cellIndex === 4 ? "issue-cell" : ""
                                        }
                                      >
                                        {cell}
                                      </td>
                                    ))}
                                  </tr>
                                )
                              )}
                            </tbody>
                          </table>
                        </div>
                      ) : (
                        <div className="raw-data">
                          <h5>原始数据:</h5>
                          <pre>{response.annex2}</pre>
                        </div>
                      )}
                    </div>
                  </div>
                </>
              )}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
} 