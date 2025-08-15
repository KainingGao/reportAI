import { useEffect, useState } from "react";
import * as mammoth from "mammoth";
import * as ExcelJS from "exceljs";
import "./styles.css";

// Add interface for the new mapping format (commented out for now)
/* interface ScheduleMapping {
  plannedDate: string;
  region: string;
  town: string;
  factoryName: string;
  tenant: string;
} */

export default function SafetyReportPage() {
  const [files, setFiles] = useState<File[]>([]);
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

  // Update state for user inputs - remove individual settings, keep only the new mapping
  const [scheduleMappingText, setScheduleMappingText] = useState("");

  // Per-file actual check dates (editable by user). Defaults to file last modified date
  const [actualDates, setActualDates] = useState<string[]>([]);
  const [activeAnnex, setActiveAnnex] = useState<"annex1" | "annex2">("annex1");
  const [smartSortActive, setSmartSortActive] = useState<boolean>(false);

  // Ordering state
  const [originalOrder, setOriginalOrder] = useState<number[]>([]);
  const [displayOrder, setDisplayOrder] = useState<number[]>([]);
  const [sortMode, setSortMode] = useState<"original" | "asc" | "desc">("original");
  const [draggingIndex, setDraggingIndex] = useState<number | null>(null);

  

  // 定义表头常量
  const annex1Header =
    "区域\t镇/街道\t出租方名称\t承租方名称\t计划核查时间\t实际核查时间";
  const annex2Header =
    "核查机构名称\t地区\t厂中厂名称\t核查时间\t存在问题\t重大隐患数量\t一般隐患数量\t隐患总数量\t现场隐患\t管理隐患\t是否属于涉爆粉尘、金属熔融企业";

  // Headers for consolidated table view
  const ANNEX1_HEADERS = [
    "区域",
    "镇/街道",
    "出租方名称",
    "承租方名称",
    "计划核查时间",
    "实际核查时间",
  ];
  const ANNEX2_HEADERS = [
    "核查机构名称",
    "地区",
    "厂中厂名称",
    "核查时间",
    "存在问题",
    "重大隐患数量",
    "一般隐患数量",
    "隐患总数量",
    "现场隐患",
    "管理隐患",
    "是否属于涉爆粉尘、金属熔融企业",
  ];

  // Helper to parse issues into rich JSX for frontend display (match Excel's layout intent)
  const renderIssuesContent = (issuesText: string) => {
    if (!issuesText) return null;
    const parts = issuesText.split(/(出租方：|承租方：)/);
    const sections: Array<{ label: string; company: string; items: string[] }> = [];
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      if (part === "出租方：" || part === "承租方：") {
        const content = (parts[i + 1] || "").trim();
        // Extract company name (before first numbered item like n、)
        const companyMatch = content.match(/^([^1-9]*?)(\s*\d+、)/);
        let company = "";
        let rest = content;
        if (companyMatch) {
          company = (companyMatch[1] || "").trim();
          rest = content.substring(companyMatch[1].length).trim();
        }
        // Split numbered items
        const items = rest
          ? rest.replace(/(\d+、)/g, "\n$1").split("\n").map((s) => s.trim()).filter(Boolean)
          : [];
        sections.push({ label: part.replace("：", ""), company, items });
      }
    }
    return (
      <div className="issues-cell">
        {sections.map((sec, idx) => (
          <div key={idx} className="issues-section">
            <span className="issues-label">{sec.label}：</span>
            {sec.company && <span className="issues-company">{sec.company}</span>}
            {sec.items.length > 0 && (
              <div className="issues-list">
                {sec.items.map((it, i) => (
                  <div key={i} className="issues-item">{it}</div>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
    );
  };

  // Build consolidated rows for current annex, respecting display order; optionally apply smart sort for Annex2
  const getSmartResponseOrder = (): number[] => {
    const baseOrder = (displayOrder.length ? displayOrder : responses.map((_, i) => i)).filter(
      (i) => responses[i]?.status === "completed"
    );
    if (!smartSortActive) return baseOrder;
    // Derive order from Annex2 first line: date(index 3) asc, then factory(index 2) asc
    const withKeys = baseOrder.map((idx) => {
      const resp = responses[idx];
      const text = resp?.annex2 || "";
      const firstLine = text
        .split("\n")
        .map((l) => l.trim())
        .filter(Boolean)[0] || "";
      const cols = firstLine ? firstLine.split("\t") : [];
      const date = cols[3] || "";
      const factory = cols[2] || "";
      return { idx, date, factory };
    });
    withKeys.sort((a, b) => {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      return a.factory.localeCompare(b.factory);
    });
    return withKeys.map((k) => k.idx);
  };

  const getConsolidatedRows = (annex: "annex1" | "annex2"): string[][] => {
    const order = getSmartResponseOrder();
    const ordered = order.map((origIndex) => responses[origIndex]).filter((r) => !!r);
    const allRows: string[][] = [];
    ordered.forEach((resp) => {
      const text = annex === "annex1" ? resp.annex1 : resp.annex2;
      if (!text) return;
      const rows = text
        .trim()
        .split("\n")
        .filter((line) => line.trim())
        .map((line) => line.split("\t"));
      rows.forEach((r) => allRows.push(r));
    });
    return allRows;
  };

  // Helper function to parse the schedule mapping (commented out for now)
  /* const parseScheduleMapping = (text: string): ScheduleMapping[] => {
    if (!text.trim()) return [];
    
    const lines = text.split('\n').filter(line => line.trim());
    return lines.map(line => {
      const parts = line.split('->').map(part => part.trim());
      if (parts.length >= 4) {
        const [plannedDate, regionTown, factoryName, tenant] = parts;
        const [region, town] = regionTown.includes('/') 
          ? regionTown.split('/').map(s => s.trim())
          : [regionTown, regionTown]; // fallback if no slash
        
        return {
          plannedDate,
          region,
          town,
          factoryName,
          tenant
        };
      }
      return null;
    }).filter(Boolean) as ScheduleMapping[];
  }; */

  // Helper function to get file last modified date
  const getFileActualDate = (file: File): string => {
    const date = new Date(file.lastModified);
    return date.toISOString().split('T')[0]; // Format as YYYY-MM-DD
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const newFiles = Array.from(e.target.files).slice(0, 100);
      setFiles(newFiles);

      setResponses(
        newFiles.map((file) => ({
          fileName: file.name,
          annex1: "",
          annex2: "",
          status: "idle",
        }))
      );

      // Initialize actual dates with file's last modified date (YYYY-MM-DD)
      setActualDates(newFiles.map((file) => getFileActualDate(file)));

      // Initialize ordering
      const initOrder = newFiles.map((_, i) => i);
      setOriginalOrder(initOrder);
      setDisplayOrder(initOrder);
      setSortMode("original");
    }
  };

  // Allow user to manually edit the actual check date for each file
  const handleDateChange = (index: number, value: string) => {
    setActualDates((prev) => {
      const next = [...prev];
      next[index] = value;
      return next;
    });
  };

  // Compute ordering by date
  const computeOrderByDate = (mode: "asc" | "desc"): number[] => {
    const indices = [...(displayOrder.length ? displayOrder : originalOrder)];
    const getDateFor = (i: number) => actualDates[i] || (files[i] ? getFileActualDate(files[i]) : "");
    indices.sort((a, b) => {
      const da = getDateFor(a);
      const db = getDateFor(b);
      if (da === db) return a - b;
      return mode === "asc" ? (da < db ? -1 : 1) : (da > db ? -1 : 1);
    });
    return indices;
  };

  // Toggle sort mode
  const handleToggleSort = () => {
    setSortMode((prev) => {
      const next = prev === "original" ? "asc" : prev === "asc" ? "desc" : "original";
      if (next === "original") {
        setDisplayOrder([...originalOrder]);
      } else if (next === "asc") {
        setDisplayOrder(computeOrderByDate("asc"));
      } else {
        setDisplayOrder(computeOrderByDate("desc"));
      }
      return next;
    });
  };

  // Keep display order in sync when dates change under sorted modes
  useEffect(() => {
    if (files.length === 0) return;
    if (sortMode === "asc") {
      setDisplayOrder(computeOrderByDate("asc"));
    } else if (sortMode === "desc") {
      setDisplayOrder(computeOrderByDate("desc"));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [actualDates]);

  // Drag & drop handlers
  const onDragStart = (origIndex: number) => {
    setDraggingIndex(origIndex);
  };

  const onDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  const onDropOverItem = (targetOrigIndex: number) => {
    if (draggingIndex === null || draggingIndex === targetOrigIndex) return;
    setSortMode("original");
    setDisplayOrder((prev) => {
      const fromPos = prev.indexOf(draggingIndex);
      const toPos = prev.indexOf(targetOrigIndex);
      if (fromPos === -1 || toPos === -1) return prev;
      const next = [...prev];
      const [moved] = next.splice(fromPos, 1);
      next.splice(toPos, 0, moved);
      return next;
    });
    setDraggingIndex(null);
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

    // Parse the schedule mapping (for future use)
    // const mappings = parseScheduleMapping(scheduleMappingText);

    try {
      await Promise.all(
        (displayOrder.length ? displayOrder : files.map((_, i) => i)).map(async (origIndex) => {
          try {
            const file = files[origIndex];
            // 从DOCX中提取文本
            const extractedText = await extractTextFromDocx(file);

            // Use user-edited actual check date if provided, otherwise fallback to file's last modified date
            const actualCheckDate = actualDates[origIndex] || getFileActualDate(file);

            // 准备prompt with the new mapping format
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
              - 实际核查时间: ${actualCheckDate}
              - 计划安排信息（格式：计划核查日期->区域/镇街道->厂中厂名称->承租方）:
              ${scheduleMappingText}
              
              请根据文档内容中找到的承租方或厂中厂名称，匹配上述计划安排信息来确定：区域、镇/街道、厂中厂名称、计划核查时间。
              如果在计划安排中找不到匹配的信息，请根据文档内容尽量推断这些信息。

              文档内容：
              ${extractedText.substring(0, 40000)}

              在返回时，请在第一部分和第二部分之间添加一行，内容为四个大写字母：XXXX

              返回例子1：
              张家港	经开区	张家港市杨舍镇农联村股份经济合作社	张家港市海达兴纺机有限公司	2025-06-23	${actualCheckDate}
              XXXX
              常州市安平安全技术服务有限公司	张家港	农联村村级租用厂房	${actualCheckDate}	出租方：张家港市杨舍镇农联村股份经济合作社 1、8楼安全出口指示灯不亮 承租方：苏州凡赛特材料科技有限公司1、9楼安全出口指示灯不亮 2、消火栓箱未见点检记录 3、消火栓箱未张贴操作说明 4、注塑机安全风险告知牌未划分风险等级和未明确管理责任人员	0	15	15	14	1	否

              返回例子2：
              张家港	经开区	张家港市杨舍镇徐丰村股份经济合作社	张家港市创新线业有限公司	2025-06-23	${actualCheckDate}
              XXXX
              常州市安平安全技术服务有限公司	张家港	徐丰村村级租用厂房	${actualCheckDate}	出租方：张家港市杨舍镇徐丰村股份经济合作社 1、出租方公告栏内各企业较大风险未更新 2、出租方公告栏内各企业安全风险四色图未更新 承租方;张家港市创新线业有限公司 1、货架未见限重标识 2、消火栓箱内放置灭火器 3、车间内通道堵塞 4、配电盒未张贴警示标识 5、电缆槽盒未跨接 6、绝缘胶垫未见检测合格标签 7、灭火器箱前堆放杂物 8、防腐剂放置点未见MSDS 9、较大风险公告栏未及时更新 10、清洁剂使用完放置在车间现场	0	12	12	10	2	否
              
              //所以你的回答只应该有像这样的三行，不要再有其他东西了
              //日期格式统一用2025-xx-xx
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
                i === origIndex
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
                i === origIndex
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
    const order = (displayOrder.length ? displayOrder : responses.map((_, i) => i));
    const completedResponses = order
      .map((i) => responses[i])
      .filter(resp => resp.status === "completed");

    if (completedResponses.length === 0) return;

    // Create 附件1 worksheet
    const annex1Sheet = workbook.addWorksheet("附件1");
    
    // Add headers for 附件1
    const annex1Headers = ["区域", "镇/街道", "出租方名称", "承租方名称", "计划核查时间", "实际核查时间"];
    annex1Sheet.addRow(annex1Headers);
    
    // Style header row for 附件1 (no fill, fixed height)
    const annex1HeaderRow = annex1Sheet.getRow(1);
    annex1HeaderRow.font = { name: "宋体", size: 9, bold: true };
    annex1HeaderRow.alignment = { horizontal: "center", vertical: "middle" };
    annex1HeaderRow.height = 39;
    
    // Add data rows for 附件1 using consolidated rows to respect current ordering (and smart sort)
    const annex1Rows = getConsolidatedRows("annex1");
    annex1Rows.forEach((row) => {
      const dataRow = annex1Sheet.addRow(row);
      dataRow.alignment = { vertical: "middle", wrapText: true };
      // Set font for all cells in row
      dataRow.eachCell((cell: any, colNumber: number) => {
        cell.font = { name: "宋体", size: 9 };
        // Center-align date columns (5 and 6, 1-based)
        if (colNumber === 5 || colNumber === 6) {
          cell.alignment = { ...cell.alignment, horizontal: "center" };
        }
      });
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

    // Add borders to all cells in 附件1 and set uniform row height (skip header for fill)
    annex1Sheet.eachRow((row: any, rowNumber: number) => {
      row.height = rowNumber === 1 ? 39 : 39;
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
    
    // Style header row for 附件2 (no fill, fixed height)
    const annex2HeaderRow = annex2Sheet.getRow(1);
    annex2HeaderRow.font = { name: "宋体", size: 9, bold: true };
    annex2HeaderRow.alignment = { horizontal: "center", vertical: "middle" };
    annex2HeaderRow.height = 39;
    
    // Add data rows for 附件2, possibly smart-sorted
    const annex2Rows = getConsolidatedRows("annex2");
    annex2Rows.forEach((row) => {
      // Convert numeric columns to numbers
      // Numeric columns: 重大隐患数量(5), 一般隐患数量(6), 隐患总数量(7), 现场隐患(8), 管理隐患(9)
      const processedRow = row.map((value, index) => {
        if ([5, 6, 7, 8, 9].includes(index)) {
          const numValue = parseInt(value);
          return isNaN(numValue) ? 0 : numValue;
        }
        return value;
      });
      const dataRow = annex2Sheet.addRow(processedRow);
      dataRow.alignment = { vertical: "middle", wrapText: true };
      // Set font for all cells in row
      dataRow.eachCell((cell: any, colNumber: number) => {
        cell.font = { name: "宋体", size: 9 };
        // Set number format for numeric columns
        if ([6, 7, 8, 9, 10].includes(colNumber)) { // 1-based indexing for columns
          (cell as any).numFmt = '0';
        }
        // Center-align date (4) and number columns (6-10, 1-based)
        if (colNumber === 4 || [6, 7, 8, 9, 10].includes(colNumber)) {
          cell.alignment = { ...cell.alignment, horizontal: "center" };
        }
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

    // Add borders to all cells in 附件2 and set uniform row height (skip header for fill)
    annex2Sheet.eachRow((row: any, rowNumber: number) => {
      row.height = rowNumber === 1 ? 39 : 393;
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

  // Copy buttons removed per request

  // Removed per consolidated table view
  
  return (
    <div className="app">
      <div className="container">
        <h1>安全周报自动化系统</h1>

        <form onSubmit={handleSubmit}>
          <div className="input-section">
            <h2>基本信息填写</h2>

            <div className="input-group">
              <label>计划安排信息 (格式: 计划核查日期-&gt;区域/镇街道-&gt;厂中厂名称-&gt;承租方):</label>
              <textarea
                value={scheduleMappingText}
                onChange={(e) => setScheduleMappingText(e.target.value)}
                placeholder="例如:&#10;2025/6/23-&gt;张家港/经开区-&gt;农联村村级租用厂房-&gt;张家港市海达兴纺机有限公司&#10;2025/6/24-&gt;张家港/经开区-&gt;徐丰村村级租用厂房-&gt;张家港市创新线业有限公司&#10;每行一个计划，实际核查日期将自动使用文件的最后修改时间"
                rows={5}
              />
            </div>
          </div>

          <div className="file-upload">
            <label htmlFor="file-upload">上传检查文档 (最多100个):</label>
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
              <div className="sort-controls">
                <button type="button" onClick={handleToggleSort} disabled={isLoading}>
                  排序: {sortMode === "original" ? "原始" : sortMode === "asc" ? "核查日期↑" : "核查日期↓"}
                </button>
              </div>
            )}
            {files.length > 0 && (
              <div className="file-list">
                <p>已选择文件 ({files.length}):</p>
                <div className="file-names-container">
                  {(displayOrder.length ? displayOrder : files.map((_, i) => i)).map((origIndex) => {
                    const file = files[origIndex];
                    return (
                      <div
                        key={origIndex}
                        className="file-item-simple"
                        draggable
                        onDragStart={() => onDragStart(origIndex)}
                        onDragOver={onDragOver}
                        onDrop={() => onDropOverItem(origIndex)}
                      >
                        <div className="drag-handle" title="拖拽排序" />
                        <div className="file-name">{file.name}</div>
                        <div className="file-date">
                          <label>
                            实际核查日期:
                            <input
                              type="date"
                              value={actualDates[origIndex] || getFileActualDate(file)}
                              onChange={(e) => handleDateChange(origIndex, e.target.value)}
                              disabled={isLoading}
                            />
                          </label>
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
          <div className="actions toolbar-actions">
            <button
              className="excel-download-btn"
              onClick={downloadExcelFile}
              disabled={!responses.some((r) => r.status === "completed")}
            >
              📊 下载Excel文件
            </button>
            <button
              type="button"
              className={`smart-sort-btn ${smartSortActive ? 'active' : ''}`}
              onClick={() => setSmartSortActive((s) => !s)}
            >
              智能排序（按日期内厂中厂分组）
            </button>
          </div>
        )}

        <div className="results-container">
          {/* Annex toggle */}
          <div className="annex-toggle" role="tablist" aria-label="Annex Switcher">
            <button
              type="button"
              className={`annex-toggle-btn ${activeAnnex === "annex1" ? "active" : ""}`}
              onClick={() => setActiveAnnex("annex1")}
              role="tab"
              aria-selected={activeAnnex === "annex1"}
            >
              附件1
            </button>
            <button
              type="button"
              className={`annex-toggle-btn ${activeAnnex === "annex2" ? "active" : ""}`}
              onClick={() => setActiveAnnex("annex2")}
              role="tab"
              aria-selected={activeAnnex === "annex2"}
            >
              附件2
            </button>
          </div>

          {/* Consolidated table */}
          <div className="table-container">
            <table>
              <thead>
                <tr>
                  {(activeAnnex === "annex1" ? ANNEX1_HEADERS : ANNEX2_HEADERS).map((header, i) => (
                    <th key={i}>{header}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {getConsolidatedRows(activeAnnex).map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {row.map((cell, cellIndex) => {
                      if (activeAnnex === "annex2" && cellIndex === 4) {
                        return (
                          <td key={cellIndex} className="issue-cell">
                            {renderIssuesContent(cell)}
                          </td>
                        );
                      }
                      return <td key={cellIndex}>{cell}</td>;
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
} 