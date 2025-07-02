# 安全周报自动化系统 (Safety Report Automation System)

A modern React + TypeScript application for automating safety inspection report generation, built with Vite.

## 🚀 Features

- 📄 **Document Processing** - Upload and process DOCX safety inspection documents
- 🤖 **AI Integration** - Automatically extract and format data using AI (DeepSeek)
- 📊 **Excel Export** - Generate professionally formatted Excel files with multiple sheets
- 📋 **Copy to Clipboard** - One-click copying of formatted data for annexes
- 🎨 **Professional Formatting** - Auto-styled Excel sheets with borders, colors, and proper column sizing
- ⚡ **Vite** - Lightning fast build tool and dev server
- ⚛️ **React 18** - Latest React with concurrent features
- 🔷 **TypeScript** - Full type safety and IntelliSense
- 🎨 **Modern CSS** - Beautiful gradient design with glassmorphism effects
- 📱 **Responsive** - Mobile-first responsive design
- 🔍 **ESLint** - Code linting and formatting
- 🔥 **Hot Module Replacement** - Instant updates during development

## 🛠️ Getting Started

### Prerequisites

Make sure you have Node.js installed (version 16 or higher):
- [Download Node.js](https://nodejs.org/)

### Installation

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Start the development server:**
   ```bash
   npm run dev
   ```

3. **Open your browser:**
   Navigate to `http://localhost:5173` to see your app running!

## 📖 How to Use

1. **Fill in Basic Information:**
   - 区域 (Region): Default is "张家港"
   - 镇/街道 (Town/Street): Default is "经开区"
   - 核查时间 (Inspection Date): Select the inspection date
   - 厂中厂映射关系 (Factory Mapping): Optional mapping information

2. **Upload Documents:**
   - Upload up to 10 DOCX safety inspection documents
   - The system supports .docx and .doc file formats

3. **Generate Reports:**
   - Click "生成安全检查报告" to process all documents
   - The system will extract text and generate formatted reports using AI

4. **Download or Copy Results:**
   - **📊 Download Excel**: Get a professionally formatted Excel file with separate sheets for 附件1 and 附件2
   - **📋 Copy Data**: Copy all "附件1" (Annex 1) or "附件2" (Annex 2) data with one click
   - **📱 View Tables**: View processed data in organized, responsive tables for each document

## 📜 Available Scripts

- `npm run dev` - Start development server
- `npm run build` - Build for production
- `npm run preview` - Preview production build locally
- `npm run lint` - Run ESLint

## 📁 Project Structure

```
safety-report-automation/
├── public/
│   └── vite.svg           # Favicon
├── src/
│   ├── App.tsx           # Main App component with business logic
│   ├── styles.css        # Complete application styles
│   ├── main.tsx         # React entry point
│   └── index.css        # Global base styles
├── index.html           # HTML template
├── package.json         # Dependencies and scripts
├── tsconfig.json        # TypeScript configuration
├── vite.config.ts       # Vite configuration
├── .eslintrc.cjs       # ESLint configuration
├── .gitignore          # Git ignore patterns
└── README.md           # Documentation
```

## 🎨 Design Features

The application includes a beautiful modern design with:
- Gradient backgrounds and glassmorphism effects
- Smooth animations and hover effects
- Responsive tables and mobile-optimized layout
- Status badges with animated processing indicators
- Professional form styling with focus effects
- Prominent Excel download button with distinctive orange styling

## 🔧 Built With

- [React](https://reactjs.org/) - UI library
- [TypeScript](https://www.typescriptlang.org/) - Type safety
- [Vite](https://vitejs.dev/) - Build tool
- [Mammoth.js](https://github.com/mwilliamson/mammoth.js/) - DOCX text extraction
- [ExcelJS](https://github.com/exceljs/exceljs) - Excel file generation and formatting
- [DeepSeek API](https://platform.deepseek.com/) - AI-powered data processing
- [ESLint](https://eslint.org/) - Code linting

## 📋 Data Processing

The system processes safety inspection documents and generates two types of annexes:

**附件1 (Annex 1) - Basic Information:**
- 区域 (Region)
- 镇/街道 (Town/Street)
- 出租方名称 (Lessor Name)
- 承租方名称 (Lessee Name)
- 计划核查时间 (Planned Inspection Time)
- 实际核查时间 (Actual Inspection Time)

**附件2 (Annex 2) - Safety Details:**
- 核查机构名称 (Inspection Organization)
- 地区 (Area)
- 厂中厂名称 (Factory-in-Factory Name)
- 核查时间 (Inspection Time)
- 存在问题 (Existing Issues)
- 重大隐患数量 (Major Hazards Count)
- 一般隐患数量 (General Hazards Count)
- 隐患总数量 (Total Hazards Count)
- 现场隐患 (On-site Hazards)
- 管理隐患 (Management Hazards)
- 是否属于涉爆粉尘、金属熔融企业 (Explosive Dust/Metal Melting Enterprise)

## 📊 Excel Features

The generated Excel files include:

**📋 Professional Formatting:**
- **Two separate worksheets**: 附件1 and 附件2 for organized data
- **Colored headers**: Blue for 附件1, Green for 附件2
- **Auto-sized columns**: Optimized width for readability
- **Cell borders**: Clean, professional table appearance
- **Alternating row colors**: Enhanced readability
- **Text wrapping**: Long content properly displayed

**📐 Column Specifications:**
- **附件1**: Standard 20-character width for all columns
- **附件2**: 
  - Standard columns: 15-20 characters
  - "存在问题" column: 50 characters (extra wide for detailed issues)
  - Numeric columns: 15 characters

**💾 File Naming:**
- Format: `安全检查报告_YYYY-MM-DD.xlsx`
- Automatic date stamping
- Chinese filename support

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

---

Happy coding! 🎉 