# Web Spreadsheet App (Excel-Style)

A minimal Excel-style spreadsheet web application with:

- Editable grid (cells A1, B2…)
- Formatting (bold, italic, alignment)
- 5-color system (White, Black, Red, Dark Blue, Green)
- Insert images (draggable + resizable)
- Undo / Redo system
- Sheet tabs (create, switch)
- Excel file import (.xlsx)
- KeyTips (Excel ALT shortcuts)
- Formula engine (SUM, AVERAGE, IF, IFERROR, MEDIAN, etc.)
- Clean React + TypeScript + Tailwind codebase

## ✨ Features

### Grid
- Click to select cells  
- Type to edit  
- Drag to select range  
- Supports hundreds of rows & columns  

### Formatting
- Bold / Italic  
- Alignment (left / center / right)  
- Text color & Fill color (5 fixed colors)  
- Works instantly and saved in state  

### Sheets
- Add new sheet  
- Switch between sheets  
- Fully independent sheet data  

### Images
- Insert PNG/JPG  
- Drag to move  
- Resize with handles  
- Saved in sheet layer  

### Undo / Redo
- Full stack-based undo/redo  
- Works for:  
  - cell values  
  - formatting  
  - images  
  - sheet edits  

### File Import
- Upload `.xlsx` file  
- Parses values & basic formatting  
- Replaces current sheet  

### KeyTips (Excel ALT shortcuts)
- ALT → H → B for Bold  
- ALT → H → I for Italic  
- ALT → H → O for Font Color  
- ALT → H → G for Fill Color  

## 🛠️ Tech Stack

- React  
- TypeScript  
- TailwindCSS  
- Vite  
- SheetJS (xlsx import)  

## 📦 Running Locally

```bash
npm install
npm run dev
