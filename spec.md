# Minimal Spreadsheet Web Application MVP

## Overview
A comprehensive web-based spreadsheet application with advanced features including real-time collaboration, extensive formula support, data visualization, and enterprise-grade functionality similar to Excel or Google Sheets. Features an Excel-style ribbon interface for professional user experience with complete Excel-like interactivity and navigation.

## New Feature: Complete Project Source Code Merger
- Enhanced export functionality to merge the complete project into a single continuous file
- Combines all backend (Motoko canister code) and frontend (React + TypeScript) source files
- Logical file ordering: backend files first (`main.mo`, `authorization/*.mo`, `Storage.mo`, etc.), then frontend files (`App.tsx`, components, hooks, lib files, styles)
- Preserves readability with clear section dividers for context between backend and frontend sections
- Removes redundant import paths while maintaining code functionality
- Ensures the resulting file represents the entire functioning application's source code
- Accessible through File menu as "Export Complete Project Source"
- Generated file maintains proper code structure and formatting for both Motoko and React/TypeScript code
- Single continuous file format for complete project viewing and analysis

## Authentication & Security
- Internet Identity integration for user authentication
- Input sanitization to prevent XSS attacks
- Safe formula rendering and execution
- Rate limiting and access control with role-based permissions (owner, editor, viewer)

## Core Spreadsheet Features

### Grid Interface & Navigation
- Dynamic spreadsheet grid supporting unlimited rows and columns (A-Z, AA-ZZ, etc.)
- High-performance virtualized rendering for large datasets (up to 10,000+ rows × 50+ columns) with 60 FPS target
- Excel-style grid with bold column headers (A, B, C...) and numbered row headers (1, 2, 3...)
- Thin grid lines with blue selection border for active cells
- Advanced cell selection system with visual highlighting and border indication
- Multi-range selection using Shift and Ctrl modifiers
- Inline cell editing with type detection (number, text, date, boolean, currency, percentage)
- Column and row resizing functionality
- Freeze panes capability
- Hide and unhide rows/columns
- Go-to cell navigation feature
- Support for wrapped text and merged cell outlines
- Smooth scrolling with optimized rendering to prevent re-renders of non-visible cells
- Selection persistence during scrolling operations

### Excel-Style Interactivity System

#### Cell Selection & Navigation
- Single-click cell highlighting with blue border and Name Box update
- Mouse drag range selection (e.g., A1:C5) with visual feedback
- Shift + Arrow key selection expansion
- Ctrl/Cmd + Click for multi-range selection
- Selection remains visible during scrolling operations
- Enter or F2 opens inline editing mode
- Esc exits editing mode
- Drag-fill handle for auto-fill series functionality
- Hover highlight for active cell indication
- Instant feedback with immediate selection and formula bar updates

#### Comprehensive Keyboard Shortcuts
- Navigation shortcuts: Arrow keys, Ctrl + Arrows, Ctrl + Home
- Editing shortcuts: Enter, Shift + Enter, Tab, Shift + Tab, Esc, F2
- Selection shortcuts: Shift + Arrows, Ctrl + A, Ctrl + Space, Shift + Space
- Clipboard shortcuts: Ctrl + C/X/V, Delete
- Undo/Redo shortcuts: Ctrl + Z/Y
- Save shortcut: Ctrl + S
- Formatting shortcuts: Ctrl + B/I/U, Ctrl + 1
- Insert/Delete shortcuts: Ctrl + + / Ctrl + -
- Find/Replace shortcuts: Ctrl + F/H
- Sheet navigation shortcuts: Ctrl + PgUp/PgDn
- Cross-platform compatibility for Windows, macOS, and Linux

#### Excel-Style KeyTips System
- Alt key activation (or Ctrl+Option on macOS) to enter KeyTip mode
- Small keytip badges displayed above ribbon tabs and toolbar buttons
- Full key sequence support (e.g., Alt → H → B for Bold formatting)
- Small overlay showing current key sequence (e.g., [Alt] → [H] → [B])
- Esc key or 2.5-second timeout to cancel KeyTip mode
- Key state machine managing states: idle → awaitingTab → awaitingCommand → executed/cancelled
- Only captures Alt sequences when spreadsheet or ribbon has focus
- Does not override browser shortcuts like Alt+F4

##### Default KeyTip Mappings
- Ribbon Tabs: F (File), H (Home), N (Insert), P (Page Layout), M (Formulas), A (Data), R (Review), W (View)
- Home Tab: B (Bold), I (Italic), U (Underline), F (Font), S (Size), L (Align Left), C (Center), R (Align Right), M (Merge), Z (Undo), Y (Redo), O (Font Color), G (Fill Color)
- Insert Tab: T (Table), C (Chart), P (Picture), I (Image)
- Data Tab: F (Filter), S (Sort), T (Text to Columns)

##### KeyTip Settings & Configuration
- Settings toggle to enable/disable KeyTips functionality
- macOS modifier configuration options (Ctrl+Option / Alt / Cmd)
- Extensible API: registerKeyTip(tab, key, label, handler) for custom buttons

#### Touch Gesture Support
- Pinch-to-zoom functionality for mobile and tablet devices
- Double-tap to edit cell content
- Touch-optimized selection and navigation

### Excel-Style Interface Components

#### Ribbon Interface
- Top ribbon with tabs: File, Home, Insert, Page Layout, Formulas, Data, Review, View, Help
- Each tab displays respective toolbar groups with icons and text labels
- Home tab includes: Clipboard, Font, Alignment, Number, Styles, Cells, Editing groups
- Insert tab includes: Tables, Illustrations (with Image button), Charts, Links groups
- Responsive design that collapses less-used groups into "More" menu on narrow screens
- ARIA roles and accessible labels for screen readers
- Keyboard navigation support within ribbon
- KeyTip badge rendering with absolute positioning and .keytip-badge CSS class

#### File Menu
- File menu dropdown with options for spreadsheet management
- "Upload Excel File (.xlsx)" option that opens native file picker restricted to .xlsx files
- "Export Complete Project Source" option to merge all backend and frontend code into single continuous file with logical ordering and clear section dividers
- Excel file import functionality with client-side parsing using xlsx library via dynamic import
- Import process replaces current sheet data and clears undo/redo history
- Success and error feedback messages for import operations

#### Quick Access Toolbar
- Positioned to the left of the ribbon
- Customizable icons for Save, Undo, Redo, Print, and Share
- Small icon buttons for quick access to frequently used functions

#### Name Box and Formula Bar
- Name Box showing current cell reference (e.g., "A1", "B5:D10") with instant updates
- Formula Bar with placeholder text "Enter value or formula (e.g. =SUM(A1:A10))"
- Save/Confirm button next to Formula Bar
- Formula bar expands for long formulas
- Immediate updates reflecting current selection

#### Toolbar Controls
- Font family dropdown selector
- Font size selector with common sizes
- Bold, italic, underline formatting buttons
- Font Color button with dropdown palette containing exactly 5 fixed colors: White (#FFFFFF), Black (#000000), Red (#FF0000), Dark Blue (#00008B), Green (#008000)
- Fill Color button with dropdown palette containing exactly the same 5 fixed colors: White (#FFFFFF), Black (#000000), Red (#FF0000), Dark Blue (#00008B), Green (#008000)
- Text alignment options (left, center, right, justify)
- Merge cells functionality
- Number format dropdown (General, Number, Currency, Percentage, Date, etc.)
- Border styling options
- Conditional formatting tools
- Sort and filter controls
- Chart creation tools
- Image insertion button in Insert tab

#### Context Menu
- Right-click context menu for cells with options:
  - Cut, Copy, Paste
  - Insert/Delete rows and columns
  - Format Cells dialog
  - Insert Comment
  - Clear Contents
  - Cell Protection settings

#### Sheet Management
- Sheet tabs at bottom of interface with navigation controls
- Clickable sheet tabs to switch between sheets within the same spreadsheet file
- Active sheet tab highlighted with distinct visual styling (selected appearance)
- "+ New Sheet" button positioned next to existing sheet tabs
- Create new sheets with automatic naming sequence (Sheet2, Sheet3, Sheet4, etc.)
- Automatic switching to newly created sheet
- Sheet tab scrolling for multiple sheets
- Each sheet maintains independent data (cells, formatting, images)
- New sheets initialize with blank grid structure and reset undo/redo stacks
- Active sheet state management with activeSheetId tracking
- Smooth transitions between sheets without UI lag or state corruption
- Sheet switching preserves all data per sheet and restores content instantly

#### Status Bar
- Bottom status bar with:
  - Sheet navigation controls
  - Zoom controls with percentage display
  - Workbook statistics (cell count, sum of selected cells)

#### Collapsible Sidebar
- "My Spreadsheets" panel converted to collapsible sidebar
- Icon toggle button to show/hide sidebar
- File management and organization features

### Excel File Import System

#### Import Interface
- "Upload Excel File (.xlsx)" option in File menu dropdown
- Native file picker restricted to .xlsx file format
- Client-side parsing using xlsx library loaded via dynamic import

#### Parsing & Conversion
- Extract cell values from Excel workbook
- Parse basic formatting including bold, italic, and horizontal alignment (left, center, right)
- Handle merged cell ranges and convert to internal MergeInfo model
- Convert parsed data to internal SheetState model with cells Record<string, Cell>
- Each Cell contains value, style properties (hAlign, vAlign, bold, italic), and merge information

#### Import Process
- Replace current active sheet data with imported Excel content
- Clear undo/redo history after successful import
- Maintain full compatibility with existing grid, formatting, and editing features
- Save imported data to backend using existing persistence APIs

#### Error Handling & Feedback
- Display user-friendly error messages for invalid or corrupted Excel files
- Show success toast notification "Excel file imported successfully" on completion
- Graceful handling of unsupported Excel features

### Image Management System

#### Image Insertion
- Insert → Image button in ribbon under Insert tab
- Native file picker supporting PNG and JPG file formats
- Images displayed as absolutely positioned elements above the spreadsheet grid
- Initial placement anchored to currently selected cell's top-left corner
- Position calculated based on grid cell's DOM coordinates

#### Image Interactivity
- Drag functionality to reposition images anywhere over the grid
- Resize functionality via visible corner handles
- Smooth mouse interactions without jitter or lost events during dragging/resizing
- Images maintain position relative to grid during scrolling

#### Image Data Structure
- Each image stored with properties: id, src (base64), x, y, width, height, anchorCell
- Images collection added to sheet state as Record<string, ImageData>
- Image state persisted with spreadsheet data
- Loaded spreadsheets restore all images with their positions and sizes

#### Image Undo/Redo Integration
- Image insertion, moving, and resizing integrated with existing undo/redo system
- Undo removes or reverts image modifications
- Redo reinstates image changes
- History actions recorded for all image operations

### Cell Formatting System

#### Centralized Formatting Controller
- Centralized `applyFormatToSelection(formatObj)` function in spreadsheet controller/store
- Merges format objects (e.g., `{ bold: true }`, `{ fontColor: '#FF0000' }`, `{ fillColor: '#008000' }`) into each selected cell's `meta.format`
- Batches formatting updates for performance optimization
- Maintains formatting state consistency across all selected cells

#### Cell Format Storage
- Each cell stores formatting data in `cell.meta.format` object
- Supports formatting properties: bold, italic, underline, font family, font size, fontColor, fillColor, text color, background color, alignment, borders
- Format data persisted to backend and synchronized with collaborators in real-time

#### Visual Format Rendering
- Cell renderer reads `cell.meta.format` and applies corresponding inline styles or CSS classes
- Dynamic styling for bold, italic, underline, font properties, colors, alignment, and borders
- Font color applied as `style.color = format.fontColor ?? "inherit"`
- Fill color applied as `style.backgroundColor = format.fillColor ?? "transparent"`
- Performance-optimized rendering that minimizes re-renders during format changes

#### Toolbar Format State Reflection
- Selection state updates to reflect current formatting of selected cells
- Toolbar buttons show active, inactive, or indeterminate states when formatting is mixed across selection
- Real-time toolbar state updates as selection changes

#### Format Command Integration
- Ribbon and KeyTip command handlers (Alt→H→B, Alt→H→I, Alt→H→O, Alt→H→G) wired to call `applyFormatToSelection`
- Toolbar buttons toggle correctly based on current selection formatting state
- Context menu formatting options integrated with centralized formatting system

#### Fixed Color Palette System - CRITICAL IMPLEMENTATION
- Font Color button opens dropdown palette displaying exactly 5 fixed colors: White (#FFFFFF), Black (#000000), Red (#FF0000), Dark Blue (#00008B), Green (#008000)
- Fill Color button opens dropdown palette displaying exactly the same 5 fixed colors: White (#FFFFFF), Black (#000000), Red (#FF0000), Dark Blue (#00008B), Green (#008000)
- Color selection must update the selected cells' `format.fontColor` and `format.fillColor` properties in the cell data model
- Grid renderer must apply these colors using inline styles: `style={{ color: cell.format?.fontColor ?? "inherit", backgroundColor: cell.format?.fillColor ?? "transparent" }}`
- Color changes must trigger immediate grid re-rendering to show visual changes in the spreadsheet
- Color formatting must be saved to backend and persist across reloads and sheet switches
- Font color and fill color changes must be integrated with undo/redo system as atomic operations
- Color selection applies chosen color to all selected cells as single atomic undo/redo action
- Color palettes appear as dropdown or popover near the respective buttons
- No custom color picker or extended color palette support
- Efficient re-rendering for single or multiple selected cells
- Performance-optimized color application that does not degrade grid performance
- Instant visual feedback when color is applied to any selection size

### Data Operations
- Copy, cut, and paste operations (keyboard shortcuts and context menu)
- Drag-fill functionality for auto-completing series
- Undo and redo operations with history tracking and stack persistence across edits
- Undo/redo entries for all formatting actions (Bold, Italic, Underline, Align, Merge, Font Color, Fill Color, etc.)
- Undo/redo actions scoped to currently active sheet only
- Sheet switching isolates undo/redo history to prevent mixing stacks between sheets
- Find and replace functionality across sheets
- Import and export support for CSV, XLSX, and JSON formats using SheetJS library
- Excel file import with client-side parsing and data conversion

### Formula System
- Comprehensive formula bar with parser and evaluator
- Mathematical functions: SUM, AVERAGE, MIN, MAX, ROUND, ROUNDUP, ROUNDDOWN
- Lookup functions: VLOOKUP, HLOOKUP, INDEX, MATCH
- Logical functions: IF, AND, OR, NOT, IFERROR
- Text functions: CONCAT, CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER
- Date/Time functions: TODAY, NOW, DATE, YEAR, MONTH, DAY, EDATE, DATEDIF
- Statistical functions: COUNT, COUNTA, COUNTIF, SUMIF, MEDIAN, STDEV
- Array literals and range support
- Dependency graph for automatic recalculation when referenced cells change
- Error handling with standard error types: #REF!, #DIV/0!, #NAME?, #VALUE!, #N/A

### Data Visualization
- Chart creation functionality supporting bar, line, and pie charts
- Charts automatically update when source data changes
- Chart customization and formatting options accessible through ribbon

## File & Sheet Management

### Spreadsheet Files
- Create new spreadsheet files
- Save and load existing spreadsheet files
- File management in collapsible sidebar for organizing spreadsheets
- Each spreadsheet stored with unique identifier and metadata
- Excel file import functionality for loading external .xlsx files
- Complete project source code merger functionality for comprehensive code analysis with logical file ordering and clear section dividers

### Sheet Management
- Multiple sheets per spreadsheet file with tab navigation
- Create new sheets using "+ New Sheet" button positioned next to existing sheet tabs
- Automatic naming sequence for new sheets (Sheet2, Sheet3, Sheet4, etc.) based on existing sheets
- Automatic switching to newly created sheet
- Clickable sheet tabs to switch between sheets within the same spreadsheet file
- Active sheet tracking with activeSheetId state management
- Each sheet maintains independent data including cell values, formatting, and images
- New sheets initialize with blank grid structure and reset undo/redo stacks
- Sheet tabs displayed at bottom of interface with scroll controls
- Default naming convention (Sheet1, Sheet2, etc.)
- Active sheet tab visually highlighted with selected styling
- Smooth sheet switching with instant data restoration and no UI lag
- Data preservation per sheet when switching between tabs

## Collaboration Features

### Real-Time Collaboration
- Multiple users can edit simultaneously
- Real-time updates using WebSocket or ICP event model
- Multiple cursor display showing other users' positions
- Cell locking to prevent conflicts during editing
- Real-time formatting change synchronization across all collaborators

### Sharing & Permissions
- Generate sharing links for spreadsheets
- Permission levels: view, comment, and edit access
- Role-based access control managed in backend

### Version Control
- Version history tracking for all changes
- Rollback capability to previous versions
- Change attribution to users

## Data Persistence & Backend

### Storage
- Persistent storage for all spreadsheet data, sheets, and cells
- Multiple sheets stored within a single spreadsheet file with independent data
- Cell formatting data storage in `meta.format` objects including fontColor and fillColor properties must be properly persisted and restored
- Image data storage with base64 encoding and positioning information
- Version history storage for rollback functionality
- User authentication and permission data storage
- Undo/redo stack persistence for seamless user experience per sheet
- Active sheet state persistence and restoration
- Sheet-specific data isolation and management

### API Operations
- CRUD operations for spreadsheets, sheets, and cells
- Sheet creation operations with automatic naming and initialization
- Image data persistence and retrieval operations
- Formatting change persistence and real-time broadcasting including font and fill color data must work correctly
- Dependency graph management for formula recalculation
- Real-time event broadcasting for collaboration
- Access control enforcement for all operations
- Performance-optimized operations for large dataset handling
- Active sheet management and switching operations
- Multi-sheet data persistence and restoration
- Complete project source code merger operations with proper backend and frontend code organization and section markers

## User Interface Design

### Layout & Responsiveness
- Excel-style professional interface with ribbon, toolbar, and grid layout
- Responsive design that works on desktop and tablet devices
- Collapsible sidebar for file management
- Formula bar and Name Box positioned above the grid
- Sheet tabs and status bar at bottom of interface
- Ribbon toolbar with comprehensive formatting and function options
- Touch-optimized interface for mobile and tablet devices

### Accessibility
- ARIA roles and labels for all ribbon components, grid, cells, and formula bar
- Keyboard navigation support throughout the interface
- Screen reader compatibility for toolbar buttons and formula bar
- High contrast support for grid and interface elements
- Comprehensive accessibility compliance for all interactive elements
- KeyTip accessibility features: ARIA labels for badges, focusable badges, high contrast support, aria-live announcements for sequence updates

### Performance Optimization
- High-performance virtualized rendering for handling large datasets with 60 FPS target
- Efficient formula recalculation using dependency graphs
- Optimized real-time collaboration updates
- Smooth scrolling implementation with optimized rendering
- Prevention of unnecessary re-renders for non-visible cells
- Maintained performance with Excel-style interface enhancements
- Batched formatting updates to minimize re-renders during format operations

## Help & Documentation

### Keyboard Shortcuts Panel
- Comprehensive Help → Keyboard Shortcuts panel
- KeyTip mappings documentation with default key sequences
- Printable cheat sheet for keyboard shortcuts and KeyTips
- Accessibility notes and configuration options
- Cross-platform shortcut variations (Windows, macOS, Linux)

## Testing & Quality Assurance

### Automated Testing
- Unit tests for formula parser and evaluator logic
- Unit tests for backend CRUD operations and business logic
- Unit tests for KeyTip state machine logic
- Unit tests for `applyFormatToSelection` function verifying correct format merging and undo behavior
- Unit tests for fixed 5-color palette font color and fill color formatting functionality must verify actual cell data updates and grid rendering
- Unit tests for image insertion, positioning, and persistence functionality
- Unit tests for Excel file import parsing and data conversion
- Unit tests for sheet creation, naming, and switching functionality
- Unit tests for active sheet state management and data isolation
- Unit tests for complete project source code merger functionality with proper file ordering and section organization
- End-to-end tests for editing workflows, import/export functionality, and collaboration features
- End-to-end tests for KeyTip sequences (Alt → H → B, Alt → N → C, Alt → N → I, Alt → H → O, Alt → H → G)
- End-to-end test: select range → Alt+H→B → verify cells become bold visually and in model state
- End-to-end tests for fixed color palette formatting: select cells → click Font Color → select from 5 colors → verify text color changes in grid and persists
- End-to-end tests for fixed fill color palette formatting: select cells → click Fill Color → select from 5 colors → verify background color changes in grid and persists
- End-to-end tests for image insertion workflow: Insert → Image → file selection → placement → drag/resize
- End-to-end tests for Excel import workflow: File → Upload Excel → file selection → parsing → data loading
- End-to-end tests for sheet management: create new sheet → verify naming → switch between sheets → verify data independence
- End-to-end tests for sheet switching: click tab → verify active sheet change → verify data preservation → verify undo/redo isolation
- End-to-end tests for complete project source code merger: File → Export Complete Project Source → verify comprehensive code generation with logical ordering and section markers
- UI tests for ribbon functionality and Excel-style interface components
- Performance tests for large dataset handling and smooth interactivity
- Cross-platform compatibility testing for keyboard shortcuts and gestures

### Documentation
- Developer README with architecture overview, setup instructions, testing procedures, and deployment guidelines
- User guide documenting supported functions, formula examples, keyboard shortcuts, KeyTip usage, ribbon functionality, image insertion features, Excel import functionality, sheet management features, fixed 5-color palette formatting features, and complete project source code merger functionality
- Deployment scripts for Internet Computer canisters
- Live demo URL provided after deployment

## Technical Implementation Notes
- Frontend built with React, TypeScript, and Tailwind CSS
- Icon set integration (Lucide or Font Awesome) for consistent ribbon styling
- Backend implemented as Motoko canister on Internet Computer
- Real-time features using ICP's query/update model or WebSocket integration
- Formula evaluation runs in frontend for performance
- All user data and spreadsheet content stored persistently in backend
- Image data stored as base64 strings with positioning metadata
- Excel file parsing using xlsx library loaded via dynamic import for client-side processing
- Complete project source code merger generates single continuous file containing all backend and frontend modules with logical ordering and clear section markers
- High-performance virtualization and smooth scrolling implementation
- Comprehensive keyboard shortcut system with cross-platform support
- KeyTip system with state machine and extensible API
- Touch gesture integration for mobile and tablet devices
- Application content and interface in English language
- Feature branch delivery with staging demo and video demonstration
