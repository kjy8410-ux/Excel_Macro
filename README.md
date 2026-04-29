# Excel Automation Toolkit (VBA Macro)

## Overview

This project provides a set of Excel VBA macros designed to improve data processing efficiency
through automation of repetitive tasks such as string parsing, data formatting, and visualization.

The toolkit is implemented in both:

* `.xlsm` (Excel macro-enabled workbook)
* `.bas` (exported VBA module for reuse and version control)

---

## File Structure

```id="4k9b2x"
.
├── Excel Macro.xlsm
└── Excel Macro.bas
```

---

## Key Features

### 1. String Parsing Automation

Automated splitting of text data into columns based on various delimiters.

Supported delimiters:

* Space
* Comma (,)
* Underscore (_)
* Dash (-)
* Special characters (◇, ☆, §)

This allows quick processing of structured or semi-structured data without manual Excel operations.

---

### 2. Data Cleaning Utilities

* Reset text-to-column formatting (`T_Cancel`)
* Remove unnecessary or custom cell styles

Helps maintain clean and consistent Excel workbooks.

---

### 3. RGB Visualization Tool

* Reads RGB values from columns (A, B, C)
* Applies the corresponding color to the target cell

Use case:

* Color validation
* Visual inspection of RGB datasets

---

## Macro List

| Macro Name     | Description                      |
| -------------- | -------------------------------- |
| `T_Space()`    | Split text by space              |
| `T_Comma()`    | Split text by comma              |
| `T_Diamond()`  | Split text by ◇                  |
| `T_Star()`     | Split text by ☆                  |
| `T_SS()`       | Split text by §                  |
| `T_UdBar()`    | Split text by underscore (_)     |
| `T_Dash()`     | Split text by dash (-)           |
| `T_Cancel()`   | Reset text parsing configuration |
| `셀스타일삭제()`     | Remove custom cell styles        |
| `PreviewRGB()` | Display RGB color preview        |

---

## Usage

### Option 1: Using Excel Macro File (.xlsm)

1. Open `Excel Macro.xlsm`
2. Enable macros
3. Run macros from:

   * Developer tab → Macros
   * Assigned buttons (if configured)

---

### Option 2: Importing VBA Module (.bas)

1. Open Excel VBA Editor (`Alt + F11`)
2. Import `Excel Macro.bas`
3. Run macros directly from the VBA environment

---

## Use Cases

* Data preprocessing for engineering or production datasets
* Parsing test logs or structured strings
* Cleaning large Excel files
* Visualizing RGB or numerical data quickly

---

## Technical Highlights

* Built using Excel VBA
* Utilizes `TextToColumns` for efficient string parsing
* Modular macro structure for reuse
* Supports both interactive and batch-style operations

---

## Purpose

This toolkit was developed to:

* Reduce manual Excel operations
* Improve workflow efficiency
* Provide reusable automation tools for data handling

---

## Future Improvements

* Add GUI-based macro selection interface
* Extend delimiter options dynamically
* Improve error handling and input validation
* Add logging for macro execution

---
