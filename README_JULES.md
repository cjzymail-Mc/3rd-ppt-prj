# Jules PowerPoint Automation Toolkit

This repository contains Python scripts for generating pixel-perfect PowerPoint presentations using COM automation (`win32com.client`), as requested.

## Prerequisites

*   **Operating System**: Windows (Required for COM automation)
*   **Software**: Microsoft PowerPoint and Excel installed.
*   **Python**: Python 3.8+
*   **Dependencies**:
    ```bash
    pip install pywin32 xlwings
    ```

## Scripts Overview

### 1. Analysis (`analyze_templates_com.py`)

This script analyzes the difference between `0-空白 ppt 模板.pptx` (Blank Template) and `1-标准 ppt 模板.pptx` (Standard Template) to identify the "added" shapes that need to be updated.

*   **Usage**: `python analyze_templates_com.py`
*   **Output**: `shape_detail_com.json` (contains details of added shapes like Name, Position, Font).
*   **Note**: I have already performed this analysis (using a temporary method) and included the `shape_detail_com.json` file so you can skip this step if you wish.

### 2. Build Presentation (`build_codex_ppt_com.py`)

This is the main script that generates the final presentation `gemini-jules.pptx`.

*   **Logic**:
    1.  Reads data from `2-问卷模板.xlsx` using `xlwings`.
    2.  Calculates statistics (Score, Grade, Counts, etc.).
    3.  Opens `1-标准 ppt 模板.pptx` and saves it as `gemini-jules.pptx`.
    4.  Updates the content of the identified shapes (Text and Charts) while preserving all original formatting.
*   **Usage**: `python build_codex_ppt_com.py`
*   **Output**: `gemini-jules.pptx`

### 3. Verification (`verify_ppt_com.py`)

This script verifies the visual fidelity of the generated presentation against the standard template.

*   **Logic**:
    1.  Opens both `1-标准 ppt 模板.pptx` and `gemini-jules.pptx`.
    2.  Compares every shape on the first slide for position, size, and font properties.
    3.  Reports any discrepancies > 1 point.
*   **Usage**: `python verify_ppt_com.py`

## Class Updates (`src/Class_030.py`)

I have updated `src/Class_030.py` to include new classes `Text_Grade` and `Text_Score` based on the styles found in the standard template, as requested.

## Notes

*   All scripts are designed to run in the background (`Visible=False`) but can be set to True for debugging.
*   The scripts rely on the specific shape names found in `shape_detail_com.json` (e.g., "矩形 11"). If you rename shapes in the template, you must update the `SHAPE_MAPPING` in `build_codex_ppt_com.py`.
