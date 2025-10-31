# DOCX Formatting MCP Server - Usage Examples

This document provides examples of how to use each tool in the DOCX Formatting MCP Server.

## Prerequisites

Before using these tools, you need:
1. An Azure Blob Storage account
2. A container in your storage account
3. Either:
   - Blob URIs with SAS (Shared Access Signature) tokens, OR
   - `AZURE_STORAGE_CONNECTION_STRING` environment variable set

## Example Blob URI Format

```
https://mystorageaccount.blob.core.windows.net/mycontainer/mydocument.docx?sp=racwdl&st=2024-01-01T00:00:00Z&se=2024-12-31T23:59:59Z&sv=2021-06-08&sr=b&sig=SIGNATURE_HERE
```

## Tool Examples

### 1. create_docx - Create a New Document

Creates a new DOCX file in Azure Blob Storage.

**Example 1: Basic Document**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "title": "Monthly Report",
  "initial_content": "This is the monthly report for October 2024."
}
```

**Example 2: Minimal Document**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/blank.docx?sas_token"
}
```

---

### 2. read_docx - Read Document Content

Reads and extracts all content, structure, and formatting from a DOCX file.

**Example:**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token"
}
```

**Response Structure:**
```json
{
  "success": true,
  "blob_uri": "...",
  "content": {
    "paragraphs": [
      {
        "text": "Monthly Report",
        "style": "Heading 1",
        "alignment": "CENTER",
        "runs": [
          {
            "text": "Monthly Report",
            "bold": true,
            "italic": false,
            "underline": false,
            "font_size": 16,
            "font_name": "Calibri"
          }
        ]
      }
    ],
    "tables": [...],
    "sections": [
      {
        "page_width": 8.5,
        "page_height": 11.0,
        "left_margin": 1.0,
        "right_margin": 1.0,
        "top_margin": 1.0,
        "bottom_margin": 1.0
      }
    ]
  }
}
```

---

### 3. insert_text - Add Text to Document

Inserts text at specified positions in the document.

**Example 1: Append Text**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "text": "This is a new paragraph at the end.",
  "position": "end"
}
```

**Example 2: Insert at Beginning**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "text": "CONFIDENTIAL - Internal Use Only",
  "position": "start",
  "style": "Heading 2"
}
```

**Example 3: Insert at Specific Position**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "text": "Inserted paragraph",
  "position": "at_index",
  "paragraph_index": 2,
  "style": "Normal"
}
```

---

### 4. edit_text - Modify Existing Text

Edits text in a specific paragraph while optionally preserving formatting.

**Example 1: Edit with Preserved Formatting**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "paragraph_index": 0,
  "new_text": "Updated Monthly Report - November 2024",
  "preserve_formatting": true
}
```

**Example 2: Edit without Preserving Formatting**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "paragraph_index": 3,
  "new_text": "This paragraph has been completely replaced.",
  "preserve_formatting": false
}
```

---

### 5. insert_or_edit_table - Manage Tables

Inserts new tables or edits existing ones.

**Example 1: Insert New Table**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "table_data": [
    ["Name", "Department", "Score"],
    ["Alice", "Engineering", "95"],
    ["Bob", "Marketing", "88"],
    ["Charlie", "Sales", "92"]
  ],
  "position": "end",
  "style": "Table Grid"
}
```

**Example 2: Edit Existing Table**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "table_index": 0,
  "table_data": [
    ["Name", "Department", "Score", "Status"],
    ["Alice", "Engineering", "95", "Approved"],
    ["Bob", "Marketing", "88", "Pending"]
  ]
}
```

**Example 3: Financial Table**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/budget.docx?sas_token",
  "table_data": [
    ["Category", "Q1", "Q2", "Q3", "Q4"],
    ["Revenue", "$100K", "$120K", "$135K", "$150K"],
    ["Expenses", "$80K", "$85K", "$90K", "$95K"],
    ["Profit", "$20K", "$35K", "$45K", "$55K"]
  ],
  "style": "Light Grid Accent 1"
}
```

---

### 6. insert_or_edit_image - Manage Images

Inserts images from base64 encoded data or Azure Blob URIs.

**Example 1: Insert Image from Azure Blob**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "image_data": "https://mystorageaccount.blob.core.windows.net/images/chart.png?sas_token",
  "width_inches": 5.0,
  "position": "end"
}
```

**Example 2: Insert Image from Base64**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "image_data": "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDA...",
  "width_inches": 3.0,
  "position": "start"
}
```

**Example 3: Insert Logo (Small)**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/letterhead.docx?sas_token",
  "image_data": "https://mystorageaccount.blob.core.windows.net/images/logo.png?sas_token",
  "width_inches": 2.0,
  "position": "start"
}
```

---

### 7. format_document - Apply Formatting

Applies comprehensive formatting to text and page margins.

**Example 1: Format Specific Paragraph**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "paragraph_index": 0,
  "bold": true,
  "font_name": "Arial",
  "font_size": 24,
  "alignment": "center",
  "font_color_rgb": [0, 0, 255]
}
```

**Example 2: Format All Paragraphs**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "font_name": "Times New Roman",
  "font_size": 12,
  "alignment": "justify"
}
```

**Example 3: Set Page Margins**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "left_margin_inches": 1.5,
  "right_margin_inches": 1.5,
  "top_margin_inches": 1.0,
  "bottom_margin_inches": 1.0
}
```

**Example 4: Comprehensive Formatting**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "paragraph_index": 2,
  "bold": false,
  "italic": true,
  "underline": false,
  "font_name": "Georgia",
  "font_size": 11,
  "font_color_rgb": [64, 64, 64],
  "alignment": "left"
}
```

**Example 5: Highlight Text with Color**
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/docs/report.docx?sas_token",
  "paragraph_index": 5,
  "bold": true,
  "font_color_rgb": [255, 0, 0],
  "font_size": 14
}
```

---

## Complete Workflow Example

Here's a complete example workflow creating a formatted report:

### Step 1: Create Document
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/reports/quarterly.docx?sas_token",
  "title": "Q4 2024 Quarterly Report"
}
```

### Step 2: Add Introduction
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/reports/quarterly.docx?sas_token",
  "text": "This report summarizes our performance in Q4 2024.",
  "position": "end"
}
```

### Step 3: Insert Performance Table
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/reports/quarterly.docx?sas_token",
  "table_data": [
    ["Metric", "Target", "Actual", "Variance"],
    ["Revenue", "$1M", "$1.2M", "+20%"],
    ["Users", "10K", "12K", "+20%"],
    ["Satisfaction", "85%", "92%", "+7%"]
  ],
  "style": "Table Grid"
}
```

### Step 4: Add Chart Image
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/reports/quarterly.docx?sas_token",
  "image_data": "https://mystorageaccount.blob.core.windows.net/charts/q4-performance.png?sas_token",
  "width_inches": 6.0
}
```

### Step 5: Format Title
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/reports/quarterly.docx?sas_token",
  "paragraph_index": 0,
  "bold": true,
  "font_size": 18,
  "alignment": "center",
  "font_color_rgb": [0, 51, 102]
}
```

### Step 6: Set Page Margins
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/reports/quarterly.docx?sas_token",
  "left_margin_inches": 1.0,
  "right_margin_inches": 1.0,
  "top_margin_inches": 1.0,
  "bottom_margin_inches": 1.0
}
```

### Step 7: Read Final Document
```json
{
  "blob_uri": "https://mystorageaccount.blob.core.windows.net/reports/quarterly.docx?sas_token"
}
```

---

## Error Handling

All tools return JSON responses with a `success` field:

**Success Response:**
```json
{
  "success": true,
  "message": "Operation completed successfully",
  "blob_uri": "..."
}
```

**Error Response:**
```json
{
  "success": false,
  "error": "Error description here"
}
```

Common errors:
- Invalid blob URI or SAS token
- Paragraph index out of range
- Invalid table data (empty rows/columns)
- Unsupported image format
- Invalid color RGB values (must be 0-255)

---

## Tips and Best Practices

1. **SAS Token Permissions**: Ensure your SAS tokens have read and write permissions
2. **Image Sizes**: Keep width between 1-7 inches for best results
3. **Table Data**: Ensure all rows have the same number of columns
4. **Font Names**: Use common fonts like Arial, Times New Roman, Calibri for compatibility
5. **Colors**: Use RGB values in range 0-255
6. **Paragraph Indices**: Use `read_docx` first to determine available paragraph indices
7. **Margins**: Standard margins are 1 inch; adjust based on your needs
8. **Base64 Images**: For large images, prefer Azure Blob URIs over base64

---

## Testing with MCP Inspector

1. Start the server in SSE mode
2. Connect MCP Inspector
3. Select a tool from the list
4. Fill in the parameters (use your Azure Blob URIs)
5. Click "Run Tool"
6. View the JSON response

---

## Deployment to Azure Container Apps (ACA)

This server is designed to work with Azure Container Apps. When deploying:

1. Set environment variables:
   - `PORT`: The port for SSE endpoint (default: 3001)
   - `LOG_LEVEL`: Logging level (default: DEBUG)
   - Optionally: `AZURE_STORAGE_CONNECTION_STRING`

2. The server supports SSE transport which is required for ACA deployment

3. Start command: `python -m src sse`

4. Health endpoint: The server exposes an SSE endpoint on the configured port

