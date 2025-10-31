# DOCX Formatting MCP Server

This is an MCP Server in Python for creating, reading, and fully editing DOCX files with Azure Blob Storage support. It includes the following features:

- **Create DOCX**: Create new DOCX files in Azure Blob Storage
- **Read DOCX**: Extract and read content, formatting, tables, and structure from DOCX files
- **Insert Text**: Add text to documents at specified positions
- **Edit Text**: Modify existing text while preserving or changing formatting
- **Table Management**: Insert and edit tables with customizable styles
- **Image Management**: Insert images from base64 or Azure Blob URIs
- **Formatting**: Complete control over fonts, sizes, colors, alignment, bold, italic, underline, and page margins
- **Connect to Agent Builder**: Test and debug the MCP server through Agent Builder
- **Debug in [MCP Inspector](https://github.com/modelcontextprotocol/inspector)**: Debug using the MCP Inspector
- **SSE Support**: Designed for deployment to Azure Container Apps (ACA) with SSE endpoint

## Get started with the DOCX Formatting MCP Server

> **Prerequisites**
>
> To run the MCP Server in your local dev machine, you will need:
>
> - [Python](https://www.python.org/)
> - (*Optional - if you prefer uv*) [uv](https://github.com/astral-sh/uv)
> - [Python Debugger Extension](https://marketplace.visualstudio.com/items?itemName=ms-python.debugpy)

## Prepare environment

There are two approaches to set up the environment for this project. You can choose either one based on your preference.

> Note: Reload VSCode or terminal to ensure the virtual environment python is used after creating the virtual environment.

| Approach | Steps |
| -------- | ----- |
| Using `uv` | 1. Create virtual environment: `uv venv` <br>2. Run VSCode Command "***Python: Select Interpreter***" and select the python from created virtual environment <br>3. Install dependencies (include dev dependencies): `uv pip install -r pyproject.toml --extra dev` |
| Using `pip` | 1. Create virtual environment: `python -m venv .venv` <br>2. Run VSCode Command "***Python: Select Interpreter***" and select the python from created virtual environment<br>3. Install dependencies (include dev dependencies): `pip install -e .[dev]` | 

After setting up the environment, you can run the server in your local dev machine via Agent Builder as the MCP Client to get started:
1. Open VS Code Debug panel. Select `Debug in Agent Builder` or press `F5` to start debugging the MCP server.
2. Set up Azure Blob Storage connection string in your environment or use Azure Blob URIs with SAS tokens.
3. Use AI Toolkit Agent Builder to test the server with document operations like creating, reading, or formatting DOCX files.
4. Click `Run` to test the server.

**Congratulations**! You have successfully run the DOCX Formatting MCP Server in your local dev machine via Agent Builder as the MCP Client.
![DebugMCP](https://raw.githubusercontent.com/microsoft/windows-ai-studio-templates/refs/heads/dev/mcpServers/mcp_debug.gif)

## What's included in the template

| Folder / File| Contents                                     |
| ------------ | -------------------------------------------- |
| `.vscode`    | VSCode files for debugging                   |
| `.aitk`      | Configurations for AI Toolkit                |
| `src`        | The source code for the DOCX formatting MCP server |
| `inspector`  | MCP Inspector configuration and dependencies |

## How to debug the DOCX Formatting MCP Server

> Notes:
> - [MCP Inspector](https://github.com/modelcontextprotocol/inspector) is a visual developer tool for testing and debugging MCP servers.
> - All debugging modes support breakpoints, so you can add breakpoints to the tool implementation code.

| Debug Mode | Description | Steps to debug |
| ---------- | ----------- | --------------- |
| Agent Builder | Debug the MCP server in the Agent Builder via AI Toolkit. | 1. Open VS Code Debug panel. Select `Debug in Agent Builder` and press `F5` to start debugging the MCP server.<br>2. Use AI Toolkit Agent Builder to test the server with DOCX operations. Server will be auto-connected to the Agent Builder.<br>3. Click `Run` to test the server with your prompts. |
| MCP Inspector | Debug the MCP server using the MCP Inspector. | 1. Install [Node.js](https://nodejs.org/)<br> 2. Set up Inspector: `cd inspector` && `npm install` <br> 3. Open VS Code Debug panel. Select `Debug SSE in Inspector (Edge)` or `Debug SSE in Inspector (Chrome)`. Press F5 to start debugging.<br> 4. When MCP Inspector launches in the browser, click the `Connect` button to connect this MCP server.<br> 5. Then you can `List Tools`, select a tool (like create_docx, read_docx, format_document), input parameters including Azure Blob URIs, and `Run Tool` to debug your server code.<br> |

## Default Ports and customizations

| Debug Mode | Ports | Definitions | Customizations | Note |
| ---------- | ----- | ------------ | -------------- |-------------- |
| Agent Builder | 3001 | [tasks.json](.vscode/tasks.json) | Edit [launch.json](.vscode/launch.json), [tasks.json](.vscode/tasks.json), [\_\_init\_\_.py](src/__init__.py), [mcp.json](.aitk/mcp.json) to change above ports. | N/A |
| MCP Inspector | 3001 (Server); 5173 and 3000 (Inspector) | [tasks.json](.vscode/tasks.json) | Edit [launch.json](.vscode/launch.json), [tasks.json](.vscode/tasks.json), [\_\_init\_\_.py](src/__init__.py), [mcp.json](.aitk/mcp.json) to change above ports.| N/A |

## Azure Blob Storage Setup

This MCP server works with Azure Blob Storage. To use it:

1. **Azure Blob URIs with SAS tokens**: The simplest approach is to use blob URIs with SAS (Shared Access Signature) tokens that include the necessary permissions.
   ```
   https://<account>.blob.core.windows.net/<container>/<filename>.docx?<sas-token>
   ```

2. **Connection String**: Alternatively, you can set the `AZURE_STORAGE_CONNECTION_STRING` environment variable.

## Available Tools

The server provides the following tools:

1. **create_docx**: Create a new DOCX file in Azure Blob Storage
2. **read_docx**: Read and extract content, formatting, and structure from a DOCX file
3. **insert_text**: Insert text at specified positions with optional styling
4. **edit_text**: Edit existing paragraph text with optional formatting preservation
5. **insert_or_edit_table**: Insert new tables or edit existing ones
6. **insert_or_edit_image**: Insert images from base64 or Azure Blob URIs
7. **format_document**: Apply text formatting (bold, italic, font, size, color, alignment) and page margins

## Feedback

If you have any feedback or suggestions for this template, please open an issue on the [AI Toolkit GitHub repository](https://github.com/microsoft/vscode-ai-toolkit/issues)
# docx_mcp_formatting
# docx_mcp_formatting
