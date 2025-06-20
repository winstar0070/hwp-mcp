# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

HWP-MCP is a Model Context Protocol (MCP) server that enables AI models like Claude to control the Korean word processor HWP (한글) through Windows COM automation. This Windows-only project acts as a bridge between Claude and HWP, allowing document automation tasks.

## Development Commands

```bash
# Install dependencies
pip install -r requirements.txt
pip install mcp

# Run all tests
pytest src/__tests__/

# Run specific test file
pytest src/__tests__/test_font_features.py
pytest src/__tests__/test_advanced_features.py
pytest src/__tests__/test_hwp_controller.py

# Run tests with coverage
pytest --cov=src src/__tests__/

# Run the MCP server (usually started automatically by Claude)
python hwp_mcp_stdio_server.py
```

## Architecture

### Core Flow
1. Claude sends commands via MCP protocol → `hwp_mcp_stdio_server.py`
2. Server routes commands to appropriate tool functions
3. Tools use `HwpController` to interact with HWP via COM
4. Results are returned to Claude through MCP response

### Core Components

1. **hwp_mcp_stdio_server.py**: Main entry point that sets up the MCP server
   - Uses FastMCP library for MCP protocol handling
   - Registers all HWP control tools with @mcp.tool() decorator
   - Manages global HWP controller instance (singleton)
   - Handles stdio communication with Claude

2. **src/tools/hwp_controller.py**: Central controller for HWP interactions
   - Singleton pattern for managing single HWP instance
   - Core methods: `connect()`, `create_new_document()`, `save_document()`, `insert_text()`
   - Special handling for table vs non-table text insertion
   - Provides `get_advanced_features()` method for accessing advanced functionality

3. **src/tools/hwp_table_tools.py**: Specialized module for table operations
   - `insert_table()`: Creates tables with specified dimensions
   - `fill_table_data()`: Populates table cells (supports JSON strings and Python lists)
   - `fill_column_numbers()`: Fills a column with sequential numbers

4. **src/tools/hwp_advanced_features.py**: Advanced features module
   - Document formatting: `set_page()`, `set_header_footer()`, `set_paragraph()`
   - Content manipulation: `insert_image()`, `find_replace()`, `insert_shape()`
   - Export/Import: `export_pdf()`, `save_as_template()`, `apply_template()`
   - Document structure: `create_toc()`

### Key Design Patterns

- **MCP Tool Registration**: Each operation is a separate @mcp.tool() function
- **COM Automation**: Uses pywin32 to control HWP through Windows COM interface
- **Error Handling**: All operations wrapped in try-except with logging
- **Security Module**: Uses FilePathCheckerModule.dll to bypass HWP security dialogs
- **Modular Features**: Features grouped into logical modules (table, advanced, etc.)

### Important Implementation Details

1. **Text Input Behavior**:
   ```python
   # Inside tables: Direct insertion
   if in_table:
       self._insert_text_direct(text)
   # Outside tables: Move to end first
   else:
       self.hwp.Run("MoveDocEnd")
       self._insert_text_direct(text)
   ```

2. **Font Styling Methods**:
   - `insert_text_with_font()`: Set style → Insert text
   - `apply_font_to_selection()`: Select text → Apply style
   - Bold/Italic use explicit 1/0 values (not boolean)

3. **Unit Conversions**:
   - HWPUNIT: 1mm = 2834.64 HWPUNIT
   - Font size: 1pt = 100 HWPUNIT
   - Paper size: stored in 0.01mm units

4. **Batch Operations**:
   ```python
   hwp_batch_operations([
       {"operation": "hwp_create"},
       {"operation": "hwp_insert_text", "params": {"text": "Hello"}}
   ])
   ```

5. **Global Instance Management**:
   - `hwp_controller` and `hwp_table_tools` are global singletons
   - Created on first access via `get_hwp_controller()`

## Adding New Features

1. **Add MCP Tool** in hwp_mcp_stdio_server.py:
   ```python
   @mcp.tool()
   def hwp_new_feature(param1: str, param2: int = None) -> str:
       """Tool description for Claude."""
       hwp = get_hwp_controller()
       # Implementation
   ```

2. **Implement Logic** in appropriate module:
   - Simple operations → hwp_controller.py
   - Table-related → hwp_table_tools.py
   - Complex features → hwp_advanced_features.py or new module

3. **Add Tests** in src/__tests__/:
   - Follow existing test patterns
   - Test both success and error cases

4. **Update Documentation**:
   - Add example to README.md
   - Update changelog with feature description