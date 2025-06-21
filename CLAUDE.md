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
pytest src/__tests__/test_command_parser.py

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
   - Provides `get_document_features()` method for accessing document features
   - `is_connected()` method for connection status checking

3. **src/tools/hwp_table_tools.py**: Specialized module for table operations
   - `insert_table()`: Creates tables with specified dimensions
   - `fill_table_with_data()`: Populates table cells (supports JSON strings and Python lists)
   - `create_table_with_data()`: Creates and fills table in one operation
   - Advanced table features: `_apply_table_style()`, `_sort_table()`, `_merge_cells()`, `_split_cell()`

4. **src/tools/hwp_advanced_features.py**: Advanced features module
   - Document formatting: `set_page()`, `set_header_footer()`, `set_paragraph()`
   - Content manipulation: `insert_image()`, `find_replace()`, `insert_shape()`
   - Export/Import: `export_pdf()`, `save_as_template()`, `apply_template()`
   - Document structure: `create_toc()`

5. **src/tools/hwp_document_features.py**: Document editing advanced features
   - Notes: `insert_footnote()`, `insert_endnote()`
   - Navigation: `insert_hyperlink()`, `insert_bookmark()`, `goto_bookmark()`
   - Review: `insert_comment()`, `search_and_highlight()`
   - Security: `insert_watermark()`, `set_document_password()`
   - Fields: `insert_field()` for date, time, page numbers, etc.

6. **src/tools/hwp_exceptions.py**: Custom exception classes
   - Hierarchical exception structure: HwpError → specific exceptions
   - Connection, Document, Table, Image, PDF, Template, Field, Parameter, Operation errors
   - `handle_hwp_error` decorator for consistent error handling

7. **src/tools/constants.py**: Centralized constants
   - Unit conversions: HWPUNIT_PER_MM, HWPUNIT_PER_PT
   - Table limits and defaults
   - Image formats, colors, field types
   - Page sizes, orientations, margins

8. **src/tools/config.py**: Runtime configuration management
   - HwpConfig dataclass with all configurable options
   - ConfigManager singleton for global config access
   - JSON-based config persistence
   - Environment variable support

9. **src/tools/hwp_utils.py**: Common utility functions to reduce code duplication
   - Decorators: `@require_hwp_connection`, `@safe_hwp_operation`
   - Common operations: `set_font_properties()`, `move_to_table_cell()`
   - Data processing: `parse_table_data()`, `validate_table_coordinates()`
   - Error handling: `execute_with_retry()`, `log_operation_result()`

### Key Design Patterns

- **MCP Tool Registration**: Each operation is a separate @mcp.tool() function
- **COM Automation**: Uses pywin32 to control HWP through Windows COM interface
- **Error Handling**: Custom exception hierarchy with specific error types
- **Security Module**: Uses FilePathCheckerModule.dll to bypass HWP security dialogs
- **Modular Features**: Features grouped into logical modules
- **Configuration**: Centralized config and constants management

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

3. **Unit Conversions** (from constants.py):
   - HWPUNIT: 1mm = 2834.64 HWPUNIT
   - Font size: 1pt = 100 HWPUNIT
   - Paper size: stored in HWPUNIT

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
   - Document features → hwp_document_features.py
   - Complex features → hwp_advanced_features.py or new module

3. **Add Exception Classes** if needed in hwp_exceptions.py:
   - Inherit from appropriate base class (HwpError or specific category)
   - Include meaningful error messages

4. **Add Constants** in constants.py:
   - Group related constants together
   - Use descriptive names and comments

5. **Add Tests** in src/__tests__/:
   - Follow existing test patterns
   - Test both success and error cases

6. **Update Documentation**:
   - Add example to README.md
   - Update changelog with feature description

## Korean Language Considerations

When working with Korean text:
- HWP is designed for Korean documents, so Korean text handling is native
- Use UTF-8 encoding throughout the project
- Test with various Korean text inputs including special characters

## Security Module Path

The security module path can be configured via:
1. config.py: `security_module_path` setting
2. Environment variable: Set in HWP_MCP_CONFIG path
3. Default: "D:/hwp-mcp/security_module/FilePathCheckerModuleExample.dll"