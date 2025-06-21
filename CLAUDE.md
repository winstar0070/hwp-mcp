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
pytest src/__tests__/test_hwp_controller.py
pytest src/__tests__/test_hwp_utils.py
pytest src/__tests__/test_hwp_exceptions.py
pytest src/__tests__/test_config.py
pytest src/__tests__/test_hwp_batch_processor.py
pytest src/__tests__/test_font_features.py
pytest src/__tests__/test_advanced_features.py
pytest src/__tests__/test_document_features.py
pytest src/__tests__/test_table_advanced.py
pytest src/__tests__/test_integration.py
pytest src/__tests__/test_chart_batch.py
pytest src/__tests__/test_command_parser.py

# Run tests with coverage
pytest --cov=src src/__tests__/

# Run the MCP server (usually started automatically by Claude)
python hwp_mcp_stdio_server.py
```

## Architecture

### Core Flow
1. Claude sends commands via MCP protocol → `hwp_mcp_stdio_server.py`
2. Server routes commands to appropriate tool functions (46 tools total)
3. Tools use `HwpController` to interact with HWP via COM
4. Results are returned to Claude through MCP response

### Core Components

1. **hwp_mcp_stdio_server.py**: Main entry point that sets up the MCP server
   - Uses FastMCP library for MCP protocol handling
   - Registers all HWP control tools with @mcp.tool() decorator
   - Manages global HWP controller instance (singleton)
   - Handles stdio communication with Claude
   - Logging configured to hwp_mcp_stdio_server.log

2. **src/tools/hwp_controller.py**: Central controller for HWP interactions
   - Singleton pattern for managing single HWP instance
   - Core methods: `connect()`, `create_new_document()`, `save_document()`, `insert_text()`
   - Special handling for table vs non-table text insertion
   - Provides `get_document_features()`, `get_advanced_features()`, `get_chart_features()` methods
   - `is_connected()` method for connection status checking
   - Uses decorators from hwp_utils for connection validation

3. **src/tools/hwp_table_tools.py**: Specialized module for table operations
   - `insert_table()`: Creates tables with specified dimensions
   - `fill_table_with_data()`: Populates table cells (supports JSON strings and Python lists)
   - `create_table_with_data()`: Creates and fills table in one operation
   - Advanced table features: `_apply_table_style()`, `_sort_table()`, `_merge_cells()`, `_split_cell()`
   - Standalone functions accessible via MCP: `apply_table_style()`, `sort_table_by_column()`, etc.

4. **src/tools/hwp_advanced_features.py**: Advanced features module
   - Document formatting: `set_page()`, `set_header_footer()`, `set_paragraph()`
   - Content manipulation: `insert_image()`, `find_replace()`, `insert_shape()`
   - Export/Import: `export_pdf()`, `save_as_template()`, `apply_template()`
   - Document structure: `create_toc()`
   - Special handling for Korean number sequences and vertical text

5. **src/tools/hwp_document_features.py**: Document editing advanced features
   - Notes: `insert_footnote()`, `insert_endnote()`
   - Navigation: `insert_hyperlink()`, `insert_bookmark()`, `goto_bookmark()`
   - Review: `insert_comment()`, `search_and_highlight()`
   - Security: `insert_watermark()`, `set_document_password()`
   - Fields: `insert_field()` for date, time, page numbers, etc.

6. **src/tools/hwp_chart_features.py**: Chart and visualization features
   - Chart insertion: `insert_chart()` with various types (column, bar, line, pie)
   - Equation editor: `insert_equation()` with basic LaTeX support
   - Templates: `insert_equation_template()` for common equations
   - Diagrams: `insert_diagram()` for simple flowcharts
   - LaTeX to HWP equation conversion

7. **src/tools/hwp_batch_processor.py**: Batch processing and transactions
   - Transaction support: `transaction()` context manager with automatic rollback
   - Batch execution: `execute_batch()` with error handling
   - Large data handling: `insert_large_table_data()` with chunk processing
   - Multi-document processing: `process_multiple_documents()`
   - Progress callbacks for long operations

8. **src/tools/hwp_exceptions.py**: Custom exception classes
   - Hierarchical exception structure: HwpError → specific exceptions
   - Connection, Document, Table, Image, PDF, Template, Field, Parameter, Operation, Batch errors
   - `handle_hwp_error` decorator for consistent error handling
   - Transaction-specific exceptions with rollback support

9. **src/tools/constants.py**: Centralized constants
   - Unit conversions: HWPUNIT_PER_MM = 2834.64, HWPUNIT_PER_PT = 100
   - Table limits: TABLE_MAX_ROWS = 100, TABLE_MAX_COLS = 100
   - Image formats, colors, field types
   - Page sizes, orientations, margins
   - Chart types and equation templates
   - Korean-specific constants (NUMBER_SEQUENCE_KOREAN, VERTICAL_KOREAN)

10. **src/tools/config.py**: Runtime configuration management
    - HwpConfig dataclass with all configurable options
    - ConfigManager singleton for global config access
    - JSON-based config persistence
    - Environment variable support (HWP_MCP_CONFIG)
    - Security module path configuration

11. **src/tools/hwp_utils.py**: Common utility functions to reduce code duplication
    - Decorators: `@require_hwp_connection`, `@safe_hwp_operation`
    - Common operations: `set_font_properties()`, `move_to_table_cell()`
    - Data processing: `parse_table_data()`, `validate_table_coordinates()`
    - Error handling: `execute_with_retry()`, `log_operation_result()`

12. **src/tools/error_handling_guide.py**: Enhanced error handling framework
    - `enhanced_error_handler` decorator for consistent error processing
    - Exception type mapping for automatic conversion
    - File path and table coordinate validation functions
    - User-friendly error message formatting

### Key Design Patterns

- **MCP Tool Registration**: Each operation is a separate @mcp.tool() function
- **COM Automation**: Uses pywin32 to control HWP through Windows COM interface
- **Enhanced Error Handling**: Specific exception types with user-friendly messages
- **Security Module**: Uses FilePathCheckerModule.dll to bypass HWP security dialogs
- **Modular Features**: Features grouped into logical modules (12 total in src/tools/)
- **Configuration**: Centralized config and constants management
- **Singleton Pattern**: Global instances for HwpController and ConfigManager
- **Context Managers**: Transaction support for batch operations
- **Decorator Pattern**: Connection validation and error handling via decorators
- **Test Coverage**: Comprehensive unit tests with 119 test cases covering core modules

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

4. **Batch Operations Format**:
   ```python
   hwp_batch_operations([
       {"action": "insert_text", "params": {"text": "Hello"}},
       {"action": "insert_table", "params": {"rows": 3, "cols": 3}}
   ])
   ```

5. **Transaction Support**:
   ```python
   with batch_processor.transaction():
       # Multiple operations
       # Automatic rollback on error
   ```

6. **Global Instance Management**:
   - `hwp_controller` and `hwp_table_tools` are global singletons
   - Created on first access via `get_hwp_controller()`
   - Thread-safe with Lock mechanism

7. **Table Data Format**:
   - Accepts JSON strings or Python lists
   - Auto-converts to 2D string arrays
   - Empty cells handled as empty strings

## Adding New Features

1. **Add MCP Tool** in hwp_mcp_stdio_server.py:
   ```python
   @mcp.tool()
   def hwp_new_feature(param1: str, param2: int = None) -> str:
       """Tool description for Claude."""
       hwp = get_hwp_controller()
       if not hwp:
           return "Error: Failed to connect to HWP program"
       # Implementation
   ```

2. **Implement Logic** in appropriate module:
   - Simple operations → hwp_controller.py
   - Table-related → hwp_table_tools.py
   - Document features → hwp_document_features.py
   - Charts/equations → hwp_chart_features.py
   - Batch operations → hwp_batch_processor.py
   - Complex features → hwp_advanced_features.py or new module

3. **Add Exception Classes** if needed in hwp_exceptions.py:
   - Inherit from appropriate base class (HwpError or specific category)
   - Include meaningful error messages

4. **Add Constants** in constants.py:
   - Group related constants together
   - Use descriptive names and comments

5. **Add Tests** in src/__tests__/:
   - Follow existing test patterns with Mock objects
   - Test both success and error cases
   - Use fixtures for common setup

6. **Update Documentation**:
   - Add example to README.md
   - Update changelog with feature description

## Korean Language Considerations

When working with Korean text:
- HWP is designed for Korean documents, so Korean text handling is native
- Use UTF-8 encoding throughout the project
- Test with various Korean text inputs including special characters
- Consider Korean-specific formatting needs (vertical text, etc.)
- Special constants: NUMBER_SEQUENCE_KOREAN = "1부터 10까지", VERTICAL_KOREAN = "세로"

## Security Module Path

The security module path can be configured via:
1. config.py: `security_module_path` setting
2. Environment variable: Set in HWP_MCP_CONFIG path
3. Default: "D:/hwp-mcp/security_module/FilePathCheckerModuleExample.dll"

## Common Error Patterns

1. **HWP Not Running**: Check `is_connected()` before operations
2. **Table Position**: Use `PositionToFieldEx()` to verify cursor is in table
3. **File Not Found**: Use absolute paths with `os.path.abspath()`
4. **Parameter Validation**: Check ranges in constants.py (TABLE_MAX_ROWS, etc.)
5. **Transaction Failures**: Always use context manager for automatic cleanup
6. **Empty Table Cells**: Handle as empty strings, not None
7. **Font Settings**: Use explicit 1/0 for boolean properties
8. **Enhanced Error Handling**: Use `enhanced_error_handler` decorator for consistent error processing
9. **Exception Specificity**: Catch specific exceptions (OSError, AttributeError, ValueError) rather than generic Exception

## User Preferences (from ~/.claude/CLAUDE.md)

- 작업이 완료되면 자동으로 git commit과 push 수행 (코드 문법 확인 필수)
- 한국어로 답변
- 웹 검색 적극 활용