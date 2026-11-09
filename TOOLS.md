# Word MCP Server — Tool Reference

> 81 tools across 13 modules.

---

### Document Lifecycle

| Tool | Description |
|------|-------------|
| `word_get_status` | Query current Word/document state (NO_WORD/NO_DOC/DOC_ACTIVE/DIALOG) |
| `word_document` | Universal entry: open by path or use active doc |
| `word_open` | Open existing .docx file |
| `word_save` | Save current document (auto `.bak` backup) |
| `word_save_as` | Save with new path (auto-detect format) |
| `word_close` | Close current document |
| `word_get_info` | Get word/paragraph/page count etc. |
| `word_get_structure` | Get heading outline with paragraph indices |
| `word_quit` | Quit Microsoft Word entirely |

### Content Input

| Tool | Description |
|------|-------------|
| `word_insert_paragraph` | Insert paragraph breaks |
| `word_insert_break` | Insert page break or horizontal line |
| `word_insert_list` | Insert bullet or numbered list |
| `word_set_columns` | Set section column count (newsletter layout) |
| `word_insert_section_break` | Insert section break (nextPage/continuous/evenPage/oddPage) |
| `word_insert_file` | Insert content of another .docx at cursor |
| `word_backspace` | Delete characters before cursor |
| `word_get_text` | Read full document text |
| `word_get_paragraph` | Read a specific paragraph's text |
| `word_get_table_data` | Extract table content as structured data (rows×columns grid) |
| `word_get_comments` | List all comments with author info |
| `word_get_bookmarks` | List all bookmarks |
| `word_get_lists` | List all bullet/numbered lists with hierarchy |
| `word_get_sections` | List sections with page setup info (orientation, columns, page size) |
| `word_export_to_pdf` | Export document to PDF |

### Streaming Document (Recommended Content Path)

| Tool | Description |
|------|-------------|
| `word_stream_start` | Start a streaming session: create document, configure page setup, base styles (`baseStyleProfile` for font/paragraph presets per style name), and template |
| `word_stream_block` | Write Markdown blocks into the streaming document (instant preview in Word). Supports all Markdown syntax: headings, bold, italic, lists, tables, code blocks, blockquotes, links |
| `word_stream_end` | End streaming session: save document, optionally export to PDF. Returns block count, character count, and elapsed time |

### Links, References & Comments

| Tool | Description |
|------|-------------|
| `word_add_hyperlink` | Add hyperlink (supports screenTip & subAddress) |
| `word_add_footnote` | Add footnote (auto-numbered) |
| `word_add_comment` | Add comment |

### Find, Selection & Navigation

| Tool | Description |
|------|-------------|
| `word_find_text` | Find text and position cursor at the match |
| `word_find_replace` | Find and replace text |
| `word_go_to` | Navigate to page/section/bookmark/end |
| `word_go_to_paragraph` | Navigate to specific paragraph by 1-based index |
| `word_select_all` | Select all content |
| `word_select_text` | Select text by character range |
| `word_select_current` | Select word or paragraph at cursor |
| `word_delete` | Delete selected content |

### Clipboard & Undo

| Tool | Description |
|------|-------------|
| `word_clipboard` | Copy/cut/paste clipboard content |
| `word_undo_redo` | Undo or redo last action(s) |
| `word_get_cursor_info` | Get cursor position & selection info |

### Formatting

| Tool | Description |
|------|-------------|
| `word_set_font` | Set font name/size/bold/italic/color/strikethrough/superscript/subscript |
| `word_set_paragraph` | Set alignment/indent/spacing |
| `word_apply_style` | Apply named style (Heading 1, Title, etc.) |
| `word_set_page_setup` | Set margins/orientation/paper size |
| `word_set_properties` | Set document metadata (title, author, etc.) |
| `word_list_styles` | List available styles |
| `word_set_body_indent` | Apply first-line indent to all 'Normal' paragraphs (Chinese academic papers) |
| `word_set_track_changes` | Enable/disable Track Changes (revision markup) |
| `word_track_changes_apply` | Accept or reject all tracked changes |

### Tables

| Tool | Description |
|------|-------------|
| `word_insert_table` | Insert table with optional data |
| `word_edit_cell` | Edit table cell text |
| `word_edit_cells` | Batch-fill multiple cells with 2D data array |
| `word_add_table_row` | Add row (optional data) |
| `word_delete_table_row` | Delete row |
| `word_add_table_column` | Add column (left of specified column or append) |
| `word_delete_table_column` | Delete column |
| `word_set_table_borders` | Set table border style/color/width |
| `word_set_table_shading` | Set table or header row background color |
| `word_merge_table_cells` | Merge cell range |
| `word_set_column_width` | Set column width in points |
| `word_set_row_height` | Set row height in points |
| `word_set_cell_font` | Set font for specific cell |
| `word_set_cell_vertical_alignment` | Set cell vertical alignment (top/center/bottom) |
| `word_apply_table_style` | Apply built-in Word table style by name |
| `word_table_to_text` | Convert table to plain text (remove table structure) |
| `word_text_to_table` | Convert selected delimited text to table |

### Media

| Tool | Description |
|------|-------------|
| `word_insert_image` | Insert image from file |
| `word_insert_chart` | Insert chart (column/bar/line/pie/area, child process sets data) |
| `word_insert_textbox` | Insert textbox (position/size/orientation) |

### Document Structure

| Tool | Description |
|------|-------------|
| `word_set_page_region` | Set page header or footer |
| `word_set_page_numbers` | Add page numbers |
| `word_insert_toc` | Insert table of contents |
| `word_add_bookmark` | Add bookmark |
| `word_set_watermark` | Set or remove watermark (e.g. "DRAFT") |

### Semantic Navigation

| Tool | Description |
|------|-------------|
| `word_locate` | Resolve heading/paragraph/table/bookmark position (read-only) |
| `word_select_at` | Move cursor to semantic location (heading/paragraph/table/bookmark) |
| `word_insert_at` | Insert Markdown at semantic location |
| `word_edit_cell_at` | Edit table cell at semantic location |

### Variable Replacement

| Tool | Description |
|------|-------------|
| `word_replace_variables` | Replace {{placeholder}} variables in document with provided values |

### Batch Operations

| Tool | Description |
|------|-------------|
| `word_batch_ops` | Execute multiple operations in a single call for performance |

### Diagnostics

| Tool | Description |
|------|-------------|
| `word_where_am_i` | Get current cursor position context (heading, section, table, etc.) |

---

Back to [README](./README.md)
