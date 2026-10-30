"""
Word Office MCP Tool Definitions
Defines all tools that the AI agent can invoke inside Microsoft Word.
Each tool specifies its JSON schema, description, and whether it is destructive
(modifies or deletes existing document content).
"""

from typing import List, Dict, Any, Optional

WORD_TOOLS: List[Dict[str, Any]] = [
    {
        "name": "word_get_document_outline",
        "description": "Reads the current Word document outline: all paragraphs (index, text snippet, style), existing tables (index, row/column counts), and active cursor/selection status. Call this first to understand the document structure.",
        "inputSchema": {
            "type": "object",
            "properties": {},
            "required": []
        },
        "destructive": False
    },
    {
        "name": "word_read_paragraphs",
        "description": "Reads the full text, style, and formatting of a range of paragraphs by index.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "start_index": {
                    "type": "integer",
                    "description": "0-based starting paragraph index."
                },
                "end_index": {
                    "type": "integer",
                    "description": "0-based ending paragraph index (inclusive). If omitted, reads through start_index only."
                }
            },
            "required": ["start_index"]
        },
        "destructive": False
    },
    {
        "name": "word_search",
        "description": "Searches the document for occurrences of a text string and returns paragraph indices and matched ranges.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "search_text": {
                    "type": "string",
                    "description": "The exact text to find in the document."
                },
                "match_case": {
                    "type": "boolean",
                    "description": "Whether to perform a case-sensitive search. Default false.",
                    "default": False
                }
            },
            "required": ["search_text"]
        },
        "destructive": False
    },
    {
        "name": "word_insert_paragraph",
        "description": "Inserts a new paragraph with a specific Word built-in style. Location can be 'end' (document end), 'start', 'cursor', or relative to an index ('after:N', 'before:N').",
        "inputSchema": {
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "The text content of the paragraph."
                },
                "style": {
                    "type": "string",
                    "description": "Word built-in style name, e.g. 'Heading 1', 'Heading 2', 'Heading 3', 'Normal', 'Title', 'Subtitle', 'Quote', 'Intense Quote'. Default 'Normal'.",
                    "default": "Normal"
                },
                "location": {
                    "type": "string",
                    "description": "Where to insert: 'end', 'start', 'cursor', 'after:<index>', or 'before:<index>'. Default 'end'.",
                    "default": "end"
                }
            },
            "required": ["text"]
        },
        "destructive": False
    },
    {
        "name": "word_insert_list",
        "description": "Inserts a native bulleted or numbered list into Word at the specified location.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "items": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List item strings."
                },
                "list_type": {
                    "type": "string",
                    "enum": ["bullet", "number"],
                    "description": "Type of list: 'bullet' or 'number'. Default 'bullet'.",
                    "default": "bullet"
                },
                "location": {
                    "type": "string",
                    "description": "Where to insert: 'end', 'cursor', 'after:<index>', 'before:<index>'. Default 'end'.",
                    "default": "end"
                }
            },
            "required": ["items"]
        },
        "destructive": False
    },
    {
        "name": "word_insert_table",
        "description": "Inserts a native Microsoft Word table populated with the given columns/headers and rows, applying executive styling and auto-fit. Use this for tables rather than raw text.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "headers": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of header column names (e.g. ['product_name', 'quantity', 'price'])."
                },
                "columns": {
                    "type": "integer",
                    "description": "Number of columns (used if headers is not specified, generates default Colonna 1, Colonna 2...)."
                },
                "rows": {
                    "description": "2D array of row cell values, or an integer specifying number of rows."
                },
                "table_style": {
                    "type": "string",
                    "description": "Word native table style name (e.g. 'Grid Table 4 - Accent 1', 'Table Grid', 'Grid Table 1 Light'). Default 'Grid Table 4 - Accent 1'.",
                    "default": "Grid Table 4 - Accent 1"
                },
                "location": {
                    "type": "string",
                    "description": "Where to insert: 'end', 'cursor', 'after:<index>', 'before:<index>'. Default 'end'.",
                    "default": "end"
                }
            },
            "required": ["headers", "rows"]
        },
        "destructive": False
    },
    {
        "name": "word_format_table",
        "description": "Formats an existing native Word table by 0-based index: changes style, header row formatting, or cell alignments.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "0-based table index in the document."
                },
                "style": {
                    "type": "string",
                    "description": "Built-in Word table style name."
                },
                "header_bold": {
                    "type": "boolean",
                    "description": "Whether to bold the top header row. Default true.",
                    "default": True
                },
                "autofit": {
                    "type": "boolean",
                    "description": "Whether to auto-fit table columns to contents. Default true.",
                    "default": True
                }
            },
            "required": ["table_index"]
        },
        "destructive": False
    },
    {
        "name": "word_set_table_cell",
        "description": "Sets the text content and optional formatting of a specific cell in an existing Word table.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "0-based table index."
                },
                "row_index": {
                    "type": "integer",
                    "description": "0-based row index (0 is typically the header row)."
                },
                "col_index": {
                    "type": "integer",
                    "description": "0-based column index."
                },
                "text": {
                    "type": "string",
                    "description": "New text value for the cell."
                },
                "bold": {
                    "type": "boolean",
                    "description": "Optional bold setting."
                }
            },
            "required": ["table_index", "row_index", "col_index", "text"]
        },
        "destructive": False
    },
    {
        "name": "word_add_table_row",
        "description": "Appends a new row to an existing Word table.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "0-based table index."
                },
                "values": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Cell values for the new row."
                }
            },
            "required": ["table_index", "values"]
        },
        "destructive": False
    },
    {
        "name": "word_format_text",
        "description": "Applies character formatting (bold, italic, font color, highlight, size) to a paragraph or to specific text inside a paragraph.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "paragraph_index": {
                    "type": "integer",
                    "description": "0-based paragraph index."
                },
                "search_text": {
                    "type": "string",
                    "description": "Optional substring to format inside the paragraph. If omitted, formats the entire paragraph."
                },
                "bold": {
                    "type": "boolean",
                    "description": "Set text bold."
                },
                "italic": {
                    "type": "boolean",
                    "description": "Set text italic."
                },
                "underline": {
                    "type": "boolean",
                    "description": "Set text underlined."
                },
                "color": {
                    "type": "string",
                    "description": "Hex color string (e.g. '#1E3A8A' or '#DC2626')."
                },
                "highlight": {
                    "type": "string",
                    "description": "Highlight color name (e.g. 'Yellow', 'Green', 'Cyan')."
                },
                "size_pt": {
                    "type": "number",
                    "description": "Font size in points (e.g. 14, 18)."
                },
                "alignment": {
                    "type": "string",
                    "enum": ["left", "center", "right", "justified"],
                    "description": "Paragraph text alignment: 'left', 'center', 'right', or 'justified'."
                },
                "style": {
                    "type": "string",
                    "description": "Paragraph style to apply: 'Title', 'Subtitle', 'Heading 1', 'Heading 2', 'Normal'."
                }
            },
            "required": ["paragraph_index"]
        },
        "destructive": False
    },
    {
        "name": "word_apply_style",
        "description": "Changes the Word paragraph style of an existing paragraph (e.g. from Normal to Heading 1, Title, etc.).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "paragraph_index": {
                    "type": "integer",
                    "description": "0-based paragraph index."
                },
                "style": {
                    "type": "string",
                    "description": "Built-in Word style name (e.g. 'Heading 1', 'Heading 2', 'Normal', 'Title')."
                }
            },
            "required": ["paragraph_index", "style"]
        },
        "destructive": False
    },
    {
        "name": "word_add_comment",
        "description": "Inserts an inline Word review comment on a paragraph or a specific phrase in a paragraph.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "paragraph_index": {
                    "type": "integer",
                    "description": "0-based paragraph index."
                },
                "comment": {
                    "type": "string",
                    "description": "The review comment text."
                },
                "search_text": {
                    "type": "string",
                    "description": "Optional specific phrase to attach the comment to."
                }
            },
            "required": ["paragraph_index", "comment"]
        },
        "destructive": False
    },
    {
        "name": "word_insert_page_break",
        "description": "Inserts a native page break at the specified location.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "location": {
                    "type": "string",
                    "description": "Where to insert: 'end', 'cursor', 'after:<index>', 'before:<index>'. Default 'end'.",
                    "default": "end"
                }
            },
            "required": []
        },
        "destructive": False
    },
    {
        "name": "word_select",
        "description": "Moves the user's cursor / selection in Microsoft Word to the specified paragraph index, making it visibly focused.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "paragraph_index": {
                    "type": "integer",
                    "description": "0-based paragraph index to select and bring into view."
                }
            },
            "required": ["paragraph_index"]
        },
        "destructive": False
    },
    {
        "name": "word_set_paragraph_text",
        "description": "Overwrites the entire text of an existing paragraph. Destructive: requires user approval if modifying existing user content.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "paragraph_index": {
                    "type": "integer",
                    "description": "0-based paragraph index to rewrite."
                },
                "new_text": {
                    "type": "string",
                    "description": "The new replacement text."
                }
            },
            "required": ["paragraph_index", "new_text"]
        },
        "destructive": True
    },
    {
        "name": "word_replace_text",
        "description": "Finds and replaces text across the document. Destructive: requires user approval before modifying existing content.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "search_text": {
                    "type": "string",
                    "description": "The text to find."
                },
                "replacement_text": {
                    "type": "string",
                    "description": "The replacement text."
                },
                "match_case": {
                    "type": "boolean",
                    "description": "Whether search is case-sensitive. Default false.",
                    "default": False
                }
            },
            "required": ["search_text", "replacement_text"]
        },
        "destructive": True
    },
    {
        "name": "word_delete_paragraph",
        "description": "Deletes a paragraph from the document by index. Destructive: requires user approval.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "paragraph_index": {
                    "type": "integer",
                    "description": "0-based paragraph index to remove."
                }
            },
            "required": ["paragraph_index"]
        },
        "destructive": True
    },
    {
        "name": "word_delete_table",
        "description": "Deletes an entire native Word table by 0-based index. Use this to remove duplicate or unwanted tables. Destructive: requires user approval.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "0-based table index to delete."
                }
            },
            "required": ["table_index"]
        },
        "destructive": True
    },
    {
        "name": "word_delete_column",
        "description": "Deletes a column from an existing native Word table by 0-based column index in-place. Never delete and recreate the whole table just to drop a column; use this tool.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "0-based table index (default 0)."
                },
                "col_index": {
                    "type": "integer",
                    "description": "0-based column index to delete (0 is the first column)."
                }
            },
            "required": ["col_index"]
        },
        "destructive": True
    },
    {
        "name": "word_delete_row",
        "description": "Deletes a row from an existing native Word table by 0-based row index in-place.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "0-based table index (default 0)."
                },
                "row_index": {
                    "type": "integer",
                    "description": "0-based row index to delete."
                }
            },
            "required": ["row_index"]
        },
        "destructive": True
    },
    {
        "name": "word_read_table",
        "description": "Reads the full content and dimensions of an existing Word table by 0-based index.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "0-based table index to read."
                }
            },
            "required": ["table_index"]
        },
        "destructive": False
    },
    {
        "name": "word_clear_document",
        "description": "Clears all content from the document body, leaving a fresh blank document. Destructive: requires user approval.",
        "inputSchema": {
            "type": "object",
            "properties": {},
            "required": []
        },
        "destructive": True
    },
    {
        "name": "word_insert_image",
        "description": "Inserts an image or generated diagram into the Word document at the specified location. Supports descriptive topic keywords (e.g. 'quantum circuit', 'diagram', 'chart', 'workflow', 'architecture'), base64 data URIs, or image URLs. High-resolution graphics are rendered natively and inserted with an aligned caption directly underneath.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "source": {
                    "type": "string",
                    "description": "Image source: descriptive topic keyword (e.g. 'quantum circuit', 'workflow', 'chart'), HTTPS URL, or base64 data string."
                },
                "caption": {
                    "type": "string",
                    "description": "Optional caption text placed directly underneath the picture."
                },
                "location": {
                    "type": "string",
                    "description": "Where to insert: 'end', 'start', 'cursor', 'after:<index>', 'before:<index>'. Default 'end'.",
                    "default": "end"
                },
                "width_px": {
                    "type": "number",
                    "description": "Optional image width in points/pixels (e.g. 250, 400)."
                },
                "height_px": {
                    "type": "number",
                    "description": "Optional image height in points/pixels."
                }
            },
            "required": ["source"]
        },
        "destructive": False
    },
    {
        "name": "word_generate_image",
        "description": "Generates a realistic AI image or illustration using Google Gemini Imagen 3 from a text prompt and inserts it directly into the Word document with an aligned caption. Active ONLY when a Gemini API key is configured in Taskpane Settings.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "prompt": {
                    "type": "string",
                    "description": "Detailed description of the image to generate (e.g. 'modern data center with illuminated server racks')."
                },
                "caption": {
                    "type": "string",
                    "description": "Optional caption text placed directly underneath the picture."
                },
                "location": {
                    "type": "string",
                    "description": "Where to insert: 'end', 'start', 'cursor', 'after:<index>', 'before:<index>'. Default 'end'.",
                    "default": "end"
                },
                "aspect_ratio": {
                    "type": "string",
                    "enum": ["16:9", "1:1", "4:3", "3:4", "9:16"],
                    "description": "Aspect ratio of the generated image. Default '16:9'.",
                    "default": "16:9"
                },
                "width_px": {
                    "type": "number",
                    "description": "Optional image width in points/pixels."
                }
            },
            "required": ["prompt"]
        },
        "destructive": False
    },
    {
        "name": "word_batch_update_table_cells",
        "description": "Updates multiple cells across an existing Word table in a single atomic batch. Highly efficient for filling out forms, matrices, and project management templates.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "0-based index of the target Word table (default 0)."
                },
                "updates": {
                    "type": "array",
                    "description": "List of cell updates to apply.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "row": {"type": "integer", "description": "0-based row index."},
                            "col": {"type": "integer", "description": "0-based column index."},
                            "text": {"type": "string", "description": "The text content to write into the cell."},
                            "bold": {"type": "boolean", "description": "Optional bold formatting for the cell text."}
                        },
                        "required": ["row", "col", "text"]
                    }
                }
            },
            "required": ["updates"]
        },
        "destructive": False
    },
    {
        "name": "word_fill_form_fields",
        "description": "Smart form filler for templates. Searches for field labels (e.g., 'Codice', 'Titolo', 'Sponsor', 'Budget complessivo') across document tables and fills the corresponding target value cells automatically without needing exact row/column coordinates.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "Optional 0-based table index. If omitted, searches all tables in the document."
                },
                "fields": {
                    "type": "object",
                    "description": "Key-value pairs mapping label name to the desired value text (e.g. {'Codice': 'PRJ-2026', 'Titolo': 'AI Voice Assistant'}).",
                    "additionalProperties": {"type": "string"}
                }
            },
            "required": ["fields"]
        },
        "destructive": False
    }
]

_TOOL_MAP = {t["name"]: t for t in WORD_TOOLS}


def get_tool_definitions() -> List[Dict[str, Any]]:
    return WORD_TOOLS


def get_tool(name: str) -> Optional[Dict[str, Any]]:
    return _TOOL_MAP.get(name)


def is_destructive(name: str) -> bool:
    tool = _TOOL_MAP.get(name)
    return bool(tool and tool.get("destructive"))
