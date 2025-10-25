---
name: markdown-word-converter
description: Convert Markdown files to Word documents with Mermaid diagram support and custom templates. Simple two-step process: mmdc for diagrams, pandoc for conversion.
license: Apache 2.0
allowed-tools: ["Read", "Write", "Bash"]
---

# Markdown to Word Converter

Convert Markdown files to Word documents with Mermaid diagram support.

## Overview

A minimalist converter that:
- Converts Mermaid diagrams to PNG images using mmdc
- Transforms Markdown to Word using pandoc
- Supports custom Word templates
- Generates automatic table of contents

## When to Use

Use when you need to:
- Convert Markdown documentation to Word format
- Export documents with Mermaid diagrams
- Apply consistent styling with templates

## Prerequisites

Required tools:
1. **Pandoc** - Install from https://pandoc.org/installing.html
2. **Mermaid CLI** - Install with: `npm install -g @mermaid-js/mermaid-cli`

## Usage

### Basic Conversion
```bash
python scripts/convert.py input.md
```

### Custom Output File
```bash
python scripts/convert.py input.md output.docx
```

### Custom Template
```bash
python scripts/convert.py input.md --template template.docx
```

## Process

1. **Validate input** - Check if Markdown file exists
2. **Process diagrams** - Convert Mermaid charts to PNG using mmdc
3. **Convert to Word** - Use pandoc with template to generate .docx
4. **Generate output** - Create final document with table of contents

## File Structure

```
markdown-word-converter/
├── SKILL.md              # This file
├── scripts/
│   └── convert.py        # Main conversion script
└── assets/
    └── template.docx     # Default Word template
```

## Error Handling

The script handles:
- Missing input files
- Missing dependencies (pandoc, mmdc)
- Invalid Mermaid syntax
- Template file issues

## Template Support

The script automatically uses `assets/template.docx` if available, or you can specify a custom template with `--template`.