---
name: markdown-word-converter
description: Convert Markdown files to Word documents with support for Mermaid diagrams, custom templates, and table of contents generation. This skill should be used when users need to transform Markdown documentation into professionally formatted Word documents while preserving charts, diagrams, and document structure.
license: Apache 2.0
allowed-tools: ["Read", "Write", "Bash"]
---

# Markdown to Word Converter

Convert Markdown files to professional Word documents with support for Mermaid diagrams, custom templates, and automatic table of contents generation.

## Overview

This skill provides a cross-platform solution for converting Markdown documents to Word format (.docx) while preserving:
- Document structure and formatting
- Mermaid diagrams (converted to PNG images)
- Tables, lists, and code blocks
- Automatic table of contents
- Custom Word template styling

## When to Use This Skill

Use this skill when you need to:
- Convert technical documentation to Word format
- Transform Markdown reports into professional documents
- Export Markdown content with Mermaid diagrams to Word
- Apply consistent styling using custom templates
- Generate printable or editable documents from Markdown

## Prerequisites

Before using this skill, ensure the following dependencies are installed:

### Required Tools
1. **Pandoc** - Universal document converter
   - Windows: Download from https://pandoc.org/installing.html
   - macOS: `brew install pandoc`
   - Linux: `sudo apt install pandoc`

2. **Mermaid CLI** - For diagram conversion
   - Install with: `npm install -g @mermaid-js/mermaid-cli`

### Dependency Check
To check if dependencies are installed:
```bash
python scripts/install_dependencies.py
```

## Usage Instructions

### Basic Conversion
To convert a Markdown file to Word:
```bash
python scripts/convert.py input.md
```
This will create `input.docx` in the same directory.

### Specify Output File
To specify a custom output filename:
```bash
python scripts/convert.py input.md -o output.docx
```

### Without Template
To convert without using the default Word template:
```bash
python scripts/convert.py input.md --no-template
```

### Check Dependencies
To check if all dependencies are installed:
```bash
python scripts/convert.py --check-deps
```

## Workflow Process

1. **Validate Input**: Check if the input Markdown file exists
2. **Check Dependencies**: Verify that pandoc and mmdc are installed
3. **Process Mermaid Diagrams**: Convert any Mermaid diagrams to PNG images using mmdc
4. **Convert to Word**: Use pandoc with the custom template to generate the final Word document
5. **Generate Output**: Create the final .docx file with table of contents and proper formatting

## Supported Features

### Markdown Elements
- Headers and document hierarchy
- Text formatting (bold, italic, etc.)
- Lists (ordered and unordered)
- Links and images
- Code blocks and inline code
- Tables
- Blockquotes

### Mermaid Diagrams
- Flowcharts
- Sequence diagrams
- Gantt charts
- Class diagrams
- State diagrams
- Pie charts
- All standard Mermaid diagram types

### Word Document Features
- Automatic table of contents
- Custom styling from template
- Embedded images (converted diagrams)
- Proper document formatting
- Cross-platform compatibility

## File Structure

```
markdown-word-converter/
├── SKILL.md                    # This file - skill definition
├── scripts/
│   ├── convert.py             # Main conversion script
│   └── install_dependencies.py # Dependency checker and installer
├── assets/
│   └── template.docx          # Default Word template
└── references/
    └── usage_examples.md      # Additional usage examples
```

## Error Handling

The conversion process includes comprehensive error handling for:
- Missing input files
- Missing dependencies
- Invalid Markdown syntax
- Mermaid diagram conversion failures
- Template file issues
- Pandoc conversion errors

## Custom Templates

To use a custom Word template:
1. Replace `assets/template.docx` with your custom template
2. The template should include proper styles for:
   - Heading levels (H1, H2, etc.)
   - Normal text
   - Code blocks
   - Tables
   - Lists

## Troubleshooting

### Common Issues

**"pandoc not found"**
- Install pandoc from https://pandoc.org/installing.html
- Ensure pandoc is in your system PATH

**"mmdc not found"**
- Install Node.js if not already installed
- Run: `npm install -g @mermaid-js/mermaid-cli`

**"Template not found"**
- Ensure `assets/template.docx` exists
- Use `--no-template` flag to skip template usage

**Mermaid diagrams not converting**
- Check that mmdc is properly installed
- Verify Mermaid syntax in your Markdown
- Use `--no-template` flag as a fallback

### Getting Help

For additional support:
1. Check the dependency installation guide
2. Verify your Markdown syntax
3. Test with simple files first
4. Check script output for detailed error messages

## Performance Notes

- Large documents (>10MB) may take longer to process
- Complex Mermaid diagrams increase conversion time
- Temporary files are automatically cleaned up
- Memory usage scales with document size

## Best Practices

1. **Test with small files first** before converting large documents
2. **Validate Mermaid syntax** to ensure proper diagram conversion
3. **Use consistent heading levels** for better table of contents generation
4. **Check template compatibility** with your target Word version
5. **Verify all images and diagrams** in the final output