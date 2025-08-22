# File to Markdown Converter

A comprehensive Python script that converts Excel files (CSV/XLSX) and PDFs to markdown format optimized for LLM consumption.

## Installation

First, clone the repository:
```bash
git clone git@github.com:Dazarang/files_to_md.git
cd files_to_md
```

This project uses [uv](https://github.com/astral-sh/uv) for dependency management.

```bash
# 1) Ensure the pinned Python is available (reads .python-version if present)
uv python install

# 2) Create & populate the env from pyproject.toml / uv.lock
uv sync

# 3) (Optional) activate the virtual environment manually
# Note: uv handles this automatically, but you can activate if preferred
source .venv/bin/activate  # On macOS/Linux
# .venv\Scripts\activate   # On Windows

# 4) (Optional) run commands inside the env
uv run python -V
```

## Features

- **Excel/CSV Support**: Convert CSV, XLSX, and XLS files to structured markdown tables
- **PDF Processing**: Extract text and tables from PDFs with multiple extraction strategies
- **LLM-Optimized**: Output formatted specifically for easy AI/LLM consumption
- **Robust Error Handling**: Comprehensive edge case management and validation
- **Memory Efficient**: Handles large files with configurable memory limits
- **Rich CLI**: Command-line interface with progress indicators
- **Highly Configurable**: Multiple table formatting and output options
- **Batch Processing**: Convert multiple files at once with progress tracking
- **Directory Processing**: Recursively process entire directories
- **Duplicate Safety**: Automatic filename conflict resolution with suffix numbering

### 1. Single and Multiple File Processing

Convert a single file with straightforward command syntax:
```bash
uv run python file_to_md.py data.csv
```

Convert multiple files in a single command with comprehensive progress tracking:
```bash
uv run python file_to_md.py file1.csv file2.xlsx file3.pdf
```

### 2. Directory Processing
Process entire directories with support for filtering by file extensions:
```bash
# Process all supported files in a directory
uv run python file_to_md.py /data_folder/

# Process specific file types only
uv run python file_to_md.py /data_folder/ --extensions "xlsx,pdf"

# Recursive processing (including subdirectories)
uv run python file_to_md.py /project/ --recursive
```

### 3. Duplicate Filename Protection
Automatically handles filename conflicts by adding numbered suffixes:
- `data.md` → `data_1.md` → `data_2.md` (etc.)
- Prevents accidental overwriting of existing files

## Usage

### Basic Usage

```bash
# Convert a single file
uv run python file_to_md.py data.csv

# Convert multiple files
uv run python file_to_md.py file1.csv file2.xlsx file3.pdf

# Convert all supported files in a directory
uv run python file_to_md.py /path/to/folder/

# Convert recursively (including subdirectories)
uv run python file_to_md.py /path/to/folder/ --recursive

# Mixed: files and directories
uv run python file_to_md.py data.csv /path/to/folder/ report.xlsx
```

### Advanced Options

```bash
# Single file with custom output
uv run python file_to_md.py data.csv -o custom_name.md

# Multiple files with output directory
uv run python file_to_md.py file1.csv file2.xlsx --output-dir /path/to/output/

# Directory processing with specific extensions
uv run python file_to_md.py /folder/ --extensions "xlsx,pdf"

# Recursive directory processing
uv run python file_to_md.py /folder/ --recursive --extensions "csv,xlsx"


# Limit processing for large files
uv run python file_to_md.py large_file.xlsx --max-rows 1000 --max-cols 20

# Process specific Excel sheets
uv run python file_to_md.py workbook.xlsx --sheets "Sheet1,Summary,Data"

# Customize table formatting
uv run python file_to_md.py data.csv --table-format grid    # Grid format
uv run python file_to_md.py data.csv --table-format simple  # Simple format

# Skip metadata generation
uv run python file_to_md.py data.csv --no-metadata

# Memory management
uv run python file_to_md.py large_file.pdf --memory-limit 2048  # 2GB limit

# Verbose output for debugging
uv run python file_to_md.py problematic_file.pdf --verbose
```

## Supported File Types

### Excel/CSV Files
- **CSV**: Comma-separated values with automatic encoding detection
- **XLSX**: Modern Excel format with multi-sheet support
- **XLS**: Legacy Excel format

### PDF Files
- **Text extraction**: Multiple strategies for robust text extraction
- **Table detection**: Automatic table identification and formatting
- **Layout preservation**: Maintains document structure when possible
- **Multi-page support**: Processes entire documents with progress tracking

## Output Features

### Optimized for LLMs
- Clean markdown formatting for easy parsing
- Structured tables with clear headers
- Metadata sections for context
- Summary statistics for numerical data
- Consistent formatting across all file types

## Examples

### Batch Processing Multiple Files
```bash
# Convert multiple specific files
uv run python file_to_md.py report1.xlsx data.csv analysis.pdf \\
  --output-dir converted_files/ \\
  --table-format github \\
  --verbose
```

### Directory Processing with Filtering
```bash
# Process only Excel and PDF files in a directory
uv run python file_to_md.py /data_folder/ \\
  --extensions "xlsx,pdf" \\
  --output-dir markdown_output/ \\
  --max-rows 1000
```

### Recursive Directory Processing
```bash
# Process all supported files in directory tree
uv run python file_to_md.py /project_docs/ \\
  --recursive \\
\
  --table-format grid
```

### CSV with Custom Options
```bash
uv run python file_to_md.py employee_data.csv \\
  --max-rows 500 \\
  --table-format grid \\
\
  --verbose
```

### Excel Multi-Sheet Processing
```bash
uv run python file_to_md.py financial_report.xlsx \\
  --sheets "Summary,Q1,Q2,Q3,Q4" \\
  --output-dir reports/ \\
  --table-format github
```

## Output Examples

### Batch Processing Results
When processing multiple files, you'll see comprehensive statistics:

```
Successfully Converted Files
+----------------------------------------------------------+
| Input File       | Output File      | Size               |
|------------------+------------------+--------------------|
| sample_data.csv  | sample_data.md   | 419 -> 1,101 bytes |
| sample_data.xlsx | sample_data.md   | 5,756 -> 998 bytes |
+----------------------------------------------------------+

Overall Statistics
+-----------------------------------------+
| Metric                    | Value       |
|---------------------------+-------------|
| Total files processed     | 2           |
| Successfully converted    | 2           |
| Failed conversions        | 0           |
| Total original size       | 6,175 bytes |
| Total output size         | 2,099 bytes |
| Overall compression ratio | 33.99%      |
+-----------------------------------------+
```

### Duplicate Filename Handling
```bash
# First conversion
uv run python file_to_md.py data.csv
# Output: data.md

# Second conversion of same file
uv run python file_to_md.py data.csv
# Output: data_1.md (automatically numbered)

# Third conversion
uv run python file_to_md.py data.csv  
# Output: data_2.md
```

## Edge Cases Handled

### File Validation
- Empty or corrupted files
- Non-existent file paths
- Insufficient permissions
- Memory constraints
- Duplicate filename conflicts (automatic resolution)

### Data Processing
- Mixed data types in columns
- Unicode and special characters
- Very large datasets (chunking)
- Password-protected files (graceful failure)
- Non-standard CSV delimiters
- Complex Excel formulas and formatting

### Batch Processing
- Progress tracking for large batches
- Individual file error isolation (one failure doesn't stop batch)
- Comprehensive reporting of successes and failures
- Memory management across multiple files

## Error Handling

The script provides comprehensive error handling with:
- **Graceful degradation**: Multiple extraction methods for PDFs
- **Detailed logging**: Rich console output with progress tracking
- **Memory monitoring**: Prevents system overload
- **Encoding detection**: Handles various text encodings
- **Validation checks**: Pre-flight validation of inputs
- **Batch resilience**: Individual file failures don't stop batch processing
- **Duplicate protection**: Automatic filename conflict resolution

## Performance Considerations

- **Memory efficient**: Processes large files without loading entirely into memory
- **Configurable limits**: Set maximum rows/columns for processing
- **Batch processing**: Handles large datasets in chunks
- **Progress indicators**: Real-time feedback for long operations
- **Parallel processing**: Future enhancement for concurrent file processing

## Troubleshooting

### Common Issues

1. **Unicode errors**: Use `--verbose` to see encoding detection details
2. **Memory issues**: Reduce `--memory-limit` or use `--max-rows`/`--max-cols`
3. **PDF extraction fails**: Try different strategies or check if PDF is text-based
4. **Excel password protection**: Remove protection before conversion
5. **Large file timeouts**: Process in smaller chunks using row/column limits
6. **Filename conflicts**: Script automatically handles with numbered suffixes
7. **Directory permissions**: Ensure read access to source and write access to output directories

### Command Line Help

```bash
uv run python file_to_md.py --help
```

## Technical Architecture

The script uses a modular architecture with:

- **FileConverter**: Main orchestration class
- **Multiple PDF strategies**: pdfplumber → PyMuPDF → PyPDF2 fallback chain
- **Pandas integration**: Robust data processing and table formatting
- **Rich UI**: Professional CLI with progress indicators
- **Comprehensive logging**: Detailed operational feedback
- **Batch processing engine**: Efficient handling of multiple files
- **Directory traversal**: Recursive and non-recursive file discovery
- **Duplicate resolution**: Automatic filename conflict handling

## Dependencies

Core dependencies managed via uv:
- `pandas`: Data processing and Excel/CSV reading
- `openpyxl`, `xlrd`: Excel file support
- `PyPDF2`, `pdfplumber`, `pymupdf`: PDF processing
- `click`: CLI framework
- `rich`: Enhanced console output
- `tabulate`: Table formatting
- Plus various utilities for encoding, validation, and performance monitoring

## License

This project is provided as-is for educational and productivity purposes.