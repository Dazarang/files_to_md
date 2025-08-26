#!/usr/bin/env python3
"""
File to Markdown Converter
===========================

A comprehensive Python script that converts Excel files (CSV/XLSX) and PDFs 
to markdown format for LLM consumption.

Features:
- Excel/CSV to structured markdown tables
- PDF to markdown with table preservation
- Robust error handling and edge case management
- Memory-efficient processing for large files
- Rich CLI interface with progress indicators
- Configurable output formatting
"""

import os
import sys
import logging
import traceback
from pathlib import Path
from typing import Optional, Dict, Any, List, Union
from dataclasses import dataclass
from enum import Enum
import glob

import click
import pandas as pd
import fitz  # PyMuPDF
import pdfplumber
from PyPDF2 import PdfReader
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeRemainingColumn
from rich.table import Table
from rich.panel import Panel
from rich.logging import RichHandler
from tabulate import tabulate
import chardet
import magic
import validators
import psutil
from pdf_ocr import PDFOCRProcessor, OCRConfig, OCRQuality, create_ocr_processor


class FileType(Enum):
    """Supported file types for conversion."""
    CSV = "csv"
    XLSX = "xlsx"
    XLS = "xls" 
    PDF = "pdf"
    UNKNOWN = "unknown"


@dataclass
class ConversionConfig:
    """Configuration settings for file conversion."""
    # Output settings
    output_dir: Optional[Path] = None
    output_filename: Optional[str] = None
    
    # Excel/CSV settings
    max_rows: Optional[int] = None
    max_cols: Optional[int] = None
    include_index: bool = False
    sheet_names: Optional[List[str]] = None
    
    # PDF settings
    extract_images: bool = False
    preserve_layout: bool = True
    table_detection: bool = True
    ocr_enabled: bool = False
    ocr_quality: str = "balanced"
    ocr_language: str = "eng"
    ocr_force: bool = False
    
    # Markdown formatting
    table_format: str = "github"  # github, grid, simple, etc.
    add_metadata: bool = True
    chunk_large_tables: bool = True
    max_table_width: int = 120
    
    # Performance settings
    memory_limit_mb: int = 1024
    batch_size: int = 1000
    verbose: bool = False


class FileConverter:
    """Main converter class that orchestrates file conversion to markdown."""
    
    def __init__(self, config: ConversionConfig):
        self.config = config
        self.console = Console()
        self.logger = self._setup_logging()
        
    def _setup_logging(self) -> logging.Logger:
        """Configure logging with rich formatting."""
        logger = logging.getLogger("file_converter")
        logger.setLevel(logging.DEBUG if self.config.verbose else logging.INFO)
        
        if not logger.handlers:
            handler = RichHandler(console=self.console, show_time=True, show_path=True)
            formatter = logging.Formatter("%(message)s")
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            
        return logger
    
    def detect_file_type(self, file_path: Path) -> FileType:
        """Detect file type using multiple methods."""
        try:
            # First try by extension
            suffix = file_path.suffix.lower()
            type_map = {
                '.csv': FileType.CSV,
                '.xlsx': FileType.XLSX,
                '.xls': FileType.XLS,
                '.pdf': FileType.PDF
            }
            
            if suffix in type_map:
                detected_type = type_map[suffix]
                self.logger.debug(f"Detected type by extension: {detected_type.value}")
                return detected_type
            
            # Fallback to magic number detection
            try:
                file_type = magic.from_file(str(file_path), mime=True)
                self.logger.debug(f"MIME type detected: {file_type}")
                
                mime_map = {
                    'text/csv': FileType.CSV,
                    'application/csv': FileType.CSV,
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': FileType.XLSX,
                    'application/vnd.ms-excel': FileType.XLS,
                    'application/pdf': FileType.PDF
                }
                
                return mime_map.get(file_type, FileType.UNKNOWN)
            except Exception as e:
                self.logger.warning(f"Magic number detection failed: {e}")
                return FileType.UNKNOWN
                
        except Exception as e:
            self.logger.error(f"File type detection failed: {e}")
            return FileType.UNKNOWN
    
    def validate_file(self, file_path: Path) -> bool:
        """Validate input file exists and is accessible."""
        try:
            if not file_path.exists():
                self.logger.error(f"File does not exist: {file_path}")
                return False
            
            if not file_path.is_file():
                self.logger.error(f"Path is not a file: {file_path}")
                return False
            
            if file_path.stat().st_size == 0:
                self.logger.error(f"File is empty: {file_path}")
                return False
            
            # Check file size vs memory limit
            file_size_mb = file_path.stat().st_size / (1024 * 1024)
            available_memory_mb = psutil.virtual_memory().available / (1024 * 1024)
            
            if file_size_mb > self.config.memory_limit_mb:
                self.logger.warning(f"File size ({file_size_mb:.1f}MB) exceeds memory limit ({self.config.memory_limit_mb}MB)")
                
            if file_size_mb > available_memory_mb * 0.5:  # Use max 50% of available memory
                self.logger.warning(f"File size ({file_size_mb:.1f}MB) may cause memory issues")
            
            return True
            
        except Exception as e:
            self.logger.error(f"File validation failed: {e}")
            return False
    
    def convert_file(self, file_path: Path) -> Optional[str]:
        """Main conversion method that routes to appropriate converter."""
        try:
            # Validate input
            if not self.validate_file(file_path):
                return None
            
            # Detect file type
            file_type = self.detect_file_type(file_path)
            if file_type == FileType.UNKNOWN:
                self.logger.error(f"Unsupported file type: {file_path}")
                return None
            
            self.logger.info(f"Converting {file_type.value.upper()} file: {file_path.name}")
            
            # Route to appropriate converter
            converters = {
                FileType.CSV: self._convert_csv,
                FileType.XLSX: self._convert_excel,
                FileType.XLS: self._convert_excel,
                FileType.PDF: self._convert_pdf
            }
            
            converter_func = converters[file_type]
            markdown_content = converter_func(file_path)
            
            if markdown_content:
                # Add metadata if requested
                if self.config.add_metadata:
                    metadata = self._generate_metadata(file_path, file_type)
                    markdown_content = metadata + "\n\n" + markdown_content
                
                self.logger.info(f"Successfully converted {file_path.name}")
                return markdown_content
            else:
                self.logger.error(f"Conversion failed for {file_path.name}")
                return None
                
        except Exception as e:
            self.logger.error(f"Conversion error: {e}")
            if self.config.verbose:
                self.logger.error(traceback.format_exc())
            return None
    
    def _generate_metadata(self, file_path: Path, file_type: FileType) -> str:
        """Generate metadata header for the markdown file."""
        try:
            stat = file_path.stat()
            file_size = stat.st_size
            mod_time = stat.st_mtime
            
            metadata = f"""---
# File Conversion Metadata
**Source File:** {file_path.name}  
**File Type:** {file_type.value.upper()}  
**File Size:** {file_size:,} bytes ({file_size / (1024*1024):.2f} MB)  
**Last Modified:** {pd.to_datetime(mod_time, unit='s').strftime('%Y-%m-%d %H:%M:%S')}  
**Converted On:** {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}  
---"""
            return metadata
        except Exception as e:
            self.logger.warning(f"Could not generate metadata: {e}")
            return ""
    
    def _convert_csv(self, file_path: Path) -> Optional[str]:
        """Convert CSV file to markdown."""
        try:
            # Detect encoding
            with open(file_path, 'rb') as f:
                raw_data = f.read(10000)  # Read first 10KB for encoding detection
                encoding_result = chardet.detect(raw_data)
                encoding = encoding_result.get('encoding', 'utf-8')
            
            self.logger.debug(f"Detected encoding: {encoding}")
            
            # Read CSV with pandas
            read_kwargs = {
                'encoding': encoding,
                'encoding_errors': 'replace',  # Handle encoding issues gracefully
                'low_memory': False,
            }
            
            # Apply limits if specified
            if self.config.max_rows:
                read_kwargs['nrows'] = self.config.max_rows
                
            df = pd.read_csv(file_path, **read_kwargs)
            
            # Apply column limit
            if self.config.max_cols:
                df = df.iloc[:, :self.config.max_cols]
            
            return self._dataframe_to_markdown(df, file_path.stem)
            
        except Exception as e:
            self.logger.error(f"CSV conversion failed: {e}")
            return None
    
    def _convert_excel(self, file_path: Path) -> Optional[str]:
        """Convert Excel file to markdown."""
        try:
            markdown_parts = []
            
            # Read Excel file
            if file_path.suffix.lower() == '.xlsx':
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            else:  # .xls
                excel_file = pd.ExcelFile(file_path, engine='xlrd')
            
            sheet_names = self.config.sheet_names or excel_file.sheet_names
            
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TimeRemainingColumn(),
                console=self.console
            ) as progress:
                task = progress.add_task("Converting sheets...", total=len(sheet_names))
                
                for sheet_name in sheet_names:
                    try:
                        # Read sheet
                        read_kwargs = {}
                        if self.config.max_rows:
                            read_kwargs['nrows'] = self.config.max_rows
                        
                        df = pd.read_excel(excel_file, sheet_name=sheet_name, **read_kwargs)
                        
                        # Apply column limit
                        if self.config.max_cols:
                            df = df.iloc[:, :self.config.max_cols]
                        
                        if not df.empty:
                            sheet_markdown = self._dataframe_to_markdown(df, sheet_name)
                            if len(excel_file.sheet_names) > 1:
                                markdown_parts.append(f"## Sheet: {sheet_name}\n\n{sheet_markdown}")
                            else:
                                markdown_parts.append(sheet_markdown)
                        else:
                            markdown_parts.append(f"## Sheet: {sheet_name}\n\n*Sheet is empty*\n")
                            
                        progress.advance(task)
                        
                    except Exception as e:
                        self.logger.warning(f"Failed to process sheet '{sheet_name}': {e}")
                        markdown_parts.append(f"## Sheet: {sheet_name}\n\n*Error processing sheet: {str(e)}*\n")
                        progress.advance(task)
            
            return "\n\n".join(markdown_parts) if markdown_parts else None
            
        except Exception as e:
            self.logger.error(f"Excel conversion failed: {e}")
            return None
    
    def _convert_pdf(self, file_path: Path) -> Optional[str]:
        """Convert PDF file to markdown using multiple extraction methods."""
        try:
            # Apply OCR if enabled
            if self.config.ocr_enabled:
                self.logger.info("OCR is enabled, checking if PDF needs OCR processing...")
                ocr_processor = create_ocr_processor(
                    quality=self.config.ocr_quality,
                    language=self.config.ocr_language,
                    force_ocr=self.config.ocr_force,
                    preserve_layout=self.config.preserve_layout
                )
                
                # Check if PDF needs OCR
                if not ocr_processor.check_pdf_has_text(file_path) or self.config.ocr_force:
                    self.logger.info("Running OCR on PDF...")
                    ocr_output = file_path.parent / f"{file_path.stem}_ocr{file_path.suffix}"
                    processed_path = ocr_processor.process_pdf(
                        input_path=file_path,
                        output_path=ocr_output,
                        check_existing_text=not self.config.ocr_force
                    )
                    
                    if processed_path and processed_path != file_path:
                        self.logger.info(f"OCR completed, using processed file: {processed_path}")
                        file_path = processed_path
                    elif processed_path == file_path:
                        self.logger.info("PDF already has text, using original file")
                    else:
                        self.logger.warning("OCR processing failed")
                        # Try alternative image-based OCR
                        self.logger.info("Attempting alternative image-based OCR...")
                        ocr_text = ocr_processor.ocr_pdf_via_images(
                            file_path, 
                            language=self.config.ocr_language
                        )
                        if ocr_text:
                            # Save the OCR text to a temporary markdown file
                            temp_md = file_path.parent / f"{file_path.stem}_ocr_text.md"
                            with open(temp_md, 'w', encoding='utf-8') as f:
                                f.write(f"# OCR Extracted Text from {file_path.name}\n\n")
                                f.write(ocr_text)
                            self.logger.info(f"Image-based OCR succeeded, saved to: {temp_md}")
                            # Return the OCR text directly instead of continuing with regular extraction
                            return ocr_text
                        else:
                            self.logger.warning("Image-based OCR also failed, continuing with original file")
                            if not self.config.verbose:
                                self.logger.info("Run with --verbose for detailed error information")
                else:
                    self.logger.info("PDF already has searchable text, skipping OCR")
            
            markdown_parts = []
            
            # Method 1: Try pdfplumber for structured extraction (best for tables)
            try:
                markdown_content = self._extract_pdf_with_pdfplumber(file_path)
                if markdown_content and len(markdown_content.strip()) > 100:  # Has substantial content
                    return markdown_content
            except Exception as e:
                self.logger.warning(f"PDFplumber extraction failed: {e}")
            
            # Method 2: Fallback to PyMuPDF for general text extraction
            try:
                markdown_content = self._extract_pdf_with_pymupdf(file_path)
                if markdown_content and len(markdown_content.strip()) > 50:
                    return markdown_content
            except Exception as e:
                self.logger.warning(f"PyMuPDF extraction failed: {e}")
            
            # Method 3: Last resort - PyPDF2 for basic text extraction
            try:
                markdown_content = self._extract_pdf_with_pypdf2(file_path)
                if markdown_content:
                    return markdown_content
            except Exception as e:
                self.logger.warning(f"PyPDF2 extraction failed: {e}")
            
            self.logger.error("All PDF extraction methods failed")
            return None
            
        except Exception as e:
            self.logger.error(f"PDF conversion failed: {e}")
            return None
    
    def _extract_pdf_with_pdfplumber(self, file_path: Path) -> Optional[str]:
        """Extract PDF content using pdfplumber (best for tables and layout)."""
        markdown_parts = []
        
        with pdfplumber.open(file_path) as pdf:
            total_pages = len(pdf.pages)
            
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TimeRemainingColumn(),
                console=self.console
            ) as progress:
                task = progress.add_task(f"Processing PDF pages...", total=total_pages)
                
                for i, page in enumerate(pdf.pages, 1):
                    try:
                        page_content = []
                        
                        # Extract tables first (higher priority)
                        if self.config.table_detection:
                            tables = page.extract_tables()
                            for j, table in enumerate(tables):
                                if table and len(table) > 1:  # Has header + data
                                    # Convert table to DataFrame for better formatting
                                    df = pd.DataFrame(table[1:], columns=table[0])
                                    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
                                    
                                    if not df.empty:
                                        table_md = self._dataframe_to_markdown(df, f"Page {i} Table {j+1}")
                                        page_content.append(f"### Page {i} - Table {j+1}\n\n{table_md}")
                        
                        # Extract remaining text (excluding table areas if possible)
                        text = page.extract_text()
                        if text:
                            text = text.strip()
                            if text and len(text) > 20:  # Avoid tiny fragments
                                # Clean up text formatting
                                text = self._clean_extracted_text(text)
                                page_content.append(f"### Page {i} - Text Content\n\n{text}")
                        
                        if page_content:
                            markdown_parts.append(f"## Page {i}\n\n" + "\n\n".join(page_content))
                        
                        progress.advance(task)
                        
                    except Exception as e:
                        self.logger.warning(f"Failed to process page {i}: {e}")
                        progress.advance(task)
        
        return "\n\n---\n\n".join(markdown_parts) if markdown_parts else None
    
    def _extract_pdf_with_pymupdf(self, file_path: Path) -> Optional[str]:
        """Extract PDF content using PyMuPDF (good balance of features)."""
        markdown_parts = []
        
        doc = fitz.open(file_path)
        total_pages = len(doc)
        
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TimeRemainingColumn(),
            console=self.console
        ) as progress:
            task = progress.add_task(f"Processing PDF pages...", total=total_pages)
            
            for page_num in range(total_pages):
                try:
                    page = doc.load_page(page_num)
                    
                    # Extract text
                    text = page.get_text()
                    if text:
                        text = text.strip()
                        if text and len(text) > 20:
                            text = self._clean_extracted_text(text)
                            markdown_parts.append(f"## Page {page_num + 1}\n\n{text}")
                    
                    progress.advance(task)
                    
                except Exception as e:
                    self.logger.warning(f"Failed to process page {page_num + 1}: {e}")
                    progress.advance(task)
        
        doc.close()
        return "\n\n---\n\n".join(markdown_parts) if markdown_parts else None
    
    def _extract_pdf_with_pypdf2(self, file_path: Path) -> Optional[str]:
        """Extract PDF content using PyPDF2 (basic text extraction)."""
        markdown_parts = []
        
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                total_pages = len(pdf_reader.pages)
                
                for i, page in enumerate(pdf_reader.pages):
                    try:
                        text = page.extract_text()
                        if text:
                            text = text.strip()
                            if text and len(text) > 20:
                                text = self._clean_extracted_text(text)
                                markdown_parts.append(f"## Page {i + 1}\n\n{text}")
                    except Exception as e:
                        self.logger.warning(f"Failed to extract text from page {i + 1}: {e}")
                        
        except Exception as e:
            self.logger.error(f"PyPDF2 extraction failed: {e}")
            return None
        
        return "\n\n---\n\n".join(markdown_parts) if markdown_parts else None
    
    def _clean_extracted_text(self, text: str) -> str:
        """Clean and format extracted text for better markdown presentation."""
        try:
            lines = text.split('\n')
            cleaned_lines = []
            
            for line in lines:
                line = line.strip()
                if line:  # Skip empty lines
                    # Remove excessive whitespace
                    line = ' '.join(line.split())
                    cleaned_lines.append(line)
            
            # Join lines with proper spacing
            cleaned_text = '\n\n'.join(cleaned_lines)
            
            # Remove excessive line breaks
            while '\n\n\n' in cleaned_text:
                cleaned_text = cleaned_text.replace('\n\n\n', '\n\n')
            
            return cleaned_text
            
        except Exception as e:
            self.logger.warning(f"Text cleaning failed: {e}")
            return text  # Return original if cleaning fails
    
    def _dataframe_to_markdown(self, df: pd.DataFrame, title: str = "") -> str:
        """Convert pandas DataFrame to well-formatted markdown."""
        try:
            if df.empty:
                return f"### {title}\n\n*No data available*\n" if title else "*No data available*\n"
            
            # Clean DataFrame
            df = df.copy()
            
            # Store original dtypes before conversion for statistics
            original_df = df.copy()
            
            # Handle NaN values
            df = df.fillna('')
            
            # Convert all columns to string to avoid formatting issues
            for col in df.columns:
                df[col] = df[col].astype(str)
            
            # Truncate very wide tables
            if len(df.columns) > 20:
                self.logger.warning(f"Table has {len(df.columns)} columns, truncating to first 20")
                df = df.iloc[:, :20]
                df['...'] = '...'  # Indicate truncation
            
            # Handle large tables by chunking if requested
            markdown_parts = []
            if title:
                markdown_parts.append(f"### {title}\n")
            
            if self.config.chunk_large_tables and len(df) > 100:
                # Chunk large tables for better readability
                chunk_size = 50
                for i in range(0, len(df), chunk_size):
                    chunk = df.iloc[i:i+chunk_size]
                    chunk_title = f"Rows {i+1}-{min(i+chunk_size, len(df))}"
                    
                    table_md = tabulate(
                        chunk, 
                        headers='keys', 
                        tablefmt=self.config.table_format,
                        showindex=self.config.include_index,
                        maxcolwidths=self.config.max_table_width // len(chunk.columns)
                    )
                    
                    markdown_parts.append(f"#### {chunk_title}\n\n{table_md}\n")
            else:
                # Regular table conversion
                table_md = tabulate(
                    df, 
                    headers='keys', 
                    tablefmt=self.config.table_format,
                    showindex=self.config.include_index,
                    maxcolwidths=self.config.max_table_width // len(df.columns) if len(df.columns) > 0 else None
                )
                markdown_parts.append(f"{table_md}\n")
            
            # Add summary statistics for numerical data (using original DataFrame before string conversion)
            numeric_cols = original_df.select_dtypes(include=['float64', 'int64']).columns
            if len(numeric_cols) > 0:
                try:
                    df_numeric = original_df[numeric_cols].apply(pd.to_numeric, errors='coerce')
                    summary = df_numeric.describe()
                    if not summary.empty:
                        summary_md = tabulate(
                            summary, 
                            headers='keys', 
                            tablefmt=self.config.table_format,
                            showindex=True
                        )
                        markdown_parts.append(f"#### Summary Statistics\n\n{summary_md}\n")
                except Exception as e:
                    self.logger.debug(f"Could not generate summary statistics: {e}")
            
            return "\n".join(markdown_parts)
            
        except Exception as e:
            self.logger.error(f"DataFrame to markdown conversion failed: {e}")
            return f"### {title}\n\n*Error converting data to markdown: {str(e)}*\n" if title else f"*Error converting data to markdown: {str(e)}*\n"
    
    def _get_unique_output_path(self, base_path: Path) -> Path:
        """Generate a unique output path by adding suffix numbers if file exists."""
        if not base_path.exists():
            return base_path
        
        # Extract parts
        stem = base_path.stem
        suffix = base_path.suffix
        parent = base_path.parent
        
        # Try numbered suffixes
        counter = 1
        while True:
            new_name = f"{stem}_{counter}{suffix}"
            new_path = parent / new_name
            if not new_path.exists():
                self.logger.info(f"File exists, using unique name: {new_path.name}")
                return new_path
            counter += 1
            
            # Safety check to avoid infinite loop
            if counter > 9999:
                raise ValueError("Could not generate unique filename after 9999 attempts")
    
    def save_markdown(self, content: str, original_file: Path) -> Optional[Path]:
        """Save markdown content to file with duplicate handling."""
        try:
            # Determine output file path
            if self.config.output_filename:
                output_file = Path(self.config.output_filename)
            else:
                output_file = original_file.with_suffix('.md')
            
            if self.config.output_dir:
                output_file = self.config.output_dir / output_file.name
            
            # Ensure output directory exists
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Handle duplicate filenames
            output_file = self._get_unique_output_path(output_file)
            
            # Write content
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            self.logger.info(f"Markdown saved to: {output_file}")
            return output_file
            
        except Exception as e:
            self.logger.error(f"Failed to save markdown: {e}")
            return None
    
    def find_files_in_directory(self, directory: Path, extensions: List[str] = None) -> List[Path]:
        """Find all supported files in a directory."""
        if extensions is None:
            extensions = ['.csv', '.xlsx', '.xls', '.pdf']
        
        files = []
        try:
            for ext in extensions:
                pattern = f"*{ext}"
                found_files = list(directory.glob(pattern))
                files.extend(found_files)
                
                # Also search case-insensitive on Windows
                if os.name == 'nt':
                    pattern_upper = f"*{ext.upper()}"
                    found_files_upper = list(directory.glob(pattern_upper))
                    files.extend(found_files_upper)
            
            # Remove duplicates and sort
            files = sorted(list(set(files)))
            self.logger.info(f"Found {len(files)} supported files in {directory}")
            return files
            
        except Exception as e:
            self.logger.error(f"Error scanning directory {directory}: {e}")
            return []
    
    def convert_multiple_files(self, file_paths: List[Path]) -> Dict[Path, Optional[Path]]:
        """Convert multiple files and return results mapping."""
        results = {}
        
        if not file_paths:
            self.logger.warning("No files to convert")
            return results
        
        total_files = len(file_paths)
        self.console.print(f"\n[bold blue]Processing {total_files} files...[/bold blue]")
        
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            TimeRemainingColumn(),
            console=self.console
        ) as progress:
            task = progress.add_task("Converting files...", total=total_files)
            
            success_count = 0
            failure_count = 0
            
            for i, file_path in enumerate(file_paths, 1):
                try:
                    progress.update(task, description=f"Converting {file_path.name} ({i}/{total_files})")
                    
                    # Convert individual file
                    markdown_content = self.convert_file(file_path)
                    
                    if markdown_content:
                        # Save markdown
                        output_path = self.save_markdown(markdown_content, file_path)
                        if output_path:
                            results[file_path] = output_path
                            success_count += 1
                        else:
                            results[file_path] = None
                            failure_count += 1
                    else:
                        results[file_path] = None
                        failure_count += 1
                    
                    progress.advance(task)
                    
                except Exception as e:
                    self.logger.error(f"Failed to convert {file_path.name}: {e}")
                    results[file_path] = None
                    failure_count += 1
                    progress.advance(task)
        
        # Summary
        self.console.print(f"\n[bold green]Conversion completed![/bold green]")
        self.console.print(f"+ Successfully converted: {success_count} files")
        if failure_count > 0:
            self.console.print(f"- Failed conversions: {failure_count} files")
        
        return results


# CLI Interface
@click.command()
@click.argument('input_path', type=click.Path(exists=True, path_type=Path), nargs=-1, required=True)
@click.option('-o', '--output', type=click.Path(path_type=Path), help='Output file path (single file only)')
@click.option('--output-dir', type=click.Path(path_type=Path), help='Output directory')
@click.option('--max-rows', type=int, help='Maximum rows to process')
@click.option('--max-cols', type=int, help='Maximum columns to process')
@click.option('--sheets', help='Comma-separated list of sheet names to process (Excel only)')
@click.option('--table-format', default='github', 
              help='Markdown table format (github, grid, simple, etc.)')
@click.option('--no-metadata', is_flag=True, help='Skip metadata generation')
@click.option('--memory-limit', type=int, default=1024, help='Memory limit in MB')
@click.option('--extensions', default='csv,xlsx,xls,pdf', 
              help='File extensions to process when scanning directories (comma-separated)')
@click.option('--recursive', '-r', is_flag=True, help='Process directories recursively')
@click.option('--ocr', is_flag=True, help='Enable OCR for unsearchable PDFs')
@click.option('--ocr-quality', type=click.Choice(['fast', 'balanced', 'high']), default='balanced',
              help='OCR quality level')
@click.option('--ocr-language', default='eng', 
              help='OCR language(s) - e.g., "eng", "deu", "eng+fra"')
@click.option('--ocr-force', is_flag=True, 
              help='Force OCR even if PDF already has text')
@click.option('--verbose', '-v', is_flag=True, help='Enable verbose logging')
def main(input_path, output, output_dir, max_rows, max_cols, sheets, 
         table_format, no_metadata, memory_limit, extensions, recursive, 
         ocr, ocr_quality, ocr_language, ocr_force, verbose):
    """Convert Excel/CSV/PDF files to markdown format for LLM consumption.
    
    INPUT_PATH can be:\n
    - Single file: file_to_md.py data.csv
    
    - Multiple files: file_to_md.py file1.csv file2.xlsx file3.pdf
    
    - Directory: file_to_md.py /path/to/folder/ (processes all supported files)
    
    - Mixed: file_to_md.py file1.csv /path/to/folder/ file2.pdf
    """
    
    # Process extensions
    supported_extensions = [f'.{ext.strip().lower().lstrip(".")}' for ext in extensions.split(',')]
    
    # Create configuration  
    config = ConversionConfig(
        output_dir=output_dir,
        output_filename=output.name if output else None,
        max_rows=max_rows,
        max_cols=max_cols,
        sheet_names=sheets.split(',') if sheets else None,
        table_format=table_format,
        add_metadata=not no_metadata,
        memory_limit_mb=memory_limit,
        ocr_enabled=ocr,
        ocr_quality=ocr_quality,
        ocr_language=ocr_language,
        ocr_force=ocr_force,
        verbose=verbose
    )
    
    console = Console()
    
    try:
        # Create converter
        converter = FileConverter(config)
        
        # Collect all files to process
        all_files = []
        directories_found = []
        
        for path in input_path:
            if path.is_file():
                all_files.append(path)
            elif path.is_dir():
                directories_found.append(path)
                if recursive:
                    # Recursive search
                    for ext in supported_extensions:
                        pattern = f"**/*{ext}"
                        found_files = list(path.rglob(pattern))
                        all_files.extend(found_files)
                        
                        # Case-insensitive search on Windows
                        if os.name == 'nt':
                            pattern_upper = f"**/*{ext.upper()}"
                            found_files_upper = list(path.rglob(pattern_upper))
                            all_files.extend(found_files_upper)
                else:
                    # Non-recursive search
                    found_files = converter.find_files_in_directory(path, supported_extensions)
                    all_files.extend(found_files)
        
        # Remove duplicates and sort
        all_files = sorted(list(set(all_files)))
        
        # Validation
        if not all_files:
            console.print("[bold red]ERROR: No supported files found to convert[/bold red]")
            if directories_found:
                console.print(f"Searched directories: {[str(d) for d in directories_found]}")
                console.print(f"Supported extensions: {supported_extensions}")
            sys.exit(1)
        
        # Check for single file with output filename
        if len(all_files) > 1 and output:
            console.print("[bold red]ERROR: Cannot specify output filename (-o) when processing multiple files[/bold red]")
            console.print("Use --output-dir instead to specify output directory")
            sys.exit(1)
        
        # Display banner
        input_summary = f"{len(all_files)} file(s)" if len(all_files) > 1 else str(all_files[0].name)
        console.print(Panel.fit(
            "[bold blue]File to Markdown Converter[/bold blue]\n"
            f"Input: [green]{input_summary}[/green]\n"
            f"Extensions: [cyan]{', '.join(supported_extensions)}[/cyan]",
            border_style="blue"
        ))
        
        # Process files
        if len(all_files) == 1:
            # Single file processing
            input_file = all_files[0]
            markdown_content = converter.convert_file(input_file)
            
            if markdown_content:
                # Save markdown
                output_path = converter.save_markdown(markdown_content, input_file)
                
                if output_path:
                    # Display success
                    console.print(f"\n[bold green]SUCCESS: Conversion completed successfully![/bold green]")
                    console.print(f"Output saved to: [cyan]{output_path}[/cyan]")
                    
                    # Display file stats
                    stats_table = Table(title="Conversion Statistics")
                    stats_table.add_column("Metric", style="cyan")
                    stats_table.add_column("Value", style="green")
                    
                    original_size = input_file.stat().st_size
                    output_size = output_path.stat().st_size
                    stats_table.add_row("Original file size", f"{original_size:,} bytes")
                    stats_table.add_row("Markdown file size", f"{output_size:,} bytes")
                    stats_table.add_row("Content length", f"{len(markdown_content):,} characters")
                    stats_table.add_row("Compression ratio", f"{output_size/original_size:.2%}")
                    
                    console.print(stats_table)
                else:
                    console.print("[bold red]ERROR: Failed to save markdown file[/bold red]")
                    sys.exit(1)
            else:
                console.print("[bold red]ERROR: Conversion failed[/bold red]")
                sys.exit(1)
        else:
            # Multiple files processing
            results = converter.convert_multiple_files(all_files)
            
            # Display detailed results
            success_files = [f for f, output in results.items() if output is not None]
            failed_files = [f for f, output in results.items() if output is None]
            
            if success_files:
                success_table = Table(title="Successfully Converted Files")
                success_table.add_column("Input File", style="green")
                success_table.add_column("Output File", style="cyan")
                success_table.add_column("Size", style="yellow")
                
                total_original_size = 0
                total_output_size = 0
                
                for input_file in success_files:
                    output_path = results[input_file]
                    original_size = input_file.stat().st_size
                    output_size = output_path.stat().st_size
                    
                    total_original_size += original_size
                    total_output_size += output_size
                    
                    success_table.add_row(
                        input_file.name,
                        output_path.name,
                        f"{original_size:,} -> {output_size:,} bytes"
                    )
                
                console.print(success_table)
                
                # Overall statistics
                overall_table = Table(title="Overall Statistics")
                overall_table.add_column("Metric", style="cyan")
                overall_table.add_column("Value", style="green")
                
                overall_table.add_row("Total files processed", f"{len(all_files)}")
                overall_table.add_row("Successfully converted", f"{len(success_files)}")
                overall_table.add_row("Failed conversions", f"{len(failed_files)}")
                overall_table.add_row("Total original size", f"{total_original_size:,} bytes")
                overall_table.add_row("Total output size", f"{total_output_size:,} bytes")
                overall_table.add_row("Overall compression ratio", f"{total_output_size/total_original_size:.2%}" if total_original_size > 0 else "N/A")
                
                console.print(overall_table)
            
            if failed_files:
                console.print(f"\n[bold red]Failed to convert {len(failed_files)} file(s):[/bold red]")
                for failed_file in failed_files:
                    console.print(f"  - {failed_file.name}")
            
            # Exit with error code if any failures
            if failed_files:
                sys.exit(1)
            
    except KeyboardInterrupt:
        console.print("\n[yellow]WARNING: Conversion interrupted by user[/yellow]")
        sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]ERROR: Unexpected error: {e}[/bold red]")
        if verbose:
            console.print(f"[red]{traceback.format_exc()}[/red]")
        sys.exit(1)


if __name__ == "__main__":
    main()