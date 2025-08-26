#!/usr/bin/env python3
"""
PDF OCR Module
==============

Provides OCR functionality to make unsearchable PDFs searchable using Tesseract and ocrmypdf.
"""

import logging
import io
from pathlib import Path
from typing import Optional, Dict, Any
from dataclasses import dataclass
from enum import Enum
import tempfile
import shutil

import ocrmypdf
import fitz  # PyMuPDF
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn
try:
    from PIL import Image
    import pytesseract
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


class OCRQuality(Enum):
    """OCR quality levels."""
    FAST = "fast"
    BALANCED = "balanced"
    HIGH = "high"


@dataclass
class OCRConfig:
    """Configuration for OCR processing."""
    quality: OCRQuality = OCRQuality.BALANCED
    language: str = "eng"
    rotate_pages: bool = True
    deskew: bool = False  # Disable deskew to avoid unpaper dependency
    clean: bool = False  # Disable clean to avoid unpaper dependency
    force_ocr: bool = False
    skip_text: bool = True
    remove_background: bool = False
    optimize_level: int = 0  # Disable optimization by default to avoid pngquant dependency
    oversample: int = 300
    preserve_layout: bool = True
    
    def to_ocrmypdf_args(self) -> Dict[str, Any]:
        """Convert config to ocrmypdf arguments."""
        args = {
            'language': self.language,
            'rotate_pages': self.rotate_pages,
            'oversample': self.oversample,
        }
        
        # Only add optimize if > 0 to avoid pngquant dependency
        if self.optimize_level > 0:
            args['optimize'] = self.optimize_level
        
        # force_ocr and skip_text are mutually exclusive
        if self.force_ocr:
            args['force_ocr'] = True
        else:
            args['skip_text'] = self.skip_text
        
        # Only add these if explicitly enabled to avoid unpaper dependency
        if self.deskew:
            args['deskew'] = True
        if self.clean:
            args['clean'] = True
            args['clean_final'] = True
        if self.remove_background:
            args['remove_background'] = True
        
        # Preserve layout for technical documents
        if self.preserve_layout:
            args['pdf_renderer'] = 'sandwich'
            # Don't set tesseract_config here as it causes issues
            # The default OCR settings work fine
        
        # Adjust settings based on quality
        if self.quality == OCRQuality.FAST:
            args['oversample'] = 200
            # Only optimize if pngquant is available
            # args['optimize'] = 1
        elif self.quality == OCRQuality.HIGH:
            args['oversample'] = 400
            # Only enable these features if dependencies are installed:
            # - pngquant for optimize: sudo apt-get install pngquant
            # - unpaper for deskew/clean: sudo apt-get install unpaper
            # args['optimize'] = 3
            # args['deskew'] = True
            # args['clean'] = True
        
        return args


class PDFOCRProcessor:
    """Handles OCR processing for PDF files."""
    
    def __init__(self, config: Optional[OCRConfig] = None):
        self.config = config or OCRConfig()
        self.console = Console()
        self.logger = self._setup_logging()
    
    def _setup_logging(self) -> logging.Logger:
        """Configure logging."""
        logger = logging.getLogger("pdf_ocr")
        if not logger.handlers:
            logger.setLevel(logging.INFO)
        return logger
    
    def check_pdf_has_text(self, pdf_path: Path, min_text_length: int = 100) -> bool:
        """
        Check if PDF already has searchable text.
        
        Args:
            pdf_path: Path to PDF file
            min_text_length: Minimum text length to consider as searchable
            
        Returns:
            True if PDF has searchable text, False otherwise
        """
        try:
            doc = fitz.open(str(pdf_path))
            total_text = ""
            
            for page in doc:
                text = page.get_text()
                total_text += text
                
                # Early return if we found enough text
                if len(total_text.strip()) > min_text_length:
                    doc.close()
                    return True
            
            doc.close()
            return len(total_text.strip()) > min_text_length
            
        except Exception as e:
            self.logger.error(f"Error checking PDF for text: {e}")
            return False
    
    def get_pdf_info(self, pdf_path: Path) -> Dict[str, Any]:
        """
        Get information about a PDF file.
        
        Args:
            pdf_path: Path to PDF file
            
        Returns:
            Dictionary with PDF information
        """
        try:
            doc = fitz.open(str(pdf_path))
            info = {
                'page_count': len(doc),
                'has_text': False,
                'file_size_mb': pdf_path.stat().st_size / (1024 * 1024),
                'metadata': doc.metadata,
            }
            
            # Check for text
            text_sample = ""
            for i, page in enumerate(doc):
                if i >= 3:  # Check first 3 pages
                    break
                text_sample += page.get_text()
            
            info['has_text'] = len(text_sample.strip()) > 100
            doc.close()
            
            return info
            
        except Exception as e:
            self.logger.error(f"Error getting PDF info: {e}")
            return {}
    
    def process_pdf(
        self, 
        input_path: Path, 
        output_path: Optional[Path] = None,
        check_existing_text: bool = True,
        prompt_on_existing: bool = False,
        attempt_repair: bool = True
    ) -> Optional[Path]:
        """
        Process a PDF file with OCR.
        
        Args:
            input_path: Path to input PDF
            output_path: Path for output PDF (if None, adds '_ocr' suffix)
            check_existing_text: Check if PDF already has text
            prompt_on_existing: Prompt user if PDF already has text
            
        Returns:
            Path to processed PDF if successful, None otherwise
        """
        if not input_path.exists():
            self.logger.error(f"Input file does not exist: {input_path}")
            return None
        
        # Generate output path if not provided
        if output_path is None:
            output_path = input_path.parent / f"{input_path.stem}_ocr{input_path.suffix}"
        
        # Check if PDF already has text
        if check_existing_text and self.check_pdf_has_text(input_path):
            self.logger.info(f"PDF already has searchable text: {input_path}")
            
            if prompt_on_existing:
                response = input("Continue with OCR anyway? (y/n): ")
                if response.lower() != 'y':
                    self.logger.info("Skipping OCR as PDF already has text")
                    return None
            elif not self.config.force_ocr:
                self.logger.info("Skipping OCR as PDF already has text (use force_ocr to override)")
                return input_path  # Return original path as it's already searchable
        
        # Get PDF info for logging
        pdf_info = self.get_pdf_info(input_path)
        self.logger.info(f"Processing PDF: {input_path.name}")
        self.logger.info(f"  Pages: {pdf_info.get('page_count', 'unknown')}")
        self.logger.info(f"  Size: {pdf_info.get('file_size_mb', 0):.2f} MB")
        self.logger.info(f"  Has text: {pdf_info.get('has_text', False)}")
        
        # Prepare ocrmypdf arguments
        ocr_args = self.config.to_ocrmypdf_args()
        
        # If PDF might be problematic, try with minimal settings first
        actual_input = input_path
        if attempt_repair:
            try:
                # Try to re-save the PDF using PyMuPDF to fix potential issues
                doc = fitz.open(str(input_path))
                if len(doc) > 0:
                    temp_repaired = input_path.parent / f"{input_path.stem}_temp_repaired.pdf"
                    doc.save(str(temp_repaired), garbage=4, deflate=True, clean=True)
                    doc.close()
                    
                    if temp_repaired.exists():
                        self.logger.debug("Created repaired PDF for OCR processing")
                        actual_input = temp_repaired
                else:
                    doc.close()
            except Exception as e:
                self.logger.debug(f"Could not repair PDF: {e}")
        
        try:
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                console=self.console
            ) as progress:
                task = progress.add_task(
                    f"Running OCR on {input_path.name}...", 
                    total=None
                )
                
                # Run OCR
                ocrmypdf.ocr(
                    input_file=str(actual_input),
                    output_file=str(output_path),
                    **ocr_args
                )
                
                progress.update(task, completed=100)
            
            # Clean up temporary repaired file if it was created
            if attempt_repair and actual_input != input_path and actual_input.exists():
                try:
                    actual_input.unlink()
                    self.logger.debug("Cleaned up temporary repaired PDF")
                except:
                    pass
            
            # Verify output
            if output_path.exists():
                if self.check_pdf_has_text(output_path):
                    self.logger.info(f"OCR completed successfully: {output_path}")
                    
                    # Log file size comparison
                    original_size = input_path.stat().st_size / (1024 * 1024)
                    new_size = output_path.stat().st_size / (1024 * 1024)
                    size_change = ((new_size - original_size) / original_size) * 100
                    
                    self.logger.info(f"  Original size: {original_size:.2f} MB")
                    self.logger.info(f"  New size: {new_size:.2f} MB")
                    self.logger.info(f"  Size change: {size_change:+.1f}%")
                    
                    return output_path
                else:
                    self.logger.warning("Output PDF may not have searchable text")
                    return output_path
            else:
                self.logger.error("Output file was not created")
                return None
                
        except ocrmypdf.exceptions.PriorOcrFoundError:
            self.logger.info("PDF already contains text (detected by ocrmypdf)")
            return input_path
            
        except ocrmypdf.exceptions.InputFileError as e:
            self.logger.error(f"Input file error: {e}")
            self.logger.error("This PDF may be corrupted or use unsupported features")
            # Clean up temporary repaired file if it exists
            if attempt_repair and actual_input != input_path and actual_input.exists():
                try:
                    actual_input.unlink()
                except:
                    pass
            return None
            
        except ocrmypdf.exceptions.SubprocessOutputError as e:
            self.logger.error(f"OCR subprocess failed: {e}")
            if "Ghostscript" in str(e):
                self.logger.error("Ghostscript failed to process this PDF. The file may:")
                self.logger.error("  - Be corrupted or malformed")
                self.logger.error("  - Use unsupported PDF features")
                self.logger.error("  - Contain complex graphics that can't be rasterized")
            elif "tesseract" in str(e).lower():
                self.logger.error("Tesseract OCR failed. Check that:")
                self.logger.error("  - Tesseract is properly installed")
                self.logger.error("  - The specified language pack is available")
            # Clean up temporary repaired file if it exists
            if attempt_repair and actual_input != input_path and actual_input.exists():
                try:
                    actual_input.unlink()
                except:
                    pass
            return None
            
        except Exception as e:
            self.logger.error(f"OCR failed: {e}")
            self.logger.debug(f"Error type: {type(e).__name__}")
            
            # Clean up temporary repaired file if it exists
            if attempt_repair and actual_input != input_path and actual_input.exists():
                try:
                    actual_input.unlink()
                except:
                    pass
            
            return None
    
    def ocr_pdf_via_images(self, pdf_path: Path, language: str = "eng") -> Optional[str]:
        """
        Alternative OCR method: Convert PDF pages to images and OCR them.
        This works when Ghostscript fails to process the PDF directly.
        
        Args:
            pdf_path: Path to PDF file
            language: OCR language
            
        Returns:
            Extracted text or None if failed
        """
        if not HAS_PIL:
            self.logger.error("PIL/Pillow not installed. Cannot use image-based OCR.")
            self.logger.error("Install with: pip install pillow pytesseract")
            return None
            
        try:
            self.logger.info("Attempting OCR via image extraction...")
            doc = fitz.open(str(pdf_path))
            all_text = []
            
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                console=self.console
            ) as progress:
                task = progress.add_task(
                    f"OCR via images...", 
                    total=len(doc)
                )
                
                for page_num, page in enumerate(doc):
                    try:
                        # Render page as image (higher DPI for better OCR)
                        mat = fitz.Matrix(300/72, 300/72)  # 300 DPI
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        
                        # Convert to PIL Image
                        img_data = pix.pil_tobytes(format="PNG")
                        img = Image.open(io.BytesIO(img_data))
                        
                        # OCR the image
                        text = pytesseract.image_to_string(img, lang=language)
                        
                        if text.strip():
                            all_text.append(f"\n{'='*50}")
                            all_text.append(f"PAGE {page_num + 1}")
                            all_text.append('='*50)
                            all_text.append(text)
                        
                        progress.advance(task)
                        
                    except Exception as e:
                        self.logger.warning(f"Failed to OCR page {page_num + 1}: {e}")
                        progress.advance(task)
            
            doc.close()
            
            if all_text:
                self.logger.info("Successfully extracted text via image OCR")
                return '\n'.join(all_text)
            else:
                self.logger.warning("No text extracted via image OCR")
                return None
                
        except Exception as e:
            self.logger.error(f"Image-based OCR failed: {e}")
            return None
    
    def extract_text_from_pdf(self, pdf_path: Path) -> Optional[str]:
        """
        Extract all text from a PDF.
        
        Args:
            pdf_path: Path to PDF file
            
        Returns:
            Extracted text as string, or None if failed
        """
        try:
            doc = fitz.open(str(pdf_path))
            full_text = []
            
            for page_num, page in enumerate(doc, 1):
                text = page.get_text()
                if text.strip():
                    full_text.append(f"\n{'='*50}")
                    full_text.append(f"PAGE {page_num}")
                    full_text.append('='*50)
                    full_text.append(text)
            
            doc.close()
            
            return '\n'.join(full_text) if full_text else None
            
        except Exception as e:
            self.logger.error(f"Failed to extract text: {e}")
            return None


def create_ocr_processor(
    quality: str = "balanced",
    language: str = "eng",
    force_ocr: bool = False,
    preserve_layout: bool = True
) -> PDFOCRProcessor:
    """
    Factory function to create an OCR processor with common settings.
    
    Args:
        quality: OCR quality level ("fast", "balanced", or "high")
        language: OCR language(s) - can be combined like "eng+fra"
        force_ocr: Force OCR even if PDF has text
        preserve_layout: Preserve original PDF layout
        
    Returns:
        Configured PDFOCRProcessor instance
    """
    quality_enum = OCRQuality(quality.lower())
    
    config = OCRConfig(
        quality=quality_enum,
        language=language,
        force_ocr=force_ocr,
        preserve_layout=preserve_layout
    )
    
    return PDFOCRProcessor(config)