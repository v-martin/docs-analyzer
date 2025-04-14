from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

class DocumentValidator:
    """A class that provides detailed validation of DOCX documents according to the given rules"""
    
    def __init__(self, doc_path):
        self.doc_path = doc_path
        self.doc = Document(doc_path)
        self.issues = []
        
    def validate_all(self):
        """Run all validation checks and return found issues"""
        self.issues = []
        
        # Run all validation methods
        self.validate_page_format()
        self.validate_font()
        self.validate_paragraphs()
        self.validate_headings()
        self.validate_page_numbering()
        self.validate_tables()
        self.validate_figures()
        self.validate_equations()
        self.validate_references()
        self.validate_appendices()
        
        return self.issues
        
    def validate_page_format(self):
        """Validate page size, margins, and orientation"""
        for section in self.doc.sections:
            # Unit conversion corrections 
            # Different conversion options - try EMU-based conversion
            # 914400 EMU = 1 inch = 25.4 mm
            # 12700 EMU = 1 point = 0.352778 mm
            
            try:
                # Try different conversion approaches to see which works best
                # Approach 1: standard twip conversion
                conversion_factor_1 = 25.4 / 1440  # 1440 twips = 1 inch = 25.4 mm
                
                # Approach 2: direct size check using docx attributes (if available)
                if hasattr(section, 'page_height_emu') and hasattr(section, 'page_width_emu'):
                    emu_to_mm = 25.4 / 914400
                    width_mm_direct = section.page_width_emu * emu_to_mm
                    height_mm_direct = section.page_height_emu * emu_to_mm
                    width_mm = width_mm_direct
                    height_mm = height_mm_direct
                else:
                    width_mm = section.page_width * conversion_factor_1
                    height_mm = section.page_height * conversion_factor_1
                
                # Check common page sizes with tolerance
                is_a4 = (209 <= width_mm <= 211 and 296 <= height_mm <= 298)
                is_letter = (214 <= width_mm <= 216 and 278 <= height_mm <= 280)  # 8.5" x 11"
                
                if not (is_a4 or is_letter):
                    # Relax the constraint a bit - just check if it's close to A4
                    if 180 <= width_mm <= 240 and 250 <= height_mm <= 330:
                        # It's reasonably close to A4/Letter, don't report an issue
                        pass
                    else:
                        self.issues.append(f"Размер страницы не соответствует формату A4 (210×297 мм). Текущий размер: {width_mm:.1f}×{height_mm:.1f} мм.")
                
                # Get page margin measurements with the same conversion factor
                left_margin_mm = section.left_margin * conversion_factor_1
                right_margin_mm = section.right_margin * conversion_factor_1
                top_margin_mm = section.top_margin * conversion_factor_1
                bottom_margin_mm = section.bottom_margin * conversion_factor_1
                
                # Apply tolerance to margin checks
                if not 25 <= left_margin_mm <= 35:  # 30mm ± 5mm
                    self.issues.append(f"Левое поле должно быть 30 мм. Текущее: {left_margin_mm:.1f} мм.")
                
                if not (10 <= right_margin_mm <= 20):  # 15mm ± 5mm
                    self.issues.append(f"Правое поле должно быть 15 мм (допускается 10 мм). Текущее: {right_margin_mm:.1f} мм.")
                
                if not 15 <= top_margin_mm <= 25:  # 20mm ± 5mm
                    self.issues.append(f"Верхнее поле должно быть 20 мм. Текущее: {top_margin_mm:.1f} мм.")
                
                if not 15 <= bottom_margin_mm <= 25:  # 20mm ± 5mm
                    self.issues.append(f"Нижнее поле должно быть 20 мм. Текущее: {bottom_margin_mm:.1f} мм.")
                    
            except Exception as e:
                self.issues.append(f"Ошибка при проверке размеров страницы: {str(e)}")
    
    def validate_font(self):
        """Validate the font type and size in the document"""
        for paragraph in self.doc.paragraphs:
            if not paragraph.text.strip():
                continue
                
            for run in paragraph.runs:
                if run.font.name and 'Times New Roman' not in run.font.name:
                    self.issues.append(f"Шрифт должен быть Times New Roman. Текущий: {run.font.name} в тексте: '{run.text[:20]}...'")
                
                if run.font.size is not None:
                    font_size_pt = run.font.size.pt
                    if font_size_pt < 12:
                        self.issues.append(f"Размер шрифта должен быть не менее 12 пт. Текущий: {font_size_pt} пт в тексте: '{run.text[:20]}...'")
    
    def validate_paragraphs(self):
        """Validate paragraph formatting including line spacing and indentation"""
        # Define patterns for headings that should not be checked for justification
        heading_patterns = [
            r'^(СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ПРИЛОЖЕНИЕ|СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ|СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ|ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ)',
            r'^(\d+)(\.\d+)*\s+[А-ЯЁ]',  # Heading with numbers like "1", "1.1", "1.1.1" etc.
            r'^Рисунок\s+\d+',  # Figure captions
            r'^Таблица\s+\d+'   # Table captions
        ]
        
        for paragraph in self.doc.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue
                
            # Skip checking alignment for headings and captions
            is_heading = False
            for pattern in heading_patterns:
                if re.match(pattern, text, re.IGNORECASE):
                    is_heading = True
                    break
                    
            # Also check for TOC entries specifically
            is_toc_entry = '\t' in paragraph.text or text.endswith('...')
            
            # Check text alignment only for regular paragraphs
            if not is_heading and not is_toc_entry and paragraph.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                self.issues.append(f"Основной текст должен быть выровнен по ширине. Текущее выравнивание: {paragraph.alignment} в тексте: '{text[:20]}...'")
            
            # Check paragraph indentation
            if paragraph.paragraph_format.first_line_indent is not None:
                # Convert to cm using the correct conversion factor
                conversion_factor = 2.54 / 1440  # 1440 twips = 1 inch = 2.54 cm
                first_line_indent_cm = paragraph.paragraph_format.first_line_indent * conversion_factor
                
                # Skip checking indentation for headings and captions
                if not is_heading and not is_toc_entry and not 1.2 <= first_line_indent_cm <= 1.3:
                    self.issues.append(f"Абзацный отступ должен быть 1,25 см. Текущий: {first_line_indent_cm:.2f} см в тексте: '{text[:20]}...'")
            
            # Check line spacing
            if paragraph.paragraph_format.line_spacing is not None:
                line_spacing = paragraph.paragraph_format.line_spacing
                if not 1.4 <= line_spacing <= 1.6:
                    self.issues.append(f"Межстрочный интервал должен быть полуторным (1.5). Текущий: {line_spacing:.1f} в тексте: '{text[:20]}...'")
    
    def validate_headings(self):
        """Validate heading format and structure"""
        # Find all headings
        headings = []
        heading_patterns = [
            (r'^(СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ПРИЛОЖЕНИЕ|СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ|СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ|ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ)(\s*$|\.{3}.*)', 0),
            (r'^(\d+)\s+[А-ЯЁ]', 1),
            (r'^(\d+\.\d+)\s+[А-ЯЁ]', 2),
            (r'^(\d+\.\d+\.\d+)\s+[А-ЯЁ]', 3)
        ]
        
        for paragraph in self.doc.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue
                
            # Check if paragraph is a heading based on patterns
            for pattern, level in heading_patterns:
                if re.match(pattern, text, re.IGNORECASE):
                    headings.append((text, level, paragraph))
                    break
        
        # Validate each heading
        for text, level, paragraph in headings:
            # Check heading format
            if level == 0:  # Main structural elements
                # Should be centered, uppercase, without numbering
                # Check specifically when the alignment is 0 (None) but should be 1 (CENTER)
                if paragraph.alignment is None or paragraph.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    # Skip if it contains tab character, it might be in TOC
                    if '\t' not in text and not text.endswith('...'):
                        self.issues.append(f"Заголовок '{text}' должен быть выровнен по центру.")
                
                # Basic structural elements should be uppercase
                # Skip check if it contains tabs (might be TOC)
                if not text.isupper() and '\t' not in text and not text.endswith('...'):
                    # Check if first part is uppercase
                    if not text.split()[0].isupper():
                        self.issues.append(f"Заголовок '{text}' должен быть написан прописными буквами.")
            
            else:  # Regular headings (chapters, sections, subsections)
                # Should have proper numbering, start with capital letter, no period at end
                if text.endswith('.'):
                    self.issues.append(f"Заголовок '{text}' не должен оканчиваться точкой.")
                
                # Check bold formatting more carefully
                if paragraph.runs:
                    is_bold = True
                    for run in paragraph.runs:
                        if run.text.strip() and not run.bold:
                            is_bold = False
                            break
                            
                    if not is_bold:
                        self.issues.append(f"Заголовок '{text}' должен быть выделен полужирным шрифтом.")
                
                # Validate numbering sequence (basic check)
                if level == 1:
                    try:
                        number = int(re.match(r'(\d+)', text).group(1))
                        if number < 1:
                            self.issues.append(f"Некорректная нумерация заголовка '{text}'.")
                    except:
                        self.issues.append(f"Некорректная нумерация заголовка '{text}'.")
    
    def validate_page_numbering(self):
        """Check page numbering in the document"""
        try:
            has_page_numbers = False
            for section in self.doc.sections:
                if hasattr(section, 'footer') and section.footer and section.footer.paragraphs:
                    footer_paragraphs = section.footer.paragraphs
                    if any(p.text.strip() for p in footer_paragraphs):
                        has_page_numbers = True
                        # Check if page numbers are centered
                        for p in footer_paragraphs:
                            if p.text.strip() and p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                                self.issues.append("Номера страниц должны быть размещены в центре нижней части страницы.")
                        break
            
            if not has_page_numbers:
                self.issues.append("Не обнаружена нумерация страниц. Номера должны быть в центре нижней части страницы.")
        except Exception as e:
            self.issues.append(f"Ошибка при проверке нумерации страниц: {str(e)}")
    
    def validate_tables(self):
        """Validate table formatting"""
        if not self.doc.tables:
            return
            
        for i, table in enumerate(self.doc.tables, 1):
            # Check if table has a caption
            has_caption = False
            caption_paragraph = None
            caption_pattern = r"^Таблица\s+" + str(i) + r"(\s+[-–—]\s+|\s*$)"
            
            for paragraph in self.doc.paragraphs:
                if re.match(caption_pattern, paragraph.text.strip()):
                    has_caption = True
                    caption_paragraph = paragraph
                    break
            
            if not has_caption:
                self.issues.append(f"Таблица #{i} не имеет подписи. Подпись должна быть в формате 'Таблица {i} - Наименование'.")
            elif caption_paragraph:
                # Table captions should be left-aligned
                if caption_paragraph.alignment != WD_ALIGN_PARAGRAPH.LEFT and caption_paragraph.alignment is not None:
                    self.issues.append(f"Подпись таблицы '{caption_paragraph.text[:30]}...' должна быть выровнена по левому краю.")
            
            # Check for table referenced in text
            has_reference = False
            ref_pattern = r'таблиц[аеуыи]?\s+' + str(i) + r'\b'
            
            for paragraph in self.doc.paragraphs:
                if re.search(ref_pattern, paragraph.text, re.IGNORECASE):
                    has_reference = True
                    break
            
            if not has_reference:
                self.issues.append(f"На таблицу #{i} нет ссылки в тексте. На все таблицы должны быть ссылки.")
    
    def validate_figures(self):
        """Validate figure (illustrations) formatting"""
        # Find figures by looking for captions
        figure_captions = []
        caption_paragraphs = []
        
        for paragraph in self.doc.paragraphs:
            if re.match(r'^Рисунок\s+\d+', paragraph.text.strip()):
                figure_captions.append(paragraph.text.strip())
                caption_paragraphs.append(paragraph)
        
        if not figure_captions:
            return
            
        # Validate each figure
        for i, (caption, paragraph) in enumerate(zip(figure_captions, caption_paragraphs), 1):
            # Check caption format
            if not re.match(r'^Рисунок\s+\d+\s+[\-–—]\s+\w', caption):
                self.issues.append(f"Неправильный формат подписи рисунка: '{caption}'. Должно быть 'Рисунок Номер - Наименование'.")
            
            # Figure captions should be centered
            if paragraph.alignment != WD_ALIGN_PARAGRAPH.CENTER and paragraph.alignment is not None:
                self.issues.append(f"Подпись рисунка '{caption[:30]}...' должна быть выровнена по центру.")
            
            # Check if numbered correctly
            match = re.search(r'Рисунок\s+(\d+)', caption)
            if match:
                number = int(match.group(1))
                if number != i:
                    self.issues.append(f"Нарушена последовательность нумерации рисунков. Рисунок с подписью '{caption}' имеет номер {number}, ожидается {i}.")
            
            # Check for figure referenced in text
            has_reference = False
            ref_pattern = r'рисун[окаеи][ак]?\s+' + str(i) + r'\b'
            
            for paragraph in self.doc.paragraphs:
                if re.search(ref_pattern, paragraph.text, re.IGNORECASE):
                    has_reference = True
                    break
            
            if not has_reference:
                self.issues.append(f"На рисунок {i} нет ссылки в тексте. На все рисунки должны быть ссылки.")
    
    def validate_equations(self):
        """Validate equation formatting"""
        # Safer approach - iterate through paragraphs with index
        paragraphs = list(self.doc.paragraphs)
        for i, paragraph in enumerate(paragraphs):
            text = paragraph.text.strip()
            # Very basic equation detection
            if ('=' in text and sum(c.isalpha() for c in text) > 0 and 
                sum(c.isdigit() or c in '+-*/=()[]{}' for c in text) > 0):
                
                # Check for proper spacing around equation
                prev_index = i - 1
                next_index = i + 1
                
                if prev_index >= 0 and paragraphs[prev_index].text.strip():
                    if not paragraphs[prev_index].paragraph_format.space_after:
                        self.issues.append(f"Перед формулой '{text[:30]}...' должна быть пустая строка.")
                
                if next_index < len(paragraphs) and paragraphs[next_index].text.strip():
                    if not paragraph.paragraph_format.space_after:
                        self.issues.append(f"После формулы '{text[:30]}...' должна быть пустая строка.")
                
                # Check for equation numbering
                if not re.search(r'\(\d+\.\d+\)$', text) and not re.search(r'\(\d+\)$', text):
                    self.issues.append(f"Формула '{text[:30]}...' должна иметь номер в скобках в конце строки.")
    
    def validate_references(self):
        """Validate references format"""
        has_references_section = False
        references_paragraphs = []
        in_references_section = False
        
        # Find references section - more robust approach
        for paragraph in self.doc.paragraphs:
            if "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in paragraph.text.upper():
                has_references_section = True
                in_references_section = True
                continue
            
            # Check if we've moved to another section
            if in_references_section and paragraph.text.strip() and any(
                section in paragraph.text.upper() 
                for section in ["ПРИЛОЖЕНИЕ", "СПИСОК СОКРАЩЕНИЙ", "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ"]
            ):
                in_references_section = False
            
            if in_references_section and paragraph.text.strip():
                references_paragraphs.append(paragraph.text.strip())
        
        if not has_references_section:
            self.issues.append("Не найден раздел 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ'.")
            return
        
        # Check references format and numbering
        for i, ref in enumerate(references_paragraphs, 1):
            if not ref.startswith(f"{i}."):
                self.issues.append(f"Источник '{ref[:50]}...' должен иметь номер {i}.")
            
            # Check for citation in text - safer approach
            has_citation = False
            citation_pattern = r"\[\s*" + str(i) + r"\s*\]"
            for paragraph in self.doc.paragraphs:
                if re.search(citation_pattern, paragraph.text):
                    has_citation = True
                    break
            
            if not has_citation and len(references_paragraphs) > 0:  # Only check if we found references
                self.issues.append(f"На источник #{i} нет ссылки в тексте. На все источники должны быть ссылки.")
    
    def validate_appendices(self):
        """Validate appendices format"""
        try:
            has_appendices = False
            for paragraph in self.doc.paragraphs:
                if paragraph.text.strip().startswith("Приложение "):
                    has_appendices = True
                    # Validate appendix heading format
                    if not paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                        self.issues.append(f"Заголовок '{paragraph.text}' должен быть выровнен по центру.")
                    
                    # Check appendix designation
                    match = re.match(r'Приложение\s+([А-Я])', paragraph.text.strip())
                    if not match:
                        self.issues.append(f"Неправильный формат обозначения приложения: '{paragraph.text}'. Должно быть 'Приложение' и буква русского алфавита.")
                    else:
                        letter = match.group(1)
                        if letter in 'ЁЗЙОЧЬЫЪ':
                            self.issues.append(f"Недопустимая буква '{letter}' для обозначения приложения.")
            
            # If there are appendices, check for mentions in text
            if has_appendices:
                has_reference = False
                for paragraph in self.doc.paragraphs:
                    p_text = paragraph.text.strip()
                    if re.search(r'приложени[еия]', p_text, re.IGNORECASE) and not p_text.startswith("Приложение "):
                        has_reference = True
                        break
                
                if not has_reference:
                    self.issues.append("На приложения нет ссылок в тексте. На все приложения должны быть ссылки.")
        except Exception as e:
            self.issues.append(f"Ошибка при проверке приложений: {str(e)}")
