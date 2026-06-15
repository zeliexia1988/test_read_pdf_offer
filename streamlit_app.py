#!/usr/bin/env python3
"""
PDF Quote Extractor - Streamlit Web App
从PDF报价单中提取信息并导出Excel
"""

import streamlit as st
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import json
from pathlib import Path

# 页面配置
st.set_page_config(
    page_title="PDF 报价提取工具",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义样式
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
    }
    </style>
    """, unsafe_allow_html=True)

TARGET_COLUMNS = [
    'Numéro',
    'Libellé / Désignation',
    'Unité',
    'Quantité',
    'PU EXW excluding tax $',
    'Montant Total EXW excluding tax $'#!/usr/bin/env python3
"""
PDF Quote Extractor - Streamlit Web App
从PDF报价单中提取信息并导出Excel
"""

import streamlit as st
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import json
from pathlib import Path

# OCR相关导入
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# 页面配置
st.set_page_config(
    page_title="PDF 报价提取工具",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义样式
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
    }
    </style>
    """, unsafe_allow_html=True)

TARGET_COLUMNS = [
    'Numéro',
    'Libellé / Désignation',
    'Unité',
    'Quantité',
    'PU EXW excluding tax $',
    'Montant Total EXW excluding tax $'
]

class QuoteExtractor:
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
        self.tables = []
        self.text_content = ""
        self.extracted_data = []
    
    def extract(self):
        """主提取流程"""
        try:
            with pdfplumber.open(self.pdf_file) as pdf:
                # 检查是否为扫描版本
                if len(pdf.pages) > 0:
                    first_page = pdf.pages[0]
                    text = first_page.extract_text()
                    
                    # 如果文本为空，尝试OCR
                    if not text or len(text.strip()) < 10:
                        if OCR_AVAILABLE:
                            st.info("📄 检测到扫描版PDF，正在使用OCR识别...")
                            return self._extract_with_ocr()
                        else:
                            st.error("❌ 这是一个扫描版本的PDF（图片）")
                            st.warning("OCR模块未安装。请联系管理员或使用可复制文本的PDF。")
                            return False
                
                # 正常的文本提取流程
                for page_num, page in enumerate(pdf.pages, 1):
                    page_tables = page.extract_tables()
                    if page_tables:
                        for table in page_tables:
                            self.tables.append({
                                'page': page_num,
                                'data': table
                            })
                    
                    text = page.extract_text() or ""
                    self.text_content += f"\n--- Page {page_num} ---\n{text}"
            
            self._find_and_parse_quote_table()
            return True
        
        except Exception as e:
            st.error(f"❌ 读取PDF错误: {e}")
            return False
    
    def _extract_with_ocr(self):
        """使用OCR处理扫描版PDF"""
        try:
            # 将PDF转换为图像
            images = convert_from_bytes(self.pdf_file.read())
            st.info(f"📸 已转换为 {len(images)} 张图像，正在OCR识别...")
            
            progress_bar = st.progress(0)
            
            for page_num, image in enumerate(images, 1):
                # 使用pytesseract进行OCR
                text = pytesseract.image_to_string(image, lang='fra+eng')
                self.text_content += f"\n--- Page {page_num} ---\n{text}"
                
                progress_bar.progress(page_num / len(images))
            
            progress_bar.empty()
            
            # OCR后尝试从文本中提取表格
            self._extract_table_from_text()
            st.success("✅ OCR识别完成！")
            return True
        
        except Exception as e:
            st.error(f"❌ OCR处理失败: {e}")
            st.info("💡 建议：获取可复制文本的PDF版本或在线转换")
            return False
    
    def _extract_table_from_text(self):
        """从OCR识别的文本中提取表格数据"""
        lines = self.text_content.split('\n')
        current_row = []
        
        for line in lines:
            line = line.strip()
            if not line:
                if current_row:
                    self.extracted_data.append(current_row)
                    current_row = []
                continue
            
            # 简单的行解析
            parts = line.split()
            if len(parts) >= 3:
                current_row = [''] * len(TARGET_COLUMNS)
                
                # 试图识别数字和文本
                current_row[1] = line  # 描述放在第二列
                
                # 查找数字
                for i, part in enumerate(parts):
                    try:
                        num = float(part.replace(',', '.'))
                        if i == len(parts) - 1:  # 最后一个数字是总金额
                            current_row[5] = num
                        elif i == len(parts) - 2:  # 倒数第二个是单价
                            current_row[4] = num
                    except:
                        pass
        
        if current_row:
            self.extracted_data.append(current_row)
    
    def _find_and_parse_quote_table(self):
        """查找并解析报价表格"""
        if not self.tables:
            st.warning("⚠️ 未找到表格，尝试从文本提取...")
            return
        
        quote_table_info = self._identify_quote_table()
        
        if quote_table_info:
            self._parse_table_based_quote(quote_table_info)
        else:
            largest_table = max(self.tables, key=lambda x: len(x['data']))
            self._parse_table_based_quote(largest_table)
    
    def _identify_quote_table(self):
        """智能识别报价表格"""
        for table_info in self.tables:
            table = table_info['data']
            if len(table) < 2:
                continue
            
            header = ' '.join(str(cell).lower() for cell in table[0] if cell)
            
            keywords = ['numéro', 'désignation', 'description', 'unité', 
                       'quantité', 'qty', 'prix', 'pu', 'montant', 'total',
                       'reference', 'libellé', 'article']
            keyword_count = sum(1 for kw in keywords if kw in header)
            
            # 降低匹配阈值，从3降到2
            if keyword_count >= 2:
                return table_info
        
        return None
    
    def _parse_table_based_quote(self, table_info):
        """从表格解析报价数据"""
        table = table_info['data']
        
        cleaned_table = self._clean_table(table)
        
        if len(cleaned_table) < 2:
            return
        
        header = cleaned_table[0]
        col_mapping = self._map_columns(header)
        
        for row in cleaned_table[1:]:
            if not any(cell.strip() for cell in row):
                continue
            
            extracted_row = self._map_row_to_target(row, col_mapping)
            if any(extracted_row):
                self.extracted_data.append(extracted_row)
    
    def _clean_table(self, table):
        """清理表格数据"""
        cleaned = []
        for row in table:
            if any(cell for cell in row if cell):
                cleaned_row = [
                    str(cell).strip() if cell else ""
                    for cell in row
                ]
                cleaned.append(cleaned_row)
        return cleaned
    
    def _map_columns(self, header_row):
        """智能列映射"""
        mapping = {}
        header_lower = [str(h).lower() for h in header_row]
        
        rules = {
            'Numéro': ['numéro', 'n°', 'num', 'reference', 'ref', 'article', 'code'],
            'Libellé / Désignation': ['libellé', 'désignation', 'description', 'item', 'product', 'designation', 'article name'],
            'Unité': ['unité', 'unit', 'u', 'um', 'pcs', 'pièce'],
            'Quantité': ['quantité', 'qty', 'qté', 'quantity', 'q', 'qté.'],
            'PU EXW excluding tax $': ['pu', 'prix', 'price', 'unit price', 'unitaire', 'pu ht', 'unit', 'price/unit', 'pu exw'],
            'Montant Total EXW excluding tax $': ['montant', 'total', 'amount', 'montant total', 'montant ht', 'total ht', 'total amount'],
        }
        
        for target_col, keywords in rules.items():
            for idx, header in enumerate(header_lower):
                if any(kw in header for kw in keywords):
                    mapping[idx] = target_col
                    break
        
        return mapping
    
    def _map_row_to_target(self, row, col_mapping):
        """将行数据映射到目标列"""
        result = [''] * len(TARGET_COLUMNS)
        
        for src_idx, target_col_name in col_mapping.items():
            if src_idx < len(row):
                target_idx = TARGET_COLUMNS.index(target_col_name)
                result[target_idx] = row[src_idx]
        
        return result
    
    def to_excel(self):
        """生成Excel字节流"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Devis"
        
        # 表头
        for col_idx, col_name in enumerate(TARGET_COLUMNS, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = col_name
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 数据
        for row_idx, row_data in enumerate(self.extracted_data, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value


if __name__ == "__main__":
    main()
