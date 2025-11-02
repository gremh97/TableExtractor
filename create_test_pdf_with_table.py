#!/usr/bin/env python3
"""
테스트용 테이블이 포함된 PDF 생성
"""

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import os

def create_test_pdf_with_table():
    """테이블이 포함된 테스트 PDF 생성"""
    
    # temperal_pdf 디렉토리 생성
    temperal_dir = "temperal_pdf"
    os.makedirs(temperal_dir, exist_ok=True)
    
    # PDF 파일 경로
    pdf_path = os.path.join(temperal_dir, "test_pdfplumber_table.pdf")
    
    # PDF 문서 생성
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    elements = []
    
    # 스타일 가져오기
    styles = getSampleStyleSheet()
    
    # 제목 추가
    title = Paragraph("테스트 문서 - pdfplumber 테이블 감지용", styles['Title'])
    elements.append(title)
    elements.append(Spacer(1, 20))
    
    # 설명 텍스트
    description = Paragraph("이 문서는 pdfplumber 테이블 감지 기능을 테스트하기 위해 생성되었습니다.", styles['Normal'])
    elements.append(description)
    elements.append(Spacer(1, 30))
    
    # 테이블 1 - 간단한 데이터
    table1_data = [
        ['항목', '수량', '단가', '합계'],
        ['사과', '10개', '1,000원', '10,000원'],
        ['바나나', '5개', '1,500원', '7,500원'],
        ['오렌지', '8개', '1,200원', '9,600원'],
        ['합계', '', '', '27,100원']
    ]
    
    table1 = Table(table1_data)
    table1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    elements.append(Paragraph("표 1: 과일 판매 현황", styles['Heading2']))
    elements.append(Spacer(1, 10))
    elements.append(table1)
    elements.append(Spacer(1, 40))
    
    # 테이블 2 - 다른 형태의 데이터
    table2_data = [
        ['지역', '2022년', '2023년', '증감률'],
        ['서울', '1,234,567', '1,345,678', '+9.0%'],
        ['부산', '456,789', '523,456', '+14.6%'],
        ['대구', '234,567', '267,890', '+14.2%'],
        ['인천', '345,678', '378,901', '+9.6%']
    ]
    
    table2 = Table(table2_data)
    table2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightblue),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    elements.append(Paragraph("표 2: 지역별 인구 현황", styles['Heading2']))
    elements.append(Spacer(1, 10))
    elements.append(table2)
    elements.append(Spacer(1, 20))
    
    # 마무리 텍스트
    footer = Paragraph("pdfplumber를 사용하여 위의 두 테이블을 정확히 감지하고 추출할 수 있는지 테스트합니다.", styles['Normal'])
    elements.append(footer)
    
    # PDF 생성
    doc.build(elements)
    print(f"테스트 PDF 생성 완료: {pdf_path}")
    
    return pdf_path

if __name__ == "__main__":
    create_test_pdf_with_table()