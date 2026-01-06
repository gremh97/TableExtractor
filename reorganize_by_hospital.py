#!/usr/bin/env python3
"""
엑셀 파일의 URL과 page title을 참고하여 동일한 병원별로 데이터를 재배열하는 스크립트
"""

import pandas as pd
import os
import shutil
from pathlib import Path
from urllib.parse import urlparse
import re
from collections import defaultdict, Counter

def extract_hospital_identifier(url, page_title):
    """URL과 page title에서 병원 식별자를 추출"""
    if pd.isna(url) or pd.isna(page_title):
        return "unknown"
    
    # URL에서 도메인 추출
    try:
        domain = urlparse(str(url)).netloc.lower()
        domain = domain.replace('www.', '')
    except:
        domain = str(url).lower()
    
    # page title에서 병원명 추출 시도
    title = str(page_title).lower()
    
    # 병원명 패턴들
    hospital_patterns = [
        r'([가-힣]+병원)',
        r'([가-힣]+의료원)',
        r'([가-힣]+대학병원)',
        r'([가-힣]+의료센터)',
        r'([가-힣]+센터)',
        r'([a-zA-Z]+\s*hospital)',
        r'([a-zA-Z]+\s*medical\s*center)',
    ]
    
    hospital_name = None
    for pattern in hospital_patterns:
        match = re.search(pattern, title)
        if match:
            hospital_name = match.group(1).strip()
            break
    
    # 병원명이 없으면 도메인을 기준으로
    if not hospital_name:
        if domain:
            hospital_name = domain.split('.')[0]
        else:
            hospital_name = "unknown"
    
    return hospital_name

def main():
    # 경로 설정
    base_path = Path('/Users/gremh/tablemagnifier/collect-data')
    excel_path = base_path / 'Medical_Table_Results.xlsx'
    source_dir = base_path / 'Medical' / 'Context' / 'Origin'
    target_dir = base_path / 'Medical_revised' / 'Context' / 'Origin'
    
    # 타겟 디렉토리 생성
    target_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"Reading Excel file: {excel_path}")
    
    # 엑셀 파일 읽기
    try:
        df = pd.read_excel(excel_path)
        print(f"Excel file loaded with {len(df)} rows")
        print(f"Columns: {df.columns.tolist()}")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    # URL과 page title 컬럼 찾기
    url_col = None
    title_col = None
    
    for col in df.columns:
        col_lower = col.lower()
        if 'url' in col_lower:
            url_col = col
        if 'title' in col_lower or 'page' in col_lower:
            title_col = col
    
    if not url_col or not title_col:
        print(f"Could not find URL column: {url_col} or title column: {title_col}")
        print("Available columns:", df.columns.tolist())
        # 첫 번째와 두 번째 컬럼을 사용
        if len(df.columns) >= 2:
            url_col = df.columns[0]
            title_col = df.columns[1]
            print(f"Using first two columns: {url_col}, {title_col}")
    
    # 병원별로 그룹화
    hospital_groups = defaultdict(list)
    
    for idx, row in df.iterrows():
        url = row.get(url_col, '') if url_col else ''
        title = row.get(title_col, '') if title_col else ''
        
        hospital_id = extract_hospital_identifier(url, title)
        hospital_groups[hospital_id].append({
            'index': idx,
            'url': url,
            'title': title,
            'original_file': f'M_origin_{idx}'
        })
    
    print(f"\nFound {len(hospital_groups)} distinct hospitals:")
    for hospital, items in hospital_groups.items():
        print(f"  {hospital}: {len(items)} items")
    
    # 병원별로 파일 복사 및 재명명
    hospital_counter = 0
    file_counter = defaultdict(int)
    summary_data = []
    
    # 기존 파일들 확인
    existing_files = list(source_dir.glob('M_origin_*'))
    print(f"\nFound {len(existing_files)} files in source directory")
    
    for hospital_name, items in hospital_groups.items():
        print(f"\nProcessing hospital {hospital_counter}: {hospital_name}")
        
        for item in items:
            original_idx = item['index']
            
            # 원본 파일 찾기 (다양한 확장자)
            source_files = list(source_dir.glob(f"M_origin_{original_idx}.*"))
            
            if not source_files:
                print(f"  Warning: No source file found for M_origin_{original_idx}")
                continue
            
            for source_file in source_files:
                # 새 파일명 생성
                new_filename = f"M_origin_{hospital_counter}_{file_counter[hospital_counter]}{source_file.suffix}"
                target_file = target_dir / new_filename
                
                # 파일 복사
                try:
                    shutil.copy2(source_file, target_file)
                    print(f"  Copied: {source_file.name} -> {new_filename}")
                    file_counter[hospital_counter] += 1
                except Exception as e:
                    print(f"  Error copying {source_file.name}: {e}")
        
        summary_data.append({
            'hospital_number': hospital_counter,
            'hospital_name': hospital_name,
            'file_count': file_counter[hospital_counter]
        })
        
        hospital_counter += 1
    
    # 요약 파일 생성
    summary_path = base_path / 'Medical_revised' / 'M_origin_summary.txt'
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    
    total_files = sum(data['file_count'] for data in summary_data)
    
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write("Medical Origin Data Summary\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"Total number of data files: {total_files}\n")
        f.write(f"Number of distinct hospitals: {len(summary_data)}\n\n")
        f.write("Distribution by hospital:\n")
        f.write("-" * 30 + "\n")
        
        for data in summary_data:
            f.write(f"Hospital {data['hospital_number']:2d} ({data['hospital_name']:30s}): {data['file_count']:3d} files\n")
        
        f.write("-" * 30 + "\n")
        f.write(f"Total: {total_files:3d} files\n")
    
    print(f"\n✅ Summary written to: {summary_path}")
    print(f"✅ Reorganization complete!")
    print(f"   - Total files processed: {total_files}")
    print(f"   - Number of hospitals: {len(summary_data)}")
    print(f"   - Output directory: {target_dir}")

if __name__ == "__main__":
    main()