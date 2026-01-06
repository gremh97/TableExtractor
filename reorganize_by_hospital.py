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
from PIL import Image, ImageChops
import numpy as np

class UnionFind:
    """Union-Find 자료구조로 중복 그룹을 찾기 위한 클래스"""
    def __init__(self):
        self.parent = {}
        self.rank = {}
    
    def find(self, x):
        if x not in self.parent:
            self.parent[x] = x
            self.rank[x] = 0
        if self.parent[x] != x:
            self.parent[x] = self.find(self.parent[x])
        return self.parent[x]
    
    def union(self, x, y):
        px, py = self.find(x), self.find(y)
        if px == py:
            return
        if self.rank[px] < self.rank[py]:
            px, py = py, px
        self.parent[py] = px
        if self.rank[px] == self.rank[py]:
            self.rank[px] += 1
    
    def get_groups(self):
        groups = defaultdict(list)
        for x in self.parent:
            groups[self.find(x)].append(x)
        return [group for group in groups.values() if len(group) > 1]

def images_are_identical(img1_path, img2_path):
    """두 이미지가 시각적으로 동일한지 비교 (유연한 기준)"""
    try:
        # 이미지 열기
        img1 = Image.open(img1_path)
        img2 = Image.open(img2_path)
        
        # 크기가 다르면 동일하지 않음
        if img1.size != img2.size:
            return False
        
        # 이미지 모드가 다르면 RGB로 변환
        if img1.mode != img2.mode:
            img1 = img1.convert('RGB')
            img2 = img2.convert('RGB')
        
        # 픽셀별 차이 계산
        diff = ImageChops.difference(img1, img2)
        
        # numpy 배열로 변환하여 더 정밀한 비교
        import numpy as np
        diff_array = np.array(diff)
        
        # 전체 픽셀 대비 다른 픽셀의 비율 계산
        total_pixels = diff_array.size
        if total_pixels == 0:
            return True
            
        # 0이 아닌 픽셀(차이가 있는 픽셀) 개수
        different_pixels = np.count_nonzero(diff_array)
        similarity_ratio = 1.0 - (different_pixels / total_pixels)
        
        # 99.5% 이상 유사하면 동일한 이미지로 판정 (기존 기준)
        threshold = 0.995
        is_similar = similarity_ratio >= threshold
        
        # 디버깅 정보 (필요시)
        if is_similar and different_pixels > 0:
            print(f"    📊 Similar images found: {img1_path.name} ≈ {img2_path.name} (similarity: {similarity_ratio:.4f})")
        
        return is_similar
        
    except Exception as e:
        print(f"Error comparing {img1_path} and {img2_path}: {e}")
        return False

def calculate_image_hash(img_path):
    """이미지의 해시값을 계산하여 빠른 비교 지원"""
    try:
        img = Image.open(img_path)
        # 작은 크기로 리사이즈하여 해시 계산
        img = img.resize((8, 8), Image.Resampling.LANCZOS).convert('L')
        pixels = list(img.getdata())
        
        # 평균값 기준으로 이진화
        avg = sum(pixels) / len(pixels)
        bits = ''.join('1' if pixel >= avg else '0' for pixel in pixels)
        return int(bits, 2)
    except:
        return None

def are_images_perceptually_similar(img1_path, img2_path):
    """퍼셉션 해시를 이용한 빠른 유사도 검사"""
    hash1 = calculate_image_hash(img1_path)
    hash2 = calculate_image_hash(img2_path)
    
    if hash1 is None or hash2 is None:
        return False
    
    # 해밍 거리 계산 (다른 비트 수)
    xor = hash1 ^ hash2
    hamming_distance = bin(xor).count('1')
    
    # 해밍 거리가 5 이하면 유사한 이미지로 판정
    return hamming_distance <= 5

def extract_hospital_number_from_filename(filename):
    """파일명에서 병원번호 추출"""
    # M_origin_0_1.png -> 0
    # M_table_0_1_0.png -> 0
    match = re.search(r'M_(origin|table)_(\d+)_', filename)
    if match:
        return int(match.group(2))
    return None

def find_and_remove_duplicates():
    """병원별로 동일한 이미지 찾기 및 중복 제거"""
    
    # 디렉토리 설정
    origin_dir = Path('/Users/gremh/tablemagnifier/collect-data/Medical_revised/Context/Origin')
    table_dir = Path('/Users/gremh/tablemagnifier/collect-data/Medical_revised/Table')
    
    print("🔍 Analyzing images for visual duplicates within same hospital groups...")
    print("=" * 80)
    
    # 병원별로 파일들 그룹화
    hospital_files = defaultdict(lambda: {'origin': [], 'table': []})
    
    # Origin 파일들 수집
    if origin_dir.exists():
        for file_path in origin_dir.glob('*.png'):
            hospital_num = extract_hospital_number_from_filename(file_path.name)
            if hospital_num is not None:
                hospital_files[hospital_num]['origin'].append(file_path)
    
    # Table 파일들 수집
    if table_dir.exists():
        for file_path in table_dir.glob('*.png'):
            hospital_num = extract_hospital_number_from_filename(file_path.name)
            if hospital_num is not None:
                hospital_files[hospital_num]['table'].append(file_path)
    
    total_duplicates_found = 0
    files_to_remove = []
    
    # 각 병원별로 중복 검사
    for hospital_num in sorted(hospital_files.keys()):
        files = hospital_files[hospital_num]
        origin_files = files['origin']
        table_files = files['table']
        
        print(f"\n  Hospital {hospital_num}: {len(origin_files)} origins, {len(table_files)} tables")
        
        duplicates_in_hospital = []
        uf = UnionFind()  # Union-Find for grouping duplicates
        
        # Origin 파일들끼리 비교
        if len(origin_files) > 1:
            print(f"    🔍 Checking {len(origin_files)} origin files for duplicates...")
            for i in range(len(origin_files)):
                for j in range(i + 1, len(origin_files)):
                    # 먼저 빠른 해시 비교
                    if are_images_perceptually_similar(origin_files[i], origin_files[j]):
                        # 해시가 유사하면 정밀 비교
                        if images_are_identical(origin_files[i], origin_files[j]):
                            duplicates_in_hospital.append(('origin', origin_files[i], origin_files[j]))
                            uf.union(origin_files[i].name, origin_files[j].name)
        
        # Table 파일들끼리 비교
        if len(table_files) > 1:
            print(f"    🔍 Checking {len(table_files)} table files for duplicates...")
            for i in range(len(table_files)):
                for j in range(i + 1, len(table_files)):
                    # 먼저 빠른 해시 비교
                    if are_images_perceptually_similar(table_files[i], table_files[j]):
                        # 해시가 유사하면 정밀 비교
                        if images_are_identical(table_files[i], table_files[j]):
                            duplicates_in_hospital.append(('table', table_files[i], table_files[j]))
                            uf.union(table_files[i].name, table_files[j].name)
        
        # 중복 발견 시 처리
        if duplicates_in_hospital:
            print(f"    🔍 Found {len(duplicates_in_hospital)} duplicate pairs")
            
            # 집합 형태로 중복 그룹 처리
            duplicate_groups = uf.get_groups()
            if duplicate_groups:
                for i, group in enumerate(duplicate_groups, 1):
                    # 중복 제거: 각 그룹에서 하나만 남기고 나머지 삭제
                    if len(group) > 1:
                        # 첫 번째 파일을 남기고 나머지를 삭제 목록에 추가
                        sorted_group = sorted(group)
                        keep_file = sorted_group[0]
                        remove_files = sorted_group[1:]
                        
                        print(f"      🗑️  Keep: {keep_file}, Remove: {', '.join(remove_files)}")
                        
                        for filename in remove_files:
                            # origin 또는 table 디렉토리에서 파일 찾기
                            origin_path = origin_dir / filename
                            table_path = table_dir / filename
                            
                            if origin_path.exists():
                                files_to_remove.append(origin_path)
                            elif table_path.exists():
                                files_to_remove.append(table_path)
                
                total_duplicates_found += len(duplicates_in_hospital)
        else:
            print(f"    ✅ No duplicates found")
    
    # 실제 파일 삭제 수행
    if files_to_remove:
        print(f"\n🗑️  Removing {len(files_to_remove)} duplicate files...")
        for file_path in files_to_remove:
            try:
                file_path.unlink()
                print(f"    Removed: {file_path.name}")
            except Exception as e:
                print(f"    Error removing {file_path.name}: {e}")
        
        print(f"✅ Duplicate removal completed! Removed {len(files_to_remove)} files.")
    
    print("\n" + "=" * 80)
    if total_duplicates_found > 0:
        print(f"📊 Total duplicate pairs found and removed: {total_duplicates_found}")
    else:
        print("✅ No visual duplicates found within hospital groups!")

def extract_hospital_identifier(url, page_title, index):
    """URL과 page title에서 병원 식별자를 추출"""
    
    # 특정 인덱스 범위에 대한 수동 매핑 (먼저 체크)
    if index == 47:
        return "hospital_47"  # 독립 병원
    elif index == 48:
        return "hospital_48"  # 독립 병원
    elif index in [49, 50, 51]:
        return "davoshospital"  # 같은 병원
    elif index in [52, 53, 54, 55]:
        return "knmc"  # 같은 병원
    elif index == 56:
        return "knmc"  # knmc 병원에 포함
    
    if pd.isna(url) or pd.isna(page_title):
        return f"unknown_{index}"  # 각각 독립적인 unknown으로 처리
    
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
            hospital_name = f"unknown_{index}"
    
    return hospital_name

def renumber_files_after_duplicate_removal(origin_dir, table_dir):
    """중복 제거 후 파일들을 연속적으로 재넘버링"""
    import re
    from collections import defaultdict
    
    print("🔄 Renumbering files after duplicate removal...")
    print("=" * 60)
    
    # 병원별 파일들 수집
    hospital_files = defaultdict(lambda: {'origin': [], 'table': []})
    
    # Origin 파일들 수집
    for file_path in origin_dir.glob('M_origin_*.png'):
        match = re.search(r'M_origin_(\d+)_(\d+)', file_path.name)
        if match:
            hospital_num = int(match.group(1))
            hospital_files[hospital_num]['origin'].append(file_path)
    
    # Table 파일들 수집
    for file_path in table_dir.glob('M_table_*.png'):
        match = re.search(r'M_table_(\d+)_(\d+)_(\d+)', file_path.name)
        if match:
            hospital_num = int(match.group(1))
            hospital_files[hospital_num]['table'].append(file_path)
    
    # 각 병원별로 재넘버링
    total_renamed = 0
    for hospital_num in sorted(hospital_files.keys()):
        origin_files = sorted(hospital_files[hospital_num]['origin'], 
                            key=lambda x: int(re.search(r'M_origin_(\d+)_(\d+)', x.name).group(2)))
        table_files = sorted(hospital_files[hospital_num]['table'],
                           key=lambda x: (int(re.search(r'M_table_(\d+)_(\d+)_(\d+)', x.name).group(2)),
                                        int(re.search(r'M_table_(\d+)_(\d+)_(\d+)', x.name).group(3))))
        
        print(f"  Hospital {hospital_num}: {len(origin_files)} origins, {len(table_files)} tables")
        
        # Origin 파일들 재넘버링
        for new_origin_idx, file_path in enumerate(origin_files):
            old_name = file_path.name
            new_name = f"M_origin_{hospital_num}_{new_origin_idx}.png"
            
            if old_name != new_name:
                new_path = file_path.parent / new_name
                file_path.rename(new_path)
                print(f"    Origin: {old_name} -> {new_name}")
                total_renamed += 1
        
        # Table 파일들을 원본별로 그룹화하여 재넘버링
        table_by_origin = defaultdict(list)
        for file_path in table_files:
            match = re.search(r'M_table_(\d+)_(\d+)_(\d+)', file_path.name)
            if match:
                origin_idx = int(match.group(2))
                table_by_origin[origin_idx].append(file_path)
        
        # 새로운 origin 인덱스에 맞춰 table 재넘버링
        origin_mapping = {}
        for new_origin_idx, old_file_path in enumerate(origin_files):
            # 이전 origin 번호 추출 (파일명이 이미 변경된 경우 고려)
            try:
                match = re.search(r'M_origin_(\d+)_(\d+)', old_file_path.name)
                if match:
                    old_origin_idx = int(match.group(2))
                else:
                    old_origin_idx = new_origin_idx
            except:
                old_origin_idx = new_origin_idx
            origin_mapping[old_origin_idx] = new_origin_idx
        
        # 각 원본별로 테이블 재넘버링
        for old_origin_idx, tables in table_by_origin.items():
            if old_origin_idx in origin_mapping:
                new_origin_idx = origin_mapping[old_origin_idx]
                sorted_tables = sorted(tables, 
                                     key=lambda x: int(re.search(r'M_table_(\d+)_(\d+)_(\d+)', x.name).group(3)))
                
                for new_table_idx, file_path in enumerate(sorted_tables):
                    old_name = file_path.name
                    new_name = f"M_table_{hospital_num}_{new_origin_idx}_{new_table_idx}.png"
                    
                    if old_name != new_name:
                        new_path = file_path.parent / new_name
                        file_path.rename(new_path)
                        print(f"    Table: {old_name} -> {new_name}")
                        total_renamed += 1
    
    print("=" * 60)
    print(f"✅ Renumbering completed! Renamed {total_renamed} files.")
    print("   All files now have continuous numbering within each hospital group.")

def main():
    # 경로 설정
    base_path = Path('/Users/gremh/tablemagnifier/collect-data')
    excel_path = base_path / 'Medical_Table_Results.xlsx'
    source_dir = base_path / 'Medical' / 'Context' / 'Origin'
    source_tables_dir = base_path / 'Medical' / 'tables_refined'
    target_dir = base_path / 'Medical_revised' / 'Context' / 'Origin'
    target_tables_dir = base_path / 'Medical_revised' / 'Table'
    
    # 타겟 디렉토리 생성
    target_dir.mkdir(parents=True, exist_ok=True)
    target_tables_dir.mkdir(parents=True, exist_ok=True)
    
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
        
        hospital_id = extract_hospital_identifier(url, title, idx)
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
                    
                    # 해당 원본에 대한 테이블 이미지들도 복사
                    original_table_dir = source_tables_dir / f"M_origin_{original_idx}"
                    if original_table_dir.exists():
                        table_files = list(original_table_dir.glob("M_table_*.png"))
                        for table_idx, table_file in enumerate(table_files):
                            # 새로운 테이블 파일명: M_table_{병원번호}_{원본번호}_{테이블인덱스}
                            new_table_filename = f"M_table_{hospital_counter}_{file_counter[hospital_counter]}_{table_idx}.png"
                            target_table_file = target_tables_dir / new_table_filename
                            
                            try:
                                shutil.copy2(table_file, target_table_file)
                                print(f"    Table copied: {table_file.name} -> {new_table_filename}")
                            except Exception as e:
                                print(f"    Error copying table {table_file.name}: {e}")
                    
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
    
    # 중복 제거 및 재정리
    print(f"\n🔍 Checking for duplicates...")
    find_and_remove_duplicates()
    
    # 파일 재넘버링
    print(f"\n🔄 Renumbering files after duplicate removal...")
    renumber_files_after_duplicate_removal(target_dir, target_tables_dir)
    
    print(f"\n🎉 Complete workflow finished!")

if __name__ == "__main__":
    main()