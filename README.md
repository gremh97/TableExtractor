# Table Magnifier

의료 웹사이트에서 테이블 데이터를 자동으로 추출하고 시각화하는 도구입니다.

## 기능

- **웹페이지 테이블 자동 추출**: Selenium을 이용한 동적 웹페이지 테이블 캡처
- **HTML 패널 테이블 추출**: JavaScript로 숨겨진 테이블이 있는 단일페이지 사이트 지원
- **PDF 테이블 추출**: pdfplumber를 이용한 PDF 내 테이블 정확한 영역 추출
- **한글 폰트 지원**: 웹브라우저 스타일 렌더링으로 한글 텍스트 완벽 지원
- **엑셀 데이터베이스**: 모든 추출 결과를 Excel 파일로 체계적 관리

## 설치 및 사용법

### 1. 환경 설정

```bash
# 가상환경 생성
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows

# 필요한 패키지 설치
pip install -r requirements.txt
```

### 2. 웹페이지 테이블 추출

```bash
# URL 목록을 urls.txt에 추가하고 실행
python continuous_table_extractor.py
```

### 3. PDF 테이블 추출

```bash
# temperal_pdf 폴더에 PDF 파일 추가하고 실행
python pdf_processor_pdfplumber.py
```

## 출력 파일 구조

```
Medical/
├── Context/Origin/          # 원본 파일 저장소
│   ├── M_origin_*.png      # 전체 웹페이지 스크린샷
│   ├── M_origin_*.pdf      # PDF 원본 파일
│   └── M_table_*.html      # 추출된 테이블 HTML 원본
├── Table/                  # 테이블 이미지 저장소
│   └── M_table_*.png       # 개별 테이블 캡처 이미지
└── Medical_Table_Results.xlsx  # 통합 데이터베이스
```

## 지원하는 사이트 유형

### 1. 일반 웹페이지
- Selenium WebDriver를 사용한 동적 콘텐츠 처리
- 테이블 자동 감지 및 개별 캡처

### 2. HTML 패널 방식 사이트
- `davoshospital.co.kr` 등 단일페이지 내 숨겨진 테이블
- `<div class="panel">` 블록 내 테이블 직접 추출
- requests + BeautifulSoup4를 이용한 빠른 처리

### 3. PDF 문서
- pdfplumber + pdf2image를 이용한 정확한 테이블 영역 감지
- 300 DPI 고해상도 테이블 이미지 생성

## 주요 개선사항

- ✅ **한글 폰트 완벽 지원**: "Malgun Gothic", "맑은 고딕" 폰트 우선 적용
- ✅ **웹브라우저 스타일 렌더링**: 실제 웹페이지와 동일한 테이블 모양 보존
- ✅ **정확한 테이블 영역 추출**: 컬럼 잘림 현상 해결
- ✅ **SSL 인증서 문제 해결**: 인증서 오류가 있는 사이트도 처리 가능

## 파일 설명

- `continuous_table_extractor.py`: 메인 웹페이지 테이블 추출 도구
- `pdf_processor_pdfplumber.py`: PDF 테이블 추출 도구
- `urls.txt`: 처리할 URL 목록
- `Medical_Table_Results.xlsx`: 통합 결과 데이터베이스

## 요구사항

- Python 3.7+
- Chrome 브라우저 (Selenium WebDriver용)
- 필요한 Python 패키지는 requirements.txt 참조

## 라이센스

MIT License