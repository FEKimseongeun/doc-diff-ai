[-> EN VERSION README](https://github.com/FEKimseongeun/doc-diff-ai/blob/master/README-EN.md)

# 📑 문서 비교 서비스 (Document Comparison Service)

AI 기반 문서 비교 서비스로, 원본과 개정된 문서(Word, PDF, Excel)를 비교하여 텍스트, 서식, 표, 이미지 등 다양한 변경 사항을 자동으로 식별합니다.

---

## 🚀 주요 기능

- **멀티 포맷 지원**: DOCX, PDF, XLSX 파일 비교 가능  
- **종합적인 변경 감지**:
  - 텍스트 내용 변경(추가, 삭제, 수정)
  - 서식 변경(글꼴, 색상, 스타일, 테두리 등)
  - 표 구조 및 내용 변경
  - 이미지 내용 및 크기 변경
  - 구조적 변경(페이지 수, 시트 수 등)
- **시각적 리포트**: 변경 사항이 하이라이트된 상세 HTML 리포트 생성
- **웹 인터페이스**: 직관적인 웹 애플리케이션 제공
- **보고서 다운로드**: 상세 비교 리포트 다운로드 가능

---

## ⚙️ 설치 방법

### 사전 준비

- Python 3.8 이상  
- pip (Python 패키지 관리자)

### 설치 절차

1. **프로젝트 다운로드**:
   ```bash
   git clone <repository-url>
   cd document-comparison-service
   
2. **가상환경 생성** (권장):
   ```bash
   python -m venv venv
   
   # Windows:
   venv\Scripts\activate
   
   # macOS/Linux:
   source venv/bin/activate

3. **의존성 패키지 설치**:
   ```bash
   pip install -r requirements.txt

4. **애플리케이션 실행**:
   ```bash
   python app.py
5. **웹 인터페이스 접속**:
   브라우저에서 http://localhost:5000 접속

---

## ⚡ 빠른 시작

### Windows
```bash
# start_service.bat 더블클릭 실행
# 또는 수동 실행:
start_service.bat
```
---

## 📘 사용법

### 웹 인터페이스
1. 문서 업로드: 원본과 개정본 문서를 업로드
2. 비교 실행: "Compare Documents" 버튼 클릭
3. 결과 확인: 변경 사항 요약 검토
4. 리포트 다운로드: 하이라이트된 HTML 리포트 다운로드

### 지원 파일 형식
- Microsoft Word (.docx): 텍스트, 서식, 표, 이미지
- PDF (.pdf): 텍스트 및 페이지 구조
- Excel (.xlsx): 셀 값, 서식, 시트 구조

### 파일 크기 제한
- 최대 파일 크기: 50MB (파일당)
- 권장: 10MB 이하 (성능 최적화)

---

## 🌐 API 엔드포인트
### POST /compare
두 문서를 비교하고 변경 분석 결과 반환
**Request:**
- original: 원본 문서 파일
- revised: 개정 문서 파일
**Reponse:**
  ```json
  {
  "success": true,
  "changes": {
    "text_changes": [...],
    "formatting_changes": [...],
    "table_changes": [...],
    "image_changes": [...],
    "structural_changes": [...],
    "summary": {
      "total_changes": 15,
      "text_changes_count": 8,
      "formatting_changes_count": 3,
      "table_changes_count": 2,
      "image_changes_count": 1,
      "structural_changes_count": 1
      }
    },
    "report_path": "comparison_report_20231201_143022.html"
  }

### GET /download_report/<filename>
  생성된 비교 리포트 다운로드

---

## 🏗️ 아키텍처
### 핵심 컴포넌트
1. DocumentParser: 다양한 문서 포맷에서 콘텐츠 추출
2. ChangeDetector: 문서 간 변경 사항 식별
3. ReportGenerator: HTML/JSON 리포트 생성
4. Web Interface: Flask 기반 웹 애플리케이션

###처리 파이프라인
1. 파싱: 텍스트, 서식, 표, 이미지 추출
2. 비교: 차이점 분석
3. 리포트 생성: 종합 보고서 작성
4. 시각화: 사용자 친화적 인터페이스 제공

### 🔎 변경 감지 알고리즘
- 텍스트 변경: difflib 기반 라인별 비교-
- 서식 변경: 글꼴, 크기, 색상, 스타일, 테두리 비교
- 표 변경: 행/열 구조 및 셀 내용 비교
- 이미지 변경: SSIM(Structural Similarity Index) 활용, 크기/속성 비교

---

## 📂 프로젝트 구조
```
document-comparison-service/
├── app.py                 # Flask 웹 애플리케이션
├── document_parser.py     # 문서 파싱 로직
├── change_detector.py     # 변경 감지 알고리즘
├── report_generator.py    # 리포트 생성
├── templates/             # 웹 UI 템플릿
│   └── index.html
├── static/                # 정적 리소스
├── uploads/               # 임시 파일 저장
├── reports/               # 리포트 저장
├── requirements.txt       # 의존성 패키지
├── start_service.bat      # Windows 실행 스크립트
├── start_service.sh       # Unix/Linux/Mac 실행 스크립트
└── README.md
```

---

## 📌 향후 개선 예정
- PPTX, RTF 등 추가 포맷 지원
- 배치 처리 기능
- API 인증 및 속도 제한
- 클라우드 스토리지 연동
- 실시간 협업 기능
- 고급 시각화 기능
- 머신러닝 기반 변경 분류

  
