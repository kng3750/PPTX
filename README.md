# PPTX 템플릿 · 문서 삽입

PPTX 템플릿 1장과 문서 파일을 업로드하면, 템플릿의 배경·머릿말은 그대로 두고 **본문 영역에만** 문서 내용을 삽입합니다.

## 지원 형식

- **템플릿**: `.pptx`
- **문서**: `.txt`, `.docx`, `.doc`

## 설치 및 실행

```bash
# 가상환경 생성 및 활성화 (권장)
python -m venv venv
venv\Scripts\activate   # Windows

# 패키지 설치
pip install -r requirements.txt

# 실행
python app.py
```

브라우저에서 http://127.0.0.1:5000 접속

## 사용 방법

1. PPTX 템플릿 파일(1장) 업로드
2. 삽입할 문서(.txt / .docx / .doc) 업로드
3. **결과 PPTX 생성** 클릭
4. 생성된 `result.pptx` 다운로드

## 템플릿 조건

- 슬라이드에 **본문용 placeholder**(텍스트 상자)가 있어야 합니다.
- 일반적인 PowerPoint 레이아웃의 "본문" 영역이 자동으로 인식됩니다.
- 타이틀/머릿말 placeholder는 변경되지 않습니다.
