# 붙임1. 소프트웨어사업 하도급 계약 승인신청서 자동 생성기

`사업대금 계산식_예시.xlsx` 엑셀 파일을 업로드하면 `계약금액정보` 및 `투입인력 총괄표` 시트를 자동 분석하여, 「소프트웨어 진흥법 시행규칙 [별지 제9호서식]」 기반 **붙임1. 하도급 계약 승인신청서(.docx)**를 생성하는 웹 애플리케이션입니다.

## 주요 기능
- **엑셀 자동 분석**: 사업명, 총 계약금액, 주사업자 계약금액, 하도급 업체별 금액(B), 비율(B/A), 투입인력 및 담당업무 자동 추출
- **계약금액(A) 기준 전환**: 전체 계약금액 기준 / 주사업자 분담금액 기준 원클릭 전환 및 실시간 비율 재계산
- **웹 UI 편집**: 수급인·하수급인 정보, 계약기간, 하도급 내역 및 사유를 화면에서 바로 수정
- **워드(.docx) 생성**: 업체별 개별 다운로드 및 선택 업체 전체 일괄 ZIP 다운로드 지원

## Render.com 배포 방법
1. 이 폴더의 파일들을 GitHub 저장소(Repository)에 업로드합니다.
2. [Render.com](https://render.com)에 로그인 후 **New + → Web Service** (또는 **Blueprint**)를 선택합니다.
3. GitHub 저장소를 연결합니다. (`render.yaml`이 포함되어 있어 Blueprint 배포 시 설정이 자동 적용됩니다.)
4. 수동 Web Service로 설정할 경우 아래 값을 입력합니다.
   - **Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT`
5. **Create Web Service**를 클릭하면 배포가 완료되고 웹 URL이 발급됩니다.

## 로컬 실행 방법
```bash
pip install -r requirements.txt
python app.py
```
브라우저에서 `http://localhost:5000` 으로 접속합니다.
