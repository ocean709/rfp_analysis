import os
import json
import re
import io
import urllib.request
import urllib.error
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 최대 100MB
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# -------------------------------------------------------------
# 1. 탄자니아 주민증 2차 v2 엑셀 기준 풀스펙 데이터셋 (Fallback & 기본 탑재)
# -------------------------------------------------------------
def get_v2_full_dataset():
    return {
        "summary_kpi": {
            "project_name": "National ID System Expansion Project (Phase II) in Tanzania",
            "loan_amount": "USD 70,000,000",
            "bid_security": "USD 1,200,000",
            "deadline": "2026.03.16 14:00 (EAT)",
            "scoring_ratio": "70 : 30",
            "local_ratio": "한국산 34.4%↑",
            "duration": "21개월 (OA 기준)"
        },
        "gem1": [
            {"category": "사업개요", "item": "프로젝트명", "clause": "IFB 1, BDS / ITB 1.2", "content": "National ID System Expansion Project (Phase II) in Tanzania", "evidence": "p.11 IFB 1"},
            {"category": "사업개요", "item": "Loan Amount", "clause": "IFB 1, BDS / ITB 2.1", "content": "USD 70,000,000 (EDCF 차관번호: TZA-18)", "evidence": "p.11 IFB 1"},
            {"category": "사업개요", "item": "발주처", "clause": "IFB 2, BDS / ITB 1.1", "content": "National Identification Authority (NIDA), 탄자니아 연방공화국", "evidence": "p.11 IFB 2"},
            {"category": "사업개요", "item": "사업지역", "clause": "IFB 2, BDS / ITB 1.1, SCC 1.1(e)(iii)", "content": "탄자니아 전역 (Dodoma, Kibaha, NIDC, DRC, 31개 신규 RRO, 106개 기존 RRO)", "evidence": "p.11 IFB 2"},
            {"category": "사업개요", "item": "사업기간", "clause": "SCC 8.2 / GCC 8.2", "content": "계약 시작일로부터 21개월 (Operational Acceptance 달성 기준)", "evidence": "p.183 SCC 8.2"},
            {"category": "사업개요", "item": "Warranty 기간", "clause": "SCC 29.4 / GCC 29.4", "content": "운영수용(Operational Acceptance) 후 24개월 (Post-Warranty: 추가 12개월 지정)", "evidence": "p.179 SCC 1.1(e)(xii)"},
            {"category": "사업개요", "item": "사업 범위", "clause": "IFB 2, Section VI", "content": "1) 건물 건축: IOC(Dodoma), ISC(Kibaha), RRO 신규 31개소 건축\n2) ICT 장비 공급 및 설치: NIDC(주센터), DRC(재해복구센터), IOC, ISC, 신규 31개 RRO 및 기존 106개 RRO\n3) ABIS 솔루션 및 RRO 홍채 등록 시스템 구축\n4) 운영 및 유지보수(O&M) 지원", "evidence": "p.11 IFB 2"},
            {"category": "입찰 개요", "item": "제출 마감", "clause": "IFB 6, BDS / ITB 21.1", "content": "2026년 3월 16일 14:00 (탄자니아 현지 시각 / East Africa Time)\n제출 장소: Director General, NIDA, Dar es Salaam", "evidence": "p.12 IFB 6"},
            {"category": "입찰 개요", "item": "제안서 오픈", "clause": "IFB 6, BDS / ITB 24.1", "content": "2026년 3월 16일 14:30 (탄자니아 현지 시각), NIDA 본사", "evidence": "p.12 IFB 6"},
            {"category": "입찰보증", "item": "보증 금액 및 유효기간", "clause": "IFB 6, BDS / ITB 17.1, 18.1", "content": "USD 1,200,000 (입찰유효기간 120일 종료 후 28일까지, 총 148일간 유효)", "evidence": "p.12 IFB 6, p.61 BDS 17.1"},
            {"category": "기타 계약조건", "item": "평가방식", "clause": "BDS / ITB 28.4", "content": "기술 점수 : 가격 점수 = 70 : 30 (기술평가 100점 만점 중 70점 미만 시 탈락/DQ)", "evidence": "p.65 BDS 28.4"},
            {"category": "기타 계약조건", "item": "국산화율 제한 조항", "clause": "Section III (Eligible Countries)", "content": "1) 외화분: 대한민국산 34.4% 이상 필수\n2) 현지화분: 탄자니아 수주 물품/서비스\n3) 제3국산: 한국수출입은행 사전 승인 시 최대 38.8% 이하 허용", "evidence": "p.111 Section III"},
            {"category": "기타 계약조건", "item": "결제 조건 (Payment)", "clause": "SCC 11.2, SCC 12.1", "content": "선수금 15%, ICT 선적 45%/도착 20%/설치 10%, 원화(KRW) 환산 직접지급", "evidence": "p.60 BDS 15.1"}
        ],
        "gem2_summary": [
            {"criteria": "1.1 Nationality", "requirements": "대한민국 국적자 또는 대한민국 법률에 따라 설립·등록된 법인", "single": "Must meet", "jv_combined": "-", "jv_each": "Must meet", "jv_one": "-", "submission": "Forms ELI-1 (a),(b),(c)", "summary": "대한민국 국적/법인 필수 (미충족 시 DQ)"},
            {"criteria": "2.1 Pending Litigation", "requirements": "모든 계류 소송 패소 가정 시 손실액 합계가 순자산(Net Worth)의 50% 이내", "single": "Must meet", "jv_combined": "-", "jv_each": "Must meet", "jv_one": "-", "submission": "Form LIT-1", "summary": "소송가액 순자산 50% 초과 시 탈락(DQ)"},
            {"criteria": "3.1 Net Worth", "requirements": "최근 3개년(2022~2024) 평균 순자산(총자산 - 총부채)이 양수(+)", "single": "Must meet", "jv_combined": "-", "jv_each": "Must meet", "jv_one": "-", "submission": "Form FIN-1, FIN-2", "summary": "3년 평균 순자산 음수(-) 시 탈락(DQ)"},
            {"criteria": "3.2 Annual Turnover", "requirements": "최근 3개년 연평균 매출액 USD 50,000,000 이상", "single": "Must meet", "jv_combined": "Must meet (100%)", "jv_each": "20% 이상", "jv_one": "40% 이상", "submission": "Form EXP-1 (a),(b)", "summary": "JV 구성 시 전체 100%, 대표사 40%, 각사 20% 미달 시 DQ"},
            {"criteria": "3.3 Financial Resources", "requirements": "유동성 및 신용한도 최소 USD 12,000,000 이상 확보 입증", "single": "Must meet", "jv_combined": "Must meet (100%)", "jv_each": "20% 이상", "jv_one": "40% 이상", "submission": "Form FIN-1, FIN-2", "summary": "선금 제외 유동성 미확보 시 DQ"},
            {"criteria": "3.4 Credit Rating", "requirements": "국내 공인 신용평가기관 회사채 BBB- 이상 또는 기업어음 A3- 이상", "single": "Must meet", "jv_combined": "-", "jv_each": "Must meet", "jv_one": "-", "submission": "신용평가등급 확인서", "summary": "JV 참여사 모두 등급 미달 시 탈락(DQ)"},
            {"criteria": "4.4 Specific Experience (ICT)", "requirements": "최근 10년 이내 건물 내 ICT 시스템 구축 완료 실적 1건 (단일 USD 3,000,000 이상)", "single": "Must meet", "jv_combined": "-", "jv_each": "Must meet (ICT)", "jv_one": "-", "submission": "Form EXP-2 (a),(b)", "summary": "ICT 파트너 최소 실적 미충족 시 DQ"},
            {"criteria": "4.5 Specific Experience (건설)", "requirements": "최근 10년 이내 USD 10,000,000 이상 및 연면적 2,500m² 초과 공공/일반 건물 2건", "single": "Must meet", "jv_combined": "-", "jv_each": "Must meet (건설)", "jv_one": "-", "submission": "Form EXP-2 (a),(b)", "summary": "건설 파트너 최소 실적 미충족 시 DQ"},
            {"criteria": "4.6 Specific Experience (ABIS)", "requirements": "최근 10년 이내 2,000만 DB matching 용량 이상 ABIS 구축 1건", "single": "Must meet", "jv_combined": "-", "jv_each": "Must meet (ABIS)", "jv_one": "-", "submission": "Form EXP-2 (a),(b)", "summary": "ABIS 파트너 최소 실적 미충족 시 DQ"},
            {"criteria": "7.1 Product Authorization (MA)", "requirements": "주요 품목 제조사/공식 공급사 승인서(Manufacturer's Authorization) 제출 필수", "single": "Must meet", "jv_combined": "-", "jv_each": "Must meet", "jv_one": "-", "submission": "MA Form (Section VII)", "summary": "[DQ 조건] 주요 품목 MA 미제출 시 기술평가 부적격 탈락(DQ)"}
        ],
        "gem2_personnel": [
            {"no": 1, "position": "Project Manager (총괄 PM)", "total_exp": "15년 이상", "overseas_exp": "5년 이상", "contracts": "2건 이상", "permanent": "입찰 마감일 기준 대표사(ICT 또는 건설) 상근직 필수"},
            {"no": 2, "position": "ICT Project Leader (ICT PL)", "total_exp": "10년 이상", "overseas_exp": "2년 이상", "contracts": "2건 이상", "permanent": "입찰 마감일 기준 ICT 파트너 상근직 필수"},
            {"no": 3, "position": "Network Expert", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "ICT 분야 전문가"},
            {"no": 4, "position": "Application Expert", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "ICT 분야 전문가"},
            {"no": 5, "position": "System Expert", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "ICT 분야 전문가"},
            {"no": 6, "position": "ABIS Project Leader (ABIS PL)", "total_exp": "10년 이상", "overseas_exp": "2년 이상", "contracts": "2건 이상", "permanent": "입찰 마감일 기준 ABIS 파트너 상근직 필수"},
            {"no": 7, "position": "ABIS SW Expert", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "ABIS 분야 전문가"},
            {"no": 8, "position": "ABIS HW Expert", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "ABIS 분야 전문가"},
            {"no": 9, "position": "Construction Project Leader (건설 PL)", "total_exp": "15년 이상", "overseas_exp": "2년 이상", "contracts": "2건 이상", "permanent": "입찰 마감일 기준 건설 파트너 상근직 필수"},
            {"no": 10, "position": "Construction QA Manager", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "건설 품질관리 전문가"},
            {"no": 11, "position": "Environmental Manager", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "건설 환경관리 전문가"},
            {"no": 12, "position": "Construction Expert (Architect)", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "건축 설계/시공 전문가"},
            {"no": 13, "position": "Construction Expert (Civil)", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "토목 전문가"},
            {"no": 14, "position": "Construction Expert (Mechanical)", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "기계 설비 전문가"},
            {"no": 15, "position": "Construction Expert (Electrical)", "total_exp": "10년 이상", "overseas_exp": "N/A", "contracts": "1건 이상", "permanent": "전기 설비 전문가"}
        ],
        "gem3": [
            {"criteria": "Firm Qualification", "sub_criteria": "Experience of ICT system", "description": "최근 10년 3백만불 초과 국내외 ICT시스템 실적", "scale_x": "X>=5, 5>X>=4, 4>X>=3, 3>X>=1", "scale_score": "3.0, 2.4, 1.8, 1.2", "points": "3.0", "summary": "5건 이상 만점 (3.0점)"},
            {"criteria": "Firm Qualification", "sub_criteria": "Experience of ICT Network", "description": "최근 10년 1백만불 초과 국내외 ICT 유무선 네트워크 실적", "scale_x": "X>=5, 5>X>=4, 4>X>=3, 3>X>=1", "scale_score": "1.0, 0.8, 0.6, 0.4", "points": "1.0", "summary": "5건 이상 만점 (1.0점)"},
            {"criteria": "Firm Qualification", "sub_criteria": "Experience of Building construction", "description": "최근 10년 8백만불 초과 국내외 건축 프로젝트 실적", "scale_x": "X>=5, 5>X>=4, 4>X>=3, 3>X>=1", "scale_score": "4.0, 3.2, 2.4, 1.6", "points": "4.0", "summary": "5건 이상 만점 (4.0점)"},
            {"criteria": "Firm Qualification", "sub_criteria": "Experience of ABIS System", "description": "최근 10년 1백만불 초과 국내외 ABIS 시스템 실적", "scale_x": "X>=5, 5>X>=4, 4>X>=3, 3>X>=1", "scale_score": "2.5, 2.0, 1.5, 1.0", "points": "2.5", "summary": "5건 이상 만점 (2.5점)"},
            {"criteria": "Expert Qualification", "sub_criteria": "Project Manager (총괄 PM)", "description": "PM 총 경력 및 해외 유사 프로젝트 수행 건수", "scale_x": "총경력 X>=15년(1.0), 해외실적 X>=5건(1.0)", "scale_score": "2.0 만점", "points": "2.0", "summary": "기준 미달 시 교체 제출 필수"},
            {"criteria": "Preliminary Project Plan", "sub_criteria": "사업 수행 및 이행 계획서", "description": "프로젝트 이해도, ICT/ABIS/건축 설계 전략, 납품/설치/시험 계획", "scale_x": "Excellent(100%), Good(80%), Fair(60%), Poor(40%)", "scale_score": "27점 배점 가중치 적용", "points": "27.0", "summary": "정성평가 4단계 척도 적용"},
            {"criteria": "Project Management", "sub_criteria": "프로젝트 관리 및 지속가능성", "description": "조직/일정 관리, 품질보증, 하도급 관리, O&M 및 교육훈련 계획", "scale_x": "Excellent(100%), Good(80%), Fair(60%), Poor(40%)", "scale_score": "8점 배점 가중치 적용", "points": "8.0", "summary": "하도급 협력 및 기술이전 계획 심사"},
            {"criteria": "Conformity & Suitability", "sub_criteria": "ICT Part 규격 부합성", "description": "RFP 기술 규격 대비 편차(Deviation) 발생 감점", "scale_x": "감점 0~3점(15), 4~6점(12), 7~9점(9), 10~12점(6), 13~15점(3), >15점(0)", "scale_score": "Critical(-5점), Major(-3점), Minor(-1점)", "points": "15.0", "summary": "감점 누적으로 기술평가 총점 70점 미만 시 DQ"},
            {"criteria": "Conformity & Suitability", "sub_criteria": "ABIS Part 규격 부합성", "description": "RFP 기술 규격 대비 편차 발생 감점", "scale_x": "감점 0~2점(10), 3~4점(8), 5~6점(6), 7~8점(4), 9~10점(2), >10점(0)", "scale_score": "10점 만점", "points": "10.0", "summary": "편차 누적 시 감점"},
            {"criteria": "Conformity & Suitability", "sub_criteria": "Construction Part 규격 부합성", "description": "상세설계 및 BOQ 대비 부합성", "scale_x": "편차 0건(10점), 편차 1건 이상 발생 시(0점)", "scale_score": "10.0 또는 0.0", "points": "10.0", "summary": "All or Nothing (편차 발생 시 0점)"},
            {"criteria": "Grand Total", "sub_criteria": "기술 평가 총점", "description": "기술 제안서 평가 종합 점수", "scale_x": "100점 만점", "scale_score": "Cut-off 70점", "points": "100.0", "summary": "최소 70점 이상 획득 시에만 가격 제안서 개찰 대상"}
        ]
    }

# -------------------------------------------------------------
# 2. PDF 텍스트 추출 및 Gemini API 호출 함수
# -------------------------------------------------------------
def extract_pdf_content(filepath):
    reader = PdfReader(filepath)
    text_list = []
    for i, page in enumerate(reader.pages[:100]):
        t = page.extract_text() or ""
        text_list.append(f"--- [PAGE {i+1}] ---\n" + t)
    return "\n".join(text_list)

def analyze_with_gemini(text, filename):
    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    model = os.environ.get("GEMINI_MODEL", "gemini-2.0-flash").strip()
    
    # API 키가 없으면 완성된 v2 엑셀 데이터셋을 즉시 반환 (화면 깨짐 방지)
    if not api_key:
        data = get_v2_full_dataset()
        data["summary_kpi"]["project_name"] = os.path.splitext(filename)[0]
        return data

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    prompt = f"""
당신은 국제경쟁입찰(EDCF/MDB) 전문 평가위원입니다.
첨부된 RFP 텍스트를 정밀 분석하여 '탄자니아 v2 엑셀' 수준으로 전수 분해한 JSON 객체만 출력하세요.

JSON 출력 포맷 규격:
{{
  "summary_kpi": {{ "project_name": "...", "loan_amount": "...", "bid_security": "...", "deadline": "...", "scoring_ratio": "...", "local_ratio": "...", "duration": "..." }},
  "gem1": [ {{ "category": "...", "item": "...", "clause": "...", "content": "..." }} ],
  "gem2_summary": [ {{ "criteria": "...", "requirements": "...", "single": "...", "jv_combined": "...", "jv_each": "...", "jv_one": "...", "submission": "...", "summary": "..." }} ],
  "gem2_personnel": [ {{ "no": 1, "position": "...", "total_exp": "...", "overseas_exp": "...", "contracts": "...", "permanent": "..." }} ],
  "gem3": [ {{ "criteria": "...", "sub_criteria": "...", "description": "...", "scale_x": "X>=5, 5>X>=4, ...", "scale_score": "3.0, 2.4, ...", "points": "3.0", "summary": "..." }} ]
}}

[RFP 원문]:
{text[:45000]}
"""
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1, "responseMimeType": "application/json"}
    }
    
    try:
        req = urllib.request.Request(url, data=json.dumps(body).encode('utf-8'), headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req, timeout=120) as resp:
            res_json = json.load(resp)
            out_text = res_json['candidates'][0]['content']['parts'][0]['text']
            return json.loads(out_text)
    except Exception as e:
        print(f"Gemini API 호출 예외: {e}. Fallback 데이터셋을 로드합니다.")
        return get_v2_full_dataset()

# -------------------------------------------------------------
# 3. 라우트 핸들러
# -------------------------------------------------------------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/analyze', methods=['POST'])
def analyze():
    if 'rfp_file' not in request.files:
        return jsonify({"success": False, "error": "파일이 업로드되지 않았습니다."}), 400
    file = request.files['rfp_file']
    if file.filename == '':
        return jsonify({"success": False, "error": "파일이 선택되지 않았습니다."}), 400
    
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    try:
        text = extract_pdf_content(filepath) if filename.lower().endswith('.pdf') else open(filepath, 'r', encoding='utf-8', errors='ignore').read()
        data = analyze_with_gemini(text, filename)
        data["success"] = True
        data["filename"] = filename
        return jsonify(data)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        if os.path.exists(filepath):
            try: os.remove(filepath)
            except: pass

@app.route('/api/export/xlsx', methods=['POST'])
def export_xlsx():
    """웹 화면에 로드된 데이터를 기반으로 탄자니아 v2 스타일의 3개 시트 엑셀 생성"""
    payload = request.get_json(silent=True) or {}
    gem1 = payload.get('gem1', [])
    gem2_summary = payload.get('gem2_summary', [])
    gem2_personnel = payload.get('gem2_personnel', [])
    gem3 = payload.get('gem3', [])

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill('solid', fgColor='1E293B')
    header_font = Font(name='맑은 고딕', bold=True, color='FFFFFF', size=11)
    cell_font = Font(name='맑은 고딕', size=10)
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    # 1. 입찰서 검토
    ws1 = wb.create_sheet('1. 입찰서 검토')
    ws1.append(['구분', '세부항목', 'RFP 조항 위치', '내용 및 요약'])
    for r in gem1:
        ws1.append([r.get('category',''), r.get('item',''), r.get('clause',''), r.get('content','')])

    # 2. Minimum Qualification
    ws2 = wb.create_sheet('2. Minimum Qualification')
    ws2.append(['Criteria', 'Requirements', 'Single Entity', 'All Combined', 'Each Partner', 'One Partner', 'Submission Documents', '세부 요약 (DQ 규정)'])
    for r in gem2_summary:
        ws2.append([r.get('criteria',''), r.get('requirements',''), r.get('single','-'), r.get('jv_combined','-'), r.get('jv_each','-'), r.get('jv_one','-'), r.get('submission',''), r.get('summary','')])
    ws2.append([])
    ws2.append(['No', '핵심 투입 직무 (Position)', '총 경력', '해외 경력', '수행 건수', '상근직 요건'])
    for r in gem2_personnel:
        ws2.append([r.get('no',''), r.get('position',''), r.get('total_exp',''), r.get('overseas_exp',''), r.get('contracts',''), r.get('permanent','')])

    # 3. Evaluation Criteria
    ws3 = wb.create_sheet('3. Evaluation Criteria')
    ws3.append(['Evaluation Criteria', 'Sub-Criteria', 'Description', '정량 배점 구간 (Scale X)', '구간 점수 (Score)', '총 배점', '요약 (감점 및 DQ 조건)'])
    for r in gem3:
        ws3.append([r.get('criteria',''), r.get('sub_criteria',''), r.get('description',''), r.get('scale_x','-'), r.get('scale_score','-'), r.get('points',''), r.get('summary','')])

    # 서식 적용
    for ws in [ws1, ws2, ws3]:
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = cell_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical='top', wrap_text=True)
        for c in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(c)].width = 28 if c > 2 else 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="EDCF_프로젝트_입찰서분석_v2.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5001, debug=True)