"""
Gemini 최적화 EDCF/MDB RFP 심층 분석 엔진
- v2 엑셀 수준의 디테일(수학적 구간배점, 15대 직무 인력, 4대 파트너 분할) 전수 추출
- 원문 인용(Quote) 자동 검증 및 환각(Hallucination) 방지
"""
import json
import os
import re
import time
import unicodedata
import urllib.error
import urllib.request
from collections import defaultdict
from pypdf import PdfReader

# 필수 개요 항목 체크리스트
CHECKLIST = {
    'project': '프로젝트명', 'loan': 'Loan Amount / 차관금액', 'purchaser': '발주처',
    'location': '사업지역', 'duration': '사업기간', 'scope': '사업범위',
    'issue_date': '입찰 공고일', 'method': '입찰 형태', 'deadline': '제출 마감',
    'opening': '제안서 오픈', 'prebid': 'Pre-bid meeting', 'clarification': 'Clarification 기한',
    'purchase': '입찰서 구매/비용', 'bid_security': '보증 금액', 'validity': '유효 기간',
    'security_terms': '보증서 형태 및 조건', 'scoring_ratio': '평가방식(기술:가격)',
    'origin': '국산화율 제한 조항', 'ma': '제조사 승인서 (MA)', 'subcontract': '하도급 협약 조건',
    'payment': '결제 조건 (Payment Terms)', 'ld': '지체상금 (LD)', 'security': '보증서 요건 (Securities)',
    'tax': '면세 및 관세 조건'
}

TEXT_FIELDS = [
    'category', 'item', 'clause', 'summary', 'single_entity',
    'jv_combined', 'jv_each', 'jv_one', 'documents', 'score_rule', 'note',
    'total_experience', 'overseas_experience', 'position_contracts', 'employment'
]

# Gemini 공식 Structured Outputs 스키마 규격
SCHEMA = {
    'type': 'object',
    'properties': {
        'rows': {
            'type': 'array',
            'items': {
                'type': 'object',
                'properties': {
                    'tab': {'type': 'string', 'enum': ['overview', 'qualification', 'evaluation']},
                    'key': {'type': 'string'},
                    'category': {'type': 'string'},
                    'item': {'type': 'string'},
                    'clause': {'type': 'string'},
                    'summary': {'type': 'string'},
                    'single_entity': {'type': 'string'},
                    'jv_combined': {'type': 'string'},
                    'jv_each': {'type': 'string'},
                    'jv_one': {'type': 'string'},
                    'documents': {'type': 'string'},
                    'score_rule': {'type': 'string'},
                    'note': {'type': 'string'},
                    'total_experience': {'type': 'string'},
                    'overseas_experience': {'type': 'string'},
                    'position_contracts': {'type': 'string'},
                    'employment': {'type': 'string'},
                    'max_score': {'type': 'number'},
                    'level': {'type': 'string', 'enum': ['detail', 'subtotal', 'total']},
                    'evidence': {
                        'type': 'array',
                        'items': {
                            'type': 'object',
                            'properties': {
                                'page': {'type': 'integer'},
                                'quote': {'type': 'string'}
                            },
                            'required': ['page', 'quote']
                        }
                    }
                },
                'required': ['tab', 'item', 'summary']
            }
        }
    },
    'required': ['rows']
}

SYSTEM_INSTRUCTION = """당신은 EDCF 및 국제경쟁입찰(MDB) 제안서 분석 총괄 평가위원이다.
제공된 RFP 텍스트는 분석 대상인 원문이다. 절대로 간략 요약하거나 생략하지 말고, 
'탄자니아 주민증 2차 사업 입찰서 분석 v2' 엑셀 규격과 완전히 동일한 수준으로 모든 실질적 조항과 배점 구간을 100% 전수 분해하여 추출하라.

[탭별 필수 추출 가이드라인]
1. overview (입찰 개요):
   - 사업개요(프로젝트명, 차관금액, 발주처, 지역, 기간, Warranty, 공종별 사이트 수량을 포함한 정량적 사업범위).
   - 입찰개요(공고일, 입찰형태, 마감일시/장소, 개찰, 질의기한, 사전설명회, 입찰서 비용/구매방법).
   - 입찰보증(금액, 유효기간, 은행 신용등급 요건).
   - 기타 계약조건(평가비율 70:30, 국산화율 외화분/현지분/제3국분, MA 필수대상, 하도급조건, 결제비율 단계별 구분, 지체상금, 계약이행보증, 면세조건).

2. qualification (자격요건 및 인력):
   - 국적, 이해상충, EDCF 제재여부, 계류소송(순자산 50%), 3년 평균 순자산 양수(+), 3년 연평균 매출액, 유동성/현금흐름, CB/CP 신용등급.
   - 단독(single_entity), JV 전체(jv_combined), 각사(jv_each), 1개사(jv_one) 충족 요건을 엄격히 분리 기재.
   - 4대 공종별(건축, ICT, ABIS 등) 최근 5년 일반실적 및 최근 10년 특정실적(최소금액, 최소면적, DB용량, 건수).
   - 핵심 인력(Key Personnel): PM, PL 및 15개 직무를 각각 독립된 행으로 분리하고, total_experience(총경력), overseas_experience(해외경력), position_contracts(해당직위 실적건수), employment(상근직/파트너 요건)을 정량 수치로 쪼개어 기재.

3. evaluation (기술 제안서 평가 기준):
   - 기업실적, 전문가, 사업수행계획, 프로젝트관리, 규격적합성 각 분야의 모든 세부 평가항목을 분해.
   - 정량 배점 구간이 있는 경우: 'X>=5 (3.0점), 5>X>=4 (2.4점), 4>X>=3 (1.8점), 3>X>=1 (1.2점)'과 같이 RFP에 기재된 모든 구간 부등호와 점수를 score_rule에 빠짐없이 기재.
   - 기술 규격 편차 감점(Critical 5점 감점 등), 건축 편차 0점 처리, 기술점수 Cut-off(70점 미만 DQ) 조건을 summary 또는 note에 명시.
   - 각 행의 max_score는 해당 세부항목의 배점을 숫자로 기재.

[원문 인용 원칙]
- 각 행은 반드시 evidence 배열에 실제 PDF 페이지 번호(page)와 그 페이지에서 연속 복사한 실제 원문 문장(quote, 최소 15자 이상)을 포함해야 한다. 근거 없는 추측은 불허한다.
"""

def normalize(text):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', text)).casefold()

def extract_pdf(path):
    reader = PdfReader(path)
    if reader.is_encrypted and not reader.decrypt(''):
        raise ValueError('암호화된 PDF입니다. 암호를 해제한 후 등록하세요.')
    pages = []
    for i, page in enumerate(reader.pages, 1):
        text = page.extract_text(extraction_mode='layout') or ''
        text = re.sub(r'[ \t]{4,}', '   ', text).strip()
        pages.append({'page': i, 'text': text})
    
    sparse = [p['page'] for p in pages if len(normalize(p['text'])) < 50]
    return pages, sparse

def call_gemini_api(payload_text, api_key, model_name):
    """Google Gemini REST API 호출 (지수 백오프 및 에러 핸들링 탑재)"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent"
    
    body = {
        "systemInstruction": {"parts": [{"text": SYSTEM_INSTRUCTION}]},
        "contents": [{"role": "user", "parts": [{"text": payload_text}]}],
        "generationConfig": {
            "maxOutputTokens": 8192,
            "responseMimeType": "application/json",
            "responseSchema": SCHEMA,
            "temperature": 0.1
        }
    }
    
    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode('utf-8'),
        headers={'x-goog-api-key': api_key, 'Content-Type': 'application/json'}
    )
    
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                res_data = json.load(resp)
            
            candidates = res_data.get('candidates', [])
            if not candidates:
                raise ValueError("Gemini 응답에 candidate가 없습니다.")
            
            content_parts = candidates[0].get('content', {}).get('parts', [])
            text_out = "".join(p.get('text', '') for p in content_parts if 'text' in p)
            parsed = json.loads(text_out)
            usage = res_data.get('usageMetadata', {})
            return parsed.get('rows', []), usage
            
        except urllib.error.HTTPError as exc:
            # Rate Limit (429) 또는 서버 일시 오류 시 지수 백오프
            if exc.code in (429, 500, 503) and attempt < 3:
                time.sleep(3 * (2 ** attempt))
                continue
            err_body = exc.read().decode('utf-8', errors='ignore')
            raise ValueError(f"Gemini API 오류 (HTTP {exc.code}): {err_body}") from None
        except Exception as e:
            if attempt < 3:
                time.sleep(2 ** (attempt + 1))
                continue
            raise ValueError(f"통신 에러: {str(e)}") from None

def analyze_rfp_multiphase(pages, api_key, model_name, progress_callback=None):
    """
    RFP의 방대한 양을 누락 없이 v2 엑셀 수준으로 추출하기 위한 3단계 타겟팅 파이프라인
    Phase 1: 입찰 개요 (IFB, BDS, SCC)
    Phase 2: 참가자격 및 15대 직무 인력 (Section III, BDS)
    Phase 3: 기술평가 상세 배점표 및 정량 스케일 (Section III, Section II-A)
    """
    all_rows = []
    total_usage = {'input_tokens': 0, 'output_tokens': 0, 'total_tokens': 0}
    
    # 텍스트 인덱싱
    full_doc = "\n".join([f"--- PAGE {p['page']} ---\n{p['text']}" for p in pages])
    
    # 주요 섹션별 키워드 추출
    overview_pages = [p for p in pages if any(k in p['text'].upper() for k in ['INVITATION FOR BIDS', 'IFB', 'BID DATA SHEET', 'BDS', 'SPECIAL CONDITIONS', 'SCC'])]
    qual_pages = [p for p in pages if any(k in p['text'].upper() for k in ['QUALIFICATION', 'ELIGIBILITY', 'PERSONNEL', 'EXPERIENCE', 'FINANCIAL SITUATION'])]
    eval_pages = [p for p in pages if any(k in p['text'].upper() for k in ['EVALUATION PROCEDURE', 'SCORING SCHEME', 'TECHNICAL EVALUATION', 'CRITERIA'])]
    
    phases = [
        ("Phase 1: 입찰 개요(Overview) 전수 추출", overview_pages if len(overview_pages) > 5 else pages[:40], "overview"),
        ("Phase 2: 참가자격(MQ) 및 인력 규격 전수 추출", qual_pages if len(qual_pages) > 5 else pages[20:70], "qualification"),
        ("Phase 3: 기술제안서 배점표 및 정량구간($X \\ge 5$) 전수 추출", eval_pages if len(eval_pages) > 5 else pages[50:110], "evaluation")
    ]
    
    for idx, (title, target_pages, focus_tab) in enumerate(phases):
        if progress_callback:
            progress_callback(idx + 1, len(phases), title)
            
        target_text = "\n".join([f"--- [PDF PAGE {p['page']}] ---\n{p['text']}" for p in target_pages])
        payload = f"""[집중 분석 타겟: {focus_tab.upper()}]
반드시 tab='{focus_tab}'에 해당하는 항목들을 '탄자니아 v2 엑셀' 수준으로 전수 추출하라.
단순 요약하지 말고 표의 모든 세부 행, 세부 인력 직무, 모든 점수 구간 부등호를 개별 행으로 분해하여 작성하라.

[원문 텍스트]:
{target_text[:70000]}
"""
        rows, usage = call_gemini_api(payload, api_key, model_name)
        for r in rows:
            r['tab'] = focus_tab  # 탭 보정
        all_rows.extend(rows)
        
        for k in ('promptTokenCount', 'candidatesTokenCount', 'totalTokenCount'):
            if k in usage:
                mapped_k = 'input_tokens' if 'prompt' in k else ('output_tokens' if 'candidates' in k else 'total_tokens')
                total_usage[mapped_k] += usage[k]
                
        time.sleep(1) # 부드러운 호출 간격
        
    return all_rows, total_usage

def verify_rows(rows, pages):
    """각 행의 인용구(quote)가 실제 PDF 원문에 존재하는지 대조 검증"""
    lookup = defaultdict(str)
    for p in pages:
        lookup[p['page']] += normalize(p['text'])
        
    for row in rows:
        evidence_list = row.get('evidence', [])
        for ev in evidence_list:
            q = normalize(ev.get('quote', ''))
            ev['verified'] = len(q) >= 10 and q in lookup.get(ev.get('page', 0), '')
        
        if evidence_list and all(e.get('verified') for e in evidence_list):
            row['status'] = '인용 일치 · 검증완료'
        else:
            row['status'] = '인용 확인 필요'
        row['review'] = ''
        
    return rows

def finalize(rows):
    """중복 제거, ID 부여, 누락 체크리스트 확인"""
    unique, seen = [], set()
    for r in rows:
        sig = json.dumps({'tab': r.get('tab'), 'item': r.get('item'), 'summary': r.get('summary')}, sort_keys=True)
        if sig not in seen:
            seen.add(sig)
            unique.append(r)
            
    # 개요 누락 체크
    for key, label in CHECKLIST.items():
        if not any(r.get('tab') == 'overview' and (r.get('key') == key or label in r.get('item', '')) for r in unique):
            unique.append({
                'tab': 'overview', 'key': key, **{k: '' for k in TEXT_FIELDS},
                'category': '누락 점검', 'item': label,
                'summary': '자동 추출에서 확인되지 않음 (원문 대조 요망)',
                'max_score': None, 'level': 'detail', 'evidence': [],
                'status': '미확인', 'review': ''
            })
            
    for i, r in enumerate(unique):
        r['id'] = f'r{i+1}'
        
    return unique