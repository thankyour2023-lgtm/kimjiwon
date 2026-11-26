import json, io
import numpy as np
import openpyxl

# 분류 수준별로 읽어 둔 투입산출표를 메모리에 보관한다.
# 키: 'basic'(기본부문 380) / 'small'(소분류 165) / 'mid'(중분류 83) / 'large'(대분류 33)
MODELS = {}

# 상품분류표에서 읽은 부문 연결 정보
# MAP_UP[기본부문코드] = {'basic':.., 'small':.., 'mid':.., 'large':..}
MAP_UP = {}
# NAME_OF[(level, code)] = 부문명 (분류표 기준)
NAME_OF = {}

LEVELS = ['basic', 'small', 'mid', 'large']
LEVEL_KR = {'basic': '기본부문', 'small': '소분류', 'mid': '중분류', 'large': '대분류'}

# 최종수요 항목 코드 → (이름, 묶음)  묶음: 소비 / 투자 / 수출
FD_ITEMS = [
    ('9111', '민간소비지출', '소비'),
    ('9112', '정부소비지출', '소비'),
    ('9121', '민간고정자본형성', '투자'),
    ('9122', '정부고정자본형성', '투자'),
    ('9131', '재고증감', '투자'),
    ('9132', '귀중품 순취득', '투자'),
    ('9140', '수출', '수출'),
]


def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)


def _is_sector(code):
    # 산업(상품) 부문코드만 True. 합계·부가가치 등 요약행은 9로 시작한다.
    c = str(code)
    return c != 'None' and not c.startswith('9')


def load_mapping(raw):
    # 상품분류.xlsx: 4열 쌍(기본부문, 소분류, 중분류, 대분류)이 이어지며
    # 하위 분류의 코드가 바뀌는 행에만 상위 코드/이름이 채워져 있다.
    global MAP_UP, NAME_OF
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.values)
    wb.close()
    MAP_UP = {}
    NAME_OF = {}
    cur = [None, None, None, None]        # [기본, 소, 중, 대] 코드
    order = ['basic', 'small', 'mid', 'large']
    # 데이터는 4번째 행(인덱스 3)부터 시작: 코드/상품명이 2칸씩 8열
    for r in rows[3:]:
        if r is None:
            continue
        changed = False
        for k in range(4):
            code = r[2 * k] if len(r) > 2 * k else None
            name = r[2 * k + 1] if len(r) > 2 * k + 1 else None
            if code is not None and str(code).strip():
                cur[k] = str(code).strip()
                if name is not None and str(name).strip():
                    NAME_OF[(order[k], cur[k])] = str(name).strip()
                changed = True
        if cur[0] is not None and changed and r[0] is not None:
            MAP_UP[cur[0]] = {order[k]: cur[k] for k in range(4)}
    if len(MAP_UP) < 300:
        raise ValueError('분류 연결표를 읽지 못했습니다(기본부문 %d개).' % len(MAP_UP))
    return emit({'mapped': len(MAP_UP)})


def _matrix(book, sheet, codes):
    rows = book[sheet]
    ri = {str(r[0]): i for i, r in enumerate(rows) if r and r[0] is not None}
    ci = {str(c): i for i, c in enumerate(rows[4])}
    a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
    if not np.isfinite(a).all():
        raise ValueError(sheet + ': 비수치 또는 결측 데이터')
    return a


def initialize(level, raw):
    # 한 분류 수준의 워크북을 읽고 계수 항등식을 검증한 뒤 메모리에 저장한다.
    if level not in LEVELS:
        raise ValueError('알 수 없는 분류 수준: ' + str(level))
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)

    base = book['국산투입계수(Ad)']
    codes = [str(c) for c in base[4][2:] if _is_sector(c)]
    if len(codes) != len(set(codes)) or len(codes) < 10:
        raise ValueError(level + ': 부문코드 구조가 예상과 다릅니다.')
    nm = {str(r[0]): str(r[1]) for r in base[6:] if r and r[0] is not None}
    names = [nm.get(c, c) for c in codes]
    n = len(codes)

    L = _matrix(book, '생산유발계수', codes)
    M = _matrix(book, '수입유발계수', codes)
    V = _matrix(book, '부가가치유발계수', codes)
    Ad = _matrix(book, '국산투입계수(Ad)', codes)
    Am = _matrix(book, '수입투입계수(Am)', codes)

    # 부가가치계(9690) 행 — 검증용
    arows = book['총투입계수(A)']
    aci = {str(c): i for i, c in enumerate(arows[4])}
    vr = next(r for r in arows if r and str(r[0]) == '9690')
    v = np.array([vr[aci[c]] for c in codes], dtype=float)

    residuals = {
        '생산': float(np.max(np.abs((np.eye(n) - Ad) @ L - np.eye(n)))),
        '수입': float(np.max(np.abs(Am @ L - M))),
        '부가가치': float(np.max(np.abs(v[:, None] * L - V))),
    }
    if max(residuals.values()) >= 1e-6:
        raise ValueError('계수 검증 실패: ' + str(residuals))

    # 영향력계수(열합/평균)·감응도계수(행합/평균) 사전 계산
    col = L.sum(axis=0)
    row = L.sum(axis=1)
    influence = col / col.mean()
    sensitivity = row / row.mean()

    # 최종수요 항목별 생산유발액 시트를 코드→행 벡터로 보관
    fd_sheet = book['최종수요 항목별 생산유발액']
    fci = {str(c): i for i, c in enumerate(fd_sheet[4])}
    fd_row = {}
    for r in fd_sheet[6:]:
        if r and r[0] is not None and _is_sector(r[0]):
            fd_row[str(r[0])] = {code: (float(r[fci[code]]) if code in fci and isinstance(r[fci[code]], (int, float)) else 0.0)
                                 for code, _, _ in FD_ITEMS}

    MODELS[level] = {
        'book': book, 'sheets': sheets, 'codes': codes, 'names': names,
        'index': {c: i for i, c in enumerate(codes)},
        'L': L, 'M': M, 'V': V,
        'influence': influence, 'sensitivity': sensitivity,
        'prod_mult': col, 'imp_coef': M.sum(axis=0), 'va_coef': V.sum(axis=0),
        'fd_row': fd_row,
    }
    return emit({
        'level': level, 'levelKr': LEVEL_KR[level],
        'sectors': [{'code': c, 'name': nm.get(c, c)} for c in codes],
        'sheets': sheets, 'residuals': residuals,
    })


def _need(level):
    if level not in MODELS:
        raise ValueError(LEVEL_KR.get(level, level) + ' 자료가 아직 로딩되지 않았습니다.')
    return MODELS[level]


def explore(level, sheet, rowcode):
    m = _need(level)
    rows = m['book'][sheet]
    choices = [{'code': str(r[0]), 'name': str(r[1])}
               for r in rows[6:] if r and len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if r and str(r[0]) == rowcode), None)
    if chosen is None:
        chosen = next(r for r in rows[6:] if r and r[0] is not None and r[1] is not None)
    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None:
            continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)):
            val = float(val)
        elif val is not None:
            val = str(val)
        cells.append({'code': str(rows[4][i]), 'name': str(rows[5][i] or ''), 'value': val})
    unit = str(rows[3][0] or '계수').replace('係数', '계수')
    if '단위' not in unit:
        unit = '계수 / 비율 (단위 없음). 행·열 명칭으로 의미를 확인하세요.'
    return emit({'rows': choices, 'selected': str(chosen[0]), 'name': str(chosen[1]),
                 'unit': unit, 'cells': cells})


def simulate(level, payload):
    m = _need(level)
    codes, L, M, V = m['codes'], m['L'], m['M'], m['V']
    shocks = json.loads(payload)
    f = np.zeros(len(codes))
    for s in shocks:
        a = float(s['amount'])
        if not np.isfinite(a) or a < 0:
            raise ValueError('수요 증가액은 유한한 0 이상의 수여야 합니다.')
        f[m['index'][s['code']]] += a
    if not f.sum() > 0:
        raise ValueError('한 산업 이상의 양의 수요 증가액을 입력하세요.')
    x = L @ f; imp = M @ f; va = V @ f
    result = [{'code': c, 'name': n, 'production': float(a), 'imports': float(b), 'valueAdded': float(d)}
              for c, n, a, b, d in zip(codes, m['names'], x, imp, va)]
    result.sort(key=lambda r: r['production'], reverse=True)
    return emit({'demand': float(f.sum()), 'production': float(x.sum()), 'imports': float(imp.sum()),
                 'valueAdded': float(va.sum()), 'multiplier': float(x.sum() / f.sum()), 'rows': result})


# ---- 분석 A·B·C·D ----

def analysis(which, args_json):
    args = json.loads(args_json) if args_json else {}
    if which == 'A':
        return _analysis_A(args)
    if which == 'B':
        return _analysis_B(args)
    if which == 'C':
        return _analysis_C(args)
    if which == 'D':
        return _analysis_D(args)
    raise ValueError('알 수 없는 분석: ' + str(which))


def _code_at(level, basic_code):
    # 기본부문 코드를 해당 분류 수준의 코드로 올려 매핑
    if basic_code not in MAP_UP:
        raise ValueError('분류표에 없는 기본부문 코드: ' + str(basic_code))
    return MAP_UP[basic_code][level]


def _analysis_A(args):
    # 같은 산업을 4개 분류 수준에서 비교: 생산유발계수(승수), 수입유발계수, 부가가치유발계수
    basic_code = str(args['basic'])
    missing = [lv for lv in LEVELS if lv not in MODELS]
    if missing:
        return emit({'need': missing})
    rows = []
    for lv in LEVELS:
        m = MODELS[lv]
        code = _code_at(lv, basic_code)
        i = m['index'][code]
        rows.append({
            'level': lv, 'levelKr': LEVEL_KR[lv], 'code': code,
            'name': m['names'][i], 'sectors': len(m['codes']),
            'prod': float(m['prod_mult'][i]),
            'imp': float(m['imp_coef'][i]),
            'va': float(m['va_coef'][i]),
        })
    return emit({'basic': basic_code, 'rows': rows})


def _analysis_B(args):
    # 영향력계수 × 감응도계수 산점도. 선택 산업(기본부문 기준) 강조.
    level = args['level']
    m = _need(level)
    sel_code = None
    if args.get('basic'):
        try:
            sel_code = _code_at(level, str(args['basic']))
        except Exception:
            sel_code = None
    pts = [{'code': c, 'name': m['names'][i],
            'infl': float(m['influence'][i]), 'sens': float(m['sensitivity'][i])}
           for i, c in enumerate(m['codes'])]
    return emit({'level': level, 'levelKr': LEVEL_KR[level], 'points': pts, 'selected': sel_code})


def _analysis_C(args):
    # 선택 산업의 생산이 어떤 최종수요(소비/투자/수출)에서 유발되는가
    level = args['level']
    m = _need(level)
    code = _code_at(level, str(args['basic'])) if args.get('basic') else args['code']
    fd = m['fd_row'].get(str(code))
    if fd is None:
        raise ValueError('최종수요 유발 자료에 없는 부문: ' + str(code))
    items = []
    groups = {'소비': 0.0, '투자': 0.0, '수출': 0.0}
    total = 0.0
    for c, label, grp in FD_ITEMS:
        v = fd.get(c, 0.0)
        items.append({'code': c, 'label': label, 'group': grp, 'value': v})
        groups[grp] += v
        total += v
    i = m['index'][str(code)]
    return emit({'level': level, 'levelKr': LEVEL_KR[level], 'code': str(code),
                 'name': m['names'][i], 'total': total,
                 'groups': [{'label': k, 'value': groups[k]} for k in ['소비', '투자', '수출']],
                 'items': items})


def _analysis_D(args):
    # 해당 산업 최종수요(수출 등) 1,000백만원 증가 시 국내 부가가치 vs 수입 누출
    level = args['level']
    m = _need(level)
    code = _code_at(level, str(args['basic'])) if args.get('basic') else args['code']
    i = m['index'][str(code)]
    amount = float(args.get('amount', 1000))
    prod = float(m['prod_mult'][i]) * amount
    imp = float(m['imp_coef'][i]) * amount
    va = float(m['va_coef'][i]) * amount
    return emit({'level': level, 'levelKr': LEVEL_KR[level], 'code': str(code),
                 'name': m['names'][i], 'amount': amount,
                 'production': prod, 'imports': imp, 'valueAdded': va,
                 'impCoef': float(m['imp_coef'][i]), 'vaCoef': float(m['va_coef'][i]),
                 'prodMult': float(m['prod_mult'][i])})
