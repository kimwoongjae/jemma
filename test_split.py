import glob
import re
import pandas as pd

def parse_bundle(product_name, option_name, qty):
    items = []
    qty = int(qty) if pd.notna(qty) and str(qty).isdigit() else 1
    if pd.isna(option_name) or not str(option_name).strip():
        return [{'sub_product': product_name, 'sub_option': '기본', 'sub_qty': qty}]
    
    opt_str = str(option_name).strip()
    
    # 1. '+' 로 묶인 세트 분리
    if '+' in opt_str:
        parts = [p.strip() for p in opt_str.split('+') if p.strip()]
        for p in parts:
            # 혹시 파트 내에 수량 표시가 있는지 체크 (예: PEARL 2개)
            m = re.search(r'(\d+)\s*(?:개|box|ea)?$', p, re.IGNORECASE)
            sub_q = qty
            item_name = p
            items.append({
                'sub_product': product_name,
                'sub_option': item_name,
                'sub_qty': sub_q
            })
    else:
        # 단일 옵션이지만 혹시 숫자로만 되어 있는 경우 (예: 옵션 4 -> 4개입 등)
        items.append({
            'sub_product': product_name,
            'sub_option': opt_str,
            'sub_qty': qty
        })
    return items

for f in glob.glob('*.xlsx'):
    print(f"=== TESTING FILE: {f} ===")
    df = pd.read_excel(f)
    print(f"Total Rows: {len(df)}")
    expanded_rows = []
    for idx, row in df.iterrows():
        split_items = parse_bundle(row.get('상품명', ''), row.get('옵션이름', ''), row.get('상품수량', 1))
        for item in split_items:
            expanded_rows.append({
                '주문번호': row.get('주문번호'),
                '수취인': row.get('수취인이름'),
                '상품명': item['sub_product'],
                '옵션': item['sub_option'],
                '수량': item['sub_qty']
            })
    print(f"Original: {len(df)} rows -> Expanded: {len(expanded_rows)} item rows")
    # 샘플 3개 출력
    print("Sample expanded items:")
    for ex in expanded_rows[:4]:
        print(f"  {ex['주문번호']} | {ex['수취인']} | {ex['상품명']} [{ex['옵션']}] x {ex['수량']}")
    print()
