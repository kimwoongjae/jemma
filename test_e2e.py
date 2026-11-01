import requests
import json

BASE_URL = 'http://127.0.0.1:5050'


def test_full_pipeline():
    print("1. Testing Dashboard Stats API...")
    r = requests.get(f"{BASE_URL}/api/dashboard/stats")
    print("  Stats response:", r.json())

    print("\n2. Testing Excel Upload Preview with '(프리비_엘모사기미스틱)배송리스트1028.xlsx'...")
    with open('(프리비_엘모사기미스틱)배송리스트1028.xlsx', 'rb') as f:
        files = {'file': ('(프리비_엘모사기미스틱)배송리스트1028.xlsx', f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        r = requests.post(f"{BASE_URL}/api/upload/preview", files=files)
    preview = r.json()
    print(f"  Preview rows: {preview['total_excel_rows']}, Split items: {preview['total_split_items']}")
    
    print("\n3. Testing Confirm Upload...")
    r = requests.post(f"{BASE_URL}/api/upload/confirm", json=preview)
    print("  Confirm result:", r.json())

    print("\n4. Testing Orders List API...")
    r = requests.get(f"{BASE_URL}/api/orders?limit=5")
    orders = r.json()
    print(f"  Fetched {len(orders)} sample orders. First order split items:")
    for it in orders[0]['items']:
        print(f"    - {it['item_product_name']} [{it['item_option_name']}] x {it['quantity']}")

    first_order_id = orders[0]['id']
    print(f"\n5. Testing Issue Tracking for Order ID {first_order_id}...")
    r = requests.post(f"{BASE_URL}/api/shipping/issue-tracking", json={'order_ids': [first_order_id]})
    print("  Tracking issued:", r.json())

    print(f"\n6. Testing Shipping Labels API for Order ID {first_order_id}...")
    r = requests.get(f"{BASE_URL}/api/shipping/labels?order_ids={first_order_id}")
    labels = r.json()
    print("  Label data courier:", labels[0]['courier'], "Tracking:", labels[0]['tracking_no'])

    print(f"\n7. Testing Dispatch & Stock Deduction for Order ID {first_order_id}...")
    r = requests.post(f"{BASE_URL}/api/shipping/dispatch", json={'order_ids': [first_order_id]})
    print("  Dispatch result:", r.json())

    print("\n8. Testing Inventory List...")
    r = requests.get(f"{BASE_URL}/api/inventory")
    prods = r.json()
    print(f"  Total products in WMS: {len(prods)}")
    for p in prods[:3]:
        print(f"    - [{p['sku']}] {p['product_name']} ({p['option_name']}) Stock: {p['stock_quantity']}")

    print("\n9. Final Dashboard Stats:")
    r = requests.get(f"{BASE_URL}/api/dashboard/stats")
    print("  Final Stats:", r.json())

if __name__ == '__main__':
    test_full_pipeline()
