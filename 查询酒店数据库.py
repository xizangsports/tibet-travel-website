#!/usr/bin/env python3
"""西藏体旅·酒店数据库查询工具"""
import sqlite3, sys, os

db = r'G:\CODEX HERMES\西藏体旅官网\hotels.db'

if not os.path.exists(db):
    print(f"❌ 数据库不存在: {db}")
    sys.exit(1)

conn = sqlite3.connect(db)
cur = conn.cursor()

def print_table(rows, cols):
    if not rows:
        print("  无结果")
        return
    # Print header
    header = ' | '.join(cols)
    print(f"  {'='*len(header)}")
    print(f"  {header}")
    print(f"  {'='*len(header)}")
    for r in rows:
        vals = [str(v)[:25] if v else '—' for v in r]
        print(f"  {' | '.join(vals)}")
    print(f"  ({len(rows)} 条记录)")

def main():
    while True:
        print("\n" + "="*50)
        print("🏔️ 西藏体旅·酒店数据库查询")
        print("="*50)
        print("1. 按城市查看酒店")
        print("2. 按价格区间查看")
        print("3. 搜索酒店（关键词）")
        print("4. 查看有2026年价格的酒店")
        print("5. 统计概览")
        print("0. 退出")
        
        choice = input("\n请选择 (0-5): ").strip()
        
        if choice == '1':
            cur.execute("SELECT DISTINCT city FROM hotels ORDER BY city")
            cities = [r[0] for r in cur.fetchall()]
            print("\n城市列表:")
            for i, c in enumerate(cities, 1):
                cnt = cur.execute("SELECT COUNT(*) FROM hotels WHERE city=?", (c,)).fetchone()[0]
                print(f"  {i}. {c} ({cnt}家)")
            sel = input("\n选择城市编号: ").strip()
            try:
                city = cities[int(sel)-1]
                rows = cur.execute('''SELECT name, star, room_type, price_low, price_mid, price_high, phone, price_tier 
                                       FROM hotels WHERE city=? ORDER BY 
                                       CASE WHEN has_2026_price=1 THEN 0 ELSE 1 END,
                                       price_tier''', (city,)).fetchall()
                print_table(rows, ['酒店名称','星级','房型','淡季','平季','旺季','电话','价格区间'])
            except:
                print("无效选择")

        elif choice == '2':
            print("\n价格区间:")
            tiers = ['暂无价格','💰 经济(≤300)','💰💰 中端(300-500)','💰💰💰 中高端(500-1000)','💰💰💰💰 高端(1000+)']
            for i, t in enumerate(tiers, 1):
                cnt = cur.execute("SELECT COUNT(*) FROM hotels WHERE price_tier=?", (t,)).fetchone()[0]
                print(f"  {i}. {t} ({cnt}家)")
            sel = input("\n选择: ").strip()
            try:
                tier = tiers[int(sel)-1]
                rows = cur.execute('''SELECT city, name, star, room_type, price_low, phone
                                       FROM hotels WHERE price_tier=? ORDER BY city''', (tier,)).fetchall()
                print_table(rows, ['城市','酒店','星级','房型','最低价','电话'])
            except:
                print("无效选择")

        elif choice == '3':
            kw = input("\n输入关键词: ").strip()
            rows = cur.execute('''SELECT city, name, star, phone, price_tier
                                   FROM hotels WHERE name LIKE ? OR city LIKE ? OR address LIKE ?
                                   ORDER BY city''', (f'%{kw}%', f'%{kw}%', f'%{kw}%')).fetchall()
            print_table(rows, ['城市','酒店','星级','电话','价格区间'])

        elif choice == '4':
            rows = cur.execute('''SELECT city, name, star, room_type, price_low, price_mid, price_high, phone
                                   FROM hotels WHERE has_2026_price=1 ORDER BY city, price_tier''').fetchall()
            print_table(rows, ['城市','酒店','星级','房型','淡季','平季','旺季','电话'])

        elif choice == '5':
            print("\n📊 统计概览:")
            cur.execute('''SELECT city, COUNT(*), SUM(has_2026_price), 
                           SUM(CASE WHEN phone NOT IN ('','None','—') THEN 1 ELSE 0 END)
                           FROM hotels GROUP BY city ORDER BY COUNT(*) DESC''')
            for r in cur.fetchall():
                print(f"  {r[0]}: {r[1]}家 (有2026价:{r[2]}, 有电话:{r[3]})")
            print(f"  总计: {cur.execute('SELECT COUNT(*) FROM hotels').fetchone()[0]} 条")

        elif choice == '0':
            break

    conn.close()

if __name__ == '__main__':
    main()
