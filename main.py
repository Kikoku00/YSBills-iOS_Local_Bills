import os
import csv
import sqlite3
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
DB_FILE = os.path.join(BASE_DIR, "account.db")
CSV_FILE = os.path.join(BASE_DIR, "export_bill.csv")

TIME_FMT = "%Y-%m-%d %H:%M:%S"



def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS bills (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            category TEXT NOT NULL,
            amount REAL NOT NULL,
            note TEXT DEFAULT '',
            created_at TEXT
        )
    ''')
    conn.commit()
    conn.close()



def migrate_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("PRAGMA table_info(bills)")
    cols = [row[1] for row in c.fetchall()]
    if "created_at" not in cols:
        c.execute("ALTER TABLE bills ADD COLUMN created_at TEXT")
        conn.commit()
        print("🔄 数据库已升级：新增 created_at 时间字段", flush=True)
    conn.close()



def add_bill(bill_type, category, amount, note):
    now = datetime.now().strftime(TIME_FMT)
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        "INSERT INTO bills (type, category, amount, note, created_at) VALUES (?, ?, ?, ?, ?)",
        (bill_type, category, amount, note, now)
    )
    conn.commit()
    conn.close()
    print(f"✅记录添加完成（时间：{now}）", flush=True)



def show_all():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, type, category, amount, note, created_at FROM bills ORDER BY id")
    rows = c.fetchall()
    conn.close()
    if not rows:
        print("暂无记账记录！", flush=True)
        return
    print("\n====全部记账记录====", flush=True)
    print("序号｜时间｜收支类型｜分类｜金额｜备注", flush=True)
    for r in rows:
        t = r[5] if r[5] else "-"
        print(f"{r[0]}｜{t}｜{r[1]}｜{r[2]}｜{r[3]:.2f}｜{r[4]}", flush=True)



def calculate():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT SUM(amount) FROM bills WHERE type='收入'")
    income = c.fetchone()[0] or 0
    c.execute("SELECT SUM(amount) FROM bills WHERE type='支出'")
    expend = c.fetchone()[0] or 0
    conn.close()
    print("\n====收支统计结果====", flush=True)
    print(f"总收入：{income:.2f}", flush=True)
    print(f"总支出：{expend:.2f}", flush=True)
    print(f"当前结余：{income - expend:.2f}", flush=True)



def filter_by_category():
    cat = input("请输入要查询的分类(如学习、餐饮)：").strip()
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        "SELECT id, type, category, amount, note, created_at FROM bills WHERE category=? ORDER BY id",
        (cat,)
    )
    rows = c.fetchall()
    conn.close()
    print(f"\n====分类【{cat}】的账目====", flush=True)
    if not rows:
        print("没有找到该分类的记录", flush=True)
        return
    for r in rows:
        t = r[5] if r[5] else "-"
        print(f"{r[0]}｜{t}｜{r[1]}｜{r[2]}｜{r[3]:.2f}｜{r[4]}", flush=True)



def delete_item():
    show_all()
    num = input("\n输入需要删除的记录序号：").strip()
    if not num.isdigit():
        print("序号输入无效", flush=True)
        return
    bill_id = int(num)
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM bills WHERE id=?", (bill_id,))
    conn.commit()
    if c.rowcount > 0:
        print("删除成功！", flush=True)
    else:
        print("序号不存在", flush=True)
    conn.close()



def export_csv():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT type, category, amount, note, created_at FROM bills ORDER BY id")
    rows = c.fetchall()
    conn.close()
    if not rows:
        print("暂无记录，无法导出！", flush=True)
        return
    with open(CSV_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["收支类型", "分类", "金额", "备注", "时间"])
        for row in rows:
            writer.writerow(row)
    print(f"✅成功导出文件：{CSV_FILE}", flush=True)
    print("提示：CSV文件可以用Excel/WPS打开，多数记账软件支持导入", flush=True)



def _read_csv_rows(path):
    """优先 UTF-8，失败则尝试 GBK（兼容 Excel 另存的 CSV）"""
    last_err = None
    for enc in ("utf-8-sig", "gbk"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return list(csv.reader(f)), enc
        except UnicodeDecodeError as e:
            last_err = e
    raise ValueError(f"无法识别文件编码（{last_err}），请另存为 UTF-8 编码的 CSV")


def _looks_like_header(row):
    """判断第一行是否为表头：同时出现‘收支类型/type’和‘金额/amount’才认为是表头"""
    text = " ".join(str(cell).strip().lower() for cell in row[:4])
    has_type = ("收支类型" in text) or ("type" in text)
    has_amount = ("金额" in text) or ("amount" in text)
    return has_type and has_amount


def _parse_time(s):
    """尝试把字符串解析为标准时间格式，失败返回 None"""
    s = s.strip()
    if not s:
        return None
    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d",
        "%Y/%m/%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).strftime(TIME_FMT)
        except ValueError:
            continue
    return None


def _print_import_errors(errors):
    for line_no, reason in errors[:10]:
        print(f"  第{line_no}行：{reason}", flush=True)
    if len(errors) > 10:
        print(f"  ... 其余 {len(errors) - 10} 条略", flush=True)


def import_csv():
    path = input(f"请输入CSV文件路径（直接回车使用 {os.path.basename(CSV_FILE)}）：").strip()
    if not path:
        path = CSV_FILE
    elif not os.path.isabs(path):
        path = os.path.join(BASE_DIR, path)

    if not os.path.exists(path):
        print(f"❌ 文件不存在：{path}", flush=True)
        return

    try:
        raw_rows, enc = _read_csv_rows(path)
    except Exception as e:
        print(f"❌ 读取文件失败：{e}", flush=True)
        return

    start = 1 if raw_rows and _looks_like_header(raw_rows[0]) else 0
    valid_rows = []
    errors = []
    now_str = datetime.now().strftime(TIME_FMT)
    time_unrecognized = 0

    for idx in range(start, len(raw_rows)):
        row = raw_rows[idx]
        line_no = idx + 1

        # 跳过空行
        if not row or all(str(cell).strip() == "" for cell in row):
            continue

        if len(row) < 3:
            errors.append((line_no, "列数不足，至少需要 收支类型/分类/金额"))
            continue

        bill_type = row[0].strip()
        category = row[1].strip()
        amount_str = row[2].strip()
        note = row[3].strip() if len(row) >= 4 else ""

        if bill_type not in ("收入", "支出"):
            errors.append((line_no, f"收支类型无效：{bill_type!r}（应为 收入/支出）"))
            continue
        if not category:
            errors.append((line_no, "分类为空"))
            continue

        clean_amount = (amount_str.replace(",", "")
                        .replace("￥", "").replace("¥", "").replace("$", "").strip())
        try:
            amount = float(clean_amount)
        except ValueError:
            errors.append((line_no, f"金额格式错误：{amount_str!r}"))
            continue

        created = None
        if len(row) >= 5:
            created = _parse_time(row[4])
            if created is None and row[4].strip():
                time_unrecognized += 1
        if created is None:
            created = now_str

        valid_rows.append((bill_type, category, amount, note, created))

    if not valid_rows:
        print("⚠️ 没有可导入的有效记录", flush=True)
        if errors:
            print(f"共 {len(errors)} 条无效记录：", flush=True)
            _print_import_errors(errors)
        return

    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.executemany(
        "INSERT INTO bills (type, category, amount, note, created_at) VALUES (?, ?, ?, ?, ?)",
        valid_rows
    )
    conn.commit()
    conn.close()

    print(f"✅ 导入完成：成功 {len(valid_rows)} 条（文件编码：{enc}）", flush=True)
    if time_unrecognized:
        print(f"⚠️ 有 {time_unrecognized} 条记录的时间无法识别，已使用当前时间", flush=True)
    if errors:
        print(f"⚠️ 跳过 {len(errors)} 条无效记录：", flush=True)
        _print_import_errors(errors)
    print("提示：本功能不会自动去重。如需避免重复导入，请先查看现有账目。", flush=True)



def main():
    init_db()
    migrate_db()
    print(f"数据文件位置：{DB_FILE}", flush=True)
    while True:
        print("\n====简易个人收支记账系统（SQLite版）====", flush=True)
        print("1. 添加记账记录", flush=True)
        print("2. 查看全部账目", flush=True)
        print("3. 收支统计(总收入/总支出/结余)", flush=True)
        print("4. 按分类查询账目", flush=True)
        print("5. 删除记录", flush=True)
        print("6. 导出CSV文件", flush=True)
        print("7. 从CSV导入账目", flush=True)
        print("8. 退出程序", flush=True)
        choice = input("请输入功能编号：").strip()

        if choice == "1":
            tp = input("收入还是支出？(输入：收入/支出)：").strip()
            category = input("分类（餐饮/交通/学习等）：").strip()
            money_str = input("金额：").strip()
            try:
                amount = float(money_str)
            except ValueError:
                print("金额格式错误！", flush=True)
                continue
            note = input("备注：").strip()
            add_bill(tp, category, amount, note)
        elif choice == "2":
            show_all()
        elif choice == "3":
            calculate()
        elif choice == "4":
            filter_by_category()
        elif choice == "5":
            delete_item()
        elif choice == "6":
            export_csv()
        elif choice == "7":
            import_csv()
        elif choice == "8":
            print("程序结束", flush=True)
            break
        else:
            print("输入编号无效，请重新选择", flush=True)


if __name__ == "__main__":
    main()