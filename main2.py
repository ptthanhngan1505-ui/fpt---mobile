import pandas as pd
import re


# ============================================================
# 1. ĐỌC FILE EXCEL
# ============================================================

FILE_EXCEL = "fptshop_dien_thoai.xlsx"


def load_data():
    df = pd.read_excel(FILE_EXCEL)

    # Xóa các dòng hoàn toàn trống
    df = df.dropna(how="all")

    return df


# ============================================================
# 2. TÌM TÊN CỘT
# ============================================================

def find_column(df, keywords):

    for column in df.columns:

        column_text = str(column).lower()

        for keyword in keywords:

            if keyword.lower() in column_text:
                return column

    return None


def get_columns(df):

    columns = {}

    # Tên sản phẩm
    columns["name"] = find_column(
        df,
        [
            "tên sản phẩm",
            "ten san pham",
            "ten_san_pham",
            "tên",
            "ten"
        ]
    )

    # Giá bán
    columns["price"] = find_column(
        df,
        [
            "giá bán",
            "gia ban",
            "gia_ban",
            "price"
        ]
    )

    # Màn hình
    columns["screen"] = find_column(
        df,
        [
            "màn hình",
            "man hinh",
            "man_hinh",
            "screen"
        ]
    )

    # Hệ điều hành
    columns["os"] = find_column(
        df,
        [
            "hệ điều hành",
            "he dieu hanh",
            "he_dieu_hanh",
            "operating system"
        ]
    )

    # Thông số
    columns["spec"] = find_column(
        df,
        [
            "thông số",
            "thong so",
            "thong_so",
            "spec"
        ]
    )

    # Pin
    columns["battery"] = find_column(
        df,
        [
            "dung lượng pin",
            "dung luong pin",
            "pin",
            "battery"
        ]
    )

    # Chip
    columns["chip"] = find_column(
        df,
        [
            "chip",
            "cpu",
            "vi xử lý",
            "vi xu ly"
        ]
    )

    return columns


# ============================================================
# 3. CHUYỂN GIÁ THÀNH SỐ
# ============================================================

def convert_price(value):

    if pd.isna(value):
        return None

    text = str(value)

    numbers = re.findall(r"\d+", text)

    if not numbers:
        return None

    try:
        return int("".join(numbers))

    except:
        return None


# ============================================================
# 4. LẤY KÍCH THƯỚC MÀN HÌNH
# ============================================================

def extract_screen(value):

    if pd.isna(value):
        return None

    text = str(value).replace(",", ".")

    match = re.search(
        r"(\d+(?:\.\d+)?)\s*(?:inch|in)",
        text,
        re.IGNORECASE
    )

    if match:

        return float(
            match.group(1)
        )

    return None


# ============================================================
# 5. LẤY DUNG LƯỢNG PIN
# ============================================================

def extract_battery(value):

    if pd.isna(value):
        return None

    text = str(value).lower()

    match = re.search(
        r"(\d{3,5})\s*mah",
        text
    )

    if match:

        return int(
            match.group(1)
        )

    return None


# ============================================================
# 6. TÌM ĐIỆN THOẠI TẦM GIÁ
# ============================================================

def find_by_price(df, cols):

    if cols["price"] is None:

        print(
            "\nKhông tìm thấy cột Giá bán!"
        )

        return

    try:

        min_price = int(
            input("Nhập giá thấp nhất: ")
        )

        max_price = int(
            input("Nhập giá cao nhất: ")
        )

    except ValueError:

        print(
            "Giá phải nhập bằng số!"
        )

        return

    print()
    print("=" * 70)
    print("ĐIỆN THOẠI TRONG TẦM GIÁ")
    print("=" * 70)

    count = 0

    for _, row in df.iterrows():

        price = convert_price(
            row[cols["price"]]
        )

        if price is None:
            continue

        if min_price <= price <= max_price:

            name = row[cols["name"]]

            print(
                f"{name} | {price:,} VNĐ"
            )

            count += 1

    print()
    print(
        f"Tìm thấy: {count} sản phẩm"
    )


# ============================================================
# 7. TÌM ĐIỆN THOẠI THEO HÃNG
# ============================================================

def find_by_brand(df, cols):

    if cols["name"] is None:

        print(
            "\nKhông tìm thấy cột tên sản phẩm!"
        )

        return

    brand = input(
        "Nhập hãng cần tìm: "
    ).strip().lower()

    print()
    print("=" * 70)
    print("ĐIỆN THOẠI THEO HÃNG")
    print("=" * 70)

    count = 0

    for _, row in df.iterrows():

        name = str(
            row[cols["name"]]
        )

        if brand in name.lower():

            if cols["price"]:

                price = convert_price(
                    row[cols["price"]]
                )

            else:

                price = None

            if price:

                print(
                    f"{name} | {price:,} VNĐ"
                )

            else:

                print(name)

            count += 1

    print()
    print(
        f"Tìm thấy: {count} sản phẩm"
    )


# ============================================================
# 8. TÌM THEO HỆ ĐIỀU HÀNH
# ============================================================

def find_by_os(df, cols):

    if cols["os"] is None:

        print(
            "\nKhông tìm thấy cột Hệ điều hành!"
        )

        return

    os_input = input(
        "Nhập hệ điều hành cần tìm: "
    ).strip().lower()

    print()
    print("=" * 70)
    print("ĐIỆN THOẠI THEO HỆ ĐIỀU HÀNH")
    print("=" * 70)

    count = 0

    for _, row in df.iterrows():

        os_value = str(
            row[cols["os"]]
        )

        if os_input in os_value.lower():

            name = row[cols["name"]]

            print(
                f"{name} | {os_value}"
            )

            count += 1

    print()
    print(
        f"Tìm thấy: {count} sản phẩm"
    )


# ============================================================
# 9. SO SÁNH THEO MÀN HÌNH
# ============================================================

def compare_screen(df, cols):

    if cols["screen"] is None:

        print(
            "\nKhông tìm thấy cột Màn hình!"
        )

        return

    results = []

    for _, row in df.iterrows():

        screen = extract_screen(
            row[cols["screen"]]
        )

        if screen is not None:

            name = row[cols["name"]]

            results.append(
                (name, screen)
            )

    results.sort(
        key=lambda x: x[1],
        reverse=True
    )

    print()
    print("=" * 70)
    print("SO SÁNH KÍCH THƯỚC MÀN HÌNH")
    print("=" * 70)

    for name, screen in results:

        print(
            f"{name} | {screen} inch"
        )


# ============================================================
# 10. SO SÁNH DUNG LƯỢNG PIN
# ============================================================

def compare_battery(df, cols):

    results = []

    for _, row in df.iterrows():

        name = row[cols["name"]]

        battery = None

        # Nếu có cột Pin riêng
        if cols["battery"]:

            battery = extract_battery(
                row[cols["battery"]]
            )

        # Nếu không có thì tìm trong Thông số
        if battery is None and cols["spec"]:

            battery = extract_battery(
                row[cols["spec"]]
            )

        if battery is not None:

            results.append(
                (name, battery)
            )

    results.sort(
        key=lambda x: x[1],
        reverse=True
    )

    print()
    print("=" * 70)
    print("SO SÁNH DUNG LƯỢNG PIN")
    print("=" * 70)

    if not results:

        print(
            "Không tìm thấy thông tin dung lượng pin!"
        )

        return

    for name, battery in results:

        print(
            f"{name} | {battery:,} mAh"
        )


# ============================================================
# 11. HIỂN THỊ TẤT CẢ ĐIỆN THOẠI
# ============================================================

def show_all(df, cols):

    print()
    print("=" * 70)
    print("TẤT CẢ ĐIỆN THOẠI")
    print("=" * 70)

    print(
        f"Tổng số sản phẩm: {len(df)}"
    )

    print()

    for i, (_, row) in enumerate(
        df.iterrows(),
        start=1
    ):

        name = row[cols["name"]]

        if cols["price"]:

            price = convert_price(
                row[cols["price"]]
            )

        else:

            price = None

        if price:

            print(
                f"{i}. {name} | "
                f"{price:,} VNĐ"
            )

        else:

            print(
                f"{i}. {name}"
            )


# ============================================================
# 12. TÌM ĐIỆN THOẠI THEO CHIP
# ============================================================

def find_by_chip(df, cols):

    if cols["chip"] is None:

        print(
            "\nKhông tìm thấy cột Chip!"
        )

        return

    chip_input = input(
        "Nhập chip cần tìm: "
    ).strip().lower()

    print()
    print("=" * 70)
    print("ĐIỆN THOẠI THEO CHIP")
    print("=" * 70)

    count = 0

    for _, row in df.iterrows():

        chip = str(
            row[cols["chip"]]
        )

        if chip_input in chip.lower():

            name = row[cols["name"]]

            print(
                f"{name} | {chip}"
            )

            count += 1

    print()
    print(
        f"Tìm thấy: {count} sản phẩm"
    )


# ============================================================
# 13. MENU
# ============================================================

def show_menu():

    print()
    print("=" * 60)
    print("        QUẢN LÝ ĐIỆN THOẠI FPT SHOP")
    print("=" * 60)

    print("1. Tìm điện thoại tầm giá")
    print("2. Tìm điện thoại theo hãng")
    print("3. Tìm theo hệ điều hành")
    print("4. So sánh theo màn hình")
    print("5. So sánh dung lượng pin")
    print("6. Hiển thị tất cả điện thoại")
    print("7. Tìm điện thoại theo chip")
    print("8. Thoát")

    print("=" * 60)


# ============================================================
# 14. CHƯƠNG TRÌNH CHÍNH
# ============================================================

def main():

    try:

        df = load_data()

        print(
            f"Đã đọc {len(df)} sản phẩm từ Excel."
        )

    except Exception as e:

        print(
            f"Lỗi đọc file Excel: {e}"
        )

        return

    cols = get_columns(df)

    # Kiểm tra các cột tìm được
    print("\nCác cột được nhận diện:")

    for key, value in cols.items():

        print(
            f"- {key}: {value}"
        )

    while True:

        show_menu()

        choice = input(
            "Nhập lựa chọn: "
        ).strip()

        if choice == "1":

            find_by_price(
                df,
                cols
            )

        elif choice == "2":

            find_by_brand(
                df,
                cols
            )

        elif choice == "3":

            find_by_os(
                df,
                cols
            )

        elif choice == "4":

            compare_screen(
                df,
                cols
            )

        elif choice == "5":

            compare_battery(
                df,
                cols
            )

        elif choice == "6":

            show_all(
                df,
                cols
            )

        elif choice == "7":

            find_by_chip(
                df,
                cols
            )

        elif choice == "8":

            print(
                "\nĐã thoát chương trình."
            )

            break

        else:

            print(
                "\nLựa chọn không hợp lệ!"
            )


# ============================================================
# CHẠY
# ============================================================

if __name__ == "__main__":

    main()