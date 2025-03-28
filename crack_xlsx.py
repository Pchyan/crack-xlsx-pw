import openpyxl
from tqdm import tqdm
import sys
import os

def try_password(file_path, password):
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True, password=password)
        return True
    except:
        return False

def crack_xlsx(file_path, wordlist_path=None):
    if not os.path.exists(file_path):
        print(f"錯誤：找不到檔案 {file_path}")
        return None

    # 如果沒有提供密碼字典，使用預設的簡單密碼列表
    if wordlist_path is None:
        passwords = [
            "password", "123456", "admin", "12345678", "qwerty",
            "111111", "123123", "abc123", "password123", "123456789"
        ]
    else:
        try:
            with open(wordlist_path, 'r', encoding='utf-8') as f:
                passwords = [line.strip() for line in f if line.strip()]
        except Exception as e:
            print(f"讀取密碼字典時發生錯誤：{str(e)}")
            return None

    print(f"開始嘗試破解 {file_path}")
    print(f"共載入 {len(passwords)} 個密碼")

    for password in tqdm(passwords, desc="嘗試密碼中"):
        if try_password(file_path, password):
            print(f"\n成功！密碼是：{password}")
            return password

    print("\n未能找到正確的密碼")
    return None

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("使用方法：python crack_xlsx.py <xlsx檔案路徑> [密碼字典路徑]")
        sys.exit(1)

    file_path = sys.argv[1]
    wordlist_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    crack_xlsx(file_path, wordlist_path) 