import itertools
import string
from tqdm import tqdm
import argparse

def generate_password_dict(min_length=4, max_length=16, output_file='password_dict.txt'):
    # 定義字符集
    lowercase = string.ascii_lowercase
    uppercase = string.ascii_uppercase
    digits = string.digits
    special_chars = string.punctuation

    all_chars = lowercase + uppercase + digits + special_chars
    
    print(f"開始生成密碼字典...")
    print(f"密碼長度範圍：{min_length} - {max_length}")
    print(f"字符集大小：{len(all_chars)}")
    
    # 計算總密碼數量
    total_passwords = sum(len(all_chars) ** length for length in range(min_length, max_length + 1))
    print(f"預計生成密碼數量：{total_passwords:,}")
    
    # 生成密碼並寫入文件
    with open(output_file, 'w', encoding='utf-8') as file:
        for length in range(min_length, max_length + 1):
            print(f"\n生成長度為 {length} 的密碼...")
            for password_tuple in tqdm(itertools.product(all_chars, repeat=length), 
                                     total=len(all_chars) ** length,
                                     desc=f"長度 {length}"):
                password = ''.join(password_tuple)
                file.write(password + '\n')
    
    print(f"\n密碼字典已生成完成！")
    print(f"檔案位置：{output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='生成密碼字典工具')
    parser.add_argument('-m', '--min', type=int, default=4,
                      help='最小密碼長度（預設：4）')
    parser.add_argument('-M', '--max', type=int, default=8,
                      help='最大密碼長度（預設：8）')
    parser.add_argument('-o', '--output', type=str, default='password_dict.txt',
                      help='輸出檔案名稱（預設：password_dict.txt）')
    
    args = parser.parse_args()
    
    # 檢查參數是否合理
    if args.min < 1:
        print("錯誤：最小密碼長度必須大於 0")
        exit(1)
    if args.max < args.min:
        print("錯誤：最大密碼長度必須大於或等於最小密碼長度")
        exit(1)
    if args.max > 16:
        print("警告：密碼長度超過 16 位可能會產生極大的檔案，建議使用較小的範圍")
        response = input("是否繼續？(y/N): ")
        if response.lower() != 'y':
            print("已取消操作")
            exit(0)
    
    generate_password_dict(min_length=args.min, max_length=args.max, output_file=args.output) 