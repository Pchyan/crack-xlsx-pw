import PyInstaller.__main__
import os
import shutil

def build():
    # 清理之前的建置檔案
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
        
    # PyInstaller 參數
    params = [
        'gui.py',  # 主程式
        '--name=Excel密碼破解工具',  # 執行檔名稱
        '--windowed',  # 使用 GUI 模式
        '--onefile',  # 打包成單一檔案
        '--icon=icon.ico',  # 程式圖示（如果有）
        '--add-data=README.md;.',  # 加入說明文件
        '--clean',  # 清理暫存檔
        '--noconfirm',  # 不詢問確認
        '--hidden-import=openpyxl',
        '--hidden-import=tqdm'
    ]
    
    # 執行打包
    PyInstaller.__main__.run(params)
    
    # 複製必要檔案到 dist 目錄
    dist_dir = 'dist/Excel密碼破解工具'
    if not os.path.exists(dist_dir):
        os.makedirs(dist_dir)
    
    # 複製 README
    if os.path.exists('README.md'):
        shutil.copy2('README.md', dist_dir)
    
    print("\n打包完成！")
    print(f"執行檔位置：{os.path.abspath('dist/Excel密碼破解工具.exe')}")
    print(f"完整程式目錄：{os.path.abspath(dist_dir)}")

if __name__ == "__main__":
    build() 