import os

# 扫描 VBA_Export 目录下的所有源码文件夹
base_dir = r'VBA_Export'
folders = ['Modules', 'Classes', 'Documents', 'Forms']
extensions = ('.bas', '.cls', '.frm')

def convert_to_gbk():
    count = 0
    for folder in folders:
        folder_path = os.path.join(base_dir, folder)
        if not os.path.exists(folder_path):
            continue
            
        for filename in os.listdir(folder_path):
            if filename.endswith(extensions):
                file_path = os.path.join(folder_path, filename)
                
                try:
                    # 先尝试以 utf-8 读取
                    with open(file_path, 'r', encoding='utf-8', newline='') as f:
                        content = f.read()
                except UnicodeDecodeError:
                    # 如果已经是 GBK 或其他编码，则跳过，防止二次转换导致乱码
                    print(f"Skipping {file_path}: Not UTF-8 or already GBK.")
                    continue
                
                # 统一换行符为 CRLF (\r\n) 以符合 VBA 编辑器规范
                content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
                
                # 写回为 GBK (ANSI)
                with open(file_path, 'w', encoding='gbk', errors='ignore', newline='') as f:
                    f.write(content)
                
                print(f"Converted: {file_path} to GBK (CRLF)")
                count += 1

    print(f"\nTotal converted: {count} files.")

if __name__ == "__main__":
    convert_to_gbk()
