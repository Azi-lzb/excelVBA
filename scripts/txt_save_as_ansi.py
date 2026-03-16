# -*- coding: utf-8 -*-
"""
将文本文件（如 .bas、.cls、.txt）用文本方式打开后另存为 ANSI 格式（Windows 下一般为 GBK/CP936）。
适用于 VBA 模块 .bas/.cls 转成 ANSI 后供 Excel 导入不乱码。不依赖 VBA，纯 Python。

用法:
  python txt_save_as_ansi.py <输入文件> [输出文件]
  仅一个参数时覆盖原文件（先读后写）。

示例:
  python txt_save_as_ansi.py VBA_Export\\Modules\\追加列按批注.bas
  python txt_save_as_ansi.py C:\\temp\\a.bas C:\\temp\\a_ansi.bas
  python txt_save_as_ansi.py C:\\temp\\a.txt
"""
import os
import sys

# Windows 中文系统下 ANSI 一般为 GBK (CP936)
ANSI_ENCODING = "gbk"


def detect_and_read(path):
    """尝试用 UTF-8 / UTF-8-BOM / GBK / GB2312 读取，返回 (content, used_encoding)。"""
    with open(path, "rb") as f:
        raw = f.read(4)
    if raw.startswith(b"\xef\xbb\xbf"):
        with open(path, "r", encoding="utf-8-sig") as f:
            return f.read(), "utf-8-sig"
    for enc in ("utf-8", "gbk", "gb2312"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read(), enc
        except (UnicodeDecodeError, LookupError):
            continue
    return None, None


def save_as_ansi(input_path, output_path=None):
    """
    将文件以文本方式读取后另存为 ANSI（GBK）格式，支持 .bas / .cls / .txt 等。
    output_path 为空时则覆盖输入文件。
    返回 True 成功，False 失败。
    """
    if output_path is None:
        output_path = input_path

    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    if not os.path.isfile(input_path):
        print("错误：找不到文件 " + input_path, file=sys.stderr)
        return False

    content, used_enc = detect_and_read(input_path)
    if content is None:
        print("错误：无法识别文件编码，请确保为 UTF-8 或 GBK/GB2312。", file=sys.stderr)
        return False

    try:
        with open(output_path, "w", encoding=ANSI_ENCODING) as f:
            f.write(content)
    except UnicodeEncodeError as e:
        print("错误：存在无法用 ANSI(GBK) 表示的字符：" + str(e), file=sys.stderr)
        return False
    except OSError as e:
        print("错误：写入失败 " + str(e), file=sys.stderr)
        return False

    if output_path == input_path:
        print("已覆盖为 ANSI(GBK)：", input_path, "（原编码：%s）" % used_enc)
    else:
        print("已另存为 ANSI(GBK)：", output_path, "（原编码：%s）" % used_enc)
    return True


def main():
    if len(sys.argv) < 2:
        print(__doc__.strip())
        print("\n请提供至少一个参数：输入文件路径（如 .bas / .cls / .txt）。")
        sys.exit(1)

    inp = sys.argv[1].strip()
    out = sys.argv[2].strip() if len(sys.argv) > 2 else None

    ok = save_as_ansi(inp, out)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
