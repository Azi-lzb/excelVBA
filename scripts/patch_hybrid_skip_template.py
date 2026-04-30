# -*- coding: utf-8 -*-
"""给「执行表头混合比对」增加护栏：源文件若与模板同路径则跳过，避免 424。"""
from pathlib import Path

BAS = Path(__file__).resolve().parent.parent / "VBA_Export" / "Modules" / "功能3dot14批量转时序.bas"
ENC = "gbk"

GUARD = (
    '    If StrComp(CStr(fileItem), templatePath, vbTextCompare) = 0 Then\r\n'
    '        skipBooks = skipBooks + 1\r\n'
    '        RunLog_WriteRow HYBRID_COMPARE_LOG_KEY, "跳过", CStr(fileItem), "", "", "跳过", "与模板同一文件", ""\r\n'
    '        GoTo NextHybridSource\r\n'
    '    End If\r\n'
)

OLD_OPEN = 'Set sourceWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True, UpdateLinks:=0)'
NEW_OPEN = GUARD + '    ' + OLD_OPEN

OLD_CLOSE = 'SafeCloseWorkbook sourceWb\r\n'
NEW_CLOSE = OLD_CLOSE + 'NextHybridSource:\r\n'


def main():
    # 用 bytes 读 + 手动 decode，避免 read_text 把 \r\n 归一成 \n
    src = BAS.read_bytes().decode(ENC)
    start = src.find("Public Sub 执行表头混合比对")
    # 注意要找 label 形式（带冒号），避免命中 `On Error GoTo HybridCompareErrHandler`
    end = src.find("HybridCompareErrHandler:", start)
    if start < 0 or end < 0:
        raise SystemExit("未定位到执行表头混合比对 Sub")
    block = src[start:end]

    if "NextHybridSource:" in block:
        print("[skip] 已经打过补丁了")
        return

    open_count = block.count(OLD_OPEN)
    if open_count != 1:
        raise SystemExit(f"预期 1 处 Open(source)，实际 {open_count}")
    patched = block.replace(OLD_OPEN, NEW_OPEN, 1)

    close_count = patched.count(OLD_CLOSE)
    if close_count != 1:
        raise SystemExit(f"预期 1 处 SafeCloseWorkbook sourceWb + CRLF，实际 {close_count}")
    patched = patched.replace(OLD_CLOSE, NEW_CLOSE, 1)

    new_src = src[:start] + patched + src[end:]
    BAS.write_bytes(new_src.encode(ENC))
    print(f"[ok] 已补丁: {BAS}")


if __name__ == "__main__":
    main()
