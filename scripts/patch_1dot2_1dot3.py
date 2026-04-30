# -*- coding: utf-8 -*-
"""
修复 1dot2 和 1dot3 模块的三个问题：
  1. 1dot3 批量/单个 Sub 调用精确版删除外资行 → 改为增强版（支持模糊匹配）
  2. 1dot3 SaveAs 错误处理漏洞（On Error Resume Next 下二次失败被吞）
  3. 1dot3 批量 Sub 每文件后的无意义 1 秒等待
  4. 1dot2 批量 Sub 缺少 Calculation / EnableEvents 关闭（批量写入时重复触发重算）
"""
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MOD = os.path.join(BASE, 'VBA_Export', 'Modules')


def read_gbk(path):
    with open(path, 'r', encoding='gbk', newline='') as f:
        return f.read()


def write_utf8(path, content):
    with open(path, 'w', encoding='utf-8', newline='') as f:
        f.write(content)


def assert_replaced(old, new, content, label):
    if old not in content:
        print(f'  [WARN] 未找到替换目标: {label}')
        return content
    result = content.replace(old, new, 1)
    print(f'  [OK] {label}')
    return result


# ─── 1dot3 ────────────────────────────────────────────────────────────────────
def patch_1dot3():
    path = os.path.join(MOD, '功能1dot3剔除外资行.bas')
    content = read_gbk(path)

    # 修复1：批量/单个 Sub 均改调增强版（两处 deletedRows = deletedRows + 删除外资行(）
    content = assert_replaced(
        'deletedRows = deletedRows + 删除外资行(tempWs, foreignBankDict)',
        'deletedRows = deletedRows + 删除外资行增强版(tempWs, foreignBankDict)',
        content,
        '1dot3-fix1 单个Sub：改调增强版'
    )
    # 批量Sub 里还有一处（上面只替换了第一处，再替换一次）
    content = assert_replaced(
        'deletedRows = deletedRows + 删除外资行(tempWs, foreignBankDict)',
        'deletedRows = deletedRows + 删除外资行增强版(tempWs, foreignBankDict)',
        content,
        '1dot3-fix1 批量Sub：改调增强版'
    )

    # 修复2：SaveAs 错误处理逻辑
    old_save = (
        "            '保存副本\r\n"
        "            On Error Resume Next\r\n"
        "            tempWb.SaveAs fileName:=savePath, FileFormat:=sourceWb.FileFormat\r\n"
        "            If Err.Number <> 0 Then\r\n"
        "                '如果保存失败，尝试使用默认格式\r\n"
        "                tempWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook\r\n"
        "            End If\r\n"
        "            \r\n"
        "            If Err.Number = 0 Then\r\n"
        "                保存成功文件数 = 保存成功文件数 + 1\r\n"
        '                RunLog_WriteRow "1.3 删除分机构表的外资行", "处理文件", newName, "", "", "成功", "删除 " & deletedRows & " 行外资行", ""\r\n'
        '                Debug.Print "文件 " & fileCount & ": " & newName & " (删除 " & deletedRows & " 行外资行)"\r\n'
        "            Else\r\n"
        '                RunLog_WriteRow "1.3 删除分机构表的外资行", "处理文件", newName, "", "", "失败", "保存失败 " & Err.Number, ""\r\n'
        '                Debug.Print "文件 " & fileCount & ": 保存失败 - " & newName\r\n'
        "            End If\r\n"
        "            On Error GoTo 0"
    )
    new_save = (
        "            '保存副本\r\n"
        "            Dim saveOk As Boolean\r\n"
        "            Dim saveErrNum As Long\r\n"
        "            saveOk = False\r\n"
        "            saveErrNum = 0\r\n"
        "            On Error Resume Next\r\n"
        "            Err.Clear\r\n"
        "            tempWb.SaveAs fileName:=savePath, FileFormat:=sourceWb.FileFormat\r\n"
        "            If Err.Number = 0 Then\r\n"
        "                saveOk = True\r\n"
        "            Else\r\n"
        "                Err.Clear\r\n"
        "                tempWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook\r\n"
        "                saveOk = (Err.Number = 0)\r\n"
        "                saveErrNum = Err.Number\r\n"
        "            End If\r\n"
        "            On Error GoTo 0\r\n"
        "            \r\n"
        "            If saveOk Then\r\n"
        "                保存成功文件数 = 保存成功文件数 + 1\r\n"
        '                RunLog_WriteRow "1.3 删除分机构表的外资行", "处理文件", newName, "", "", "成功", "删除 " & deletedRows & " 行外资行", ""\r\n'
        '                Debug.Print "文件 " & fileCount & ": " & newName & " (删除 " & deletedRows & " 行外资行)"\r\n'
        "            Else\r\n"
        '                RunLog_WriteRow "1.3 删除分机构表的外资行", "处理文件", newName, "", "", "失败", "保存失败 " & saveErrNum, ""\r\n'
        '                Debug.Print "文件 " & fileCount & ": 保存失败 - " & newName\r\n'
        "            End If"
    )
    content = assert_replaced(old_save, new_save, content, '1dot3-fix2 SaveAs 错误处理')

    # 修复3：删除无意义的 1 秒等待
    content = assert_replaced(
        '            Application.Wait (Now + TimeValue("0:00:01"))\r\n',
        '',
        content,
        '1dot3-fix3 删除 Application.Wait'
    )

    write_utf8(path, content)
    print('  => 1dot3 已写回（UTF-8，待 convert.py 转 GBK）')


# ─── 1dot2 ────────────────────────────────────────────────────────────────────
def patch_1dot2():
    path = os.path.join(MOD, '功能1dot2机构名称标准化.bas')
    content = read_gbk(path)

    # 修复4：批量 Sub 中补充关闭重算和事件
    old_app_off = (
        "        Application.ScreenUpdating = False\r\n"
        "        Application.DisplayAlerts = False\r\n"
        "        \r\n"
        "        totalMapped = 0"
    )
    new_app_off = (
        "        Application.ScreenUpdating = False\r\n"
        "        Application.DisplayAlerts = False\r\n"
        "        Application.Calculation = xlCalculationManual\r\n"
        "        Application.EnableEvents = False\r\n"
        "        \r\n"
        "        totalMapped = 0"
    )
    content = assert_replaced(old_app_off, new_app_off, content, '1dot2-fix4a 关闭 Calculation/EnableEvents')

    old_app_on = (
        "        Application.ScreenUpdating = True\r\n"
        "        Application.DisplayAlerts = True\r\n"
        "        \r\n"
        '        RunLog_WriteRow "1.2 机构名称标准化", "完成"'
    )
    new_app_on = (
        "        Application.ScreenUpdating = True\r\n"
        "        Application.DisplayAlerts = True\r\n"
        "        Application.Calculation = xlCalculationAutomatic\r\n"
        "        Application.EnableEvents = True\r\n"
        "        \r\n"
        '        RunLog_WriteRow "1.2 机构名称标准化", "完成"'
    )
    content = assert_replaced(old_app_on, new_app_on, content, '1dot2-fix4b 恢复 Calculation/EnableEvents')

    write_utf8(path, content)
    print('  => 1dot2 已写回（UTF-8，待 convert.py 转 GBK）')


if __name__ == '__main__':
    print('=== patch 1dot2 ===')
    patch_1dot2()
    print('=== patch 1dot3 ===')
    patch_1dot3()
    print('\n全部完成，请运行 python convert.py 转换为 GBK。')
