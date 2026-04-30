# -*- coding: utf-8 -*-
"""
批量重命名 VBA 模块文件，并同步更新 Attribute VB_Name。
读写均使用 GBK 编码，保证中文不乱码。
"""

import os
import re

MODULES_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "VBA_Export", "Modules")

RENAME_MAP = {
    "vbaSync":                              "vbaSync",
    "初始化配置":                            "初始化_配置",
    "功能1dot6模板与源数据表格校验":           "校验_模板源数据",
    "功能1dot7注入校验区域与汇总校验结果":      "校验_注入校验结果",
    "功能1dot9指标关系校验":                  "校验_指标关系",
    "功能1dot5地区数值核对":                  "机构_地区数值核对",
    "功能1dot2机构名称标准化":                "机构_名称标准化",
    "功能1dot1拆分村镇银行数据":              "机构_村镇银行",
    "功能1dot4外汇页眉修改":                  "机构_外汇页眉",
    "功能1dot3剔除外资行":                    "机构_剔除外资行",
    "功能1dot6提取工作表数据":                "机构_提取工作表",
    "功能1dot8涉农贷款比上月修正":             "机构_涉农贷款修正",
    "功能2dot1数值核对":                      "汇总_数值核对",
    "功能2dot2dot1按使用区域汇总":            "汇总_按区域",
    "功能2dot2dot2按批注汇总":               "汇总_按批注",
    "功能2dot3批量重命名文件":                "转换_批量重命名",
    "功能2dot4批量Excel格式转换":             "转换_Excel格式",
    "功能3dot9追加列按批注":                  "单元格_追加列",
    "功能3dot11批量Word格式转换":             "转换_Word格式",
    "功能3dot13批量修改Sheet名":              "转换_修改Sheet名",
    "功能3dot5取消合并并填充":                "单元格_取消合并",
    "功能3dot6合并相邻同值":                  "单元格_合并同值",
    "功能3dot10分列数字符号中文":              "单元格_分列",
    "功能3dot11按批注检查重复数据":            "去重_按批注检查",
    "功能3dot14按批注删除重复行":              "去重_按批注删除",
    "功能3dot12按配置检查重复数据":            "去重_按配置检查",
    "功能3dot13按配置删除重复行":              "去重_按配置删除",
    "功能3dot15按配置去重追加到目标":          "去重_按配置追加",
    "功能3dot14批量转时序":                   "时序_维度比对",
    "功能3dot15批量输出打印版":               "打印_输出操作",
    "功能3dot7注入VBA到源文件":               "VBA_注入",
    "功能3dot8清除目标工作簿VBA":             "VBA_清除",
}

def update_vb_name(content: str, new_name: str) -> str:
    return re.sub(
        r'^Attribute VB_Name = ".*?"',
        f'Attribute VB_Name = "{new_name}"',
        content,
        count=1,
        flags=re.MULTILINE,
    )

def main():
    results = []
    for old_name, new_name in RENAME_MAP.items():
        old_path = os.path.join(MODULES_DIR, old_name + ".bas")
        new_path = os.path.join(MODULES_DIR, new_name + ".bas")

        if not os.path.exists(old_path):
            results.append(f"[跳过] 文件不存在: {old_name}.bas")
            continue

        if old_path == new_path:
            results.append(f"[无变化] {old_name}.bas")
            continue

        if os.path.exists(new_path):
            results.append(f"[冲突] 目标已存在，跳过: {new_name}.bas")
            continue

        with open(old_path, "r", encoding="gbk", errors="replace") as f:
            content = f.read()

        content = update_vb_name(content, new_name)

        with open(new_path, "w", encoding="gbk") as f:
            f.write(content)

        os.remove(old_path)
        results.append(f"[完成] {old_name}.bas  →  {new_name}.bas")

    print("\n".join(results))
    print(f"\n共处理 {len(results)} 项")

if __name__ == "__main__":
    main()
