# -*- coding: utf-8 -*-
"""跨底稿建指标池：批量解码多个工作簿里的 EDB metric，输出可检索的指标池。

动机：Wind 安装目录里的指标库（etc/edb/EDBNew*.chs、etc/wdf/Indicator.xml、
etc/Matrix/wind_IndicatorTree.dat、res/wings/.../formulaBasicIndicator.json）
都是加密、压缩或只含财务指标，实测无法提取 EDB 代码。
真正可用的指标库是用户自己历史底稿里的连接批注——批量扫一遍比逐个打开快得多。

用法：
  python scripts/pool_metrics.py --roots "<指标库目录A>" "<指标库目录B>" --out pool.json
  python scripts/pool_metrics.py --roots "<指标库目录>" --terms 福建 人口 --out pool.json

仅标准库 + 同目录 edb_tools；不启动 Excel，不改任何文件。
"""
import argparse
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from edb_tools import inspect_book  # noqa: E402

SKIP_DIRS = {".venv", "venv", "node_modules", "site-packages", "__pycache__",
             "$RECYCLE.BIN", "AppData", ".git"}


def walk_books(roots, min_size, max_size):
    out = []
    for root in roots:
        if not os.path.isdir(root):
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
            for fn in filenames:
                if not fn.lower().endswith((".xlsx", ".xlsm")):
                    continue
                p = os.path.join(dirpath, fn)
                try:
                    size = os.path.getsize(p)
                except OSError:
                    continue
                if min_size <= size <= max_size:
                    out.append((size, p))
    # 大文件更可能有 Wind 连接，优先扫
    out.sort(key=lambda x: -x[0])
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--roots", nargs="+", required=True, help="要扫描的目录")
    ap.add_argument("--out", required=True, help="指标池 JSON 输出路径")
    ap.add_argument("--terms", nargs="*", default=[],
                    help="只保留名称/代码同时包含全部关键词的指标")
    ap.add_argument("--max-files", type=int, default=20, help="最多解码几个工作簿")
    ap.add_argument("--min-size", type=int, default=200_000)
    ap.add_argument("--max-size", type=int, default=40_000_000)
    a = ap.parse_args()

    books = walk_books(a.roots, a.min_size, a.max_size)
    pool, failed = {}, []
    scanned = 0
    for size, path in books[: a.max_files]:
        try:
            book = inspect_book(path)
        except Exception as ex:  # 单个坏文件不影响整体
            failed.append({"path": path, "error": f"{type(ex).__name__}: {ex}"})
            continue
        scanned += 1
        for anchor in book.get("anchors", []):
            for metric in anchor.get("config", {}).get("metric", []):
                code = metric.get("code")
                if not code:
                    continue
                entry = pool.setdefault(code, {
                    "code": code,
                    "name": metric.get("name"),
                    "unit": metric.get("unit"),
                    "freq": metric.get("freq"),
                    "changeRecord": metric.get("changeRecord", []),
                    "metric": metric,
                    "sources": [],
                })
                entry["sources"].append({
                    "book": os.path.basename(path),
                    "path": path,
                    "sheet": anchor.get("sheet"),
                    "anchor": anchor.get("cell"),
                })

    items = list(pool.values())
    if a.terms:
        lowered = [t.casefold() for t in a.terms]
        items = [v for v in items
                 if all(t in " ".join(str(v.get(k, "")) for k in
                                      ("code", "name", "unit")).casefold()
                        for t in lowered)]

    payload = {
        "scope": "用户历史底稿中的 EDB metric，非 Wind 全库",
        "scanned_books": scanned,
        "candidate_books": len(books),
        "failed": failed,
        "unique_metrics": len(pool),
        "returned": len(items),
        "terms": a.terms,
        "metrics": items,
    }
    Path(a.out).write_text(json.dumps(payload, ensure_ascii=False, indent=1),
                           encoding="utf-8")

    print(f"books: {len(books)} candidates, {scanned} decoded, {len(failed)} failed")
    print(f"unique metrics in pool: {len(pool)}; returned after --terms: {len(items)}")
    for v in items:
        src = v["sources"][0]["book"] if v["sources"] else "?"
        print(f"  {v['code']:<10} | {v.get('name')} | {v.get('unit')} | "
              f"freq={v.get('freq')} | {src}")
    if failed:
        print(f"failed books: {[f['path'] for f in failed]}")


if __name__ == "__main__":
    main()
