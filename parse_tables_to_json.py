# -*- coding: utf-8 -*-
"""
独自タグから表を復元し、階層ヘッダーつきJSONを出力するスクリプト。
- <"図表ネーム"> …… 直近の表のシート名候補（「表」を含む場合のみ採用）
- <"表...">        …… テーブル開始
- <"行"> / <"行_罫なし"> …… 行区切り
- <"G...">内容     …… セル。 '＝C◯_C◯' / '＝C◯_T◯' で rowspan/colspan
- <KG> は ' / ' に変換、その他の <...> は除去（<sub> 等の装飾含む）

JSON構造（例）:
{
  "tables": [
    {
      "id": "表2023_0016_3a",
      "name": "表3a 脂質異常症診断基準",
      "columns": [
        {"path": ["LDLコレステロール"], "key": "LDLコレステロール"},
        {"path": ["血清脂質","C"], "key": "血清脂質|C"},
        {"path": ["血清脂質","TG"], "key": "血清脂質|TG"},
        ...
      ],
      "rows": [
        {"LDLコレステロール": "140mg/dL以上", "血清脂質|C": "→", "血清脂質|TG": "↑↑↑", ...},
        ...
      ]
    }
  ]
}
"""
import sys
import re
import json
from typing import List, Tuple, Dict, Any

# ---------- 正規表現 ----------
CAPTION_RE      = re.compile(r'^<"図表ネーム">(.*)$')
TABLE_START_RE  = re.compile(r'^<"表[^"]*">$')
ROW_RE          = re.compile(r'^<"行[^"]*">$')   # <"行"> と <"行_罫なし"> をまとめて扱う
CELL_RE         = re.compile(r'^<"G[^"]*">(.*)$')
SPAN_RE         = re.compile(r'＝([CT])(\d+)_([CT])(\d+)')  # ＝C2_C1 / ＝C1_T2 など
HEADER_HINT_RE  = re.compile(r'こ色')  # タグ中に「色」っぽい指定があるセルをヘッダー候補に

# ---------- ユーティリティ ----------
def strip_inline_tags(text: str) -> str:
    text = text.replace('<KG>', ' / ')
    text = re.sub(r'<[^>]+>', '', text)
    return re.sub(r'\s+', ' ', text).strip()

def parse_span(tag: str) -> Tuple[int, int]:
    m = SPAN_RE.search(tag)
    if not m:
        return 1, 1
    rowspan = int(m.group(2))
    colspan = int(m.group(4))
    return rowspan, colspan

def parse_cell_line(line: str) -> Tuple[str, int, int, bool]:
    # line の左側（">"の直前）をタグとして扱い、span と header ヒントを抽出
    tag = line.split('">', 1)[0]
    rs, cs = parse_span(tag)
    header_hint = bool(HEADER_HINT_RE.search(tag)) or (rs > 1 or cs > 1)
    m = CELL_RE.match(line)
    val = strip_inline_tags(m.group(1)) if m else ''
    return val, rs, cs, header_hint

def sanitize_name(name: str) -> str:
    name = name.replace(':', '：')
    name = re.sub(r'[\\/*?\[\]]', ' ', name)
    return name

# ---------- コア配置（値+メタ） ----------
def place_cells_with_meta(rows_spec: List[List[Tuple[str,int,int,bool]]]):
    """
    ★FIX: 占有マップ(occ)で空き判定。上段スパン直下を誤って「空き」と見做さない。
    戻り値:
        grid_vals  : List[List[str]]            可視グリッド（結合部は ''）
        merges     : List[(r1,c1,r2,c2)]        1-based 結合範囲
        top_lefts  : Dict[(r,c)] -> (rs,cs,header_hint)
    """
    grid_vals: List[List[str]] = []
    merges: List[Tuple[int,int,int,int]] = []
    top_lefts: Dict[Tuple[int,int], Tuple[int,int,bool]] = {}
    occ: Dict[Tuple[int,int], bool] = {}  # 1-based (row,col)

    row_idx = 0
    for row in rows_spec:
        row_idx += 1
        while len(grid_vals) < row_idx:
            grid_vals.append([])

        col = 1
        for val, rs, cs, hdr in row:
            # 次の空き列を occ で確認
            while occ.get((row_idx, col), False):
                col += 1
            # グリッド拡張
            for rr in range(row_idx, row_idx + rs):
                while len(grid_vals) < rr:
                    grid_vals.append([])
                while len(grid_vals[rr - 1]) < col + cs - 1:
                    grid_vals[rr - 1].append('')
            while len(grid_vals[row_idx - 1]) < col:
                grid_vals[row_idx - 1].append('')
            # 値を配置
            grid_vals[row_idx - 1][col - 1] = val
            top_lefts[(row_idx, col)] = (rs, cs, hdr)
            # スパン占有
            for rr in range(row_idx, row_idx + rs):
                for cc in range(col, col + cs):
                    occ[(rr, cc)] = True
            if rs > 1 or cs > 1:
                merges.append((row_idx, col, row_idx + rs - 1, col + cs - 1))
            col += cs

    return grid_vals, merges, top_lefts

def fill_merged_labels(grid_vals: List[List[str]], merges: List[Tuple[int,int,int,int]]) -> List[List[str]]:
    """結合範囲を同一ラベルで埋めたグリッド（ラベルプレビュー用）"""
    filled = [row[:] for row in grid_vals]
    for r1, c1, r2, c2 in merges:
        label = filled[r1-1][c1-1]
        for r in range(r1-1, r2):
            for c in range(c1-1, c2):
                filled[r][c] = label
    # 単独セルの '' はそのまま
    return filled

# ---------- 論理カラム統合と値抽出 ----------
def header_path_for_column(label_grid: List[List[str]], c: int, header_depth: int) -> List[str]:
    """
    指定列のヘッダーパスを構築する。
    連続する同じ値はスキップして、階層的なパスを作る。
    """
    path = []
    last = None
    for r in range(header_depth):
        if c < len(label_grid[r]):
            s = label_grid[r][c].strip()
            if s and s != last:
                path.append(s)
                last = s
    return path if path else [f"col{c+1}"]

def build_logical_columns(label_grid: List[List[str]], header_depth: int) -> List[Dict[str, Any]]:
    """
    連続する同一ヘッダーパスを1つの論理カラムに統合する。
    
    Returns:
        [{path: [...], key: "...", col_range: (start, end)}, ...]
        col_rangeは1-basedの範囲（start, end）を表す
    """
    if not label_grid:
        return []
    
    col_count = max((len(row) for row in label_grid), default=0)
    if col_count == 0:
        return []
    
    logical_columns = []
    c = 0
    
    while c < col_count:
        # 現在の列のパス取得
        current_path = tuple(header_path_for_column(label_grid, c, header_depth))
        start_col = c + 1  # 1-based
        
        # 連続する同じパスの範囲を見つける
        while c + 1 < col_count:
            next_path = tuple(header_path_for_column(label_grid, c + 1, header_depth))
            if next_path == current_path:
                c += 1
            else:
                break
        
        end_col = c + 1  # 1-based
        
        logical_columns.append({
            "path": list(current_path),
            "col_range": (start_col, end_col)
        })
        c += 1
    
    # ユニークキーの生成（非連続重複にサフィックス付与）
    seen_keys = {}
    for col in logical_columns:
        base_key = "|".join(col["path"])
        if base_key in seen_keys:
            seen_keys[base_key] += 1
            col["key"] = f"{base_key}_{seen_keys[base_key]}"
        else:
            seen_keys[base_key] = 1
            col["key"] = base_key
    
    return logical_columns

def extract_value_from_range(row_cells: List[str], col_range: Tuple[int, int], 
                            policy: str = "first_nonempty", separator: str = " / ") -> str:
    """
    論理カラムの範囲から値を抽出する。
    
    Args:
        row_cells: 行のセル値リスト（0-based）
        col_range: (start, end) 1-basedの列範囲
        policy: "first_nonempty", "last_nonempty", "concat"
        separator: concat時のセパレータ
    
    Returns:
        抽出された値
    """
    start, end = col_range
    values = []
    
    for c in range(start, end + 1):
        if c - 1 < len(row_cells):
            values.append(row_cells[c - 1])
        else:
            values.append("")
    
    non_empty = [v for v in values if v.strip()]
    
    if policy == "last_nonempty":
        return non_empty[-1] if non_empty else ""
    elif policy == "concat":
        return separator.join(non_empty)
    else:  # first_nonempty (default)
        return non_empty[0] if non_empty else ""

def detect_classification_spans(merges: List[Tuple[int, int, int, int]], 
                               label_grid: List[List[str]], 
                               header_depth: int,
                               target_col: int = 1) -> List[Dict[str, Any]]:
    """
    区分見出し（縦長の分類セル）を検出する。
    幅1列（c1==c2）かつ rowspan>=2 のマージを対象とする。
    
    Returns:
        [{label: "...", r1: start_row, r2: end_row}, ...]
    """
    class_spans = []
    
    for r1, c1, r2, c2 in merges:
        # 幅1列、データ領域内、縦結合
        if c1 == c2 == target_col and r1 > header_depth and r2 > r1:
            if r1 - 1 < len(label_grid) and c1 - 1 < len(label_grid[r1 - 1]):
                label = label_grid[r1 - 1][c1 - 1]
                class_spans.append({
                    "label": label,
                    "r1": r1,
                    "r2": r2
                })
    
    return class_spans

def get_classification_for_row(row_num: int, class_spans: List[Dict[str, Any]]) -> str:
    """
    指定行の分類ラベルを取得する（フラットモード用）。
    """
    for span in class_spans:
        if span["r1"] <= row_num <= span["r2"]:
            return span["label"]
    return ""



# ---------- ヘッダー行の自動判定 ----------
def infer_header_depth_from_first_row(rows_spec) -> int:
    """
    新しいヘッダー深度推定ロジック：
    表の最初の行（最初に出現する <"行"> ブロック）に含まれるセルのrowspanの最大値を返す。
    
    規則：
    - 表の最初の行に出現する各セルの "前半のC/Tに付く数値"（= 行方向の連結数 / rowspan）の最大値を header depth とする
    - 例：最初の行に ＝C2_C1, ＝C1_C2, ＝C2_C1 があれば max(2,1,2)=2 行がヘッダー
    - セルがない/数値が取れない場合は 1 を返す（フォールバック）
    
    Args:
        rows_spec: [[(val, rowspan, colspan, header_hint), ...], ...] 形式のリスト
    
    Returns:
        int: ヘッダー深度（最低1）
    """
    if not rows_spec:
        return 1
    
    # 最初の行を取得
    first_row = rows_spec[0] if rows_spec else []
    
    max_rs = 1
    for cell in first_row:
        # cell は (value, rowspan, colspan, header_hint) 形式
        # rowspan は2番目の要素（インデックス1）
        try:
            rs = int(cell[1]) if len(cell) > 1 else 1
            if rs > max_rs:
                max_rs = rs
        except (ValueError, TypeError, IndexError):
            # 数値に変換できない場合はスキップ
            pass
    
    return max_rs  # 少なくとも1行はヘッダー

# ---------- テーブルパース ----------
def parse_text_to_tables(text: str, keep_dividers: bool = False):
    """
    テキストから表を抽出する。
    区切り行（セルのない行）は既定でスキップ、keep_dividers=Trueで保持。
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    tables = []
    current_caption = None
    current_table = None
    current_row = []
    
    def flush_row():
        nonlocal current_row, current_table
        if current_table is not None:
            if current_row:
                # セルがある通常の行
                current_table['rows'].append(current_row)
            elif keep_dividers:
                # セルのない区切り行をメタデータとして保持
                current_table['rows'].append([])  # 空行として記録
            # keep_dividers=Falseの場合、セルのない行は単にスキップ
            current_row = []
    
    def flush_table():
        nonlocal current_table, tables
        if current_table is not None:
            flush_row()
            tables.append(current_table)
            current_table = None

    for ln in lines:
        mcap = CAPTION_RE.match(ln)
        if mcap:
            current_caption = strip_inline_tags(mcap.group(1))
            continue

        if TABLE_START_RE.match(ln):
            flush_table()
            table_id = ln[2:-2]  # <" と "> を剥がす
            name = current_caption if (current_caption and '表' in current_caption) else table_id
            current_table = {'id': table_id, 'name': sanitize_name(name), 'rows': []}
            continue

        if ROW_RE.match(ln):
            # 行区切り（<"行">, <"行_罫なし">, <"行_破線"> など）
            flush_row()
            continue

        if CELL_RE.match(ln) and current_table is not None:
            val, rs, cs, hdr = parse_cell_line(ln)
            current_row.append((val, rs, cs, hdr))
            continue

        # 図の説明や注釈は今回は無視

    flush_table()
    return tables

# ---------- JSON構築 ----------
def table_to_json(tbl: Dict[str, Any], 
                  manual_header_depth: int = None,
                  value_policy: str = "concat",
                  concat_separator: str = " / ",
                  group_key: str = "分類",
                  nested_mode: bool = False,
                  keep_dividers: bool = False,
                  add_classification: bool = False) -> Dict[str, Any]:
    """
    テーブルデータをJSON形式に変換する（論理カラム統合版）。
    
    Args:
        tbl: テーブルデータ（'id', 'name', 'rows'を含む辞書）
        manual_header_depth: 手動指定のヘッダー深度（Noneの場合は自動推定）
        value_policy: 値抽出ポリシー ("first_nonempty", "last_nonempty", "concat")
        concat_separator: concat時のセパレータ
        group_key: フラットモードでの分類列名
        nested_mode: Trueの場合ネストJSON出力
        keep_dividers: 区切り行を保持するか
        add_classification: 区分見出しを分類列として追加するか
    
    Returns:
        JSON形式のテーブルデータ
    """
    rows_spec = tbl['rows']
    if not rows_spec:
        return {"id": tbl['id'], "name": tbl['name'], "columns": [], "rows": []}
    
    # 空行（区切り行）を除外してから処理（keep_dividers=Falseの場合）
    if not keep_dividers:
        rows_spec = [row for row in rows_spec if row]
    
    if not rows_spec:
        return {"id": tbl['id'], "name": tbl['name'], "columns": [], "rows": []}
    
    grid_vals, merges, top_lefts = place_cells_with_meta(rows_spec)
    
    # ヘッダー深度の決定
    if manual_header_depth is not None:
        header_depth = manual_header_depth
    else:
        header_depth = infer_header_depth_from_first_row(rows_spec)
    
    # ラベルグリッドの作成（結合範囲を埋める）
    label_grid = fill_merged_labels(grid_vals, merges)
    
    # 論理カラムの構築（連続同一パスを統合）
    logical_columns = build_logical_columns(label_grid, header_depth)
    
    # 区分見出しの検出（add_classificationが有効な場合のみ処理）
    class_spans = []
    if add_classification:
        class_spans = detect_classification_spans(merges, label_grid, header_depth)
    
    # フラットモードで分類列がある場合、先頭に分類カラムを追加
    if class_spans and not nested_mode and add_classification:
        # 分類カラムを先頭に挿入
        logical_columns.insert(0, {
            "path": [group_key],
            "key": group_key,
            "col_range": None  # 特別な列
        })
    
    # データ行の処理
    data_rows = []
    
    for r_idx in range(header_depth, len(grid_vals)):
        row_num = r_idx + 1  # 1-based
        
        # keep_dividersでかつ空行の場合
        if keep_dividers and r_idx < len(rows_spec) and not rows_spec[r_idx]:
            data_rows.append({"divider": True})
            continue
        
        row_obj = {}
        
        # フラットモードで分類値を追加（add_classificationが有効な場合のみ）
        if class_spans and not nested_mode and add_classification:
            classification = get_classification_for_row(row_num, class_spans)
            row_obj[group_key] = classification
        
        # 各論理カラムの値を抽出
        for col in logical_columns:
            if col.get("col_range"):  # 通常のデータカラム
                value = extract_value_from_range(
                    grid_vals[r_idx], 
                    col["col_range"], 
                    value_policy, 
                    concat_separator
                )
                row_obj[col["key"]] = value
        
        data_rows.append(row_obj)
    
    # ネストモードの場合、グループごとに行を分割
    if nested_mode and class_spans and add_classification:
        groups = []
        for span in class_spans:
            group_rows = []
            for r_idx in range(span["r1"] - 1, span["r2"]):  # 0-based indexに変換
                if r_idx - header_depth >= 0 and r_idx - header_depth < len(data_rows):
                    group_rows.append(data_rows[r_idx - header_depth])
            if group_rows:
                groups.append({
                    "label": span["label"],
                    "rows": group_rows
                })
        
        # グループに属さない行も追加
        grouped_rows = set()
        for span in class_spans:
            for r in range(span["r1"], span["r2"] + 1):
                grouped_rows.add(r - header_depth - 1)
        
        ungrouped_rows = []
        for i, row in enumerate(data_rows):
            if i not in grouped_rows:
                ungrouped_rows.append(row)
        
        if ungrouped_rows:
            groups.insert(0, {"label": "", "rows": ungrouped_rows})
        
        return {
            "id": tbl['id'],
            "name": tbl['name'],
            "header_depth": header_depth,
            "columns": [{"path": col["path"], "key": col["key"]} for col in logical_columns 
                       if col.get("col_range")],  # 分類列を除外
            "groups": groups
        }
    
    # フラットモード出力
    return {
        "id": tbl['id'],
        "name": tbl['name'],
        "header_depth": header_depth,
        "columns": [{"path": col["path"], "key": col["key"]} for col in logical_columns],
        "rows": data_rows
    }

def main():
    import argparse
    
    # CLIパーサーの設定
    parser = argparse.ArgumentParser(description='独自タグから表を復元し、階層ヘッダーつきJSONを出力')
    parser.add_argument('input_file', help='入力テキストファイル')
    parser.add_argument('output_file', nargs='?', default='tables.json', help='出力JSONファイル（デフォルト: tables.json）')
    
    # ヘッダー関連
    parser.add_argument('--header-depth', type=int, metavar='N',
                        help='ヘッダー深度を手動で指定（省略時は自動推定）')
    
    # 値抽出ポリシー
    parser.add_argument('--value-policy', choices=['first_nonempty', 'last_nonempty', 'concat'],
                        default='concat',
                        help='横結合セルからの値抽出ポリシー（デフォルト: concat）')
    parser.add_argument('--concat-sep', default=' / ',
                        help='concat時のセパレータ（デフォルト: " / "）')
    
    # 区分見出し関連
    parser.add_argument('--add-classification', action='store_true',
                        help='区分見出し（縦長セル）を分類列として追加')
    parser.add_argument('--group-key', default='分類',
                        help='フラットモードでの分類列名（デフォルト: 分類）')
    parser.add_argument('--nested', action='store_true',
                        help='ネストJSONモードで出力（--add-classificationと併用）')
    
    # 区切り行関連
    parser.add_argument('--keep-dividers', action='store_true',
                        help='データのない区切り行をメタデータとして保持')
    
    args = parser.parse_args()
    
    # header-depthのバリデーション
    manual_header_depth = None
    if args.header_depth is not None:
        if args.header_depth < 1:
            print(f"警告: --header-depth {args.header_depth} は無効です。1に矯正します。")
            manual_header_depth = 1
        else:
            manual_header_depth = args.header_depth
            print(f"手動指定: header_depth = {manual_header_depth}")

    with open(args.input_file, "r", encoding="utf-8") as f:
        text = f.read()

    # テーブルの解析（区切り行の扱いを指定）
    parsed = parse_text_to_tables(text, keep_dividers=args.keep_dividers)
    
    # 各テーブルをJSON形式に変換
    out = {"tables": []}
    for tbl in parsed:
        json_table = table_to_json(
            tbl, 
            manual_header_depth=manual_header_depth,
            value_policy=args.value_policy,
            concat_separator=args.concat_sep,
            group_key=args.group_key,
            nested_mode=args.nested,
            keep_dividers=args.keep_dividers,
            add_classification=args.add_classification
        )
        out["tables"].append(json_table)

    with open(args.output_file, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"OK: wrote {args.output_file} with {len(out['tables'])} table(s).")
    
    # オプション情報の表示
    if args.value_policy != 'first_nonempty':
        print(f"  値抽出ポリシー: {args.value_policy}")
    if args.nested:
        print(f"  ネストモード: 有効")
    if args.keep_dividers:
        print(f"  区切り行保持: 有効")
    if args.add_classification:
        print(f"  分類列追加: 有効 (列名: {args.group_key})")

if __name__ == "__main__":
    main()
