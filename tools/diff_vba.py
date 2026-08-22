"""
tools/diff_vba.py  — xlsm 実物と src/*.bas の乖離を検出する（最重要）
使い方: py tools/diff_vba.py [xlsmファイル] [srcディレクトリ]
デフォルト: py tools/diff_vba.py デイトレ用.xlsm src

終了コード:
  0 = すべて一致（コミット可能）
  1 = 差分あり（要更新）

乖離がある場合は export_vba.py を実行してから再コミットすること。
"""
import sys, os, re, difflib
sys.stdout.reconfigure(encoding='utf-8')

def normalize(code: str) -> list[str]:
    """比較用に行末空白と末尾の空行を正規化"""
    lines = [line.rstrip() for line in code.splitlines()]
    # 末尾の空行を除去（olevba と書き出しの改行数の差を無視する）
    while lines and lines[-1] == '':
        lines.pop()
    return lines

def main():
    xlsm = sys.argv[1] if len(sys.argv) > 1 else 'デイトレ用.xlsm'
    srcdir = sys.argv[2] if len(sys.argv) > 2 else 'src'

    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        print("ERROR: oletools が必要です: py -m pip install oletools")
        sys.exit(1)

    if not os.path.exists(xlsm):
        print(f"ERROR: {xlsm} が見つかりません")
        sys.exit(1)

    p = VBA_Parser(xlsm)

    # xlsm からモジュールコードを収集
    xlsm_modules: dict[str, str] = {}
    for _, _, name, code in p.extract_macros():
        if re.match(r'^Sheet\d+\.cls$', name):
            continue
        xlsm_modules[name] = code

    any_diff = False
    checked = 0

    for xlsm_name, xlsm_code in sorted(xlsm_modules.items()):
        # src/ 側のファイル名を決定
        if xlsm_name == 'frmController.frm':
            src_name = 'frmController.frm.txt'
            # イベントコードのみで比較するため xlsm 側もフィルタ
            lines = xlsm_code.splitlines()
            filtered = []
            skip = False
            for line in lines:
                if line.startswith('VERSION ') or line.startswith('BEGIN'):
                    skip = True
                if line.strip() == 'END':
                    skip = False
                    continue
                if not skip and not line.startswith('Attribute '):
                    filtered.append(line)
            xlsm_cmp = '\n'.join(filtered)
        elif xlsm_name.endswith('.cls'):
            src_name = xlsm_name
            lines = [l for l in xlsm_code.splitlines() if not l.startswith('Attribute ')]
            xlsm_cmp = '\n'.join(lines)
        else:
            src_name = xlsm_name
            xlsm_cmp = xlsm_code

        src_path = os.path.join(srcdir, src_name)
        checked += 1

        if not os.path.exists(src_path):
            print(f"[MISSING] {src_path} が src/ に存在しません")
            any_diff = True
            continue

        try:
            with open(src_path, encoding='cp932') as f:
                src_code = f.read()
        except UnicodeDecodeError:
            print(f"[ENC ERR] {src_path} が CP932 でデコードできません（UTF-8 で保存されている可能性）")
            any_diff = True
            continue

        a = normalize(xlsm_cmp)
        b = normalize(src_code)

        if a != b:
            any_diff = True
            diff = list(difflib.unified_diff(
                b, a,
                fromfile=f'src/{src_name}',
                tofile=f'xlsm/{xlsm_name}',
                lineterm=''
            ))
            print(f"\n[DIFF] {src_name}")
            # 差分の先頭30行だけ表示
            for line in diff[:30]:
                print(f"  {line}")
            if len(diff) > 30:
                print(f"  ... (差分 {len(diff)} 行、30行で打ち切り)")
        else:
            print(f"[SAME] {src_name}")

    print(f"\n--- 結果: {checked} モジュールを検査 ---")
    if any_diff:
        print("差分あり。export_vba.py を実行して src/ を更新してください。")
        sys.exit(1)
    else:
        print("すべて一致。コミット可能です。")

if __name__ == '__main__':
    main()
