"""
tools/export_vba.py  — xlsm から VBA を src/ に書き出す
使い方: py tools/export_vba.py [xlsmファイル] [出力ディレクトリ]
デフォルト: py tools/export_vba.py デイトレ用.xlsm src

出力規則（CLAUDE.md §3.4 準拠）:
  - 標準モジュール (.bas): Attribute VB_Name 行を残す（VBE Import 用）
  - ドキュメントモジュール (.cls): Attribute 行を全除去（貼り付け専用）
  - frmController.frm: frmController.frm.txt として保存（イベントコードのみ）
  - Sheet*.cls: スキップ（VBA内容なし）
  - エンコード: CP932 / CRLF 必須
  - olevba が返す \r\n を吸収し、\r\r\n（二重CR）にならないよう書き出す
"""
import sys, os, re
sys.stdout.reconfigure(encoding='utf-8')


def write_cp932_crlf(path: str, lines: list) -> int:
    """CP932 / CRLF で書き出す。olevba の \\r\\n を吸収して二重CRを防ぐ。"""
    content = '\r\n'.join(line.rstrip('\r') for line in lines) + '\r\n'
    with open(path, 'wb') as f:
        f.write(content.encode('cp932'))
    return len(lines)


def main():
    xlsm = sys.argv[1] if len(sys.argv) > 1 else 'デイトレ用.xlsm'
    outdir = sys.argv[2] if len(sys.argv) > 2 else 'src'

    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        print("ERROR: oletools が必要です: py -m pip install oletools")
        sys.exit(1)

    if not os.path.exists(xlsm):
        print(f"ERROR: {xlsm} が見つかりません")
        sys.exit(1)

    os.makedirs(outdir, exist_ok=True)
    p = VBA_Parser(xlsm)
    written = []

    for _, _, name, code in p.extract_macros():
        # Sheet*.cls はスキップ（コードなし）
        if re.match(r'^Sheet\d+\.cls$', name):
            continue

        if name == 'frmController.frm':
            # VERSION/BEGIN...END ブロックと Attribute 行を除去
            lines = code.splitlines()
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
            out_path = os.path.join(outdir, 'frmController.frm.txt')
            n = write_cp932_crlf(out_path, filtered)
            written.append((out_path, n))

        elif name.endswith('.cls'):
            # ThisWorkbook.cls: Attribute 行を除去
            lines = [l for l in code.splitlines() if not l.startswith('Attribute ')]
            out_path = os.path.join(outdir, name)
            n = write_cp932_crlf(out_path, lines)
            written.append((out_path, n))

        elif name.endswith('.bas'):
            # 標準モジュール: そのまま（Attribute VB_Name は残す）
            out_path = os.path.join(outdir, name)
            n = write_cp932_crlf(out_path, code.splitlines())
            written.append((out_path, n))

    print(f"[export_vba] {xlsm} → {outdir}/ に {len(written)} ファイルを書き出しました")
    for path, lines in sorted(written):
        print(f"  {path:50s} {lines:4d}行")


if __name__ == '__main__':
    main()
