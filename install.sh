#!/bin/bash
# powerpoint-converter のセットアップ。sudo は使わない（足りないものは入れ方を表示する）。何度実行してもよい。
#
#   1. LibreOffice（Impress）と、Python から操作するための python3-uno があるか確かめる
#   2. powerpoint-converter コマンドを ~/.local/bin に作る
#   3. LibreOffice を一度起動して、専用のプロファイルを作っておく（初回の表示を速くする）
set -euo pipefail
cd "$(dirname "$(readlink -f "$0")")"
HERE="$PWD"
ok=1

python3 -c 'import sys; sys.exit(sys.version_info < (3, 10))' \
  || { echo "Python 3.10 以上が必要"; exit 1; }

# ---------------------------------------------------------------- 1. LibreOffice
if command -v soffice >/dev/null || command -v libreoffice >/dev/null; then
  echo "LibreOffice: $(soffice --version 2>/dev/null | head -1)"
else
  echo "LibreOffice が無い。入れる:  sudo apt install libreoffice-impress"
  ok=0
fi
if python3 -c 'import uno' 2>/dev/null; then
  echo "python3-uno: あり"
else
  echo "python3-uno が無い。入れる:  sudo apt install python3-uno"
  ok=0
fi
command -v fc-match >/dev/null || echo "注意: fc-match が無い（フォントの置き換えを調べられない）。sudo apt install fontconfig"

# ---------------------------------------------------------------- 2. コマンド
mkdir -p "$HOME/.local/bin"
ln -sfn "$HERE/powerpoint_converter.sh" "$HOME/.local/bin/powerpoint-converter"
echo "powerpoint-converter: $HOME/.local/bin/powerpoint-converter → $HERE/powerpoint_converter.sh"
case ":$PATH:" in *":$HOME/.local/bin:"*) ;; *) echo "注意: ~/.local/bin が PATH に無い。~/.bashrc に足すこと" ;; esac

# ---------------------------------------------------------------- 3. プロファイル
if [ "$ok" = 1 ]; then
  python3 - <<'PY'
import sys
sys.path.insert(0, ".")
from powerpoint_converter.office import Office, OfficeError
for name in ("serve", "present", "cli", "bulk1", "bulk2", "bulk3", "bulk4"):
    o = Office(name)
    try:
        o._run(lambda d, c: None)
        print(f"LibreOffice のプロファイル（{name}）: 準備した")
    except OfficeError as e:
        print(f"注意: {e}")
    finally:
        o.release()
PY
  mkdir -p data
  echo
  echo "完了。どこからでも  powerpoint-converter  を実行すると、ブラウザで data/ の一覧が開く。"
else
  echo
  echo "上の足りないものを入れてから、もう一度  bash install.sh"
  exit 1
fi
