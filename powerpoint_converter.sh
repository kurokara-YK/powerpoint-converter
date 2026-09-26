#!/usr/bin/env bash
# powerpoint-converter の起動スクリプト。インストール不要（Python の依存ライブラリは無い）。
# ~/.local/bin/powerpoint-converter からシンボリックリンクで呼ばれても動くよう、実体の場所を辿る。
here="$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")"
PYTHONPATH="$here${PYTHONPATH:+:$PYTHONPATH}" exec python3 -m powerpoint_converter "$@"
