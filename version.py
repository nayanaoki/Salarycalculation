# -*- coding: utf-8 -*-
"""アプリのバージョン情報と更新配信設定。

リリースのたびに __version__ を上げ、配信側の version.json も同じ番号に更新する。
"""

__version__ = "1.1.0"

APP_NAME = "給与自動計算"

# 起動時に取得する version.json の URL。
# GitHub Releases を使う場合は、リポジトリ直下に version.json を置き、
# その RAW URL（または常に最新を指す固定 URL）を指定する。
#   例: https://raw.githubusercontent.com/<owner>/<repo>/main/version.json
UPDATE_URL = "https://raw.githubusercontent.com/nayanaoki/Salarycalculation/main/version.json"
