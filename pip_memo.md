pip / setuptools を 3.7 対応の最終版に更新
python -m pip install --upgrade "pip<21.0" "setuptools<60" "wheel"

proxy を環境変数で設定（PowerShell の例）
setx HTTP_PROXY  http://192.168.1.1:3550
setx HTTPS_PROXY http://192.168.1.1:3550

Nuitka のバージョンを指定してインストール
pip install nuitka==0.6.19
