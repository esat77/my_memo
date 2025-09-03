import os
import sys
import msvcrt
import tempfile
import atexit
import ctypes
import argparse

# from pystray import Icon, Menu, MenuItem
# from PIL import Image, ImageDraw

# # アイコン画像を作成（簡易な円形）
# def create_image():
#     image = Image.new("RGB", (64, 64), "white")
#     draw = ImageDraw.Draw(image)
#     draw.ellipse((16, 16, 48, 48), fill="blue")
#     return image

# # 「終了」メニューが選択されたとき
# def on_quit(icon, item):
#     icon.stop()

# # ----------------------------------------
# # タスクトレイに表示
# # ----------------------------------------
# icon = Icon("hello")
# icon.icon = create_image()
# icon.menu = Menu(MenuItem("Exit", on_quit))

# if os.name == "nt":
#     ctypes.windll.kernel32.SetConsoleOutputCP(65001)
#     sys.stdout.reconfigure(encoding='utf-8')

# ----------------------------------------
# argparse で引数を取得（デフォルト "hello"）
# ----------------------------------------
parser = argparse.ArgumentParser(description="引数がなければデフォルトで 'hello' を表示するアプリ")
parser.add_argument("message", nargs="?", default="hello", help="表示するメッセージ（省略時は 'hello'）")
args = parser.parse_args()
message = args.message

# ----------------------------------------
# 排他制御：すでに実行中なら終了
# ----------------------------------------
lock_file_path = os.path.join(tempfile.gettempdir(), 'my_unique_app.lock')

try:
    lock_file = open(lock_file_path, 'w')
    msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
except OSError:
    print("このアプリはすでに起動しています。")
    input("何かキーを押してください...")
    # icon.stop()

# 終了時にロック解除 & ファイル削除
def cleanup():
    try:
        lock_file.close()
        os.remove(lock_file_path)
    except Exception:
        pass

atexit.register(cleanup)

# ----------------------------------------
# ウィンドウを最小化
# ----------------------------------------
SW_MINIMIZE = 6
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), SW_MINIMIZE)

# ----------------------------------------
# メイン処理
# ----------------------------------------
print(message)
# icon.run()
input()

