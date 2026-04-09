"""
カードラボ「最終見積書」自動作成 ランチャー
ダブルクリックでStreamlitアプリを起動し、ブラウザで開く
"""
import sys
import os
import threading
import webbrowser
import time

def main():
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    app_path = os.path.join(base_dir, 'app.py')
    port = 8501

    print('=' * 50)
    print('  カードラボ「最終見積書」自動作成')
    print('  起動中...')
    print('=' * 50)
    print()

    # 3秒後にブラウザを開く
    def open_browser():
        time.sleep(3)
        webbrowser.open(f'http://localhost:{port}')
        print(f'  ブラウザで開きました: http://localhost:{port}')
        print()
        print('  ※ このウィンドウを閉じるとアプリが終了します')

    threading.Thread(target=open_browser, daemon=True).start()

    # Streamlitを直接起動（subprocess不要）
    os.chdir(base_dir)
    os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
    os.environ['STREAMLIT_SERVER_PORT'] = str(port)
    os.environ['STREAMLIT_SERVER_ADDRESS'] = 'localhost'
    os.environ['STREAMLIT_BROWSER_GATHER_USAGE_STATS'] = 'false'
    os.environ['STREAMLIT_THEME_PRIMARY_COLOR'] = '#d4549c'
    os.environ['STREAMLIT_SERVER_FILE_WATCHER_TYPE'] = 'none'
    os.environ['STREAMLIT_GLOBAL_DEVELOPMENT_MODE'] = 'false'

    from streamlit.web import cli as stcli
    sys.argv = ['streamlit', 'run', app_path]
    stcli.main()

if __name__ == '__main__':
    main()
