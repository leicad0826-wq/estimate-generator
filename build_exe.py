"""
PyInstallerでExeをビルドするスクリプト
使い方: python build_exe.py
"""
import PyInstaller.__main__
import streamlit
import os

streamlit_dir = os.path.dirname(streamlit.__file__)

PyInstaller.__main__.run([
    'launcher.py',
    '--name=カードラボ最終見積書',
    '--onedir',
    '--console',
    # アプリ本体とアセット
    '--add-data=app.py;.',
    '--add-data=mascot.png;.',
    # Streamlit全体をバンドル
    f'--add-data={streamlit_dir};streamlit',
    # hidden imports
    '--hidden-import=streamlit',
    '--hidden-import=streamlit.runtime.scriptrunner',
    '--hidden-import=streamlit.web.cli',
    '--hidden-import=streamlit.runtime.caching',
    '--hidden-import=openpyxl',
    '--hidden-import=lxml',
    '--hidden-import=lxml.etree',
    '--hidden-import=PIL',
    '--hidden-import=numpy',
    '--hidden-import=pyxlsb',
    '--hidden-import=pyarrow',
    '--hidden-import=altair',
    '--hidden-import=pydeck',
    '--hidden-import=toml',
    '--hidden-import=click',
    '--hidden-import=tornado',
    '--hidden-import=watchdog',
    '--hidden-import=jinja2',
    '--hidden-import=protobuf',
    '--hidden-import=cachetools',
    '--hidden-import=gitdb',
    '--hidden-import=tenacity',
    '--hidden-import=narwhals',
    '--hidden-import=jsonschema',
    '--collect-all=streamlit',
    '--collect-all=altair',
    '--noconfirm',
    '--clean',
])
