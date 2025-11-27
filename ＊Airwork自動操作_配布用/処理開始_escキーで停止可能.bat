@echo off
setlocal enabledelayedexpansion

:PASSWORD_INPUT
cls
echo 実行確認

set /p password="パスワードを入力してください: "

cd /d "%~dp0"

rem 仮想環境のPythonパスを設定
set VENV_PYTHON=%USERPROFILE%\airwork_automation_env\Scripts\python.exe

rem 仮想環境が存在するか確認
if not exist "%VENV_PYTHON%" (
    echo.
    echo エラー: Python仮想環境が見つかりません
    echo.
    echo 先に「初回のみ_環境構築プログラム起動」フォルダ内の
    echo 「右クリック→「Powershellで実行」で起動.ps1」を実行してください。
    echo.
    pause
    exit /b 1
)

rem パスワード検証
"%VENV_PYTHON%" program\new_automation.py !password! --verify
if errorlevel 1 (
    echo パスワードが違います。再度入力してください。
    timeout /t 2 > nul
    goto PASSWORD_INPUT
)

rem パスワードが正しい場合のみ実行確認を表示
echo.
echo ============================================
echo 【説明】
echo - Escape キーを押すと処理を即座に終了します
echo - Alt + Space キーで処理を一時停止・再開できます
echo - この処理でダウンロードおよびスクリーンショットされたファイルは、全て"\Downloads\pdf"に保存されます。
echo - 規定年齢以上により担当者へ共有されなかった応募者については、あとで担当者が本人まで直接連絡をいれます。「未対応」フラグのままにしておいてください。
echo.
echo 【確認事項】
echo - 「ダウンロード」フォルダに「pdf」フォルダはありますか？
echo - Edgeの設定で、ダウンロード先のフォルダを上記「pdf」に変更しましたか？
echo   詳細は同梱の解説動画（冒頭シーン）をご確認ください。
echo.
echo 【注意事項】
echo マウスを画面の四隅に置いている状態で実行すると
echo エラーになります。四隅から離して実行してください。
echo ============================================
echo.
choice /c YN /m "プログラムを実行しますか？(Y=はい / N=いいえ)"
if errorlevel 2 (
    echo プログラムを終了します
    pause
    exit
)

rem 実行確認がYesの場合のみ実行
cls
"%VENV_PYTHON%" program\new_automation.py !password!
pause
