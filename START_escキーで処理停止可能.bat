@echo off
setlocal enabledelayedexpansion

:PASSWORD_INPUT
cls
echo 実行確認

set /p password="パスワードを入力してください: "

cd /d "%~dp0program"

rem パスワード検証
python new_automation.py !password! --verify
if errorlevel 1 (
    echo パスワードが違います。再度入力してください。
    timeout /t 2 > nul
    goto PASSWORD_INPUT
)

rem パスワードが正しい場合のみ実行確認を表示
echo.
echo ============================================
echo 【操作説明】
echo - Escape キーを押すと処理を強制終了します
echo - Alt + Space キーで処理を一時停止・再開できます
echo - この処理でダウンロードおよびスクリーンショットされたファイルは、全て"\Downloads\pdf"に保存されます。
echo - 規定年齢より上の応募者に関しては、後ほど別担当者が対応しますので、「未対応」フラグのまま何もしなくて大丈夫です。
echo.
echo 【確認事項】
echo - 「ダウンロード」フォルダに「pef」フォルダはありますか？
echo - Edgeの設定で、ダウンロード先のフォルダを上記の「pdf」に変更しましたか？
echo.
echo 【注意事項】
echo マウスが画面の四隅に置いてある状態で実行すると
echo エラーになります。四隅から離して再度実行してください。
echo.
echo 稀にレジュメを添えない応募者がいますが、
echo その場合は代わりに応募者情報画面を自動でスクショされます。
echo ============================================
echo.
choice /c YN /m "プログラムを実行しますか？(Y=はい / N=いいえ)"
if errorlevel 2 (
    echo プログラムを終了します
    pause
    exit
)

rem 実行確認でYesの場合のみ実行
cls
python new_automation.py !password!
pause