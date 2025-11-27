@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

echo ================================================
echo config.yaml 暗号化ツール
echo ================================================
echo.
echo config.yaml を暗号化して config.enc を生成します
echo.
echo ================================================
echo.

cd /d "%~dp0program"

:MENU
echo ----------------------------------------
echo [1] 暗号化
echo     - config.yaml から config.enc を作成
echo.
echo [2] 復号（確認用）
echo     - config.enc から config.yaml を復元
echo.
echo [3] 終了
echo ----------------------------------------
echo.
set /p choice="番号を入力してください (1-3): "

if "%choice%"=="1" goto ENCRYPT
if "%choice%"=="2" goto DECRYPT
if "%choice%"=="3" goto END
echo.
echo 無効な選択です。もう一度選択してください。
echo.
goto MENU

:ENCRYPT
echo.
echo ------------------------------------------------
echo 暗号化を開始します
echo ------------------------------------------------
python encrypt_config.py encrypt
echo.
if errorlevel 1 (
    echo エラーが発生しました。
) else (
    echo.
    echo [成功] 暗号化が完了しました
    echo.
    echo 次のステップ:
    echo 1. program\config.enc ファイルをメンバーに共有
    echo 2. パスワードは別途、安全な方法で伝達
    echo 3. program\config.yaml は削除せず安全な場所に保管
)
echo.
pause
goto MENU

:DECRYPT
echo.
echo ------------------------------------------------
echo 復号を開始します
echo ------------------------------------------------
python encrypt_config.py decrypt
echo.
if errorlevel 1 (
    echo エラーが発生しました。
) else (
    echo [成功] 復号が完了しました
)
echo.
pause
goto MENU

:END
echo.
echo ツールを終了します
echo.
pause
exit

