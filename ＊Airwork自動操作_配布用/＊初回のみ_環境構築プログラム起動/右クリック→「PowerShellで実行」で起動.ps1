# ========================================
# Python 環境自動セットアップスクリプト
# ========================================
# 
# 【目的】
# 他のメンバーでも、ローカルで Python 環境を安全に自動セットアップできるようにする。
#
# 【処理内容】
# 1. Pythonインストーラーの存在確認
# 2. ユーザー同意の取得
# 3. サイレントインストール（PATH追加）
# 4. 仮想環境(venv)作成
# 5. requirements.txtのインストール
# 6. 完了メッセージ表示
#
# 【実行方法】
# このファイルを右クリック → 「PowerShellで実行」
#
# ========================================

# エラー時に停止
$ErrorActionPreference = "Stop"

# 文字コード設定（日本語対応）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# カラー出力用関数
function Write-ColorOutput($ForegroundColor) {
    $fc = $host.UI.RawUI.ForegroundColor
    $host.UI.RawUI.ForegroundColor = $ForegroundColor
    if ($args) {
        Write-Output $args
    }
    $host.UI.RawUI.ForegroundColor = $fc
}

# ヘッダー表示
Write-Host ""
Write-ColorOutput Yellow "=========================================="
Write-ColorOutput Yellow "  Python 環境自動セットアップ"
Write-ColorOutput Yellow "=========================================="
Write-Host ""

# スクリプトのディレクトリを取得
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "作業ディレクトリ: $scriptDir"
Write-Host ""

# Pythonインストーラーのパス
$pythonInstallerName = "python-3.12.10-amd64.exe"
$pythonInstallerPath = Join-Path $scriptDir $pythonInstallerName

# インストーラーの存在確認
Write-Host "Pythonインストーラーを確認中..."
if (-Not (Test-Path $pythonInstallerPath)) {
    Write-ColorOutput Red "エラー: Pythonインストーラーが見つかりません"
    Write-Host ""
    Write-Host "以下のファイルを同じフォルダに配置してください："
    Write-Host "  - $pythonInstallerName"
    Write-Host ""
    Write-Host "配置場所: $scriptDir"
    Write-Host ""
    Read-Host "Enterキーを押して終了してください"
    exit 1
}
Write-ColorOutput Green "? インストーラーを確認しました"
Write-Host ""

# ユーザー同意の取得
Write-ColorOutput Cyan "【重要】このスクリプトは以下の処理を実行します："
Write-Host "  1. Python 3.12.10 のインストール"
Write-Host "  2. 仮想環境(venv)の作成"
Write-Host "  3. 必要なパッケージのインストール"
Write-Host ""
Write-Host "インストール先: C:\Users\$env:USERNAME\AppData\Local\Programs\Python\Python312"
Write-Host "仮想環境作成先: C:\Users\$env:USERNAME\airwork_automation_env"
Write-Host ""

$confirmation = Read-Host "Pythonをインストールしてもよろしいですか？ (y/n)"
if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
    Write-ColorOutput Yellow "セットアップをキャンセルしました"
    Read-Host "Enterキーを押して終了してください"
    exit 0
}

Write-Host ""
Write-ColorOutput Cyan "=========================================="
Write-ColorOutput Cyan "  セットアップを開始します"
Write-ColorOutput Cyan "=========================================="
Write-Host ""

# Python がすでにインストールされているか確認
Write-Host "既存のPythonインストールを確認中..."
$pythonPath = "C:\Users\$env:USERNAME\AppData\Local\Programs\Python\Python312\python.exe"
$pythonInstalled = Test-Path $pythonPath

if ($pythonInstalled) {
    Write-ColorOutput Yellow "? Python 3.12 は既にインストールされています"
    Write-Host ""
} else {
    # Pythonのサイレントインストール
    Write-Host "Python 3.12.10 をインストール中..."
    Write-Host "（この処理には数分かかる場合があります）"
    Write-Host ""
    
    try {
        $installArgs = @(
            "/quiet",
            "InstallAllUsers=0",
            "PrependPath=1",
            "Include_test=0",
            "Include_pip=1",
            "Include_doc=0",
            "Include_launcher=1",
            "InstallLauncherAllUsers=0"
        )
        
        $process = Start-Process -FilePath $pythonInstallerPath -ArgumentList $installArgs -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-ColorOutput Green "? Pythonのインストールが完了しました"
            Write-Host ""
        } else {
            throw "インストーラーがエラーコード $($process.ExitCode) で終了しました"
        }
    } catch {
        Write-ColorOutput Red "エラー: Pythonのインストールに失敗しました"
        Write-Host $_.Exception.Message
        Read-Host "Enterキーを押して終了してください"
        exit 1
    }
}

# Pythonパスの確認
if (-Not (Test-Path $pythonPath)) {
    Write-ColorOutput Red "エラー: Pythonが正しくインストールされませんでした"
    Write-Host "パス: $pythonPath"
    Read-Host "Enterキーを押して終了してください"
    exit 1
}

# 仮想環境のパス
$venvPath = "C:\Users\$env:USERNAME\airwork_automation_env"
$venvPython = Join-Path $venvPath "Scripts\python.exe"
$venvPip = Join-Path $venvPath "Scripts\pip.exe"

# 仮想環境の作成
Write-Host "仮想環境を作成中..."
if (Test-Path $venvPath) {
    Write-ColorOutput Yellow "既存の仮想環境が見つかりました。削除して再作成します..."
    Remove-Item -Path $venvPath -Recurse -Force
}

try {
    & $pythonPath -m venv $venvPath
    Write-ColorOutput Green "? 仮想環境を作成しました"
    Write-Host "  場所: $venvPath"
    Write-Host ""
} catch {
    Write-ColorOutput Red "エラー: 仮想環境の作成に失敗しました"
    Write-Host $_.Exception.Message
    Read-Host "Enterキーを押して終了してください"
    exit 1
}

# requirements.txt の確認
$requirementsPath = Join-Path $scriptDir "requirements.txt"
if (-Not (Test-Path $requirementsPath)) {
    Write-ColorOutput Red "エラー: requirements.txt が見つかりません"
    Write-Host "場所: $requirementsPath"
    Read-Host "Enterキーを押して終了してください"
    exit 1
}

# pipのアップグレード
Write-Host "pip をアップグレード中..."
try {
    & $venvPython -m pip install --upgrade pip --quiet
    Write-ColorOutput Green "? pip をアップグレードしました"
    Write-Host ""
} catch {
    Write-ColorOutput Yellow "警告: pip のアップグレードに失敗しましたが、続行します"
    Write-Host ""
}

# パッケージのインストール
Write-Host "必要なパッケージをインストール中..."
Write-Host "（この処理には数分かかる場合があります）"
Write-Host ""

try {
    & $venvPip install -r $requirementsPath
    Write-Host ""
    Write-ColorOutput Green "? パッケージのインストールが完了しました"
    Write-Host ""
} catch {
    Write-ColorOutput Red "エラー: パッケージのインストールに失敗しました"
    Write-Host $_.Exception.Message
    Read-Host "Enterキーを押して終了してください"
    exit 1
}

# 完了メッセージ
Write-Host ""
Write-ColorOutput Green "=========================================="
Write-ColorOutput Green "  セットアップが完了しました！"
Write-ColorOutput Green "=========================================="
Write-Host ""
Write-Host "【インストール情報】"
Write-Host "  Python: $pythonPath"
Write-Host "  仮想環境: $venvPath"
Write-Host ""
Write-Host "【次のステップ】"
Write-Host "  1. このフォルダの1つ上の階層に戻ってください"
Write-Host "  2. 「処理開始_escキーで停止可能.bat」をダブルクリック"
Write-Host "  3. パスワードを入力してプログラムを実行"
Write-Host ""
Write-Host "【手動で実行する場合】"
Write-Host "  $venvPython program\new_automation.py [パスワード]"
Write-Host ""
Write-ColorOutput Cyan "詳細は README_配布用.md をご覧ください"
Write-Host ""
Read-Host "Enterキーを押して終了してください"

