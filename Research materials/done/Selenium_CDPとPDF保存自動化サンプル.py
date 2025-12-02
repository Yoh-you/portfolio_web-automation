import logging
import os
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
from urllib.parse import urlparse

import pandas as pd
import pyautogui # (後で削除可能だが、念のため残す)
import tkinter as tk
import yaml
from retry import retry
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from tkinter import messagebox
# WebDriverのサービス（CDP実行に重要）
from selenium.webdriver.edge.service import Service as EdgeService 

# win32comのインポート（Outlook操作用）
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    win32com = None
    HAS_WIN32COM = False

# ==============================================================================
# 設定・データクラス
# ==============================================================================

@dataclass
class AutomationConfig:
    """設定ファイル (config.yaml) の構造を定義するデータクラス"""
    edge_path: str
    download_folder: str
    template_path: str
    url: str
    wait_time: Dict[str, int]
    login_info: Dict[str, str]
    webdriver_path: Optional[str] = None
    secrets_file: str = "secrets.yaml"


class AutomationError(Exception):
    """カスタム例外クラス"""
    pass


# 仮のパスワードチェック関数（本番環境に合わせて調整してください）
def check_password(password: str) -> bool:
    # 実際にはハッシュ化されたパスワードとの比較を行う
    return password == "hoge_password"


# ==============================================================================
# メインの自動化クラス
# ==============================================================================

class AutomationScript:
    """
    CDP (Chrome DevTools Protocol) を使用して、ダウンロードフォルダを設定し、
    PDFファイルのURLを直接使って安定的にダウンロードする自動化スクリプト。
    """
    def __init__(self, config_path: str):
        self.setup_logging()
        
        # 設定のロード
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = yaml.safe_load(f)
        self.config = AutomationConfig(**config_data)

        self.driver: Optional[webdriver.Edge] = None
        self.browser_wait: Optional[WebDriverWait] = None
        self.download_dir = Path(self.config.download_folder).expanduser().resolve()
        
        # ダウンロードフォルダの作成
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.logger.info(f"ダウンロードフォルダ: {self.download_dir}")
        
        # 強制終了と一時停止のためのモニタリングスレッドを開始（コードのロジックを維持）
        self.running = True
        self.paused = False
        self.pause_lock = threading.Lock()
        threading.Thread(target=self._monitor_esc, daemon=True).start()
        
        
    def setup_logging(self) -> None:
        """ロギング設定を初期化する"""
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG) 
        log_handler = logging.StreamHandler()
        log_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
        self.logger.addHandler(log_handler)
        
    def _monitor_esc(self) -> None:
        """ESCキーが押されたら強制終了する"""
        try:
            import keyboard
            while self.running:
                if keyboard.is_pressed('esc'):
                    self.logger.warning("ESCキー検知：強制終了")
                    self.cleanup()
                    os._exit(0)
                time.sleep(0.3)
        except ImportError:
            self.logger.warning("keyboardライブラリが見つかりません。ESCによる強制終了は無効です。")
        except Exception as e:
            self.logger.error(f"ESCモニタリング中にエラーが発生: {e}")
            
    def wait_if_paused(self) -> None:
        """一時停止状態の場合に待機する"""
        while self.paused:
            time.sleep(0.5)


    def _get_webdriver_options(self) -> EdgeOptions:
        """Edge WebDriverのオプションを設定する"""
        options = EdgeOptions()
        if self.config.edge_path:
            options.binary_location = self.config.edge_path
        
        options.add_argument("--start-maximized")
        # ダウンロード設定: ダウンロードダイアログを表示しない (CDPで制御するため)
        options.add_experimental_option(
            "prefs", {
                "download.prompt_for_download": False,
                # CDPでディレクトリを設定するため、ここでは特に設定しないが、
                # ブラウザのデフォルト設定を上書きしたい場合に以下を設定
                # "download.default_directory": str(self.download_dir),
            }
        )
        return options

    
    def initialize_driver(self) -> None:
        """WebDriverを初期化し、CDPによるダウンロード設定を行う"""
        try:
            options = self._get_webdriver_options()
            
            # WebDriverのパスが指定されている場合はサービスを設定
            service = EdgeService(self.config.webdriver_path) if self.config.webdriver_path else EdgeService()

            self.driver = webdriver.Edge(options=options, service=service)
            self.driver.set_page_load_timeout(60)
            self.browser_wait = WebDriverWait(self.driver, self.config.wait_time.get('browser', 15))
            
            self.driver.get(self.config.url)
            self.logger.info("WebDriver の起動と初期URLへのアクセスが完了しました")

            # 📌 CDP: ダウンロード設定の初期化
            # ダウンロード先フォルダを設定し、プロンプト表示を無効にする (重要)
            self.driver.execute_cdp_cmd(
                "Page.setDownloadBehavior", {
                    "behavior": "allow",
                    "downloadPath": str(self.download_dir),
                    "eventsEnabled": True,
                }
            )
            self.logger.info(f"CDPダウンロード動作を設定しました。パス: {self.download_dir}")

        except Exception as e:
            self.logger.error("WebDriverの初期化中にエラーが発生しました")
            self.logger.exception(e)
            raise AutomationError(f"WebDriverの初期化失敗: {e}")


    def login(self) -> None:
        """ログイン処理の模擬（実際のロジックに置き換えてください）"""
        # 以前のコードのロジックを維持
        wait = self.browser_wait
        if not wait:
            raise AutomationError("WebDriverが初期化されていません")
            
        self.logger.info("ログイン処理を開始します")
        
        # 1. ログインページへの遷移 (例: ボタンクリック)
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='__next']/div/main/div[2]/div[2]/a"))).click()
            time.sleep(self.config.wait_time.get('click', 2))
        except Exception:
            self.logger.warning("ログイン遷移ボタンが見つかりません。すでにログイン画面にいると仮定します。")

        # 2. ユーザー名とパスワードの入力
        try:
            user_input = wait.until(EC.presence_of_element_located((By.ID, "account")))
            user_input.clear()
            user_input.send_keys(self.config.login_info['username'])
            
            pwd_input = wait.until(EC.presence_of_element_located((By.ID, "password")))
            pwd_input.clear()
            pwd_input.send_keys(self.config.login_info['password'])
            
            # 3. ログインボタンのクリック
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mainContent']/div/div[2]/div[4]/input"))).click()
            time.sleep(self.config.wait_time.get('browser', 6))
            self.logger.info("ログイン処理を完了しました")
        except Exception as e:
            self.logger.error("ログイン要素の特定または操作中にエラーが発生しました。")
            self.logger.exception(e)
            raise AutomationError("ログイン処理失敗")


    def navigate_entries(self) -> None:
        """応募者一覧画面へ遷移する"""
        if not self.browser_wait:
            raise AutomationError("WebDriverが未初期化です")
        self.logger.info("応募者一覧へ遷移します")
        # 例: ナビゲーションリンクをクリック
        self.browser_wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='__next']/header/div/nav/ul/li[3]/a"))).click()
        time.sleep(self.config.wait_time.get('browser', 4))

    
    def search_and_open(self, full_name: str) -> str:
        """応募者を検索し、詳細画面を開き、レジュメタブを開く"""
        if not self.browser_wait or not self.driver:
            raise AutomationError("WebDriverが初期化されていません")
            
        wait = self.browser_wait
        self.logger.info(f"応募者を検索します: {full_name}")
        
        # 1. 検索
        search_box = wait.until(EC.presence_of_element_located((By.NAME, "searchWord")))
        search_box.clear()
        search_box.send_keys(full_name)
        
        # 2. 検索ボタンをクリック
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='applicationList']/form/div/button"))).click()
        time.sleep(self.config.wait_time.get('click', 2))
        
        # 3. 検索結果の最初の行をクリックして詳細を開く
        try:
            first_row = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "table tbody tr:first-child"))
            )
            first_row.click()
            self.logger.info("最初の行をクリックして詳細を開きました")
        except Exception:
            self.logger.warning("行クリックに失敗。詳細画面への別の遷移ロジックが必要です。")
            raise AutomationError("詳細画面への遷移に失敗")

        # 4. レジュメ（PDF）を開くボタンの href 属性を取得
        # 新しいタブを開く操作は不要（URLを直接使うため）
        resume_button = wait.until(EC.presence_of_element_located((By.XPATH, "//a[@data-la='entry_detail_resume_btn_click']")))
        pdf_url = resume_button.get_attribute("href")
        
        if not pdf_url:
            raise AutomationError("レジュメボタンのURL(href)を取得できませんでした")
            
        self.logger.info(f"レジュメPDFのURLを取得しました: {pdf_url[:80]}...")
        
        return pdf_url


    def download_pdf_from_url(self, pdf_url: str, file_name: str) -> Path:
        """
        CDPでダウンロード設定されたブラウザのセッションを利用して、
        PDFファイルのURLから直接ダウンロードを実行する。
        """
        if not self.driver:
            raise AutomationError("WebDriverが未初期化です")
            
        self.logger.info(f"URLからのPDFダウンロードを開始します: {file_name}")
        
        # ファイルがダウンロードされるのを待つために、既存のファイルをチェック
        initial_files = set(os.listdir(self.download_dir))
        
        # ダウンロード開始前にブラウザがダウンロード時に使用するであろうファイル名を予測
        # URLの最後の部分または提供された file_name から一時ファイル名を予測
        guessed_browser_name = Path(urlparse(pdf_url).path).name
        # URLのファイル名に日本語が含まれる場合、ブラウザがURLエンコードして保存することがあるため
        # 確実な予測は難しいため、ダウンロードフォルダの変更をモニタリングする

        # 1. ダウンロードをトリガーするためにPDFのURLに遷移
        # ダウンロードはメインウィンドウで行うため、タブ切り替えは不要
        self.driver.get(pdf_url)
        
        # EdgeがPDFをダウンロードファイルとして処理するのを待機
        time.sleep(self.config.wait_time.get('download_start', 3)) 

        # 2. ダウンロード完了を待機し、ファイルを見つける
        timeout = self.config.wait_time.get('download_complete', 60)
        start_time = time.time()
        final_file_path: Optional[Path] = None

        while time.time() - start_time < timeout:
            current_files = set(os.listdir(self.download_dir))
            new_files = current_files - initial_files
            
            # .crdownload や .tmp などの一時ファイルを無視し、確定したファイルを探す
            completed_downloads = [
                self.download_dir / f 
                for f in new_files 
                if f.lower().endswith('.pdf') and not f.lower().endswith('.crdownload') and not f.lower().endswith('.tmp')
            ]
            
            if completed_downloads:
                # ダウンロードされたファイルを見つけた
                final_file_path = completed_downloads[0]
                self.logger.info(f"ダウンロード完了を確認しました: {final_file_path.name}")
                break
            
            time.sleep(1)
        
        if final_file_path is None:
            self.logger.error("PDFのダウンロードがタイムアウトしました。")
            raise AutomationError("PDFダウンロード失敗 (タイムアウト)")

        # 3. ファイル名の変更（整理のため）
        safe_file_name = "".join(c for c in file_name if c.isalnum() or c in (' ', '.', '_', '-')).strip()
        new_path = self.download_dir / f"{safe_file_name}.pdf"
        
        # リネーム
        try:
            final_file_path.rename(new_path)
            self.logger.info(f"ファイル名を変更しました: {new_path.name}")
            return new_path
        except Exception as e:
            self.logger.warning(f"ファイル名変更に失敗しました ({final_file_path.name} -> {new_path.name}): {e}。元の名前で続行します。")
            return final_file_path # 変更できなくても、ダウンロードは成功


    def save_screenshot(self, applicant_id: str) -> None:
        """エラー発生時のスクリーンショットを保存する"""
        if self.driver:
            screenshot_path = self.download_dir / f"error_{applicant_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            self.driver.save_screenshot(str(screenshot_path))
            self.logger.info(f"エラー発生時のスクリーンショットを保存しました: {screenshot_path}")


    def send_email(self, email_address: str, subject: str, body_text: str, attachment_path: Path) -> None:
        """Outlookを使用してメールを送信する模擬"""
        if not HAS_WIN32COM:
            self.logger.warning("win32comがインストールされていません。メール送信処理をスキップします。")
            return
            
        self.logger.info(f"メール送信処理を開始します: To={email_address}")
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0) # 0=olMailItem
            mail.To = email_address
            mail.Subject = subject
            mail.Body = body_text
            
            if attachment_path.exists():
                mail.Attachments.Add(str(attachment_path))
            
            mail.Send()
            self.logger.info("メールを送信しました")
        except Exception as e:
            self.logger.error("Outlookメール送信中にエラーが発生しました")
            self.logger.exception(e)


    def process_data(self) -> None:
        """メイン処理ループ"""
        # 以前のコードのロジックを維持
        csv_files = list(self.download_dir.glob("応募一覧_*.csv"))
        if not csv_files:
            self.logger.error("処理するCSVファイルが見つかりませんでした。")
            return

        latest_csv = max(csv_files, key=os.path.getctime)
        self.logger.info(f"CSVを読み込みます: {latest_csv}")

        # PandasでCSVを読み込み（エンコーディングは環境に合わせて調整）
        df = pd.read_csv(latest_csv, encoding='shift_jis')
        df = df.iloc[:, [1, 4, 8, 29, 36]]
        df.columns = ["B", "E", "I", "AD", "AK"] 
        data_to_process = df.to_dict('records')
        
        self.logger.info(f"応募者を検索します... 処理対象: {len(data_to_process)}件")

        # 処理開始時のメインタブのハンドルを取得
        main_handle = self.driver.window_handles[0] if self.driver else None

        for i, row in enumerate(data_to_process):
            self.wait_if_paused()
            applicant_id = row.get("B", f"UNKNOWN_{i}")
            applicant_name = row.get("E", "不明な応募者")
            self.logger.info(f"--- 処理 {i+1}/{len(data_to_process)}: 応募者ID={applicant_id}, 氏名={applicant_name} ---")
            
            try:
                # 1. 応募者の詳細画面を開き、PDFのURLを取得
                # この操作でメインウィンドウのURLがPDF URLに変わることを許容
                pdf_url = self.search_and_open(applicant_name)
                
                # 2. PDFファイルを直接ダウンロード
                saved_pdf_path = self.download_pdf_from_url(
                    pdf_url=pdf_url,
                    file_name=f"レジュメ_{applicant_id}_{applicant_name}"
                )
                
                # 3. メール送信
                self.send_email(
                    email_address=row.get("AD", ""), 
                    subject=row.get("AK", f"レジュメ送付（{applicant_name}）"), 
                    body_text=row.get("I", ""),
                    attachment_path=saved_pdf_path
                )
                
            except AutomationError as e:
                self.logger.error(f"応募者 {applicant_id} の処理中に自動化エラーが発生しました。エラー={e}")
                self.save_screenshot(applicant_id)
            except Exception as exc:
                self.logger.error(f"応募者 {applicant_id} の処理中に予期せぬエラーが発生しました")
                self.logger.exception(exc)
                self.save_screenshot(applicant_id)
            finally:
                # ダウンロード後、必ず元の応募者一覧ページに戻る
                if self.driver:
                    self.driver.get(self.config.url) # ホームに戻るか、
                    self.navigate_entries() # 応募者一覧に戻る
                    time.sleep(self.config.wait_time.get('browser', 3))


    def cleanup(self) -> None:
        """終了処理"""
        self.running = False
        if self.driver:
            self.logger.info("WebDriverを終了します")
            self.driver.quit()

    def run(self) -> None:
        """スクリプト全体を実行する"""
        try:
            self.initialize_driver()
            self.login()
            self.navigate_entries()
            
            self.logger.info("--- データ処理を開始します ---")
            self.process_data()
            self.logger.info("スクリプトが正常に完了しました。")
            
        except AutomationError:
            self.logger.error("自動化スクリプトの実行が停止しました。")
        except Exception as exc:
            self.logger.error("処理中に致命的なエラーが発生しました")
            self.logger.exception(exc)
        finally:
            self.cleanup()


def main() -> None:
    # パスワードチェックのロジックは、元のコードに倣って維持します
    if len(sys.argv) < 2:
        print("パスワードが必要です")
        sys.exit(1)
    
    password = sys.argv[1]
    
    if not check_password(password):
        print("パスワードが違います")
        sys.exit(1)
    
    # 設定ファイルのパス
    config_path = Path(__file__).parent / "config.yaml"
    
    if not config_path.exists():
        print(f"設定ファイルが見つかりません: {config_path}")
        sys.exit(1)

    automation = AutomationScript(str(config_path))
    automation.run()


if __name__ == "__main__":
    main()