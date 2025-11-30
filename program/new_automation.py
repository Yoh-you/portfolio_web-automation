import logging
import os
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import pyautogui
import tkinter as tk
from pywinauto import Desktop, timings
from pywinauto.findwindows import ElementNotFoundError
import yaml
import keyboard
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from retry import retry
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from tkinter import messagebox

try:
    import win32com.client
except ImportError:
    win32com = None


@dataclass
class AutomationConfig:
    edge_path: str
    download_folder: str
    template_path: str
    url: str
    wait_time: Dict[str, int]
    login_info: Dict[str, str]
    webdriver_path: Optional[str] = None
    secrets_file: str = "secrets.yaml"


class AutomationError(Exception):
    pass


def check_password(password: str) -> bool:
    key = b'yCh_OE7jEoX7S9aUMuk-CCNiJT_GIfb1ZHkLO8b5jbw='
    cipher = b'gAAAAABnwUPuHzq3v84PkAwHhqeiqE2WnXY-2IxdOoZeOHfyeKVTToWfh89_8WQRTPGxJFJ40aoorfQLbb0-3pMRgX-cA2m41g=='
    f = Fernet(key)
    try:
        return password == f.decrypt(cipher).decode()
    except Exception:
        return False


class AutomationScript:
    def __init__(self, config_path: str):
        self.setup_logging()
        self.config = self.load_config(config_path)
        # pyautoguiの設定を追加
        pyautogui.PAUSE = 0.5  # 各操作間に0.5秒の待機を設定
        pyautogui.FAILSAFE = True  # フェイルセーフを有効化
        # マウスを画面中央に移動（フォーカス問題の回避）
        screen_width, screen_height = pyautogui.size()
        pyautogui.moveTo(screen_width // 2, screen_height // 2)
        
        self.driver: Optional[webdriver.Edge] = None
        self.browser_wait: Optional[WebDriverWait] = None
        self.running = True
        self.pause_lock = threading.Lock()
        self.paused = False
        threading.Thread(target=self._monitor_esc, daemon=True).start()
        threading.Thread(target=self._monitor_pause, daemon=True).start()

    def setup_logging(self) -> None:
        log_dir = Path(__file__).parent / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f"automation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [%(levelname)s] %(message)s',
            handlers=[logging.FileHandler(log_file), logging.StreamHandler()],
        )
        self.logger = logging.getLogger(__name__)

    def load_config(self, path: str) -> AutomationConfig:
        config_path = Path(path)
        if not config_path.exists():
            raise AutomationError(f"設定ファイルが見つかりません: {config_path}")
        with open(config_path, encoding='utf-8') as f:
            data = yaml.safe_load(f) or {}
        secrets_name = data.get('secrets_file') or 'secrets.yaml'
        secrets_path = config_path.parent / secrets_name
        if not secrets_path.exists():
            raise AutomationError(f"secretsファイルが見つかりません: {secrets_path}")
        with open(secrets_path, encoding='utf-8') as f:
            secrets = yaml.safe_load(f) or {}
        for required in ('url', 'login_info'):
            if required not in secrets:
                raise AutomationError(f"secrets.yaml に {required} が必要です")
            data[required] = secrets[required]
        return AutomationConfig(**data)

    def _monitor_esc(self) -> None:
        while self.running:
            if keyboard.is_pressed('esc'):
                self.logger.warning("ESCキー検知：強制終了")
                self.cleanup()
                os._exit(0)
            time.sleep(0.3)

    def _monitor_pause(self) -> None:
        while self.running:
            if keyboard.is_pressed('alt') and keyboard.is_pressed('space'):
                with self.pause_lock:
                    self.paused = not self.paused
                    msg = "一時停止" if self.paused else "再開"
                    self.logger.info(f"▶ {msg}")
                time.sleep(0.5)
            time.sleep(0.3)

    def wait_if_paused(self) -> None:
        while self.paused:
            time.sleep(0.5)

    def start_webdriver(self) -> webdriver.Edge:
        options = EdgeOptions()
        options.use_chromium = True
        options.add_argument("--start-maximized")
        download_folder = Path(self.config.download_folder).expanduser().resolve()
        download_folder.mkdir(parents=True, exist_ok=True)
        prefs = {
            "download.default_directory": str(download_folder),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
        options.add_experimental_option("prefs", prefs)
        edge_binary = (self.config.edge_path or "").replace("%s", "").strip()
        if edge_binary:
            options.binary_location = edge_binary
        self.logger.info("WebDriver の起動を試みます")
        if self.config.webdriver_path:
            service = webdriver.EdgeService(self.config.webdriver_path)
            driver = webdriver.Edge(service=service, options=options)
        else:
            driver = webdriver.Edge(options=options)
        driver.set_page_load_timeout(60)
        return driver

    def login(self) -> None:
        if not self.browser_wait:
            raise AutomationError("WebDriverが初期化されていません")
        wait = self.browser_wait
        self.logger.info("ログイン処理を開始します")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='__next']/div/main/div[2]/div[2]/a"))).click()
        time.sleep(self.config.wait_time.get('click', 2))
        user = wait.until(EC.presence_of_element_located((By.ID, "account")))
        user.clear()
        user.send_keys(self.config.login_info['username'])
        pwd = wait.until(EC.presence_of_element_located((By.ID, "password")))
        pwd.clear()
        pwd.send_keys(self.config.login_info['password'])
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mainContent']/div/div[2]/div[4]/input"))).click()
        time.sleep(self.config.wait_time.get('browser', 6))

    def navigate_entries(self) -> None:
        if not self.browser_wait:
            raise AutomationError("WebDriverが初期化されていません")
        self.logger.info("応募者一覧へ遷移します")
        self.browser_wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='__next']/header/div/nav/ul/li[3]/a"))).click()
        time.sleep(self.config.wait_time.get('browser', 4))

    def filter_entries(self, status_value: str = "01") -> None:
        wait = self.browser_wait
        self.logger.info(f"ステータスを{status_value} に設定して検索します")
        select_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//select[@name='selectionStatus' and @data-select='selectBox']")))
        Select(select_element).select_by_value(status_value)
        self.click_search()
        time.sleep(self.config.wait_time.get('browser', 4))

    def click_search(self) -> None:
        if not self.browser_wait:
            raise AutomationError("WebDriverが未初期化です")
        self.browser_wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='applicationList']/form/div/button"))).click()
        time.sleep(self.config.wait_time.get('click', 2))

    def download_entries(self) -> None:
        if not self.browser_wait:
            raise AutomationError("WebDriverが未初期化です")
        self.logger.info("CSVダウンロードを開始します")
        # JavaScriptで直接クリックして、確実に1回だけ実行
        download_button = self.browser_wait.until(
            EC.presence_of_element_located((By.XPATH, "//button[@data-la='entries_download_btn_click']"))
        )
        # Seleniumのclick()ではなく、JavaScriptで実行して確実に1回だけ
        self.driver.execute_script("arguments[0].click();", download_button)
        time.sleep(self.config.wait_time.get('browser', 6))

    def get_latest_csv(self) -> str:
        folder = Path(self.config.download_folder).expanduser().resolve()
        self.logger.info(f"ダウンロードフォルダからCSVを探します: {folder}")
        csv_files = list(folder.glob("*.csv"))
        if not csv_files:
            raise AutomationError("CSVファイルが見つかりません")
        latest = max(csv_files, key=lambda f: f.stat().st_ctime)
        self.logger.info(f"最新CSV: {latest}")
        return str(latest)

    def process_data(self, csv_path: str) -> pd.DataFrame:
        self.logger.info(f"CSVを読み込みます: {csv_path}")
        df = pd.read_csv(csv_path)
        df = df.iloc[:, [1, 4, 8, 29, 36]]
        df.columns = ["B", "E", "I", "AD", "AK"]
        return df

    def confirm_csv_data(self, df: pd.DataFrame) -> bool:
        """B/E/I/AD/AK列の内容をダイアログで表示"""
        preview = df[['B', 'E', 'I', 'AD', 'AK']].head(5).to_string(index=False)
        root = tk.Tk()
        root.withdraw()  # メインウィンドウは非表示
        
        # ダイアログを確実に前面に表示する設定
        root.attributes('-topmost', True)  # 常に最前面に表示
        root.lift()  # ウィンドウを前面に移動
        root.focus_force()  # フォーカスを強制的に取得
        
        # ダイアログを表示
        result = messagebox.askokcancel(
            "CSVデータ確認",
            f"下記のデータで処理を開始します。\n\n{preview}\n\nOK→続行 / キャンセル→中断"
        )
        
        root.destroy()
        return result

    def search_and_open(self, full_name: str) -> None:
        if not self.browser_wait or not self.driver:
            raise AutomationError("WebDriverが初期化されていません")
        self.logger.info(f"応募者を検索します: {full_name}")
        search_box = self.browser_wait.until(EC.presence_of_element_located((By.NAME, "searchWord")))
        search_box.clear()
        search_box.send_keys(full_name)
        self.click_search()
        rows = self.browser_wait.until(EC.presence_of_all_elements_located((By.XPATH, "//td[contains(@class, 'styles_tdSelectionStatus')]")))
        if not rows:
            raise AutomationError("検索結果が見つかりません")
        self.logger.info("対応状況セルを開いて詳細画面へ遷移します")
        try:
            first_row = self.browser_wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "table tbody tr:first-child"))
            )
            first_row.click()
            self.logger.info("最初の行全体をクリックしました")
        except Exception:
            self.logger.warning("行全体のクリックに失敗したため、セルを再試行します")
            first_cell = self.browser_wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "table tbody tr:first-child td:first-child"))
            )
            first_cell.click()
            self.logger.info("セルをクリックして詳細を開きました")
        time.sleep(1)
        resume_button = self.browser_wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@data-la='entry_detail_resume_btn_click']")))
        current = list(self.driver.window_handles)
        resume_button.click()
        self.wait_for_new_tab(len(current))

    def wait_for_new_tab(self, previous_count: int, timeout: int = 15) -> None:
        if not self.driver:
            raise AutomationError("WebDriverが未初期化です")
        start = time.time()
        while time.time() - start < timeout:
            handles = self.driver.window_handles
            if len(handles) > previous_count:
                self.driver.switch_to.window(handles[-1])
                return
            time.sleep(0.5)
        raise AutomationError("新しいタブが開きません")

    def save_pdf(self) -> bool:
        self.logger.info("PDF保存処理を行います")
        for attempt in range(3):
            try:
                time.sleep(4)
                # マウスを画面中央に移動（右クリックメニューを表示するため）
                screen_width, screen_height = pyautogui.size()
                pyautogui.moveTo(screen_width // 2, screen_height // 2)
                time.sleep(0.5)
                pyautogui.click(button='right')
                time.sleep(3)
                pyautogui.press('s')  # 保存メニューを選択
                time.sleep(2)
                pyautogui.press('enter')  # 保存ダイアログが表示されるのを待つ
                time.sleep(5)
                download_folder = Path(self.config.download_folder).expanduser().resolve()
                if not self._handle_save_as_dialog(download_folder):
                    raise AutomationError("名前を付けて保存ダイアログの操作に失敗しました")
                time.sleep(self.config.wait_time.get('save_pdf', 4))
                pyautogui.hotkey('ctrl', 'Q')
                time.sleep(4)
                return True
            except Exception as exc:
                self.logger.warning(f"PDF保存失敗 attempt={attempt+1}: {exc}")
                time.sleep(1)
        return False

    def _handle_save_as_dialog(self, target_folder: Path) -> bool:
        timeout = self.config.wait_time.get('save_pdf', 5) + 10
        dialog_patterns = [
            r"^名前を付けて保存.*Microsoft Edge",
            r"^名前を付けて保存.*",
            r".*Save As.*",
        ]
        self.logger.info(f"保存ダイアログ検出を開始 patterns={dialog_patterns} timeout={timeout}")
        save_dialog = self._wait_for_save_dialog(dialog_patterns, timeout)
        if save_dialog is None:
            self.logger.warning("保存ダイアログ接続失敗: ダイアログが見つかりませんでした")
            return False

        file_name_edit = self._find_control(
            save_dialog,
            [
                {"title": "ファイル名:", "control_type": "Edit"},
                {"title": "File name:", "control_type": "Edit"},
                {"control_type": "Edit", "found_index": 0},
            ],
        )
        if file_name_edit is None:
            self.logger.warning("ファイル名欄が見つかりませんでした")
            return False

        try:
            default_name = file_name_edit.texts()[0]
            self.logger.debug(f"保存ダイアログ既定ファイル名: {default_name}")
        except (IndexError, AttributeError):
            default_name = "download.pdf"
            self.logger.debug("保存ダイアログ既定ファイル名取得に失敗、download.pdf を使用")
        target_path = target_folder / default_name
        self.logger.info(f"保存先パスを設定: {target_path}")
        file_name_edit.set_edit_text(str(target_path))

        save_button = self._find_control(
            save_dialog,
            [
                {"title": "保存", "control_type": "Button"},
                {"title": "Save", "control_type": "Button"},
                {"control_type": "Button", "found_index": 0},
            ],
        )
        if save_button is None:
            self.logger.warning("保存ボタンが見つかりませんでした")
            return False

        try:
            save_button.click()
            return True
        except Exception as exc:
            self.logger.warning(f"保存ボタンのクリックに失敗しました: {exc}")
            return False

    def _wait_for_save_dialog(self, title_patterns, timeout: int):
        deadline = time.time() + timeout
        attempt = 1
        while time.time() < deadline:
            desktop = Desktop(backend="uia")
            for pattern in title_patterns:
                try:
                    self.logger.debug(f"[保存ダイアログ探索] attempt={attempt} pattern={pattern}")
                    dialog = desktop.window(title_re=pattern)
                    dialog.wait("visible", timeout=0.5)
                    self.logger.info(f"保存ダイアログ検出成功 pattern={pattern}")
                    return dialog.wrapper_object()
                except ElementNotFoundError:
                    continue
                except timings.TimeoutError:
                    continue
                except Exception as exc:
                    self.logger.debug(f"保存ダイアログ取得中に例外: {exc}")
                    continue
            attempt += 1
            time.sleep(0.5)
        self.logger.warning("保存ダイアログ検出に失敗しました（タイムアウト）")
        return None

    def _find_control(self, dialog, candidates):
        for props in candidates:
            try:
                control = dialog.child_window(**props).wrapper_object()
                self.logger.debug(f"コントロール検出成功 props={props}")
                return control
            except ElementNotFoundError:
                continue
            except Exception as exc:
                self.logger.debug(f"コントロール検索で例外発生 {props}: {exc}")
        return None

    def collect_attachment(self) -> Path:
        self.logger.info("添付ファイル候補を探します")
        folder = Path(self.config.download_folder).expanduser().resolve()
        candidates = list(folder.glob("*.pdf")) + list(folder.glob("*.png"))
        if not candidates:
            raise AutomationError("添付ファイルが見つかりません")
        latest = max(candidates, key=lambda f: f.stat().st_mtime)
        return latest

    def send_email(self, branch: str, ak_value: str, applicant_address: str = "") -> None:
        if win32com is None:
            raise AutomationError("win32com がインポートできません")
        self.logger.info(f"{branch} へのメールを生成します")
        mail = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
        mail.To = ""
        mail.CC = ""
        mail.Subject = f"{branch} レジュメ送付"
        mail.Body = f"住所: {applicant_address}\nAK: {ak_value}"
        attachment = self.collect_attachment()
        mail.Attachments.Add(str(attachment))
        mail.Send()
        self.logger.info(f"メール送信: {attachment.name}")

    def close_overlay(self) -> None:
        if not self.browser_wait:
            return
        try:
            btn = self.browser_wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@data-la='overlay_entry_detail_close_btn_click']")))
            btn.click()
            time.sleep(1)
        except Exception:
            pass

    def show_dialog(self, message: str, is_error: bool = False) -> None:
        root = tk.Tk()
        root.withdraw()
        if is_error:
            messagebox.showerror("エラー", message)
        else:
            messagebox.showinfo("通知", message)
        root.destroy()

    def save_screenshot(self, prefix: str) -> None:
        screenshot_dir = Path("screenshots")
        screenshot_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pyautogui.screenshot(str(screenshot_dir / f"{prefix}_{timestamp}.png"))

    def cleanup(self) -> None:
        self.running = False
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass

    def run(self) -> None:
        try:
            self.driver = self.start_webdriver()
            self.browser_wait = WebDriverWait(self.driver, self.config.wait_time.get('browser', 6) + 10)
            self.driver.get(self.config.url)
            self.login()
            self.navigate_entries()
            self.filter_entries()
            self.download_entries()
            csv_path = self.get_latest_csv()
            df = self.process_data(csv_path)
            if not self.confirm_csv_data(df):
                self.show_dialog("CSV確認で中断しました", is_error=True)
                return
            for _, row in df.iterrows():
                self.wait_if_paused()
                try:
                    if int(row["E"]) >= 55:
                        self.logger.info("55歳以上のためスキップ")
                        continue
                    self.search_and_open(row["B"])
                    time.sleep(self.config.wait_time.get('browser', 3))
                    if not self.save_pdf():
                        self.save_screenshot(row["B"])
                    self.send_email(row["AD"], row["AK"], row["I"])
                finally:
                    pyautogui.hotkey('ctrl', 'w')
                    if self.driver:
                        self.driver.switch_to.window(self.driver.window_handles[0])
                    self.close_overlay()
        except Exception as exc:
            self.logger.error("処理中に致命的なエラーが発生しました")
            self.logger.exception(exc)
            raise
        finally:
            self.cleanup()


def main() -> None:
    if len(sys.argv) < 2:
        print("パスワードが必要です")
        sys.exit(1)
    
    password = sys.argv[1]
    
    # --verifyオプションが指定された場合の検証モード
    if len(sys.argv) > 2 and sys.argv[2] == "--verify":
        if not check_password(password):
            print("パスワードが違います")
            sys.exit(1)
        sys.exit(0)  # 検証成功で終了
    
    # 通常実行モード
    if not check_password(password):
        print("パスワードが違います")
        sys.exit(1)
    
    config_path = Path(__file__).parent / "config.yaml"
    automation = AutomationScript(str(config_path))
    try:
        automation.run()
    except Exception as exc:
        automation.logger.error("全体処理で致命的なエラーが発生しました")
        automation.logger.exception(exc)
        raise


if __name__ == "__main__":
    main()
