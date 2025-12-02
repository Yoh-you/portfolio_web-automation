"""
Airwork自動化スクリプト - Selenium版（実装完了）
=============================================

目的: SeleniumでAirworkのWeb操作部分を自動化

状態: ✅ 実際のAirwork要素に対応済み（2025-11-18更新）

実装済み機能:
    ✅ ログイン処理
    ✅ 応募者ページへの遷移
    ✅ 選考ステータス設定
    ✅ CSVダウンロード
    ✅ 応募者検索
    ✅ 応募者選択
    ✅ 履歴書を開く
    ✅ 詳細ページを閉じる

使用方法:
    1. AIRWORK_URL、USERNAME、PASSWORDを設定
    2. python airwork_selenium_sample.py を実行
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import logging
from pathlib import Path
import pandas as pd
import pyautogui
import glob
import os
import keyboard


# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s'
)
logger = logging.getLogger(__name__)


# 緊急停止フラグ
emergency_stop_flag = False


def check_emergency_stop():
    """
    Escapeキーが押されたかチェック
    押されていた場合は緊急停止フラグを立てる
    """
    global emergency_stop_flag
    if keyboard.is_pressed('esc'):
        emergency_stop_flag = True
        logger.warning("🛑 Escapeキーが検出されました！緊急停止します...")
        return True
    return False


class AirworkSeleniumAutomation:
    """Selenium版Airwork自動化クラス"""

    # 環境変数 %USERPROFILE% を展開してパスを取得
    user_profile = os.path.expandvars(r'%USERPROFILE%')
    
    # 目的のパスを構築
    target_path = os.path.join(user_profile, 'Downloads', 'pdf')

    def __init__(self, url: str, username: str, password: str, download_dir: str = None):
        """
        初期化
        
        Args:
            url: AirworkのURL
            username: ユーザー名
            password: パスワード
            download_dir: ダウンロード先ディレクトリ（Noneの場合はtarget_pathを使用）
        """
        self.url = url
        self.username = username
        self.password = password
        # download_dirが指定されていない場合は、クラス変数のtarget_pathを使用
        self.download_dir = download_dir if download_dir else self.target_path
        self.driver = None
        
    def start_browser(self):
        """ブラウザを起動（ダウンロード先設定付き）"""
        try:
            logger.info("ブラウザを起動中...")
            
            # Edgeオプションを設定
            options = webdriver.EdgeOptions()
            
            # ダウンロード先を指定
            if self.download_dir:
                prefs = {
                    "download.default_directory": self.download_dir,
                    "download.prompt_for_download": False,
                    "download.directory_upgrade": True,
                    "safebrowsing.enabled": True
                }
                options.add_experimental_option("prefs", prefs)
                logger.info(f"ダウンロード先を設定: {self.download_dir}")
            
            self.driver = webdriver.Edge(options=options)
            self.driver.maximize_window()  # ウィンドウを最大化
            logger.info("✓ ブラウザ起動成功")
            return True
        except Exception as e:
            logger.error(f"ブラウザ起動エラー: {str(e)}")
            return False
    
    def open_airwork(self):
        """Airworkサイトを開く"""
        try:
            logger.info(f"Airworkを開いています: {self.url}")
            self.driver.get(self.url)
            
            # ページ読み込み完了を待機
            WebDriverWait(self.driver, 10).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            
            logger.info("✓ ページ読み込み完了")
            return True
        except Exception as e:
            logger.error(f"ページ読み込みエラー: {str(e)}")
            return False
    
    def login(self):
        """
        ログイン処理
        
        実際のAirworkの要素に対応済み：
        - ログインボタン: <a class="styles_loginButton_XULR9...">
        - ユーザー名: id="account"
        - パスワード: id="password"
        - ログイン実行: <input type="submit" class="primary">
        """
        try:
            logger.info("ログイン処理を開始...")
            
            # 最初のログインボタンをクリック（トップページ）
            login_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='__next']/div/main/div[2]/div[2]/a"))
            )
            login_button.click()
            logger.info("✓ ログインボタンをクリック")
            time.sleep(2)
            
            # ユーザー名入力（id="account"）
            username_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "account"))
            )
            username_field.clear()
            username_field.send_keys(self.username)
            logger.info("✓ ユーザー名を入力")
            
            # パスワード入力（id="password"）
            password_field = self.driver.find_element(By.ID, "password")
            password_field.clear()
            password_field.send_keys(self.password)
            logger.info("✓ パスワードを入力")
            
            # ログイン実行（Submitボタンをクリック）
            submit_button = self.driver.find_element(By.XPATH, "//*[@id='mainContent']/div/div[2]/div[4]/input")
            submit_button.click()
            logger.info("✓ ログイン実行ボタンをクリック")
            
            # ログイン完了を待機（応募者メニューが表示されるまで）
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//a[@href='/entries']"))
            )
            
            logger.info("✓ ログイン成功")
            return True
            
        except TimeoutException:
            logger.error("❌ ログイン処理がタイムアウトしました")
            self.save_screenshot("login_timeout")
            return False
        except Exception as e:
            logger.error(f"❌ ログインエラー: {str(e)}")
            self.save_screenshot("login_error")
            return False
    
    def navigate_to_search_page(self):
        """
        検索ページへ遷移（image3, image4, image5に相当）
        
        実際のAirworkの要素に対応済み：
        - 応募者リンク: <a href="/entries">応募者</a>
        """
        try:
            logger.info("検索ページへ遷移中...")
            
            # 「応募者」リンクをクリック
            menu_item = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@href='/entries']"))
            )
            menu_item.click()
            time.sleep(2)
            
            logger.info("✓ 検索ページへ遷移完了")
            return True
            
        except Exception as e:
            logger.error(f"❌ ページ遷移エラー: {str(e)}")
            self.save_screenshot("navigation_error")
            return False
    
    def set_selection_status(self, status_value="01"):
        """
        選考ステータスを設定して検索
        
        実際のAirworkの要素に対応済み：
        - 選択メニュー: name="selectionStatus"
        - 選択肢: <option value="01">未対応</option>
        - 検索ボタン: class="styles_searchButton__aRKjk"
        
        Args:
            status_value: ステータス値（デフォルト: "01" = 未対応）
        """
        try:
            logger.info("選考ステータスを設定中...")
            
            # 選択メニューを見つける
            select_element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.NAME, "selectionStatus"))
            )
            
            # Selectオブジェクトを使用
            from selenium.webdriver.support.ui import Select
            select = Select(select_element)
            select.select_by_value(status_value)
            
            logger.info(f"✓ ステータスを設定: {status_value}")
            time.sleep(1)
            
            # 検索ボタンをクリック
            search_button = self.driver.find_element(
                By.XPATH, "//*[@id='applicationList']/form/div/button"
            )
            search_button.click()
            time.sleep(2)
            
            logger.info("✓ 検索実行完了")
            return True
            
        except Exception as e:
            logger.error(f"❌ ステータス設定エラー: {str(e)}")
            self.save_screenshot("status_error")
            return False
    
    def download_csv(self):
        """
        CSVダウンロード
        
        実際のAirworkの要素に対応済み：
        - ダウンロードボタン: data-la="entries_download_btn_click"
        """
        try:
            logger.info("CSVダウンロードを開始...")
            
            # ダウンロードボタンを待機してクリック
            download_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-la='entries_download_btn_click']"))
            )
            download_button.click()
            
            logger.info("✓ ダウンロードボタンをクリック")
            
            # ダウンロード完了を待機（7秒）
            time.sleep(7)
            
            logger.info("✓ CSVダウンロード完了")
            return True
            
        except Exception as e:
            logger.error(f"❌ ダウンロードエラー: {str(e)}")
            self.save_screenshot("download_error")
            return False
    
    def search_applicant(self, full_name: str):
        """
        応募者を検索
        
        実際のAirworkの要素に対応済み：
        - 検索ボックス: name="searchWord"
        - 検索ボタン: class="styles_searchButton__aRKjk"
        
        Args:
            full_name: 検索する氏名（フルネーム）
        """
        try:
            logger.info(f"応募者を検索中: {full_name}")
            
            # 検索ボックスを見つける
            search_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.NAME, "searchWord"))
            )
            
            # 検索キーワードをクリアして入力
            search_box.clear()
            search_box.send_keys(full_name)
            logger.info(f"✓ 検索キーワードを入力: {full_name}")
            
            # 検索ボタンをクリック
            search_button = self.driver.find_element(
                By.XPATH, "//*[@id='applicationList']/form/div/button"
            )
            search_button.click()
            
            # 検索結果の表示を待機
            time.sleep(2)
            
            logger.info("✓ 検索実行完了")
            return True
            
        except Exception as e:
            logger.error(f"❌ 検索エラー: {str(e)}")
            self.save_screenshot("search_error")
            return False
    
    def select_applicant(self):
        """
        検索結果から応募者を選択
        
        実際のAirworkの要素に対応済み：
        - 応募者リストの選択要素: data-select="selectBoxTable"
        """
        try:
            logger.info("応募者を選択中...")
            
            # 候補をクリック（最初の応募者の選択ボックス）
            candidate = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//select[@data-select='selectBoxTable']"))
            )
            candidate.click()
            
            time.sleep(2)
            logger.info("✓ 応募者を選択")
            return True
            
        except Exception as e:
            logger.error(f"❌ 選択エラー: {str(e)}")
            self.save_screenshot("select_error")
            return False
    
    def open_resume(self):
        """
        履歴書を開く
        
        実際のAirworkの要素に対応済み：
        - 履歴書を開くボタン: data-la="entry_detail_resume_btn_click"
        """
        try:
            logger.info("履歴書を開いています...")
            
            # 履歴書を開くボタンをクリック
            resume_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@data-la='entry_detail_resume_btn_click']"))
            )
            resume_button.click()
            
            time.sleep(3)
            logger.info("✓ 履歴書を開きました")
            
            # ここからPyAutoGUIに切り替え（PDF操作）
            logger.info("→ PyAutoGUIでPDF保存処理へ")
            
            return True
            
        except Exception as e:
            logger.error(f"❌ 履歴書オープンエラー: {str(e)}")
            self.save_screenshot("resume_open_error")
            return False
    
    def read_csv_cell_b2(self, csv_filename: str = None):
        """
        CSVファイルのB2セル（2行目、2列目）の値を取得
        
        Args:
            csv_filename: CSVファイル名（指定しない場合は最新のCSVを自動検索）
            
        Returns:
            str: B2セルの値、エラー時はNone
        """
        try:
            logger.info("CSVファイルを読み込んでいます...")
            
            # CSVファイルのパスを決定
            if csv_filename:
                csv_path = Path(self.download_dir) / csv_filename
            else:
                # 最新のCSVファイルを検索
                csv_files = glob.glob(os.path.join(self.download_dir, "*.csv"))
                if not csv_files:
                    logger.error("❌ CSVファイルが見つかりません")
                    return None
                csv_path = max(csv_files, key=os.path.getctime)
                logger.info(f"最新のCSVファイル: {csv_path}")
            
            # CSVを読み込み（エンコーディングを自動判定）
            try:
                df = pd.read_csv(csv_path, encoding='utf-8')
            except:
                try:
                    df = pd.read_csv(csv_path, encoding='shift_jis')
                except:
                    df = pd.read_csv(csv_path, encoding='cp932')
            
            # B2セルの値を取得（行1、列1 - 0-indexedなので）
            if len(df) > 0 and len(df.columns) > 1:
                b2_value = df.iloc[0, 1]  # 1行目（0-indexed）、2列目（0-indexed）
                logger.info(f"✓ B2セルの値を取得: {b2_value}")
                return str(b2_value)
            else:
                logger.error("❌ CSVファイルのサイズが不足しています")
                return None
                
        except Exception as e:
            logger.error(f"❌ CSV読み込みエラー: {str(e)}")
            self.save_screenshot("csv_read_error")
            return None
    
    def click_first_applicant_status_cell(self):
        """
        最初の応募者の行をクリック（詳細ページを開く）
        
        修正：対応状況セル内のプルダウンを避けて、行全体または応募者名セルをクリック
        """
        try:
            logger.info("最初の応募者の行をクリック中...")
            
            # 方法1: テーブルの最初の行をクリック
            try:
                first_row = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "table tbody tr:first-child"))
                )
                first_row.click()
                logger.info("✓ 応募者行をクリックしました（行全体）")
            except:
                # 方法2（フォールバック）: 最初の行の最初のセルをクリック
                logger.info("行全体のクリックに失敗、最初のセルをクリックします...")
                first_cell = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "table tbody tr:first-child td:first-child"))
                )
                first_cell.click()
                logger.info("✓ 応募者セルをクリックしました（最初のセル）")
            
            time.sleep(2)  # 詳細ページの読み込み待機
            return True
            
        except Exception as e:
            logger.error(f"❌ 応募者行クリックエラー: {str(e)}")
            self.save_screenshot("row_click_error")
            return False
    
    def select_interview_adjustment(self):
        """
        「面接調整開始」を選択
        
        実際のAirworkの要素に対応済み：
        - 選択肢: <option value="04">面接調整開始</option>
        """
        try:
            logger.info("「面接調整開始」を選択中...")
            
            # 親selectを特定し、その中のoption[value='04']を選択
            select_element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//select[@data-select='selectBoxTable']")
                )
            )
            
            from selenium.webdriver.support.ui import Select
            select = Select(select_element)
            select.select_by_value("04")
            
            time.sleep(1)
            logger.info("✓ 「面接調整開始」を選択しました")
            return True
            
        except Exception as e:
            logger.error(f"❌ 面接調整開始選択エラー: {str(e)}")
            self.save_screenshot("interview_select_error")
            return False
    
    def save_pdf_from_resume_page(self):
        """
        レジュメページでPDFを保存（PyAutoGUI操作）
        
        手順:
        1. 右クリック
        2. 下矢印キー4回
        3. Enter
        4. 3秒待機
        5. Enter
        6. Ctrl+Wでページを閉じる
        """
        try:
            logger.info("レジュメページでPDF保存操作を開始...")
            
            # 少し待機
            time.sleep(2)
            
            # 右クリック
            pyautogui.rightClick()
            logger.info("✓ 右クリック実行")
            time.sleep(0.5)
            
            # 下矢印キーを4回押す
            for i in range(4):
                pyautogui.press('down')
                time.sleep(0.2)
            logger.info("✓ 下矢印キーを4回押しました")
            
            # Enter
            pyautogui.press('enter')
            logger.info("✓ 1回目のEnter実行")
            
            # 3秒待機
            time.sleep(3)
            
            # 2回目のEnter
            pyautogui.press('enter')
            logger.info("✓ 2回目のEnter実行")
            
            # 少し待機
            time.sleep(1)
            
            # Ctrl+Wでページを閉じる
            pyautogui.hotkey('ctrl', 'w')
            logger.info("✓ Ctrl+Wでページを閉じました")
            
            time.sleep(2)
            logger.info("✓ PDF保存処理完了")
            return True
            
        except Exception as e:
            logger.error(f"❌ PDF保存エラー: {str(e)}")
            self.save_screenshot("pdf_save_error")
            return False
    
    def close_detail_page(self):
        """
        詳細ページを閉じる（resume_closeに相当）
        
        実際のAirworkの要素に対応済み：
        - 閉じるボタン: data-la="overlay_entry_detail_close_btn_click" (imgタグ)
        """
        try:
            logger.info("詳細ページを閉じています...")
            
            # 閉じるボタンをクリック（imgタグ）
            close_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//img[@data-la='overlay_entry_detail_close_btn_click']"))
            )
            close_button.click()
            
            time.sleep(2)
            logger.info("✓ 詳細ページを閉じました")
            return True
            
        except Exception as e:
            logger.error(f"❌ ページクローズエラー: {str(e)}")
            return False
    
    def save_screenshot(self, filename_prefix: str):
        """
        スクリーンショット保存（デバッグ用）
        
        Args:
            filename_prefix: ファイル名のプレフィックス
        """
        try:
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = Path(f"screenshots/{filename_prefix}_{timestamp}.png")
            screenshot_path.parent.mkdir(exist_ok=True)
            
            self.driver.save_screenshot(str(screenshot_path))
            logger.info(f"スクリーンショット保存: {screenshot_path}")
        except Exception as e:
            logger.error(f"スクリーンショット保存エラー: {str(e)}")
    
    def quit(self):
        """ブラウザを終了"""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("✓ ブラウザを終了しました")
            except Exception as e:
                logger.error(f"ブラウザ終了エラー: {str(e)}")


def main():
    """
    メイン処理（調査フロー実装済み）
    
    フロー:
    1. CSVダウンロード（指定ディレクトリへ）
    2. CSVのB2セルを取得して検索
    3. 対応状況セルをクリック（Selenium - Option B）
    4. レジュメボタンをクリック
    5. PDF保存操作（PyAutoGUI）
    6. 詳細ページを閉じる
    7. 「面接調整開始」を選択
    """
    
    # TODO: 実際の認証情報に置き換え
    AIRWORK_URL = ""
    USERNAME = ""
    PASSWORD = ""
    
    # ダウンロード先ディレクトリ（Noneを指定すると、クラス変数のtarget_pathが使用される）
    DOWNLOAD_DIR = None
    
    # ダウンロードディレクトリを作成（automationインスタンス作成後にパスが確定するため、後で作成）
    automation = AirworkSeleniumAutomation(AIRWORK_URL, USERNAME, PASSWORD, DOWNLOAD_DIR)
    
    # ダウンロードディレクトリを作成
    Path(automation.download_dir).mkdir(parents=True, exist_ok=True)
    
    try:
        logger.info("🛑 緊急停止: いつでもEscapeキーを押すと処理を中断できます")
        
        # ブラウザ起動
        if check_emergency_stop() or not automation.start_browser():
            return
        
        # Airworkを開く
        if check_emergency_stop() or not automation.open_airwork():
            return
        
        # ログイン
        if check_emergency_stop() or not automation.login():
            return
        
        # 検索ページへ遷移
        if check_emergency_stop() or not automation.navigate_to_search_page():
            return
        
        # 選考ステータスを「未対応」に設定して検索
        if check_emergency_stop() or not automation.set_selection_status("01"):
            return
        
        # 【1】CSVダウンロード
        if check_emergency_stop():
            return
        logger.info("=" * 50)
        logger.info("【処理1】CSVダウンロード開始")
        if not automation.download_csv():
            return
        
        # 【2】CSVのB2セルを読み取って検索
        if check_emergency_stop():
            return
        logger.info("=" * 50)
        logger.info("【処理2】CSVのB2セルを読み取り")
        b2_value = automation.read_csv_cell_b2()
        if not b2_value:
            logger.error("❌ B2セルの値が取得できませんでした")
            return
        
        logger.info(f"B2セルの値で検索します: {b2_value}")
        if check_emergency_stop() or not automation.search_applicant(b2_value):
            return
        
        # 【3】対応状況セルをクリック（Selenium）
        if check_emergency_stop():
            return
        logger.info("=" * 50)
        logger.info("【処理3】対応状況セルをクリック（Selenium使用）")
        if not automation.click_first_applicant_status_cell():
            return
        
        time.sleep(1)  # クリック後の待機
        
        # 【4】レジュメボタンをクリック
        if check_emergency_stop():
            return
        logger.info("=" * 50)
        logger.info("【処理4】レジュメボタンをクリック")
        if not automation.open_resume():
            return
        
        # 【5】PDF保存操作（PyAutoGUI）
        if check_emergency_stop():
            return
        logger.info("=" * 50)
        logger.info("【処理5】PDF保存操作（PyAutoGUI）")
        if not automation.save_pdf_from_resume_page():
            return
        
        # 【6】詳細ページを閉じる
        if check_emergency_stop():
            return
        logger.info("=" * 50)
        logger.info("【処理6】詳細ページを閉じる")
        automation.close_detail_page()
        
        # 【7】面接調整開始を選択
        if check_emergency_stop():
            return
        logger.info("=" * 50)
        logger.info("【処理7】「面接調整開始」を選択")
        if not automation.select_interview_adjustment():
            logger.warning("⚠️ 面接調整開始の選択に失敗しましたが、処理を続行します")
        
        logger.info("=" * 50)
        logger.info("✅ 全処理完了！調査成功！")
        logger.info("=" * 50)
        
    except Exception as e:
        logger.error(f"❌ 予期せぬエラー: {str(e)}")
        automation.save_screenshot("unexpected_error")
    
    finally:
        # ブラウザを終了
        input("\nEnterキーを押すとブラウザを閉じます...")
        automation.quit()


if __name__ == "__main__":
    print("=" * 60)
    print("Airwork Selenium自動化 - 完全Selenium実装版")
    print("=" * 60)
    print("\n【実装フロー】")
    print("1. ✅ CSVダウンロード（C:\\Users\\□LMC□本社⑧\\Downloads\\pdf）")
    print("2. ✅ CSVのB2セルを取得して検索")
    print("3. ✅ 対応状況セルをクリック（Selenium - 環境差異に強い！）")
    print("4. ✅ レジュメボタンをクリック")
    print("5. ✅ PDF保存操作（PyAutoGUI）")
    print("6. ✅ 詳細ページを閉じる")
    print("7. ✅ 「面接調整開始」を選択")
    print("\n【新機能】")
    print("✅ ダウンロード先ディレクトリ指定")
    print("✅ CSV読み込み＆B2セル取得")
    print("✅ Selenium要素クリック（座標不要！）")
    print("✅ 右クリック→矢印キー操作")
    print("✅ 面接調整開始の自動選択")
    print("\n【設定が必要な項目】")
    print("1. ⚠️ main()内のURL、USERNAME、PASSWORD")
    print("=" * 60)
    
    response = input("\nテストを実行しますか？ (y/n): ")
    if response.lower() == 'y':
        main()
    else:
        print("テストをキャンセルしました。")
        print("main()関数内の認証情報を設定してから実行してください。")

