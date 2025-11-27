import pyautogui
from PIL import Image, ImageEnhance
import time
import os
import webbrowser
import pyperclip
import pandas as pd
import logging
import yaml
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any
from dataclasses import dataclass
from retry import retry
import keyboard
import sys
import threading
import tkinter as tk
from tkinter import messagebox
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import queue
import getpass
import base64

# プログラム実行パスワード設定（setup.pyで生成された情報）
def check_password(password: str) -> bool:
    # 暗号化キー（実際の運用では別途安全に保管）
    KEY = b'yCh_OE7jEoX7S9aUMuk-CCNiJT_GIfb1ZHkLO8b5jbw='
    
    # 正しいPASSを暗号化した値
    ENCRYPTED_VALID_PASSWORD = b'gAAAAABnwUPuHzq3v84PkAwHhqeiqE2WnXY-2IxdOoZeOHfyeKVTToWfh89_8WQRTPGxJFJ40aoorfQLbb0-3pMRgX-cA2m41g=='

    # パスワードの検証
    f = Fernet(KEY)
    try:
        decrypted_password = f.decrypt(ENCRYPTED_VALID_PASSWORD).decode()
        return password == decrypted_password
    except:
        return False

def main():
    if len(sys.argv) < 2:
        print("認証エラー：パスワードが必要です")
        sys.exit(1)
    
    password = sys.argv[1]
    
    if len(sys.argv) > 2 and sys.argv[2] == "--verify":
        if not check_password(password):
            print("パスワードが違います")
            sys.exit(1)
        sys.exit(0)
    
    # 通常実行モード
    if not check_password(password):
        print("認証エラー：無効なパスワードです")
        sys.exit(1)

# メインの処理...
    print("認証成功！プログラムを実行します")


if __name__ == "__main__":
    main()
    
# 設定クラスの定義
@dataclass
class AutomationConfig:
    edge_path: str
    download_folder: str
    template_path: str
    url: str
    wait_time: Dict[str, int]
    base_image_path: str
    image_paths: Dict[str, Any]
    confidence_level: float
    login_info: Dict[str, str]
    secrets_file: str = "secrets.yaml"

class AutomationError(Exception):
    """自動化処理の例外クラス"""
    pass

class AutomationScript:

    """
    Airワーク自動化スクリプト
    - ログイン処理
    - CSVデータの取得と処理
    - テキスト検索と画像認識
    - メール送信
    を自動化します
    """

    def __init__(self, config_path: str):
        """
        初期化処理
        Args:
            config_path: 設定ファイルのパス
        """
        self.setup_logging()
        self.config = self.load_config(config_path)
        self.setup_environment()
        self.image_paths = self.load_image_paths()  # 画像パスを読み込む
        self.running = True  # 実行状態フラグ
        self.paused = False  # 一時停止フラグ
        self.pause_lock = threading.Lock()  # スレッドセーフな状態管理
        self.start_esc_monitor()
        self.start_pause_monitor()  # Alt+Space監視スレッド

    def start_esc_monitor(self):
        """ESCキー監視スレッドの開始"""
        def monitor_esc():
            while self.running:
                try:
                    if keyboard.is_pressed('esc'):
                        self.logger.warning("ESCキーが押されたため、プログラムを終了します")
                        self.cleanup()
                        os._exit(0)  # 強制終了
                    time.sleep(0.1)
                except Exception as e:
                    self.logger.error(f"ESCキー監視エラー: {str(e)}")

        self.esc_thread = threading.Thread(target=monitor_esc, daemon=True)
        self.esc_thread.start()
    
    def start_pause_monitor(self):
        """Alt+Spaceキー監視スレッドの開始"""
        def monitor_pause():
            while self.running:
                try:
                    # Alt+Spaceキーの組み合わせを監視
                    if keyboard.is_pressed('alt') and keyboard.is_pressed('space'):
                        with self.pause_lock:
                            self.paused = not self.paused  # トグル（切り替え）
                            if self.paused:
                                self.logger.warning("⏸ プログラムを一時停止しました（Alt+Spaceで再開）")
                            else:
                                self.logger.info("▶ プログラムを再開しました")
                        time.sleep(0.5)  # 連続押下防止
                    time.sleep(0.1)
                except Exception as e:
                    self.logger.error(f"一時停止監視エラー: {str(e)}")
        
        self.pause_thread = threading.Thread(target=monitor_pause, daemon=True)
        self.pause_thread.start()
    
    def wait_if_paused(self):
        """一時停止中であれば待機"""
        while self.paused and self.running:
            time.sleep(0.5)

    def show_dialog(self, message: str, is_error: bool = False):
        """ダイアログを表示"""
        try:
            root = tk.Tk()
            root.withdraw()  # メインウィンドウを非表示
            
            if is_error:
                messagebox.showerror("エラー", message)
            else:
                messagebox.showinfo("通知", message)
                
            root.destroy()
        except Exception as e:
            self.logger.error(f"ダイアログ表示エラー: {str(e)}")
    
    def setup_logging(self):
        """ログ設定"""
        from pathlib import Path
        script_dir = Path(__file__).parent  # スクリプトのディレクトリを取得
        log_dir = script_dir / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f"automation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [%(levelname)s] %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def decrypt_config_file(self, encrypted_path: Path, password: str) -> str:
        """
        暗号化された設定ファイルを復号
        
        Args:
            encrypted_path: 暗号化ファイルのパス (config.enc)
            password: 復号パスワード
            
        Returns:
            str: 復号されたYAML文字列
        """
        try:
            # 暗号化ファイルを読み込み
            with open(encrypted_path, 'rb') as f:
                salt = f.read(16)  # 最初の16バイトはソルト
                encrypted_data = f.read()
            
            # パスワードからキーを生成
            kdf = PBKDF2HMAC(
                algorithm=hashes.SHA256(),
                length=32,
                salt=salt,
                iterations=100000,
            )
            key = kdf.derive(password.encode())
            key = base64.urlsafe_b64encode(key)
            fernet = Fernet(key)
            
            # 復号
            decrypted_data = fernet.decrypt(encrypted_data)
            return decrypted_data.decode('utf-8')
            
        except Exception as e:
            return None

    def load_config(self, config_path: str) -> AutomationConfig:
        """設定ファイルの読み込み（暗号化対応）"""
        try:
            config_file = Path(config_path)
            encrypted_file = config_file.parent / "config.enc"
            
            # config.yamlが存在する場合はそれを使用
            if config_file.exists():
                self.logger.info("config.yaml を読み込みます")
                with open(config_path, 'r', encoding='utf-8') as f:
                    config_data = yaml.safe_load(f) or {}
                config_data = self._inject_secrets(config_file, config_data)
                return AutomationConfig(**config_data)
            
            # config.yamlが無く、config.encが存在する場合は復号して使用
            elif encrypted_file.exists():
                self.logger.info("暗号化された config.enc を検出しました")
                
                # パスワード入力（環境変数から取得、なければ入力）
                config_password = os.environ.get('CONFIG_PASSWORD')
                
                # パスワードが正しくない場合は再入力を促す
                while True:
                    if not config_password:
                        print("\n" + "=" * 50)
                        print("設定ファイル復号")
                        print("=" * 50)
                        config_password = getpass.getpass("config.enc の復号パスワードを入力してください: ")
                    
                    # 復号を試行
                    yaml_content = self.decrypt_config_file(encrypted_file, config_password)
                    
                    if yaml_content is not None:
                        # 復号成功
                        config_data = yaml.safe_load(yaml_content) or {}
                        config_data = self._inject_secrets(config_file, config_data)
                        self.logger.info("config.enc の復号に成功しました")
                        return AutomationConfig(**config_data)
                    else:
                        # 復号失敗 - パスワードが間違っている
                        print("パスワードが違います")
                        config_password = None  # 再入力を促すためにリセット
            
            else:
                raise AutomationError("config.yaml も config.enc も見つかりません")
                
        except AutomationError:
            raise
        except Exception as e:
            raise AutomationError(f"設定ファイルの読み込みに失敗: {str(e)}")
    
    def _inject_secrets(self, config_file: Path, config_data: Dict[str, Any]) -> Dict[str, Any]:
        """secrets.yamlの情報を設定に注入"""
        secrets_filename = config_data.get('secrets_file') or 'secrets.yaml'
        config_data['secrets_file'] = secrets_filename
        
        secrets_path = config_file.parent / secrets_filename
        if not secrets_path.exists():
            raise AutomationError(f"機密情報ファイルが見つかりません: {secrets_path}")
        
        with open(secrets_path, 'r', encoding='utf-8') as f:
            secrets_data = yaml.safe_load(f) or {}
        
        required_keys = ['url', 'login_info']
        for key in required_keys:
            if key not in secrets_data:
                raise AutomationError(f"secrets.yaml に {key} が定義されていません")
            config_data[key] = secrets_data[key]
        
        return config_data
    
    def wait_for_image(self, image_name: str, timeout: int = 30, use_grayscale: bool = True) -> bool:
        """
        画像が表示されるまで待機（クリックはしない）
        Args:
            image_name: 画像の名前
            timeout: タイムアウトまでの秒数（デフォルト30秒）
            use_grayscale: グレースケール検索を使うか（デフォルトTrue）
        Returns:
            bool: 成功したかどうか
        """
        self.logger.info(f"{image_name}の表示を待機中...")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                image_path = self.image_paths.get(image_name)
                if not image_path:
                    raise AutomationError(f"未定義の画像: {image_name}")
                
                location = pyautogui.locateOnScreen(
                    image_path,
                    confidence=self.config.confidence_level,
                    grayscale=use_grayscale
                )
                
                if location:
                    self.logger.info(f"{image_name}の表示を確認")
                    return True
                    
                time.sleep(1)  # 1秒待機して再試行
                
            except Exception as e:
                self.logger.debug(f"待機中... {str(e)}")
                time.sleep(1)
                
        self.logger.error(f"{image_name}が見つかりませんでした（タイムアウト）")
        return False

    def wait_and_click_image(self, image_name: str, timeout: int = 30, click_position: str = 'center', use_grayscale: bool = True) -> bool:
        """
        画像が表示されるまで待機してクリック
        Args:
            image_name: 画像の名前
            timeout: タイムアウトまでの秒数（デフォルト30秒）
            click_position: クリック位置（'center', 'left', 'right', 'top'）
            use_grayscale: グレースケール検索を使うか（デフォルトTrue）
        Returns:
            bool: 成功したかどうか
        """
        self.logger.info(f"{image_name}の表示を待機中...")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                image_path = self.image_paths.get(image_name)
                if not image_path:
                    raise AutomationError(f"未定義の画像: {image_name}")
                
                location = pyautogui.locateOnScreen(
                    image_path,
                    confidence=self.config.confidence_level,
                    grayscale=use_grayscale
                )
                
                if location:
                    # クリック位置の計算
                    if click_position == 'left':
                        click_x = location.left + (location.width * 0.25)  # 左から25%の位置
                        click_y = location.top + (location.height // 2)
                    elif click_position == 'right':
                        click_x = location.left + (location.width * 0.75)  # 右から25%の位置
                        click_y = location.top + (location.height // 2)
                    elif click_position == 'top':
                        click_x = location.left + (location.width // 2)  # 中央のX座標
                        click_y = location.top + (location.height * 0.25)  # 上部25%の位置
                    else:  # centerの場合
                        click_x = location.left + (location.width // 2)
                        click_y = location.top + (location.height // 2)
                    
                    pyautogui.click(click_x, click_y)
                    self.logger.info(f"{image_name}のクリックに成功")
                    return True
                    
                time.sleep(1)  # 1秒待機して再試行
                
            except Exception as e:
                self.logger.debug(f"待機中... {str(e)}")
                time.sleep(1)
                
        self.logger.error(f"{image_name}が見つかりませんでした（タイムアウト）")
        return False

    def setup_environment(self):
        """環境設定"""
        try:
            # フェイルセーフの調整（必要に応じて）
            pyautogui.FAILSAFE = True  # 安全機能は維持
            pyautogui.PAUSE = 0.5
            
            # マウス初期位置を画面中央に設定
            screen_width, screen_height = pyautogui.size()
            pyautogui.moveTo(screen_width // 2, screen_height // 2)
            
            # ブラウザの登録
            webbrowser.register('edge', None, 
                              webbrowser.BackgroundBrowser(self.config.edge_path))
                              
        except pyautogui.FailSafeException:
            # FailSafeExceptionはそのまま再投げしてメインハンドラーでキャッチさせる
            raise
        except Exception as e:
            raise AutomationError(f"環境設定に失敗: {str(e)}")

    def load_image_paths(self) -> Dict[str, str]:
        """画像パスを動的に生成"""
        try:
            base_path = self.config.base_image_path
            base_format = self.config.image_paths['base_format']
            start = self.config.image_paths['start_number']
            end = self.config.image_paths['end_number']
            
            paths = {
                f'image{num}': base_format.format(
                    base_path=base_path, 
                    number=num
                )
                for num in range(start, end + 1)
            }
            
            # 特定の画像の別名を追加
            special_images = {
                'login': paths['image1'],
                'login_half': paths['image2'],  # ログイン情報入力用
                'resume_open': paths['image9'],
                'download': paths['image7'],
                'resume_close': paths['image12'],
                'search_box': paths['image6'],
                'not_supported': paths['image8'],
                'a_date': paths['image8'],
                'iab': paths['image13'],
                'fail': paths['image12'],
                'save': paths['image10'],
                'filename_field': paths['image11'],  # ファイル名欄用
                'wait_image_1': f"{self.config.base_image_path}/wait_image_1.png",  # Excel待機用
                'wait_image_2': f"{self.config.base_image_path}/wait_image_2.png",  # 履歴書オープン直後待機用
                'wait_image_3': f"{self.config.base_image_path}/wait_image_3.png",  # PDF保存メニュー待機用
                'wait_image_4': f"{self.config.base_image_path}/wait_image_4.png",  # Excel検索欄待機用
                'wait_image_5': f"{self.config.base_image_path}/wait_image_5.png",  # 右クリックメニュー待機用
                'wait_image_6': f"{self.config.base_image_path}/wait_image_6.png",  # メールフォーマット待機用
                'wait_image_7': f"{self.config.base_image_path}/wait_image_7.png",  # 添付メニュー遷移待機用
                'wait_image_8': f"{self.config.base_image_path}/wait_image_8.png",  # このPC参照表示待機用
            }
            paths.update(special_images)
            
            # デバッグ用：待機画像のパス確認
            self.logger.info("=== 待機画像のパス確認 ===")
            for wait_img in ['wait_image_1', 'wait_image_2', 'wait_image_3', 'wait_image_4', 'wait_image_5', 'wait_image_6', 'wait_image_7', 'wait_image_8']:
                img_path = special_images.get(wait_img)
                self.logger.info(f"{wait_img}のパス: {img_path}")
                
                # ファイルの存在確認
                if img_path:
                    full_path = Path(img_path)
                    if full_path.exists():
                        self.logger.info(f"  → ✓ {wait_img}が見つかりました")
                    else:
                        self.logger.warning(f"  → ✗ {wait_img}が見つかりません！")
                        # 絶対パスでも確認
                        abs_path = Path(__file__).parent / img_path
                        if abs_path.exists():
                            self.logger.info(f"  → ✓ 絶対パスでは見つかりました: {abs_path}")
                        else:
                            self.logger.warning(f"  → ✗ 絶対パスでも見つかりません: {abs_path}")
            self.logger.info("========================")
            
            return paths
        except Exception as e:
            raise AutomationError(f"画像パス生成に失敗: {str(e)}")

    @retry(tries=3, delay=2, backoff=2)
    def click_image(self, image_name: str, use_grayscale: bool = False) -> None:
        """画像認識によるクリック（リトライ機能付き）"""
        image_path = self.image_paths.get(image_name)  # 修正
        if not image_path:
            raise AutomationError(f"未定義の画像: {image_name}")

        self.logger.info(f"画像クリック試行: {image_name}")
        image = pyautogui.locateOnScreen(
            image_path, 
            confidence=self.config.confidence_level,
            grayscale=use_grayscale
        )
        
        if image:
            pyautogui.click(image)
            time.sleep(self.config.wait_time.get('click', 2))
            self.logger.info(f"画像クリック成功: {image_name}")
        else:
            raise AutomationError(f"画像が見つかりません: {image_name}")
    

    def get_latest_csv(self) -> str:
        """最新のCSVファイルを取得"""
        try:
            download_folder = Path(self.config.download_folder).expanduser().resolve()
            csv_files = list(download_folder.glob('*.csv'))
            if not csv_files:
                raise AutomationError("CSVファイルが見つかりません")
            
            latest_csv = max(csv_files, key=lambda f: f.stat().st_ctime)
            self.logger.info(f"最新CSVファイル: {latest_csv}")
            return str(latest_csv)
        except Exception as e:
            raise AutomationError(f"CSVファイル取得エラー: {str(e)}")

    def process_data(self, csv_path: str) -> pd.DataFrame:
        """CSVデータの処理"""
        try:
            df = pd.read_csv(csv_path)
            self.logger.info("CSVファイルの列名:")
            self.logger.info(df.columns.tolist())

            # インデックス番号で列を取得（0がA, 1がB, 4がE, 8がI, 29がAD, 36がAK）
            df_bx = pd.DataFrame()
            df_bx['B'] = df.iloc[:, 1]    # B列（氏名）
            df_bx['E'] = df.iloc[:, 4]    # E列（年齢）
            df_bx['I'] = df.iloc[:, 8]    # I列（アドレス）
            df_bx['AD'] = df.iloc[:, 29]   # AD列（支店名）
            df_bx['AK'] = df.iloc[:, 36]   # AK列（新規追加データ）

            # AD列の文字列処理（既存のコード）
            def clean_branch_name(text: str) -> str:
                if not isinstance(text, str):
                    return ""
                parts = text.split('　')
                if len(parts) > 1:
                    self.logger.info(f"支店名抽出: '{text}' → '{parts[-1]}'")
                    return parts[-1]
                return text

            df_bx['AD'] = df_bx['AD'].apply(clean_branch_name)
            
            self.logger.info("データ処理完了")
            return df_bx
        except Exception as e:
            raise AutomationError(f"データ処理エラー: {str(e)}")

    @retry(tries=2, delay=1)
    def find_and_click_text(self, text: str) -> None:
        try:
            self.logger.info(f"検索対象のフルネーム: '{text}'")

            # 検索欄をクリック
            self.logger.info("検索欄をクリック")
            if not self.wait_and_click_image('search_box', timeout=30, click_position='left'):
                raise AutomationError("検索欄が見つかりません")
            time.sleep(1)
            
            # フルネームを検索欄に入力
            pyautogui.hotkey('ctrl', 'a')
            pyperclip.copy(text)
            pyautogui.hotkey('ctrl', 'v')
            
            # 検索ボタンをクリック
            self.logger.info("検索ボタンをクリック")
            if not self.wait_and_click_image('search_box', timeout=30, click_position='right'):
                raise AutomationError("検索ボタンが見つかりません")
            
            time.sleep(2)

            # image10の左側をクリック
            if not self.wait_and_click_image('a_date', timeout=30, click_position='left'):
                raise AutomationError("候補が見つかりません")
            time.sleep(2)

        except Exception as e:
            self.logger.error(f"テキスト検索エラー: {str(e)}")
            raise

    def send_email(self, branch_name: str, recipient: str, ak_value: str, applicant_address: str = "") -> None:
        """メール送信とその後の処理"""
        try:
            # Adobe PDF画面閉じる
            pyautogui.hotkey('ctrl', 'Q')
            time.sleep(3)

            # Excelウィンドウをアクティブ化
            window = pyautogui.getWindowsWithTitle(
                "outlookmail_送付フォーマット.xlsx")[0]
            window.activate()
            
            # Excel内の特定要素が表示されるまで待機
            self.logger.info("Excel内の要素表示を待機中...")
            if not self.wait_for_image('wait_image_1', timeout=30):
                raise AutomationError("Excelファイルの内容表示待機がタイムアウトしました")
            
            # メール作成までの処理
            pyautogui.hotkey('ctrl', 'f')
            self.logger.info("Excel検索ダイアログの表示を待機中...")
            if not self.wait_for_image('wait_image_4', timeout=30):
                raise AutomationError("Excel検索ダイアログの表示待機がタイムアウトしました")
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1)
            pyautogui.hotkey('delete')
            time.sleep(1)
            pyperclip.copy(recipient)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(2)
            pyautogui.press('esc')
            time.sleep(2)
            
            # カーソル移動とメニュー表示
            for _ in range(4):
                pyautogui.press('tab')
            
            pyautogui.hotkey('shift', 'f10')  # 右クリックメニュー
            self.logger.info("右クリックメニューの表示を待機中...")
            if not self.wait_for_image('wait_image_5', timeout=30):
                raise AutomationError("右クリックメニューの表示待機がタイムアウトしました")
            time.sleep(1)
            pyautogui.press('o')
            time.sleep(1)
            pyautogui.press('o')
            pyautogui.press('enter')
            time.sleep(3)
            
            # メールフォーマット画面が表示されるまで待機
            self.logger.info("メールフォーマット画面の表示を待機中...")
            if not self.wait_for_image('wait_image_6', timeout=30):
                raise AutomationError("メールフォーマット画面の表示待機がタイムアウトしました")
            
            # 支店名を貼り付け
            pyperclip.copy(branch_name)
            pyautogui.hotkey('ctrl', 'v')
            self.logger.info("支店名をペーストしました")

            # スペースキーを1回入力
            pyautogui.press('space')

            # 職種データをペースト
            pyperclip.copy(ak_value)
            pyautogui.hotkey('ctrl', 'v')
            self.logger.info("職種データをペーストしました")

            # 改行してからアドレスをペースト
            pyautogui.press('enter')
            time.sleep(2)
            
            # 応募者のアドレスをペースト（スクショの場合）
            if applicant_address:
                pyperclip.copy(applicant_address)
                pyautogui.hotkey('ctrl', 'v')
                self.logger.info(f"応募者アドレスをペーストしました: {applicant_address}")
                time.sleep(2)

            # 最新の添付ファイル名を取得（拡張子付き）
            attachment_filename = self.get_latest_attachment_filename()
            
            # 新しいファイル添付方法（エクスプローラー経由）
            self.logger.info(f"ファイルを添付します: {attachment_filename}")
            
            # 1. Alt+H → A → F → B でこのPC参照を開く
            pyautogui.hotkey('alt', 'h')
            time.sleep(1)
            pyautogui.press('a')
            time.sleep(1)
            pyautogui.press('f')
            self.logger.info("添付メニューの表示を待機中...")
            if not self.wait_for_image('wait_image_7', timeout=30):
                raise AutomationError("添付メニューの表示待機がタイムアウトしました")
            time.sleep(1)
            pyautogui.press('b')
            time.sleep(1)
            self.logger.info("参照ダイアログの表示を待機中...")
            if not self.wait_for_image('wait_image_8', timeout=30):
                raise AutomationError("参照ダイアログの表示待機がタイムアウトしました")
            
            # 2. Alt+Dでフォルダパスの欄をアクティブ化
            self.logger.info("フォルダパスの欄をアクティブ化")
            pyautogui.hotkey('alt', 'd')
            time.sleep(2)
            
            # 3. download_folderのパスを貼り付け（絶対パスを取得）
            download_folder_path = str(Path(self.config.download_folder).expanduser().resolve())
            self.logger.info(f"フォルダパスを入力: {download_folder_path}")
            pyperclip.copy(download_folder_path)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(5)  # フォルダ内容の読み込み完了を待つ
            
            # 4. image11（ファイル名欄）の右側をクリック
            self.logger.info("ファイル名欄にカーソルを表示")
            if not self.wait_and_click_image('filename_field', timeout=30, click_position='right'):
                raise AutomationError("ファイル名欄が見つかりません")
            time.sleep(3)  # フィールドがアクティブになるまで待機
            
            # フィールドがアクティブになったことを確認するため、もう一度クリック
            pyautogui.click()  # 同じ位置を再クリック
            time.sleep(1)
            
            # 既存の内容をクリア
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            
            # 5. ファイル名を貼り付けてEnter
            self.logger.info(f"ファイル名を入力: {attachment_filename}")
            pyperclip.copy(attachment_filename)
            time.sleep(0.5)  # クリップボードへのコピー完了待機
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(3)  # ペースト後の待機時間を増加
            pyautogui.press('enter')
            time.sleep(2)
            
            # メールフォーマット画面に戻るまで待機
            self.logger.info("メールフォーマット画面の表示を待機中...")
            if not self.wait_for_image('wait_image_6', timeout=30):
                raise AutomationError("メールフォーマット画面の表示待機がタイムアウトしました")
            
            time.sleep(self.config.wait_time.get('email', 2))
            
            # メール送信
            pyautogui.hotkey('ctrl', 'enter')
            self.logger.info(f"メール送信完了: {recipient}")
            
            
            # Excel内の特定要素が表示されるまで待機（メール送信後のExcel画面復帰確認）
            self.logger.info("メール送信後のExcel画面を待機中...")
            if not self.wait_for_image('wait_image_1', timeout=30):
                self.logger.warning("Excelファイルの内容表示待機がタイムアウトしました（メール送信後）")

            # ブラウザ画面に戻る処理
            pyautogui.hotkey('alt', 'tab')
            time.sleep(5)  # ブラウザへの切り替え完了を待つ
            
        except Exception as e:
            raise AutomationError(f"メール送信エラー: {str(e)}")

    def wait_for_excel(self, timeout: int = 30) -> bool:
        """Excelファイルが開くのを待機"""
        self.logger.info("Excelファイルの表示を待機中...")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                windows = pyautogui.getWindowsWithTitle("outlookmail_送付フォーマット.xlsx")
                if windows:
                    self.logger.info("Excelファイルの表示を確認")
                    return True
                time.sleep(1)
            except Exception as e:
                self.logger.debug(f"Excel待機中... {str(e)}")
                time.sleep(1)
        
        return False

    def get_latest_pdf(self) -> bool:
        """PDFファイルが正常に保存されたか確認"""
        try:
            download_folder = Path(self.config.download_folder).expanduser().resolve()
            pdf_files = list(download_folder.glob('*.pdf'))
            if not pdf_files:
                self.logger.error("PDFファイルが見つかりません")
                return False
            
            latest_pdf = max(pdf_files, key=lambda f: f.stat().st_mtime)
            file_time = latest_pdf.stat().st_mtime
            current_time = time.time()
            
            # 最新のPDFファイルが5分以内に作成されていれば成功とみなす
            if current_time - file_time < 300:
                self.logger.info(f"PDFファイルが正常に保存されました: {latest_pdf}")
                return True
            
            self.logger.warning("最近作成されたPDFファイルが見つかりません")
            return False
                
        except Exception as e:
            self.logger.error(f"PDF保存確認エラー: {str(e)}")
            return False

    def save_pdf_with_retry(self, retry_count=3, delay=5):
        """PDFとして保存する処理（リトライあり）"""
        self.logger.info("PDFとして保存を試行")
        
        for attempt in range(retry_count):
            try:
                # レジュメページが開くまで待機
                time.sleep(5)
                
                # 右クリックする
                self.logger.info("右クリックを実行")
                pyautogui.click(button='right')
                time.sleep(4)
                
                # 保存メニューの画像をクリック
                self.logger.info("保存メニューをクリック")
                if not self.wait_and_click_image('save', timeout=30):
                    raise AutomationError("保存メニューが見つかりません")
                self.logger.info("PDF保存ダイアログの表示を待機中...")
                if not self.wait_for_image('wait_image_3', timeout=30):
                    raise AutomationError("PDF保存ダイアログが表示されませんでした")
                time.sleep(2)
                
                # Enterを押して保存実行
                self.logger.info("Enterキーを押下")
                pyautogui.press('enter')
                time.sleep(5)  # 保存完了まで待機

                # Adobe PDF画面閉じる
                pyautogui.hotkey('ctrl', 'Q')
                time.sleep(5)

                # PDFが作成されたか確認
                if self.get_latest_pdf():
                    self.logger.info("PDF保存が成功")
                    return True
                
                self.logger.warning(f"PDF保存失敗、リトライ {attempt+1}/{retry_count}")
                time.sleep(delay)
                
            except Exception as e:
                self.logger.error(f"PDF保存エラー: {str(e)}")
                time.sleep(delay)
        
        self.logger.error("PDF保存が複数回失敗")
        return False

    def save_screenshot(self, filename: str = None) -> str:
        """
        現在の画面をスクリーンショットとして保存
        Args:
            filename: 保存するファイル名（指定しない場合はタイムスタンプ付き）
        Returns:
            str: 保存したスクリーンショットのパス
        """
        try:
            # スクリーンショット保存先（PDFと同じフォルダに統一）
            screenshot_dir = Path(self.config.download_folder).expanduser().resolve()
            screenshot_dir.mkdir(parents=True, exist_ok=True)
            
            # スクリーンショットのファイル名
            if filename:
                screenshot_path = screenshot_dir / filename
            else:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_path = screenshot_dir / f"screenshot_{timestamp}.png"
            
            # スクリーンショット取得と保存
            screenshot = pyautogui.screenshot()
            screenshot.save(str(screenshot_path))
            
            self.logger.info(f"スクリーンショット保存: {screenshot_path}")
            return str(screenshot_path)
            
        except Exception as e:
            self.logger.error(f"スクリーンショット保存エラー: {str(e)}")
            raise AutomationError(f"スクリーンショット保存に失敗: {str(e)}")

    def get_latest_attachment(self) -> str:
        """
        最新の添付ファイル（PDF or スクリーンショット）を取得
        Returns:
            str: 添付するファイルのパス
        """
        try:
            # PDFフォルダのパス（スクショもPDFフォルダに保存される）
            pdf_folder = Path(self.config.download_folder).expanduser().resolve()

            latest_file = None
            latest_time = 0

            # PDFファイルとPNGファイルをチェック
            if pdf_folder.exists():
                # PDFファイルをチェック
                pdf_files = list(pdf_folder.glob('*.pdf'))
                if pdf_files:
                    pdf_latest = max(pdf_files, key=lambda f: f.stat().st_mtime)
                    latest_time = pdf_latest.stat().st_mtime
                    latest_file = pdf_latest

                # PNGファイル（スクリーンショット）をチェック
                png_files = list(pdf_folder.glob('*.png'))
                if png_files:
                    png_latest = max(png_files, key=lambda f: f.stat().st_mtime)
                    if png_latest.stat().st_mtime > latest_time:
                        latest_file = png_latest

            if latest_file:
                self.logger.info(f"添付ファイル: {latest_file}")
                return str(latest_file)
            else:
                raise AutomationError("添付可能なファイルが見つかりません")

        except Exception as e:
            self.logger.error(f"添付ファイル取得エラー: {str(e)}")
            raise AutomationError(f"添付ファイル取得に失敗: {str(e)}")

    def get_latest_attachment_filename(self) -> str:
        """
        最新の添付ファイルのファイル名のみを取得（拡張子付き）
        Returns:
            str: ファイル名（例: "resume.pdf" or "山田太郎_東京オフィス.png"）
        """
        try:
            attachment_path = self.get_latest_attachment()
            filename = Path(attachment_path).name
            self.logger.info(f"添付ファイル名: {filename}")
            return filename
        except Exception as e:
            self.logger.error(f"ファイル名取得エラー: {str(e)}")
            raise AutomationError(f"ファイル名取得に失敗: {str(e)}")

    def confirm_csv_data(self, df_bx):
        """B/E/I/AD/AK列の内容をダイアログで表示し、10秒後に自動でOK"""
        preview = df_bx[['B', 'E', 'I', 'AD', 'AK']].head(5).to_string(index=False)
        root = tk.Tk()
        root.withdraw()
        
        # 10秒後に自動的にEnterキーを押すスレッド
        def auto_confirm():
            time.sleep(10)
            try:
                pyautogui.press('enter')
                self.logger.info("10秒経過：自動的にOKを押しました")
            except Exception as e:
                self.logger.error(f"自動確認エラー: {str(e)}")
        
        # 自動確認スレッドを開始
        auto_thread = threading.Thread(target=auto_confirm, daemon=True)
        auto_thread.start()
        
        result = messagebox.askokcancel(
            "CSVデータ確認（10秒後に自動続行）",
            f"下記のデータで処理を開始します。\n\n{preview}\n\n※10秒後に自動的に続行します\nOK→続行 / キャンセル→中断"
        )
        root.destroy()
        return result

    def send_notification_email(self, recipient: str, subject: str, body: str) -> None:
        """
        処理完了通知メールを送信（Outlook COM経由）
        Args:
            recipient: 受信者のメールアドレス
            subject: 件名
            body: 本文
        """
        try:
            import win32com.client
            
            self.logger.info(f"通知メールを送信します: {recipient}")
            
            # Outlookアプリケーションを起動
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            # 新規メールアイテムを作成（0はolMailItem定数）
            mail_item = outlook.CreateItem(0)
            
            # メール情報を設定
            mail_item.To = recipient.strip()
            mail_item.Subject = subject
            mail_item.Body = body
            
            # メールを送信
            mail_item.Send()
            
            self.logger.info(f"通知メール送信完了: {recipient}")
            
        except Exception as e:
            raise AutomationError(f"通知メール送信エラー: {str(e)}")

    def send_completion_notifications(self) -> None:
        """
        全処理完了後に通知先リストから通知メールを送信（Outlook COM経由）
        """
        try:
            import win32com.client
            
            self.logger.info("処理完了通知を送信します")
            
            # Excelファイルのパス
            excel_path = Path(__file__).parent / "outlookmail_送付フォーマット.xlsx"
            
            # Excelファイルを読み込み（シート名：処理完了の通知先）
            df = pd.read_excel(excel_path, sheet_name='処理完了の通知先', header=0)
            
            self.logger.info("通知先リストを読み込みました")
            self.logger.info(f"列名: {df.columns.tolist()}")
            
            # 件名を取得（E1セル = ヘッダー行のE列 = 列名から取得）
            subject = ""
            if len(df.columns) > 4:
                subject = str(df.columns[4]) if pd.notna(df.columns[4]) else ""
            self.logger.info(f"件名: {subject}")
            
            # B列の最終行を求める（NaNでない最後の行）
            # A列=氏名、B列=アドレス、E列=本文
            df_filtered = df[df.iloc[:, 1].notna()]  # B列（インデックス1）がNaNでない行
            
            if df_filtered.empty:
                self.logger.warning("通知先リストが空です")
                return
            
            self.logger.info(f"通知先件数: {len(df_filtered)}")
            
            # Outlookアプリケーションを起動
            self.logger.info("Outlookアプリケーションを起動しています...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            # 各通知先にメールを送信
            for i, row in df_filtered.iterrows():
                try:
                    name = row.iloc[0]  # A列：氏名
                    recipient = row.iloc[1]  # B列：アドレス
                    body = row.iloc[4]  # E列：本文
                    
                    # 空の宛先はスキップ
                    if not recipient or str(recipient).strip() == "" or pd.isna(recipient):
                        self.logger.info(f"行 {i + 1}: 宛先が空のためスキップしました")
                        continue
                    
                    self.logger.info(f"通知メール送信: {name} ({recipient})")
                    
                    # 新規メールアイテムを作成（0はolMailItem定数）
                    mail_item = outlook.CreateItem(0)
                    
                    # メール情報を設定
                    mail_item.To = str(recipient).strip()
                    mail_item.Subject = subject  # 全メール共通の件名（E1セルから取得）
                    mail_item.Body = str(body) if pd.notna(body) else ""
                    
                    # メールを送信
                    mail_item.Send()
                    
                    self.logger.info(f"メール送信完了: {recipient}")
                    
                except Exception as e:
                    self.logger.error(f"通知メール送信エラー（{i+1}行目）: {str(e)}")
                    continue
            
            self.logger.info("全通知メールの送信が完了しました")
            
        except Exception as e:
            self.logger.error(f"処理完了通知エラー: {str(e)}")
            raise AutomationError(f"処理完了通知に失敗: {str(e)}")

    def run(self):
        """メイン処理"""
        try:
            self.logger.info("処理開始")
            
            # ブラウザ起動
            self.logger.info("ブラウザを起動中...")
            try:
                # マウスを画面中央に移動
                screen_width, screen_height = pyautogui.size()
                pyautogui.moveTo(screen_width // 2, screen_height // 2)
                
                edge_exe = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                os.startfile(edge_exe)
                time.sleep(7)  # ブラウザ起動待機
                
                # URLを入力（Ctrl+Lでアドレスバーにフォーカス）
                pyautogui.hotkey('ctrl', 'l')
                time.sleep(3)
                pyperclip.copy(self.config.url)
                pyautogui.hotkey('ctrl', 'v')
                pyautogui.press('enter')
                
            except Exception as e:
                self.logger.error(f"ブラウザ起動エラー: {str(e)}")
                raise AutomationError("ブラウザの起動に失敗しました")
                
            # ページ読み込みのための待機
            self.logger.info("ページの読み込みを待機中...")
            time.sleep(6)
            
            # ログイン処理
            self.logger.info("ログイン画面の要素を探索中...")
            self.click_image('login', use_grayscale=False)
            time.sleep(4)

            # ログイン情報入力
            self.logger.info("ログイン情報入力フィールドを探します...")
            if self.wait_and_click_image('login_half', timeout=20, click_position='top'):
                pyperclip.copy(self.config.login_info['username'])
                pyautogui.hotkey('ctrl', 'v')

                pyautogui.press('tab')  # パスワードフィールドへ移動
                pyperclip.copy(self.config.login_info['password'])
                pyautogui.hotkey('ctrl', 'v')

                # タブを2回押してエンターでログイン
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('enter')

                # ログイン処理の待機
                self.logger.info("ログイン処理の待機中...")
                time.sleep(5)
            else:
                self.logger.info("ログイン入力画面が表示されなかったため、ホーム画面へ遷移済みと判断し処理を継続します")
                time.sleep(2)
            
            # 画像2から5までを順番にクリック
            images_to_click = [
                'image3',
                'image4',
                'image5',
                ('search_box', 'right'),
                'download',
            ]
            for item in images_to_click:
                if isinstance(item, tuple):
                    image_name, position = item
                else:
                    image_name, position = item, 'center'
                if not self.wait_and_click_image(image_name, click_position=position):
                    raise AutomationError(f"{image_name}のクリックに失敗しました")

            time.sleep(7)  # ダウンロード待機

            # CSVファイル処理
            self.logger.info("CSVファイルを処理中...")
            csv_path = self.get_latest_csv()
            df_bx = self.process_data(csv_path)

            # 追加：CSV内容の確認ダイアログ
            if not self.confirm_csv_data(df_bx):
                self.logger.info("ユーザーが処理を中断しました")
                self.show_dialog("処理を中断しました", is_error=True)
                return

            # CSVファイル処理直後にExcelファイルを開く
            self.logger.info("Excelファイルを開きます...")
            try:
                excel_path = Path(__file__).parent / "outlookmail_送付フォーマット.xlsx"
                os.startfile(str(excel_path))
                
                # Excelファイルの表示を待機
                if not self.wait_for_excel(timeout=30):
                    raise AutomationError("Excelファイルの表示待機がタイムアウトしました")
                time.sleep(2)  # 安定化のための追加待機
                
                # Excelウィンドウをアクティブ化（画像認識のため前面に表示）
                self.logger.info("Excelウィンドウをアクティブ化します")
                excel_window = pyautogui.getWindowsWithTitle("outlookmail_送付フォーマット.xlsx")[0]
                excel_window.activate()
                  # アクティブ化待機
                
                # Excel内の特定要素が表示されるまで待機
                self.logger.info("Excel内の要素表示を待機中...")
                if not self.wait_for_image('wait_image_1', timeout=30):
                    raise AutomationError("Excelファイルの内容表示待機がタイムアウトしました")
                
                time.sleep(2)
                self.logger.info("ブラウザ画面に切り替えます")
                pyautogui.hotkey('alt', 'tab')  # ブラウザ画面に戻る

            except Exception as e:
                raise AutomationError(f"Excelファイルを開けません: {str(e)}")

            # メインループ処理の修正
            for i, row in df_bx.iterrows():
                self.wait_if_paused()  # 一時停止チェック
                try:
                    self.logger.info(f"行の処理開始: {i + 1}")
                    age = int(row['E'])  # 年齢を取得
                    
                    # 55歳以上の場合は早期にスキップ（54歳以下が対象）
                    if age >= 55:
                        self.logger.info(f"55歳以上のためスキップ: {age}歳")
                        continue
                    
                    time.sleep(3)  # 画面切り替え待ち

                    # 1. 検索処理を実行
                    self.find_and_click_text(row['B'])
                    
                    # メインループ処理内の該当部分を修正
                    resume_opened = False
                    try:
                        # 履歴書を開くボタンのクリックを試行
                        resume_button_clicked = self.wait_and_click_image('resume_open', timeout=30)
                        if not resume_button_clicked:
                            self.logger.warning("履歴書を開くボタンが見つかりません。スクリーンショットを取得します")
                            
                            # 1. スクリーンショットを取得する
                            # 2. スクリーンショットのファイル名をB列_AD列の形式にする
                            screenshot_filename = f"{row['B']}_{row['AD']}.png"
                            
                            # スクリーンショット取得と保存（PDFフォルダに保存）
                            self.save_screenshot(filename=screenshot_filename)
                            
                            # 3. スクリーンショットを添付してメール送信
                            self.logger.info("スクリーンショットを添付してメールを送信します")
                            self.send_email(row['AD'], row['AD'], row['AK'], row['I'])
                            
                        else:
                            resume_opened = True
                            self.logger.info("履歴書オープン後の画面を待機中...")
                            # wait_image_2の待機をコメントアウト（タブ名が変わることがあるため）
                            # if not self.wait_for_image('wait_image_2', timeout=30):
                            #     raise AutomationError("履歴書画面の表示待機がタイムアウトしました")
                            time.sleep(7)  # 画面安定待機（画像待機の代わりに固定時間待機）
                            
                            # PDF保存処理
                            if not self.save_pdf_with_retry():
                                self.logger.warning("PDF保存に問題があったため、スクリーンショットを取得します")
                                self.save_screenshot()
                                
                            # メール送信処理（PDFの場合はアドレス不要）
                            self.send_email(row['AD'], row['AD'], row['AK'])
                        
                    except Exception as e:
                        self.logger.error(f"行{i + 1}の処理でエラー: {str(e)}")
                        # エラー発生時はスクリーンショットを取得
                        try:
                            self.save_screenshot()
                        except Exception as screenshot_error:
                            self.logger.error(f"スクリーンショット取得エラー: {str(screenshot_error)}")
                        continue

                    # 5. レジュメページを閉じる（Ctrl+W）
                    if resume_opened:
                        self.logger.info("レジュメページを閉じます")
                        pyautogui.hotkey('ctrl', 'w')
                        time.sleep(2)
                    else:
                        self.logger.info("レジュメページが開かれていないためCtrl+Wをスキップします")
                        self.logger.info("詳細ページは既に閉じられているためresume_closeの待機をします")
                    
                    # 6. 詳細ページを閉じる(resume_close)
                    if not self.wait_and_click_image('resume_close', timeout=30):
                        self.logger.warning("詳細ページの閉じるボタンが見つかりませんでした")
                    else:
                        self.logger.info("詳細ページを閉じました")
                    time.sleep(2)
                    
                    # 問題点2の修正: image10とimage11を順番にクリック
                    self.logger.info("not_supported画像を確認")
                    if not self.wait_and_click_image('not_supported', timeout=30):
                        self.logger.warning("not_supported画像が見つかりません。スクリーンショットを取得します")
                        self.save_screenshot()
                    
                    time.sleep(2)
                    
                    self.logger.info("image11(iab)をクリック")
                    if not self.wait_and_click_image('iab', timeout=30):
                        self.logger.warning("image11が見つかりませんでした")
                    time.sleep(2)
                    
                    self.logger.info(f"行の処理完了: {i + 1}")
                
                except Exception as e:
                    self.logger.error(f"行{i + 1}の処理でエラー: {str(e)}")
                    # エラーが発生しても次の行の処理へ進む
                    continue

            self.logger.info("全処理完了")
            
            # 検索欄をクリアして未対応フラグを表示
            try:
                self.logger.info("検索欄をクリアします")
                # 1. image6の左側をクリック
                if self.wait_and_click_image('search_box', timeout=10, click_position='left'):
                    time.sleep(1)
                    # 2. Ctrl+Aで全選択
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(1)
                    # 3. deleteで削除
                    pyautogui.press('delete')
                    time.sleep(1)
                    # 4. image6の右側をクリック（未対応フラグを表示）
                    if self.wait_and_click_image('search_box', timeout=10, click_position='right'):
                        self.logger.info("検索欄をクリアし、未対応フラグを表示しました")
                        time.sleep(2)  # 画面表示の安定化待機
                    else:
                        self.logger.warning("検索ボックスの右側クリックに失敗しました")
                else:
                    self.logger.warning("検索ボックスが見つかりませんでした")
            except Exception as e:
                self.logger.warning(f"検索欄のクリア処理でエラー: {str(e)}")
                # エラーが発生しても処理は続行
            
            # 全処理完了後に通知メールを送信
            try:
                self.logger.info("処理完了通知を送信します")
                self.send_completion_notifications()
            except Exception as e:
                self.logger.error(f"処理完了通知の送信に失敗しました: {str(e)}")
                # 通知メール送信失敗は致命的ではないため、処理は続行
            
        except Exception as e:
            self.logger.error(f"致命的なエラー: {str(e)}")
            raise
        
        finally:
            # 全ての処理が完了してからクリーンアップを実行
            self.cleanup()

    def cleanup(self):
        """クリーンアップ処理"""
        try:
            self.running = False  # ESCキー監視スレッドを停止
            # まず開いているウィンドウを閉じる
            for window_title in ["latest_csv_file", "outlookmail_送付フォーマット.xlsx"]:
                try:
                    window = pyautogui.getWindowsWithTitle(window_title)[0]
                    window.close()
                except Exception as e:
                    self.logger.warning(f"ウィンドウが閉じられません: {window_title}")
            
            # 一時ファイルの削除前に十分な待機時間を設定
            time.sleep(3)
            
            # デバッグ画像の削除
            debug_folder = Path(__file__).parent / "logs" / "debug_full_image"
            if debug_folder.exists():
                for file in debug_folder.glob("*.png"):
                    try:
                        file.unlink()
                    except Exception:
                        pass
        
            self.logger.info("クリーンアップ完了")
            
        except Exception as e:
            self.logger.error(f"クリーンアップエラー: {str(e)}")


if __name__ == "__main__":
    automation = None  # 例外時の参照対策
    try:
        from pathlib import Path
        script_dir = Path(__file__).parent
        # スクリプトと同じ  ディレクトリに config.yaml がある場合
        config_path = str(script_dir / "config.yaml")
        automation = AutomationScript(config_path)
        automation.logger.info("プログラムを開始します")
        
        # メイン処理の実行
        automation.run()

        # 正常終了時のダイアログ
        automation.show_dialog("処理が完了しました")
                
    except pyautogui.FailSafeException:
        print("\nマウスが画面の四隅にありますので、実行できません")
        print("マウスを中央付近に移動してから、再度実行してください。")
        if automation and automation.logger:
            automation.logger.error("FailSafeException: マウスが画面の四隅にあります")
        sys.exit(1)
    except Exception as e:
        if automation and automation.logger:
            automation.logger.error(f"エラーが発生しました: {str(e)}")
            # エラー時のダイアログ
            automation.show_dialog("エラーが起きたので中断されました", is_error=True)
        else:
            print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        input("Enterキーを押して終了")