"""
config.yaml暗号化スクリプト

使い方:
    暗号化: python encrypt_config.py encrypt
    復号:   python encrypt_config.py decrypt
"""

import sys
import getpass
import base64
from pathlib import Path
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

def derive_key(password: str, salt: bytes) -> bytes:
    """
    パスワードから暗号化キーを生成（PBKDF2使用）
    
    Args:
        password: ユーザーが入力したパスワード
        salt: ソルト（ランダムバイト列）
    
    Returns:
        bytes: 暗号化キー（Base64エンコード済み）
    """
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    key = kdf.derive(password.encode())
    return base64.urlsafe_b64encode(key)

def encrypt_file(input_path: Path, output_path: Path, password: str) -> None:
    """
    ファイルを暗号化
    
    Args:
        input_path: 暗号化元ファイル（config.yaml）
        output_path: 暗号化先ファイル（config.enc）
        password: 暗号化パスワード
    """
    # ソルトを生成（ランダム16バイト）
    import os
    salt = os.urandom(16)
    
    # パスワードからキーを生成
    key = derive_key(password, salt)
    fernet = Fernet(key)
    
    # ファイルを読み込んで暗号化
    with open(input_path, 'rb') as f:
        data = f.read()
    
    encrypted_data = fernet.encrypt(data)
    
    # ソルト + 暗号化データを保存
    with open(output_path, 'wb') as f:
        f.write(salt)  # 最初の16バイトはソルト
        f.write(encrypted_data)
    
    print(f"✓ 暗号化完了: {input_path} → {output_path}")
    print(f"  ソルト長: {len(salt)} バイト")
    print(f"  暗号化データ長: {len(encrypted_data)} バイト")

def decrypt_file(input_path: Path, output_path: Path, password: str) -> None:
    """
    ファイルを復号
    
    Args:
        input_path: 暗号化ファイル（config.enc）
        output_path: 復号先ファイル（config.yaml）
        password: 復号パスワード
    """
    # 暗号化ファイルを読み込み
    with open(input_path, 'rb') as f:
        salt = f.read(16)  # 最初の16バイトはソルト
        encrypted_data = f.read()
    
    # パスワードからキーを生成
    key = derive_key(password, salt)
    fernet = Fernet(key)
    
    try:
        # 復号
        decrypted_data = fernet.decrypt(encrypted_data)
        
        # ファイルに保存
        with open(output_path, 'wb') as f:
            f.write(decrypted_data)
        
        print(f"✓ 復号完了: {input_path} → {output_path}")
        
    except Exception as e:
        print(f"✗ 復号失敗: パスワードが間違っているか、ファイルが破損しています")
        print(f"  エラー: {str(e)}")
        sys.exit(1)

def main():
    """メイン処理"""
    
    # 引数チェック
    if len(sys.argv) < 2:
        print("使い方:")
        print("  暗号化: python encrypt_config.py encrypt")
        print("  復号:   python encrypt_config.py decrypt")
        sys.exit(1)
    
    mode = sys.argv[1].lower()
    
    # スクリプトのディレクトリを取得
    script_dir = Path(__file__).parent
    
    if mode == "encrypt" or mode == "enc":
        # 暗号化モード
        print("=" * 50)
        print("config.yaml 暗号化ツール")
        print("=" * 50)
        
        input_file = script_dir / "config.yaml"
        output_file = script_dir / "config.enc"
        
        # ファイル存在チェック
        if not input_file.exists():
            print(f"✗ エラー: {input_file} が見つかりません")
            sys.exit(1)
        
        if output_file.exists():
            confirm = input(f"⚠ {output_file} は既に存在します。上書きしますか？ (y/N): ")
            if confirm.lower() != 'y':
                print("処理を中断しました")
                sys.exit(0)
        
        # パスワード入力
        password = getpass.getpass("暗号化パスワードを入力してください: ")
        password_confirm = getpass.getpass("確認のため再度入力してください: ")
        
        if password != password_confirm:
            print("✗ エラー: パスワードが一致しません")
            sys.exit(1)
        
        if len(password) < 8:
            print("✗ エラー: パスワードは8文字以上にしてください")
            sys.exit(1)
        
        # 暗号化実行
        encrypt_file(input_file, output_file, password)
        
        print("\n【重要】")
        print("1. 暗号化されたファイル (config.enc) をメンバーに共有してください")
        print("2. 元のファイル (config.yaml) は安全な場所に保管してください")
        print("3. パスワードは別途、安全な方法でメンバーに伝えてください")
        print("4. new_automation.py は暗号化ファイルを自動的に読み込みます")
        
    elif mode == "decrypt" or mode == "dec":
        # 復号モード
        print("=" * 50)
        print("config.enc 復号ツール")
        print("=" * 50)
        
        input_file = script_dir / "config.enc"
        output_file = script_dir / "config.yaml"
        
        # ファイル存在チェック
        if not input_file.exists():
            print(f"✗ エラー: {input_file} が見つかりません")
            sys.exit(1)
        
        if output_file.exists():
            confirm = input(f"⚠ {output_file} は既に存在します。上書きしますか？ (y/N): ")
            if confirm.lower() != 'y':
                print("処理を中断しました")
                sys.exit(0)
        
        # パスワード入力
        password = getpass.getpass("復号パスワードを入力してください: ")
        
        # 復号実行
        decrypt_file(input_file, output_file, password)
        
        print("\n✓ config.yaml が復元されました")
        
    else:
        print(f"✗ エラー: 不明なモード '{mode}'")
        print("使い方:")
        print("  暗号化: python encrypt_config.py encrypt")
        print("  復号:   python encrypt_config.py decrypt")
        sys.exit(1)

if __name__ == "__main__":
    main()

