import subprocess
import sys

def get_installed_packages():
    """インストール済みパッケージのリストを取得"""
    result = subprocess.run(
        [sys.executable, "-m", "pip", "freeze"],
        capture_output=True,
        text=True
    )
    installed = {}
    for line in result.stdout.strip().split('\n'):
        if line and '==' in line:
            package_name = line.split('==')[0].lower()
            installed[package_name] = line
    return installed

def get_required_packages(requirements_file='requirements.txt'):
    """requirements.txtから必要なパッケージのリストを取得"""
    required = set()
    try:
        with open(requirements_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                # コメントと空行をスキップ
                if line and not line.startswith('#'):
                    # パッケージ名を抽出（バージョン指定子を除く）
                    package_name = line.split('==')[0].split('>=')[0].split('<=')[0].split('~=')[0].split('!=')[0].split('>')[0].split('<')[0].strip().lower()
                    required.add(package_name)
    except FileNotFoundError:
        print(f"エラー: {requirements_file} が見つかりません。")
        sys.exit(1)
    return required

def main():
    print("インストール済みパッケージを確認中...")
    installed_packages = get_installed_packages()
    
    print("requirements.txt を読み込み中...")
    required_packages = get_required_packages()
    
    # requirements.txtにないパッケージを特定
    packages_to_remove = []
    for package_name in installed_packages:
        if package_name not in required_packages:
            packages_to_remove.append(package_name)
    
    if not packages_to_remove:
        print("\n削除するパッケージはありません。")
        return
    
    print(f"\n以下の {len(packages_to_remove)} 個のパッケージが requirements.txt に含まれていません:")
    for package in sorted(packages_to_remove):
        print(f"  - {package}")
    
    # 確認
    response = input("\nこれらのパッケージを削除しますか? (y/N): ")
    if response.lower() != 'y':
        print("キャンセルしました。")
        return
    
    # パッケージを削除
    print("\nパッケージを削除中...")
    for package in packages_to_remove:
        print(f"削除中: {package}")
        subprocess.run(
            [sys.executable, "-m", "pip", "uninstall", "-y", package],
            capture_output=True
        )
    
    print("\n完了しました！")

if __name__ == "__main__":
    main()
