import subprocess
import sys
from datetime import datetime

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def run():
    print("=" * 60)
    print("🚀 깃허브(GitHub) 자동 업로드를 시작합니다...")
    print("=" * 60)
    print()

    # 1. git add -A
    print("[1/3] 변경된 파일 추가 중 (git add -A)...")
    res1 = subprocess.run(["git", "add", "-A"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    if res1.returncode != 0:
        print("❌ git add 실패:", res1.stderr)
        input("\n엔터 키를 누르면 창이 닫힙니다...")
        return

    # 2. git commit
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    commit_msg = f"update: 사이트 기능 업데이트 ({now_str})"
    print(f"[2/3] 커밋 생성 중: {commit_msg}")
    res2 = subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    print(res2.stdout)
    if "nothing to commit" in res2.stdout.lower():
        print("💡 변경된 새 파일이 없거나 이미 최신 상태입니다.")

    # 3. git push
    print("[3/3] 깃허브로 전송 중 (git push origin main)...")
    res3 = subprocess.run(["git", "push", "origin", "main"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    print(res3.stdout)
    print(res3.stderr)

    if res3.returncode == 0:
        print("=" * 60)
        print("✅ 깃허브 업로드 대성공!")
        print("🌐 https://github.com/kimwoongjae/jemma")
        print()
        print("👉 Render에서 1~2분 뒤 온라인 사이트가 자동 업데이트됩니다!")
        print("=" * 60)
    else:
        print("❌ 업로드 중 오류 발생:", res3.stderr)

    print()
    input("창을 닫으려면 엔터(Enter) 키를 누르세요...")

if __name__ == '__main__':
    run()
