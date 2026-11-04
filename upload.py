"""Commit local changes and push the main branch to GitHub."""

import subprocess
import sys
from datetime import datetime
from pathlib import Path


REPO_DIR = Path(__file__).resolve().parent
BRANCH = "main"
REMOTE = "origin"


def run_git(*args: str) -> subprocess.CompletedProcess[str]:
    """Run Git in this script's repository and preserve both output streams."""
    return subprocess.run(
        ["git", *args],
        cwd=REPO_DIR,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )


def show_output(result: subprocess.CompletedProcess[str]) -> None:
    if result.stdout:
        print(result.stdout.rstrip())
    if result.stderr:
        print(result.stderr.rstrip())


def fail(message: str, result: subprocess.CompletedProcess[str] | None = None) -> int:
    print(f"\n❌ {message}")
    if result is not None:
        show_output(result)
    return 1


def output_of(*args: str) -> tuple[str | None, subprocess.CompletedProcess[str] | None]:
    result = run_git(*args)
    if result.returncode != 0:
        show_output(result)
        return None, result
    return result.stdout.strip(), None


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    print("=" * 60)
    print("🚀 GitHub 업로드를 시작합니다")
    print(f"저장소: {REPO_DIR}")
    print("=" * 60)

    root, error = output_of("rev-parse", "--show-toplevel")
    if error:
        return fail("이 폴더가 Git 저장소가 아닙니다.", error)
    if Path(root).resolve() != REPO_DIR.resolve():
        return fail(f"업로드 스크립트 위치와 Git 저장소 루트가 다릅니다. (Git 루트: {root})")

    branch, error = output_of("branch", "--show-current")
    if error:
        return fail("현재 브랜치를 확인하지 못했습니다.", error)
    if branch != BRANCH:
        return fail(f"현재 브랜치는 '{branch}'입니다. '{BRANCH}' 브랜치에서 실행해주세요.")

    remote_url, error = output_of("remote", "get-url", REMOTE)
    if error:
        return fail(f"'{REMOTE}' 원격 저장소가 설정되어 있지 않습니다.", error)
    print(f"원격 저장소: {remote_url}")

    print("\n[1/4] GitHub의 최신 main 정보를 확인합니다...")
    fetched = run_git("fetch", REMOTE, BRANCH)
    show_output(fetched)
    if fetched.returncode != 0:
        return fail("GitHub에서 최신 정보를 가져오지 못했습니다. 네트워크와 로그인 상태를 확인해주세요.")

    # Do not create commits on a stale/diverged branch or overwrite remote work.
    remote_is_ancestor = run_git("merge-base", "--is-ancestor", f"{REMOTE}/{BRANCH}", "HEAD")
    if remote_is_ancestor.returncode not in (0, 1):
        return fail("로컬 브랜치와 GitHub 브랜치의 관계를 확인하지 못했습니다.", remote_is_ancestor)
    if remote_is_ancestor.returncode == 1:
        local_is_ancestor = run_git("merge-base", "--is-ancestor", "HEAD", f"{REMOTE}/{BRANCH}")
        if local_is_ancestor.returncode == 0:
            return fail("GitHub에만 새 커밋이 있습니다. 먼저 'git pull --rebase origin main'을 실행하고 다시 업로드해주세요.")
        if local_is_ancestor.returncode != 1:
            return fail("브랜치 상태를 확인하지 못했습니다.", local_is_ancestor)
        return fail("로컬과 GitHub의 main 브랜치가 서로 달라 자동 업로드를 중단했습니다. 충돌을 먼저 정리해주세요.")

    print("\n[2/4] 변경된 파일을 추가합니다...")
    added = run_git("add", "-A")
    if added.returncode != 0:
        return fail("git add에 실패했습니다.", added)

    staged = run_git("diff", "--cached", "--quiet")
    if staged.returncode == 1:
        message = f"update: 사이트 기능 업데이트 ({datetime.now().astimezone().strftime('%Y-%m-%d %H:%M:%S')})"
        print(f"\n[3/4] 커밋을 생성합니다: {message}")
        committed = run_git("commit", "-m", message)
        show_output(committed)
        if committed.returncode != 0:
            return fail("커밋 생성에 실패했습니다. 위 오류를 확인해주세요.")
    elif staged.returncode == 0:
        print("\n[3/4] 새로 커밋할 파일은 없습니다.")
    else:
        return fail("스테이징된 변경 파일을 확인하지 못했습니다.", staged)

    counts, error = output_of("rev-list", "--left-right", "--count", f"{REMOTE}/{BRANCH}...HEAD")
    if error:
        return fail("GitHub와 로컬 커밋 수를 비교하지 못했습니다.", error)
    try:
        behind, ahead = (int(part) for part in counts.split())
    except (ValueError, TypeError):
        return fail(f"커밋 비교 결과를 해석할 수 없습니다: {counts}")
    if behind:
        return fail("GitHub에 로컬에 없는 커밋이 있습니다. 원격 변경을 먼저 반영해주세요.")

    local_sha, error = output_of("rev-parse", "HEAD")
    if error:
        return fail("로컬 커밋을 확인하지 못했습니다.", error)

    if ahead:
        print(f"\n[4/4] 로컬 커밋 {ahead}개를 GitHub에 전송합니다...")
        pushed = run_git("push", REMOTE, BRANCH)
        show_output(pushed)
        if pushed.returncode != 0:
            return fail("GitHub 전송에 실패했습니다. 위 오류를 확인해주세요.")
    else:
        print("\n[4/4] 로컬과 GitHub가 이미 같은 커밋입니다.")

    refs, error = output_of("ls-remote", "--heads", REMOTE, f"refs/heads/{BRANCH}")
    if error:
        return fail("푸시 후 GitHub 반영 여부를 확인하지 못했습니다.", error)
    remote_sha = refs.split()[0] if refs else None
    if remote_sha != local_sha:
        return fail(f"GitHub 반영 확인에 실패했습니다. 로컬={local_sha}, GitHub={remote_sha or '브랜치 없음'}")

    print("\n" + "=" * 60)
    print("✅ GitHub 반영을 확인했습니다.")
    print(f"브랜치: {REMOTE}/{BRANCH}")
    print(f"커밋: {local_sha[:12]}")
    print(f"저장소: {remote_url}")
    print("Render 자동 배포는 GitHub 반영 후 시작됩니다.")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
