"""
엑셀 트래커 파일 감시 → 변경 시 대시보드 자동 갱신.

두 가지 트리거로 동작:
  1) **저장 감지 (즉시)**: 5초 간격으로 엑셀 mtime을 확인, 변경 감지 시 즉시 갱신
  2) **주기적 강제 갱신 (안전망)**: 5분마다 변경 여부와 무관하게 한 번 갱신

종료: Ctrl+C 또는 콘솔창 닫기
"""
from __future__ import annotations

import os
import sys
import time
import subprocess
from datetime import datetime
from pathlib import Path

# 설정
POLL_INTERVAL_SEC = 5         # 엑셀 mtime 체크 주기
DEBOUNCE_SEC = 2              # 저장 완료까지 대기
PERIODIC_REFRESH_MIN = 5      # 강제 갱신 주기 (분)

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

try:
    from config import EXCEL_PATH
except Exception as e:
    print(f"[ERROR] config.py 로드 실패: {e}")
    sys.exit(1)

EXCEL = Path(EXCEL_PATH)
MAKE_DASHBOARD = HERE / "make_dashboard.py"


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_mtime() -> float:
    try:
        return EXCEL.stat().st_mtime
    except FileNotFoundError:
        return 0.0


def run_dashboard_update(reason: str) -> bool:
    print(f"\n[{now_str()}] 대시보드 갱신 중... ({reason})")
    try:
        proc = subprocess.run(
            [sys.executable, str(MAKE_DASHBOARD), "--quiet"],
            cwd=str(HERE),
            capture_output=True,
            text=True,
            timeout=120,
        )
        if proc.returncode == 0:
            print(f"[{now_str()}] [OK] 대시보드 갱신 완료")
            return True
        else:
            print(f"[{now_str()}] [FAIL] 갱신 실패 (exit {proc.returncode})")
            if proc.stderr:
                print(f"   stderr: {proc.stderr[:500]}")
            return False
    except subprocess.TimeoutExpired:
        print(f"[{now_str()}] [WARN] 갱신 시간 초과 (120초)")
        return False
    except Exception as e:
        print(f"[{now_str()}] [ERR] 갱신 실행 에러: {e}")
        return False


def main():
    print("=" * 60)
    print(" 엑셀 트래커 감시 + 대시보드 자동 갱신")
    print("=" * 60)
    print(f" 감시 파일 : {EXCEL.name}")
    print(f" 폴링 주기 : {POLL_INTERVAL_SEC}초 (저장 감지)")
    print(f" 강제 갱신 : {PERIODIC_REFRESH_MIN}분마다")
    print(f" 시작 시각 : {now_str()}")
    print("=" * 60)
    print("종료하려면 Ctrl+C 또는 콘솔창을 닫으세요.\n")

    if not EXCEL.exists():
        print(f"[WARN] 엑셀 파일이 없어요: {EXCEL}")
        print("       config.py의 EXCEL_PATH를 확인하세요.")
        print("       파일이 생길 때까지 대기합니다...")

    # 시작 시 한 번 갱신
    run_dashboard_update("시작 시 초기 갱신")

    last_mtime = get_mtime()
    last_periodic = time.time()

    while True:
        try:
            time.sleep(POLL_INTERVAL_SEC)

            # 1) 저장 감지 - mtime 변경
            current_mtime = get_mtime()
            if current_mtime > 0 and current_mtime != last_mtime:
                # 저장 완료 보장을 위해 짧게 대기
                time.sleep(DEBOUNCE_SEC)
                confirmed_mtime = get_mtime()
                print(f"\n[{now_str()}] 엑셀 저장 감지됨")
                run_dashboard_update("엑셀 저장됨")
                last_mtime = confirmed_mtime
                last_periodic = time.time()
                continue

            # 2) 주기적 강제 갱신
            elapsed_min = (time.time() - last_periodic) / 60
            if elapsed_min >= PERIODIC_REFRESH_MIN:
                run_dashboard_update(f"{PERIODIC_REFRESH_MIN}분 주기 갱신")
                last_periodic = time.time()
                last_mtime = get_mtime()

        except KeyboardInterrupt:
            print(f"\n\n[{now_str()}] 감시 종료 (사용자 중단)")
            break
        except Exception as e:
            print(f"[{now_str()}] [WARN] 루프 에러 (계속 진행): {e}")
            time.sleep(POLL_INTERVAL_SEC)


if __name__ == "__main__":
    main()
