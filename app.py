"""
app.py — 會計對帳系統
一次上傳三份檔案，自動跑完整流程並產出所有結果。

流程：
  Step 1  銀行對帳單 + 帳務查詢 → 未入帳清單_YYMM.xlsx + 對帳報告.html
  Step 2  未入帳清單（Step1產出）+ 帳務查詢 + 未銷帳明細 → 經辦人欄 + 經辦人報告.html
"""

from __future__ import annotations

import contextlib
import datetime
import io
import os
import re
import tempfile
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / ".env")

LOG_DIR = BASE_DIR / "log"
LOG_DIR.mkdir(exist_ok=True)

st.set_page_config(page_title="會計對帳系統", page_icon="🏦", layout="centered")
st.title("🏦 會計對帳系統")

# ── session_state 初始化 ───────────────────────────────────────────────────
if "run_results" not in st.session_state:
    st.session_state.run_results = None   # {summary, log_path, files}

# ── 三欄上傳 ───────────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)

with col1:
    bank_uploads = st.file_uploader(
        "銀行對帳單",
        accept_multiple_files=True,
        type=["xls", "xlsx"],
        help="可上傳多個月份，檔名需含「銀行對帳單-YYMM」",
    )
with col2:
    acc_uploads = st.file_uploader(
        "帳務查詢",
        accept_multiple_files=True,
        type=["xlsx", "xlsm"],
    )
with col3:
    ar_upload = st.file_uploader(
        "未銷帳明細表",
        type=["xlsm"],
    )

st.divider()

# ── 執行按鈕 ───────────────────────────────────────────────────────────────
can_run = bool(bank_uploads and acc_uploads)
if not can_run:
    st.info("請上傳**銀行對帳單**與**帳務查詢**後即可開始（未銷帳明細為選填）。")

if st.button("🚀 開始對帳", disabled=not can_run, type="primary", use_container_width=True):
    # 清掉上次結果
    st.session_state.run_results = None

    run_ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"run_{run_ts}.txt"
    log_file = open(log_path, "w", encoding="utf-8")

    summary_msgs: list[tuple[str, str]] = []   # (level, text)
    output_files: list[dict] = []              # {label, filename, mime, data}

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmpdir:

        # ── 寫入上傳檔案 ──
        bank_paths: dict[str, str] = {}
        for f in bank_uploads:
            path = os.path.join(tmpdir, f.name)
            with open(path, "wb") as fp:
                fp.write(f.getbuffer())
            m = re.search(r"銀行對帳單-(\d{5,8})", f.name)
            if m:
                bank_paths[m.group(1)[:5]] = path

        acc_paths: list[str] = []
        for f in acc_uploads:
            path = os.path.join(tmpdir, f.name)
            with open(path, "wb") as fp:
                fp.write(f.getbuffer())
            acc_paths.append(path)

        if ar_upload:
            ar_path = os.path.join(tmpdir, ar_upload.name)
            with open(ar_path, "wb") as fp:
                fp.write(ar_upload.getbuffer())

        if not bank_paths:
            log_file.close()
            st.error("找不到有效月份碼，請確認銀行對帳單檔名格式：銀行對帳單-YYMM*.xls(x)")
            st.stop()

        # ── Step 1 ──
        step1_results: dict = {}
        with st.spinner("Step 1 · 銀行對帳進行中..."):
            try:
                import reconcile_core as rc
                log_file.write(f"=== Step 1 開始 {run_ts} ===\n")
                captured = io.StringIO()
                with contextlib.redirect_stdout(captured):
                    step1_results = rc.run_pipeline(bank_paths, acc_paths, tmpdir)
                log_file.write(captured.getvalue())
                log_file.flush()
                months_done = sorted(step1_results.keys())
                summary_msgs.append(("success", f"✅ Step 1 完成，處理月份：{', '.join(months_done)}"))
            except Exception as e:
                log_file.write(f"\n[ERROR] Step 1: {e}\n")
                log_file.close()
                st.error(f"Step 1 失敗：{e}")
                st.stop()

        # ── Step 2 ──
        step2_results: dict = {}
        if ar_upload or acc_uploads:
            with st.spinner("Step 2 · 未入帳經辦人查找進行中..."):
                try:
                    import find_handler as fh
                    history_path = str(BASE_DIR / "handler_history.json")
                    log_file.write("\n=== Step 2 開始 ===\n")
                    captured2 = io.StringIO()
                    with contextlib.redirect_stdout(captured2):
                        step2_results = fh.run_pipeline(
                            data_dir=tmpdir,
                            history_path=history_path,
                            output_dir=tmpdir,
                        )
                    log_file.write(captured2.getvalue())
                    log_file.flush()
                    total = sum(len(v) for v in step2_results.values())
                    found = sum(sum(1 for x in v if x["method"] != 4) for v in step2_results.values())
                    summary_msgs.append(("success", f"✅ Step 2 完成：{found}/{total} 筆找到經辦人"))
                except Exception as e:
                    log_file.write(f"\n[ERROR] Step 2: {e}\n")
                    summary_msgs.append(("warning", f"⚠️ Step 2 失敗（不影響 Step 1 結果）：{e}"))

        log_file.close()

        # ── 將所有輸出檔讀成 bytes 存進 session_state（temp dir 關掉前一定要讀完）──
        # Step 2 的 update_excel 是原地修改同一份 xlsx，所以只讀一次即可
        for month_code, files in sorted(step1_results.items()):
            target_date = rc.month_code_to_date(month_code)
            xlsx_fp = files.get("xlsx", "")
            if xlsx_fp and os.path.exists(xlsx_fp):
                output_files.append({
                    "label": f"📊 未入帳清單 {target_date}",
                    "filename": f"未入帳清單_{month_code}.xlsx",
                    "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "data": open(xlsx_fp, "rb").read(),
                })
            html_fp = files.get("html", "")
            if html_fp and os.path.exists(html_fp):
                output_files.append({
                    "label": f"📋 對帳報告 {target_date}",
                    "filename": f"對帳報告_{month_code}.html",
                    "mime": "text/html",
                    "data": open(html_fp, "rb").read(),
                })

        fh_html = os.path.join(tmpdir, "未入帳_經辦人報告.html")
        if os.path.exists(fh_html):
            output_files.append({
                "label": "📋 經辦人報告",
                "filename": "未入帳_經辦人報告.html",
                "mime": "text/html",
                "data": open(fh_html, "rb").read(),
            })

    # temp dir 已關閉，資料全在 output_files 裡，存進 session_state
    st.session_state.run_results = {
        "summary": summary_msgs,
        "log_path": str(log_path),
        "files": output_files,
    }
    st.rerun()

# ── 顯示結果（從 session_state 讀取，不依賴 temp dir）─────────────────────
if st.session_state.run_results:
    res = st.session_state.run_results

    for level, msg in res["summary"]:
        if level == "success":
            st.success(msg)
        elif level == "warning":
            st.warning(msg)
        else:
            st.info(msg)

    st.subheader("📥 下載結果")
    st.caption(f"執行記錄：`{res['log_path']}`")

    dl_cols = st.columns(2)
    for i, f in enumerate(res["files"]):
        dl_cols[i % 2].download_button(
            f["label"],
            data=f["data"],
            file_name=f["filename"],
            mime=f["mime"],
            key=f"dl_{i}",
            use_container_width=True,
        )
