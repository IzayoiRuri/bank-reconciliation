"""银行对账工具 (Bank Reconciliation Tool) — Streamlit Web UI

完整的银行对账操作界面：
- 上传银行流水和公司日记账
- 配置匹配参数
- 显示对账结果（汇总/匹配明细/差异/重复）
- 导出 Excel 报告
- 历史记录管理
"""

import sys
sys.path.insert(0, 'src')

import os
import tempfile
from types import SimpleNamespace
from datetime import datetime

import streamlit as st
import pandas as pd

from pipeline import run_reconciliation, ReconciliationResult
from reporter import generate_report
from history import init_db, save_reconciliation, list_history, get_history, delete_history
import config as default_config


# ═══════════════════════════════════════════════════════════════════════
# Display helpers
# ═══════════════════════════════════════════════════════════════════════

def _fmt_date(date_val):
    """Format a date value for display. Handles NaT and Timestamp."""
    if date_val is None:
        return ''
    if pd.isna(date_val):
        return ''
    if hasattr(date_val, 'strftime'):
        return date_val.strftime('%Y-%m-%d')
    return str(date_val)


def _match_type_label(match_type):
    """Return a Chinese label for match type."""
    labels = {
        'exact': '🎯 精确匹配',
        'fuzzy': '🔍 模糊匹配',
        'split': '✂️ 拆分匹配',
    }
    return labels.get(match_type, match_type)


# ═══════════════════════════════════════════════════════════════════════
# Page config
# ═══════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="银行对账工具",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🔍 银行对账工具")
st.caption("上传银行流水和公司日记账，自动完成对账匹配与差异分析")


# ═══════════════════════════════════════════════════════════════════════
# Session state initialization
# ═══════════════════════════════════════════════════════════════════════

if "result" not in st.session_state:
    st.session_state.result = None
if "report_path" not in st.session_state:
    st.session_state.report_path = None
if "reconciliation_done" not in st.session_state:
    st.session_state.reconciliation_done = False
if "history_loaded" not in st.session_state:
    st.session_state.history_loaded = False


# ═══════════════════════════════════════════════════════════════════════
# Helper: build config object from UI parameters
# ═══════════════════════════════════════════════════════════════════════

def build_config(amount_tolerance, exact_days, fuzzy_days, fuzzy_threshold):
    """Build a SimpleNamespace config from user-provided UI parameters."""
    return SimpleNamespace(
        AMOUNT_TOLERANCE=amount_tolerance,
        EXACT_MATCH_DAYS=exact_days,
        FUZZY_MATCH_DAYS=fuzzy_days,
        FUZZY_SCORE_THRESHOLD=fuzzy_threshold,
        SPLIT_MATCH_DAYS=fuzzy_days,   # use same as fuzzy window
        DUPLICATE_CHECK_DAYS=3,         # keep default
    )


# ═══════════════════════════════════════════════════════════════════════
# Helper: save uploaded file to temp location
# ═══════════════════════════════════════════════════════════════════════

def save_uploaded_file(uploaded_file):
    """Save a Streamlit UploadedFile to a temporary location and return the path."""
    if uploaded_file is None:
        return None
    suffix = os.path.splitext(uploaded_file.name)[1] or '.xlsx'
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getvalue())
    tmp.close()
    return tmp.name


# ═══════════════════════════════════════════════════════════════════════
# Helper: run reconciliation (cached for performance)
# ═══════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def cached_reconciliation(bank_path, ledger_path, amount_tolerance,
                          exact_days, fuzzy_days, fuzzy_threshold, bank_format):
    """Run reconciliation with caching to avoid recomputation on re-render."""
    cfg = build_config(amount_tolerance, exact_days, fuzzy_days, fuzzy_threshold)
    result = run_reconciliation(bank_path, ledger_path, config_module=cfg, bank_format=bank_format)
    return result


# ═══════════════════════════════════════════════════════════════════════
# Sidebar — File Upload & Configuration
# ═══════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.header("📂 文件上传")

    bank_file = st.file_uploader(
        "🏦 银行流水 (.xlsx/.xls)",
        type=["xlsx", "xls"],
        key="bank_upload",
        help="上传银行的交易流水文件",
    )

    ledger_file = st.file_uploader(
        "📒 公司日记账 (.xlsx/.xls)",
        type=["xlsx", "xls"],
        key="ledger_upload",
        help="上传公司日记账文件",
    )

    st.divider()

    bank_format = st.selectbox(
        "🏦 银行格式",
        options=["auto", "zhaoshang", "gonghang"],
        format_func=lambda x: {
            "auto": "自动检测",
            "zhaoshang": "招商银行",
            "gonghang": "工商银行（对账单）",
        }.get(x, x),
        key="bank_format",
        help="选择银行流水文件的格式。'自动检测'会先尝试标准格式，失败后自动切换对账单格式。",
    )

    st.divider()

    # ── Config ──────────────────────────────────────────────────────
    with st.expander("⚙️ 参数配置", expanded=False):
        amount_tol = st.number_input(
            "金额容差（元）",
            min_value=0.0,
            max_value=100.0,
            value=default_config.AMOUNT_TOLERANCE,
            step=0.01,
            format="%.2f",
            help="匹配时允许的最大金额差异",
        )
        exact_days = st.slider(
            "精确匹配日期容差（天）",
            min_value=1,
            max_value=10,
            value=default_config.EXACT_MATCH_DAYS,
            step=1,
            help="精确匹配时银行和日记账日期的最大允许天数差",
        )
        fuzzy_days = st.slider(
            "模糊匹配日期容差（天）",
            min_value=1,
            max_value=30,
            value=default_config.FUZZY_MATCH_DAYS,
            step=1,
            help="模糊匹配时银行和日记账日期的最大允许天数差",
        )
        fuzzy_threshold = st.slider(
            "模糊匹配阈值",
            min_value=50,
            max_value=100,
            value=default_config.FUZZY_SCORE_THRESHOLD,
            step=5,
            help="模糊匹配的最低相似度分数（0-100），越高越严格",
        )

    st.divider()

    # ── Reconcile button ────────────────────────────────────────────
    can_reconcile = bank_file is not None and ledger_file is not None

    if st.button("▶️ 开始对账", type="primary", disabled=not can_reconcile,
                 use_container_width=True):
        if not can_reconcile:
            st.warning("请先上传银行流水和公司日记账文件")

        with st.spinner("正在执行对账，请稍候..."):
            # Save uploaded files
            bank_tmp = save_uploaded_file(bank_file)
            ledger_tmp = save_uploaded_file(ledger_file)

            try:
                # Clear cache to force re-run with new files
                cached_reconciliation.clear()

                result = cached_reconciliation(
                    bank_tmp, ledger_tmp,
                    amount_tol, exact_days, fuzzy_days, fuzzy_threshold,
                    bank_format,
                )

                st.session_state.result = result
                st.session_state.reconciliation_done = True
                st.session_state.report_path = None  # reset

                # Save to history
                try:
                    save_reconciliation(
                        result,
                        report_path="",
                        notes=f"银行:{bank_file.name} 日记账:{ledger_file.name}",
                    )
                except Exception as e:
                    st.warning(f"历史记录保存失败: {e}")

                st.success(f"✅ 对账完成！匹配率: {result.match_rate:.2f}%")

            except Exception as e:
                st.error(f"对账失败: {e}")
                import traceback
                st.code(traceback.format_exc())

            finally:
                # Clean up temp files
                for p in [bank_tmp, ledger_tmp]:
                    if p and os.path.exists(p):
                        try:
                            os.unlink(p)
                        except OSError:
                            pass

        st.rerun()

    st.divider()

    # ── History ─────────────────────────────────────────────────────
    st.subheader("📋 历史记录")

    try:
        history = list_history(limit=10)
        st.session_state.history_loaded = True
    except Exception:
        history = []

    if not history:
        st.caption("暂无历史记录")
    else:
        for rec in history:
            rid = rec['id']
            dr = rec.get('date_range_start', '') or ''
            match_rate = rec.get('match_rate', 0) or 0
            rec_time = (rec.get('reconciled_at', '') or '')[:19]

            col_h1, col_h2 = st.columns([6, 1])
            with col_h1:
                label = f"#{rid} | {match_rate:.1f}% | {dr}"
                if st.button(label, key=f"hist_{rid}", use_container_width=True):
                    detail = get_history(rid)
                    if detail:
                        st.session_state.history_detail = detail
                    with st.expander(f"📋 记录 #{rid} 详情", expanded=True):
                        st.json({
                            "对账时间": detail.get('reconciled_at', ''),
                            "银行文件": detail.get('bank_file', ''),
                            "日记账文件": detail.get('ledger_file', ''),
                            "日期范围": f"{detail.get('date_range_start','')} ~ {detail.get('date_range_end','')}",
                            "银行笔数": detail.get('bank_count', 0),
                            "日记账笔数": detail.get('ledger_count', 0),
                            "匹配率": f"{detail.get('match_rate', 0):.2f}%",
                            "精确/模糊/拆分": f"{detail.get('exact_matched',0)}/{detail.get('fuzzy_matched',0)}/{detail.get('split_matched',0)}",
                            "银行独有": detail.get('unmatched_bank_count', 0),
                            "日记账独有": detail.get('unmatched_ledger_count', 0),
                            "金额差额": f"{detail.get('amount_diff', 0):.2f}",
                        })
            with col_h2:
                if st.button("🗑️", key=f"del_{rid}", help="删除此记录"):
                    delete_history(rid)
                    st.rerun()


# ═══════════════════════════════════════════════════════════════════════
# Main Area — Tabs
# ═══════════════════════════════════════════════════════════════════════

if not st.session_state.reconciliation_done:
    # No reconciliation run yet
    st.info("👈 请在左侧上传银行流水和公司日记账文件，然后点击「开始对账」按钮。")
    st.stop()


result = st.session_state.result
if result is None:
    st.error("对账结果不存在，请重新执行对账。")
    st.stop()


tab1, tab2, tab3, tab4 = st.tabs([
    "📊 对账结果",
    "⚠️ 差异交易",
    "🔄 疑似重复",
    "📥 导出",
])


# ═══════════════════════════════════════════════════════════════════════
# Tab 1: 对账结果
# ═══════════════════════════════════════════════════════════════════════

with tab1:
    st.subheader("📊 对账汇总")

    # ── Summary metric cards ────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("🏦 银行交易数", result.bank_total_count)
        st.metric("📒 日记账交易数", result.ledger_total_count)
    with c2:
        st.metric("✅ 匹配总数", result.matched_count)
        st.metric("📊 匹配率", f"{result.match_rate:.2f}%")
    with c3:
        st.metric("🎯 精确匹配", result.exact_matched)
        st.metric("🔍 模糊匹配", result.fuzzy_matched)
    with c4:
        st.metric("✂️ 拆分匹配", result.split_matched)
        dr = result.date_range
        st.metric("📅 日期范围",
                  f"{dr[0]} ~ {dr[1]}" if dr and dr[0] else "无")

    st.divider()

    # ── Amount comparison table ─────────────────────────────────────
    st.subheader("💰 金额对比")

    bank_net = result.bank_total_income - result.bank_total_expense
    ledger_net = result.ledger_total_income - result.ledger_total_expense

    amt_data = {
        "项目": ["收入", "支出", "净额"],
        "银行 (元)": [
            f"{result.bank_total_income:,.2f}",
            f"{result.bank_total_expense:,.2f}",
            f"{bank_net:,.2f}",
        ],
        "日记账 (元)": [
            f"{result.ledger_total_income:,.2f}",
            f"{result.ledger_total_expense:,.2f}",
            f"{ledger_net:,.2f}",
        ],
        "差额 (元)": [
            f"{result.bank_total_income - result.ledger_total_income:,.2f}",
            f"{result.bank_total_expense - result.ledger_total_expense:,.2f}",
            f"{bank_net - ledger_net:,.2f}",
        ],
    }

    st.dataframe(
        pd.DataFrame(amt_data),
        use_container_width=True,
        hide_index=True,
    )

    # Show balance info if available
    if result.bank_opening_balance or result.bank_closing_balance:
        bc1, bc2 = st.columns(2)
        with bc1:
            st.metric("银行期初余额", f"{result.bank_opening_balance:,.2f} 元")
        with bc2:
            st.metric("银行期末余额", f"{result.bank_closing_balance:,.2f} 元")

    st.divider()

    # ── Matched detail table ────────────────────────────────────────
    st.subheader("📋 匹配明细")

    match_type_filter = st.selectbox(
        "匹配类型筛选",
        options=["全部", "精确匹配 (exact)", "模糊匹配 (fuzzy)", "拆分匹配 (split)"],
        key="match_filter",
    )

    if result.matched_records:
        # Build display rows
        type_map = {
            "全部": None,
            "精确匹配 (exact)": "exact",
            "模糊匹配 (fuzzy)": "fuzzy",
            "拆分匹配 (split)": "split",
        }
        filter_type = type_map[match_type_filter]

        rows = []
        bank_df = result.bank_df
        ledger_df = result.ledger_df

        for i, rec in enumerate(result.matched_records):
            if filter_type and rec.match_type != filter_type:
                continue

            # Bank data
            if bank_df is not None and len(bank_df) > 0 and rec.bank_idx < len(bank_df):
                br = bank_df.iloc[rec.bank_idx]
                b_date = _fmt_date(br.get('date'))
                b_summary = str(br.get('summary', ''))
            else:
                b_date = ''
                b_summary = ''

            # Ledger data
            if isinstance(rec.ledger_idx, list):
                if ledger_df is not None and len(ledger_df) > 0:
                    li0 = rec.ledger_idx[0] if rec.ledger_idx else 0
                    if li0 < len(ledger_df):
                        lr = ledger_df.iloc[li0]
                        l_date = _fmt_date(lr.get('date'))
                        l_summary = f"{lr.get('summary', '')} (+{len(rec.ledger_idx)-1}笔)"
                    else:
                        l_date = ''
                        l_summary = ''
                else:
                    l_date = ''
                    l_summary = ''
            else:
                if ledger_df is not None and len(ledger_df) > 0 and rec.ledger_idx < len(ledger_df):
                    lr = ledger_df.iloc[rec.ledger_idx]
                    l_date = _fmt_date(lr.get('date'))
                    l_summary = str(lr.get('summary', ''))
                else:
                    l_date = ''
                    l_summary = ''

            rows.append({
                "序号": i + 1,
                "匹配类型": _match_type_label(rec.match_type),
                "银行日期": b_date,
                "银行摘要": b_summary[:60],
                "银行金额": f"{rec.bank_amount:,.2f}",
                "日记账日期": l_date,
                "日记账摘要": l_summary[:60],
                "日记账金额": f"{rec.ledger_amount:,.2f}",
                "金额差": f"{rec.amount_diff:,.2f}",
                "日期差(天)": rec.date_diff,
                "相似度": f"{rec.score:.0f}" if rec.score else "",
            })

        if rows:
            detail_df = pd.DataFrame(rows)
            st.dataframe(
                detail_df,
                use_container_width=True,
                hide_index=True,
                height=400,
            )
            st.caption(f"共 {len(rows)} 条匹配记录")
        else:
            st.info("无匹配记录")
    else:
        st.info("无匹配记录")


# ═══════════════════════════════════════════════════════════════════════
# Tab 2: 差异交易
# ═══════════════════════════════════════════════════════════════════════

with tab2:
    subtab_bank, subtab_ledger = st.tabs(["🏦 银行独有", "📒 日记账独有"])

    # ── Bank Only ───────────────────────────────────────────────────
    with subtab_bank:
        bank_unmatched = result.unmatched_bank

        if bank_unmatched is None or len(bank_unmatched) == 0:
            st.success("✅ 所有银行交易均已匹配，无独有交易。")
        else:
            st.metric("银行独有交易数", len(bank_unmatched))
            bank_amt = float(bank_unmatched['normalized_amount'].sum())
            st.metric("银行独有总金额", f"{bank_amt:,.2f}")

            st.divider()

            # Build display table
            display_cols = []
            for col in ['date', 'summary', 'normalized_amount', 'balance',
                         'counterparty', 'counterparty_acct']:
                if col in bank_unmatched.columns:
                    display_cols.append(col)

            bank_disp = bank_unmatched[display_cols].copy()
            if 'date' in bank_disp.columns:
                bank_disp['date'] = bank_disp['date'].apply(_fmt_date)
            if 'normalized_amount' in bank_disp.columns:
                bank_disp['normalized_amount'] = bank_disp['normalized_amount'].apply(
                    lambda x: f"{x:,.2f}" if pd.notna(x) else "0.00"
                )

            bank_disp.columns = ["日期", "摘要", "金额", "余额", "对方名称", "对方账号"][:len(display_cols)]

            st.dataframe(
                bank_disp,
                use_container_width=True,
                hide_index=True,
                height=400,
            )

    # ── Ledger Only ─────────────────────────────────────────────────
    with subtab_ledger:
        ledger_unmatched = result.unmatched_ledger

        if ledger_unmatched is None or len(ledger_unmatched) == 0:
            st.success("✅ 所有日记账交易均已匹配，无独有交易。")
        else:
            st.metric("日记账独有交易数", len(ledger_unmatched))
            ledger_amt = float(ledger_unmatched['normalized_amount'].sum())
            st.metric("日记账独有总金额", f"{ledger_amt:,.2f}")

            st.divider()

            # Build display table
            display_cols = []
            for col in ['date', 'summary', 'normalized_amount',
                         'counterparty_subject', 'voucher_no', 'direction']:
                if col in ledger_unmatched.columns:
                    display_cols.append(col)

            ledger_disp = ledger_unmatched[display_cols].copy()
            if 'date' in ledger_disp.columns:
                ledger_disp['date'] = ledger_disp['date'].apply(_fmt_date)
            if 'normalized_amount' in ledger_disp.columns:
                ledger_disp['normalized_amount'] = ledger_disp['normalized_amount'].apply(
                    lambda x: f"{x:,.2f}" if pd.notna(x) else "0.00"
                )

            ledger_disp.columns = ["日期", "摘要", "金额", "对方科目", "结算号", "方向"][:len(display_cols)]

            st.dataframe(
                ledger_disp,
                use_container_width=True,
                hide_index=True,
                height=400,
            )


# ═══════════════════════════════════════════════════════════════════════
# Tab 3: 疑似重复
# ═══════════════════════════════════════════════════════════════════════

with tab3:
    has_bank_dup = (result.bank_duplicates is not None
                    and len(result.bank_duplicates) > 0)
    has_ledger_dup = (result.ledger_duplicates is not None
                      and len(result.ledger_duplicates) > 0)

    if not has_bank_dup and not has_ledger_dup:
        st.success("✅ 未发现疑似重复交易。")
    else:
        if has_bank_dup:
            st.subheader(f"🏦 银行疑似重复（{len(result.bank_duplicates)} 笔）")

            bd = result.bank_duplicates.copy()
            disp_cols = []
            for col in ['date', 'summary', 'normalized_amount', 'duplicate_group_id']:
                if col in bd.columns:
                    disp_cols.append(col)

            bd_disp = bd[disp_cols].copy()
            if 'date' in bd_disp.columns:
                bd_disp['date'] = bd_disp['date'].apply(_fmt_date)
            if 'normalized_amount' in bd_disp.columns:
                bd_disp['normalized_amount'] = bd_disp['normalized_amount'].apply(
                    lambda x: f"{x:,.2f}" if pd.notna(x) else "0.00"
                )

            bd_disp.columns = ["日期", "摘要", "金额", "重复组ID"][:len(disp_cols)]

            st.dataframe(
                bd_disp,
                use_container_width=True,
                hide_index=True,
                height=300,
            )

        if has_ledger_dup:
            st.subheader(f"📒 日记账疑似重复（{len(result.ledger_duplicates)} 笔）")

            ld = result.ledger_duplicates.copy()
            disp_cols = []
            for col in ['date', 'summary', 'normalized_amount', 'duplicate_group_id']:
                if col in ld.columns:
                    disp_cols.append(col)

            ld_disp = ld[disp_cols].copy()
            if 'date' in ld_disp.columns:
                ld_disp['date'] = ld_disp['date'].apply(_fmt_date)
            if 'normalized_amount' in ld_disp.columns:
                ld_disp['normalized_amount'] = ld_disp['normalized_amount'].apply(
                    lambda x: f"{x:,.2f}" if pd.notna(x) else "0.00"
                )

            ld_disp.columns = ["日期", "摘要", "金额", "重复组ID"][:len(disp_cols)]

            st.dataframe(
                ld_disp,
                use_container_width=True,
                hide_index=True,
                height=300,
            )


# ═══════════════════════════════════════════════════════════════════════
# Tab 4: 导出
# ═══════════════════════════════════════════════════════════════════════

with tab4:
    st.subheader("📥 导出 Excel 报告")

    st.markdown("""
    **报告包含以下内容：**

    1. **对账汇总** — 银行/日记账统计、匹配统计、差异分析、金额差额
    2. **匹配明细** — 所有匹配交易的详细信息（日期/摘要/金额/类型/相似度）
    3. **银行独有** — 仅在银行流水中出现的交易
    4. **日记账独有** — 仅在日记账中出现的交易
    5. **疑似重复** — 银行和日记账中可能重复的交易
    """)

    # Build default filename
    dr = result.date_range
    if dr and dr[0]:
        default_name = f"对账报告_{dr[0]}_{dr[1]}.xlsx"
    else:
        default_name = f"对账报告_{datetime.now().strftime('%Y%m%d')}.xlsx"

    if st.button("📥 生成并下载 Excel 报告", type="primary",
                 use_container_width=True):
        with st.spinner("正在生成报告..."):
            try:
                # Generate report to a temp file
                tmp_report = tempfile.NamedTemporaryFile(
                    delete=False, suffix='.xlsx'
                )
                tmp_report.close()

                report_path = generate_report(result, tmp_report.name)
                st.session_state.report_path = report_path

                # Read the file for download
                with open(report_path, 'rb') as f:
                    report_bytes = f.read()

                st.download_button(
                    label="⬇️ 下载报告",
                    data=report_bytes,
                    file_name=default_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                st.success(f"✅ 报告已生成: {default_name}")

                # Update history with report path
                try:
                    history_list = list_history(limit=1)
                    if history_list:
                        import sqlite3
                        conn = sqlite3.connect(default_config.DB_PATH)
                        conn.execute(
                            "UPDATE reconciliations SET report_path=? WHERE id=?",
                            (report_path, history_list[0]['id']),
                        )
                        conn.commit()
                        conn.close()
                except Exception:
                    pass  # Non-critical

            except Exception as e:
                st.error(f"报告生成失败: {e}")
                import traceback
                st.code(traceback.format_exc())

    # If report was already generated in this session
    if st.session_state.report_path and os.path.exists(st.session_state.report_path):
        st.divider()
        st.caption(f"已生成的报告: {os.path.basename(st.session_state.report_path)}")

        with open(st.session_state.report_path, 'rb') as f:
            st.download_button(
                label="⬇️ 重新下载报告",
                data=f.read(),
                file_name=os.path.basename(st.session_state.report_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
