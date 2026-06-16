"""银行对账工具 — CustomTkinter 桌面版

双击即可启动，无需浏览器。
支持招商银行和工商银行（对账单）两种格式。
"""

import sys
import os
import threading
import tempfile
from datetime import datetime
from tkinter import filedialog, messagebox
import tkinter as tk

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

import customtkinter as ctk
import pandas as pd

from pipeline import run_reconciliation
from reporter import generate_report
from history import init_db, save_reconciliation, list_history, get_history, delete_history
from config import BANK_FORMATS

# ═══════════════════════════════════════════════════════════════════════
# App config
# ═══════════════════════════════════════════════════════════════════════

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

APP_TITLE = "银行对账工具"
WINDOW_WIDTH = 1100
WINDOW_HEIGHT = 750


# ═══════════════════════════════════════════════════════════════════════
# Main Application
# ═══════════════════════════════════════════════════════════════════════

class BankReconciliationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.minsize(900, 600)

        # State
        self.bank_path = None
        self.ledger_path = None
        self.result = None
        self.report_path = None
        self.is_running = False

        # Init DB
        try:
            init_db()
        except Exception:
            pass

        self._build_ui()

    # ── UI Construction ─────────────────────────────────────────────

    def _build_ui(self):
        # ── Grid layout ──────────────────────────────
        self.grid_columnconfigure(0, weight=0)  # sidebar fixed
        self.grid_columnconfigure(1, weight=1)  # main content
        self.grid_rowconfigure(0, weight=1)

        # ── Sidebar ──────────────────────────────────
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.sidebar.grid_propagate(False)

        ctk.CTkLabel(self.sidebar, text="📂 文件选择",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(15, 5))

        # Bank file
        self.bank_label = ctk.CTkLabel(self.sidebar, text="未选择银行流水",
                                        text_color="gray", wraplength=240)
        self.bank_label.pack(pady=(5, 2))
        ctk.CTkButton(self.sidebar, text="🏦 选择银行流水",
                      command=self._choose_bank).pack(pady=(0, 10), padx=15)

        # Ledger file
        self.ledger_label = ctk.CTkLabel(self.sidebar, text="未选择日记账",
                                          text_color="gray", wraplength=240)
        self.ledger_label.pack(pady=(5, 2))
        ctk.CTkButton(self.sidebar, text="📒 选择日记账",
                      command=self._choose_ledger).pack(pady=(0, 10), padx=15)

        # Bank format
        ctk.CTkLabel(self.sidebar, text="🏦 银行格式",
                     font=ctk.CTkFont(size=13)).pack(pady=(10, 2))
        self.format_var = ctk.StringVar(value="auto")
        self.format_menu = ctk.CTkOptionMenu(
            self.sidebar,
            values=["auto", "zhaoshang", "gonghang"],
            variable=self.format_var,
            command=self._on_format_change,
        )
        self.format_menu.pack(pady=(0, 10), padx=15)
        self._on_format_change("auto")

        # Parameters
        ctk.CTkLabel(self.sidebar, text="⚙️ 匹配参数",
                     font=ctk.CTkFont(size=13)).pack(pady=(10, 5))

        # Amount tolerance
        ctk.CTkLabel(self.sidebar, text="金额容差（元）").pack()
        self.amount_var = ctk.DoubleVar(value=0.01)
        ctk.CTkEntry(self.sidebar, textvariable=self.amount_var, width=120).pack(pady=(2, 8))

        # Date tolerance
        ctk.CTkLabel(self.sidebar, text="精确匹配日期容差（天）").pack()
        self.exact_days_var = ctk.IntVar(value=3)
        ctk.CTkSlider(self.sidebar, from_=1, to=10, number_of_steps=9,
                      variable=self.exact_days_var).pack(pady=(2, 2), padx=15)
        self.exact_days_label = ctk.CTkLabel(self.sidebar, text="3 天")
        self.exact_days_label.pack(pady=(0, 8))
        self.exact_days_var.trace_add("write", lambda *_: self.exact_days_label.configure(
            text=f"{self.exact_days_var.get()} 天"))

        # Fuzzy threshold
        ctk.CTkLabel(self.sidebar, text="模糊匹配阈值").pack()
        self.fuzzy_var = ctk.IntVar(value=70)
        ctk.CTkSlider(self.sidebar, from_=50, to=100, number_of_steps=10,
                      variable=self.fuzzy_var).pack(pady=(2, 2), padx=15)
        self.fuzzy_label = ctk.CTkLabel(self.sidebar, text="70%")
        self.fuzzy_label.pack(pady=(0, 8))
        self.fuzzy_var.trace_add("write", lambda *_: self.fuzzy_label.configure(
            text=f"{self.fuzzy_var.get()}%"))

        # Run button
        self.run_btn = ctk.CTkButton(self.sidebar, text="🚀 开始对账",
                                      command=self._run_reconciliation,
                                      height=40,
                                      font=ctk.CTkFont(size=15, weight="bold"))
        self.run_btn.pack(pady=(10, 5), padx=15)

        # Progress
        self.progress = ctk.CTkProgressBar(self.sidebar, mode="indeterminate")
        self.status_label = ctk.CTkLabel(self.sidebar, text="")

        # Export button
        self.export_btn = ctk.CTkButton(self.sidebar, text="📥 导出 Excel 报告",
                                         command=self._export_report,
                                         state="disabled")
        self.export_btn.pack(pady=(10, 5), padx=15)

        # History section
        ctk.CTkLabel(self.sidebar, text="📜 历史记录",
                     font=ctk.CTkFont(size=13)).pack(pady=(10, 5))
        self.history_frame = ctk.CTkScrollableFrame(self.sidebar, height=120)
        self.history_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self._refresh_history()

        # ── Main content ─────────────────────────────
        self.main = ctk.CTkFrame(self, corner_radius=0)
        self.main.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(0, weight=1)

        # Tab view
        self.tabview = ctk.CTkTabview(self.main)
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.tabview.add("📊 对账结果")
        self.tabview.add("⚠️ 差异交易")
        self.tabview.add("🔄 疑似重复")
        self.tabview.add("📥 导出")

        # Tab: 对账结果
        self.summary_text = ctk.CTkTextbox(self.tabview.tab("📊 对账结果"),
                                            font=ctk.CTkFont(size=14),
                                            wrap="word")
        self.summary_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.summary_text.insert("1.0", "👈 请在左侧选择银行流水和日记账文件，然后点击「开始对账」。\n\n"
                                 "支持格式：\n"
                                 "  • 招商银行 — 标准 xlsx 流水\n"
                                 "  • 工商银行 — 对账单格式（.xlsx）")

        # Tab: 差异交易
        self.diff_text = ctk.CTkTextbox(self.tabview.tab("⚠️ 差异交易"),
                                         font=ctk.CTkFont(size=13),
                                         wrap="word")
        self.diff_text.pack(fill="both", expand=True, padx=5, pady=5)

        # Tab: 疑似重复
        self.dup_text = ctk.CTkTextbox(self.tabview.tab("🔄 疑似重复"),
                                        font=ctk.CTkFont(size=13),
                                        wrap="word")
        self.dup_text.pack(fill="both", expand=True, padx=5, pady=5)

        # Tab: 导出
        self.export_text = ctk.CTkTextbox(self.tabview.tab("📥 导出"),
                                           font=ctk.CTkFont(size=14),
                                           wrap="word")
        self.export_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.export_text.insert("1.0", "对账完成后可导出 Excel 报告。\n\n"
                                "报告包含 5 个工作表：\n"
                                "  1. 对账汇总\n"
                                "  2. 匹配明细\n"
                                "  3. 银行独有\n"
                                "  4. 日记账独有\n"
                                "  5. 疑似重复\n\n"
                                "点击左侧「📥 导出 Excel 报告」按钮即可。")

    # ── Format dropdown handler ─────────────────────────────────────

    def _on_format_change(self, choice):
        label_map = {
            "auto": "自动检测",
            "zhaoshang": "招商银行",
            "gonghang": "工商银行（对账单）",
        }
        self.format_menu.configure(text=label_map.get(choice, choice))

    # ── File choosers ───────────────────────────────────────────────

    def _choose_bank(self):
        path = filedialog.askopenfilename(
            title="选择银行流水",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.bank_path = path
            self.bank_label.configure(text=os.path.basename(path), text_color="white")

    def _choose_ledger(self):
        path = filedialog.askopenfilename(
            title="选择日记账",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.ledger_path = path
            self.ledger_label.configure(text=os.path.basename(path), text_color="white")

    # ── Run reconciliation ─────────────────────────────────────────

    def _run_reconciliation(self):
        if not self.bank_path or not self.ledger_path:
            messagebox.showwarning("提示", "请先选择银行流水和日记账文件。")
            return

        if self.is_running:
            return

        self.is_running = True
        self.run_btn.configure(state="disabled", text="⏳ 对账中...")
        self.progress.pack(pady=(5, 2), padx=15)
        self.progress.start()
        self.status_label.pack(pady=(0, 5))
        self.status_label.configure(text="正在执行对账，请稍候...")

        # Run in background thread
        thread = threading.Thread(target=self._do_reconciliation, daemon=True)
        thread.start()

    def _do_reconciliation(self):
        try:
            from types import SimpleNamespace

            cfg = SimpleNamespace(
                AMOUNT_TOLERANCE=self.amount_var.get(),
                EXACT_MATCH_DAYS=self.exact_days_var.get(),
                FUZZY_MATCH_DAYS=7,
                FUZZY_SCORE_THRESHOLD=self.fuzzy_var.get(),
                SPLIT_MATCH_DAYS=7,
                DUPLICATE_CHECK_DAYS=3,
            )

            bank_fmt = self.format_var.get()
            result = run_reconciliation(
                self.bank_path, self.ledger_path,
                config_module=cfg,
                bank_format=bank_fmt,
            )

            self.result = result
            self.after(0, self._on_reconciliation_done)

        except Exception as e:
            import traceback
            err = f"{e}\n{traceback.format_exc()}"
            self.after(0, lambda: self._on_reconciliation_error(err))

    def _on_reconciliation_done(self):
        self.is_running = False
        self.progress.stop()
        self.progress.pack_forget()
        self.status_label.pack_forget()
        self.run_btn.configure(state="normal", text="🚀 开始对账")
        self.export_btn.configure(state="normal")

        result = self.result

        # ── Summary tab ──────────────────────────────
        summary = self.summary_text
        summary.delete("1.0", "end")

        lines = []
        lines.append("=" * 60)
        lines.append("  对账汇总")
        lines.append("=" * 60)
        lines.append("")
        lines.append(f"  🏦 银行交易总数:      {result.bank_total_count:>6} 笔")
        lines.append(f"  📒 日记账交易总数:    {result.ledger_total_count:>6} 笔")
        lines.append("")
        lines.append(f"  ✅ 匹配总数:          {result.matched_count:>6} 笔")
        lines.append(f"  📊 匹配率:            {result.match_rate:>8.2f}%")
        lines.append(f"  🎯 精确匹配:          {result.exact_matched:>6} 笔")
        lines.append(f"  🔍 模糊匹配:          {result.fuzzy_matched:>6} 笔")
        lines.append(f"  ✂️  拆分匹配:          {result.split_matched:>6} 笔")
        lines.append("")

        dr = result.date_range
        if dr and dr[0]:
            lines.append(f"  📅 日期范围:          {dr[0]} ~ {dr[1]}")
        lines.append("")
        lines.append("-" * 60)
        lines.append("  金额对比")
        lines.append("-" * 60)
        bank_net = result.bank_total_income - result.bank_total_expense
        ledger_net = result.ledger_total_income - result.ledger_total_expense
        lines.append(f"  {'项目':<12} {'银行（元）':>18} {'日记账（元）':>18} {'差额（元）':>18}")
        lines.append(f"  {'收入':<12} {result.bank_total_income:>18,.2f} {result.ledger_total_income:>18,.2f} {result.bank_total_income - result.ledger_total_income:>18,.2f}")
        lines.append(f"  {'支出':<12} {result.bank_total_expense:>18,.2f} {result.ledger_total_expense:>18,.2f} {result.bank_total_expense - result.ledger_total_expense:>18,.2f}")
        lines.append(f"  {'净额':<12} {bank_net:>18,.2f} {ledger_net:>18,.2f} {bank_net - ledger_net:>18,.2f}")

        if result.bank_opening_balance or result.bank_closing_balance:
            lines.append("")
            lines.append(f"  银行期初余额: {result.bank_opening_balance:,.2f} 元")
            lines.append(f"  银行期末余额: {result.bank_closing_balance:,.2f} 元")

        # Matched details (first 50)
        if result.matched_records:
            lines.append("")
            lines.append("-" * 60)
            lines.append("  匹配明细（精确/模糊/拆分，显示前 50 条）")
            lines.append("-" * 60)
            for i, rec in enumerate(result.matched_records[:50]):
                mt = {"exact": "精确", "fuzzy": "模糊", "split": "拆分"}.get(rec.match_type, rec.match_type)
                lines.append(f"  #{i+1:>3} [{mt}] 银行 {rec.bank_amount:>12,.2f} ⇔ 日记账 {rec.ledger_amount:>12,.2f}  |  日期差 {rec.date_diff}天" +
                           (f"  相似度 {rec.score:.0f}" if rec.score else ""))

        summary.insert("1.0", "\n".join(lines))

        # ── Diff tab ─────────────────────────────────
        self._fill_diffs(result)

        # ── Duplicate tab ────────────────────────────
        self._fill_duplicates(result)

        # ── Export tab ───────────────────────────────
        exp = self.export_text
        exp.delete("1.0", "end")
        exp_lines = []
        exp_lines.append("对账已完成！点击左侧「📥 导出 Excel 报告」按钮保存。")
        exp_lines.append("")
        exp_lines.append(f"匹配率: {result.match_rate:.2f}%")
        exp_lines.append(f"日期范围: {dr[0]} ~ {dr[1]}" if dr and dr[0] else "日期范围: 无")
        exp_lines.append("")
        exp_lines.append("报告将包含 5 个工作表：")
        exp_lines.append("  1. 对账汇总 — 总数、总额、匹配率")
        exp_lines.append("  2. 匹配明细 — 每笔交易如何匹配")
        exp_lines.append("  3. 银行独有 — 银行有但日记账无的交易")
        exp_lines.append("  4. 日记账独有 — 日记账有但银行无的交易")
        exp_lines.append("  5. 疑似重复 — 同源同日同金额的重复交易")
        exp.insert("1.0", "\n".join(exp_lines))

        # ── Save to history ──────────────────────────
        try:
            save_reconciliation(
                result, report_path="",
                notes=f"银行:{os.path.basename(self.bank_path)} 日记账:{os.path.basename(self.ledger_path)}",
            )
            self._refresh_history()
        except Exception:
            pass

        messagebox.showinfo("完成", f"✅ 对账完成！\n匹配率: {result.match_rate:.2f}%")

    def _on_reconciliation_error(self, err):
        self.is_running = False
        self.progress.stop()
        self.progress.pack_forget()
        self.status_label.pack_forget()
        self.run_btn.configure(state="normal", text="🚀 开始对账")
        messagebox.showerror("对账失败", err)

    def _fill_diffs(self, result):
        diff = self.diff_text
        diff.delete("1.0", "end")
        lines = []

        # Bank only
        bank_unmatched = result.unmatched_bank
        if bank_unmatched is not None and len(bank_unmatched) > 0:
            lines.append("=" * 60)
            lines.append(f"  🏦 银行独有交易 — {len(bank_unmatched)} 笔")
            lines.append("=" * 60)
            bank_amt = float(bank_unmatched['normalized_amount'].sum())
            lines.append(f"  总金额: {bank_amt:,.2f} 元")
            lines.append("")
            for _, row in bank_unmatched.iterrows():
                d = row.get('date')
                date_str = d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') and pd.notna(d) else ''
                lines.append(f"  {date_str} | {row.get('normalized_amount', 0):>12,.2f} | {row.get('summary', '')[:50]}")
        else:
            lines.append("✅ 所有银行交易均已匹配，无独有交易。")

        lines.append("")
        lines.append("")

        # Ledger only
        ledger_unmatched = result.unmatched_ledger
        if ledger_unmatched is not None and len(ledger_unmatched) > 0:
            lines.append("=" * 60)
            lines.append(f"  📒 日记账独有交易 — {len(ledger_unmatched)} 笔")
            lines.append("=" * 60)
            ledger_amt = float(ledger_unmatched['normalized_amount'].sum())
            lines.append(f"  总金额: {ledger_amt:,.2f} 元")
            lines.append("")
            for _, row in ledger_unmatched.iterrows():
                d = row.get('date')
                date_str = d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') and pd.notna(d) else ''
                lines.append(f"  {date_str} | {row.get('normalized_amount', 0):>12,.2f} | {row.get('summary', '')[:50]}")
        else:
            lines.append("✅ 所有日记账交易均已匹配，无独有交易。")

        diff.insert("1.0", "\n".join(lines))

    def _fill_duplicates(self, result):
        dup = self.dup_text
        dup.delete("1.0", "end")
        lines = []

        bank_dups = result.bank_duplicates
        if bank_dups is not None and len(bank_dups) > 0:
            lines.append("=" * 60)
            lines.append(f"  🏦 银行内部重复 — {len(bank_dups)} 笔")
            lines.append("=" * 60)
            for _, row in bank_dups.iterrows():
                d = row.get('date')
                date_str = d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') and pd.notna(d) else ''
                lines.append(f"  {date_str} | {row.get('normalized_amount', 0):>12,.2f} | {row.get('summary', '')[:50]}")
        else:
            lines.append("🏦 银行: 未发现重复交易。")

        lines.append("")

        ledger_dups = result.ledger_duplicates
        if ledger_dups is not None and len(ledger_dups) > 0:
            lines.append("=" * 60)
            lines.append(f"  📒 日记账内部重复 — {len(ledger_dups)} 笔")
            lines.append("=" * 60)
            for _, row in ledger_dups.iterrows():
                d = row.get('date')
                date_str = d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') and pd.notna(d) else ''
                lines.append(f"  {date_str} | {row.get('normalized_amount', 0):>12,.2f} | {row.get('summary', '')[:50]}")
        else:
            lines.append("📒 日记账: 未发现重复交易。")

        dup.insert("1.0", "\n".join(lines))

    # ── Export ──────────────────────────────────────────────────────

    def _export_report(self):
        if self.result is None:
            messagebox.showwarning("提示", "请先执行对账。")
            return

        save_path = filedialog.asksaveasfilename(
            title="保存 Excel 报告",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile=f"对账报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        )

        if not save_path:
            return

        try:
            generate_report(self.result, save_path)
            self.report_path = save_path
            messagebox.showinfo("导出成功", f"✅ 报告已保存到:\n{save_path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    # ── History ─────────────────────────────────────────────────────

    def _refresh_history(self):
        for widget in self.history_frame.winfo_children():
            widget.destroy()

        try:
            history = list_history(limit=20)
        except Exception:
            history = []

        if not history:
            ctk.CTkLabel(self.history_frame, text="暂无历史记录",
                         text_color="gray").pack(pady=10)
            return

        for rec in history:
            rid = rec['id']
            mr = rec.get('match_rate', 0) or 0
            dr = rec.get('date_range_start', '') or ''
            label = f"#{rid}  {mr:.1f}%  {dr}"

            frame = ctk.CTkFrame(self.history_frame)
            frame.pack(fill="x", pady=2, padx=2)

            btn = ctk.CTkButton(frame, text=label, anchor="w",
                                command=lambda r=rid: self._show_history_detail(r))
            btn.pack(side="left", fill="x", expand=True, padx=(0, 2))

            del_btn = ctk.CTkButton(frame, text="🗑", width=30,
                                     command=lambda r=rid: self._delete_history(r))
            del_btn.pack(side="right")

    def _show_history_detail(self, rid):
        try:
            detail = get_history(rid)
        except Exception:
            return

        if not detail:
            return

        info = (
            f"历史记录 #{rid}\n"
            f"{'='*40}\n"
            f"对账时间: {detail.get('reconciled_at', '')}\n"
            f"银行文件: {detail.get('bank_file', '')}\n"
            f"日记账文件: {detail.get('ledger_file', '')}\n"
            f"日期范围: {detail.get('date_range_start','')} ~ {detail.get('date_range_end','')}\n"
            f"银行笔数: {detail.get('bank_count', 0)}\n"
            f"日记账笔数: {detail.get('ledger_count', 0)}\n"
            f"匹配率: {detail.get('match_rate', 0):.2f}%\n"
            f"精确/模糊/拆分: {detail.get('exact_matched',0)}/{detail.get('fuzzy_matched',0)}/{detail.get('split_matched',0)}\n"
            f"银行独有: {detail.get('unmatched_bank_count', 0)}\n"
            f"日记账独有: {detail.get('unmatched_ledger_count', 0)}\n"
            f"金额差额: {detail.get('amount_diff', 0):.2f}"
        )
        messagebox.showinfo(f"历史记录 #{rid}", info)

    def _delete_history(self, rid):
        if messagebox.askyesno("确认删除", f"确定要删除历史记录 #{rid} 吗？"):
            try:
                delete_history(rid)
                self._refresh_history()
            except Exception as e:
                messagebox.showerror("删除失败", str(e))


# ═══════════════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = BankReconciliationApp()
    app.mainloop()
