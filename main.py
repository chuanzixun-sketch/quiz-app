"""
工业级单机刷题 App — Flet 实现
=================================
- 数据源：同目录下的 questions.xls（Excel / HTML 自适应加载）
- 目标平台：Android APK（通过 flet pack 交叉编译）
- 设计原则：纯离线、无状态逻辑、高度健壮
"""

import random
import traceback
from pathlib import Path

import flet as ft

# ── 常量 ──────────────────────────────────────────────
DATA_FILE = Path(__file__).parent / "questions.xls"

# 原始列名 → 内部规范化列名的映射表
#   keys:    Pandas 读入后 strip 过的列名
#   values:  内部使用的英文 key
COLUMN_MAP = {
    "题目类型":   "qtype",
    "选择题题干": "stem",
    "正确答案":   "answer",
    "答案解析":   "explanation",
    "难易度":     "difficulty",
    "知识点":     "topic",
    "标签":       "tags",
    "选项数":     "opt_count",
    "选项A":      "opt_a",
    "选项B":      "opt_b",
    "选项C":      "opt_c",
    "选项D":      "opt_d",
    # HTML 格式变体（选项列名带空格）
    "选项 A":     "opt_a",
    "选项 B":     "opt_b",
    "选项 C":     "opt_c",
    "选项 D":     "opt_d",
}


# ═══════════════════════════════════════════════════════
# 一、数据管道
# ═══════════════════════════════════════════════════════

def load_data(filepath: Path) -> "pd.DataFrame":
    """四级容错加载：xlrd → openpyxl → read_html → 报错"""
    import pandas as pd

    if not filepath.exists():
        raise FileNotFoundError(f"题库文件未找到：{filepath}")

    # 1) 尝试 xlrd（老旧 .xls）
    try:
        df = pd.read_excel(str(filepath), engine="xlrd", sheet_name=0)
        return df
    except Exception:
        pass

    # 2) 尝试 openpyxl（.xlsx）
    try:
        df = pd.read_excel(str(filepath), engine="openpyxl", sheet_name=0)
        return df
    except Exception:
        pass

    # 3) 尝试 HTML 表格（Excel 另存为 .xls 但实际是 HTML 的情况）
    try:
        # read_html 返回 table 列表；默认用 lxml，不行则用 html5lib
        for backend in ["lxml", "html5lib", "bs4"]:
            try:
                tables = pd.read_html(str(filepath), encoding="utf-8", flavor=backend)
                if tables:
                    return tables[0]
            except Exception:
                continue
    except Exception:
        pass

    # 4) 所有路径都失败
    raise ValueError(
        f"无法解析题库文件 {filepath}。\n"
        "已尝试引擎：xlrd, openpyxl, HTML table (lxml/html5lib/bs4)。\n"
        "请确认文件格式正确且未损坏。"
    )


def clean_data(df: "pd.DataFrame") -> "pd.DataFrame":
    """
    严格清洗 & 类型强转：
    - 去表头空格
    - 列名映射到内部 key
    - fillna('') 防 NaN 污染
    - 所有文本列 astype(str).str.strip()
    - 正确答案 str.upper() 保证 "A"/"B"/"C"/"D"
    - 过滤掉表头行重复混入的数据行
    """
    import pandas as pd

    # 1) 表头去噪
    df.columns = df.columns.str.strip()

    # 2) 规范化选项列名：去掉空格（"选项 A" → "选项A"）
    rename_map = {}
    for col in df.columns:
        stripped = col.replace(" ", "").replace("　", "")
        if stripped != col:
            rename_map[col] = stripped
    if rename_map:
        df.rename(columns=rename_map, inplace=True)

    # 3) 映射到内部 key
    df.rename(columns={k: v for k, v in COLUMN_MAP.items() if k in df.columns}, inplace=True)

    # 4) 确认必要列存在
    required = ["stem", "answer"]
    for col in required:
        if col not in df.columns:
            raise KeyError(f"题库缺少必要列。已识别列：{list(df.columns)}，缺少：{col}")

    # 5) 过滤掉表头行重复混入的数据（第一行数据值等于列名）
    first_val = str(df.iloc[0]["answer"]).strip()
    if first_val in ("正确答案", "answer"):
        df = df.iloc[1:].copy()

    # 6) NaN → '' 对所有列
    df = df.fillna("")

    # 7) 文本列严格强转
    text_cols = ["stem", "answer", "explanation", "topic", "difficulty", "tags",
                 "opt_a", "opt_b", "opt_c", "opt_d"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # 8) 其他可能列也处理
    if "qtype" in df.columns:
        df["qtype"] = df["qtype"].astype(str).str.strip()
    if "opt_count" in df.columns:
        df["opt_count"] = df["opt_count"].astype(str).str.strip()

    # 9) 答案强转大写
    df["answer"] = df["answer"].str.upper().str.strip()

    # 10) 过滤无效行（题干为空 / 答案不是 A-D 单字母）
    df = df[df["stem"] != ""].copy()
    df = df[df["answer"].str.match(r"^[ABCD]$", na=False)].copy()

    # 11) 重置索引
    df.reset_index(drop=True, inplace=True)

    return df


def get_option_text(row: dict, letter: str) -> str:
    """获取某道题的某个选项文本；不存在则返回空字符串。"""
    key = f"opt_{letter.lower()}"
    return row.get(key, "")


# ═══════════════════════════════════════════════════════
# 二、主应用类
# ═══════════════════════════════════════════════════════

class QuizApp:
    def __init__(self, page: ft.Page, df: "pd.DataFrame"):
        self.page = page
        self.df_full = df                          # 全量题库
        self.df = df.copy()                        # 当前显示子集（按分类/错题筛选后）
        self.all_topics = sorted(df["topic"].unique().tolist())

        # ── 应用状态 ──
        self.current_index = 0
        self.answered = False                      # 当前题目是否已作答
        self.is_random = False
        self.selected_category = "全部"
        self.wrong_only = False                    # 只看错题模式
        self.wrong_set: set = set()                # 答错题目的原始索引集合
        self.score_correct = 0
        self.score_total = 0

        # 随机模式下的打乱索引
        self._shuffled_indices: list = []

        # ── UI 控件引用（延迟绑定）──
        self._build_ui()

    # ── 属性 ──────────────────────────────────────────

    @property
    def total(self) -> int:
        return len(self.df)

    @property
    def current_row(self) -> dict:
        """当前题目数据（DataFrame 行 → dict）"""
        if self.total == 0:
            return {}
        idx = self._effective_index()
        return self.df.iloc[idx].to_dict()

    def _effective_index(self) -> int:
        """映射逻辑索引 → DataFrame 行号（考虑随机打乱）"""
        if self.is_random and self._shuffled_indices:
            return self._shuffled_indices[self.current_index]
        return self.current_index

    def _original_index(self) -> int:
        """当前题目在 df_full 中的原始索引（用于错题记录）"""
        row = self.current_row
        return row.get("__orig_idx__", self._effective_index())

    # ── 数据筛选 ──────────────────────────────────────

    def apply_filter(self):
        """根据 selected_category 和 wrong_only 重建 df 并重置进度"""
        if self.wrong_only and self.wrong_set:
            # 只看错题
            wrong_indices = list(self.wrong_set)
            self.df = self.df_full.iloc[wrong_indices].copy()
            self.df["__orig_idx__"] = wrong_indices
        elif self.selected_category == "全部":
            self.df = self.df_full.copy()
        else:
            mask = self.df_full["topic"] == self.selected_category
            self.df = self.df_full[mask].copy()

        self.df.reset_index(drop=True, inplace=True)
        self.current_index = 0
        self.answered = False
        self._shuffled_indices = []
        if self.is_random and self.total > 0:
            self._shuffle()

    def _shuffle(self):
        """生成随机打乱索引"""
        self._shuffled_indices = list(range(self.total))
        random.shuffle(self._shuffled_indices)

    # ── 导航 ──────────────────────────────────────────

    def go_prev(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.answered = False
            self._refresh_question()

    def go_next(self):
        if self.current_index < self.total - 1:
            self.current_index += 1
            self.answered = False
            self._refresh_question()
        elif self.current_index == self.total - 1:
            # 最后一题 → 进入完成总结界面
            self.current_index += 1
            self._refresh_question()

    def go_restart(self):
        """重置所有进度并从头开始"""
        self.current_index = 0
        self.answered = False
        self.score_correct = 0
        self.score_total = 0
        self.wrong_set.clear()
        if self.is_random and self.total > 0:
            self._shuffle()
        self._refresh_question()

    def go_random(self):
        """随机跳转到一道新题"""
        if self.total <= 1:
            return
        old = self.current_index
        while self.current_index == old and self.total > 1:
            self.current_index = random.randint(0, self.total - 1)
        self.answered = False
        self._refresh_question()

    # ── 答题逻辑 ──────────────────────────────────────

    def on_select(self, chosen: str):
        """用户点击选项 A/B/C/D"""
        if self.answered or self.total == 0:
            return
        self.answered = True
        self.score_total += 1

        row = self.current_row
        correct = row.get("answer", "").upper()
        is_correct = (chosen == correct)

        if is_correct:
            self.score_correct += 1
        else:
            orig_idx = self._original_index()
            self.wrong_set.add(orig_idx)

        # 更新 UI 反馈
        self._show_feedback(chosen, correct)
        self._update_info_bar()

    # ── UI 构建（一次性）───────────────────────────────

    def _build_ui(self):
        page = self.page
        page.window_width = 420
        page.window_height = 820

        # ── 顶部信息栏 ──
        self.info_topic = ft.Text("", size=14, weight=ft.FontWeight.W_500)
        self.info_progress = ft.Text("", size=14, weight=ft.FontWeight.W_500)
        self.info_difficulty = ft.Text("", size=12)
        self.info_row = ft.Container(
            content=ft.ResponsiveRow(
                controls=[
                    ft.Column(col={"sm": 7}, controls=[self.info_topic, self.info_difficulty]),
                    ft.Column(col={"sm": 5}, controls=[self.info_progress],
                              horizontal_alignment=ft.CrossAxisAlignment.END),
                ],
            ),
            padding=ft.padding.symmetric(horizontal=16, vertical=10),
            bgcolor=ft.colors.SURFACE_CONTAINER_HIGHEST,
        )

        # 进度条
        self.progress_bar = ft.ProgressBar(value=0, height=6, color=ft.colors.PRIMARY)

        # ── 题干区域（可滚动）──
        self.qtype_badge = ft.Chip(label=ft.Text(""), leading=ft.Icon(ft.icons.QUIZ))
        self.stem_text = ft.Text("", size=16, weight=ft.FontWeight.W_500)

        # 选项按钮
        self.opt_buttons: list[ft.ElevatedButton] = []
        for letter in ["A", "B", "C", "D"]:
            btn = ft.ElevatedButton(
                text="",
                width=500,
                height=48,
                style=ft.ButtonStyle(
                    bgcolor=ft.colors.SURFACE_VARIANT,
                    color=ft.colors.ON_SURFACE,
                    shape=ft.RoundedRectangleBorder(radius=10),
                    padding=ft.padding.all(12),
                    text_style=ft.TextStyle(size=15),
                ),
                on_click=lambda e, l=letter: self.on_select(l),
            )
            self.opt_buttons.append(btn)

        # 答案解析面板（初始隐藏）
        self.explanation_text = ft.Text("", size=14, color=ft.colors.ON_SURFACE_VARIANT)
        self.explanation_container = ft.Container(
            content=ft.Column([
                ft.Divider(height=1),
                ft.Text("📖 解析", size=14, weight=ft.FontWeight.W_600),
                self.explanation_text,
            ]),
            padding=ft.padding.all(16),
            bgcolor=ft.colors.SURFACE_CONTAINER,
            border_radius=10,
            visible=False,
        )

        # 完成状态控件
        self.completion_icon = ft.Icon(ft.icons.CHECK_CIRCLE, size=48, color=ft.colors.PRIMARY, visible=False)
        self.completion_text = ft.Text("", size=16, visible=False)
        self.completion_detail = ft.Text("", size=13, color=ft.colors.ON_SURFACE_VARIANT, visible=False)
        self.btn_restart = ft.ElevatedButton(
            "🔄 重新开始", visible=False,
            on_click=lambda _: self.go_restart(),
            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
        )

        self.empty_text = ft.Text("📭 当前筛选条件下没有题目", size=16, visible=False)

        # ── 可滚动内容区 ──
        self.content_list = ft.ListView(
            expand=True,
            spacing=12,
            padding=ft.padding.all(16),
            controls=[
                self.completion_icon,
                self.completion_text,
                self.completion_detail,
                self.btn_restart,
                self.empty_text,
                self.qtype_badge,
                self.stem_text,
                self.opt_buttons[0],
                self.opt_buttons[1],
                self.opt_buttons[2],
                self.opt_buttons[3],
                self.explanation_container,
            ],
        )

        # ── 底部导航栏（固定）──
        self.btn_prev = ft.ElevatedButton(
            "◀ 上一题", on_click=lambda _: self.go_prev(),
            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
        )
        self.btn_next = ft.ElevatedButton(
            "下一题 ▶", on_click=lambda _: self.go_next(),
            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
        )
        self.btn_random = ft.OutlinedButton(
            "🎲 随机", on_click=lambda _: self.go_random(),
            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
        )
        self.nav_score = ft.Text("", size=13, weight=ft.FontWeight.W_500)

        nav_row_1 = ft.Row(
            controls=[self.btn_prev, self.nav_score, self.btn_next],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
        )
        nav_row_2 = ft.Row(
            controls=[self.btn_random],
            alignment=ft.MainAxisAlignment.CENTER,
        )

        # 筛选控件
        self.category_dd = ft.Dropdown(
            label="分类筛选",
            options=[ft.dropdown.Option("全部", "全部")] +
                    [ft.dropdown.Option(t, t) for t in self.all_topics],
            value="全部",
            dense=True,
            on_change=self._on_category_change,
        )
        self.random_switch = ft.Switch(
            label="随机模式", value=False, on_change=self._on_random_toggle,
        )
        self.wrong_switch = ft.Switch(
            label=f"错题本 ({len(self.wrong_set)})", value=False,
            on_change=self._on_wrong_toggle,
        )
        self.filter_row = ft.ResponsiveRow(
            controls=[
                ft.Column(col={"sm": 6}, controls=[self.category_dd]),
                ft.Column(col={"sm": 3}, controls=[self.random_switch]),
                ft.Column(col={"sm": 3}, controls=[self.wrong_switch]),
            ],
        )

        bottom_nav = ft.Container(
            content=ft.Column([
                self.filter_row,
                ft.Divider(height=1),
                nav_row_1,
                nav_row_2,
            ], spacing=6),
            padding=ft.padding.symmetric(horizontal=16, vertical=8),
            bgcolor=ft.colors.SURFACE,
            shadow=ft.BoxShadow(blur_radius=8, color=ft.colors.with_opacity(0.15, ft.colors.BLACK)),
        )

        # ── 组装页面 ──
        page.add(
            ft.Column([
                self.info_row,
                self.progress_bar,
                self.content_list,
                bottom_nav,
            ], spacing=0, expand=True)
        )

        # 首次渲染
        self._refresh_question()

    # ── 状态刷新 ──────────────────────────────────────

    def _refresh_question(self):
        """完整刷新当前题目显示"""
        if self.total == 0:
            self._render_empty()
            return

        row = self.current_row

        # 完成状态
        is_done = self.current_index >= self.total
        self.completion_icon.visible = is_done
        self.completion_text.visible = is_done
        self.completion_detail.visible = is_done
        self.btn_restart.visible = is_done
        self.empty_text.visible = False
        if is_done:
            total_q = self.score_total
            correct_q = self.score_correct
            pct = (correct_q / total_q * 100) if total_q > 0 else 0
            wrong_q = total_q - correct_q
            self.completion_icon.name = ft.icons.CHECK_CIRCLE if pct >= 60 else ft.icons.REFRESH
            self.completion_icon.color = ft.colors.PRIMARY if pct >= 60 else ft.colors.ORANGE_400
            self.completion_text.value = (
                f"🎉 本轮刷题结束！\n"
                f"正确率：{correct_q}/{total_q}（{pct:.0f}%）"
            )
            self.completion_detail.value = (
                f"✅ 正确：{correct_q} 题\n"
                f"❌ 错误：{wrong_q} 题\n"
                f"📝 错题本累计：{len(self.wrong_set)} 题\n"
                f"📂 当前分类：{self.selected_category}"
            )
            self._set_question_controls_visible(False)
            self.explanation_container.visible = False
            # 底部按钮切换
            self.btn_prev.visible = False
            self.btn_next.visible = False
            self.btn_random.visible = False
            self.random_switch.visible = False
            self.wrong_switch.visible = False
            self.category_dd.visible = False
            self.nav_score.value = ""
            self.info_progress.value = "🏁 完成"
            self.progress_bar.value = 1.0
            self.page.update()
            return

        # 解析当前行
        qtype = row.get("qtype", "")
        stem = row.get("stem", "")
        difficulty = row.get("difficulty", "")
        topic = row.get("topic", "")
        explanation = row.get("explanation", "")
        correct = row.get("answer", "").upper()

        # 更新信息栏
        self.info_topic.value = f"🏷️ 知识点：{topic}" if topic else ""
        self.info_difficulty.value = f"⭐ 难度：{difficulty} | 📋 {qtype}" if difficulty or qtype else ""
        self.info_progress.value = f"🔢 {self.current_index + 1} / {self.total}"
        self.progress_bar.value = (self.current_index + 1) / self.total if self.total > 0 else 0

        # 题目类型标签
        self.qtype_badge.label.value = qtype if qtype else "题目"
        self.qtype_badge.visible = bool(qtype)

        # 题干
        self.stem_text.value = stem

        # 选项按钮
        option_labels = ["A", "B", "C", "D"]
        for i, letter in enumerate(option_labels):
            opt_text = get_option_text(row, letter)
            btn = self.opt_buttons[i]
            if opt_text:
                btn.text = f"{letter}. {opt_text}"
                btn.visible = True
                btn.disabled = False
                btn.style = ft.ButtonStyle(
                    bgcolor=ft.colors.SURFACE_VARIANT,
                    color=ft.colors.ON_SURFACE,
                    shape=ft.RoundedRectangleBorder(radius=10),
                    padding=ft.padding.all(12),
                    text_style=ft.TextStyle(size=15),
                )
            else:
                btn.visible = False

        # 隐藏解析
        self.explanation_container.visible = False
        self.explanation_text.value = ""

        # 显示问题相关控件，隐藏完成/空态控件
        self._set_question_controls_visible(True)
        self.completion_icon.visible = False
        self.completion_text.visible = False
        self.completion_detail.visible = False
        self.btn_restart.visible = False
        self.empty_text.visible = False

        # 恢复底部控件可见性（从完成界面返回时）
        self.btn_prev.visible = True
        self.btn_next.visible = True
        self.btn_random.visible = True
        self.random_switch.visible = True
        self.wrong_switch.visible = True
        self.category_dd.visible = True

        # 更新导航按钮状态
        self.btn_prev.disabled = (self.current_index == 0)
        self.btn_next.disabled = (self.current_index >= self.total - 1)

        # 更新得分
        self._update_nav_score()

        # 更新错题开关标签
        self.wrong_switch.label = f"错题本 ({len(self.wrong_set)})"

        self.page.update()

    def _set_question_controls_visible(self, visible: bool):
        """批量设置题干区域控件的可见性（选项按钮的个体状态由 _refresh_question 单独控制）。"""
        self.qtype_badge.visible = visible
        self.stem_text.visible = visible
        if not visible:
            for btn in self.opt_buttons:
                btn.visible = False

    def _render_empty(self):
        """当前筛选无题目时的兜底 UI"""
        self.completion_icon.visible = False
        self.completion_text.visible = False
        self.completion_detail.visible = False
        self.btn_restart.visible = False
        self.empty_text.visible = True
        self._set_question_controls_visible(False)
        self.explanation_container.visible = False
        self.btn_prev.disabled = True
        self.btn_next.disabled = True
        self.nav_score.value = "0 / 0"
        self.info_topic.value = "📭 无题目"
        self.info_progress.value = ""
        self.progress_bar.value = 0
        self.page.update()

    def _show_feedback(self, chosen: str, correct: str):
        """答题后的视觉反馈"""
        row = self.current_row

        # 更新选项按钮颜色
        for i, letter in enumerate(["A", "B", "C", "D"]):
            btn = self.opt_buttons[i]
            if not btn.visible:
                continue
            btn.disabled = True
            if letter == correct:
                # 正确答案始终显示绿色
                btn.style = ft.ButtonStyle(
                    bgcolor=ft.colors.GREEN_ACCENT_700,
                    color=ft.colors.WHITE,
                    shape=ft.RoundedRectangleBorder(radius=10),
                    padding=ft.padding.all(12),
                    text_style=ft.TextStyle(size=15),
                )
            elif letter == chosen and chosen != correct:
                # 用户选错了 → 红色
                btn.style = ft.ButtonStyle(
                    bgcolor=ft.colors.RED_ACCENT_400,
                    color=ft.colors.WHITE,
                    shape=ft.RoundedRectangleBorder(radius=10),
                    padding=ft.padding.all(12),
                    text_style=ft.TextStyle(size=15),
                )

        # 显示解析
        explanation = row.get("explanation", "")
        correct_text = get_option_text(row, correct)
        if explanation:
            self.explanation_text.value = f"✅ 正确答案：{correct}. {correct_text}\n\n{explanation}"
        else:
            self.explanation_text.value = f"✅ 正确答案：{correct}. {correct_text}"
        self.explanation_container.visible = True

        self._update_nav_score()
        self.page.update()

    def _update_info_bar(self):
        row = self.current_row
        self.info_difficulty.value = (
            f"⭐ 难度：{row.get('difficulty', '')} | 📋 {row.get('qtype', '')}"
        )
        self.progress_bar.value = (self.current_index + 1) / self.total if self.total > 0 else 0

    def _update_nav_score(self):
        if self.score_total > 0:
            pct = self.score_correct / self.score_total * 100
            self.nav_score.value = f"✅ {self.score_correct} / {self.score_total} ({pct:.0f}%)"
        else:
            self.nav_score.value = ""

    # ── 事件处理 ──────────────────────────────────────

    def _on_category_change(self, e):
        self.selected_category = self.category_dd.value or "全部"
        self.wrong_only = False
        self.wrong_switch.value = False
        self.apply_filter()
        self._refresh_question()

    def _on_random_toggle(self, e):
        self.is_random = self.random_switch.value
        if self.is_random and self.total > 0:
            self._shuffle()
            self.current_index = 0
        else:
            self._shuffled_indices = []
            self.current_index = 0
        self.answered = False
        self._refresh_question()

    def _on_wrong_toggle(self, e):
        self.wrong_only = self.wrong_switch.value
        if self.wrong_only:
            self.is_random = False
            self.random_switch.value = False
            self.category_dd.value = "全部"
        self.apply_filter()
        self._refresh_question()


# ═══════════════════════════════════════════════════════
# 三、兼容性包装 & 入口
# ═══════════════════════════════════════════════════════

def build_assets_dir():
    """确保 flet pack 打包时的 assets 目录存在（用于包含 .xls 数据文件）。"""
    assets = Path(__file__).parent / "assets"
    assets.mkdir(exist_ok=True)
    return assets


def main(page: ft.Page):
    # 全局配置（布局由 QuizApp 内部控制，此处不做 scroll 设定以免冲突）
    page.theme_mode = ft.ThemeMode.LIGHT
    page.title = "刷题 App"
    page.vertical_alignment = ft.VerticalAlignment.TOP
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.padding = 0
    page.spacing = 0

    # 加载数据
    try:
        df_raw = load_data(DATA_FILE)
        df_clean = clean_data(df_raw)
    except Exception as exc:
        page.add(
            ft.Container(
                content=ft.Column([
                    ft.Icon(ft.icons.ERROR_OUTLINE, size=64, color=ft.colors.RED_400),
                    ft.Text("题库加载失败", size=22, weight=ft.FontWeight.W_600),
                    ft.Text(str(exc), size=14, color=ft.colors.ON_SURFACE_VARIANT),
                    ft.Text(traceback.format_exc(), size=11, color=ft.colors.OUTLINE),
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=12),
                padding=ft.padding.all(40),
                alignment=ft.alignment.center,
                expand=True,
            )
        )
        page.update()
        return

    if len(df_clean) == 0:
        page.add(
            ft.Container(
                content=ft.Column([
                    ft.Icon(ft.icons.SEARCH_OFF, size=64, color=ft.colors.ORANGE_400),
                    ft.Text("题库为空", size=22, weight=ft.FontWeight.W_600),
                    ft.Text("Excel 文件中未解析到有效题目，请检查数据格式。", size=14),
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=12),
                padding=ft.padding.all(40),
                alignment=ft.alignment.center,
                expand=True,
            )
        )
        page.update()
        return

    # 启动主应用
    QuizApp(page, df_clean)
    page.update()


# ── 入口 ─────────────────────────────────────────────
if __name__ == "__main__":
    build_assets_dir()
    ft.app(target=main, view=ft.AppView.FLET_APP)
