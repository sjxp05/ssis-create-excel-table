# -*- coding: utf-8 -*-
"""
데이터 검증 자동화 시스템 — UI 스켈레톤
=======================================
사업지침(PDF) + 조견표(Excel) 업로드 → 검증 → 오류 수정 → 단가표 생성/검증

* 이 파일은 UI 껍데기만 포함합니다. 백엔드 연결 지점은 전부 `# TODO:` 로 표시.
* 백엔드 연결 지점 모음: 최하단 `BackendHooks` 클래스 하나만 채우면 됩니다.

실행:      python app.py            (pip install PyQt6)
EXE 빌드:  pyinstaller --onefile --windowed --name "데이터검증자동화" app.py

구조
----
  [1] Theme        : 색상/폰트 상수 + 공용 스타일 함수
  [2] Models       : ValidationError / UploadedFile 데이터 모델
  [3] Widgets      : 재사용 위젯 (UploadCard, ValidationTable, ErrorDetailPanel)
  [4] Pages        : 탭별 페이지 5개
  [5] MainWindow   : 조립 + 상태 관리
  [6] BackendHooks : ★ 백엔드 로직을 연결하는 유일한 지점 ★
"""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field
from typing import Callable, Optional

from PyQt6.QtCore import Qt, QDate, pyqtSignal
from PyQt6.QtGui import QBrush, QColor, QFont
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QFileDialog, QFrame, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton,
    QScrollArea, QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget,
)


# ═════════════════════════════════════════════════════════════════════════════
# [1] Theme — 색상/폰트/공용 스타일
# ═════════════════════════════════════════════════════════════════════════════
class Theme:
    PRIMARY     = "#1b3d7a"
    PRIMARY_HOV = "#274e94"
    SECONDARY   = "#eef1f7"
    BACKGROUND  = "#f8f8f7"
    CARD        = "#ffffff"
    BORDER      = "#e2e2e0"
    MUTED       = "#f0f0ee"
    MUTED_FG    = "#6b6b78"
    FG          = "#1a1a1f"

    RED_BG, RED_BORDER, RED, RED_DARK       = "#fef2f2", "#fecaca", "#dc2626", "#991b1b"
    GREEN_BG, GREEN_BORDER, GREEN, GREEN_DARK = "#f0fdf4", "#86efac", "#16a34a", "#166534"

    FONT_SANS = "Noto Sans KR"
    FONT_MONO = "Consolas"


T = Theme  # 짧은 별칭


def primary_button(text: str, height: int = 40, font_size: int = 13) -> QPushButton:
    b = QPushButton(text)
    b.setFixedHeight(height)
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    b.setStyleSheet(
        f"QPushButton{{background:{T.PRIMARY}; color:white; border:none; border-radius:4px;"
        f" font-size:{font_size}px; font-weight:600; padding:0 16px;}}"
        f"QPushButton:hover{{background:{T.PRIMARY_HOV};}}"
        f"QPushButton:disabled{{background:{T.MUTED}; color:{T.MUTED_FG};}}")
    return b


def outline_button(text: str, height: int = 32) -> QPushButton:
    b = QPushButton(text)
    b.setFixedHeight(height)
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    b.setStyleSheet(
        f"QPushButton{{background:{T.CARD}; color:{T.PRIMARY}; border:1px solid {T.PRIMARY};"
        f" border-radius:4px; font-size:11px; font-weight:500; padding:0 12px;}}"
        f"QPushButton:hover{{background:{T.SECONDARY};}}"
        f"QPushButton:disabled{{color:{T.MUTED_FG}; border-color:{T.BORDER};}}")
    return b


def section_label(text: str) -> QLabel:
    l = QLabel(text)
    l.setStyleSheet(f"font-size:10px; font-weight:600; color:{T.MUTED_FG}; letter-spacing:1px;")
    return l


TABLE_QSS = (
    f"QTableWidget{{background:white; alternate-background-color:{T.BACKGROUND};"
    f" gridline-color:{T.BORDER}; font-size:11px; font-family:'{T.FONT_MONO}';"
    f" border:1px solid {T.BORDER};}}"
    f"QHeaderView::section{{background:{T.PRIMARY}; color:white;"
    f" border:1px solid rgba(255,255,255,0.2); padding:6px 10px; font-size:11px;}}")


# ═════════════════════════════════════════════════════════════════════════════
# [2] Models — 백엔드와 UI 사이의 데이터 계약
# ═════════════════════════════════════════════════════════════════════════════
@dataclass
class ValidationError:
    """검증 오류 1건. 백엔드 결과를 이 형태로 변환해서 UI에 넘기면 됩니다."""
    id: str                 # 고유 ID
    sheet: str              # 시트명
    cell: str               # 셀 주소 표기 (예: "E3")
    row: int                # 테이블 행 인덱스 (0-base, 데이터 기준)
    col: int                # 테이블 열 인덱스 (0-base, 데이터 기준)
    field: str              # 항목명
    current: str            # 현재 값
    expected: str           # 올바른 값
    rule: str               # 적용 규칙 설명
    page: int               # 사업지침 근거 페이지
    section: str            # 사업지침 조항 (예: "제4조 제1항")
    fixed: bool = False     # 수정 여부 (UI가 관리)


@dataclass
class SheetData:
    """검증 대상 시트 1개 분량의 표 데이터."""
    filename: str = ""
    sheet_name: str = ""
    headers: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    errors: list[ValidationError] = field(default_factory=list)

    @property
    def open_errors(self) -> list[ValidationError]:
        return [e for e in self.errors if not e.fixed]

    @property
    def fixed_errors(self) -> list[ValidationError]:
        return [e for e in self.errors if e.fixed]


@dataclass
class UploadedFile:
    path: str
    name: str
    size: int

    @property
    def size_text(self) -> str:
        s = self.size
        if s < 1024:
            return f"{s} B"
        if s < 1024 ** 2:
            return f"{s / 1024:.1f} KB"
        return f"{s / 1024 ** 2:.1f} MB"


# ═════════════════════════════════════════════════════════════════════════════
# [3] Widgets — 재사용 위젯
# ═════════════════════════════════════════════════════════════════════════════
class UploadCard(QFrame):
    """드래그앤드롭 + 클릭 업로드 카드."""
    fileSelected = pyqtSignal(object)   # UploadedFile

    def __init__(self, title: str, desc: str, exts: list[str], icon: str = "📄"):
        super().__init__()
        self.exts = exts
        self.file: Optional[UploadedFile] = None
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 16)
        lay.setSpacing(12)

        top = QHBoxLayout()
        top.setSpacing(12)
        self._icon = QLabel(icon)
        self._icon.setFixedSize(36, 36)
        self._icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        top.addWidget(self._icon, 0, Qt.AlignmentFlag.AlignTop)
        txt = QVBoxLayout()
        txt.setSpacing(2)
        t = QLabel(title)
        t.setStyleSheet(f"font-weight:600; font-size:13px; color:{T.FG}; background:transparent;")
        d = QLabel(desc)
        d.setWordWrap(True)
        d.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG}; background:transparent;")
        txt.addWidget(t)
        txt.addWidget(d)
        top.addLayout(txt, 1)
        lay.addLayout(top)

        self._status = QLabel()
        lay.addWidget(self._status)
        self._render(state="idle")

    # ── 상태 렌더링 ──────────────────────────────────────────
    def _render(self, state: str):
        """state: idle | hover | done"""
        icon_css = f"border-radius:4px; font-size:16px;"
        if state == "done" and self.file:
            self.setStyleSheet(f"UploadCard{{background:#f6fdf8; border:1px solid {T.GREEN_BORDER}; border-radius:4px;}}")
            self._icon.setText("✔")
            self._icon.setStyleSheet(f"background:#dcfce7; color:{T.GREEN}; {icon_css}")
            self._status.setText(f"✔  {self.file.name}  ·  {self.file.size_text}")
            self._status.setStyleSheet(
                f"font-size:11px; color:{T.GREEN}; padding:8px 10px; border-radius:4px;"
                f" background:{T.GREEN_BG}; font-family:'{T.FONT_MONO}';")
        else:
            border = f"2px dashed {T.PRIMARY}" if state == "hover" else f"2px dashed {T.BORDER}"
            bg = T.SECONDARY if state == "hover" else T.CARD
            self.setStyleSheet(f"UploadCard{{background:{bg}; border:{border}; border-radius:4px;}}")
            self._icon.setStyleSheet(f"background:{T.SECONDARY}; color:{T.PRIMARY}; {icon_css}")
            self._status.setText("⬆  클릭하거나 파일을 끌어다 놓으세요")
            self._status.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG}; padding:8px 10px; background:transparent;")

    # ── 파일 세팅 ────────────────────────────────────────────
    def set_path(self, path: str):
        size = os.path.getsize(path) if os.path.exists(path) else 0
        self.file = UploadedFile(path=path, name=os.path.basename(path), size=size)
        self._render("done")
        self.fileSelected.emit(self.file)

    # ── 이벤트 ───────────────────────────────────────────────
    def mousePressEvent(self, e):
        filt = f"파일 ({' '.join('*' + x for x in self.exts)})"
        path, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", filt)
        if path:
            self.set_path(path)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            self._render("hover")
            e.acceptProposedAction()

    def dragLeaveEvent(self, e):
        self._render("done" if self.file else "idle")

    def dropEvent(self, e):
        urls = e.mimeData().urls()
        if urls:
            self.set_path(urls[0].toLocalFile())
        else:
            self._render("done" if self.file else "idle")


class ValidationTable(QTableWidget):
    """오류 셀 하이라이트 + 클릭 시 상세 패널 오픈용 표."""
    errorClicked = pyqtSignal(object)   # ValidationError

    def __init__(self):
        super().__init__()
        self.data: Optional[SheetData] = None
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.verticalHeader().setVisible(False)
        self.setAlternatingRowColors(True)
        self.setStyleSheet(TABLE_QSS)
        self.cellClicked.connect(self._on_click)

    def set_data(self, data: SheetData):
        self.data = data
        self.refresh()

    def _error_at(self, row: int, col: int) -> Optional[ValidationError]:
        if not self.data:
            return None
        for e in self.data.errors:
            if e.row == row and e.col == col:
                return e
        return None

    def refresh(self):
        if not self.data:
            self.setRowCount(0)
            self.setColumnCount(0)
            return
        self.setColumnCount(len(self.data.headers) + 1)
        self.setRowCount(len(self.data.rows))
        self.setHorizontalHeaderLabels(["#"] + self.data.headers)
        for r, row in enumerate(self.data.rows):
            num = QTableWidgetItem(str(r + 1))
            num.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            num.setForeground(QBrush(QColor(T.MUTED_FG)))
            self.setItem(r, 0, num)
            for c, val in enumerate(row):
                it = QTableWidgetItem(str(val))
                err = self._error_at(r, c)
                if err:
                    if err.fixed:
                        it.setText(err.current)
                        it.setBackground(QColor(T.GREEN_BG))
                        it.setForeground(QBrush(QColor(T.GREEN_DARK)))
                        it.setToolTip("수정 완료")
                    else:
                        it.setBackground(QColor(T.RED_BG))
                        it.setForeground(QBrush(QColor(T.RED_DARK)))
                        it.setToolTip(f"오류: {err.rule}")
                        f = it.font()
                        f.setBold(True)
                        it.setFont(f)
                self.setItem(r, c + 1, it)
        self.resizeColumnsToContents()
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.setColumnWidth(0, 36)

    def _on_click(self, r: int, c: int):
        err = self._error_at(r, c - 1)
        if err:
            self.errorClicked.emit(err)


class ErrorDetailPanel(QFrame):
    """우측 오류 상세 패널: 값 비교 / 규칙 / 근거 / 자동수정·직접입력."""
    closed = pyqtSignal()
    fixRequested = pyqtSignal(object, str)   # (ValidationError, value) value="" → 자동수정

    def __init__(self):
        super().__init__()
        self.error: Optional[ValidationError] = None
        self._mode = "auto"
        self.setFixedWidth(330)
        self.setStyleSheet(f"QFrame{{background:{T.CARD}; border-left:1px solid {T.BORDER};}}")

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self._build_header())
        root.addWidget(self._build_body(), 1)
        root.addWidget(self._build_footer())
        self._apply_mode()

    # ── 헤더 ────────────────────────────────────────────────
    def _build_header(self) -> QWidget:
        head = QFrame()
        head.setStyleSheet(f"background:{T.RED_BG}; border:none; border-bottom:1px solid {T.BORDER};")
        hl = QHBoxLayout(head)
        hl.setContentsMargins(18, 14, 12, 14)
        col = QVBoxLayout()
        col.setSpacing(4)
        row = QHBoxLayout()
        row.setSpacing(8)
        badge = QLabel("⨯ 오류")
        badge.setStyleSheet(
            f"background:{T.RED_BG}; color:{T.RED}; border:1px solid {T.RED_BORDER};"
            f" border-radius:4px; padding:2px 8px; font-size:11px;")
        self._loc = QLabel()
        self._loc.setStyleSheet(f"color:{T.MUTED_FG}; font-size:11px; font-family:'{T.FONT_MONO}';"
                                f" background:transparent; border:none;")
        row.addWidget(badge)
        row.addWidget(self._loc)
        row.addStretch()
        self._field = QLabel()
        self._field.setWordWrap(True)
        self._field.setStyleSheet(f"font-weight:600; font-size:13px; color:{T.FG};"
                                  f" background:transparent; border:none;")
        col.addLayout(row)
        col.addWidget(self._field)
        hl.addLayout(col, 1)
        close = QPushButton("✕")
        close.setFixedSize(24, 24)
        close.setCursor(Qt.CursorShape.PointingHandCursor)
        close.setStyleSheet(f"QPushButton{{border:none; background:transparent; color:{T.MUTED_FG};}}"
                            f"QPushButton:hover{{color:{T.FG};}}")
        close.clicked.connect(self.closed.emit)
        hl.addWidget(close, 0, Qt.AlignmentFlag.AlignTop)
        return head

    # ── 본문 ────────────────────────────────────────────────
    def _build_body(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea{border:none; background:transparent;}")
        body = QWidget()
        body.setStyleSheet(f"background:{T.CARD};")
        bl = QVBoxLayout(body)
        bl.setContentsMargins(18, 18, 18, 18)
        bl.setSpacing(18)

        bl.addWidget(section_label("값 비교"))
        cmp_row = QHBoxLayout()
        cmp_row.setSpacing(10)
        box_cur, self._cur = self._value_box("현재 값", T.RED_BG, T.RED_BORDER, T.RED, T.RED_DARK)
        box_exp, self._exp = self._value_box("올바른 값", T.GREEN_BG, T.GREEN_BORDER, T.GREEN, T.GREEN_DARK)
        cmp_row.addWidget(box_cur)
        cmp_row.addWidget(box_exp)
        bl.addLayout(cmp_row)

        bl.addWidget(section_label("적용 규칙"))
        self._rule = QLabel()
        self._rule.setWordWrap(True)
        self._rule.setStyleSheet(f"background:{T.MUTED}; border:1px solid {T.BORDER}; border-radius:4px;"
                                 f" padding:10px; font-size:12px; color:{T.FG};")
        bl.addWidget(self._rule)

        bl.addWidget(section_label("사업지침 근거"))
        ref = QFrame()
        ref.setStyleSheet(f"QFrame{{background:{T.SECONDARY}; border:1px solid #d0d9ea; border-radius:4px;}}")
        rl = QVBoxLayout(ref)
        rl.setContentsMargins(14, 12, 14, 12)
        rl.setSpacing(8)
        self._section = QLabel()
        self._section.setStyleSheet(f"color:{T.PRIMARY}; font-weight:600; font-size:12px;"
                                    f" background:transparent; border:none;")
        prow = QHBoxLayout()
        pl = QLabel("페이지")
        pl.setStyleSheet(f"color:{T.MUTED_FG}; font-size:11px; background:transparent; border:none;")
        self._page = QLabel()
        self._page.setStyleSheet(f"background:{T.PRIMARY}; color:white; border-radius:4px; border:none;"
                                 f" padding:4px 10px; font-size:11px; font-family:'{T.FONT_MONO}';")
        prow.addWidget(pl)
        prow.addWidget(self._page)
        prow.addStretch()
        rl.addWidget(self._section)
        rl.addLayout(prow)
        bl.addWidget(ref)
        bl.addStretch()
        scroll.setWidget(body)
        return scroll

    # ── 하단 수정 영역 ───────────────────────────────────────
    def _build_footer(self) -> QWidget:
        foot = QFrame()
        foot.setStyleSheet(f"background:{T.BACKGROUND}; border:none; border-top:1px solid {T.BORDER};")
        fl = QVBoxLayout(foot)
        fl.setContentsMargins(0, 0, 0, 0)
        fl.setSpacing(0)

        tabs = QHBoxLayout()
        tabs.setSpacing(0)
        self._tab_auto = QPushButton("✦ 자동 수정")
        self._tab_manual = QPushButton("✎ 직접 입력")
        for b, mode in ((self._tab_auto, "auto"), (self._tab_manual, "manual")):
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setFixedHeight(36)
            b.clicked.connect(lambda _, m=mode: self._set_mode(m))
            tabs.addWidget(b, 1)
        fl.addLayout(tabs)

        inner = QWidget()
        inner.setStyleSheet("background:transparent;")
        il = QVBoxLayout(inner)
        il.setContentsMargins(16, 12, 16, 14)
        il.setSpacing(8)

        self._auto_hint = QLabel()
        self._auto_hint.setStyleSheet(f"background:{T.GREEN_BG}; border:1px solid #bbf7d0; border-radius:4px;"
                                      f" padding:8px 10px; font-size:11px; color:{T.GREEN};")
        self._btn_auto = primary_button("✦  자동 수정 적용", 38)
        self._btn_auto.clicked.connect(lambda: self._fix(""))
        il.addWidget(self._auto_hint)
        il.addWidget(self._btn_auto)

        self._manual_hint = QLabel("수정할 값을 직접 입력하세요.")
        self._manual_hint.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        mrow = QHBoxLayout()
        mrow.setSpacing(8)
        self._input = QLineEdit()
        self._input.setFixedHeight(34)
        self._input.setStyleSheet(
            f"QLineEdit{{border:1px solid {T.BORDER}; border-radius:4px; padding:0 10px;"
            f" font-family:'{T.FONT_MONO}'; font-size:12px; background:white;}}"
            f"QLineEdit:focus{{border:1px solid {T.PRIMARY};}}")
        self._input.returnPressed.connect(self._apply_manual)
        self._btn_apply = primary_button("적용", 34, 12)
        self._btn_apply.setFixedWidth(60)
        self._btn_apply.clicked.connect(self._apply_manual)
        mrow.addWidget(self._input, 1)
        mrow.addWidget(self._btn_apply)
        self._recommend = QLabel()
        self._recommend.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        il.addWidget(self._manual_hint)
        il.addLayout(mrow)
        il.addWidget(self._recommend)

        self._done = QLabel("✔  수정 완료")
        self._done.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._done.setFixedHeight(40)
        self._done.setStyleSheet(f"background:#dcfce7; color:{T.GREEN}; border:1px solid {T.GREEN_BORDER};"
                                 f" border-radius:4px; font-size:13px; font-weight:600;")
        il.addWidget(self._done)
        fl.addWidget(inner)
        return foot

    # ── 내부 로직 ────────────────────────────────────────────
    @staticmethod
    def _value_box(title, bg, border, tc, vc):
        box = QFrame()
        box.setStyleSheet(f"QFrame{{background:{bg}; border:1px solid {border}; border-radius:4px;}}")
        l = QVBoxLayout(box)
        l.setContentsMargins(12, 10, 12, 10)
        l.setSpacing(5)
        t = QLabel(title)
        t.setStyleSheet(f"font-size:10px; font-weight:600; color:{tc}; background:transparent; border:none;")
        v = QLabel()
        v.setWordWrap(True)
        v.setStyleSheet(f"font-family:'{T.FONT_MONO}'; font-size:13px; font-weight:600; color:{vc};"
                        f" background:transparent; border:none;")
        l.addWidget(t)
        l.addWidget(v)
        return box, v

    def _set_mode(self, mode: str):
        self._mode = mode
        self._apply_mode()

    def _apply_mode(self):
        def tab_css(active):
            if active:
                return (f"QPushButton{{background:white; color:{T.PRIMARY}; border:none;"
                        f" border-bottom:2px solid {T.PRIMARY}; font-size:11px; font-weight:600;}}")
            return (f"QPushButton{{background:transparent; color:{T.MUTED_FG}; border:none;"
                    f" border-bottom:1px solid {T.BORDER}; font-size:11px;}}")
        self._tab_auto.setStyleSheet(tab_css(self._mode == "auto"))
        self._tab_manual.setStyleSheet(tab_css(self._mode == "manual"))
        fixed = bool(self.error and self.error.fixed)
        auto = self._mode == "auto"
        self._auto_hint.setVisible(auto and not fixed)
        self._btn_auto.setVisible(auto and not fixed)
        self._manual_hint.setVisible(not auto and not fixed)
        self._input.setVisible(not auto and not fixed)
        self._btn_apply.setVisible(not auto and not fixed)
        self._recommend.setVisible(not auto and not fixed)
        self._done.setVisible(fixed)

    def show_error(self, err: ValidationError):
        self.error = err
        self._loc.setText(f"{err.sheet} · 셀 {err.cell}")
        self._field.setText(err.field)
        self._cur.setText(err.current)
        self._exp.setText(err.expected)
        self._rule.setText(err.rule)
        self._section.setText(f"📖  사업지침 {err.section}")
        self._page.setText(f"p. {err.page}")
        self._auto_hint.setText(f"→  {err.expected} 으로 자동 수정됩니다")
        self._recommend.setText(f"ⓘ 권장값: {err.expected}")
        self._input.setPlaceholderText(err.current)
        self._input.clear()
        self._apply_mode()
        self.show()

    def _apply_manual(self):
        val = self._input.text().strip()
        if not val:
            QMessageBox.warning(self, "입력 오류", "값을 입력해주세요.")
            return
        self._fix(val)

    def _fix(self, value: str):
        if self.error:
            self.fixRequested.emit(self.error, value)
            self._apply_mode()


# ═════════════════════════════════════════════════════════════════════════════
# [4] Pages — 탭별 페이지
# ═════════════════════════════════════════════════════════════════════════════
def _centered_scroll_page(max_width: int) -> tuple[QScrollArea, QVBoxLayout]:
    """가운데 정렬 스크롤 페이지 골격."""
    page = QScrollArea()
    page.setWidgetResizable(True)
    page.setStyleSheet(f"QScrollArea{{border:none; background:{T.BACKGROUND};}}")
    inner = QWidget()
    inner.setStyleSheet(f"background:{T.BACKGROUND};")
    wrap = QHBoxLayout(inner)
    col = QVBoxLayout()
    col.setSpacing(20)
    col.setContentsMargins(0, 32, 0, 32)
    holder = QWidget()
    holder.setMaximumWidth(max_width)
    holder.setLayout(col)
    wrap.addStretch()
    wrap.addWidget(holder, 1)
    wrap.addStretch()
    page.setWidget(inner)
    return page, col


def _toolbar() -> tuple[QFrame, QHBoxLayout]:
    bar = QFrame()
    bar.setStyleSheet(f"background:{T.CARD}; border-bottom:1px solid {T.BORDER};")
    l = QHBoxLayout(bar)
    l.setContentsMargins(20, 10, 20, 10)
    l.setSpacing(12)
    return bar, l


def _legend() -> QFrame:
    bar = QFrame()
    bar.setStyleSheet(f"background:{T.BACKGROUND}; border-bottom:1px solid {T.BORDER};")
    l = QHBoxLayout(bar)
    l.setContentsMargins(20, 6, 20, 6)
    lg = QLabel("범례:   ■ 오류 셀 (클릭하여 상세 보기)")
    lg.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
    l.addWidget(lg)
    l.addStretch()
    return bar


class UploadPage(QWidget):
    """탭 1: 파일 업로드."""
    validateRequested = pyqtSignal()

    def __init__(self):
        super().__init__()
        page, col = _centered_scroll_page(640)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(page)

        crumb = QLabel("①  파일 업로드   ›   검증 실행   ›   결과 확인 및 수정")
        crumb.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        title = QLabel("파일 업로드")
        title.setStyleSheet(f"font-size:20px; font-weight:600; color:{T.FG};")
        sub = QLabel("사업지침 PDF와 조견표 엑셀 파일을 업로드하면 자동으로 데이터를 검증합니다.")
        sub.setStyleSheet(f"font-size:13px; color:{T.MUTED_FG};")
        col.addWidget(crumb)
        col.addWidget(title)
        col.addWidget(sub)

        self.card_guide = UploadCard(
            "사업지침 (Business Guidebook)",
            "PDF 형식의 사업지침 문서 · 규칙 추출 및 검증 기준으로 활용",
            [".pdf"], "📄")
        self.card_sheet = UploadCard(
            "조견표 (Quick Reference)",
            "Excel 형식의 조견표 파일 · 셀 단위로 사업지침과 대조 검증",
            [".xlsx", ".xls"], "▦")
        self.card_guide.fileSelected.connect(self._refresh)
        self.card_sheet.fileSelected.connect(self._refresh)
        col.addWidget(self.card_guide)
        col.addWidget(self.card_sheet)

        self.btn_validate = primary_button("🔍  데이터 검증 시작  →", 44)
        self.btn_validate.setEnabled(False)
        self.btn_validate.clicked.connect(self.validateRequested.emit)
        col.addWidget(self.btn_validate)
        self.hint = QLabel("두 파일을 모두 업로드해야 검증을 시작할 수 있습니다.")
        self.hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        col.addWidget(self.hint)
        col.addStretch()

    def _refresh(self):
        ready = bool(self.card_guide.file and self.card_sheet.file)
        self.btn_validate.setEnabled(ready)
        self.hint.setVisible(not ready)

    def set_busy(self, busy: bool):
        self.btn_validate.setEnabled(not busy)
        self.btn_validate.setText("⟳  사업지침과 조견표 대조 중..." if busy else "🔍  데이터 검증 시작  →")


class ValidationPage(QWidget):
    """탭 2/5 공용: 표 + 오류 상세 패널 검증 화면."""
    fixRequested = pyqtSignal(object, str)      # (ValidationError, value)
    fixAllRequested = pyqtSignal()
    saveRequested = pyqtSignal()

    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        bar, bl = _toolbar()
        self.lbl_file = QLabel()
        self.lbl_file.setStyleSheet(f"font-size:11px; font-weight:600; color:{T.FG};")
        self.lbl_sheet = QLabel()
        self.lbl_sheet.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        self.lbl_errors = QLabel()
        self.lbl_errors.setStyleSheet(f"font-size:11px; color:{T.RED};")
        bl.addWidget(self.lbl_file)
        bl.addWidget(self.lbl_sheet)
        bl.addWidget(self.lbl_errors)
        bl.addStretch()
        btn_all = primary_button("✦ 전체 자동 수정", 30, 11)
        btn_all.clicked.connect(self.fixAllRequested.emit)
        self.btn_save = outline_button("⬇ 수정 파일 저장")
        self.btn_save.clicked.connect(self.saveRequested.emit)
        bl.addWidget(btn_all)
        bl.addWidget(self.btn_save)
        v.addWidget(bar)
        v.addWidget(_legend())

        body = QHBoxLayout()
        body.setSpacing(0)
        wrap = QWidget()
        wl = QVBoxLayout(wrap)
        wl.setContentsMargins(20, 16, 20, 16)
        self.table = ValidationTable()
        self.table.errorClicked.connect(self.open_detail)
        wl.addWidget(self.table)
        body.addWidget(wrap, 1)

        self.panel = ErrorDetailPanel()
        self.panel.hide()
        self.panel.closed.connect(self.panel.hide)
        self.panel.fixRequested.connect(self.fixRequested.emit)
        body.addWidget(self.panel)
        v.addLayout(body, 1)

    def set_data(self, data: SheetData):
        self.table.set_data(data)
        self.lbl_file.setText(data.filename)
        self.lbl_sheet.setText(f"시트: {data.sheet_name}")
        self.refresh()

    def refresh(self):
        self.table.refresh()
        data = self.table.data
        n = len(data.open_errors) if data else 0
        self.lbl_errors.setText(f"● 오류 {n}건")
        self.btn_save.setEnabled(bool(data) and n == 0 and bool(data.errors))
        if self.panel.isVisible() and self.panel.error:
            self.panel.show_error(self.panel.error)

    def open_detail(self, err: ValidationError):
        self.panel.show_error(err)


class ErrorListPage(QWidget):
    """탭 3: 오류 목록 (검색/일괄수정/개별수정)."""
    fixRequested = pyqtSignal(object, str)
    fixAllRequested = pyqtSignal()
    saveRequested = pyqtSignal()
    rowActivated = pyqtSignal(object)           # 행 클릭 → 검증 탭으로 이동

    COLS = ["구분", "셀", "항목", "현재 값", "올바른 값", "적용 규칙", "근거", "상태/조치"]

    def __init__(self):
        super().__init__()
        self.data: Optional[SheetData] = None
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        bar, bl = _toolbar()
        self.search = QLineEdit()
        self.search.setPlaceholderText("🔍 오류 검색...")
        self.search.setFixedSize(260, 30)
        self.search.setStyleSheet(
            f"QLineEdit{{border:1px solid {T.BORDER}; border-radius:4px; padding:0 10px;"
            f" font-size:11px; background:white;}}"
            f"QLineEdit:focus{{border:1px solid {T.PRIMARY};}}")
        self.search.textChanged.connect(self.refresh)
        bl.addWidget(self.search)
        bl.addStretch()
        self.lbl_total = QLabel()
        self.lbl_total.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        btn_all = primary_button("✦ 전체 자동 수정", 30, 11)
        btn_all.clicked.connect(self.fixAllRequested.emit)
        self.btn_save = outline_button("⬇ 수정 파일 다운로드")
        self.btn_save.clicked.connect(self.saveRequested.emit)
        bl.addWidget(self.lbl_total)
        bl.addWidget(btn_all)
        bl.addWidget(self.btn_save)
        v.addWidget(bar)

        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(True)
        self.table.setStyleSheet(
            f"QTableWidget{{background:white; alternate-background-color:{T.BACKGROUND};"
            f" gridline-color:{T.BORDER}; font-size:11px; border:none;}}"
            f"QHeaderView::section{{background:{T.MUTED}; color:{T.MUTED_FG}; border:none;"
            f" border-bottom:1px solid {T.BORDER}; padding:8px 12px; font-size:11px; font-weight:600;}}")
        self.table.setColumnCount(len(self.COLS))
        self.table.setHorizontalHeaderLabels(self.COLS)
        self.table.cellClicked.connect(self._on_click)
        v.addWidget(self.table, 1)

        self.empty = QLabel("✔  오류가 없습니다")
        self.empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty.setStyleSheet(f"font-size:13px; color:{T.MUTED_FG}; padding:60px;")
        self.empty.hide()
        v.addWidget(self.empty)

    def set_data(self, data: SheetData):
        self.data = data
        self.refresh()

    def _filtered(self) -> list[ValidationError]:
        if not self.data:
            return []
        q = self.search.text().strip()
        if not q:
            return self.data.errors
        return [e for e in self.data.errors if q in e.field or q in e.cell or q in e.rule]

    def refresh(self):
        errs = self._filtered()
        self.lbl_total.setText(f"총 {len(errs)}건")
        self.empty.setVisible(bool(self.data) and not errs)
        self.table.setVisible(not self.empty.isVisible())
        n_open = len(self.data.open_errors) if self.data else 0
        self.btn_save.setEnabled(bool(self.data and self.data.errors) and n_open == 0)

        t = self.table
        t.setRowCount(len(errs))
        for r, e in enumerate(errs):
            def cell(text, color=T.FG, bg=None, mono=False, center=False, bold=False):
                it = QTableWidgetItem(text)
                it.setForeground(QBrush(QColor("#9a9aa5" if e.fixed else color)))
                if bg:
                    it.setBackground(QColor(bg))
                if mono:
                    it.setFont(QFont(T.FONT_MONO, 9))
                if center:
                    it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if bold:
                    f = it.font(); f.setBold(True); it.setFont(f)
                return it
            t.setItem(r, 0, cell("⨯ 오류", T.RED, center=True))
            t.setItem(r, 1, cell(e.cell, mono=True, bold=True, center=True))
            t.setItem(r, 2, cell(e.field, bold=True))
            t.setItem(r, 3, cell(e.current, T.RED, T.RED_BG, mono=True))
            t.setItem(r, 4, cell(e.expected, T.GREEN, T.GREEN_BG, mono=True))
            t.setItem(r, 5, cell(e.rule, T.MUTED_FG))
            t.setItem(r, 6, cell(f"p.{e.page}", T.PRIMARY, T.SECONDARY, mono=True, center=True, bold=True))
            t.setItem(r, 7, cell("✔ 수정됨" if e.fixed else "자동 수정",
                                 T.GREEN if e.fixed else "white",
                                 None if e.fixed else T.PRIMARY, center=True, bold=True))
        t.resizeColumnsToContents()
        t.setColumnWidth(5, 380)
        t.resizeRowsToContents()

    def _on_click(self, r: int, c: int):
        errs = self._filtered()
        if r >= len(errs):
            return
        e = errs[r]
        if c == 7 and not e.fixed:
            self.fixRequested.emit(e, "")
        else:
            self.rowActivated.emit(e)


class UnitPricePage(QWidget):
    """탭 4: 단가표 생성."""
    generateRequested = pyqtSignal()
    downloadRequested = pyqtSignal()
    validateRequested = pyqtSignal()

    def __init__(self):
        super().__init__()
        page, col = _centered_scroll_page(860)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(page)

        title = QLabel("단가표 자동 생성")
        title.setStyleSheet(f"font-size:20px; font-weight:600; color:{T.FG};")
        sub = QLabel("검증된 조견표를 기준으로 단가표를 자동 생성합니다. 사업지침의 요율 기준이 자동으로 반영됩니다.")
        sub.setWordWrap(True)
        sub.setStyleSheet(f"font-size:13px; color:{T.MUTED_FG};")
        col.addWidget(title)
        col.addWidget(sub)

        # 상태 박스
        self._status_box = QFrame()
        sl = QHBoxLayout(self._status_box)
        sl.setContentsMargins(18, 16, 18, 16)
        sl.setSpacing(14)
        self._status_icon = QLabel()
        self._status_icon.setFixedSize(36, 36)
        self._status_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        stxt = QVBoxLayout()
        stxt.setSpacing(2)
        self._status_title = QLabel()
        self._status_title.setStyleSheet(f"font-size:13px; font-weight:600; color:{T.FG};"
                                         f" background:transparent; border:none;")
        self._status_desc = QLabel()
        self._status_desc.setWordWrap(True)
        self._status_desc.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};"
                                        f" background:transparent; border:none;")
        stxt.addWidget(self._status_title)
        stxt.addWidget(self._status_desc)
        sl.addWidget(self._status_icon, 0, Qt.AlignmentFlag.AlignTop)
        sl.addLayout(stxt, 1)
        col.addWidget(self._status_box)

        # 단가표 미리보기
        prev = QFrame()
        prev.setStyleSheet(f"QFrame{{border:1px solid {T.BORDER}; border-radius:4px; background:white;}}")
        pl = QVBoxLayout(prev)
        pl.setContentsMargins(0, 0, 0, 0)
        pl.setSpacing(0)
        ph = QFrame()
        ph.setStyleSheet(f"background:{T.MUTED}; border:none; border-bottom:1px solid {T.BORDER};")
        phl = QHBoxLayout(ph)
        phl.setContentsMargins(16, 10, 16, 10)
        pt = QLabel("단가표 미리보기")
        pt.setStyleSheet(f"font-size:11px; font-weight:600; color:{T.FG}; background:transparent; border:none;")
        ps = QLabel("사업지침 기준 자동 산출")
        ps.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG}; background:transparent; border:none;")
        phl.addWidget(pt)
        phl.addStretch()
        phl.addWidget(ps)
        pl.addWidget(ph)
        self.preview = QTableWidget()
        self.preview.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.preview.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.preview.verticalHeader().setVisible(False)
        self.preview.setAlternatingRowColors(True)
        self.preview.setMinimumHeight(320)
        self.preview.setStyleSheet(TABLE_QSS.replace(f"border:1px solid {T.BORDER}", "border:none"))
        pl.addWidget(self.preview)
        col.addWidget(prev)

        # 생성/다운로드
        row = QHBoxLayout()
        row.setSpacing(12)
        self.btn_generate = primary_button("▤  단가표 Excel 파일 생성", 44)
        self.btn_generate.clicked.connect(self.generateRequested.emit)
        self.btn_download = outline_button("⬇ 다운로드", 44)
        self.btn_download.hide()
        self.btn_download.clicked.connect(self.downloadRequested.emit)
        row.addWidget(self.btn_generate, 1)
        row.addWidget(self.btn_download)
        col.addLayout(row)

        self.done_banner = QLabel()
        self.done_banner.setStyleSheet(f"background:{T.GREEN_BG}; color:{T.GREEN_DARK};"
                                       f" border:1px solid {T.GREEN_BORDER}; border-radius:4px;"
                                       f" padding:12px 14px; font-size:12px; font-weight:600;")
        self.done_banner.hide()
        col.addWidget(self.done_banner)

        # 다음 단계
        self.next_box = QWidget()
        nl = QVBoxLayout(self.next_box)
        nl.setContentsMargins(0, 12, 0, 0)
        nl.setSpacing(10)
        nt = QLabel("다음 단계: 단가표 검증")
        nt.setStyleSheet(f"font-size:13px; font-weight:600; color:{T.FG};")
        nd = QLabel("생성된 단가표를 사업지침과 대조하여 데이터를 검증합니다.")
        nd.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        self.btn_validate = primary_button("🔍  단가표 검증 시작  →", 44)
        self.btn_validate.clicked.connect(self.validateRequested.emit)
        nl.addWidget(nt)
        nl.addWidget(nd)
        nl.addWidget(self.btn_validate)
        self.next_box.hide()
        col.addWidget(self.next_box)
        col.addStretch()
        self.set_ready(False)

    def set_ready(self, ready: bool, total: int = 0, fixed: int = 0):
        """조견표 검증 완료 여부에 따라 상태 박스/버튼 갱신."""
        icon_css = "border-radius:4px; font-size:16px; border:none;"
        if ready:
            self._status_box.setStyleSheet(
                f"QFrame{{background:{T.GREEN_BG}; border:1px solid {T.GREEN_BORDER}; border-radius:4px;}}")
            self._status_icon.setText("✔")
            self._status_icon.setStyleSheet(f"background:#dcfce7; color:{T.GREEN}; {icon_css}")
            self._status_title.setText("검증 완료 — 단가표 생성 가능")
            self._status_desc.setText(f"조견표 검증 결과 {total}건이 확인되었으며, {fixed}건이 수정되었습니다.")
        else:
            self._status_box.setStyleSheet(
                f"QFrame{{background:{T.CARD}; border:1px solid {T.BORDER}; border-radius:4px;}}")
            self._status_icon.setText("⚠")
            self._status_icon.setStyleSheet(f"background:{T.MUTED}; color:{T.MUTED_FG}; {icon_css}")
            self._status_title.setText("검증 미완료")
            self._status_desc.setText("먼저 파일을 업로드하고 데이터 검증을 완료하세요.")
        self.btn_generate.setEnabled(ready)

    def set_preview(self, headers: list[str], rows: list[list[str]]):
        p = self.preview
        p.setColumnCount(len(headers))
        p.setRowCount(len(rows))
        p.setHorizontalHeaderLabels(headers)
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                p.setItem(r, c, QTableWidgetItem(str(val)))
        p.resizeColumnsToContents()

    def set_generated(self, filename: str, size_text: str):
        self.btn_download.show()
        self.done_banner.setText(f"✔  단가표 생성 완료    {filename} ({size_text})")
        self.done_banner.show()
        self.next_box.show()

    def set_busy(self, which: str, busy: bool):
        """which: generate | validate"""
        if which == "generate":
            self.btn_generate.setEnabled(not busy)
            self.btn_generate.setText("⟳  단가표 생성 중..." if busy else "▤  단가표 Excel 파일 생성")
        else:
            self.btn_validate.setEnabled(not busy)
            self.btn_validate.setText("⟳  단가표 검증 중..." if busy else "🔍  단가표 검증 시작  →")


class UnitValidationPage(QWidget):
    """탭 5: 단가표 검증 (빈 상태 ↔ 검증 화면 전환)."""
    goGenerate = pyqtSignal()

    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        # 빈 상태
        self.empty = QWidget()
        el = QVBoxLayout(self.empty)
        el.setAlignment(Qt.AlignmentFlag.AlignCenter)
        el.setSpacing(14)
        ei = QLabel("▤")
        ei.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ei.setStyleSheet("font-size:36px; color:#c8c8c4;")
        et = QLabel("단가표 생성 후 검증을 시작하세요.")
        et.setAlignment(Qt.AlignmentFlag.AlignCenter)
        et.setStyleSheet(f"font-size:13px; color:{T.MUTED_FG};")
        eb = primary_button("단가표 생성 탭으로 이동", 38)
        eb.clicked.connect(self.goGenerate.emit)
        el.addWidget(ei)
        el.addWidget(et)
        el.addWidget(eb, 0, Qt.AlignmentFlag.AlignCenter)
        v.addWidget(self.empty, 1)

        # 검증 화면 (ValidationPage 재사용)
        self.validation = ValidationPage()
        self.validation.hide()
        v.addWidget(self.validation, 1)

    def show_validation(self, data: SheetData):
        self.empty.hide()
        self.validation.set_data(data)
        self.validation.show()


# ═════════════════════════════════════════════════════════════════════════════
# [5] MainWindow — 조립 + 상태 관리
# ═════════════════════════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    TAB_UPLOAD, TAB_VALIDATION, TAB_ERRORS, TAB_UNITPRICE, TAB_UNITVALID = range(5)
    TAB_LABELS = ["⬆ 파일 업로드", "▦ 조견표 검증", "⚠ 오류 목록", "▤ 단가표 생성", "✔ 단가표 검증"]

    def __init__(self, backend: "BackendHooks"):
        super().__init__()
        self.backend = backend
        self.quick: Optional[SheetData] = None   # 조견표 검증 결과
        self.unit: Optional[SheetData] = None    # 단가표 검증 결과
        self.setWindowTitle("데이터 검증 자동화 시스템")
        self.resize(1280, 800)
        self.setStyleSheet(f"QMainWindow{{background:{T.BACKGROUND};}}")

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self._build_header())
        root.addWidget(self._build_tabbar())

        # 페이지 생성
        self.page_upload = UploadPage()
        self.page_validation = ValidationPage()
        self.page_errors = ErrorListPage()
        self.page_unitprice = UnitPricePage()
        self.page_unitvalid = UnitValidationPage()

        self.stack = QStackedWidget()
        for p in (self.page_upload, self.page_validation, self.page_errors,
                  self.page_unitprice, self.page_unitvalid):
            self.stack.addWidget(p)
        root.addWidget(self.stack, 1)
        root.addWidget(self._build_footer())

        self._connect_signals()
        self.go(self.TAB_UPLOAD)

    # ── 시그널 연결 ──────────────────────────────────────────
    def _connect_signals(self):
        self.page_upload.validateRequested.connect(self.run_validation)

        self.page_validation.fixRequested.connect(lambda e, v: self.fix_error(self.quick, e, v))
        self.page_validation.fixAllRequested.connect(lambda: self.fix_all(self.quick))
        self.page_validation.saveRequested.connect(lambda: self.save_sheet(self.quick))

        self.page_errors.fixRequested.connect(lambda e, v: self.fix_error(self.quick, e, v))
        self.page_errors.fixAllRequested.connect(lambda: self.fix_all(self.quick))
        self.page_errors.saveRequested.connect(lambda: self.save_sheet(self.quick))
        self.page_errors.rowActivated.connect(self._jump_to_error)

        self.page_unitprice.generateRequested.connect(self.run_generate)
        self.page_unitprice.downloadRequested.connect(self.download_unit_price)
        self.page_unitprice.validateRequested.connect(self.run_unit_validation)

        self.page_unitvalid.goGenerate.connect(lambda: self.go(self.TAB_UNITPRICE))
        self.page_unitvalid.validation.fixRequested.connect(lambda e, v: self.fix_error(self.unit, e, v))
        self.page_unitvalid.validation.fixAllRequested.connect(lambda: self.fix_all(self.unit))
        self.page_unitvalid.validation.saveRequested.connect(lambda: self.save_sheet(self.unit))

    # ── 헤더/탭바/푸터 ───────────────────────────────────────
    def _build_header(self) -> QWidget:
        head = QFrame()
        head.setFixedHeight(48)
        head.setStyleSheet(f"background:{T.PRIMARY};")
        l = QHBoxLayout(head)
        l.setContentsMargins(20, 0, 20, 0)
        l.setSpacing(10)
        logo = QLabel("▦")
        logo.setFixedSize(28, 28)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo.setStyleSheet("background:rgba(255,255,255,0.15); color:white; border-radius:4px; font-size:14px;")
        title = QLabel("데이터 검증 자동화 시스템")
        title.setStyleSheet("color:white; font-size:13px; font-weight:600;")
        l.addWidget(logo)
        l.addWidget(title)
        l.addStretch()
        self.hdr_err = QLabel()
        self.hdr_err.setStyleSheet("background:rgba(239,68,68,0.2); color:#fecaca;"
                                   " border:1px solid rgba(248,113,113,0.3); border-radius:4px;"
                                   " padding:3px 8px; font-size:11px;")
        self.hdr_fix = QLabel()
        self.hdr_fix.setStyleSheet("background:rgba(34,197,94,0.2); color:#bbf7d0;"
                                   " border:1px solid rgba(74,222,128,0.3); border-radius:4px;"
                                   " padding:3px 8px; font-size:11px;")
        self.hdr_err.hide()
        self.hdr_fix.hide()
        date = QLabel("🕒 " + QDate.currentDate().toString("yyyy. MM. dd."))
        date.setStyleSheet("color:rgba(255,255,255,0.5); font-size:11px;")
        l.addWidget(self.hdr_err)
        l.addWidget(self.hdr_fix)
        l.addWidget(date)
        return head

    def _build_tabbar(self) -> QWidget:
        bar = QFrame()
        bar.setStyleSheet(f"background:{T.CARD}; border-bottom:1px solid {T.BORDER};")
        l = QHBoxLayout(bar)
        l.setContentsMargins(20, 0, 20, 0)
        l.setSpacing(0)
        self.tabs: list[QPushButton] = []
        for i, txt in enumerate(self.TAB_LABELS):
            b = QPushButton(txt)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setFixedHeight(42)
            b.clicked.connect(lambda _, ix=i: self.go(ix))
            l.addWidget(b)
            self.tabs.append(b)
        l.addStretch()
        return bar

    def _build_footer(self) -> QWidget:
        foot = QFrame()
        foot.setFixedHeight(28)
        foot.setStyleSheet(f"background:{T.MUTED}; border-top:1px solid {T.BORDER};")
        l = QHBoxLayout(foot)
        l.setContentsMargins(20, 0, 20, 0)
        l.setSpacing(16)
        self.foot_status = QLabel("○ 대기 중")
        self.foot_status.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        self.foot_files = QLabel()
        self.foot_files.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        ver = QLabel("공공기관 데이터 검증 자동화 시스템 v1.0")
        ver.setStyleSheet(f"font-size:11px; color:{T.MUTED_FG};")
        l.addWidget(self.foot_status)
        l.addWidget(self.foot_files)
        l.addStretch()
        l.addWidget(ver)
        return foot

    def go(self, index: int):
        self.stack.setCurrentIndex(index)
        for i, b in enumerate(self.tabs):
            active = i == index
            if active:
                b.setStyleSheet(f"QPushButton{{background:transparent; color:{T.PRIMARY}; border:none;"
                                f" border-bottom:2px solid {T.PRIMARY}; font-size:12px; font-weight:600;"
                                f" padding:0 16px;}}")
            else:
                b.setStyleSheet(f"QPushButton{{background:transparent; color:{T.MUTED_FG}; border:none;"
                                f" border-bottom:2px solid transparent; font-size:12px; padding:0 16px;}}"
                                f"QPushButton:hover{{color:{T.FG};}}")

    # ── 동작 (백엔드 호출 → UI 반영) ──────────────────────────
    def run_validation(self):
        """탭 1: 검증 실행."""
        guide = self.page_upload.card_guide.file
        sheet = self.page_upload.card_sheet.file
        if not (guide and sheet):
            return
        self.page_upload.set_busy(True)
        try:
            # TODO: 오래 걸리는 작업이면 QThread/QThreadPool 로 감싸서 UI 프리징 방지
            self.quick = self.backend.validate_quick_reference(guide.path, sheet.path)
        except Exception as ex:
            QMessageBox.critical(self, "검증 실패", str(ex))
            return
        finally:
            self.page_upload.set_busy(False)
        self.page_validation.set_data(self.quick)
        self.page_errors.set_data(self.quick)
        self.refresh_status()
        self.go(self.TAB_VALIDATION)

    def fix_error(self, data: Optional[SheetData], err: ValidationError, value: str):
        """오류 1건 수정. value="" 이면 expected 로 자동 수정."""
        if not data:
            return
        err.fixed = True
        err.current = value or err.expected
        # TODO: 수정 내역을 백엔드에도 반영해야 하면 여기서 호출 (예: backend.apply_fix(err))
        self._refresh_views(data)

    def fix_all(self, data: Optional[SheetData]):
        if not data:
            return
        for e in data.open_errors:
            e.fixed = True
            e.current = e.expected
        # TODO: 일괄 수정 백엔드 반영이 필요하면 여기서 호출
        self._refresh_views(data)

    def save_sheet(self, data: Optional[SheetData]):
        if not data:
            return
        path, _ = QFileDialog.getSaveFileName(self, "파일 저장", data.filename, "Excel (*.xlsx)")
        if not path:
            return
        try:
            self.backend.save_fixed_sheet(data, path)
            QMessageBox.information(self, "저장 완료", f"파일이 저장되었습니다.\n{path}")
        except Exception as ex:
            QMessageBox.critical(self, "저장 실패", str(ex))

    def run_generate(self):
        """탭 4: 단가표 생성."""
        self.page_unitprice.set_busy("generate", True)
        try:
            # TODO: 오래 걸리면 QThread 처리
            headers, rows, filename, size_text = self.backend.generate_unit_price(self.quick)
        except Exception as ex:
            QMessageBox.critical(self, "생성 실패", str(ex))
            return
        finally:
            self.page_unitprice.set_busy("generate", False)
        self.page_unitprice.set_preview(headers, rows)
        self.page_unitprice.set_generated(filename, size_text)

    def download_unit_price(self):
        path, _ = QFileDialog.getSaveFileName(self, "다운로드", "unit_price_list.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            self.backend.export_unit_price(path)
            QMessageBox.information(self, "저장 완료", f"파일이 저장되었습니다.\n{path}")
        except Exception as ex:
            QMessageBox.critical(self, "저장 실패", str(ex))

    def run_unit_validation(self):
        """탭 4 → 5: 단가표 검증."""
        self.page_unitprice.set_busy("validate", True)
        try:
            # TODO: 오래 걸리면 QThread 처리
            self.unit = self.backend.validate_unit_price()
        except Exception as ex:
            QMessageBox.critical(self, "검증 실패", str(ex))
            return
        finally:
            self.page_unitprice.set_busy("validate", False)
        self.page_unitvalid.show_validation(self.unit)
        self.go(self.TAB_UNITVALID)

    # ── 화면 갱신 ────────────────────────────────────────────
    def _refresh_views(self, data: SheetData):
        if data is self.quick:
            self.page_validation.refresh()
            self.page_errors.refresh()
        elif data is self.unit:
            self.page_unitvalid.validation.refresh()
        self.refresh_status()

    def _jump_to_error(self, err: ValidationError):
        self.go(self.TAB_VALIDATION)
        self.page_validation.open_detail(err)

    def refresh_status(self):
        """헤더 배지 / 푸터 / 단가표 페이지 상태 동기화."""
        if not self.quick:
            return
        n_open, n_fixed = len(self.quick.open_errors), len(self.quick.fixed_errors)
        self.hdr_err.setVisible(n_open > 0)
        self.hdr_err.setText(f"⨯ 오류 {n_open}건")
        self.hdr_fix.setVisible(n_fixed > 0)
        self.hdr_fix.setText(f"✔ 수정 {n_fixed}건")
        self.foot_status.setText("● 검증 완료")
        self.foot_status.setStyleSheet(f"font-size:11px; color:{T.GREEN};")
        files = []
        if self.page_upload.card_guide.file:
            files.append("📄 " + self.page_upload.card_guide.file.name)
        if self.page_upload.card_sheet.file:
            files.append("▦ " + self.page_upload.card_sheet.file.name)
        self.foot_files.setText("    ".join(files))
        self.page_unitprice.set_ready(True, len(self.quick.errors), n_fixed)


# ═════════════════════════════════════════════════════════════════════════════
# [6] BackendHooks — ★ 백엔드 로직 연결 지점 (여기만 채우면 됨) ★
# ═════════════════════════════════════════════════════════════════════════════
class BackendHooks:
    """
    구현된 백엔드 로직을 이 클래스의 메서드에 연결하세요.
    각 메서드는 UI가 그대로 소비할 수 있는 형태(SheetData / ValidationError)를 반환합니다.
    실패 시 예외를 던지면 UI가 에러 다이얼로그를 띄웁니다.
    """

    def validate_quick_reference(self, guide_pdf_path: str, sheet_xlsx_path: str) -> SheetData:
        """사업지침 + 조견표 → 검증 결과."""
        # TODO: 백엔드 검증 로직 호출 후 결과를 SheetData 로 변환하여 반환
        #   return SheetData(
        #       filename=os.path.basename(sheet_xlsx_path),
        #       sheet_name="...",
        #       headers=[...],                # 표 헤더
        #       rows=[[...], ...],            # 표 데이터 (문자열 2차원 리스트)
        #       errors=[ValidationError(...), ...],
        #   )
        raise NotImplementedError("validate_quick_reference 를 구현하세요.")

    def save_fixed_sheet(self, data: SheetData, save_path: str) -> None:
        """수정 완료된 시트를 엑셀로 저장."""
        # TODO: data.rows + data.errors(수정값 반영됨) 를 엑셀로 기록
        raise NotImplementedError("save_fixed_sheet 를 구현하세요.")

    def generate_unit_price(self, quick: Optional[SheetData]) -> tuple[list[str], list[list[str]], str, str]:
        """단가표 생성. 반환: (headers, rows, 파일명, 크기 텍스트)."""
        # TODO: 검증된 조견표 기반 단가표 생성 로직 호출
        raise NotImplementedError("generate_unit_price 를 구현하세요.")

    def export_unit_price(self, save_path: str) -> None:
        """생성된 단가표 엑셀 파일 내보내기."""
        # TODO: 생성 결과를 save_path 에 기록
        raise NotImplementedError("export_unit_price 를 구현하세요.")

    def validate_unit_price(self) -> SheetData:
        """생성된 단가표 검증 결과."""
        # TODO: 단가표 검증 로직 호출 후 SheetData 반환
        raise NotImplementedError("validate_unit_price 를 구현하세요.")


# ═════════════════════════════════════════════════════════════════════════════
# 엔트리 포인트
# ═════════════════════════════════════════════════════════════════════════════
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Malgun Gothic", 9))   # TODO: 배포 환경에 맞는 한글 폰트로 교체 가능
    win = MainWindow(backend=BackendHooks())
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
