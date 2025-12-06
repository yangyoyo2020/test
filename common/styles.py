from pathlib import Path
from typing import Optional, Dict
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QSettings


_CURRENT_THEME: str = 'light'


# 视觉变量：在此集中管理浅色/深色主题的关键配色与基础字体
THEME_VARS: Dict[str, Dict[str, str]] = {
    'light': {
        'PRIMARY': '#1976D2',
        'ACCENT': '#FF8C42',
        'SUCCESS': '#4CAF50',
        'WARNING': '#FF9800',
        'DANGER': '#f44336',
        'BG': '#F3F4F6',
        'SURFACE': '#FFFFFF',
        'SURFACE_ALT': '#FBFDFF',
        'HEADER_BG': '#F7F9FB',
        'TEXT': '#222222',
        'TITLE': '#2c3e50',
        'MUTED': '#7f8c8d',
        'BORDER_LIGHT': '#E6E9EE',
        'BORDER_MUTED': '#D8DDE6',
        'GRIDLINE': '#F0F0F0',
        'PLACEHOLDER_BG': '#ecf0f1',
        'SUCCESS_BG': '#d5f5e3',
        'DISABLED_BG': '#E0E0E0',
        'DISABLED_TEXT': '#9E9E9E',
    },
    'dark': {
        'PRIMARY': '#2563EB',
        'ACCENT': '#FFB775',
        'SUCCESS': '#16A34A',
        'WARNING': '#D97706',
        'DANGER': '#DC2626',
        'BG': '#121214',
        'SURFACE': '#1E1F22',
        'SURFACE_ALT': '#0f1112',
        'HEADER_BG': '#151617',
        'TEXT': '#E6E6E6',
        'TITLE': '#FFFFFF',
        'MUTED': '#9CA3AF',
        'BORDER_DARK': '#2A2A2A',
        'BORDER_MUTED': '#2C2C2E',
        'GRIDLINE': '#1f1f1f',
        'PLACEHOLDER_BG': '#0f1112',
        'SUCCESS_BG': 'rgba(52,211,153,0.06)',
        'DISABLED_BG': '#2b2b2b',
        'DISABLED_TEXT': '#6b6b6b',
    }
}


# 一个精简的基础 QSS 模板，使用 theme vars 填充，随后会把项目中的完整 qss（若存在）追加
_BASE_TEMPLATE = """
QWidget {{
  background-color: {BG};
  color: {TEXT};
  font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
  font-size: 10pt;
}}

QPushButton {{
  background-color: {PRIMARY};
  color: {SURFACE};
  border: none;
  padding: 8px 14px;
  border-radius: 6px;
  font-weight: 600;
}}
QPushButton:hover {{ background-color: {PRIMARY}; opacity: 0.95; }}

QPushButton[variant="secondary"] {{ background-color: {SUCCESS}; }}
QPushButton[variant="warning"] {{ background-color: {WARNING}; }}
QPushButton[variant="danger"] {{ background-color: {DANGER}; }}

QLineEdit, QTextEdit {{
  background: {SURFACE};
  border: 1px solid rgba(0,0,0,0.08);
  border-radius: 6px;
  padding: 6px;
}}

.QCard {{
  background-color: {SURFACE};
  border-radius: 8px;
  padding: 12px;
  border: 1px solid rgba(0,0,0,0.06);
}}
"""


def load_qss(path: Path) -> str:
    try:
        return Path(path).read_text(encoding='utf-8')
    except Exception:
        return ''


def render_base_qss(theme: str = 'light') -> str:
    """Render the minimal base QSS from `THEME_VARS` for the given theme."""
    vars = THEME_VARS.get(theme, THEME_VARS['light'])
    return _BASE_TEMPLATE.format(**vars)


def apply_theme(app: Optional[QApplication], theme: Optional[str] = None) -> None:
    """Apply theme to the application.

    Behavior:
    - If `theme` is None, try to read saved preference from QSettings (key 'theme').
    - Build a base QSS from visual variables and append project qss file if present.
    - Update module-level `_CURRENT_THEME` and persist selection.
    """
    global _CURRENT_THEME
    if app is None:
        app = QApplication.instance()
    # 读取用户偏好
    settings = QSettings('yangyoyo2020', 'test')
    saved = settings.value('theme', type=str)
    if theme is None:
        theme = saved or _CURRENT_THEME

    if theme not in THEME_VARS:
        theme = 'light'

    base_dir = Path(__file__).parent

    # 先渲染基础 qss（保证关键颜色来自中心化变量）
    qss = render_base_qss(theme)

    # 如果存在文件版的 qss（legacy），把它追加到渲染结果后面，让文件中的更具体规则生效
    file_qss = None
    if theme == 'dark':
        candidate = base_dir / 'style_dark.qss'
        if candidate.exists():
            file_qss = load_qss(candidate)
    else:
        candidate = base_dir / 'style.qss'
        if candidate.exists():
            file_qss = load_qss(candidate)

    if file_qss:
        # 替换文件 qss 中的 %KEY% 占位符为当前主题色值（向后兼容占位用法）
        theme_vars = THEME_VARS.get(theme, THEME_VARS['light'])
        for k, v in theme_vars.items():
            file_qss = file_qss.replace(f"%{k}%", v)

        # 把文件 qss 放到渲染的基础后面，允许细节覆盖
        qss = qss + '\n\n' + file_qss

    try:
        if app is not None:
            app.setStyleSheet(qss)
            _CURRENT_THEME = theme
            settings.setValue('theme', theme)
    except Exception:
        # 如果设置样式失败，保持当前主题不变
        pass


def get_current_theme() -> str:
    """Return the current theme name ('light' or 'dark')."""
    return _CURRENT_THEME


def set_theme(app: Optional[QApplication], theme: str) -> None:
    """Set and apply a specific theme (and persist it)."""
    apply_theme(app, theme=theme)


def toggle_theme(app: Optional[QApplication] = None) -> str:
    """Toggle between light and dark themes. Returns the new theme name.

    If `app` is None, the function will try to use QApplication.instance().
    """
    global _CURRENT_THEME
    if app is None:
        app = QApplication.instance()
    new_theme = 'dark' if _CURRENT_THEME == 'light' else 'light'
    apply_theme(app, theme=new_theme)
    return get_current_theme()
