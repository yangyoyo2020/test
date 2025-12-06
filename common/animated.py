from PyQt6.QtWidgets import QPushButton, QGraphicsDropShadowEffect
from PyQt6.QtCore import QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QColor


class AnimatedButton(QPushButton):
    """A QPushButton with a subtle hover elevation implemented via QGraphicsDropShadowEffect.

    On hover, the shadow blur is animated up; on leave it animates back down.
    """

    def __init__(self, *args, shadow_color=QColor(0, 0, 0, 80), hover_blur=18, duration=180, **kwargs):
        super().__init__(*args, **kwargs)
        self._effect = QGraphicsDropShadowEffect(self)
        self._effect.setOffset(0, 2)
        self._effect.setBlurRadius(0)
        self._effect.setColor(shadow_color)
        self.setGraphicsEffect(self._effect)

        self._hover_blur = hover_blur
        self._duration = duration

        self._anim = QPropertyAnimation(self._effect, b"blurRadius", self)
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._anim.setDuration(self._duration)

    def enterEvent(self, event):
        try:
            self._anim.stop()
            self._anim.setStartValue(self._effect.blurRadius())
            self._anim.setEndValue(self._hover_blur)
            self._anim.start()
        except Exception:
            pass
        super().enterEvent(event)

    def leaveEvent(self, event):
        try:
            self._anim.stop()
            self._anim.setStartValue(self._effect.blurRadius())
            self._anim.setEndValue(0)
            self._anim.start()
        except Exception:
            pass
        super().leaveEvent(event)
