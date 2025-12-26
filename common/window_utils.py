from PyQt6.QtCore import QObject
from common.logger import get_logger

# 获取全局日志器用于追踪窗口生命周期
_window_logger = None

def _get_window_logger():
    global _window_logger
    if _window_logger is None:
        try:
            _window_logger = get_logger('WindowUtils')
        except Exception:
            _window_logger = None
    return _window_logger


def open_or_activate(parent, attr_name: str, factory_callable, restore_if_minimized: bool = True):
    """通用函数：复用已打开的窗口或创建并显示新的窗口实例。

    本函数增强了对已销毁/已删除 Qt 对象的检测：如果父对象上保存的引用指向已删除的 Qt 对象，
    会自动清理并创建新的实例。创建新实例后会连接其 `destroyed` 信号以在窗口销毁时清理父属性，
    避免出现悬挂引用导致后续复用时抛出异常或程序行为异常（例如意外退出）。
    """
    logger = _get_window_logger()
    try:
        existing = getattr(parent, attr_name, None)
        if existing:
            if logger:
                logger.info(f"[WindowUtils] 尝试复用窗口: {attr_name}")
            # 检查 existing 是否为有效的 QObject（未被删除）
            try:
                if isinstance(existing, QObject):
                    # 调用 isVisible/isMinimized 可能因已删除的底层 C++ 对象而抛出，故放在 try 中
                    try:
                        is_visible = existing.isVisible()
                        is_minimized = existing.isMinimized()
                        if logger:
                            logger.info(f"[WindowUtils] 窗口 {attr_name} 状态 - 可见: {is_visible}, 最小化: {is_minimized}")
                    except Exception as e:
                        # 已被删除或不可用，清理引用，强制创建新实例
                        if logger:
                            logger.warning(f"[WindowUtils] 窗口 {attr_name} 已被删除或不可用: {e}")
                        try:
                            setattr(parent, attr_name, None)
                        except Exception:
                            pass
                        existing = None
                    else:
                        if is_visible or is_minimized:
                            try:
                                if restore_if_minimized and is_minimized:
                                    try:
                                        existing.showNormal()
                                    except Exception:
                                        pass
                                existing.activateWindow()
                                existing.raise_()
                                if logger:
                                    logger.info(f"[WindowUtils] 成功复用窗口: {attr_name}")
                            except Exception:
                                pass
                            return existing
                else:
                    # 非 QObject 对象，尝试复用（保守处理）
                    try:
                        is_visible = getattr(existing, 'isVisible', lambda: False)()
                        is_minimized = getattr(existing, 'isMinimized', lambda: False)()
                        if is_visible or is_minimized:
                            try:
                                if restore_if_minimized and is_minimized:
                                    try:
                                        existing.showNormal()
                                    except Exception:
                                        pass
                                existing.activateWindow()
                                existing.raise_()
                                if logger:
                                    logger.info(f"[WindowUtils] 成功复用非QObject窗口: {attr_name}")
                            except Exception:
                                pass
                            return existing
                    except Exception:
                        # 如果检测失败，放弃复用
                        if logger:
                            logger.warning(f"[WindowUtils] 复用窗口 {attr_name} 失败，创建新实例")
                        try:
                            setattr(parent, attr_name, None)
                        except Exception:
                            pass
                        existing = None
            except Exception as e:
                # 任何异常都触发创建新实例
                if logger:
                    logger.error(f"[WindowUtils] 复用窗口 {attr_name} 异常: {e}")
                try:
                    setattr(parent, attr_name, None)
                except Exception:
                    pass
                existing = None
    except Exception as e:
        if logger:
            logger.error(f"[WindowUtils] 检查现有窗口异常: {e}")
        existing = None

    # 未找到可复用窗口或引用不可用，创建新的
    if logger:
        logger.info(f"[WindowUtils] 创建新窗口实例: {attr_name}")
    inst = factory_callable()
    try:
        # 将实例保存到父对象上
        try:
            setattr(parent, attr_name, inst)
            if logger:
                logger.info(f"[WindowUtils] 窗口实例已保存到 {attr_name}")
        except Exception as e:
            if logger:
                logger.error(f"[WindowUtils] 保存窗口实例失败: {e}")

        # 注意：不再连接 destroyed 信号，因为：
        # 1. closeEvent 已经清理了资源
        # 2. destroyed 信号在二次事件循环中可能导致不稳定
        # 3. 父窗口持有引用，子窗口销毁时自动清理

        try:
            inst.show()
            if logger:
                logger.info(f"[WindowUtils] 窗口已显示: {attr_name}")
        except Exception as e:
            if logger:
                logger.error(f"[WindowUtils] 显示窗口失败: {e}")
    except Exception as e:
        # 创建或显示过程中出错时，尽量返回实例以便呼叫者有机会处理
        if logger:
            logger.error(f"[WindowUtils] 创建窗口过程异常: {e}")
        return inst

    return inst
