def open_or_activate(parent, attr_name: str, factory_callable, restore_if_minimized: bool = True):
    """通用函数：复用已打开的窗口或创建并显示新的窗口实例。

    - parent: 拥有窗口引用属性的对象（例如 UnifiedLoginWindow 实例）
    - attr_name: parent 上保存窗口实例的属性名（字符串）
    - factory_callable: 无参函数，返回新创建的窗口实例
    - restore_if_minimized: 若为 True，且找到的已存在窗口处于最小化状态，则调用 `showNormal()` 恢复

    返回：窗口实例（复用或新创建）
    """
    try:
        existing = getattr(parent, attr_name, None)
        if existing:
            try:
                # 如果窗口可见或处于最小化（需要恢复），优先复用
                is_visible = getattr(existing, 'isVisible', lambda: False)()
                is_minimized = getattr(existing, 'isMinimized', lambda: False)()
                if is_visible or is_minimized:
                    try:
                        # 若最小化则根据参数恢复到正常状态
                        if restore_if_minimized and is_minimized:
                            try:
                                existing.showNormal()
                            except Exception:
                                pass
                        existing.activateWindow()
                        existing.raise_()
                    except Exception:
                        pass
                    return existing
            except Exception:
                pass
    except Exception:
        pass

    # 未找到可复用窗口，创建新的
    inst = factory_callable()
    try:
        setattr(parent, attr_name, inst)
    except Exception:
        pass
    try:
        inst.show()
    except Exception:
        pass
    return inst
