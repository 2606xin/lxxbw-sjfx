# -*- coding: utf-8 -*-
"""paths.py — 兼容开发与 PyInstaller 冻结模式的资源路径"""
import os, sys


def resource_path(rel):
    """返回相对资源路径的绝对路径 (冻结时指向 _MEIPASS)"""
    base = getattr(sys, "_MEIPASS", None)
    if base is None:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel)
