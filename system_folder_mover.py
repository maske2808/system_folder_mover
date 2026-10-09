#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
system_folder_mover.py - 系统文件夹转移工具（Marvis 风格 UI）
功能：
  - 列出 Windows 常见系统文件夹（桌面 / 文档 / 下载 / 图片 / 音乐 / 视频 /
    收藏夹 / 联系人 / 搜索 / 发送到 / 3D对象 / 相机胶卷 / 开始菜单 / 启动 /
    最近 / 保存的游戏 等 Known Folders），显示当前路径与实时大小（MB / GB）
  - 可选「转移文件」或「不转移文件」：转移文件=robocopy 移动文件并更新文件夹
    指向；不转移文件=仅更新文件夹指向，文件保留原位置
  - 批量目标文件夹：选择根目录后自动为每个系统文件夹生成同名子目录；也可为
    每个系统文件夹单独指定目标文件夹
  - 配置保存 / 导入（INI）：保存当前所有文件夹的当前位置为 .ini 快照
    （[Folders] 节 名称=路径），导入后各项目标设为文件中的位置
  - 可选「允许执行注销」：以 ExitWindowsEx 注销当前账户使转移生效
  - 可选「清理垃圾」：清理用户临时文件、缩略图缓存与回收站
  - 顶部横幅显示当前系统名称与系统版本（注册表 ProductName / DisplayVersion /
    CurrentBuildNumber，Build>=22000 判定为 Windows 11）
UI 仿 Marvis：左侧导航（程序 + 底部设置）+ 右侧内容框，浅色主题，自绘无边框
标题栏 + 圆角窗口 + 右侧可收起日志栏。

运行：建议以管理员权限运行（SHSetKnownFolderPath 修改系统文件夹指向需要
管理员权限；普通权限下部分文件夹可能失败）。
依赖：Python 标准库 + ctypes（无第三方依赖）。
"""

import ctypes
import ctypes.wintypes as wt
import os
import platform
import queue
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import uuid
import winreg

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------
SINGLE_INSTANCE_MUTEX = r'Global\SystemFolderMover_SingleInstance'
WINDOW_TITLE = '系统文件夹转移 - 系统文件夹转移工具'

APP_NAME = 'FolderMover'

LOG_PANEL_W = 300

# 系统文件夹清单：名称 -> Known Folder GUID + 默认子目录名 + 是否常用 + 回退路径模板
# 回退路径用于 SHGetKnownFolderPath / 注册表均无法定位时的兜底展示（%XXX% 环境变量）
KNOWN_FOLDERS = [
    ('桌面',      '{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}', 'Desktop',      True,  '%USERPROFILE%\\Desktop'),
    ('我的文档',  '{FDD39AD0-238F-46AF-ADB4-6C85480369C7}', 'Documents',    True,  '%USERPROFILE%\\Documents'),
    ('下载',      '{374DE290-123F-4565-9164-39C4925E467B}', 'Downloads',    True,  '%USERPROFILE%\\Downloads'),
    ('图片',      '{33E28130-4E1E-4676-835A-98395C3BC3BB}', 'Pictures',     True,  '%USERPROFILE%\\Pictures'),
    ('音乐',      '{4BD8D9C6-6DC0-4A00-BE55-E34E09DCFA62}', 'Music',        True,  '%USERPROFILE%\\Music'),
    ('视频',      '{18989B1D-99B5-455B-841C-AB7C74E4DDFC}', 'Videos',       True,  '%USERPROFILE%\\Videos'),
    ('收藏夹',    '{1777F761-68AD-4D8A-87BD-30B759FA33DD}', 'Favorites',    False, '%USERPROFILE%\\Favorites'),
    ('联系人',    '{56784854-C6CB-462B-8169-88E350ACB882}', 'Contacts',     False, '%USERPROFILE%\\Contacts'),
    ('搜索',      '{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}', 'Searches',     False, '%USERPROFILE%\\Searches'),
    ('发送到',    '{8983036C-9480-4F82-B6EA-63F5B8CA2AB6}', 'SendTo',       False, '%APPDATA%\\Microsoft\\Windows\\SendTo'),
    ('3D对象',    '{31C0DD25-9439-4F12-BF41-7FFAEA741455}', '3D Objects',   False, '%USERPROFILE%\\3D Objects'),
    ('相机胶卷',  '{AB5FB87B-7CE2-4F83-915D-550846C9537B}', 'Camera Roll',  False, '%USERPROFILE%\\Pictures\\Camera Roll'),
    ('开始菜单',  '{625B52C3-B724-4A1E-9F72-FC0D2D3E0F9A}', 'Start Menu',   False, '%APPDATA%\\Microsoft\\Windows\\Start Menu'),
    ('启动',      '{B97D20BB-F46A-4C97-9B0B-221ED071251D}', 'Startup',      False, '%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Startup'),
    ('最近',      '{AE50C081-EBD2-438A-8655-8A092E34987A}', 'Recent',       False, '%APPDATA%\\Microsoft\\Windows\\Recent'),
    ('保存的游戏', '{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}', 'Saved Games', False, '%USERPROFILE%\\Saved Games'),
    # 第 8 次迭代新增：次要页 4 项
    # 临时文件无 Known Folder GUID（fid 为空串），直接以 %TEMP% 环境变量定位，
    # 转移时仅移动文件、不更新 Known Folder 指向
    ('IE缓存',    '{3524812D-4B42-4D67-8A20-A37C670D0B7A}', 'INetCache',     False, '%LOCALAPPDATA%\\Microsoft\\Windows\\INetCache'),
    ('Cookies',   '{2B0CA765-2385-4F45-AB1C-F8C0B3F95B4D}', 'Cookies',       False, '%APPDATA%\\Microsoft\\Windows\\Cookies'),
    ('临时文件',  '',                                     'Temp',          False, '%TEMP%'),
    ('链接',      '{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}', 'Links',       False, '%USERPROFILE%\\Links'),
]

# 第 8 次迭代：Known Folder 指向的目录不存在时，按顺序尝试的常见实际路径
# （仅对目录不存在的项生效，存在则优先保留系统配置的指向）
_ALTERNATE_PATHS = {
    'IE缓存': [
        '%LOCALAPPDATA%\\Microsoft\\Windows\\INetCache',
        '%LOCALAPPDATA%\\Microsoft\\Windows\\Temporary Internet Files',
    ],
    'Cookies': [
        '%APPDATA%\\Microsoft\\Windows\\Cookies',
        '%LOCALAPPDATA%\\Microsoft\\Windows\\Cookies',
    ],
}

# 第 9 次迭代：User Shell Folders 注册表值名映射（参考 pft.exe 的定位方式）。
# pft.exe 直接读取注册表友好值名（Cache/Cookies/My Music/Personal 等）得到
# 系统文件夹真实当前位置；本机 SHGetKnownFolderPath 在部分环境返回失败，
# 注册表回退若仅按 Known Folder GUID 查询会漏掉友好值名，导致落到环境变量
# 兜底。此处给出每个系统文件夹在注册表中的候选值名（友好名优先，GUID 次之），
# 与 pft.exe 的读取逻辑保持一致。
_SHELL_FOLDER_VALUE_NAMES = {
    '桌面':     ['Desktop', '{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}'],
    '我的文档': ['Personal', '{FDD39AD0-238F-46AF-ADB4-6C85480369C7}'],
    '下载':     ['{374DE290-123F-4565-9164-39C4925E467B}'],
    '图片':     ['My Pictures', '{33E28130-4E1E-4676-835A-98395C3BC3BB}'],
    '音乐':     ['My Music', '{4BD8D9C6-6DC0-4A00-BE55-E34E09DCFA62}'],
    '视频':     ['My Video', '{18989B1D-99B5-455B-841C-AB7C74E4DDFC}'],
    '收藏夹':   ['Favorites', '{1777F761-68AD-4D8A-87BD-30B759FA33DD}'],
    '联系人':   ['Contacts', '{56784854-C6CB-462B-8169-88E350ACB882}'],
    '搜索':     ['Searches', '{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}'],
    '发送到':   ['SendTo'],
    '3D对象':   ['{31C0DD25-9439-4F12-BF41-7FFAEA741455}'],
    '相机胶卷': ['{AB5FB87B-7CE2-4F83-915D-550846C9537B}'],
    '开始菜单': ['Start Menu'],
    '启动':     ['Startup'],
    '最近':     ['Recent'],
    '保存的游戏': ['Saved Games', '{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}'],
    'IE缓存':   ['Cache', '{3524812D-4B42-4D67-8A20-A37C670D0B7A}'],
    'Cookies':  ['Cookies', '{2B0CA765-2385-4F45-AB1C-F8C0B3F95B4D}'],
    '临时文件': [],
    '链接':     ['Links', '{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}'],
}

# 第 10 次迭代：批量目标重构 —— 未单独设置目标的文件夹跟随「用户（个人文件夹）」。
# 跟随项目标 = 用户目标路径 + 该文件夹相对用户目录的子路径。
# 映射参考 D:\xitong-shezhi\用户文件转移到其他盘\个人文件转移工具(DOS之家).ini；
# 配置中无映射的按 Windows 默认位置相对用户目录的路径填写。
_FOLLOW_SUBDIRS = {
    '桌面':     'Desktop',
    '我的文档': 'Documents',
    '下载':     'Downloads',
    '图片':     'Pictures',
    '音乐':     'Music',
    '视频':     'Videos',
    '收藏夹':   'Favorites',
    '联系人':   'Contacts',
    '搜索':     'Searches',
    '发送到':   'AppData\\Roaming\\Microsoft\\Windows\\SendTo',
    '3D对象':   '3D Objects',
    '相机胶卷': 'Pictures\\Camera Roll',
    '开始菜单': 'AppData\\Roaming\\Microsoft\\Windows\\Start Menu',
    '启动':     'AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup',
    '最近':     'AppData\\Roaming\\Microsoft\\Windows\\Recent',
    '保存的游戏': 'Saved Games',
    'IE缓存':   'INetCache',
    'Cookies':  'INetCookies',
    '临时文件': 'Temp',
    '链接':     'Links',
}


def _round_rect_pts(x0, y0, x1, y1, r):
    """圆角矩形点列（配合 Canvas create_polygon smooth=True 绘制大圆角矩形）。"""
    r = max(1, min(float(r), (x1 - x0) / 2.0, (y1 - y0) / 2.0))
    return [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r,
            x1, y1 - r, x1, y1, x1 - r, y1, x0 + r, y1,
            x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]


# ---------------------------------------------------------------------------
# 单实例互斥
# ---------------------------------------------------------------------------
def _acquire_single_instance():
    k = ctypes.WinDLL('kernel32', use_last_error=True)
    k.CreateMutexW.restype = wt.HANDLE
    k.CreateMutexW.argtypes = [ctypes.c_void_p, wt.BOOL, wt.LPCWSTR]
    k.WaitForSingleObject.argtypes = [wt.HANDLE, ctypes.c_uint32]
    k.WaitForSingleObject.restype = ctypes.c_uint32
    h = k.CreateMutexW(None, False, SINGLE_INSTANCE_MUTEX)
    if not h:
        return True
    if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
        # 旧实例已异常退出时 mutex 处于 abandoned：可接管
        if k.WaitForSingleObject(h, 0) == 0x80:  # WAIT_ABANDONED
            return True
        try:
            k.CloseHandle(h)
        except Exception:
            pass
        return False
    return True


def _activate_existing_window():
    """查找已运行实例的主窗口并激活（还原 + 置前 + 前台）。"""
    try:
        u = ctypes.windll.user32
        u.FindWindowW.restype = wt.HWND
        u.FindWindowW.argtypes = [wt.LPCWSTR, wt.LPCWSTR]
        u.EnumWindows.restype = wt.BOOL
        u.EnumWindows.argtypes = [ctypes.c_void_p, wt.LPARAM]
        u.GetWindowTextLengthW.restype = ctypes.c_int
        u.GetWindowTextLengthW.argtypes = [wt.HWND]
        u.GetWindowTextW.restype = ctypes.c_int
        u.GetWindowTextW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
        u.IsWindow.restype = wt.BOOL
        u.IsWindow.argtypes = [wt.HWND]
        u.ShowWindow.argtypes = [wt.HWND, ctypes.c_int]
        u.SetForegroundWindow.argtypes = [wt.HWND]
        u.SetWindowPos.argtypes = [wt.HWND, wt.HWND, ctypes.c_int,
                                   ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                   ctypes.c_uint]
        u.FlashWindow.argtypes = [wt.HWND, wt.BOOL]
        u.keybd_event.argtypes = [wt.BYTE, wt.BYTE, wt.DWORD, ctypes.c_size_t]

        def _title_ok(hwnd):
            n = u.GetWindowTextLengthW(hwnd)
            if n <= 0:
                return True
            buf = ctypes.create_unicode_buffer(n + 1)
            u.GetWindowTextW(hwnd, buf, n + 1)
            return buf.value != WINDOW_TITLE

        _found = [0]

        @ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
        def _cb(hwnd, lparam):
            if not _title_ok(hwnd):
                _found[0] = hwnd
                return False
            return True

        hwnd = u.FindWindowW(None, WINDOW_TITLE)
        if not (hwnd and u.IsWindow(hwnd)):
            u.EnumWindows(_cb, 0)
            hwnd = _found[0]
        if hwnd and u.IsWindow(hwnd):
            u.ShowWindow(hwnd, 9)
            u.SetWindowPos(hwnd, -1, 0, 0, 0, 0,
                           0x0001 | 0x0002 | 0x0040)
            u.SetWindowPos(hwnd, -2, 0, 0, 0, 0,
                           0x0001 | 0x0002)
            u.keybd_event(0x12, 0, 0, 0)
            u.keybd_event(0x12, 0, 2, 0)
            u.SetForegroundWindow(hwnd)
            u.FlashWindow(hwnd, True)
            return True
    except Exception:
        pass
    return False


# ---------------------------------------------------------------------------
# Known Folder 路径读写（SHGetKnownFolderPath / SHSetKnownFolderPath）
# ---------------------------------------------------------------------------
class _GUID(ctypes.Structure):
    """GUID 结构（ctypes.wintypes 不含 GUID，需自定义）。"""
    _fields_ = [('Data1', ctypes.c_ulong),
                ('Data2', ctypes.c_ushort),
                ('Data3', ctypes.c_ushort),
                ('Data4', ctypes.c_ubyte * 8)]


def _make_guid(fid):
    """把 '{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}' 字符串转为 GUID 结构。"""
    u = uuid.UUID(fid)
    g = _GUID()
    g.Data1 = u.time_low
    g.Data2 = u.time_mid
    g.Data3 = u.time_hi_version
    g.Data4 = (ctypes.c_ubyte * 8)(*u.bytes[8:16])
    return g


def _expand_env(path):
    """展开路径中的 %XXX% 环境变量。"""
    try:
        return os.path.expandvars(path)
    except Exception:
        return path


def _get_known_folder_path(fid, fallback=None, reg_names=None):
    """读取 Known Folder 当前路径。
    优先级：SHGetKnownFolderPath -> 注册表 User Shell Folders（fid GUID 与
    友好值名，参考 pft.exe 按 Cache/Cookies/My Music/Personal 等值名读取）
    -> 回退模板。"""
    try:
        shell32 = ctypes.windll.shell32
        ole32 = ctypes.windll.ole32
        g = _make_guid(fid)
        p = ctypes.c_wchar_p()
        hr = shell32.SHGetKnownFolderPath(ctypes.byref(g), 0, None,
                                          ctypes.byref(p))
        try:
            if hr == 0 and p.value:
                return p.value
        finally:
            if p:
                ole32.CoTaskMemFree(p)
    except Exception:
        pass
    # 回退：注册表 User Shell Folders（fid GUID + 友好值名，pft.exe 同款）
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders',
            0, winreg.KEY_READ)
        try:
            for nm in ((fid, ) + tuple(reg_names or ())):
                try:
                    v, _ = winreg.QueryValueEx(key, nm)
                    if v:
                        return _expand_env(v)
                except OSError:
                    continue
        finally:
            winreg.CloseKey(key)
    except OSError:
        pass
    # 回退：环境变量模板
    if fallback:
        v = _expand_env(fallback)
        if v:
            return v
    return None


def _get_user_profile_path():
    """用户（个人文件夹）当前位置（参考 pft.exe 的定位方式）。
    pft.exe 会尝试读取注册表 User Shell Folders 的 'UserProfile' 值名；
    该值在 Windows 默认不存在，个人文件夹实际被转移后，其子项（我的文档
    Personal）的父目录即个人文件夹根（如 F:\\yonghu）。定位优先级：
    1) 注册表 'UserProfile' 值名；2) Personal（我的文档）的父目录（存在且
    非系统用户目录根时）；3) %USERPROFILE% 兜底。"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders',
            0, winreg.KEY_READ)
        try:
            try:
                v, _ = winreg.QueryValueEx(key, 'UserProfile')
                if v:
                    return _expand_env(v)
            except OSError:
                pass
            try:
                v, _ = winreg.QueryValueEx(key, 'Personal')
                if v:
                    parent = os.path.dirname(_expand_env(v))
                    low = parent.lower()
                    if (parent and os.path.isdir(parent)
                            and low not in ('c:\\users', 'c:\\', 'c:\\users\\')):
                        return parent
            except OSError:
                pass
        finally:
            winreg.CloseKey(key)
    except OSError:
        pass
    return _expand_env('%USERPROFILE%') or None


def _set_known_folder_path(fid, new_path):
    """更新 Known Folder 指向。返回 (ok, msg)。"""
    try:
        shell32 = ctypes.windll.shell32
        g = _make_guid(fid)
        hr = shell32.SHSetKnownFolderPath(ctypes.byref(g), 0, None,
                                          ctypes.c_wchar_p(new_path))
        if hr == 0:
            return True, '文件夹指向已更新'
        return False, 'SHSetKnownFolderPath 返回 0x%08X' % (hr & 0xFFFFFFFF)
    except Exception as e:
        return False, '更新失败：%s' % e


def _get_known_folder_default(fid):
    """读取注册表 User Shell Folders 中记录的原默认路径（用于对比）。"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders',
            0, winreg.KEY_READ)
        try:
            # 值名优先用 GUID 字符串，失败回退为常见名称映射
            names = [fid]
            for n, f, d, common, fb in KNOWN_FOLDERS:
                if f == fid:
                    names = [fid, {'我的文档': 'Personal'}.get(n, n)]
                    break
            for nm in names:
                try:
                    v, t = winreg.QueryValueEx(key, nm)
                    return v
                except OSError:
                    continue
        finally:
            winreg.CloseKey(key)
    except OSError:
        pass
    return None


# 第 10 次迭代：转移用户（个人文件夹）整体时更新注册表 UserProfile 指向。
# 当前用户 SID 通过 ProfileList 下 ProfileImagePath 与 %USERPROFILE% 匹配定位。
def _set_profile_image_path(new_path):
    """更新当前用户 ProfileImagePath（HKLM，需管理员）。返回 (ok, msg)。"""
    try:
        cur = os.environ.get('USERPROFILE', '').rstrip('\\')
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList',
            0, winreg.KEY_READ)
        sid = None
        try:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(key, i)
                except OSError:
                    break
                try:
                    sk = winreg.OpenKey(key, sub, 0, winreg.KEY_READ)
                    try:
                        p = winreg.QueryValueEx(sk, 'ProfileImagePath')[0]
                    finally:
                        winreg.CloseKey(sk)
                except OSError:
                    p = None
                if p and str(p).rstrip('\\').lower() == cur.lower():
                    sid = sub
                    break
                i += 1
        finally:
            winreg.CloseKey(key)
        if not sid:
            return False, '未找到当前用户 ProfileList 项，无法更新 ProfileImagePath'
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\%s' % sid,
            0, winreg.KEY_SET_VALUE)
        try:
            winreg.SetValueEx(key, 'ProfileImagePath', 0, winreg.REG_EXPAND_SZ,
                              new_path.rstrip('\\'))
        finally:
            winreg.CloseKey(key)
        return True, '用户文件夹指向（ProfileImagePath）已更新为 %s，注销后生效' % new_path
    except OSError as e:
        return False, '更新 ProfileImagePath 失败（可能无管理员权限）：%s' % e


# ---------------------------------------------------------------------------
# 系统信息
# ---------------------------------------------------------------------------
def get_system_info():
    """返回 (系统名, 版本, 内部版本号) 三元组。"""
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                             r'SOFTWARE\Microsoft\Windows NT\CurrentVersion',
                             0, winreg.KEY_READ)
        try:
            prod = winreg.QueryValueEx(key, 'ProductName')[0]
            disp = winreg.QueryValueEx(key, 'DisplayVersion')[0]
            build = winreg.QueryValueEx(key, 'CurrentBuildNumber')[0]
        finally:
            winreg.CloseKey(key)
    except OSError:
        prod, disp, build = platform.platform(), '', ''
    try:
        build_num = int(str(build).split('.')[0])
    except Exception:
        build_num = 0
    # Build >= 22000 判定为 Windows 11（注册表 ProductName 在 Win11 下仍可能
    # 显示 Windows 10，这里做展示修正）
    if build_num >= 22000 and prod.lower().startswith('windows 10'):
        prod = prod.replace('Windows 10', 'Windows 11', 1)
    elif build_num >= 22000 and 'Windows 10' in prod:
        prod = prod.replace('Windows 10', 'Windows 11', 1)
    return prod, disp, build


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 目录大小与格式化
# ---------------------------------------------------------------------------
class _WIN32_FIND_DATAW(ctypes.Structure):
    # WIN32_FIND_DATAW 为 4 字节对齐、总长 592 字节；
    # 若不强制 _pack_=4，ctypes 会把 FILETIME(c_uint64) 按 8 字节对齐，
    # 使结构变 600 字节，导致 FindFirstFileW 写入的数据读取错位。
    _pack_ = 4
    _fields_ = [
        ('dwFileAttributes', ctypes.c_uint32),
        ('ftCreationTime', ctypes.c_uint64),
        ('ftLastAccessTime', ctypes.c_uint64),
        ('ftLastWriteTime', ctypes.c_uint64),
        ('nFileSizeHigh', ctypes.c_uint32),
        ('nFileSizeLow', ctypes.c_uint32),
        ('dwReserved0', ctypes.c_uint32),
        ('dwReserved1', ctypes.c_uint32),
        ('cFileName', ctypes.c_wchar * 260),
        ('cAlternateFileName', ctypes.c_wchar * 14),
    ]


def dir_size(path):
    """递归统计目录大小（字节）。权限失败项跳过。

    采用与参考程序 pft.exe 完全相同的 Windows API 做法：
    FindFirstFileW / FindNextFileW 递归枚举，一次系统调用同时取得
    文件名、属性与文件大小（nFileSizeHigh/Low），无需额外 stat，
    比 os.walk + getsize 快数倍，避免大小长期停留在"计算中"。
    重解析点（符号链接/联接）跳过，防止循环递归。
    """
    total = 0
    if not path or not os.path.isdir(path):
        return 0
    if not hasattr(ctypes, 'windll'):
        return 0
    k32 = ctypes.windll.kernel32
    find_first = k32.FindFirstFileW
    find_next = k32.FindNextFileW
    find_close = k32.FindClose
    find_first.argtypes = [ctypes.c_wchar_p,
                           ctypes.POINTER(_WIN32_FIND_DATAW)]
    find_first.restype = ctypes.c_void_p
    find_next.argtypes = [ctypes.c_void_p,
                          ctypes.POINTER(_WIN32_FIND_DATAW)]
    find_next.restype = ctypes.c_int
    find_close.argtypes = [ctypes.c_void_p]
    find_close.restype = ctypes.c_int
    INVALID = 0xFFFFFFFFFFFFFFFF  # 64 位 INVALID_HANDLE_VALUE
    FILE_ATTRIBUTE_DIRECTORY = 0x10
    FILE_ATTRIBUTE_REPARSE_POINT = 0x400
    stack = [path]
    while stack:
        p = stack.pop()
        search = p.rstrip('\\/') + '\\*'
        fd = _WIN32_FIND_DATAW()
        h = find_first(search, ctypes.byref(fd))
        if h == INVALID or not h:
            continue
        try:
            while True:
                name = fd.cFileName
                if name not in ('.', '..'):
                    attrs = fd.dwFileAttributes
                    if attrs & FILE_ATTRIBUTE_REPARSE_POINT:
                        pass  # 跳过符号链接/联接
                    elif attrs & FILE_ATTRIBUTE_DIRECTORY:
                        stack.append(os.path.join(p, name))
                    else:
                        total += (fd.nFileSizeHigh << 32) | fd.nFileSizeLow
                if not find_next(h, ctypes.byref(fd)):
                    break
        finally:
            find_close(h)
    return total


SIZE_COLOR_KB = '#9e9e9e'   # 灰
SIZE_COLOR_MB = '#2f7fe0'   # 蓝
SIZE_COLOR_GB = '#f07818'   # 橙
SIZE_COLOR_TB = '#e53e3e'   # 红
_SIZE_COLORS = {'KB': SIZE_COLOR_KB, 'MB': SIZE_COLOR_MB,
                'GB': SIZE_COLOR_GB, 'TB': SIZE_COLOR_TB}


def fmt_size(n):
    """字节数格式化为 4 级单位：<1024KB 显示 KB，1024KB~1024MB 显示 MB，
    1024MB~1024GB 显示 GB，>=1024GB 显示 TB。数值前不补零。

    例如 512.00 KB、45.34 MB、1.21 GB、1.50 TB。
    """
    kb = n / 1024.0
    if kb < 1024:
        return '%.2f KB' % kb
    mb = kb / 1024.0
    if mb < 1024:
        return '%.2f MB' % mb
    gb = mb / 1024.0
    if gb < 1024:
        return '%.2f GB' % gb
    tb = gb / 1024.0
    return '%.2f TB' % tb


def size_color(text):
    """按单位返回大小标签颜色：KB=灰、MB=蓝、GB=橙、TB=红。"""
    unit = text.split()[-1] if text else ''
    return _SIZE_COLORS.get(unit, '#888888')


# ---------------------------------------------------------------------------
# robocopy 移动文件
# ---------------------------------------------------------------------------
def robocopy_move(src, dst, exclude_profile=False):
    """移动目录内容（含子目录）。返回 (ok, summary)。

    exclude_profile=True 时排除 NTUSER.* 等用户配置文件核心文件
    （第 10 次迭代：转移用户文件夹整体时使用）。
    """
    try:
        os.makedirs(dst, exist_ok=True)
    except OSError as e:
        return False, '创建目标目录失败：%s' % e
    if not os.path.isdir(src):
        return True, '源目录不存在，跳过移动'
    cmd = ['robocopy', src, dst, '/MOVE', '/E', '/COPY:DAT',
           '/R:1', '/W:1', '/NFL', '/NDL', '/NP', '/NJH', '/NJS']
    if exclude_profile:
        cmd += ['/XF', 'ntuser.*']
    try:
        p = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='gbk', errors='replace',
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        out, _ = p.communicate(timeout=3600)
    except Exception as e:
        return False, 'robocopy 执行失败：%s' % e
    code = p.returncode
    if code < 8:
        return True, '移动完成（robocopy 退出码 %d）' % code
    return False, '移动失败（robocopy 退出码 %d）' % code


# ---------------------------------------------------------------------------
# 第 24 次迭代：desktop.ini 处理（确认弹窗「覆盖与覆盖」选项）
# 两种模式：
#   keep —— 保留目标路径 desktop.ini：目标已有则保留（转移前备份、转移后
#            恢复，防止被 robocopy 从源移来的 desktop.ini 覆盖）；
#            目标没有则不创建。
#   copy —— 复制原本路径 desktop.ini：源路径有 desktop.ini 则复制到目标；
#            源没有则不创建。
# ---------------------------------------------------------------------------
def _desktop_ini_prepare(src, dst, mode):
    """转移前处理。keep 模式且目标已存在 desktop.ini 时暂存备份。
    返回备份路径或 None。"""
    if mode != 'keep':
        return None
    if not dst:
        return None
    try:
        dst_ini = os.path.join(dst, 'desktop.ini')
        if not os.path.isfile(dst_ini):
            return None
        bak = os.path.join(dst, '.desktop.ini.__mover_bak__')
        if os.path.exists(bak):
            try:
                os.remove(bak)
            except OSError:
                pass
        os.rename(dst_ini, bak)
        return bak
    except OSError:
        return None


def _desktop_ini_finalize(src, dst, mode, backup):
    """转移后按模式处理 desktop.ini。"""
    if not dst:
        return
    dst_ini = os.path.join(dst, 'desktop.ini')
    try:
        if mode == 'keep':
            # 移除可能被 robocopy 从源移来的 desktop.ini（目标没有则不创建）
            if os.path.isfile(dst_ini):
                try:
                    os.remove(dst_ini)
                except OSError:
                    pass
            # 恢复目标原有 desktop.ini
            if backup and os.path.isfile(backup):
                try:
                    os.rename(backup, dst_ini)
                except OSError:
                    pass
        elif mode == 'copy':
            # 把源路径 desktop.ini 复制到目标（源没有则不创建）
            if not src:
                return
            src_ini = os.path.join(src, 'desktop.ini')
            if os.path.isfile(src_ini):
                try:
                    os.makedirs(dst, exist_ok=True)
                    import shutil
                    shutil.copy2(src_ini, dst_ini)
                except OSError:
                    pass
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 清理垃圾
# ---------------------------------------------------------------------------
def clean_temp_files():
    """清理用户临时目录内容（跳过占用文件）。返回删除项数。"""
    cnt = 0
    tmp = os.environ.get('TEMP', '')
    if not tmp or not os.path.isdir(tmp):
        return 0
    for name in os.listdir(tmp):
        p = os.path.join(tmp, name)
        try:
            if os.path.isdir(p):
                import shutil
                shutil.rmtree(p, ignore_errors=True)
                cnt += 1
            else:
                os.remove(p)
                cnt += 1
        except OSError:
            pass
    return cnt


def clean_thumb_cache():
    """删除缩略图缓存文件。返回删除项数。"""
    cnt = 0
    base = os.path.join(os.environ.get('LOCALAPPDATA', ''),
                        'Microsoft', 'Windows', 'Explorer')
    if not os.path.isdir(base):
        return 0
    for name in os.listdir(base):
        if name.lower().startswith('thumbcache_') and name.lower().endswith('.db'):
            try:
                os.remove(os.path.join(base, name))
                cnt += 1
            except OSError:
                pass
    return cnt


def clean_recycle_bin():
    """清空回收站（PowerShell Clear-RecycleBin）。返回 (ok, msg)。"""
    try:
        subprocess.run(
            ['powershell', '-NoProfile', '-Command',
             'Clear-RecycleBin -Force -ErrorAction SilentlyContinue'],
            capture_output=True, timeout=120,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        return True, '回收站已清空'
    except Exception as e:
        return False, '清空回收站失败：%s' % e


# ---------------------------------------------------------------------------
# 应用目录 / 资源路径 / 字体注册
# ---------------------------------------------------------------------------
def app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(name):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, name)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)


_BUNDLED_FONTS = ('HartwellAlt-Bold.ttf', 'fa-solid-900.ttf')


def _register_fonts():
    """注册私有字体（仅当前进程可见，不写入系统）。"""
    try:
        gdi32 = ctypes.WinDLL('gdi32')
        gdi32.AddFontResourceExW.argtypes = [wt.LPCWSTR, wt.DWORD, ctypes.c_void_p]
        gdi32.AddFontResourceExW.restype = ctypes.c_int
        for f in _BUNDLED_FONTS:
            path = resource_path(f)
            if os.path.isfile(path):
                n = gdi32.AddFontResourceExW(path, 0x10, None)  # FR_PRIVATE
                if n == 0:
                    gdi32.AddFontResourceW(path)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------
class App:
    # ---------------- 配色 ----------------
    BG_TOP = (0xfe, 0xfe, 0xfe)     # 极浅灰白渐变顶部
    BG_BOT = (0xf7, 0xf7, 0xf7)     # 极浅灰白渐变底部
    CARD_BG = '#fbfbfb'
    CARD_BORDER = '#efefef'
    NAV_ACTIVE = '#f0f0f0'
    NAV_INACTIVE = '#fbfbfb'
    NAV_TEXT_ACTIVE = '#333333'
    NAV_TEXT = '#333333'
    TEXT_MAIN = '#333333'
    TEXT_DIM = '#888888'
    ACCENT = '#35d07f'
    WARN = '#f5b73d'
    DANGER = '#ff5d5d'
    FONT = 'Microsoft YaHei UI'
    NAV_ICONS = {'程序': '\uf06e', '设置': '\uf013'}

    TITLEBAR_H = 36
    ROUND_RADIUS = 16

    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self._dragging = False
        self._bg_dirty = False
        self._drag_sync_pending = False
        self._bg_cache = {}
        self._bg_photo = None
        self._nav_anim_running = False
        self._nav_anim_items = {}
        self._nav_anim_after_id = None
        self.current_page = '程序'
        self._switch_running = False

        # 状态数据
        self.sys_name, self.sys_ver, self.sys_build = get_system_info()
        self.move_files = tk.BooleanVar(value=True)
        self.allow_logoff = tk.BooleanVar(value=False)
        self.clean_garbage = tk.BooleanVar(value=False)
        # 第 10 次迭代：批量目标即用户文件夹目标，不再使用 batch_root
        self.user_target_var = tk.StringVar(value='')
        # 每个系统文件夹: name -> {fid, default_dir, common, var_enabled,
        #                         var_target, current_path, size_text, size_bytes}
        self.folder_states = {}
        for name, fid, d, common, fb in KNOWN_FOLDERS:
            if fid:
                cur = _get_known_folder_path(fid, fb,
                                             _SHELL_FOLDER_VALUE_NAMES.get(name, []))
            else:
                # 无 Known Folder GUID（如临时文件）：直接按环境变量模板定位
                cur = _expand_env(fb)
            # Known Folder 指向的目录不存在时，按顺序尝试常见实际路径
            if cur and not os.path.isdir(cur):
                for alt in _ALTERNATE_PATHS.get(name, []):
                    p = _expand_env(alt)
                    if p and os.path.isdir(p):
                        cur = p
                        break
            self.folder_states[name] = {
                'fid': fid,
                'default_dir': d,
                'common': common,
                'var_enabled': tk.BooleanVar(value=common),
                'var_target': tk.StringVar(value=''),
                'current_path': cur or '',
                'size_text': '计算中...',
                'size_bytes': -1,
                'row_widgets': [],
                'manual_set': False,
            }
        # 用户（个人文件夹）独立状态：重要页展示；第 10 次迭代起作为跟随基准，
        # 用户目标即批量目标，转移与否由批量目标是否设置决定
        cur_user = _get_user_profile_path()
        self.user_state = {
            'fid': '{5E6C858F-0E22-4760-9AFE-EA3317B67173}',
            'default_dir': 'User',
            'common': False,
            'var_enabled': tk.BooleanVar(value=False),
            'var_target': self.user_target_var,
            'current_path': cur_user or '',
            'size_text': '计算中...',
            'size_bytes': -1,
            'row_widgets': [],
        }

        # 第 10 次迭代：用户目标变化时自动重算跟随项并刷新例外标记
        self.user_target_var.trace_add('write', self._on_user_target_change)

        self._apply_window_icon()
        self._build()
        self._apply_window_icon()
        self.root.bind('<<QueuePoll>>', lambda e: self._poll_queue(), add='+')
        self.root.after_idle(self._poll_queue)
        self._start_size_workers()
        self.root.after(300, self._ensure_sizes_shown)
        self.log('系统文件夹转移工具已就绪。')
        if not is_admin():
            self.log('提示：当前未以管理员身份运行，转移系统文件夹可能失败。建议右键“以管理员身份运行”。')

    # ---------------- 窗口图标 ----------------
    def _apply_window_icon(self):
        try:
            u32 = ctypes.windll.user32
            u32.LoadImageW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p,
                                       ctypes.c_uint, ctypes.c_int, ctypes.c_int,
                                       ctypes.c_uint]
            u32.LoadImageW.restype = ctypes.c_void_p
            u32.GetAncestor.argtypes = [ctypes.c_void_p, ctypes.c_uint]
            u32.GetAncestor.restype = ctypes.c_void_p
            u32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                         ctypes.c_void_p, ctypes.c_void_p]
            u32.SendMessageW.restype = ctypes.c_void_p
            icon_path = resource_path('app_icon.ico')
            hbig = u32.LoadImageW(None, icon_path, 1, 32, 32, 0x10)
            hsmall = u32.LoadImageW(None, icon_path, 1, 16, 16, 0x10)
            hwnd = self.root.winfo_id()
            if hwnd:
                hwnd = u32.GetAncestor(hwnd, 2)
            if hwnd:
                if hbig:
                    u32.SendMessageW(hwnd, 0x0080, 1, hbig)
                if hsmall:
                    u32.SendMessageW(hwnd, 0x0080, 0, hsmall)
        except Exception:
            pass

    # ---------------- 基础绘制 ----------------
    def _rounded(self, c, x1, y1, x2, y2, r, fill='', outline='', width=1, tags=None):
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
               x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
               x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        return c.create_polygon(pts, smooth=True, fill=fill, outline=outline,
                                width=width, tags=tags)

    def _draw_gradient(self, canvas, w, h):
        if self._dragging or self._drag_sync_pending:
            self._bg_dirty = True
            return
        if w < 2 or h < 2:
            return
        canvas.delete('bg')
        steps = max(h, 2)
        for i in range(steps):
            t = i / (steps - 1)
            r = int(self.BG_TOP[0] + (self.BG_BOT[0] - self.BG_TOP[0]) * t)
            g = int(self.BG_TOP[1] + (self.BG_BOT[1] - self.BG_TOP[1]) * t)
            b = int(self.BG_TOP[2] + (self.BG_BOT[2] - self.BG_TOP[2]) * t)
            color = '#%02x%02x%02x' % (r, g, b)
            canvas.create_line(0, i, w, i, fill=color, tags='bg')
        canvas.tag_lower('bg')
        self._bg_dirty = False

    # ---------------- 窗口结构 ----------------
    def _hwnd(self):
        try:
            return ctypes.windll.user32.GetParent(self.root.winfo_id())
        except Exception:
            return None

    def _build(self):
        self.root.title(WINDOW_TITLE)
        self.root.overrideredirect(False)
        # 第 19 次迭代：主窗口高度增加 50px（760 -> 810），最小高度同步 +50
        self.root.geometry('1180x810')
        self.root.minsize(980, 670)
        self.root.configure(bg='#f7f7f7')
        self._strip_chrome()
        self._last_rgn_size = (0, 0)
        self.root.bind('<Configure>', self._on_root_configure)

        self.bg = tk.Canvas(self.root, highlightthickness=0, bd=0)
        self.bg.place(x=0, y=0, relwidth=1, relheight=1)
        self.bg.bind('<Configure>',
                     lambda e: self._draw_gradient(self.bg, e.width, e.height))

        self._build_titlebar()

        self.nav = tk.Canvas(self.root, width=196, highlightthickness=0, bd=0,
                             bg='#fbfbfb')
        self.nav.place(x=0, y=0, relheight=1, height=0)
        self._build_nav()

        self.content = tk.Frame(self.root, bg='#f7f7f7')
        self.content.place(x=196, y=0, relwidth=1, relheight=1,
                           width=-196, height=0)

        self.pages = {}
        self.pages['程序'] = self._build_page_program(self.content)
        self.pages['设置'] = self._build_page_settings(self.content)
        self.pages['关于'] = self._build_page_about(self.content)
        for p in self.pages.values():
            p.place(relwidth=1, relheight=1)

        self.titlebar_left.lift()
        self.titlebar_mid.lift()
        self.titlebar_right.lift()

        self._switch('程序')
        self.root.after(200, self._ensure_taskbar)

    def _ensure_taskbar(self):
        try:
            hwnd = self._hwnd()
            if hwnd:
                ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                                  0x0001 | 0x0002 | 0x0004 | 0x0020)
        except Exception:
            pass

    def _strip_chrome(self):
        try:
            self.root.update_idletasks()
            hwnd = self._hwnd()
            if not hwnd:
                return
            st = ctypes.windll.user32.GetWindowLongW(hwnd, -16)
            st &= ~(0x00C00000 | 0x00080000 | 0x00010000 | 0x00040000)
            st |= 0x02000000 | 0x04000000
            ctypes.windll.user32.SetWindowLongW(hwnd, -16, st)
            ex = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
            ex |= 0x40000
            ex &= ~0x80
            ctypes.windll.user32.SetWindowLongW(hwnd, -20, ex)
            ctypes.windll.user32.SetWindowPos(
                hwnd, 0, 0, 0, 0, 0,
                0x0001 | 0x0002 | 0x0004 | 0x0020)
            self._apply_round_region(self.ROUND_RADIUS)
        except Exception:
            pass

    def _apply_round_region(self, radius):
        try:
            hwnd = self._hwnd()
            if not hwnd:
                return
            w = self.root.winfo_width()
            h = self.root.winfo_height()
            if w <= 0 or h <= 0:
                return
            self._last_rgn_size = (w, h)
            pref = 1 if radius <= 0 else 2
            try:
                ret = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    ctypes.c_void_p(hwnd), 33,
                    ctypes.byref(ctypes.c_int(pref)), 4)
                if ret == 0:
                    ctypes.windll.user32.SetWindowRgn(hwnd, None, True)
                    return
            except Exception:
                pass
            if radius <= 0:
                ctypes.windll.user32.SetWindowRgn(hwnd, None, True)
                return
            r = int(radius)
            rgn = ctypes.windll.gdi32.CreateRoundRectRgn(
                0, 0, w + 1, h + 1, r * 2, r * 2)
            if not rgn:
                return
            if not ctypes.windll.user32.SetWindowRgn(hwnd, rgn, True):
                ctypes.windll.gdi32.DeleteObject(rgn)
        except Exception:
            pass

    def _on_root_configure(self, e):
        try:
            if e.widget is not self.root:
                return
            if self._dragging:
                return
            if (e.width, e.height) == self._last_rgn_size:
                return
            self._apply_round_region(0 if self._is_maximized else self.ROUND_RADIUS)
        except Exception:
            pass

    # ---------------- 自绘标题栏 ----------------
    def _make_tb_btn(self, tb, text, cmd, hover_close=False):
        btn = tk.Label(tb, text=text, bg=tb['bg'], fg='#888888',
                       font=('Segoe UI Symbol', 11), width=3, cursor='hand2',
                       relief='flat', bd=0, highlightthickness=0)
        btn.bind('<Button-1>', lambda e: cmd())
        if hover_close:
            btn.bind('<Enter>', lambda e: btn.configure(bg='#f9e6e6'))
            btn.bind('<Leave>', lambda e: btn.configure(bg=tb['bg']))
        else:
            btn.bind('<Enter>', lambda e: btn.configure(bg='#f8f8f8'))
            btn.bind('<Leave>', lambda e: btn.configure(bg=tb['bg']))
        return btn

    def _build_titlebar(self):
        tb_left = tk.Frame(self.root, bg=self.NAV_INACTIVE,
                           height=self.TITLEBAR_H, highlightthickness=0, bd=0)
        tb_left.place(x=0, y=0, width=196, height=self.TITLEBAR_H)
        self.titlebar_left = tb_left
        tb_right = tk.Frame(self.root, bg='#f7f7f7',
                            height=self.TITLEBAR_H, highlightthickness=0, bd=0)
        tb_right.place(relx=1.0, x=-10, y=0, width=118,
                       height=self.TITLEBAR_H, anchor='ne')
        self.titlebar_right = tb_right
        tb_mid = tk.Frame(self.root, bg='#f7f7f7',
                          height=self.TITLEBAR_H, highlightthickness=0, bd=0)
        tb_mid.place(x=196, y=0, relwidth=1, width=-324,
                     height=self.TITLEBAR_H)
        self.titlebar_mid = tb_mid
        self._is_maximized = False
        self._prev_rect = None
        self._drag_off = None
        self._last_press_at = None
        self._dbl_pending = False
        self._press_x = 0
        self._press_y = 0
        self._press_hwnd = None
        self._drag_hwnd = None
        self._drag_hold = False
        self._poll_active = False

        self._tb_btn_close = self._make_tb_btn(tb_right, '\u2715', self._tb_close,
                                               hover_close=True)
        self._tb_btn_close.pack(side='right', fill='y')
        self._tb_btn_max = self._make_tb_btn(tb_right, '\u25a1', self._tb_max)
        self._tb_btn_max.pack(side='right', fill='y')
        self._tb_btn_min = self._make_tb_btn(tb_right, '\u2014', self._tb_min)
        self._tb_btn_min.pack(side='right', fill='y')

        for tb in (tb_left, tb_mid, tb_right):
            tb.bind('<Button-1>', self._tb_drag_start)
            tb.bind('<B1-Motion>', self._tb_drag_move)
            tb.bind('<ButtonRelease-1>', self._tb_drag_end)
            tb.bind('<Double-Button-1>', lambda e: self._tb_toggle_max())
            for w in tb.winfo_children():
                if isinstance(w, tk.Label) and w not in (self._tb_btn_close,
                                                         self._tb_btn_max,
                                                         self._tb_btn_min):
                    w.bind('<Button-1>', self._tb_drag_start)
                    w.bind('<B1-Motion>', self._tb_drag_move)
                    w.bind('<ButtonRelease-1>', self._tb_drag_end)
                    w.bind('<Double-Button-1>', lambda e: self._tb_toggle_max())

        hwnd = self._hwnd()
        if hwnd:
            try:
                ex = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
                ctypes.windll.user32.SetWindowLongW(hwnd, -20,
                                                    ex | 0x40000)
            except Exception:
                pass

    def _tb_drag_start(self, e):
        now = time.time()
        dbl_ms = ctypes.windll.user32.GetDoubleClickTime()
        if self._last_press_at and (now - self._last_press_at) * 1000.0 < dbl_ms:
            self._last_press_at = None
            self._dbl_pending = False
            self._tb_toggle_max()
            return
        self._last_press_at = now
        self._dbl_pending = True
        self._press_x = e.x_root
        self._press_y = e.y_root
        self._press_hwnd = self._hwnd()
        self.root.after(60, self._check_start_drag)

    def _check_start_drag(self):
        if not getattr(self, '_dbl_pending', False):
            return
        try:
            if not (ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000):
                self._dbl_pending = False
                return
            pt = wt.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
            if (abs(pt.x - self._press_x) < 4 and
                    abs(pt.y - self._press_y) < 4):
                self.root.after(50, self._check_start_drag)
                return
        except Exception:
            self._dbl_pending = False
            return
        self._dbl_pending = False
        if self._is_maximized:
            self._tb_toggle_max()
        hwnd = self._press_hwnd or self._hwnd()
        if not hwnd:
            return
        try:
            pt = wt.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
            r = wt.RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r))
            self._drag_off = (pt.x - r.left, pt.y - r.top)
            self._drag_hwnd = hwnd
            self._dragging = True
            self._drag_hold = True
            self._drag_poll()
        except Exception:
            self._drag_hwnd = None
            self._drag_off = None
            self._dragging = False
            self._drag_hold = False
            self._after_drag_sync()

    def _drag_poll(self):
        if not self._dragging:
            return
        try:
            if not (ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000):
                self._dragging = False
                self._drag_hold = False
                self._drag_hwnd = None
                self._drag_off = None
                self._after_drag_sync()
                return
            pt = wt.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
            hwnd = getattr(self, '_drag_hwnd', None)
            if hwnd and self._drag_off:
                x = pt.x - self._drag_off[0]
                y = pt.y - self._drag_off[1]
                ctypes.windll.user32.SetWindowPos(
                    hwnd, 0, x, y, 0, 0,
                    0x0001 | 0x0004 | 0x0010)
            self.root.after(5, self._drag_poll)
        except Exception:
            self._dragging = False
            self._drag_hold = False
            self._drag_hwnd = None
            self._drag_off = None
            self._after_drag_sync()

    def _tb_drag_move(self, e):
        return

    def _after_drag_sync(self):
        self._drag_sync_pending = True
        self._drag_hold = False
        # 拖拽结束：若消费未在进行则唤醒一轮（事件驱动，无常驻轮询）
        if not self._poll_active:
            self.root.after_idle(self._poll_queue)
        try:
            self.root.update_idletasks()
        except Exception:
            pass
        self.root.after_idle(self._flush_bg_redraw)

    def _flush_bg_redraw(self):
        if self._dragging:
            self._bg_dirty = True
            return
        try:
            w = self.bg.winfo_width()
            h = self.bg.winfo_height()
            if w < 2 or h < 2:
                self.root.after_idle(self._flush_bg_redraw)
                return
            self._drag_sync_pending = False
            self._draw_gradient(self.bg, w, h)
            self._bg_dirty = False
            if not self._is_maximized:
                self._apply_round_region(self.ROUND_RADIUS)
        except Exception:
            self.root.after_idle(self._flush_bg_redraw)

    def _tb_drag_end(self, e):
        if getattr(self, '_dbl_pending', False):
            self._dbl_pending = False
        if self._dragging:
            self._dragging = False
            self._after_drag_sync()

    def _tb_toggle_max(self):
        hwnd = self._hwnd()
        if not hwnd:
            return
        if self._is_maximized:
            if self._prev_rect:
                ctypes.windll.user32.SetWindowPos(
                    hwnd, 0,
                    self._prev_rect[0], self._prev_rect[1],
                    self._prev_rect[2] - self._prev_rect[0],
                    self._prev_rect[3] - self._prev_rect[1],
                    0x0004)
            self._is_maximized = False
            self._tb_btn_max.config(text='\u25a1')
            self._apply_round_region(self.ROUND_RADIUS)
        else:
            r = wt.RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r))
            self._prev_rect = (r.left, r.top, r.right, r.bottom)
            class _MONITORINFO(ctypes.Structure):
                _fields_ = [
                    ('cbSize', ctypes.c_ulong),
                    ('rcMonitor', ctypes.wintypes.RECT),
                    ('rcWork', ctypes.wintypes.RECT),
                    ('dwFlags', ctypes.c_ulong),
                ]
            mi = _MONITORINFO()
            mi.cbSize = ctypes.sizeof(_MONITORINFO)
            mon = ctypes.windll.user32.MonitorFromWindow(hwnd, 2)
            if mon and ctypes.windll.user32.GetMonitorInfoW(mon, ctypes.byref(mi)):
                rc = mi.rcWork
            else:
                sp = wt.RECT()
                ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(sp), 0)
                rc = sp
            ctypes.windll.user32.SetWindowPos(
                hwnd, 0, rc.left, rc.top,
                rc.right - rc.left, rc.bottom - rc.top,
                0x0004)
            self._is_maximized = True
            self._tb_btn_max.config(text='\u2750')
            self._apply_round_region(0)

    def _tb_min(self):
        hwnd = self._hwnd()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 6)  # SW_MINIMIZE

    def _tb_max(self):
        self._tb_toggle_max()

    def _tb_close(self):
        self._on_close()

    def _on_close(self):
        try:
            self.root.destroy()
        except Exception:
            pass

    # ---------------- 导航 ----------------
    def _build_nav(self):
        n = self.nav
        _TITLE_AREA_H = 90
        self._nav_title = n.create_text(
            44, _TITLE_AREA_H // 2, text='Masker', anchor='w',
            fill=self.NAV_TEXT, font=('Hartwell Alt Bold', 24),
            tags='nav_title')
        for ev, fn in (('<Button-1>', self._tb_drag_start),
                       ('<B1-Motion>', self._tb_drag_move),
                       ('<ButtonRelease-1>', self._tb_drag_end),
                       ('<Double-Button-1>', self._tb_toggle_max)):
            n.tag_bind('nav_title', ev, fn)

        items = ['程序']
        self._nav_tags = {}
        self._nav_icon_w = 13
        self._nav_text_x = 44 + self._nav_icon_w + 6
        self._nav_text_cx = (14 + 182) / 2.0
        y0 = 96
        for i, name in enumerate(items):
            y = y0 + i * 40
            tag = 'nav_%d' % i
            ttag = tag + '_t'
            itag = tag + '_i'
            self._nav_tags[name] = (tag, y)
            self._rounded(n, 14, y, 182, y + 34, 10, fill=self.NAV_INACTIVE,
                          outline='', tags=tag)
            n.create_text(44, y + 17, anchor='w',
                          text=self.NAV_ICONS.get(name, ''),
                          fill=self.NAV_TEXT, font=('Font Awesome 6 Free Solid', 10),
                          tags=itag)
            n.create_text(self._nav_text_x, y + 17, anchor='w', text=name,
                          fill=self.NAV_TEXT, font=(self.FONT, 10, 'bold'),
                          tags=ttag)
            n.tag_bind(tag, '<Button-1>', lambda e, nm=name: self._switch(nm))
            n.tag_bind(tag, '<Enter>', lambda e, nm=name: self._nav_hover(nm, True))
            n.tag_bind(tag, '<Leave>', lambda e, nm=name: self._nav_hover(nm, False))
            n.tag_bind(ttag, '<Button-1>', lambda e, nm=name: self._switch(nm))
            n.tag_bind(itag, '<Button-1>', lambda e, nm=name: self._switch(nm))

        name = '设置'
        tag = 'nav_settings'
        ttag = tag + '_t'
        itag = tag + '_i'
        y = 540
        self._nav_tags[name] = (tag, y)
        self._nav_settings_poly = self._rounded(
            n, 14, y, 182, y + 34, 10, fill=self.NAV_INACTIVE, outline='',
            tags=tag)
        self._nav_settings_icon = n.create_text(
            44, y + 17, anchor='w', text=self.NAV_ICONS.get(name, ''),
            fill=self.NAV_TEXT, font=('Font Awesome 6 Free Solid', 10), tags=itag)
        self._nav_settings_label = n.create_text(
            self._nav_text_x, y + 17, anchor='w', text=name, fill=self.NAV_TEXT,
            font=(self.FONT, 10, 'bold'), tags=ttag)
        n.tag_bind(tag, '<Button-1>', lambda e, nm=name: self._switch(nm))
        n.tag_bind(tag, '<Enter>', lambda e, nm=name: self._nav_hover(nm, True))
        n.tag_bind(tag, '<Leave>', lambda e, nm=name: self._nav_hover(nm, False))
        n.tag_bind(ttag, '<Button-1>', lambda e, nm=name: self._switch(nm))
        n.tag_bind(itag, '<Button-1>', lambda e, nm=name: self._switch(nm))

        self._nav_divider = n.create_line(16, 524, 180, 524, fill='#e8e8e8')
        n.bind('<Configure>', self._relayout_nav)

    def _relayout_nav(self, event=None):
        try:
            h = self.nav.winfo_height()
            if h < 200:
                return
            y = h - 52
            x1, x2, r = 14, 182, 10
            pts = [x1 + r, y, x2 - r, y, x2, y, x2, y + r,
                   x2, y + 34 - r, x2, y + 34, x2 - r, y + 34,
                   x1 + r, y + 34, x1, y + 34, x1, y + 34 - r,
                   x1, y + r, x1, y]
            self.nav.coords(self._nav_settings_poly, *pts)
            self.nav.coords(self._nav_settings_icon, 44, y + 17)
            self.nav.coords(self._nav_settings_label, self._nav_text_x, y + 17)
            self.nav.coords(self._nav_divider, 16, y - 16, 180, y - 16)
            self._nav_tags['设置'] = (self._nav_tags['设置'][0], y)
        except Exception:
            pass

    def _nav_hover(self, name, enter):
        tag, y = self._nav_tags[name]
        if name == self.current_page:
            return
        self.nav.itemconfig(tag, fill='#f0f0f0' if enter else self.NAV_INACTIVE)

    def _nav_redraw(self):
        for name, (tag, y) in self._nav_tags.items():
            active = (name == self.current_page)
            self.nav.itemconfig(tag,
                                fill=self.NAV_ACTIVE if active else self.NAV_INACTIVE)
            for i in self.nav.find_withtag(tag + '_t'):
                self.nav.itemconfig(i, fill=self.NAV_TEXT_ACTIVE if active else self.NAV_TEXT)
            for i in self.nav.find_withtag(tag + '_i'):
                self.nav.itemconfig(i, fill=self.NAV_TEXT_ACTIVE if active else self.NAV_TEXT)

    def _nav_anim_start(self):
        base = self._nav_text_base_x()
        cx = getattr(self, '_nav_text_cx', 98)
        if getattr(self, '_nav_anim_after_id', None) is not None:
            try:
                self.root.after_cancel(self._nav_anim_after_id)
            except Exception:
                pass
            self._nav_anim_after_id = None
        items = {}
        for nm, (tag, y) in self._nav_tags.items():
            ttag = tag + '_t'
            coords = self.nav.coords(ttag)
            cur = coords[0] if coords else base
            target = cx if nm == self.current_page else base
            if abs(cur - target) > 1:
                items[nm] = [ttag, cur, target]
        if not items:
            self._nav_anim_running = False
            self._nav_anim_items = {}
            return
        self._nav_anim_running = True
        self._nav_anim_items = items
        self._nav_anim_step()

    def _nav_anim_step(self):
        items = getattr(self, '_nav_anim_items', None)
        if not items:
            self._nav_anim_running = False
            return
        step = 4
        done = True
        for nm, (ttag, cur, target) in items.items():
            _co = self.nav.coords(ttag)
            y = _co[1] if len(_co) > 1 else 0
            if abs(cur - target) <= step:
                self.nav.coords(ttag, target, y)
            else:
                cur += step if target > cur else -step
                self.nav.coords(ttag, cur, y)
                items[nm][1] = cur
                done = False
        if done:
            self._nav_anim_running = False
            self._nav_anim_items = {}
            self._nav_anim_after_id = None
        else:
            self._nav_anim_after_id = self.root.after(16, self._nav_anim_step)

    def _nav_text_base_x(self):
        return getattr(self, '_nav_text_x', 60)

    def _switch(self, name):
        self.current_page = name
        for p in self.pages.values():
            p.place_forget()
        self.pages[name].place(relwidth=1, relheight=1)
        self._nav_redraw()
        self._nav_anim_start()

    # ---------------- 程序页 ----------------
    def _build_page_program(self, parent):
        page = tk.Frame(parent, bg='#f7f7f7')

        head = tk.Frame(page, bg='#f7f7f7')
        head.pack(fill='x', padx=28, pady=(40, 6))
        tk.Label(head, text='程序', bg='#f7f7f7', fg=self.TEXT_MAIN,
                 font=(self.FONT, 18, 'bold')).pack(side='left')
        tk.Label(head, text='转移系统文件夹到自定义位置',
                 bg='#f7f7f7', fg=self.TEXT_DIM,
                 font=(self.FONT, 11)).pack(side='left', padx=(10, 0))

        btn_kw = dict(bg='#f7f7f7', bd=0, relief='flat', highlightthickness=0,
                      cursor='hand2')
        self._log_icon_btn = tk.Label(head, text='\u269d', fg='#888888',
                                      font=('Segoe UI Symbol', 15), **btn_kw)
        self._log_icon_btn.pack(side='right')
        self._log_icon_btn.bind('<Button-1>', lambda e: self._toggle_log_panel())
        self._log_icon_btn.bind('<Enter>',
                                lambda e: self._log_icon_btn.configure(bg='#f0f0f0'))
        self._log_icon_btn.bind('<Leave>', lambda e: self._set_log_btn_state(
            self._log_visible))

        # 系统信息横幅
        banner = tk.Frame(page, bg='#f7f7f7')
        banner.pack(fill='x', padx=28, pady=(8, 0))
        self.sys_banner = tk.Canvas(banner, height=62, bg='#f7f7f7',
                                    highlightthickness=0, bd=0)
        self.sys_banner.pack(fill='x')
        banner.bind('<Configure>', lambda e: self._draw_sys_banner(e.width))

        # 第 12 次迭代：移除上方独立的「用户文件夹目标」批量栏，
        # 批量目标入口统一到下方「用户」行本身（名称/大小/当前/目标三列）。

        # 第 13 次迭代：用户行独立卡片，展示在页签栏上方
        self._build_user_card(page)

        # 文件夹列表（重要 / 其他 2 页签，见 _build_folder_rows）

        # 文件夹列表（滚动区）
        list_wrap = tk.Frame(page, bg='#f7f7f7')
        list_wrap.pack(fill='both', expand=True, padx=28, pady=(8, 4))
        self.list_wrap = list_wrap
        # 页签头（重要 / 其他，两按钮平分横向宽度各 50%）
        self._build_list_tab_head(list_wrap)
        self.list_canvas = tk.Canvas(list_wrap, bg='#f7f7f7',
                                     highlightthickness=0, bd=0)
        self.list_scroll = tk.Canvas(list_wrap, width=6, bg='#f7f7f7',
                                     highlightthickness=0, bd=0)
        self.list_canvas.pack(side='left', fill='both', expand=True)
        self.list_scroll.pack(side='right', fill='y')
        self.list_inner = tk.Frame(self.list_canvas, bg='#f7f7f7')
        self._list_window = self.list_canvas.create_window(
            (0, 0), window=self.list_inner, anchor='nw')
        self.list_inner.bind('<Configure>',
                             lambda e: (self.list_canvas.configure(
                                 scrollregion=self.list_canvas.bbox('all')),
                                 self._list_sb_kick()))
        self.list_canvas.bind('<Configure>',
                              lambda e: (self.list_canvas.itemconfig(
                                  self._list_window, width=e.width),
                                 self._list_sb_kick()))
        # 全局滚轮绑定：鼠标落在列表滚动区（含其内部任意子控件）时滚动页面，
        # 避免滚轮事件被子控件吞掉导致无法滚动。
        self.list_canvas.bind_all('<MouseWheel>', self._on_list_wheel_global)
        # 滚动条
        # 滚动条（红线式：无轨道背景，仅一条红色细线作为滑动块；静止自动隐藏）
        self._list_sb_hide_after = None
        self._list_sb_drag_off = 0
        self._list_sb_geo = self.list_scroll.create_rectangle(0, 0, 6, 20,
                                                              fill='', outline='')
        self._list_sb_item = self.list_scroll.create_polygon(
            *_round_rect_pts(0, 0, 6, 20, 3),
            fill='#e53935', outline='#e53935', width=1, smooth=True)
        self.list_scroll.itemconfigure(self._list_sb_item, state='hidden')
        self.list_scroll.tag_raise('all')
        self.list_scroll.bind('<ButtonPress-1>', self._list_sb_press)
        self.list_scroll.bind('<B1-Motion>', self._list_sb_drag)
        self.list_scroll.bind('<ButtonRelease-1>', lambda e: self._list_sb_kick())
        self.list_canvas.bind('<MouseWheel>', self._on_list_wheel)
        self.list_inner.bind('<MouseWheel>', self._on_list_wheel)

        # 文件夹卡片行
        self._build_folder_rows()

        # 底部执行
        foot = tk.Frame(page, bg='#f7f7f7')
        foot.pack(fill='x', padx=28, pady=(4, 14))
        self.btn_move = tk.Button(foot, text='开始转移', command=self._start_move,
                                  bg=self.ACCENT, fg='#ffffff', relief='flat', bd=0,
                                  activebackground='#2bb26c',
                                  activeforeground='#ffffff',
                                  font=(self.FONT, 12, 'bold'), cursor='hand2',
                                  padx=26, pady=5)
        self.btn_move.pack(side='left')
        tk.Label(foot, text='勾选要转移的系统文件夹，配置目标路径后点击开始转移。',
                 bg='#f7f7f7', fg='#888888',
                 font=(self.FONT, 9)).pack(side='left', padx=(14, 0))

        # 右侧日志栏
        self.log_panel = tk.Frame(self.root, bg='#f7f7f7', highlightthickness=0)
        self._log_visible = False
        self._log_anim = False
        self._log_cur = 0
        self._log_target = 0
        self.log_panel.place_forget()
        lp_sep = tk.Frame(self.log_panel, bg='#e8e8e8', width=1)
        lp_sep.pack(side='left', fill='y')
        lp_head = tk.Frame(self.log_panel, bg='#f7f7f7')
        lp_head.pack(fill='x', pady=(40, 0))
        tk.Label(lp_head, text='运行日志', bg='#f7f7f7', fg=self.TEXT_MAIN,
                 font=(self.FONT, 14, 'bold')).pack(pady=(10, 8))
        self.log_text = tk.Text(self.log_panel, height=8, bg='#f7f7f7', fg='#333333',
                                insertbackground='#333333', relief='flat', bd=0,
                                font=('Consolas', 9), state='disabled', wrap='word',
                                padx=10, pady=4)
        sb = tk.Canvas(self.log_panel, width=6, bg='#f7f7f7',
                       highlightthickness=0, bd=0)
        self._log_sb = sb
        self._log_sb_hide_after = None
        self._log_sb_drag_off = 0
        self._log_sb_geo = sb.create_rectangle(0, 0, 6, 20, fill='', outline='')
        self._log_sb_item = sb.create_polygon(
            *_round_rect_pts(0, 0, 6, 20, 3),
            fill='#e53935', outline='#e53935', width=1, smooth=True)
        sb.itemconfigure(self._log_sb_item, state='hidden')
        sb.tag_raise('all')
        self.log_text.bind('<MouseWheel>', self._log_sb_kick)
        sb.bind('<ButtonPress-1>', self._log_sb_press)
        sb.bind('<B1-Motion>', self._log_sb_drag)
        sb.bind('<ButtonRelease-1>', lambda e: self._log_sb_kick(e, release=True))
        self.log_text.configure(yscrollcommand=self._log_sb_update)
        self.log_text.pack(side='left', fill='both', expand=True)

        return page

    def _draw_sys_banner(self, w):
        c = self.sys_banner
        c.delete('all')
        self._rounded(c, 0, 0, w, 62, 12, fill=self.CARD_BG,
                      outline=self.CARD_BORDER, width=1)
        c.create_text(22, 31, anchor='w',
                      text='\uf1c0', font=('Font Awesome 6 Free Solid', 22),
                      fill=self.ACCENT)
        c.create_text(58, 17, anchor='w', text='当前系统：%s' % self.sys_name,
                      fill=self.TEXT_MAIN, font=(self.FONT, 11, 'bold'))
        c.create_text(58, 40, anchor='w',
                      text='版本：%s    内部版本：%s    %s' %
                           (self.sys_ver or '未知', self.sys_build or '未知',
                            '管理员权限' if is_admin() else '普通权限'),
                      fill=self.TEXT_DIM, font=(self.FONT, 9))

    # ---------------- 文件夹列表行 ----------------
    # ---------------- 文件夹列表（重要 / 次要 / 其他 3 页签） ----------------
    def _build_list_tab_head(self, list_wrap):
        head = tk.Frame(list_wrap, bg='#f7f7f7')
        head.pack(fill='x', pady=(0, 0))
        self._list_tab_btns = {}
        self._list_tab_cur = 'important'
        for tid, tname in (('important', '重要'), ('minor', '次要'), ('other', '其他')):
            b = tk.Label(head, text=tname, cursor='hand2',
                         font=(self.FONT, 11, 'bold'),
                         bg='#f0f0f0', fg='#888888', pady=6,
                         highlightthickness=1,
                         highlightbackground='#e8e8e8')
            b.pack(side='left', fill='x', expand=True)
            b.bind('<Button-1>', lambda e, t=tid: self._switch_list_tab(t))
            self._list_tab_btns[tid] = b

        # 第 18 次迭代：页签说明置顶并简化样式——
        # 去掉圆角浅色卡片边框，改为紧贴页签按钮栏下沿、与页签同宽的水平延展灰色小字区。
        self._list_notes = {}
        note_bar = tk.Frame(list_wrap, bg='#f0f0f0')
        note_bar.pack(fill='x')
        self._list_note_label = tk.Label(
            note_bar, text='', justify='left', anchor='w',
            bg='#f0f0f0', fg='#999999', font=(self.FONT, 9), padx=12, pady=6)
        self._list_note_label.pack(fill='x')
        note_bar.bind('<Configure>', lambda e: self._list_note_label.configure(
            wraplength=max(10, e.width - 24)))

    def _set_list_note(self, tid):
        lines = self._list_notes.get(tid, ())
        self._list_note_label.configure(text='\n'.join(lines))

    def _build_folder_rows(self):
        self._list_pages = {}
        for tid in ('important', 'minor', 'other'):
            pg = tk.Frame(self.list_inner, bg='#f7f7f7')
            self._list_pages[tid] = pg
        # 第 13 次迭代：用户行独立展示在页签栏上方（见 _build_user_card），
        # 重要页仅保留文档、下载、音乐、图片、视频
        self._build_list_page('important', [
            ('文档', '我的文档'), ('下载', '下载'), ('音乐', '音乐'),
            ('图片', '图片'), ('视频', '视频'),
        ], notes=[
            '说明：文档、下载、音乐、图片、视频是常用数据文件夹，转移可释放系统盘空间；',
            '转移会移动现有文件并更新系统指向，建议提前确认目标盘剩余空间充足；',
            '未单独设置的文件夹将跟随用户目标自动生成同名子目录，单独设置目标即为例外。',
        ])
        # 次要页：第 8 次迭代新增的 4 项 + 收藏夹、联系人、链接、搜索、保存的游戏
        # 第 17 次迭代：桌面从其他页迁入次要页（属次要/缓存类文件夹）
        # 第 19 次迭代：桌面调整到次要页第一位（其余项顺序不变）
        self._build_list_page('minor', [
            ('桌面', '桌面'),
            ('IE缓存', 'IE缓存'), ('Cookies', 'Cookies'), ('临时文件', '临时文件'),
            ('收藏夹', '收藏夹'), ('联系人', '联系人'), ('链接', '链接'),
            ('搜索', '搜索'), ('保存的游戏', '保存的游戏'),
        ], notes=[
            '说明：IE 缓存、Cookies、临时文件、收藏夹、联系人、链接、搜索、保存的游戏、桌面等',
            '多为缓存、临时或低价值数据，转移可进一步减少系统盘占用；缓存类文件夹被系统',
            '频繁读写，转移后访问速度受目标盘性能影响，建议谨慎选择目标位置。',
        ])
        # 其他页：剩余系统文件夹（第 17 次迭代：桌面已移至次要页）
        self._build_list_page('other', [
            ('发送到', '发送到'), ('3D对象', '3D对象'),
            ('相机胶卷', '相机胶卷'), ('开始菜单', '开始菜单'), ('启动', '启动'),
            ('最近', '最近'),
        ], notes=[
            '说明：发送到、3D 对象、相机胶卷、开始菜单、启动、最近等为较少使用的系统项，',
            '一般无需转移；如确需转移请谨慎，开始菜单、启动等项会影响系统界面与开机行为。',
        ])
        self._switch_list_tab('important')

    def _build_list_page(self, tid, items, notes=None):
        pg = self._list_pages[tid]
        for display, name in items:
            self._build_row(pg, display, name)
        # 第 18 次迭代：说明文本登记到置顶说明条（页签按钮栏下沿），由
        # _set_list_note 在切换页签时刷新；不再在页内底部绘制圆角说明卡。
        if notes:
            self._list_notes[tid] = list(notes)

    # 第 13 次迭代：用户行/普通文件夹行统一构建方法（原 _build_list_page 循环体）
    def _build_row(self, pg, display, name):
        if name == '用户':
            st = self.user_state
        else:
            st = self.folder_states.get(name)
            if st is None:
                return
        row = tk.Frame(pg, bg=self.CARD_BG,
                       highlightbackground=self.CARD_BORDER,
                       highlightthickness=1)
        row.pack(fill='x', pady=(0, 6))
        st['row_widgets'] = []

        # 第一行：勾选 + 名称 + 大小 + 当前路径
        line1 = tk.Frame(row, bg=self.CARD_BG)
        line1.pack(fill='x', padx=14, pady=(8, 2))
        chk = tk.Checkbutton(line1, variable=st['var_enabled'], bg=self.CARD_BG,
                             activebackground=self.CARD_BG,
                             highlightthickness=0, bd=0, cursor='hand2',
                             selectcolor='#ffffff')
        chk.pack(side='left')
        st['row_widgets'].append(chk)
        tk.Label(line1, text=display, bg=self.CARD_BG, fg=self.TEXT_MAIN,
                 font=(self.FONT, 11, 'bold'), width=10,
                 anchor='w').pack(side='left', padx=(2, 10))
        # 大小列固定 12 字符宽度（Consolas 等宽），保证各行的"当前："列起始位置对齐
        size_lbl = tk.Label(line1, text=st['size_text'], bg=self.CARD_BG,
                            fg=size_color(st['size_text']),
                            font=('Consolas', 10), width=12, anchor='w')
        size_lbl.pack(side='left')
        st['row_widgets'].append(size_lbl)
        st['size_lbl'] = size_lbl
        # 若大小统计线程在标签创建前已完成（结果缓存在 size_bytes / size_text），
        # 直接回填显示，避免消息消费早于标签创建导致大小永远不显示
        if st['size_bytes'] > 0:
            txt = fmt_size(st['size_bytes'])
            if st['size_text'] != txt:
                st['size_text'] = txt
                size_lbl.configure(text=txt, fg=size_color(txt))
        tk.Label(line1, text='当前：', bg=self.CARD_BG, fg=self.TEXT_DIM,
                 font=(self.FONT, 9)).pack(side='left', padx=(14, 0))
        # 第 11 次迭代：用户行提示文字放"当前"横向行的最右端并右对齐
        # 第 16 次迭代：导入按钮已移至第二行与保存按钮并排（见 line2）
        if name == '用户':
            tk.Label(line1, text='未单独设置的文件夹跟随用户转移；单独设置目标则为例外。',
                     bg=self.CARD_BG, fg='#999999',
                     font=(self.FONT, 9)).pack(side='right', padx=(0, 4))
        cur_lbl = tk.Label(line1, text=st['current_path'] or '（未找到）',
                           bg=self.CARD_BG, fg='#555555', cursor='hand2',
                           font=(self.FONT, 9), anchor='w')
        cur_lbl.pack(side='left', fill='x', expand=True, padx=(4, 10))
        # 第 8 次迭代：当前地址 hover 平滑加粗 / 移开恢复 / 点击用资源管理器打开
        cur_lbl.bind('<Enter>',
                     lambda e, l=cur_lbl: self._cur_lbl_anim_start(l, True))
        cur_lbl.bind('<Leave>',
                     lambda e, l=cur_lbl: self._cur_lbl_anim_start(l, False))
        cur_lbl.bind('<Button-1>',
                     lambda e, l=cur_lbl, s=st: self._open_cur_path(l, s))
        st['cur_lbl'] = cur_lbl
        st['row_widgets'].append(cur_lbl)

        # 第二行：目标路径 + 例外标记 + 浏览
        line2 = tk.Frame(row, bg=self.CARD_BG)
        line2.pack(fill='x', padx=14, pady=(2, 8))
        tk.Label(line2, text='目标：', bg=self.CARD_BG, fg=self.TEXT_DIM,
                 font=(self.FONT, 9)).pack(side='left')
        ent = tk.Entry(line2, textvariable=st['var_target'],
                       bg='#ffffff', fg='#333333', relief='flat', bd=0,
                       highlightthickness=1, highlightbackground='#e8e8e8',
                       highlightcolor=self.ACCENT, font=(self.FONT, 9))
        ent.pack(side='left', fill='x', expand=True, padx=(6, 0))
        st['row_widgets'].append(ent)
        # 第 10 次迭代：例外标记（单独设置目标、不跟随用户批量目标的文件夹）
        st['exc_lbl'] = None
        if name != '用户':
            exc_lbl = tk.Label(line2, text='例外', bg='#e53935', fg='#ffffff',
                               font=(self.FONT, 8, 'bold'), padx=6, pady=1)
            exc_lbl.pack(side='left', padx=(8, 0))
            exc_lbl.pack_forget()
            st['exc_lbl'] = exc_lbl
            st['row_widgets'].append(exc_lbl)
            # 手动编辑目标后判定例外
            ent.bind('<FocusOut>',
                     lambda e, nm=name: self._on_target_edited(nm))
        # 第 12 次迭代：用户行作为批量目标入口，目标列改用与批量目标一致的
        # 黑色「选择...」按钮（替换原灰色浏览按钮），其余文件夹行保持灰色浏览
        if name == '用户':
            btn = tk.Button(line2, text='选择...',
                            command=lambda nm=name: self._pick_single_target(nm),
                            bg='#333333', fg='#ffffff', relief='flat', bd=0,
                            activebackground='#555555',
                            activeforeground='#ffffff',
                            font=(self.FONT, 9, 'bold'), cursor='hand2',
                            padx=10, pady=1)
        else:
            btn = tk.Button(line2, text='浏览...',
                            command=lambda nm=name: self._pick_single_target(nm),
                            bg='#f0f0f0', fg='#333333', relief='flat', bd=0,
                            activebackground='#e4e4e4', activeforeground='#333333',
                            font=(self.FONT, 9), cursor='hand2', padx=10, pady=1)
        btn.pack(side='left', padx=(8, 0))
        st['row_widgets'].append(btn)
        # 第 16 次迭代：用户行「选择...」右侧并排「保存位置」「导入位置」按钮
        # （保存：把当前所有文件夹的当前位置写入 .ini 快照；导入：读取快照设目标）
        if name == '用户':
            save_btn = tk.Button(line2, text='保存位置',
                                 command=self._save_snapshot,
                                 bg=self.ACCENT, fg='#ffffff', relief='flat', bd=0,
                                 activebackground='#4cd98c',
                                 activeforeground='#ffffff',
                                 font=(self.FONT, 9, 'bold'), cursor='hand2',
                                 padx=10, pady=1)
            save_btn.pack(side='left', padx=(8, 0))
            st['row_widgets'].append(save_btn)
            imp_btn = tk.Button(line2, text='导入位置',
                                command=self._load_snapshot,
                                bg=self.ACCENT, fg='#ffffff', relief='flat', bd=0,
                                activebackground='#4cd98c',
                                activeforeground='#ffffff',
                                font=(self.FONT, 9, 'bold'), cursor='hand2',
                                padx=10, pady=1)
            imp_btn.pack(side='left', padx=(8, 0))
            st['row_widgets'].append(imp_btn)

    # 第 13 次迭代：用户行独立卡片，展示在页签栏上方（banner 与列表之间）
    def _build_user_card(self, parent):
        pg = tk.Frame(parent, bg='#f7f7f7')
        pg.pack(fill='x', padx=28, pady=(8, 0))
        self._build_row(pg, '用户', '用户')

    # ---------------- 当前地址：hover 平滑加粗 + 点击打开 ----------------
    def _cur_lbl_anim_start(self, lbl, entering):
        """鼠标进入/离开当前地址标签：6 帧颜色渐变 + 末尾切换字重。"""
        try:
            if getattr(lbl, '_hover_after', None) is not None:
                lbl.after_cancel(lbl._hover_after)
            lbl._hover_after = None
        except Exception:
            pass
        start = '#555555'
        end = self.ACCENT
        if not entering:
            start, end = end, start
        steps = 6
        dt = 20

        def _step(i):
            t = (i + 1) / float(steps)
            r = int(round(int(start[1:3], 16) +
                          (int(end[1:3], 16) - int(start[1:3], 16)) * t))
            g = int(round(int(start[3:5], 16) +
                          (int(end[3:5], 16) - int(start[3:5], 16)) * t))
            b = int(round(int(start[5:7], 16) +
                          (int(end[5:7], 16) - int(start[5:7], 16)) * t))
            try:
                lbl.configure(fg='#%02x%02x%02x' % (r, g, b))
                if i == steps - 1:
                    lbl.configure(font=(self.FONT, 9,
                                        'bold' if entering else 'normal'))
                    lbl._hover_after = None
                else:
                    lbl._hover_after = lbl.after(
                        dt, lambda: _step(i + 1))
            except Exception:
                lbl._hover_after = None
        lbl._hover_after = lbl.after(dt, lambda: _step(0))

    def _open_cur_path(self, lbl, st):
        """点击当前地址：用资源管理器打开该目录；目录不存在时提示。"""
        try:
            if getattr(lbl, '_hover_after', None) is not None:
                lbl.after_cancel(lbl._hover_after)
            lbl._hover_after = None
            lbl.configure(font=(self.FONT, 9, 'bold'), fg=self.ACCENT)
        except Exception:
            pass
        path = (st.get('current_path') or '').strip()
        if not path or not os.path.isdir(path):
            messagebox.showwarning(
                '提示', '该目录不存在，无法打开。\n\n%s' % (path or '（未找到）'),
                parent=self.root)
            return
        try:
            os.startfile(path)
            self.log('已打开目录：%s' % path)
        except Exception as e:
            messagebox.showwarning('提示', '打开目录失败：%s' % e,
                                   parent=self.root)

    def _switch_list_tab(self, tid):
        self._list_tab_cur = tid
        for k, pg in self._list_pages.items():
            pg.pack_forget()
        self._list_pages[tid].pack(fill='both', expand=True)
        self._set_list_note(tid)
        # 切换页签后回到顶部，避免残留上一页签的滚动位置导致新页内容
        # 显示偏移 / 滚动条位置错乱
        try:
            self.list_canvas.yview_moveto(0)
        except Exception:
            pass
        for k, b in self._list_tab_btns.items():
            active = (k == tid)
            b.configure(bg=self.ACCENT if active else '#f0f0f0',
                        fg='#ffffff' if active else '#888888')
        try:
            self.root.after_idle(self._list_sb_kick)
        except Exception:
            pass

    def _pick_single_target(self, name):
        if name == '用户':
            d = filedialog.askdirectory(title='为「用户（个人文件夹）」选择目标文件夹',
                                        parent=self.root)
            if d:
                self.user_target_var.set(d)
                self.log('「用户（个人文件夹）」目标文件夹已设置为：%s' % d)
            return
        st = self.folder_states[name]
        d = filedialog.askdirectory(title='为「%s」选择目标文件夹' % name,
                                    parent=self.root)
        if d:
            st['var_target'].set(d)
            st['manual_set'] = True
            self.log('「%s」目标文件夹已设置为：%s（已标记为例外，不再跟随用户文件夹）'
                     % (name, d))
            self._refresh_exception_label(name)

    def _pick_batch_root(self):
        d = filedialog.askdirectory(title='选择「用户（个人文件夹）」统一目标路径',
                                    parent=self.root)
        if d:
            self.user_target_var.set(d)
            self.log('用户文件夹目标已设置为：%s（未单独设置的文件夹将自动跟随）' % d)

    def _apply_batch(self):
        ut = self.user_target_var.get().strip()
        if not ut:
            messagebox.showwarning('提示', '请先在「用户」行设置目标路径。',
                                   parent=self.root)
            return
        self._refresh_follow_states()
        self.log('已应用跟随：用户目标 = %s，例外项保持单独目标' % ut)

    # ---------------- 第 10 次迭代：跟随 / 例外逻辑 ----------------
    def _follow_path(self, name):
        """未单独设置目标时该文件夹的跟随目标（用户目标路径 + 相对子路径）。"""
        ut = self.user_target_var.get().strip()
        if not ut:
            return ''
        sub = _FOLLOW_SUBDIRS.get(name)
        if not sub:
            sub = self.folder_states[name]['default_dir']
        return os.path.normpath(os.path.join(ut, sub))

    def _is_exception(self, name):
        """单独设置了其他目标位置（与跟随路径不同）即为例外项。"""
        if name == '用户':
            return False
        st = self.folder_states.get(name)
        if not st:
            return False
        if not self.user_target_var.get().strip():
            return False
        t = st['var_target'].get().strip()
        if not t:
            return False
        return os.path.normpath(t) != os.path.normpath(self._follow_path(name))

    def _refresh_exception_label(self, name):
        st = self.folder_states.get(name)
        lbl = (st or {}).get('exc_lbl')
        if not lbl:
            return
        try:
            if self._is_exception(name):
                lbl.pack(side='left', padx=(8, 0))
            else:
                lbl.pack_forget()
        except Exception:
            pass

    def _on_target_edited(self, name):
        """目标 Entry 失去焦点：手动编辑后记录 manual_set 并刷新例外标记。"""
        st = self.folder_states.get(name)
        if not st:
            return
        t = st['var_target'].get().strip()
        if t:
            fp = self._follow_path(name)
            if fp and os.path.normpath(t) == os.path.normpath(fp):
                st['manual_set'] = False
            else:
                st['manual_set'] = True
        self._refresh_exception_label(name)

    def _on_user_target_change(self, *args):
        """用户目标变化：自动重算跟随项、联动用户行勾选、刷新例外标记。"""
        try:
            self._refresh_follow_states()
        except Exception:
            pass

    def _refresh_follow_states(self):
        """重算所有跟随项目标；用户勾选由批量目标是否设置决定；刷新例外标记。"""
        ut = self.user_target_var.get().strip()
        for name, st in self.folder_states.items():
            if not st.get('manual_set'):
                fp = self._follow_path(name)
                st['var_target'].set(fp)
            self._refresh_exception_label(name)
        # 用户文件夹本身：设置批量目标即转移（自动勾选）；清空则不转移
        try:
            self.user_state['var_enabled'].set(bool(ut))
        except Exception:
            pass
        if not ut:
            self.log('用户文件夹目标已清空，用户文件夹将不参与转移。')
        else:
            self.log('用户文件夹目标：%s' % ut)

    # ---------------- 大小计算 ----------------
    def _start_size_workers(self):
        for name, st in self.folder_states.items():
            cur = st['current_path']
            if not cur:
                st['size_text'] = '不可用'
                continue
            threading.Thread(target=self._size_worker, args=(name, cur),
                             daemon=True).start()
        us = self.user_state
        if us['current_path']:
            threading.Thread(target=self._size_worker,
                             args=('用户', us['current_path']),
                             daemon=True).start()
        else:
            us['size_text'] = '不可用'

    def _size_worker(self, name, path):
        n = -1
        try:
            n = dir_size(path)
        except Exception as e:
            # 统计异常也回传结果（-1 表示失败），避免标签永久停留在"计算中..."
            try:
                self.q.put(('size', name, -1))
                self._wake_poll()
            except Exception:
                pass
            return
        try:
            self.q.put(('size', name, n))
            self._wake_poll()
        except Exception:
            pass

    # ---------------- 转移 ----------------
    def _start_move(self):
        if self._switch_running:
            self.log('转移正在进行中，请稍候...')
            return
        enabled = [(n, st) for n, st in self.folder_states.items()
                   if st['var_enabled'].get()]
        # 第 10 次迭代：用户文件夹是否转移由批量目标是否设置决定（勾选联动）
        ut = self.user_target_var.get().strip()
        if self.user_state['var_enabled'].get():
            if not ut:
                messagebox.showwarning(
                    '提示', '「用户（个人文件夹）」已勾选，但未设置目标路径。\n\n请先在「用户」行中设置目标路径。',
                    parent=self.root)
                return
            enabled.append(('用户', self.user_state))
        if not enabled:
            messagebox.showwarning('提示', '请至少勾选一个要转移的系统文件夹。',
                                   parent=self.root)
            return
        missing = [(n, st) for n, st in enabled
                   if n != '用户' and not st['var_target'].get().strip()]
        if missing:
            names = '、'.join(n for n, _ in missing)
            messagebox.showwarning('提示', '以下文件夹未设置目标路径：\n%s\n\n请为它们选择目标文件夹，或在「用户」行设置批量目标。' % names,
                                   parent=self.root)
            return
        mode = '转移文件' if self.move_files.get() else '不转移文件'
        self._show_start_confirm(enabled, mode)

    def _show_start_confirm(self, enabled, mode):
        dlg = tk.Toplevel(self.root)
        dlg.title('开始转移 - 配置确认')
        dlg.configure(bg='#f7f7f7')
        dlg.resizable(False, False)
        dlg.transient(self.root)
        try:
            dlg.grab_set()
        except Exception:
            pass
        try:
            dlg.attributes('-topmost', True)
        except Exception:
            pass

        pad = 20
        # 第 24 次迭代：头部一行放标题与「全貌」切换按钮
        top_row = tk.Frame(dlg, bg='#f7f7f7')
        top_row.pack(anchor='w', padx=pad, pady=(16, 6), fill='x')
        tk.Label(top_row, text='即将开始转移，请确认以下完整配置：',
                 bg='#f7f7f7', fg='#333333',
                 font=(self.FONT, 11, 'bold')).pack(side='left')
        show_all = [False]
        btn_all = tk.Button(top_row, text='全貌', command=None,
                            bg='#f7f7f7', fg=self.ACCENT, relief='flat', bd=0,
                            font=(self.FONT, 10, 'bold'), cursor='hand2',
                            padx=10, pady=4, activebackground='#f0f0f0',
                            activeforeground=self.ACCENT)
        btn_all.pack(side='right')

        ut = self.user_target_var.get().strip()
        follow_items = [(n, st) for n, st in self.folder_states.items()
                        if not self._is_exception(n)]
        exc_items = [(n, st) for n, st in self.folder_states.items()
                     if self._is_exception(n)]
        lines = ['模式：%s' % mode,
                 '清理垃圾：%s' % ('开启' if self.clean_garbage.get() else '关闭'),
                 '允许注销：%s' % ('开启' if self.allow_logoff.get() else '关闭')]
        lines.append('用户(个人文件夹)批量目标：%s'
                     % (ut or '（未设置）'))
        lines.append('跟随项：%d 个（未单独设置，跟随用户文件夹）；例外项：%d 个（单独设置目标）'
                     % (len(follow_items), len(exc_items)))
        info = '\n'.join(lines)
        tk.Label(dlg, text=info, bg='#f7f7f7', fg='#888888',
                 font=(self.FONT, 9), anchor='w', justify='left'
                 ).pack(anchor='w', padx=pad)

        body = tk.Frame(dlg, bg='#ffffff', highlightbackground='#e8e8e8',
                        highlightthickness=1)
        body.pack(fill='both', expand=True, padx=pad, pady=(2, 0))
        # 第 21 次迭代：表头与数据行共用一个 grid 容器，各列天然对齐；
        # 列宽全部内容自适应（minsize=0）；当前路径列左 padding 6 微微右移。
        # 第 24 次迭代：新增第 5 列「覆盖与覆盖」（desktop.ini 处理选项）。
        # 第 25 次迭代：第 5 列标题改为「文件夹属性ℹ️」，ℹ️ 悬停显示说明。
        body.grid_columnconfigure(0, weight=0, minsize=0)
        body.grid_columnconfigure(1, weight=1, minsize=0)
        body.grid_columnconfigure(2, weight=1, minsize=0)
        body.grid_columnconfigure(3, weight=0, minsize=0)
        body.grid_columnconfigure(4, weight=0, minsize=0)

        def _bind_tooltip(w, text):
            # 简单悬停气泡：Enter 显示、Leave 销毁，窗口置顶无边框
            tip = [None]

            def _show(_e):
                try:
                    if tip[0] is not None:
                        tip[0].destroy()
                except Exception:
                    pass
                t = tk.Toplevel(dlg)
                t.wm_overrideredirect(True)
                try:
                    t.wm_attributes('-topmost', True)
                except Exception:
                    pass
                try:
                    t.wm_attributes('-alpha', 0.95)
                except Exception:
                    pass
                tk.Label(t, text=text, bg='#333333', fg='#ffffff',
                         font=(self.FONT, 9), justify='left',
                         padx=10, pady=8, wraplength=340).pack()
                x = w.winfo_rootx()
                y = w.winfo_rooty() + w.winfo_height() + 4
                t.wm_geometry('+%d+%d' % (x, y))
                tip[0] = t

            def _hide(_e):
                if tip[0] is not None:
                    try:
                        tip[0].destroy()
                    except Exception:
                        pass
                    tip[0] = None

            w.bind('<Enter>', _show)
            w.bind('<Leave>', _hide)

        # 第 24 次迭代：表头去掉背景色（bg 与 body 一致），仅用字体加粗区分
        for col, txt in enumerate(('文件夹', '当前路径', '目标路径', '是否转移')):
            hd = tk.Label(body, text=txt, bg='#ffffff', fg='#888888',
                          font=(self.FONT, 9, 'bold'))
            hd._is_header = True
            hd.grid(
                row=0, column=col, sticky='w',
                padx=(10, 8) if col == 0 else
                ((6, 8) if col == 1 else ((0, 8) if col == 2 else
                 ((0, 10) if col == 3 else (0, 10)))),
                pady=2)
        # 第 5 列表头：「文件夹属性」+ ℹ️ 信息图标（悬停显示说明）
        hd5 = tk.Frame(body, bg='#ffffff')
        hd5._is_header = True
        tk.Label(hd5, text='文件夹属性', bg='#ffffff', fg='#888888',
                 font=(self.FONT, 9, 'bold')).pack(side='left')
        info_lb = tk.Label(hd5, text='\u2139\ufe0f', bg='#ffffff',
                           fg='#2b6cb0', font=(self.FONT, 9, 'bold'),
                           cursor='hand2')
        info_lb.pack(side='left', padx=(2, 0))
        _bind_tooltip(
            info_lb,
            '控制转移时如何处理目标文件夹的 desktop.ini 文件：\n'
            '勾选「不覆盖」：保留目标路径已有的 desktop.ini'
            '（目标路径没有则不创建）。\n'
            '不勾选：将源路径的 desktop.ini 复制到目标路径'
            '（源路径没有则不创建）。')
        hd5.grid(row=0, column=4, sticky='w', padx=(0, 10), pady=2)

        # 第 24 次迭代：每项 desktop.ini 处理选项，智能默认：目标路径已存在
        # desktop.ini → 保留目标路径；不存在 → 复制原本路径。
        # 第 25 次迭代：控件改为单个勾选框「不覆盖」（勾选=keep、不勾选=copy），
        # desktop_modes 仍保存 StringVar(keep/copy) 供转移逻辑使用，重建行时复用；
        # check_vars 保存每项「不覆盖」勾选状态，勾选/取消通过 _on_check 同步。
        desktop_modes = {}
        check_vars = {}

        def _desktop_default(dst):
            try:
                if dst and os.path.isfile(os.path.join(dst, 'desktop.ini')):
                    return 'keep'
            except OSError:
                pass
            return 'copy'

        def _desktop_var(name, dst):
            if name not in desktop_modes:
                desktop_modes[name] = tk.StringVar(
                    value=_desktop_default(dst))
            return desktop_modes[name]

        def _desktop_checkvar(name, dst):
            if name not in check_vars:
                default = _desktop_default(dst)
                check_vars[name] = tk.BooleanVar(
                    value=(default == 'keep'))
                # 同步初始化 desktop_modes，保证未点击勾选框时也有智能默认值
                if name not in desktop_modes:
                    desktop_modes[name] = tk.StringVar(value=default)
            return check_vars[name]

        def _on_check(name):
            # 勾选「不覆盖」= keep（保留目标路径）；不勾选 = copy（复制原本路径）
            desktop_modes[name].set(
                'keep' if check_vars[name].get() else 'copy')

        # 数据行由 _render 按折叠状态重建：折叠时未配置行直接不布局，
        # 不占用请求高度，展开/折叠后重新 _fit_dialog 自适应窗口尺寸。
        row_idx = [1]

        def _clear_data():
            for w in body.winfo_children():
                if getattr(w, '_is_header', False):
                    continue
                w.destroy()

        def _add_group(txt):
            w = tk.Label(body, text=txt, bg='#ffffff', fg='#666666',
                         font=(self.FONT, 9, 'bold'), anchor='w')
            w.grid(row=row_idx[0], column=0, columnspan=5, sticky='we',
                   padx=10, pady=4)
            row_idx[0] += 1

        def _add_row(name, cur, dst, on, exc=False, check_var=None):
            # 第 24 次迭代：勾选转移的行仅名称加粗，去掉浅绿背景色
            row_bg = '#ffffff'
            name_font = (self.FONT, 9, 'bold') if on else (self.FONT, 9)
            name_txt = name
            if exc:
                name_txt = name + '（例外）'
            tk.Label(body, text=name_txt, bg=row_bg, fg='#333333',
                     font=name_font, anchor='w').grid(
                row=row_idx[0], column=0, sticky='w', padx=(10, 8), pady=2)
            tk.Label(body, text=cur, bg=row_bg, fg='#888888',
                     font=('Consolas', 9), anchor='w').grid(
                row=row_idx[0], column=1, sticky='w', padx=(6, 8))
            tk.Label(body, text=dst, bg=row_bg,
                     fg=self.ACCENT if exc else '#888888',
                     font=('Consolas', 9), anchor='w').grid(
                row=row_idx[0], column=2, sticky='w', padx=(0, 8))
            tk.Label(body, text='\u2713' if on else '\u2717',
                     bg=row_bg, fg=self.ACCENT if on else '#bbbbbb',
                     font=('Segoe UI Symbol', 10), anchor='w').grid(
                row=row_idx[0], column=3, sticky='w', padx=(0, 10))
            # 第 25 次迭代：单个勾选框「不覆盖」替代两个单选，不再撑高行高。
            # 勾选 = keep（保留目标路径 desktop.ini）；不勾选 = copy（复制原本路径）
            df = tk.Frame(body, bg=row_bg)
            tk.Checkbutton(df, text='不覆盖', variable=check_var,
                           command=lambda n=name: _on_check(n),
                           bg=row_bg, fg='#666666',
                           activebackground=row_bg, activeforeground='#333333',
                           font=(self.FONT, 8), anchor='w', bd=0,
                           highlightthickness=0, cursor='hand2'
                           ).pack(anchor='w')
            df.grid(row=row_idx[0], column=4, sticky='w',
                    padx=(0, 10), pady=1)
            row_idx[0] += 1

        def _render():
            _clear_data()
            row_idx[0] = 1
            # 用户行：勾选且批量目标已设置才算已配置
            user_on = self.user_state['var_enabled'].get()
            user_cfg = bool(user_on and ut)
            if show_all[0] or user_cfg:
                _add_group('用户(个人文件夹)批量目标')
                _add_row('用户(个人文件夹)',
                         self.user_state.get('current_path') or '（未知）',
                         ut or '（未设置）', user_on,
                         check_var=_desktop_checkvar('用户', ut))
            # 跟随项 / 例外项：勾选转移且目标路径已设置才算已配置
            for group_txt, items, exc in (
                    ('跟随项（%d）' % len(follow_items), follow_items, False),
                    ('例外项（%d）' % len(exc_items), exc_items, True)):
                cfgs = [bool(st['var_enabled'].get()
                            and st['var_target'].get().strip())
                        for _, st in items]
                if show_all[0] or any(cfgs):
                    _add_group(group_txt)
                    for (name, st), cfg in zip(items, cfgs):
                        if show_all[0] or cfg:
                            dst_txt = st['var_target'].get().strip()
                            _add_row(name,
                                     st.get('current_path') or '（未知）',
                                     dst_txt or '（未设置）',
                                     st['var_enabled'].get(), exc=exc,
                                     check_var=_desktop_checkvar(name, dst_txt))
            _fit_dialog()

        def _fit_dialog():
            dlg.update_idletasks()
            # 第 20 次迭代：窗口高度自适应内容并约束在屏幕可视范围内
            scr_w = dlg.winfo_screenwidth()
            scr_h = dlg.winfo_screenheight()
            req_w = dlg.winfo_reqwidth() + 20
            req_h = dlg.winfo_reqheight() + 20
            w = max(760, req_w)
            h = req_h
            if w > scr_w - 40:
                w = scr_w - 40
            if h > scr_h - 80:
                h = scr_h - 80
            x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
            x = max(8, min(x, scr_w - w - 8))
            y = max(8, min(y, scr_h - h - 40))
            dlg.geometry('%dx%d+%d+%d' % (w, h, x, y))

        def _toggle_all():
            show_all[0] = not show_all[0]
            btn_all.config(text='折叠' if show_all[0] else '全貌')
            _render()

        btn_all.config(command=_toggle_all)
        _render()

        tk.Label(dlg, text='是否先保存当前位置快照再开始转移？',
                 bg='#f7f7f7', fg='#333333',
                 font=(self.FONT, 10, 'bold')).pack(anchor='w', padx=pad,
                                                    pady=(14, 4))

        modes = {k: v.get() for k, v in desktop_modes.items()}

        def do_save_then_start():
            dlg.destroy()
            if self._save_snapshot():
                self._begin_transfer(enabled, modes)

        def do_direct_start():
            dlg.destroy()
            self._begin_transfer(enabled, modes)

        def do_cancel():
            dlg.destroy()

        btn_frame = tk.Frame(dlg, bg='#f7f7f7')
        btn_frame.pack(fill='x', padx=pad, pady=(4, 16))
        # 第 21 次迭代：顺序调整为 不保存直接转移(左)/保存当前路径并转移(中)/取消(右)，
        # 两个主按钮颜色对调（绿色给"不保存直接转移"，深色给"保存当前路径并转移"），取消保持原样
        tk.Button(btn_frame, text='不保存直接转移', command=do_direct_start,
                  bg=self.ACCENT, fg='#ffffff', activebackground='#2bb06b',
                  activeforeground='#ffffff', relief='flat', bd=0,
                  font=(self.FONT, 10, 'bold'), cursor='hand2',
                  padx=16, pady=6).pack(side='left')
        tk.Button(btn_frame, text='保存当前路径并转移', command=do_save_then_start,
                  bg='#333333', fg='#ffffff', activebackground='#555555',
                  activeforeground='#ffffff', relief='flat', bd=0,
                  font=(self.FONT, 10, 'bold'), cursor='hand2',
                  padx=16, pady=6).pack(side='left', padx=(10, 0))
        tk.Button(btn_frame, text='取消', command=do_cancel,
                  bg='#f7f7f7', fg='#888888', activebackground='#f0f0f0',
                  activeforeground='#888888', relief='flat', bd=0,
                  font=(self.FONT, 10, 'bold'), cursor='hand2',
                  padx=16, pady=6).pack(side='left', padx=(10, 0))

        _fit_dialog()

    def _begin_transfer(self, enabled, desktop_modes=None):
        if self._switch_running:
            self.log('转移正在进行中，请稍候...')
            return
        self._switch_running = True
        # 第 24 次迭代：desktop_modes 为每项选定的 desktop.ini 处理模式
        threading.Thread(target=self._move_worker,
                         args=(list(enabled), desktop_modes or {}),
                         daemon=True).start()

    def _move_worker(self, enabled, desktop_modes=None):
        results = []
        ok_all = True
        desktop_modes = desktop_modes or {}
        try:
            for name, st in enabled:
                # 第 10 次迭代：用户（个人文件夹）整体转移走独立流程
                if name == '用户':
                    ok, m = self._move_user_folder(
                        st, desktop_modes.get(name, 'keep'))
                    if ok:
                        results.append((name, True, m))
                    else:
                        ok_all = False
                        results.append((name, False, m))
                    continue
                src = st['current_path'] or ''
                dst = st['var_target'].get().strip()
                fid = st['fid']
                # 第 24 次迭代：按确认弹窗所选模式处理 desktop.ini
                dmode = desktop_modes.get(name, 'keep')
                backup = _desktop_ini_prepare(src, dst, dmode)
                self.q.put(('log', '--- 正在处理「%s」---' % name))
                self._wake_poll()
                try:
                    # 1) 转移文件（可选）
                    if self.move_files.get():
                        if src and os.path.isdir(src):
                            self.q.put(('log', '「%s」移动文件：%s -> %s' % (name, src, dst)))
                            self._wake_poll()
                            ok, m = robocopy_move(src, dst)
                            if not ok:
                                self.q.put(('log', '「%s」文件移动失败：%s' % (name, m)))
                                self._wake_poll()
                                ok_all = False
                                results.append((name, False, m))
                                continue
                            self.q.put(('log', '「%s」%s' % (name, m)))
                            self._wake_poll()
                        else:
                            self.q.put(('log', '「%s」源目录不存在或为空，跳过移动' % name))
                            self._wake_poll()
                    else:
                        self.q.put(('log', '「%s」已选择不转移文件，仅更新文件夹指向' % name))
                        self._wake_poll()
                    # 2) 更新 Known Folder 指向（无 GUID 的文件夹如临时文件仅移动文件）
                    try:
                        os.makedirs(dst, exist_ok=True)
                    except OSError as e:
                        self.q.put(('log', '「%s」创建目标目录失败：%s' % (name, e)))
                        self._wake_poll()
                        ok_all = False
                        results.append((name, False, '创建目标目录失败'))
                        continue
                    if fid:
                        ok, m = _set_known_folder_path(fid, dst)
                        self.q.put(('log', '「%s」%s' % (name, m)))
                        self._wake_poll()
                    else:
                        ok, m = True, '无 Known Folder 指向（临时目录），仅移动文件'
                        self.q.put(('log', '「%s」%s' % (name, m)))
                        self._wake_poll()
                    if ok:
                        results.append((name, True, '已完成'))
                        st['current_path'] = dst
                    else:
                        ok_all = False
                        results.append((name, False, m))
                finally:
                    # 第 24 次迭代：desktop.ini 按所选模式收尾
                    _desktop_ini_finalize(src, dst, dmode, backup)
        finally:
            try:
                self.q.put(('move_done', results, ok_all))
                self._wake_poll()
            except Exception:
                pass

    # 第 10 次迭代：转移用户（个人文件夹）整体到批量目标路径。
    # 只移动内容并更新 ProfileImagePath 指向，不写入 Known Folder。
    # 第 24 次迭代：新增 dmode 参数按所选模式处理 desktop.ini。
    def _move_user_folder(self, st, dmode='keep'):
        src = st.get('current_path') or ''
        dst = self.user_target_var.get().strip()
        if not dst:
            return False, '未设置用户文件夹目标'
        self.q.put(('log', '--- 正在处理「用户（个人文件夹）」---'))
        self._wake_poll()
        if src and os.path.normpath(src) == os.path.normpath(dst):
            self.q.put(('log', '「用户」已在目标位置，无需移动（%s）' % dst))
            self._wake_poll()
            st['current_path'] = dst
            return True, '用户文件夹已在目标位置'
        backup = _desktop_ini_prepare(src, dst, dmode)
        try:
            if self.move_files.get():
                if src and os.path.isdir(src):
                    self.q.put(('log', '「用户」移动文件：%s -> %s（排除 NTUSER.* 配置文件）' % (src, dst)))
                    self._wake_poll()
                    ok, m = robocopy_move(src, dst, exclude_profile=True)
                    if not ok:
                        self.q.put(('log', '「用户」文件移动失败：%s' % m))
                        self._wake_poll()
                        return False, m
                    self.q.put(('log', '「用户」%s' % m))
                    self._wake_poll()
                else:
                    self.q.put(('log', '「用户」源目录不存在或为空，跳过移动'))
                    self._wake_poll()
            else:
                self.q.put(('log', '「用户」已选择不转移文件，仅更新指向'))
                self._wake_poll()
            try:
                os.makedirs(dst, exist_ok=True)
            except OSError as e:
                self.q.put(('log', '「用户」创建目标目录失败：%s' % e))
                self._wake_poll()
                return False, '创建目标目录失败'
            ok, m = _set_profile_image_path(dst)
            self.q.put(('log', '「用户」%s' % m))
            self._wake_poll()
            st['current_path'] = dst
            return ok, m
        finally:
            # 第 24 次迭代：desktop.ini 按所选模式收尾
            _desktop_ini_finalize(src, dst, dmode, backup)

    # ---------------- 注销 ----------------
    # 第 23 次迭代：参照 pft.exe（DOS之家 个人文件转移工具）的 M_SysLogoff
    # 实现修复「勾选允许执行注销后仍不注销」。
    # pft 的注销做法：先 AdjustTokenPrivileges 启用 SE_SHUTDOWN_NAME（关机/
    # 注销特权），注销前有延时缓冲（kernel32 Sleep），再调用 user32
    # ExitWindowsEx(EWX_LOGOFF, 0) 并校验结果。原实现直接
    # ExitWindowsEx(0x00, 0) 且不检查返回值：特权缺失或调用失败时静默失败，
    # 用户勾选后看不到任何注销动作。此处对齐 pft：启用关机特权 → 确认后
    # 延时缓冲（让转移文件操作/句柄收尾）→ 调用注销 → 按返回值/错误码提示。
    def _enable_shutdown_privilege(self):
        """启用 SE_SHUTDOWN_NAME 关机/注销特权（ExitWindowsEx 成功前置条件）"""
        try:
            advapi32 = ctypes.windll.advapi32
            kernel32 = ctypes.windll.kernel32
            TOKEN_ADJUST_PRIVILEGES = 0x20
            TOKEN_QUERY = 0x08
            SE_PRIVILEGE_ENABLED = 0x2

            class LUID(ctypes.Structure):
                _fields_ = [('LowPart', wt.DWORD), ('HighPart', ctypes.c_long)]

            class LUID_AND_ATTRIBUTES(ctypes.Structure):
                _fields_ = [('Luid', LUID), ('Attributes', wt.DWORD)]

            class TOKEN_PRIVILEGES(ctypes.Structure):
                _fields_ = [('PrivilegeCount', wt.DWORD),
                            ('Privileges', LUID_AND_ATTRIBUTES * 1)]

            advapi32.OpenProcessToken.restype = wt.BOOL
            advapi32.OpenProcessToken.argtypes = [wt.HANDLE, wt.DWORD,
                                                  ctypes.POINTER(wt.HANDLE)]
            advapi32.LookupPrivilegeValueW.restype = wt.BOOL
            advapi32.LookupPrivilegeValueW.argtypes = [wt.LPCWSTR, wt.LPCWSTR,
                                                       ctypes.POINTER(LUID)]
            advapi32.AdjustTokenPrivileges.restype = wt.BOOL
            advapi32.AdjustTokenPrivileges.argtypes = [
                wt.HANDLE, wt.BOOL, ctypes.POINTER(TOKEN_PRIVILEGES),
                wt.DWORD, ctypes.c_void_p, ctypes.POINTER(wt.DWORD)]

            hToken = wt.HANDLE()
            if not advapi32.OpenProcessToken(kernel32.GetCurrentProcess(),
                                             TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY,
                                             ctypes.byref(hToken)):
                self.log('启用关机特权失败：OpenProcessToken 错误 %d' %
                         kernel32.GetLastError())
                return False
            try:
                tp = TOKEN_PRIVILEGES()
                tp.PrivilegeCount = 1
                tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED
                if not advapi32.LookupPrivilegeValueW(None, 'SeShutdownPrivilege',
                                                     ctypes.byref(tp.Privileges[0].Luid)):
                    self.log('启用关机特权失败：LookupPrivilegeValueW 错误 %d' %
                             kernel32.GetLastError())
                    return False
                if not advapi32.AdjustTokenPrivileges(hToken, False,
                                                      ctypes.byref(tp), 0, None, None):
                    self.log('启用关机特权失败：AdjustTokenPrivileges 错误 %d' %
                             kernel32.GetLastError())
                    return False
                err = kernel32.GetLastError()
                if err == 1300:  # ERROR_NOT_ALL_ASSIGNED
                    self.log('警告：当前账户令牌未授予关机特权，注销可能失败。')
                    return False
                return True
            finally:
                kernel32.CloseHandle(hToken)
        except Exception as e:
            self.log('启用关机特权异常：%s' % e)
            return False

    def _do_logoff(self):
        """实际执行注销（确认弹窗 + 延时缓冲后调用，校验结果）"""
        if not self._enable_shutdown_privilege():
            self.log('警告：关机特权未启用，注销调用可能失败。')
        self.log('正在注销当前账户...')
        try:
            user32 = ctypes.WinDLL('user32', use_last_error=True)
            ret = user32.ExitWindowsEx(0x00, 0)  # EWX_LOGOFF
            if ret:
                self.log('注销指令已发出，系统即将注销当前账户。')
            else:
                err = ctypes.get_last_error()
                self.log('注销调用失败：ExitWindowsEx 返回 0（错误码 %d）。'
                         '请确认以管理员身份运行并已勾选「允许执行注销」。' % err)
                messagebox.showerror('注销失败',
                                     'ExitWindowsEx 调用失败（错误码 %d）。\n'
                                     '请确认以管理员身份运行本程序，且当前账户允许注销。' % err,
                                     parent=self.root)
        except Exception as e:
            self.log('注销失败：%s' % e)
            messagebox.showerror('错误', '注销失败：%s' % e, parent=self.root)

    def _logoff(self):
        if not self.allow_logoff.get():
            messagebox.showwarning('提示', '请先在“设置”页勾选「允许执行注销」。',
                                   parent=self.root)
            return
        if not messagebox.askyesno(
                '注销确认',
                '即将注销当前账户。\n\n请先保存所有正在编辑的文档，确认后系统将立即注销。\n\n确认注销？',
                parent=self.root):
            return
        # 第 23 次迭代：确认后延时 3 秒再实际注销，给转移文件操作/句柄
        # 收尾缓冲（参照 pft.exe 注销前延时处理），避免注销过早导致未生效
        self.log('转移已完成，将在 3 秒后自动注销当前账户，请及时保存文档...')
        try:
            self.root.after(3000, self._do_logoff)
        except Exception as e:
            self.log('注销调度失败：%s' % e)
            self._do_logoff()

    # ---------------- 清理垃圾 ----------------
    def _clean(self):
        if not self.clean_garbage.get():
            messagebox.showwarning('提示', '请先在“设置”页勾选「启用清理垃圾」。',
                                   parent=self.root)
            return
        if not messagebox.askyesno(
                '清理垃圾确认',
                '即将清理以下内容：\n'
                '  1. 用户临时文件（%TEMP% 下所有内容）\n'
                '  2. 缩略图缓存（thumbcache_*.db）\n'
                '  3. 回收站（全部清空）\n\n'
                '正在使用的文件将被自动跳过。\n\n'
                '确认清理？',
                parent=self.root):
            return
        threading.Thread(target=self._clean_worker, daemon=True).start()

    def _clean_worker(self):
        try:
            self.q.put(('log', '开始清理系统垃圾...'))
            self._wake_poll()
            n1 = clean_temp_files()
            self.q.put(('log', '临时文件：清理 %d 项' % n1))
            self._wake_poll()
            n2 = clean_thumb_cache()
            self.q.put(('log', '缩略图缓存：清理 %d 个文件' % n2))
            self._wake_poll()
            ok, m = clean_recycle_bin()
            self.q.put(('log', '回收站：%s' % m))
            self._wake_poll()
            self.q.put(('log', '垃圾清理完成。'))
            self._wake_poll()
        except Exception:
            pass

    # ---------------- 配置保存 / 导入 ----------------
    # ---------------- 第 15 次迭代：当前位置快照保存 / 导入（.ini） ----------------
    def _save_snapshot(self):
        """把用户行与所有系统文件夹的当前位置写入 .ini 快照文件（[Folders] 名称=路径）。"""
        lines = ['; 系统文件夹位置快照（第 15 次迭代生成）', '[Folders]']
        us = self.user_state
        u_cur = (us.get('current_path') or '').strip()
        lines.append('用户=%s' % u_cur if u_cur else '用户=')
        for name, st in self.folder_states.items():
            cp = (st.get('current_path') or '').strip()
            lines.append('%s=%s' % (name, cp if cp else ''))
        path = filedialog.asksaveasfilename(
            title='保存当前位置快照', defaultextension='.ini',
            initialfile='文件夹位置快照.ini',
            initialdir=app_dir(),
            filetypes=[('INI 快照', '*.ini'), ('所有文件', '*.*')],
            parent=self.root)
        if not path:
            return False
        try:
            with open(path, 'w', encoding='utf-8-sig') as f:
                f.write('\n'.join(lines) + '\n')
            self.log('当前位置快照已保存：%s（用户 + %d 个文件夹）'
                     % (path, len(self.folder_states)))
            messagebox.showinfo('完成', '当前位置快照已保存。', parent=self.root)
            return True
        except Exception as e:
            messagebox.showerror('错误', '保存失败：%s' % e, parent=self.root)
            return False

    def _load_snapshot(self):
        """读取 .ini 快照：用户行目标 = 文件中「用户」位置，各文件夹目标 = 文件中对应位置，
        全部视为独立目标（不再跟随推导）。"""
        path = filedialog.askopenfilename(
            title='导入当前位置快照', initialdir=app_dir(),
            filetypes=[('INI 快照', '*.ini'), ('所有文件', '*.*')],
            parent=self.root)
        if not path:
            return
        # 解析 [Folders] 节：名称=路径
        targets = {}
        in_folders = False
        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                for raw in f:
                    line = raw.strip()
                    if not line or line.startswith(';') or line.startswith('#'):
                        continue
                    if line.startswith('[') and line.endswith(']'):
                        in_folders = (line.strip('[]').strip() == 'Folders')
                        continue
                    if not in_folders or '=' not in line:
                        continue
                    k, v = line.split('=', 1)
                    k = k.strip()
                    v = v.strip()
                    if k:
                        targets[k] = v
        except Exception as e:
            messagebox.showerror('错误', '读取快照失败：%s' % e, parent=self.root)
            return
        if not targets:
            messagebox.showerror('错误', '快照文件中未找到 [Folders] 节的有效条目。',
                                 parent=self.root)
            return
        # 各文件夹目标 = 文件中对应位置，全部视为独立目标（例外）
        matched = 0
        for name, st in self.folder_states.items():
            if name in targets:
                st['manual_set'] = True
                st['var_target'].set(targets[name])
                matched += 1
            else:
                st['manual_set'] = False
        # 用户行目标 = 文件中「用户」位置（设置后触发跟随重算，但 manual_set 项不被覆盖）
        if '用户' in targets:
            self.user_target_var.set(targets['用户'])
        # 刷新例外标记与用户勾选联动
        self._refresh_follow_states()
        self.log('当前位置快照已导入：%s（应用 %d 个文件夹 + 用户目标）' % (path, matched))
        messagebox.showinfo('完成', '当前位置快照已导入，各项目标已按文件位置设置。',
                            parent=self.root)

    # ---------------- 设置页 ----------------
    def _build_page_settings(self, parent):
        page = tk.Frame(parent, bg='#f7f7f7')

        head = tk.Frame(page, bg='#f7f7f7')
        head.pack(fill='x', padx=28, pady=(40, 6))
        tk.Label(head, text='设置', bg='#f7f7f7', fg=self.TEXT_MAIN,
                 font=(self.FONT, 18, 'bold')).pack(side='left')
        tk.Label(head, text='\uf013', bg='#f7f7f7', fg=self.TEXT_DIM,
                 font=('Font Awesome 6 Free Solid', 13)).pack(side='left', padx=(8, 0))
        tk.Button(head, text='关于', command=lambda: self._switch('关于'),
                  bg='#333333', fg='#ffffff', activebackground='#555555',
                  activeforeground='#ffffff', relief='flat', bd=0,
                  font=(self.FONT, 10, 'bold'), cursor='hand2',
                  padx=14, pady=2).pack(side='right')

        # 转移方式
        card1 = tk.Frame(page, bg=self.CARD_BG, highlightbackground=self.CARD_BORDER,
                         highlightthickness=1)
        card1.pack(fill='x', padx=28, pady=(6, 6))
        tk.Label(card1, text='转移方式', bg=self.CARD_BG, fg=self.TEXT_MAIN,
                 font=(self.FONT, 12)).pack(anchor='w', padx=24, pady=(10, 6))
        f1 = tk.Frame(card1, bg=self.CARD_BG)
        f1.pack(fill='x', padx=24, pady=(0, 4))
        tk.Radiobutton(f1, text='转移文件', variable=self.move_files, value=True,
                       bg=self.CARD_BG, activebackground=self.CARD_BG,
                       highlightthickness=0, bd=0, cursor='hand2',
                       font=(self.FONT, 10), fg=self.TEXT_MAIN,
                       selectcolor='#ffffff').pack(anchor='w')
        tk.Label(card1, text='将现有文件移动到新目标目录后，更新系统文件夹指向（推荐）。',
                 bg=self.CARD_BG, fg=self.TEXT_DIM, font=(self.FONT, 9)
                 ).pack(anchor='w', padx=40)
        f2 = tk.Frame(card1, bg=self.CARD_BG)
        f2.pack(fill='x', padx=24, pady=(6, 4))
        tk.Radiobutton(f2, text='不转移文件', variable=self.move_files, value=False,
                       bg=self.CARD_BG, activebackground=self.CARD_BG,
                       highlightthickness=0, bd=0, cursor='hand2',
                       font=(self.FONT, 10), fg=self.TEXT_MAIN,
                       selectcolor='#ffffff').pack(anchor='w')
        tk.Label(card1, text='仅更新系统文件夹指向，原文件保留在当前位置（需手动迁移）。',
                 bg=self.CARD_BG, fg=self.TEXT_DIM, font=(self.FONT, 9)
                 ).pack(anchor='w', padx=40, pady=(0, 10))

        # 系统操作
        card2 = tk.Frame(page, bg=self.CARD_BG, highlightbackground=self.CARD_BORDER,
                         highlightthickness=1)
        card2.pack(fill='x', padx=28, pady=(6, 6))
        tk.Label(card2, text='系统操作', bg=self.CARD_BG, fg=self.TEXT_MAIN,
                 font=(self.FONT, 12)).pack(anchor='w', padx=24, pady=(10, 8))
        row2a = tk.Frame(card2, bg=self.CARD_BG)
        row2a.pack(fill='x', padx=24, pady=(0, 6))
        tk.Checkbutton(row2a, text='允许执行注销', variable=self.allow_logoff,
                       bg=self.CARD_BG, activebackground=self.CARD_BG,
                       highlightthickness=0, bd=0, cursor='hand2',
                       font=(self.FONT, 10), fg=self.TEXT_MAIN,
                       selectcolor='#ffffff').pack(side='left')
        self.btn_logoff = tk.Button(row2a, text='注销当前账户', command=self._logoff,
                                    bg=self.DANGER, fg='#ffffff', relief='flat', bd=0,
                                    activebackground='#e04a4a',
                                    activeforeground='#ffffff',
                                    font=(self.FONT, 10, 'bold'), cursor='hand2',
                                    padx=14, pady=2)
        self.btn_logoff.pack(side='left', padx=(18, 0))
        tk.Label(row2a, text='注销后文件夹指向将全部生效', bg=self.CARD_BG,
                 fg=self.TEXT_DIM, font=(self.FONT, 9)).pack(side='left', padx=(10, 0))
        row2b = tk.Frame(card2, bg=self.CARD_BG)
        row2b.pack(fill='x', padx=24, pady=(0, 10))
        tk.Checkbutton(row2b, text='启用清理垃圾', variable=self.clean_garbage,
                       bg=self.CARD_BG, activebackground=self.CARD_BG,
                       highlightthickness=0, bd=0, cursor='hand2',
                       font=(self.FONT, 10), fg=self.TEXT_MAIN,
                       selectcolor='#ffffff').pack(side='left')
        self.btn_clean = tk.Button(row2b, text='清理系统垃圾', command=self._clean,
                                   bg=self.WARN, fg='#ffffff', relief='flat', bd=0,
                                   activebackground='#e0a434',
                                   activeforeground='#ffffff',
                                   font=(self.FONT, 10, 'bold'), cursor='hand2',
                                   padx=14, pady=2)
        self.btn_clean.pack(side='left', padx=(18, 0))
        tk.Label(row2b, text='清理用户临时文件、缩略图缓存与回收站',
                 bg=self.CARD_BG, fg=self.TEXT_DIM,
                 font=(self.FONT, 9)).pack(side='left', padx=(10, 0))

        # 第 15 次迭代起：设置页不再提供配置保存/导入入口
        # （保存/导入已移到用户行：黑「选择...」右侧并排「保存位置」「导入位置」）
        return page

    # ---------------- 关于页 ----------------
    def _build_page_about(self, parent):
        page = tk.Frame(parent, bg='#f7f7f7')

        head = tk.Frame(page, bg='#f7f7f7')
        head.pack(fill='x', padx=28, pady=(40, 6))
        tk.Label(head, text='关于', bg='#f7f7f7', fg=self.TEXT_MAIN,
                 font=(self.FONT, 18, 'bold')).pack(side='left')
        tk.Button(head, text='返回', command=lambda: self._switch('设置'),
                  bg='#333333', fg='#ffffff', activebackground='#555555',
                  activeforeground='#ffffff', relief='flat', bd=0,
                  font=(self.FONT, 10, 'bold'), cursor='hand2',
                  padx=14, pady=2).pack(side='right')
        # 第 17 次迭代：关于页展示软件版本号
        tk.Label(page, text='系统文件夹转移  v1.7', bg='#f7f7f7', fg='#666666',
                 font=(self.FONT, 11, 'bold')).pack(anchor='w', padx=28, pady=(10, 0))

        card = tk.Frame(page, bg='#fbfbfb', highlightbackground=self.CARD_BORDER,
                        highlightthickness=1)
        card.pack(fill='x', padx=28, pady=(14, 10))
        lines = [
            '· 系统文件夹：SHGetKnownFolderPath 枚举 Windows 已知文件夹，三页签分组（重要 / 次要 / 其他）快速定位',
            '· 转移文件：robocopy 移动 + SHSetKnownFolderPath 更新指向；不转移文件：仅更新指向',
            '· 用户批量目标：未单独设置的文件夹自动跟随用户目标生成同名子目录，单独设置目标即为例外',
            '· 当前位置快照：.ini 一键保存 / 导入各文件夹当前位置，随时恢复或迁移到新电脑',
            '· 当前地址点击：列表内点击当前路径即可用资源管理器打开，hover 平滑加粗提示',
            '· 大小四色分级：实时统计每个系统文件夹占用空间，按大小四色分级显示',
            '· 注销：ExitWindowsEx 注销使指向全部生效；清理垃圾：清临时文件、缩略图缓存与回收站',
            '· 日志：右侧栏记录每次转移与清理命令及结果，可随时展开查看',
        ]
        for i, ln in enumerate(lines):
            tk.Label(card, text=ln, bg='#fbfbfb', fg='#888888',
                     font=(self.FONT, 9), anchor='w', justify='left').pack(fill='x', pady=2)

        return page

    # ---------------- 日志 ----------------
    def log(self, msg):
        try:
            self.log_text.configure(state='normal')
            self.log_text.insert('end', msg + '\n')
            self.log_text.see('end')
            self.log_text.configure(state='disabled')
        except Exception:
            pass

    def _set_log_btn_state(self, active):
        try:
            self._log_icon_btn.configure(
                bg='#f7f7f7' if active else '#f7f7f7',
                text='\u2347' if active else '\u269d')
        except Exception:
            pass

    def _toggle_log_panel(self):
        if self._log_anim:
            return
        self._log_anim = True
        if self._log_visible:
            self._log_target = 0
        else:
            self._log_visible = True
            self._log_cur = 0
            self.log_panel.place(relx=1.0, rely=0.0, anchor='ne',
                                 relheight=1.0, width=0)
            self.titlebar_left.lift()
            self.titlebar_mid.lift()
            self.titlebar_right.lift()
            self._set_log_btn_state(True)
            self._log_target = LOG_PANEL_W
        self._log_step()

    def _log_step(self):
        # 与参考软件保持相同总时长（约 132ms：参考 300/28*12ms，本版 300/50*22ms），
        # 但单帧步长加大、帧数减半（11 帧 -> 6 帧），大幅降低 content 逐帧重排开销，
        # 消除右侧栏展开/收起时的卡顿，观感仍与参考软件一致。
        step = 50
        if self._log_target > self._log_cur:
            self._log_cur = min(self._log_cur + step, self._log_target)
        else:
            self._log_cur = max(self._log_cur - step, 0)
        self.log_panel.place(relx=1.0, rely=0.0, anchor='ne',
                             relheight=1.0, width=self._log_cur)
        self.content.place_configure(x=196, y=0, relheight=1, relwidth=1,
                                     width=-(196 + self._log_cur), height=0)
        if self._log_cur != self._log_target:
            self.root.after(22, self._log_step)
        else:
            self._log_anim = False
            if self._log_target == 0:
                self.log_panel.place_forget()
                self.content.place_configure(x=196, y=0, relwidth=1,
                                             relheight=1, width=-196, height=0)
                self._log_visible = False
                self._set_log_btn_state(False)

    def _log_sb_kick(self, event=None, release=False):
        sb = getattr(self, '_log_sb', None)
        if sb is None:
            return
        try:
            if not sb.winfo_ismapped():
                head = getattr(self, 'log_panel', None).winfo_children()[1]
                y0 = head.winfo_y() + head.winfo_height() + 4
                h0 = self.log_panel.winfo_height() - y0 - 8
                if h0 < 40:
                    h0 = 40
                sb.place(relx=1.0, rely=0.0, anchor='ne',
                         y=y0, height=h0, width=10)
                self.root.after_idle(self._log_sb_refresh)
        except Exception:
            return
        delay = 1200 if release else 2000
        if self._log_sb_hide_after is not None:
            try:
                self.root.after_cancel(self._log_sb_hide_after)
            except Exception:
                pass
            self._log_sb_hide_after = None
        self._log_sb_hide_after = self.root.after(delay, self._log_sb_hide)

    def _log_sb_update(self, first, last):
        self._log_sb_refresh()

    def _log_sb_refresh(self):
        sb = getattr(self, '_log_sb', None)
        if sb is None or not sb.winfo_ismapped():
            return
        try:
            first, last = self.log_text.yview()
            if first == 0 and last >= 0.999:
                # 内容不足一屏/无需滚动：不显示任何滑动元素
                if self._log_sb_hide_after is not None:
                    try:
                        self.root.after_cancel(self._log_sb_hide_after)
                    except Exception:
                        pass
                    self._log_sb_hide_after = None
                try:
                    sb.place_forget()
                except Exception:
                    pass
                return
            h = sb.winfo_height()
            frac = max(0.0, min(1.0, last - first))
            thumb_h = max(20, int(h * frac))
            y0 = int(first * (h - thumb_h)) if h > thumb_h else 0
            w = sb.winfo_width()
            sb.itemconfigure(self._log_sb_item, state='normal')
            sb.coords(self._log_sb_geo, 0, y0, w, y0 + thumb_h)
            sb.coords(self._log_sb_item,
                      *_round_rect_pts(0, y0, w, y0 + thumb_h, 3))
        except Exception:
            pass

    def _log_sb_press(self, event):
        self._log_sb_kick(event)
        try:
            c = self._log_sb.coords(self._log_sb_geo)
            self._log_sb_drag_off = event.y - c[1]
        except Exception:
            self._log_sb_drag_off = 0

    def _log_sb_drag(self, event):
        self._log_sb_kick(event)
        try:
            h = self._log_sb.winfo_height()
            c = self._log_sb.coords(self._log_sb_geo)
            thumb_h = c[3] - c[1]
            track = h - thumb_h
            if track <= 0:
                return
            frac = min(1.0, max(0.0,
                                (event.y - self._log_sb_drag_off) / track))
            self.log_text.yview_moveto(frac)
        except Exception:
            pass

    def _log_sb_hide(self):
        self._log_sb_hide_after = None
        sb = getattr(self, '_log_sb', None)
        if sb is None:
            return
        try:
            if sb.winfo_ismapped():
                sb.place_forget()
        except Exception:
            pass

    # ---------------- 列表滚动条 ----------------
    def _on_list_wheel(self, event):
        try:
            delta = -1 if event.delta > 0 else 1
            self.list_canvas.yview_scroll(delta * int(abs(event.delta) / 120 or 1),
                                          'units')
        except Exception:
            pass

    def _on_list_wheel_global(self, event):
        try:
            if self.current_page != '程序':
                return
            wrap = getattr(self, 'list_wrap', None)
            if wrap is None or not wrap.winfo_ismapped():
                return
            x0 = wrap.winfo_rootx()
            y0 = wrap.winfo_rooty()
            x1 = x0 + wrap.winfo_width()
            y1 = y0 + wrap.winfo_height()
            if x0 <= event.x_root <= x1 and y0 <= event.y_root <= y1:
                self._on_list_wheel(event)
                self._list_sb_kick()
        except Exception:
            pass

    def _list_sb_kick(self):
        try:
            c = self.list_canvas
            bbox = c.bbox('all')
            if not bbox:
                return
            first, last = c.yview()
            total = max(1, bbox[3] - bbox[1])
            view = c.winfo_height()
            if view < 4 or total <= view:
                # 内容无需滚动：不显示任何滑动元素
                self._list_sb_hide_now()
                return
            # thumb 高度按内容比例：view * (view / total)
            thumb_h = max(20, int(view * view / total))
            if thumb_h >= view - 2:
                # 内容仅略高于可视区时 thumb 近乎满高，收敛为可视区高度
                thumb_h = view - 2
            track = view - thumb_h
            denom = total - view  # 内容可滚动总偏移
            # first 为可视区顶部相对 scrollregion 的比例（0..(total-view)/total）。
            # 滑块顶部 = first*total/denom*track，保证滚到底时滑块正好贴底。
            if denom > 0:
                y0 = int(first * total / denom * track)
            else:
                y0 = 0
            y0 = max(0, min(y0, track))
            w = self.list_scroll.winfo_width()
            self.list_scroll.itemconfigure(self._list_sb_item, state='normal')
            self.list_scroll.coords(self._list_sb_geo, 0, y0, w, y0 + thumb_h)
            self.list_scroll.coords(self._list_sb_item,
                                    *_round_rect_pts(0, y0, w, y0 + thumb_h, 3))
            # 停止滚动片刻后自动隐藏
            if self._list_sb_hide_after is not None:
                try:
                    self.list_scroll.after_cancel(self._list_sb_hide_after)
                except Exception:
                    pass
            self._list_sb_hide_after = self.list_scroll.after(
                800, self._list_sb_hide_now)
        except Exception:
            pass

    def _list_sb_hide_now(self):
        try:
            if self._list_sb_hide_after is not None:
                try:
                    self.list_scroll.after_cancel(self._list_sb_hide_after)
                except Exception:
                    pass
                self._list_sb_hide_after = None
            self.list_scroll.itemconfigure(self._list_sb_item, state='hidden')
        except Exception:
            pass

    def _list_sb_press(self, event):
        try:
            if self._list_sb_hide_after is not None:
                try:
                    self.list_scroll.after_cancel(self._list_sb_hide_after)
                except Exception:
                    pass
                self._list_sb_hide_after = None
            self.list_scroll.itemconfigure(self._list_sb_item, state='normal')
        except Exception:
            pass
        try:
            c = self.list_scroll.coords(self._list_sb_geo)
            self._list_sb_drag_off = event.y - c[1]
        except Exception:
            self._list_sb_drag_off = 0

    def _list_sb_drag(self, event):
        try:
            h = self.list_scroll.winfo_height()
            c = self.list_scroll.coords(self._list_sb_geo)
            thumb_h = c[3] - c[1]
            track = h - thumb_h
            if track <= 0:
                return
            frac = min(1.0, max(0.0,
                                (event.y - self._list_sb_drag_off) / track))
            self.list_canvas.yview_moveto(frac)
            self._list_sb_kick()
        except Exception:
            pass

    # ---------------- 队列消费 ----------------
    def _poll_queue(self):
        if self._drag_hold:
            self._poll_active = False
            return
        if self._poll_active:
            return
        self._poll_active = True
        try:
            while True:
                try:
                    item = self.q.get_nowait()
                except queue.Empty:
                    break
                if item[0] == 'log':
                    self.log(item[1])
                elif item[0] == 'size':
                    _, name, n = item
                    st = self.folder_states.get(name)
                    if st is None and name == '用户':
                        st = self.user_state
                    if st:
                        st['size_bytes'] = n
                        st['size_text'] = fmt_size(n) if n >= 0 else '统计失败'
                        lbl = st.get('size_lbl')
                        if lbl is not None:
                            try:
                                lbl.configure(text=st['size_text'],
                                               fg=size_color(st['size_text']))
                            except Exception:
                                pass
                elif item[0] == 'move_done':
                    _, results, ok_all = item
                    self._switch_running = False
                    for name, ok, m in results:
                        self.log('「%s」%s' % (name, '成功' if ok else '失败'))
                    if ok_all:
                        self.log('全部转移完成。建议注销账户使设置全部生效。')
                    else:
                        self.log('转移部分失败，请查看日志。')
                    self._refresh_sizes()
                    # 第 22 次迭代：修复勾选「允许执行注销」后转移完成不注销。
                    # 原实现 move_done 收尾仅重置状态/输出日志/刷新大小，
                    # 从未检查 allow_logoff 并调用 _logoff，导致勾选后不注销。
                    # 此处位于主线程（_poll_queue 由 after 调度执行），
                    # 转移线程（_move_worker）已全部结束并回传结果，是注销的正确时机。
                    if self.allow_logoff.get():
                        self.log('检测到已勾选「允许执行注销」，转移已完成，准备注销账户...')
                        self._logoff()
        finally:
            self._poll_active = False
        if not self.q.empty():
            self.root.after(0, self._poll_queue)

    def _wake_poll(self):
        # 线程安全：非主线程禁止操作 Tk（after/event_generate 非线程安全，
        # 在后台线程调用会静默失效，导致统计结果积压不刷新——第9次实测发现）。
        # 主线程定时器 _ensure_sizes_shown 会轮询队列兜底刷新。
        try:
            if threading.current_thread() is not threading.main_thread():
                return
        except Exception:
            return
        try:
            self.root.after(0, self._poll_queue)
        except Exception:
            try:
                self.root.event_generate('<<QueuePoll>>', when='tail')
            except Exception:
                pass

    def _refresh_sizes(self):
        for name, st in self.folder_states.items():
            cur = st['current_path']
            if cur and os.path.isdir(cur):
                threading.Thread(target=self._size_worker, args=(name, cur),
                                 daemon=True).start()
        us = self.user_state
        if us['current_path'] and os.path.isdir(us['current_path']):
            threading.Thread(target=self._size_worker,
                             args=('用户', us['current_path']),
                             daemon=True).start()

    def _ensure_sizes_shown(self):
        """定时兜底：先轮询队列（主线程安全），再强制刷新大小显示。"""
        try:
            self._poll_queue()
        except Exception:
            pass
        try:
            for name, st in self.folder_states.items():
                if st['size_bytes'] >= 0:
                    txt = fmt_size(st['size_bytes'])
                elif st['size_bytes'] == -1:
                    txt = '统计失败'
                else:
                    continue
                if st['size_text'] != txt:
                    st['size_text'] = txt
                    lbl = st.get('size_lbl')
                    if lbl is not None:
                        try:
                            lbl.configure(text=txt, fg=size_color(txt))
                        except Exception:
                            pass
        except Exception:
            pass
        # 用户（个人文件夹）行兜底
        try:
            us = self.user_state
            if us['size_bytes'] >= 0:
                txt = fmt_size(us['size_bytes'])
            elif us['size_bytes'] == -1:
                txt = '统计失败'
            else:
                txt = None
            if txt and us['size_text'] != txt:
                us['size_text'] = txt
                lbl = us.get('size_lbl')
                if lbl is not None:
                    try:
                        lbl.configure(text=txt, fg=size_color(txt))
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self.root.after(1000, self._ensure_sizes_shown)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def main():
    _register_fonts()
    if not _acquire_single_instance():
        _activate_existing_window()
        return
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
