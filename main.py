# -*- coding: utf-8 -*-
# Author: Wei Jia
# Project: play_ground
# File: main.py
try:
    import FreeSimpleGUI as sg
except ImportError:
    import PySimpleGUI as sg
import subprocess
import platform
import ctypes
import traceback
import base64

VERSION = "Alpha 1.10"
ICON = """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"""  # noqa

# 加载 kernel32.dll
kernel32 = ctypes.windll.kernel32

sg.set_global_icon(base64.b64decode(ICON))


def bytes_to_mb(_bytes):
    return _bytes / (1024 * 1024)


def get_bits():
    res_str = "Faild to get"
    if platform.architecture()[0] == '32bit':
        res_str = "Current system is 32-bit, nolonger to get whether if 64-bit process"
    else:
        # 定义函数指针
        IsWow64Process = kernel32.IsWow64Process

        # 定义函数参数和返回类型
        IsWow64Process.restype = ctypes.c_bool
        IsWow64Process.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_bool)]

        # 用于存储检测结果
        is_wow64 = ctypes.c_bool()

        # 获取当前进程的句柄
        current_process = ctypes.windll.kernel32.GetCurrentProcess()

        # 调用函数进行检测
        if IsWow64Process(current_process, ctypes.byref(is_wow64)):
            if is_wow64.value:
                res_str = "Current process is 32-bit process running on 64-bit system"
            else:
                res_str = "Current process is 64-bit process"

    sg.PopupOK(res_str, title="System Arch")


def get_memstatus():
    # 定义 MEMORYSTATUSEX 结构体
    class MEMORYSTATUSEX(ctypes.Structure):
        _fields_ = [
            ('dwLength', ctypes.c_ulong),
            ('dwMemoryLoad', ctypes.c_ulong),
            ('ullTotalPhys', ctypes.c_ulonglong),
            ('ullAvailPhys', ctypes.c_ulonglong),
            ('ullTotalPageFile', ctypes.c_ulonglong),
            ('ullAvailPageFile', ctypes.c_ulonglong),
            ('ullTotalVirtual', ctypes.c_ulonglong),
            ('ullAvailVirtual', ctypes.c_ulonglong),
            ('ullAvailExtendedVirtual', ctypes.c_ulonglong)
        ]

    # 创建 MEMORYSTATUSEX 结构体实例
    memory_status = MEMORYSTATUSEX()
    memory_status.dwLength = ctypes.sizeof(MEMORYSTATUSEX)

    # 调用 GlobalMemoryStatusEx 函数获取内存状态信息
    kernel32.GlobalMemoryStatusEx(ctypes.byref(memory_status))

    # 打印内存信息
    total_physical_memory = memory_status.ullTotalPhys
    available_physical_memory = memory_status.ullAvailPhys
    mem_info_str = "Total Pysical Memory: {}MB\nAvailable Physical Memory: {}MB".format(
        int(bytes_to_mb(total_physical_memory)),
        int(bytes_to_mb(available_physical_memory)))
    sg.PopupOK(mem_info_str, title="Mem Status")


def get_possible_screen_resolutions():
    try:
        import win32api
        import win32con
        import pywintypes

        res = []
        i = 0
        try:
            while True:
                ds = win32api.EnumDisplaySettings(None, i)
                _reso = (ds.PelsWidth, ds.PelsHeight)
                if _reso not in res:
                    res.append(_reso)
                i += 1
        except Exception:
            pass
        res.sort(key=lambda x: (x[0], x[1]))
        return res
    except Exception:
        return None


def get_current_system_resolution():
    import win32api
    old_width = win32api.GetSystemMetrics(0)
    old_height = win32api.GetSystemMetrics(1)
    return old_width, old_height


def change_system_resolution(new_width, new_height, recover=False):
    try:
        import win32api
        import win32con
        import pywintypes

        old_width, old_height = get_current_system_resolution()

        devmode = pywintypes.DEVMODEType()

        devmode.PelsWidth = new_width
        devmode.PelsHeight = new_height

        devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT

        res = win32api.ChangeDisplaySettings(devmode, 0)

        if recover:
            return None
        return old_width, old_height, res
    except Exception:
        sg.PopupError('Change System Resolution Failed', title='Error')


def get_installed_software_ex():
    import winreg

    software_list = []
    _name = []
    sub_keys = [r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall']
    if platform.architecture()[0] == '64bit':
        sub_keys.append(r'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall')
    for sub_key in sub_keys:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub_key, 0, winreg.KEY_ALL_ACCESS)
        for i in range(0, winreg.QueryInfoKey(key)[0] - 1):
            try:
                key_name = winreg.EnumKey(key, i)
                key_path = sub_key + '\\' + key_name
                each_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_ALL_ACCESS)
                display_name = winreg.QueryValueEx(each_key, 'DisplayName')[0]
                if display_name:
                    try:
                        publisher = winreg.QueryValueEx(each_key, 'Publisher')[0]
                    except Exception:
                        publisher = ""
                    try:
                        install_date = winreg.QueryValueEx(each_key, 'InstallDate')[0]
                    except Exception:
                        install_date = ""
                    try:
                        display_version = winreg.QueryValueEx(each_key, 'DisplayVersion')[0]
                    except Exception:
                        display_version = ""
                    try:
                        estimated_size = winreg.QueryValueEx(each_key, 'EstimatedSize')[0]
                    except Exception:
                        estimated_size = ""
                    try:
                        uninstall_string = winreg.QueryValueEx(each_key, 'UninstallString')[0]
                    except Exception:
                        uninstall_string = ""
                    try:
                        install_location = winreg.QueryValueEx(each_key, "InstallLocation")[0]
                    except Exception:
                        install_location = ""
                    if display_name and display_name not in _name:
                        software_list.append({
                            "name": display_name,
                            "publisher": publisher,
                            "date": install_date,
                            "version": display_version,
                            "size": estimated_size,
                            "uninstall": uninstall_string,
                            "location": install_location,
                        })
                        _name.append(display_name)
            except WindowsError:
                pass
    return software_list


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def search_and_run_from_winreg(command: str, _winreg_res: dict):
    _command = command.lower()
    if _command in _winreg_res:
        subprocess.Popen(_winreg_res[_command])
    elif _command + '.exe' in _winreg_res:
        subprocess.Popen(_winreg_res[_command + '.exe'])
    elif _command + '.msc' in _winreg_res:
        subprocess.Popen(_winreg_res[_command + '.msc'])
    else:
        subprocess.Popen(_command)


def main_window():
    sg.change_look_and_feel('Gray Gray Gray')
    _os_name = platform.system()
    if _os_name != 'Windows':
        sg.Window('Not supprot', layout=[[sg.Text('This program only support Windows operating system!')],
                                         [sg.Button('Cancel', sg.BUTTON_TYPE_CLOSES_WIN)]]).read()
        return None

    import winreg

    _app_paths_key = winreg.OpenKeyEx(winreg.HKEY_LOCAL_MACHINE,
                                      r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths")
    _idx = 0
    _apps = {}
    while True:
        try:
            app_name = winreg.EnumKey(_app_paths_key, _idx)
            sub_key = winreg.OpenKeyEx(_app_paths_key, app_name)
            try:
                app_path = winreg.QueryValueEx(sub_key, '')
            except OSError:
                pass
            else:
                _apps[app_name.lower()] = app_path[0]
            _idx += 1
        except OSError:
            break

    _os_family = platform.platform()
    _os_version_specific = platform.win32_ver()
    user_admin = is_admin()

    if user_admin:
        _ua_info = {'text': "This program is running under Administrator", 'text_color': "green"}
    else:
        sg.PopupOK('This Windows Settings Center Is Not Runing Under Administrator, Some Function Will Disabled, '
                   'Please Run This Software As Administrator',
                   title='Warning', button_color='yellow')
        _ua_info = {'text': "This program is not running under Administrator", 'text_color': "red"}

    _os_ver = platform.version()
    _os_machine = platform.machine()
    _os_processor = platform.processor()
    sg.set_options(font=('Helvetica', 12))

    menu_def = [['File', ['Exit']],
                ['Help', ['About']]]
    layout = [
        [sg.Menu(menu_def)],
        [sg.Text('Welcome to the Windows Settings Center', justification='center', expand_x=True,
                 font=('Calibri bold', 20))],
        [sg.Text(**_ua_info)],
        [sg.Text("OS Name: {} {} {} {}".format(_os_name, *_os_version_specific[:3]))],
        [sg.Text("OS Version: " + _os_ver)],
        [sg.Text("OS Arch: " + _os_machine)],
        [sg.Text("OS Processor: " + _os_processor)],
        [sg.Frame(title='Windows Tools', layout=[[sg.Button('Windows Version', key='winver',
                                                            button_color='white on blue'),
                                                  sg.Button('Registry Editor', key='regedit',
                                                            disabled=False if user_admin else True),
                                                  sg.Button('Control Panel', key='control')],
                                                 [sg.Button('Command Prompt',
                                                            button_color='white on black',
                                                            key='cmd'),
                                                  sg.Button('Get Mem Status', key='_mem'),
                                                  sg.Button('Display Settings',
                                                            key='control desk.cpl')],
                                                 [sg.Button('Git Bits', key='_bits'),
                                                  sg.Button('On-Screen Keyboard', key='osk',
                                                            disabled=False if user_admin else True),
                                                  sg.Button('System Infomation', key='msinfo32')],
                                                 [sg.Button('System Configuration',
                                                            key='msconfig', disabled=False if user_admin else True),
                                                  sg.Button('Internet Properties',
                                                            key='control inetcpl.cpl'),
                                                  sg.Button('System Properties', key='control sysdm.cpl')
                                                  ],
                                                 [sg.Button('Certification Manager', key='certmgr.msc',
                                                            button_color='black on orange')]],
                  expand_x=True)],
        [sg.Frame(title='Featured Function', layout=[[sg.Button('Change Resolution Directly',
                                                                button_color='black on yellow', key='_cr'),
                                                      sg.Button('Software Uninstall Tool', key='_smt',
                                                                button_color='black on yellow',
                                                                disabled=False if user_admin else True)]],
                  expand_x=True)],
        [sg.Button('Exit', button_color='white on red', key='_exit')]]

    window = sg.Window('Windows Settings Center', layout=layout, resizable=False, auto_size_buttons=True,
                       finalize=True)
    window.force_focus()
    while True:
        event, values = window.read()
        try:
            if event in [sg.WIN_CLOSED, '_exit', 'Exit']:
                break
            elif event == 'About':
                sg.PopupOK(
                    'Windows Settings Center\nVersion: {}\n2024-2025 © Wei Jia All Rights Reserved.'.format(VERSION),
                    title='About')
            elif event == 'winver':
                subprocess.Popen(event)
            elif event == 'regedit':
                subprocess.Popen(event)
            elif event == 'control':
                subprocess.Popen(event)
            elif event == 'cmd':
                subprocess.Popen(event, shell=False, creationflags=subprocess.CREATE_NEW_CONSOLE)
            elif event == '_mem':
                get_memstatus()
            elif event == '_sys_info_str':
                sg.PopupOK('Not Implement Function', title='Warning')
            elif event == '_bits':
                get_bits()
            elif event == 'osk':
                subprocess.Popen(event, shell=True)
            elif event == 'msinfo32':
                search_and_run_from_winreg(event, _apps)
            elif event == 'msconfig':
                search_and_run_from_winreg(event, _apps)
            elif event == 'control inetcpl.cpl':
                subprocess.Popen(event)
            elif event == 'control desk.cpl':
                subprocess.Popen(event)
            elif event == 'control sysdm.cpl':
                subprocess.Popen(event)
            elif event == 'certmgr.msc':
                subprocess.Popen('certmgr.msc', shell=True)
            elif event == '_cr':
                rlst = get_possible_screen_resolutions()
                rlst_str = ["{}*{}".format(i[0], i[1]) for i in rlst]
                sub_window = sg.Window(title='Set',
                                       layout=[
                                           [sg.Combo(rlst_str, key='_rlst',
                                                     default_value="*".join(
                                                         [str(_i) for _i in get_current_system_resolution()]),
                                                     readonly=True,
                                                     expand_x=True)],
                                           [sg.Button('Confirm', key='_sco', expand_x=True),
                                            sg.Button('Cancel', key='_sca', expand_x=True)]],
                                       size=(200, 75), modal=True)
                while True:
                    _e, _v = sub_window.read()
                    if _e in [sg.WIN_CLOSED, '_sca']:
                        break
                    else:
                        if _e == '_sco':
                            _oldw, _oldh, _res = change_system_resolution(
                                *list(map(lambda x: int(x), _v['_rlst'].split('*'))))
                            if _res == 0:
                                _c2 = sg.PopupOKCancel('Do You Want To Keep This Resolution Setting?', title='Confirm')
                                if _c2 == 'Cancel':
                                    change_system_resolution(_oldw, _oldh, recover=True)
                                else:
                                    break
                            else:
                                sg.PopupCancel('Change Resolution Failed')
                sub_window.close()
            elif event == '_smt':
                def size2mb(s):
                    if s == '':
                        return ''
                    else:
                        return "{} MB".format(str(round(int(s) / 1024, 2)))

                local_software = get_installed_software_ex()
                _smt_layout = []
                _smt_table_values = [
                    [i['name'], i['date'], i['location'], i['publisher'], i['version'], size2mb(i['size'])] for
                    i in local_software]
                _smt_layout.append(
                    [sg.Table(headings=["Name", "Install Date", "Install Location", "Publisher", "Version", "Size"],
                              values=_smt_table_values, vertical_scroll_only=False, justification='left',
                              enable_click_events=True, key='_smt_table', num_rows=20)])
                _smt_layout.append([sg.Text("Detail: "), sg.Text("", key='_det')])

                _smt_layout.append(
                    [sg.Button('Exit', key='_esc'), sg.Button('Uninstall', key='_uni', button_color='white on red',
                                                              disabled=True)])
                smt_window = sg.Window(title='Software Management Tool', layout=_smt_layout,
                                       return_keyboard_events=True, modal=True)
                while True:
                    sube, subv = smt_window.read()
                    if sube in [sg.WIN_CLOSED, '_esc']:
                        break
                    elif isinstance(sube, tuple):
                        if sube[0] == '_smt_table':
                            if sube[-1][0] is None or sube[-1][0] == -1:
                                pass
                            else:
                                _uninstall = local_software[sube[-1][0]]['uninstall']
                                if _uninstall:
                                    smt_window['_uni'].update(disabled=False)
                                else:
                                    smt_window['_uni'].update(disabled=True)
                                smt_window['_det'].update(value=_smt_table_values[sube[-1][0]][sube[-1][1]])
                    elif sube == '_uni':
                        _uninstall = local_software[subv['_smt_table'][0]]['uninstall']
                        _softname = local_software[subv['_smt_table'][0]]['name']
                        _uc = sg.PopupYesNo("Are You Sure You Want To Uninstall {} ?".format(_softname),
                                            title='Uninstall')
                        if _uc == 'Yes':
                            subprocess.Popen(_uninstall)
                            break
                smt_window.close()



        except Exception as e:
            if isinstance(e, FileNotFoundError):
                sg.PopupError('Enconter An Error:\n{}'.format(e),
                              'The Current Version Of Windows May Not Support This Function', title='Error')
            else:
                traceback.print_exc()
                sg.PopupError('Enconter An Error:\n{}'.format(e),
                              'This Error May Be An Internal Error, Please Contact The Devoloper')
                break
    window.close()


if __name__ == '__main__':
    main_window()
