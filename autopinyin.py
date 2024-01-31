import uiautomation as auto
import pyautogui
import pinyin
import unicodedata
import time
import re


class InputIndicatorNotFoundError(Exception):
    """Raised when the input indicator is not found"""
    pass

class TaskbarNotFoundError(Exception):
    """Raised when the taskbar is not found"""
    pass

class CandidatePanelNotFoundError(Exception):
    """Raised when the candidate panel is not found"""
    pass


class AutoPinyin(object):

    def __init__(self, ui_respond_time=0.08, type_interval=0.001) -> None:
        super().__init__()
        
        self.input_indicator = None
        self.input_experience = None
        self.candidate_ui = None
        self.candidate_panel = None
        self.previous_page_button = None
        self.next_page_button = None

        self.ui_respond_time = ui_respond_time
        self.type_interval = type_interval
    

    def find_input_indicator(self) -> None:

        if self.input_indicator is not None:
            return

        root = auto.GetRootControl()
        taskbar = None
        for control in root.GetChildren():
            if control.Name == '任务栏':
                taskbar = control
                break
        if taskbar is None:
            raise TaskbarNotFoundError('没有找到任务栏')

        for child_depth2 in taskbar.GetChildren():
            for child_depth3 in child_depth2.GetChildren():
                for child_depth4 in child_depth3.GetChildren():
                    match = re.search('托盘输入指示器', child_depth4.Name)
                    if match:
                        self.input_indicator = child_depth4
                        break

        if self.input_indicator is None:
            raise InputIndicatorNotFoundError('没有找到托盘输入指示器')


    def switch_to_chinese(self, wait_time=0) -> None:
        time.sleep(wait_time)

        self.find_input_indicator()

        # 切换到微软拼音
        if re.search('英语\(', self.input_indicator.Name):
            # Windows + 空格
            pyautogui.hotkey('win', 'space')
        while re.search('英语\(', self.input_indicator.Name):
            pass # 等待切换完成

        # 切换到中文模式
        if re.search('英语模式', self.input_indicator.Name):
            # shift
            pyautogui.hotkey('shift')
        while re.search('英语模式', self.input_indicator.Name):
            pass # 等待切换完成
    

    def switch_to_english(self, wait_time=0) -> None:
        time.sleep(wait_time)

        self.find_input_indicator()

        # 切换到微软拼音
        if re.search('英语\(', self.input_indicator.Name):
            # Windows + 空格
            pyautogui.hotkey('win', 'space')
        while re.search('英语\(', self.input_indicator.Name):
            pass # 等待切换完成

        # 切换到英文模式
        if re.search('中文模式', self.input_indicator.Name):
            # shift
            pyautogui.hotkey('shift')
        while re.search('中文模式', self.input_indicator.Name):
            pass # 等待切换完成
    

    def auto_pinyin_input(self, characters: str, wait_time=0) -> None:
        time.sleep(wait_time)

        self.switch_to_chinese()
        time.sleep(self.ui_respond_time)

        if len(characters) > 10:
            raise ValueError('输入字符数不能超过10个')
        
        if any(ord(c) < 128 or unicodedata.category(c) == 'Po' for c in characters):
            raise ValueError('输入字符不能包含ASCII字符或中文标点')

        characters_pinyin = pinyin.get(characters, format="strip")
        pyautogui.typewrite(characters_pinyin, interval=self.type_interval)
        time.sleep(self.ui_respond_time)

        if self.candidate_panel is None:
            input_experience = auto.WindowControl(searchDepth=2, Name='Windows 输入体验')
            self.candidate_ui = input_experience.MenuControl(searchDepth=1, Name='Microsoft 候选项 UI')
            self.candidate_panel = self.candidate_ui.ListControl(searchDepth=1, Name='候选项面板')
            self.previous_page_button = self.candidate_panel.ButtonControl(searchDepth=1, Name='上一页')
            self.next_page_button = self.candidate_panel.ButtonControl(searchDepth=1, Name='下一页')
        if self.candidate_panel is None:
            raise CandidatePanelNotFoundError('没有找到候选项面板')
        
        candidates = self.candidate_panel.GetChildren()
        candidate_num = 0
        prev_name = ''
        remaining_characters = characters

        while remaining_characters != '':

            hit = False

            while candidates != None and candidates[candidate_num].Name == prev_name:
                pass

            while not hit:
                candidate_num = 0
                for candidate in candidates:
                    if candidate.ControlType != auto.ControlType.ListItemControl:
                        continue
                    candidate_num += 1
                    if remaining_characters.startswith(candidate.Name):
                        hit = True
                        prev_name = candidate.Name
                        remaining_characters = remaining_characters[len(candidate.Name):]
                        pyautogui.press(str(candidate_num))
                        time.sleep(self.ui_respond_time)
                        break
                
                if not hit:
                    pyautogui.press(']')
                    time.sleep(self.ui_respond_time)