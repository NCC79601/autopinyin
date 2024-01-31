import uiautomation as auto
import pyautogui
import pinyin
import unicodedata
import time
import re

from utils import split_string, chinese_punctuation_translate


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

    def __init__(self, ui_respond_time=0.06, type_interval=0.001, split_length=5) -> None:
        super().__init__()
        
        self.input_indicator = None
        self.input_experience = None
        self.candidate_ui = None
        self.candidate_panel = None
        self.previous_page_button = None
        self.next_page_button = None

        self.ui_respond_time = ui_respond_time
        self.type_interval = type_interval

        self.split_length = split_length
    

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
        """自动输入（只能是汉字），输入完成后自动按下候选项数字键"""
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
        remaining_characters = characters

        print(f'after input, candidates: {[candidate.Name for candidate in candidates]}')

        while remaining_characters != '':

            hit = False

            while self.previous_page_button.IsEnabled:
                pass

            while not hit:
                candidate_num = 0
                for candidate in candidates:
                    if candidate.ControlType != auto.ControlType.ListItemControl:
                        continue
                    candidate_num += 1
                    if remaining_characters.startswith(candidate.Name):
                        hit = True
                        remaining_characters = remaining_characters[len(candidate.Name):]
                        print(f'hit candidate #{candidate_num}: {candidate.Name}, remaining characters: {remaining_characters}')
                        print(f'candidates: {[candidate.Name for candidate in candidates]}')
                        pyautogui.press(str(candidate_num))
                        time.sleep(self.ui_respond_time)
                        break
                
                if not hit:
                    if self.next_page_button.IsEnabled:
                        pyautogui.press(']')
                        time.sleep(self.ui_respond_time)
                    else:
                        while self.previous_page_button.IsEnabled:
                            pyautogui.press('[')


    def auto_input(self, characters: str, wait_time=0) -> None:
        """自动输入（任意字符），输入完成后自动按下候选项数字键"""
        time.sleep(wait_time)

        group = split_string(characters)

        for item in group:
            if item['type'] == '汉字':
                self.switch_to_chinese()
                str = item['string']
                # Split str into substrings of length self.split_length
                substrings = [str[i:i+self.split_length] for i in range(0, len(str), self.split_length)]
                print(f'substrings: {substrings}')
                for substring in substrings:
                    self.auto_pinyin_input(substring)
            
            elif item['type'] == '中文标点':
                self.switch_to_chinese()
                str = item['string']
                str = chinese_punctuation_translate(str)
                pyautogui.typewrite(str)
            
            else:
                self.switch_to_english()
                pyautogui.typewrite(item['string'])
        
