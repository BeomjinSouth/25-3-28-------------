import os
import win32com.client
import time

class HwpController:
    """
    한글(HWP) 문서를 제어하기 위한 컨트롤러 클래스.
    win32com을 사용하여 한글 프로그램을 자동화합니다.
    """

    def __init__(self):
        self.hwp = None
        self.visible = True
        self.is_hwp_running = False
        self.current_document_path = None

    def connect(self, visible=True, register_security_module=True):
        """
        한글 프로그램에 연결합니다.
        
        Args:
            visible (bool): 한글 창을 표시할지 여부.
            register_security_module (bool): 보안 모듈 등록 여부.
        
        Returns:
            bool: 연결 성공 여부.
        """
        try:
            self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            
            # 보안 모듈 등록 (파일 경로를 실제 환경에 맞게 수정 필요)
            if register_security_module:
                try:
                    module_path = os.path.abspath("D:/hwp-mcp/security_module/FilePathCheckerModuleExample.dll")
                    self.hwp.RegisterModule("FilePathCheckerModuleExample", module_path)
                    print("보안 모듈이 등록되었습니다.")
                except Exception as e:
                    print(f"보안 모듈 등록 실패 (무시): {e}")
            
            self.visible = visible
            self.hwp.XHwpWindows.Item(0).Visible = visible
            self.is_hwp_running = True
            return True
        except Exception as e:
            print(f"한글 프로그램 연결 실패: {e}")
            return False

    def disconnect(self):
        """
        한글 프로그램과의 연결을 종료합니다.
        
        Returns:
            bool: 종료 성공 여부.
        """
        try:
            if self.is_hwp_running:
                self.hwp = None
                self.is_hwp_running = False
            return True
        except Exception as e:
            print(f"연결 해제 실패: {e}")
            return False

    def create_new_document(self):
        """
        새 문서를 생성합니다.
        
        Returns:
            bool: 생성 성공 여부.
        """
        try:
            if not self.is_hwp_running:
                self.connect()
            self.hwp.Run("FileNew")
            self.current_document_path = None
            return True
        except Exception as e:
            print(f"새 문서 생성 실패: {e}")
            return False

    def open_document(self, file_path):
        """
        기존 문서를 엽니다.
        
        Args:
            file_path (str): 열 문서의 파일 경로.
        
        Returns:
            bool: 열기 성공 여부.
        """
        try:
            if not self.is_hwp_running:
                self.connect()
            abs_path = os.path.abspath(file_path)
            self.hwp.Open(abs_path)
            self.current_document_path = abs_path
            return True
        except Exception as e:
            print(f"문서 열기 실패: {e}")
            return False

    def save_document(self, file_path=None):
        """
        문서를 저장합니다.
        
        Args:
            file_path (str, optional): 저장할 경로. 지정하지 않으면 현재 문서를 저장합니다.
        
        Returns:
            bool: 저장 성공 여부.
        """
        try:
            if not self.is_hwp_running:
                return False
            
            if file_path:
                abs_path = os.path.abspath(file_path)
                self.hwp.SaveAs(abs_path, "HWP", "")
                self.current_document_path = abs_path
            else:
                if self.current_document_path:
                    self.hwp.Save()
                else:
                    self.hwp.SaveAs()
            return True
        except Exception as e:
            print(f"문서 저장 실패: {e}")
            return False

    def insert_text(self, text, preserve_linebreaks=True):
        """
        현재 커서 위치에 텍스트를 삽입합니다.
        
        Args:
            text (str): 삽입할 텍스트.
            preserve_linebreaks (bool): 줄바꿈 보존 여부.
        
        Returns:
            bool: 삽입 성공 여부.
        """
        try:
            if not self.is_hwp_running:
                return False
            
            if preserve_linebreaks and "\n" in text:
                lines = text.split("\n")
                for i, line in enumerate(lines):
                    if i > 0:
                        self.insert_paragraph()
                    if line.strip():
                        self._insert_text_direct(line)
                return True
            else:
                return self._insert_text_direct(text)
        except Exception as e:
            print(f"텍스트 삽입 실패: {e}")
            return False

    def _insert_text_direct(self, text):
        """
        내부 함수: 텍스트를 직접 삽입합니다.
        
        Args:
            text (str): 삽입할 텍스트.
        
        Returns:
            bool: 삽입 성공 여부.
        """
        try:
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            return True
        except Exception as e:
            print(f"텍스트 직접 삽입 실패: {e}")
            return False

    def set_font_style(self, font_name=None, font_size=None, bold=False, italic=False, underline=False, select_previous_text=False):
        """
        현재 선택된 텍스트 또는 다음 입력될 텍스트의 글꼴 스타일을 설정합니다.
        
        Args:
            font_name (str, optional): 글꼴 이름.
            font_size (int, optional): 글꼴 크기 (포인트 단위).
            bold (bool): 굵게 설정 여부.
            italic (bool): 기울임 설정 여부.
            underline (bool): 밑줄 설정 여부.
            select_previous_text (bool): 이전 텍스트를 선택할지 여부.
        
        Returns:
            bool: 설정 성공 여부.
        """
        try:
            if not self.is_hwp_running:
                return False
            
            if select_previous_text:
                self.select_last_text()
            
            self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
            
            if font_name:
                self.hwp.HParameterSet.HCharShape.FaceNameHangul = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameLatin = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameHanja = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameJapanese = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameOther = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameSymbol = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameUser = font_name
            
            if font_size:
                # 한글 단위 (10pt = 1000)
                self.hwp.HParameterSet.HCharShape.Height = font_size * 100
            
            self.hwp.HParameterSet.HCharShape.Bold = bold
            self.hwp.HParameterSet.HCharShape.Italic = italic
            self.hwp.HParameterSet.HCharShape.UnderlineType = 1 if underline else 0
            
            self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
            return True
        except Exception as e:
            print(f"글꼴 스타일 설정 실패: {e}")
            return False

    def set_font(self, font_name, font_size, bold=False, italic=False, select_previous_text=False):
        """
        글꼴을 설정합니다. (set_font_style을 내부적으로 호출)
        
        Args:
            font_name (str): 글꼴 이름.
            font_size (int): 글꼴 크기.
            bold (bool): 굵게 설정 여부.
            italic (bool): 기울임 설정 여부.
            select_previous_text (bool): 이전 텍스트 선택 여부.
        
        Returns:
            bool: 설정 성공 여부.
        """
        return self.set_font_style(
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            select_previous_text=select_previous_text
        )

    def insert_paragraph(self):
        """
        새 단락을 삽입합니다.
        
        Returns:
            bool: 삽입 성공 여부.
        """
        try:
            if not self.is_hwp_running:
                return False
            self.hwp.HAction.Run("BreakPara")
            return True
        except Exception as e:
            print(f"단락 삽입 실패: {e}")
            return False

    def select_last_text(self):
        """
        현재 단락의 마지막 입력 텍스트를 선택합니다.
        
        Returns:
            bool: 선택 성공 여부.
        """
        try:
            if not self.is_hwp_running:
                return False
            
            current_pos = self.hwp.GetPos()
            if not current_pos:
                return False
            
            self.hwp.Run("MoveLineStart")
            start_pos = self.hwp.GetPos()
            self.hwp.SetPos(*start_pos)
            self.hwp.SelectText(start_pos, current_pos)
            return True
        except Exception as e:
            print(f"텍스트 선택 실패: {e}")
            return False

    def get_text(self):
        """
        현재 문서의 전체 텍스트를 가져옵니다.
        
        Returns:
            str: 문서 텍스트.
        """
        try:
            if not self.is_hwp_running:
                return ""
            return self.hwp.GetTextFile("TEXT", "")
        except Exception as e:
            print(f"텍스트 가져오기 실패: {e}")
            return ""
