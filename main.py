"""
HWP to PDF 일괄 변환기
개발자: leedoyeon
버전: 1.0
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
from pathlib import Path
import sys

# Windows 환경에서만 import
try:
    import win32com.client
    import pythoncom
    WINDOWS_AVAILABLE = True
except ImportError:
    WINDOWS_AVAILABLE = False

class HWPToPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("HWP to PDF 일괄 변환기 v1.0")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        # 창 아이콘 설정 (선택사항)
        try:
            # self.root.iconbitmap('icon.ico')  # 아이콘 파일이 있다면 주석 해제
            pass
        except:
            pass
        
        # 선택된 파일들
        self.hwp_files = []
        self.is_converting = False
        
        # 초기 검사
        if not self.check_system_requirements():
            return
            
        self.setup_ui()
        
    def check_system_requirements(self):
        """시스템 요구사항을 확인합니다."""
        try:
            if not WINDOWS_AVAILABLE:
                messagebox.showerror(
                    "시스템 오류", 
                    "이 프로그램은 Windows에서만 실행할 수 있습니다.\n"
                    "또한 필요한 시스템 구성요소가 설치되지 않았습니다."
                )
                self.root.destroy()
                return False
                
            # 한글 프로그램 설치 확인
            if not self.check_hwp_installed():
                result = messagebox.askyesno(
                    "한글 프로그램 필요", 
                    "한글과컴퓨터의 '한글' 프로그램이 설치되어 있지 않거나\n"
                    "정상적으로 작동하지 않습니다.\n\n"
                    "한글 프로그램 없이는 HWP 파일을 변환할 수 없습니다.\n"
                    "그래도 프로그램을 계속 실행하시겠습니까?"
                )
                if not result:
                    self.root.destroy()
                    return False
                    
            return True
            
        except Exception as e:
            try:
                messagebox.showerror("초기화 오류", f"프로그램 초기화 중 오류가 발생했습니다:\n{str(e)}")
            except:
                pass
            try:
                self.root.destroy()
            except:
                pass
            return False
    
    def check_hwp_installed(self):
        """한글 프로그램 설치 여부를 확인합니다."""
        try:
            pythoncom.CoInitialize()
            hwp = win32com.client.Dispatch("HWPApplication.HwpObject")
            hwp.Quit()
            pythoncom.CoUninitialize()
            return True
        except Exception as e:
            return False
        
    def setup_ui(self):
        """UI를 설정합니다."""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(
            main_frame, 
            text="🔄 HWP to PDF 일괄 변환기", 
            font=("맑은 고딕", 18, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 버전 정보
        version_label = ttk.Label(
            main_frame, 
            text="v1.0 - 한글 문서를 PDF로 쉽게 변환하세요", 
            font=("맑은 고딕", 9)
        )
        version_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # 파일 선택 프레임
        file_frame = ttk.LabelFrame(main_frame, text="📁 HWP 파일 선택", padding="10")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 파일 선택 버튼들
        btn_frame = ttk.Frame(file_frame)
        btn_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        select_btn = ttk.Button(btn_frame, text="📂 HWP 파일 선택", command=self.select_files)
        select_btn.grid(row=0, column=0, padx=(0, 10))
        
        clear_btn = ttk.Button(btn_frame, text="🗑️ 목록 지우기", command=self.clear_files)
        clear_btn.grid(row=0, column=1, padx=(0, 10))
        
        # 선택된 파일 개수 표시
        self.file_count_label = ttk.Label(btn_frame, text="선택된 파일: 0개", font=("맑은 고딕", 9))
        self.file_count_label.grid(row=0, column=2, padx=(20, 0))
        
        # 파일 목록 표시
        list_frame = ttk.LabelFrame(main_frame, text="📋 선택된 파일 목록", padding="10")
        list_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 리스트박스와 스크롤바
        list_container = ttk.Frame(list_frame)
        list_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.file_listbox = tk.Listbox(
            list_container, 
            height=8, 
            font=("맑은 고딕", 9),
            selectmode=tk.EXTENDED
        )
        v_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.file_listbox.yview)
        h_scrollbar = ttk.Scrollbar(list_container, orient="horizontal", command=self.file_listbox.xview)
        
        self.file_listbox.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 선택된 파일 삭제 기능
        remove_btn = ttk.Button(list_frame, text="❌ 선택한 파일 제거", command=self.remove_selected_files)
        remove_btn.grid(row=1, column=0, pady=(10, 0))
        
        # 출력 폴더 설정
        output_frame = ttk.LabelFrame(main_frame, text="📁 출력 설정", padding="10")
        output_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(output_frame, text="PDF 저장 위치:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.output_path = tk.StringVar(value="원본 파일과 같은 폴더")
        self.output_entry = ttk.Entry(
            output_frame, 
            textvariable=self.output_path, 
            state="readonly", 
            width=60,
            font=("맑은 고딕", 9)
        )
        self.output_entry.grid(row=1, column=0, padx=(0, 10), sticky=(tk.W, tk.E))
        
        output_btn = ttk.Button(output_frame, text="📂 폴더 선택", command=self.select_output_folder)
        output_btn.grid(row=1, column=1)
        
        reset_output_btn = ttk.Button(output_frame, text="🔄 기본값", command=self.reset_output_folder)
        reset_output_btn.grid(row=1, column=2, padx=(5, 0))
        
        # 변환 설정
        settings_frame = ttk.LabelFrame(main_frame, text="⚙️ 변환 설정", padding="10")
        settings_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 덮어쓰기 옵션
        self.overwrite_var = tk.BooleanVar(value=True)
        overwrite_check = ttk.Checkbutton(
            settings_frame, 
            text="기존 PDF 파일이 있으면 덮어쓰기", 
            variable=self.overwrite_var
        )
        overwrite_check.grid(row=0, column=0, sticky=tk.W)
        
        # 변환 버튼과 진행률
        convert_frame = ttk.LabelFrame(main_frame, text="🚀 변환 실행", padding="10")
        convert_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        button_frame = ttk.Frame(convert_frame)
        button_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.convert_btn = ttk.Button(
            button_frame, 
            text="🔄 PDF로 변환 시작", 
            command=self.start_conversion,
            style="Accent.TButton"
        )
        self.convert_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.stop_btn = ttk.Button(
            button_frame, 
            text="⏹️ 변환 중지", 
            command=self.stop_conversion,
            state="disabled"
        )
        self.stop_btn.grid(row=0, column=1)
        
        # 진행률 표시
        progress_frame = ttk.Frame(convert_frame)
        progress_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(progress_frame, text="진행률:").grid(row=0, column=0, sticky=tk.W)
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        self.status_label = ttk.Label(
            progress_frame, 
            text="변환 대기 중... 파일을 선택하고 변환을 시작하세요.", 
            font=("맑은 고딕", 9)
        )
        self.status_label.grid(row=2, column=0, pady=(5, 0), sticky=tk.W)
        
        # 하단 정보
        info_label = ttk.Label(
            main_frame, 
            text="💡 주의: 변환 중에는 한글 프로그램이 자동으로 실행됩니다. 다른 한글 작업은 잠시 중단해주세요.",
            font=("맑은 고딕", 8),
            foreground="gray"
        )
        info_label.grid(row=7, column=0, columnspan=3, pady=(10, 0))
        
        # 그리드 가중치 설정
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)  # 파일 목록이 확장되도록
        
        file_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        list_container.columnconfigure(0, weight=1)
        list_container.rowconfigure(0, weight=1)
        
        output_frame.columnconfigure(0, weight=1)
        settings_frame.columnconfigure(0, weight=1)
        convert_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(0, weight=1)
        progress_frame.columnconfigure(0, weight=1)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def select_files(self):
        """HWP 파일들을 선택합니다."""
        files = filedialog.askopenfilenames(
            title="변환할 HWP 파일들을 선택하세요",
            filetypes=[
                ("HWP 파일", "*.hwp"),
                ("모든 파일", "*.*")
            ]
        )
        
        if files:
            # 중복 제거하며 파일 추가
            new_files = [f for f in files if f not in self.hwp_files]
            self.hwp_files.extend(new_files)
            self.update_file_list()
            
            if new_files:
                self.status_label.config(text=f"{len(new_files)}개의 새 파일이 추가되었습니다.")
    
    def clear_files(self):
        """파일 목록을 모두 지웁니다."""
        if self.hwp_files and messagebox.askyesno("확인", "선택된 모든 파일을 목록에서 제거하시겠습니까?"):
            self.hwp_files.clear()
            self.update_file_list()
            self.status_label.config(text="파일 목록이 지워졌습니다.")
    
    def remove_selected_files(self):
        """선택된 파일들을 목록에서 제거합니다."""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("안내", "제거할 파일을 목록에서 선택해주세요.")
            return
            
        # 뒤에서부터 제거 (인덱스 변화 방지)
        for i in reversed(selected_indices):
            del self.hwp_files[i]
        
        self.update_file_list()
        self.status_label.config(text=f"{len(selected_indices)}개의 파일이 제거되었습니다.")
    
    def update_file_list(self):
        """파일 목록을 업데이트합니다."""
        self.file_listbox.delete(0, tk.END)
        
        for i, file_path in enumerate(self.hwp_files):
            filename = os.path.basename(file_path)
            folder = os.path.dirname(file_path)
            display_text = f"{i+1:2d}. {filename}"
            self.file_listbox.insert(tk.END, display_text)
        
        count = len(self.hwp_files)
        self.file_count_label.config(text=f"선택된 파일: {count}개")
        
        # 변환 버튼 상태 업데이트
        if count > 0 and not self.is_converting:
            self.convert_btn.config(state="normal")
        else:
            self.convert_btn.config(state="disabled")
    
    def select_output_folder(self):
        """출력 폴더를 선택합니다."""
        folder = filedialog.askdirectory(title="PDF 파일을 저장할 폴더를 선택하세요")
        if folder:
            self.output_path.set(folder)
    
    def reset_output_folder(self):
        """출력 폴더를 기본값으로 리셋합니다."""
        self.output_path.set("원본 파일과 같은 폴더")
    
    def start_conversion(self):
        """변환을 시작합니다."""
        if not self.hwp_files:
            messagebox.showwarning("경고", "변환할 HWP 파일을 선택해주세요.")
            return
        
        # 출력 폴더 확인
        if self.output_path.get() != "원본 파일과 같은 폴더":
            if not os.path.exists(self.output_path.get()):
                messagebox.showerror("오류", "선택한 출력 폴더가 존재하지 않습니다.")
                return
        
        # UI 상태 변경
        self.is_converting = True
        self.convert_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        
        # 스레드에서 변환 실행
        self.conversion_thread = threading.Thread(target=self.convert_files)
        self.conversion_thread.daemon = True
        self.conversion_thread.start()
    
    def stop_conversion(self):
        """변환을 중지합니다."""
        self.is_converting = False
        self.status_label.config(text="변환을 중지하는 중...")
    
    def convert_files(self):
        """실제 파일 변환을 수행합니다."""
        hwp = None
        try:
            # COM 초기화
            pythoncom.CoInitialize()
            
            total_files = len(self.hwp_files)
            self.progress.config(maximum=total_files, value=0)
            
            # 한글 애플리케이션 연결
            try:
                hwp = win32com.client.Dispatch("HWPApplication.HwpObject")
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except Exception as e:
                raise Exception(f"한글 프로그램을 실행할 수 없습니다: {str(e)}")
            
            converted_count = 0
            failed_files = []
            skipped_files = []
            
            for i, hwp_file in enumerate(self.hwp_files):
                if not self.is_converting:  # 중지 요청 확인
                    break
                    
                try:
                    filename = os.path.basename(hwp_file)
                    self.root.after(0, lambda f=filename, current=i+1, total=total_files: 
                                   self.status_label.config(text=f"변환 중: {f} ({current}/{total})"))
                    
                    # 출력 경로 결정
                    if self.output_path.get() == "원본 파일과 같은 폴더":
                        output_dir = os.path.dirname(hwp_file)
                    else:
                        output_dir = self.output_path.get()
                    
                    # PDF 파일명 생성
                    base_name = Path(hwp_file).stem
                    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
                    
                    # 기존 파일 존재 확인
                    if os.path.exists(pdf_path) and not self.overwrite_var.get():
                        skipped_files.append(filename)
                        continue
                    
                    # HWP 파일 열기
                    result = hwp.Open(hwp_file)
                    if not result:
                        raise Exception("파일을 열 수 없습니다")
                    
                    # PDF로 저장
                    hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
                    hwp.HParameterSet.HFileOpenSave.filename = pdf_path
                    hwp.HParameterSet.HFileOpenSave.Format = "PDF"
                    save_result = hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
                    
                    if not save_result:
                        raise Exception("PDF 저장에 실패했습니다")
                    
                    # 파일 닫기
                    hwp.Clear(1)
                    
                    converted_count += 1
                    
                except Exception as e:
                    failed_files.append((filename, str(e)))
                    # 오류 발생 시에도 문서 닫기 시도
                    try:
                        if hwp:
                            hwp.Clear(1)
                    except:
                        pass
                
                # 진행률 업데이트
                self.root.after(0, lambda progress=i + 1: self.progress.config(value=progress))
            
            # 결과 메시지 생성
            if not self.is_converting:
                message = f"변환이 중지되었습니다.\n완료된 파일: {converted_count}개"
            else:
                message_parts = [f"변환 완료!\n\n✅ 성공: {converted_count}개"]
                
                if skipped_files:
                    message_parts.append(f"⏭️ 건너뜀: {len(skipped_files)}개 (기존 파일 존재)")
                
                if failed_files:
                    message_parts.append(f"❌ 실패: {len(failed_files)}개")
                    if len(failed_files) <= 3:
                        message_parts.append("\n실패한 파일:")
                        for filename, error in failed_files:
                            message_parts.append(f"• {filename}")
                
                message = "\n".join(message_parts)
            
            self.root.after(0, lambda msg=message: messagebox.showinfo("변환 결과", msg))
            
        except Exception as e:
            error_msg = f"변환 중 오류가 발생했습니다:\n{str(e)}\n\n한글 프로그램이 제대로 설치되어 있는지 확인해주세요."
            self.root.after(0, lambda: messagebox.showerror("오류", error_msg))
        
        finally:
            # 한글 프로그램 종료
            try:
                if hwp:
                    hwp.Quit()
            except:
                pass
            
            # COM 정리
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            
            # UI 복원
            self.root.after(0, self.reset_ui)
    
    def reset_ui(self):
        """UI를 초기 상태로 복원합니다."""
        self.is_converting = False
        self.convert_btn.config(state="normal" if self.hwp_files else "disabled")
        self.stop_btn.config(state="disabled")
        self.progress.config(value=0)
        if self.hwp_files:
            self.status_label.config(text="변환이 완료되었습니다. 새로운 변환을 시작할 수 있습니다.")
        else:
            self.status_label.config(text="변환 대기 중... 파일을 선택하고 변환을 시작하세요.")

def main():
    """메인 함수"""
    try:
        # Windows 환경 체크
        if not WINDOWS_AVAILABLE:
            root = tk.Tk()
            root.withdraw()  # 메인 창 숨기기
            messagebox.showerror(
                "시스템 오류", 
                "이 프로그램은 Windows에서만 실행할 수 있습니다.\n"
                "또한 다음 구성요소가 필요합니다:\n"
                "- 한글과컴퓨터 '한글' 프로그램\n"
                "- Windows COM 지원"
            )
            root.destroy()
            return
        
        # Tkinter 루트 생성
        root = tk.Tk()
        
        # 앱 인스턴스 생성
        app = HWPToPDFConverter(root)
        
        # 메인 루프 실행 (app 초기화가 성공한 경우에만)
        try:
            if root.winfo_exists():
                root.mainloop()
        except tk.TclError:
            pass  # 창이 이미 닫힌 경우
            
    except Exception as e:
        try:
            # 오류 창 표시
            error_root = tk.Tk()
            error_root.withdraw()
            messagebox.showerror("실행 오류", f"프로그램 실행 중 오류가 발생했습니다:\n{str(e)}")
            error_root.destroy()
        except:
            # Tkinter도 사용할 수 없는 경우
            print(f"오류: {str(e)}")
        return

if __name__ == "__main__":
    main()