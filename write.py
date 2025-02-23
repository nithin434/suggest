import keyboard
import pyperclip
import cohere
from PyQt5.QtWidgets import QStyle
from PyQt5.QtWidgets import (QApplication, QLabel, QMainWindow, QPushButton, 
                            QVBoxLayout, QWidget, QHBoxLayout, QSystemTrayIcon)
from PyQt5.QtCore import Qt, QTimer, QPoint
from PyQt5.QtGui import QIcon
import sys
import win32gui
import win32process
import psutil
import threading
import json
import os

class SuggestionCategory:
    GRAMMAR = "grammar"
    STYLE = "style"
    TONE = "tone"
    CLARITY = "clarity"

class WritingAssistant:
    def __init__(self):
        self.previous_text = ""
        self.is_active = True
        self.buffer = []
        self.typing_timer = QTimer()
        self.typing_timer.setSingleShot(True)
        self.typing_timer.timeout.connect(self.process_buffer)
        self.config = self.load_config()
        self.current_suggestions = []
        self.suggestion_index = 0
        
        # Initialize Cohere API
        self.cohere_client = cohere.Client(self.config.get('cohere_api_key', ''))
        
        # Initialize GUI
        self.app = QApplication(sys.argv)
        # Create overlay window in main thread
        self.overlay = None
        self.app.aboutToQuit.connect(self.cleanup)
        # Initialize overlay in main thread
        self.init_overlay()
        self.status_icon = None  # Add this line
        
        # Create system tray icon
        self.setup_system_tray()  # Add this line
        
        # Register global shortcuts
        self.setup_shortcuts()
        
        # Start keyboard monitoring
        self.start_monitoring()

    def init_overlay(self):
        """Initialize overlay window in main thread"""
        self.overlay = OverlayWindow(self)
        self.overlay.hide()

    def cleanup(self):
        """Cleanup resources before exit"""
        if self.typing_timer:
            self.typing_timer.stop()
        if self.status_icon:
            self.status_icon.hide()

    def load_config(self):
        try:
            with open('config.json', 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            default_config = {
                'cohere_api_key': '',
                'suggestion_delay': 1.0,
                'context_window': 100,
                'enabled_categories': [
                    SuggestionCategory.GRAMMAR,
                    SuggestionCategory.STYLE,
                    SuggestionCategory.TONE,
                    SuggestionCategory.CLARITY
                ],
                'app_specific_settings': {
                    'code_editors': ['vscode.exe', 'pycharm64.exe', 'sublime_text.exe'],
                    'doc_editors': ['winword.exe', 'wordpad.exe', 'notepad.exe'],
                    'browsers': ['chrome.exe', 'firefox.exe', 'msedge.exe']
                }
            }
            with open('config.json', 'w') as f:
                json.dump(default_config, f)
            return default_config

    def setup_shortcuts(self):
        # Global shortcuts for suggestion navigation
        keyboard.add_hotkey('ctrl+alt+right', self.next_suggestion)
        keyboard.add_hotkey('ctrl+alt+left', self.previous_suggestion)
        keyboard.add_hotkey('ctrl+alt+enter', self.accept_suggestion)
        keyboard.add_hotkey('ctrl+alt+backspace', self.reject_suggestion)
        keyboard.add_hotkey('ctrl+alt+space', self.toggle_assistant)

    def setup_system_tray(self):
        """Setup system tray icon to show assistant status"""
        from PyQt5.QtWidgets import QSystemTrayIcon
        from PyQt5.QtGui import QIcon
        
        # Create or load an icon (you'll need to provide an icon file)
        self.status_icon = QSystemTrayIcon(self.app)
        
        # Create a default icon or load from file
        # You can replace this with your own .ico file path
        icon_path = os.path.join(os.path.dirname(__file__), 'assistant_icon.ico')
        if not os.path.exists(icon_path):
            # If icon doesn't exist, use a default system icon
            self.status_icon.setIcon(self.app.style().standardIcon(QStyle.SP_ComputerIcon))
        else:
            self.status_icon.setIcon(QIcon(icon_path))
        
        self.status_icon.setToolTip('Writing Assistant Active')
        self.status_icon.show()

    def get_active_application(self):
        try:
            # Get the foreground window handle
            hwnd = win32gui.GetForegroundWindow()
            # Get the process ID
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            # Get the process name
            process = psutil.Process(pid)
            return process.name().lower()
        except Exception:
            return None

    def get_context_specific_prompt(self, text, app_name):
        if not app_name:
            return self.get_general_prompt(text)

        app_settings = self.config['app_specific_settings']
        
        if app_name in app_settings['code_editors']:
            return f"""Analyze this code snippet and suggest improvements for:
            - Code style and conventions
            - Potential bugs or issues
            - Performance optimizations
            - Documentation needs
            Text: {text}"""
            
        elif app_name in app_settings['doc_editors']:
            return f"""Review this text and suggest improvements for:
            - Grammar and spelling
            - Style and clarity
            - Tone and professionalism
            - Document structure
            Text: {text}"""
            
        elif app_name in app_settings['browsers']:
            return f"""Review this web content and suggest improvements for:
            - Clarity and conciseness
            - Tone and engagement
            - Format and structure
            Text: {text}"""
            
        return self.get_general_prompt(text)

    def get_general_prompt(self, text):
        categories = []
        if SuggestionCategory.GRAMMAR in self.config['enabled_categories']:
            categories.append("grammar and spelling")
        if SuggestionCategory.STYLE in self.config['enabled_categories']:
            categories.append("writing style")
        if SuggestionCategory.TONE in self.config['enabled_categories']:
            categories.append("tone and voice")
        if SuggestionCategory.CLARITY in self.config['enabled_categories']:
            categories.append("clarity and conciseness")
        
        categories_str = ", ".join(categories)
        return f"Review this text and suggest specific improvements for {categories_str}: {text}"

    def get_context(self):
        """Get context from surrounding text"""
        # Store last N words for context
        context_window = self.config.get('context_window', 100)
        return ''.join(self.buffer[-context_window:])

    def get_ai_suggestions(self, text):
        try:
            app_name = self.get_active_application()
            context = self.get_context()
            prompt = self.get_context_specific_prompt(f"{context} {text}", app_name)
            
            # Show typing indicator while processing
            cursor_pos = win32gui.GetCursorPos()
            self.overlay.show_typing_indicator(QPoint(*cursor_pos))
            
            response = self.cohere_client.generate(
                model='command',
                prompt=prompt,
                max_tokens=150,
                temperature=0.7,
                k=5,
                stop_sequences=["\n\n"],
                return_likelihoods='NONE'
            )
            
            # Parse multiple suggestions
            suggestions = response.generations[0].text.strip().split('\n')
            self.current_suggestions = [s.strip() for s in suggestions if s.strip()]
            self.suggestion_index = 0
            
            self.overlay.hide_typing_indicator()
            return self.current_suggestions[0] if self.current_suggestions else None
            
        except Exception as e:
            self.overlay.hide_typing_indicator()
            print(f"Error getting AI suggestions: {e}")
            return None

    def next_suggestion(self):
        if not self.current_suggestions:
            return
            
        self.suggestion_index = (self.suggestion_index + 1) % len(self.current_suggestions)
        cursor_pos = win32gui.GetCursorPos()
        self.overlay.show_suggestions(
            self.current_suggestions[self.suggestion_index],
            QPoint(*cursor_pos)
        )

    def previous_suggestion(self):
        if not self.current_suggestions:
            return
            
        self.suggestion_index = (self.suggestion_index - 1) % len(self.current_suggestions)
        cursor_pos = win32gui.GetCursorPos()
        self.overlay.show_suggestions(
            self.current_suggestions[self.suggestion_index],
            QPoint(*cursor_pos)
        )

    def accept_suggestion(self):
        if not self.current_suggestions:
            return
            
        # Implement the suggestion by simulating keyboard input
        suggestion = self.current_suggestions[self.suggestion_index]
        keyboard.write(suggestion)
        self.current_suggestions = []
        self.overlay.hide()

    def reject_suggestion(self):
        self.current_suggestions = []
        self.overlay.hide()

    def toggle_assistant(self):
        self.is_active = not self.is_active
        status = "enabled" if self.is_active else "disabled"
        self.overlay.show_status(f"Writing Assistant {status}")

    def process_buffer(self):
        """Process the text buffer and get AI suggestions"""
        if len(self.buffer) > 0:
            if self.status_icon:
                self.status_icon.setToolTip('Processing...')
            
            text = ''.join(self.buffer[-self.config['context_window']:])
            suggestion = self.get_ai_suggestions(text)
            
            if suggestion:
                cursor_pos = win32gui.GetCursorPos()
                self.overlay.show_suggestions(suggestion, QPoint(*cursor_pos))
            
            if self.status_icon:
                self.status_icon.setToolTip('Writing Assistant Active')
            
            self.buffer = []

    def start_monitoring(self):
        """Start monitoring keyboard input"""
        def on_key_event(event):
            if not self.is_active:
                return
                
            if event.event_type == keyboard.KEY_DOWN and event.name.isprintable():
                self.buffer.append(event.name)
                
                # Reset and start timer using Qt timer
                self.typing_timer.stop()
                self.typing_timer.start(int(self.config['suggestion_delay'] * 1000))

        keyboard_thread = threading.Thread(
            target=keyboard.hook,
            args=(on_key_event,),
            daemon=True
        )
        keyboard_thread.start()
        
        # Start Qt event loop
        self.app.exec_()

class OverlayWindow(QMainWindow):
    def __init__(self, assistant):
        super().__init__()
        self.assistant = assistant
        self.setup_ui()
        self.typing_indicator = QLabel()
        self.typing_indicator.setStyleSheet("""
            QLabel {
                background-color: rgba(0, 0, 0, 0.8);
                color: white;
                padding: 5px;
                border-radius: 3px;
                font-size: 10px;
            }
        """)
        self.typing_indicator.hide()
        
        # Animation dots
        self.animation_timer = QTimer()
        self.animation_timer.timeout.connect(self.update_typing_animation)
        self.animation_dots = 0

    def setup_ui(self):
        self.setWindowFlags(
            Qt.FramelessWindowHint | 
            Qt.WindowStaysOnTopHint |
            Qt.Tool
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Suggestion label
        self.label = QLabel()
        self.label.setStyleSheet("""
            QLabel {
                background-color: rgba(0, 0, 0, 0.8);
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 12px;
            }
        """)
        layout.addWidget(self.label)
        
        # Navigation buttons
        nav_layout = QHBoxLayout()
        prev_btn = QPushButton("←")
        next_btn = QPushButton("→")
        accept_btn = QPushButton("✓")
        reject_btn = QPushButton("✗")
        
        for btn in [prev_btn, next_btn, accept_btn, reject_btn]:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: rgba(0, 0, 0, 0.8);
                    color: white;
                    border: none;
                    padding: 5px;
                    border-radius: 3px;
                    min-width: 25px;
                }
                QPushButton:hover {
                    background-color: rgba(60, 60, 60, 0.8);
                }
            """)
            nav_layout.addWidget(btn)
        
        layout.addLayout(nav_layout)
        
        # Connect buttons
        prev_btn.clicked.connect(self.assistant.previous_suggestion)
        next_btn.clicked.connect(self.assistant.next_suggestion)
        accept_btn.clicked.connect(self.assistant.accept_suggestion)
        reject_btn.clicked.connect(self.assistant.reject_suggestion)

    def show_suggestions(self, text, pos):
        self.label.setText(text)
        self.label.adjustSize()
        self.move(pos.x() + 20, pos.y() + 20)
        self.resize(self.sizeHint())
        self.show()
        
    def show_status(self, text):
        self.label.setText(text)
        self.label.adjustSize()
        cursor_pos = win32gui.GetCursorPos()
        self.move(cursor_pos[0] + 20, cursor_pos[1] + 20)
        self.resize(self.sizeHint())
        self.show()
        QTimer.singleShot(2000, self.hide)

    def update_typing_animation(self):
        """Update typing indicator animation"""
        self.animation_dots = (self.animation_dots + 1) % 4
        self.typing_indicator.setText(f"Thinking{'.' * self.animation_dots}")

    def show_typing_indicator(self, pos):
        """Show typing indicator at cursor position"""
        self.typing_indicator.show()
        self.typing_indicator.move(pos.x() + 20, pos.y() + 20)
        self.animation_timer.start(500)

    def hide_typing_indicator(self):
        """Hide typing indicator"""
        self.typing_indicator.hide()
        self.animation_timer.stop()

if __name__ == "__main__":
    WritingAssistant()
