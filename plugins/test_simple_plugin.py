# plugins/test_simple_plugin.py
import tkinter as tk
from tkinter import ttk

class TestSimplePlugin:
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
    
    def get_tab_name(self):
        return "Тест плагин"
    
    def create_tab(self):
        tab_frame = ttk.Frame(self.root)
        label = ttk.Label(tab_frame, text="Тестовый плагин работает!")
        label.pack(padx=20, pady=20)
        return tab_frame

def get_plugin_class():
    return TestSimplePlugin