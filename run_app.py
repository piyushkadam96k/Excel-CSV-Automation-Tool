#!/usr/bin/env python3
"""
✅ Excel/CSV Automation Tool
Simple executable version ready to use!
"""

import os
import sys

try:
    import customtkinter as ctk
    from app import App  # Import the main app class
    
    def main():
        root = ctk.CTk()
        app = App(root)
        root.mainloop()
    
    if __name__ == "__main__":
        main()
        
except Exception as e:
    print(f"Error starting application: {e}")
    input("Press Enter to exit...")
