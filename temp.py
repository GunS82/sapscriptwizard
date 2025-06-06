# === Казахстан BEGIN ========================================================
# Main clipboard helper application with tray icon and hotkey

import tkinter as tk
import tkinter.messagebox as messagebox
import json
import keyboard # pip install keyboard
import pystray # pip install pystray Pillow
from PIL import Image # Pillow is a dependency for pystray image handling
import threading
import sys
import os

class KeyValueApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Key/Value Clipboard Helper")
        # Determine window size based on screen resolution, or set a default
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 350
        window_height = 250
        self.root.geometry(f"{window_width}x{window_height}+{int((screen_width-window_width)/2)}+{int((screen_height-window_height)/2)}") # Center the window
        self.root.withdraw() # Hide initially

        self.json_filename = "key_value_data.json"
        self.data = self.load_data(self.json_filename)

        # --- Астана BEGIN  # GUI components and logic --------------------------
        self.create_widgets()
        # --- Астана END ------------------------------------------------------

        self.setup_tray()
        self.setup_hotkey()

        # Handle window closing (hides instead of destroying)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

    def load_data(self, filename):
        if not os.path.exists(filename):
             messagebox.showerror("Ошибка", f"JSON файл не найден: {filename}\nСоздан пустой файл. Пожалуйста, заполните его.")
             try:
                 with open(filename, 'w', encoding='utf-8') as f:
                     json.dump({}, f) # Create an empty JSON file
                 return {}
             except Exception as e:
                 messagebox.showerror("Ошибка записи", f"Не удалось создать файл: {filename}\n{e}")
                 self.root.quit()
                 sys.exit(1)

        try:
            with open(filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if not isinstance(data, dict):
                     messagebox.showerror("Ошибка формата", f"JSON файл {filename} имеет неверный формат. Ожидается объект Key:Value.")
                     return {} # Return empty data if format is wrong
                return data
        except json.JSONDecodeError:
             messagebox.showerror("Ошибка декодирования", f"Не удалось декодировать JSON из файла: {filename}\nПроверьте синтаксис JSON.")
             return {} # Return empty data on decode error
        except Exception as e:
             messagebox.showerror("Ошибка загрузки", f"Произошла ошибка при загрузке данных из {filename}: {e}")
             return {} # Return empty data on other errors


    # --- Астана BEGIN  # GUI components and logic (continued) ----------------
    def create_widgets(self):
        # Clear existing widgets
        for widget in self.root.winfo_children():
            widget.destroy()

        if not self.data:
            label = tk.Label(self.root, text="Нет данных для отображения.\nЗаполните 'key_value_data.json'.", padx=10, pady=10)
            label.pack(expand=True)
            return

        # Use a Canvas and Scrollbar if many items might exceed window height
        canvas = tk.Canvas(self.root)
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Arrange buttons
        row = 0
        col = 0
        max_cols = 2 # Arrange buttons in 2 columns

        for key, value in self.data.items():
            # Create a button for each key
            button = tk.Button(
                scrollable_frame,
                text=str(key), # Ensure text is a string
                command=lambda v=str(value): self.copy_to_clipboard(v), # Ensure value is a string
                width=20 # Fixed width for buttons
            )
            button.grid(row=row, column=col, padx=5, pady=5, sticky="ew") # sticky="ew" makes buttons expand

            col += 1
            if col >= max_cols:
                col = 0
                row += 1

        # Make columns in the scrollable_frame resize properly
        for i in range(max_cols):
            scrollable_frame.grid_columnconfigure(i, weight=1)


        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


    def copy_to_clipboard(self, value):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(value)
            # Optionally provide feedback, e.g., in console or temporary label
            # print(f"Copied '{value}' to clipboard")
        except Exception as e:
            messagebox.showerror("Ошибка буфера обмена", f"Не удалось скопировать в буфер обмена: {e}")

    def show_window(self):
        if not self.root.winfo_exists(): # Check if the main window has been destroyed
             return # Don't try to show if destroyed

        # Reload data and rebuild widgets each time the window is shown
        # This allows updating the JSON file while the app is running
        self.data = self.load_data(self.json_filename)
        self.create_widgets()
        # Adjust window size based on content if necessary, or keep fixed
        # self.root.update_idletasks() # Update geometry info
        # required_height = self.root.winfo_reqheight()
        # required_width = self.root.winfo_reqwidth()
        # self.root.geometry(f"{max(window_width, required_width)}x{max(window_height, required_height)}")


        self.root.deiconify() # Show the window
        self.root.lift()      # Bring to front
        self.root.attributes('-topmost', True) # Keep on top briefly
        self.root.after_idle(self.root.attributes, '-topmost', False) # Remove topmost after idle
        self.root.focus_force() # Force focus

    def hide_window(self):
        if self.root.winfo_exists():
             self.root.withdraw() # Hide the window
    # --- Астана END ------------------------------------------------------


    def setup_tray(self):
        # Create a simple 64x64 white image for the tray icon
        # You could replace this with a custom .ico file if preferred
        try:
            width, height = 64, 64
            image = Image.new('RGB', (width, height), 'white')
        except Exception as e:
            messagebox.showerror("Ошибка изображения", f"Не удалось создать изображение для трея: {e}\nУбедитесь, что Pillow установлен.")
            self.exit_app() # Exit if tray icon cannot be created
            return


        menu = (pystray.MenuItem('Показать', self.show_window),
                pystray.MenuItem('Выход', self.exit_app))

        # Ensure the icon runs in a separate thread as icon.run() is blocking
        self.icon = pystray.Icon("Key Value Helper", image, "Key Value Helper", menu)
        threading.Thread(target=self.icon.run, daemon=True).start()


    def setup_hotkey(self):
        # Set global hotkey (e.g., Ctrl+Shift+C)
        # 'add_hotkey' blocks, so run it in a thread
        try:
            # Use a different hotkey if Ctrl+Shift+C is commonly used elsewhere
            # 'ctrl+alt+shift+k' is less likely to conflict
            hotkey_thread = threading.Thread(target=lambda: keyboard.add_hotkey("ctrl+shift+c", self.toggle_window), daemon=True)
            hotkey_thread.start()
            # print("Hotkey Ctrl+Shift+C registered.") # For debugging
        except Exception as e:
             messagebox.showerror("Ошибка горячей клавиши", f"Не удалось зарегистрировать горячую клавишу 'Ctrl+Shift+C': {e}\nПрограмма будет работать только через иконку в трее.")


    def toggle_window(self):
        # Check if the window exists and is currently visible
        if self.root.winfo_exists() and self.root.winfo_viewable():
             self.hide_window()
        else:
             self.show_window()


    def exit_app(self):
        # Stop the tray icon thread
        if hasattr(self, 'icon') and self.icon:
            self.icon.stop()
        # Exit the Tkinter main loop
        if self.root.winfo_exists():
            self.root.quit()
        sys.exit(0) # Ensure the process exits


# === Казахстан END ==========================================================


if __name__ == "__main__":
    # Create a dummy JSON file if it doesn't exist (optional, handled by load_data)
    # For initial setup, let's ensure one is created with examples if missing.
    json_filename = "key_value_data.json"
    if not os.path.exists(json_filename):
        try:
            initial_data = {
                "Пример Ключ 1": "Пример Значение 1",
                "Другой Ключ": "Другое Значение, которое может быть длиннее",
                "Еще Один": "Последнее"
            }
            with open(json_filename, 'w', encoding='utf-8') as f:
                json.dump(initial_data, f, ensure_ascii=False, indent=4)
            print(f"Создан файл-пример: {json_filename}")
        except Exception as e:
            print(f"Не удалось создать файл-пример {json_filename}: {e}")


    root = tk.Tk()
    # Hide the root window immediately after creation
    root.withdraw()

    app = KeyValueApp(root)

    # Run the Tkinter event loop. This is blocking.
    # The tray icon and hotkey run in separate threads.
    root.mainloop()

    # Clean up any lingering threads? Daemon threads should exit with the main process.
    # It's good practice to explicitly exit the process if mainloop finishes
    # without quit() being called (e.g., fatal error), but root.quit() handles it.
    # sys.exit(0) # This might be redundant after root.quit()

