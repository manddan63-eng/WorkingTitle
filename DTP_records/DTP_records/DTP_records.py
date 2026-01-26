import tkinter as tk

def main():
    root = tk.Tk()
    root.title("Тест")
    label = tk.Label(root, text="Кликни меня!")
    label.pack(padx=20, pady=20)
    root.mainloop()  # <-- Это блокирующий вызов!

if __name__ == "__main__":
    main()