import tkinter as tk
from gui import EmailSenderGUI 


def main():
    root = tk.Tk()
    app = EmailSenderGUI(root) 
    root.mainloop()


if __name__ == "__main__":
    main()