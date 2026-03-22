import tkinter as tk
from gui import EmailSenderProGUI


def main():
    root = tk.Tk()
    EmailSenderProGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()