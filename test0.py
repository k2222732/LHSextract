import tkinter as tk
from tkinter import ttk

root = tk.Tk()

# 创建一个Frame作为从属元素
frame = ttk.Frame(root)
frame.grid(row=1, column=1)

# 在Frame中放置两个元素
label1 = ttk.Label(frame, text="Label 1")
label2 = ttk.Label(frame, text="Label 2")

label1.grid(row=1, column=1)
label2.grid(row=1, column=2)

root.mainloop()
