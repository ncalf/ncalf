import tkinter as tk


def open_window():
    window = tk.Tk()

    heading = tk.Label(text="NCALF 2021")
    heading.pack()

    label = tk.Label(text="Round")
    entry = tk.Entry()
    label.pack()
    entry.pack()
    roundno = entry.get()

    print(roundno)

    window.mainloop()
