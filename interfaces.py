


def choose_folder():
    import tkinter
    root=tkinter.Tk()
    result =tkinter.filedialog.askdirectory()
    root.withdraw()
    return result
