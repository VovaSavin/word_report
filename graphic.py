import tkinter as tk

from main import run


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.fr = None
        self.title("Рапорти")
        self.geometry("400x300")
        self.create_frame()

    def create_frame(self):
        self.fr = FrameCustom()
        self.fr.place(relx=0, rely=0, relwidth=1, relheight=1)


class FrameCustom(tk.Frame):
    def __init__(self):
        super().__init__()
        self.labels = dict()
        self.entries = dict()
        self.btn_ok = None
        self.lbl_switch = None
        self.switcher = None
        self.value_switcher = tk.BooleanVar()
        self.texts = [
            "Кому: ", "Звання: ", "Дата: ",
        ]
        self["background"] = "#eef3f8"
        self.create_label()
        self.create_entries()
        self.create_buttons()

    def create_label(self):

        for x in range(len(self.texts)):
            self.labels[f"lbl_{x}"] = tk.Label(
                text=self.texts[x],
                bg=self["background"],
                font=('Comic Sans MS', 12),
            )
            self.labels[f"lbl_{x}"].grid(row=x, column=1, padx=1, pady=5, )

    def create_entries(self):
        for y in range(len(self.texts)):
            self.entries[f"ent_{y}"] = tk.Entry(
                bg="#f2f2f2",
                font=('Comic Sans MS', 12),
                bd=1,
                justify=tk.CENTER,
            )
            self.entries[f"ent_{y}"].grid(row=y, column=2, padx=1, pady=5, )

    def create_buttons(self):
        self.btn_ok = tk.Button(text="Формувати", command=self.handler_button, )
        self.btn_ok.grid(row=len(self.texts), column=2, padx=1, pady=5, )

    def handler_button(self):
        dict_data = dict()
        dict_keys = [
            "full_name", "rank", "date",
        ]
        for z in range(len(self.texts)):
            dict_data[dict_keys[z]] = self.entries[f"ent_{z}"].get()
        run(data=dict_data, switcher=True)

    def disabled_entry_all(self):
        if self.value_switcher.get():
            for x in range(len(self.texts)):
                self.entries[f"ent_{x}"].config(state=tk.DISABLED)
        else:
            for x in range(len(self.texts)):
                self.entries[f"ent_{x}"].config(state=tk.NORMAL)

    def switcher_type_report(self):
        self.lbl_switch = tk.Label(
            text="З файлу names.txt: ", bg=self["background"], font=('Comic Sans MS', 12),
        )
        self.lbl_switch.grid(row=0, column=3, padx=1, pady=5, )
        self.switcher = tk.Checkbutton(
            variable=self.value_switcher,
            command=self.disabled_entry_all,
        )
        self.switcher.grid(row=0, column=4, padx=1, pady=5, )


wnd = MainWindow()
wnd.mainloop()
