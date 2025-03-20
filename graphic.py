import tkinter as tk
from tkinter import filedialog
import threading
from main import run


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.fr = None
        self.title("Рапорти")
        self.geometry("1100x400")
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
        self.text_field = None
        self.lbl_for_text = None
        self.path_to_directory = None
        self.lbl_path = None
        self.btn_path = None
        self.btn_ok_2 = None
        self.error_lbl = None
        self.value_switcher = tk.BooleanVar()
        self.texts = [
            "Кому: ", "Звання: ", "Дата: ",
        ]
        self["background"] = "#eef3f8"
        self.create_label()
        self.create_entries()
        self.create_text_field()
        self.create_buttons()
        self.switcher_type_report()
        self.create_path_entry_get()
        self.create_btn_ok_2()

    def destroy_all(self):
        for x in range(len(self.texts)):
            self.labels[f"lbl_{x}"].destroy()
            self.entries[f"ent_{x}"].destroy()
        self.btn_ok.destroy()
        self.text_field.destroy()
        self.lbl_for_text.destroy()
        if self.lbl_switch:
            self.lbl_switch.destroy()
            self.switcher.destroy()
        self.destroy()

    def restart_init(self):
        wnd.create_frame()

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
                width=50,
            )
            self.entries[f"ent_{y}"].grid(row=y, column=2, padx=1, pady=5, sticky=tk.N, )

    def create_text_field(self):
        self.lbl_for_text = tk.Label(
            text="Текст: ",
            bg=self["background"],
            font=('Comic Sans MS', 12),
        )
        self.lbl_for_text.grid(row=len(self.texts), column=1, padx=1, pady=2, sticky=tk.N, )
        self.text_field = tk.Text(
            bg="#f2f2f2",
            font=('Comic Sans MS', 12),
            bd=1,
            height=10,
            width=50,
        )
        self.text_field.grid(row=len(self.texts), column=2, padx=1, pady=5, )

    def create_buttons(self):
        self.btn_ok = tk.Button(
            text="Формувати",
            command=self.handler_button,
            width=50,
        )
        self.btn_ok.grid(row=len(self.texts) + 1, column=2, padx=1, pady=5, )

    def handler_button(self, dev=False):
        # Окремий потік
        def task():
            list_data_entry = [
                self.entries[f"ent_0"].get(),
                self.entries[f"ent_1"].get(),
                self.entries[f"ent_2"].get(),
                self.text_field.get("1.0", tk.END),
                not self.value_switcher.get(),
            ]
            if not dev:
                if all(list_data_entry):
                    dict_data = dict()
                    dict_keys = [
                        "full_name", "rank", "date",
                    ]
                    for z in range(len(self.texts)):
                        dict_data[dict_keys[z]] = self.entries[f"ent_{z}"].get()
                    dict_data["text"] = "\t" + self.text_field.get("1.0", tk.END)
                    self.destroy_all()
                    run(data=dict_data, switcher=False)
                    self.restart_init()
                else:
                    pass
            else:
                dict_data = {
                    'full_name': 'Івану Іванову',
                    'rank': 'генерал',
                    'date': '15.03.2025',
                    'text': '\tПрошу звільнити мене за власним бажанням з лав національної гвардії України. Зобов’язуюсь відслужити два повних тижні \nу розмірі 14 календарних днів. Буду пити та гуляти ці два тижні.\n'
                }
                print("Thread")
                self.destroy_all()
                run(data=dict_data, switcher=self.value_switcher.get())
                self.restart_init()

        threading.Thread(target=task, daemon=True).start()

    def create_path_entry_get(self):
        self.lbl_path = tk.Label(
            text="Директорія: ",
            bg=self["background"],
            font=('Comic Sans MS', 12),
        )
        self.lbl_path.grid(row=1, column=3, padx=1, pady=5, )
        self.path_to_directory = tk.Entry(
            bg="#f2f2f2",
            font=('Comic Sans MS', 12),
            bd=1,
            justify=tk.CENTER,
            width=40,
        )
        self.path_to_directory.grid(row=1, column=4, padx=1, pady=5, sticky=tk.S, )

        self.btn_path = tk.Button(
            width=10,
            height=1,
            text="Шлях",
            command=self.get_path_to_directory,
        )
        self.btn_path.grid(row=1, column=5, padx=1, pady=5, sticky=tk.N, )
        self.path_to_directory.config(state=tk.DISABLED)

    def create_btn_ok_2(self):
        self.btn_ok_2 = tk.Button(
            width=10,
            height=1,
            text="Формувати",
            command=self.increment_files_docx,
        )
        self.btn_ok_2.grid(row=2, column=4, padx=1, pady=5, sticky=tk.N, )

    def increment_files_docx(self):
        if self.value_switcher.get():
            path = self.path_to_directory.get()
            if path:
                if self.error_lbl:
                    self.error_lbl.destroy()
                data = {
                    "path_to_directory": path,
                }
                run(data=data, switcher=self.value_switcher.get(), path=path)
            else:
                self.error_lbl = tk.Label(
                    text="Порожнє поле!",
                    bg=self["background"],
                    fg="red",
                    font=('Comic Sans MS', 12),
                )
                self.error_lbl.grid(row=3, column=4, padx=1, pady=5, sticky=tk.N, )
        else:
            pass

    def get_path_to_directory(self):
        path = filedialog.askdirectory()
        self.path_to_directory.delete(0, tk.END)
        self.path_to_directory.insert(0, path)

    def disabled_entry_all(self):
        if self.value_switcher.get():
            for x in range(len(self.texts)):
                self.entries[f"ent_{x}"].config(state=tk.DISABLED)
            self.text_field.config(bg="#a9a7aa")
            self.text_field.config(state=tk.DISABLED)
            self.path_to_directory.config(state=tk.NORMAL)
        else:
            for x in range(len(self.texts)):
                self.entries[f"ent_{x}"].config(state=tk.NORMAL)
            self.text_field.config(bg="#f2f2f2")
            self.text_field.config(state=tk.NORMAL)
            self.path_to_directory.config(state=tk.DISABLED)

    def switcher_type_report(self):
        self.lbl_switch = tk.Label(
            text="Вибір директорії: ", bg=self["background"], font=('Comic Sans MS', 12),
        )
        self.lbl_switch.grid(row=0, column=3, padx=1, pady=5, )
        self.switcher = tk.Checkbutton(
            variable=self.value_switcher,
            command=self.disabled_entry_all,
        )
        self.switcher.grid(row=0, column=4, padx=1, pady=5, sticky=tk.W, )


wnd = MainWindow()
wnd.mainloop()
