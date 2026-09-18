"""Local registration of bankruptcy debtors for subsequent preparations."""

from tkinter import StringVar, Toplevel, ttk

from .debtor_fallbacks import load_local, save_debtor
from .errors import InputFileError


class DebtorEditor:
    def __init__(self, parent, on_saved, path=None):
        self.path = path
        self.on_saved = on_saved
        self.records = load_local(path)
        self.window = Toplevel(parent)
        self.window.title("Cadastro de sacados (massas falidas)")
        self.window.geometry("820x530")
        self.window.transient(parent)
        self.window.grab_set()
        frame = ttk.Frame(self.window, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Cadastre o CNPJ e o nome oficial confirmados da massa falida.").pack(
            anchor="w")
        ttk.Label(frame, text="Após salvar, prepare um novo Excel para aplicar o cadastro.").pack(
            anchor="w")
        self.tree = ttk.Treeview(frame, columns=("failure", "document", "name"),
                                 show="headings", height=6)
        for column, title in (("failure", "Falência"), ("document", "CNPJ"),
                              ("name", "Nome oficial")):
            self.tree.heading(column, text=title)
        self.tree.pack(fill="both", expand=True, pady=10)
        self.tree.bind("<<TreeviewSelect>>", self.select)
        self.failure = StringVar(master=self.window)
        self.document = StringVar(master=self.window)
        self.name = StringVar(master=self.window)
        for title, variable in (("Falência (nome usado nos arquivos)", self.failure),
                                ("CNPJ", self.document), ("Nome oficial do sacado", self.name)):
            ttk.Label(frame, text=title).pack(anchor="w")
            ttk.Entry(frame, textvariable=variable).pack(fill="x")
        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=10)
        ttk.Button(actions, text="Salvar sacado", command=self.save).pack(side="left")
        ttk.Button(actions, text="Fechar", command=self.window.destroy).pack(side="right")
        self.message = StringVar(master=self.window)
        ttk.Label(frame, textvariable=self.message, wraplength=740).pack(anchor="w")
        self.refresh()

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for key, entry in sorted(self.records.items()):
            self.tree.insert("", "end", values=(key, entry["document"], entry["name"]))

    def select(self, _event=None):
        selected = self.tree.selection()
        if selected:
            failure, document, name = self.tree.item(selected[0], "values")
            self.failure.set(failure)
            self.document.set(document)
            self.name.set(name)

    def save(self):
        try:
            self.records = save_debtor(self.failure.get(), self.document.get(), self.name.get(),
                                       self.path)
        except (InputFileError, OSError) as exc:
            self.message.set(str(exc))
            return
        self.refresh()
        self.message.set("Sacado salvo. Prepare um novo Excel para aplicar o cadastro.")
        self.on_saved()
