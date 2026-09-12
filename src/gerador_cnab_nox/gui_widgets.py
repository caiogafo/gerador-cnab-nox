from __future__ import annotations

from tkinter import Canvas, Text, ttk

from .gui_theme import COLORS


def paragraph(parent, text="", *, variable=None, style="Help.TLabel", pady=(0, 10)):
    label = ttk.Label(
        parent,
        text=text,
        textvariable=variable,
        style=style,
        justify="left",
        wraplength=780,
    )
    label.pack(fill="x", anchor="w", pady=pady)
    parent.bind(
        "<Configure>",
        lambda event: label.configure(wraplength=max(100, event.width - 16)),
        add="+",
    )
    return label


class ScrollableBody(ttk.Frame):
    """Only the central content scrolls; navigation and actions stay outside."""

    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = Canvas(self, highlightthickness=0, background=COLORS["background"])
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.content = ttk.Frame(self.canvas, style="Shell.TFrame", padding=(4, 8))
        self.window = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.content.columnconfigure(0, weight=1)
        self.content.bind("<Configure>", self._region)
        self.canvas.bind("<Configure>", self._resize)

    def _region(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _resize(self, event):
        self.canvas.itemconfigure(self.window, width=event.width)

    def contains(self, widget):
        return str(widget).startswith(str(self.content) + ".")

    def reveal(self, widget, *, align_top=False):
        if not self.contains(widget) or not widget.winfo_ismapped():
            return
        self.update_idletasks()
        top = widget.winfo_rooty() - self.canvas.winfo_rooty()
        bottom = top + widget.winfo_height()
        visible = self.canvas.winfo_height()
        total = max(self.content.winfo_height(), visible)
        offset = self.canvas.canvasy(0)
        if align_top or top < 0:
            self.canvas.yview_moveto(max(0, offset + top - 12) / total)
        elif bottom > visible:
            self.canvas.yview_moveto((offset + bottom - visible + 12) / total)

    def wheel(self, event):
        if not self.contains(event.widget):
            return
        # Let native list/text widgets handle their own scrollable contents.
        if event.widget.winfo_class() in {"Text", "Treeview", "TCombobox"}:
            return
        if self.content.winfo_height() <= self.canvas.winfo_height():
            return
        delta = -1 if getattr(event, "num", 0) == 4 else 1
        if getattr(event, "delta", 0):
            delta = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(delta * 3, "units")
        return "break"


class Card(ttk.Frame):
    """Four small rounded corners and solid fills avoid expensive image tiling."""

    def __init__(self, parent):
        super().__init__(parent, style="Shell.TFrame", padding=20)
        self.background = Canvas(self, highlightthickness=0, background=COLORS["background"])
        self.background.place(x=0, y=0, relwidth=1, relheight=1, bordermode="outside")
        self.tk.call("lower", str(self.background))
        self.bind("<Configure>", self._draw, add="+")

    def _draw(self, event):
        canvas, w, h = self.background, event.width, event.height
        canvas.delete("all")
        for inset, color in ((0, COLORS["border"]), (1, COLORS["surface"])):
            canvas.create_rectangle(14, inset, w - 14, h - inset, fill=color, width=0)
            canvas.create_rectangle(inset, 14, w - inset, h - 14, fill=color, width=0)
        for image, (x, y) in zip(
            self.winfo_toplevel()._nox_card_corners,
            ((0, 0), (w - 15, 0), (0, h - 15), (w - 15, h - 15)),
            strict=True,
        ):
            canvas.create_image(x, y, image=image, anchor="nw")


class ResultPanel(Card):
    def __init__(self, parent, body_font):
        super().__init__(parent)
        self.kind = "info"
        self.title = paragraph(self, style="info.TLabel", pady=(0, 8))
        self.body = paragraph(self, style="TLabel")
        self.metrics_frame = ttk.Frame(self)
        self.metric_summary = ""
        self.issue_frame, self.issue_text = self._text_area(body_font, 6)
        self.details_button = ttk.Button(
            self, text="Detalhes técnicos", command=self.toggle_details
        )
        self.detail_frame, self.detail_text = self._text_area(body_font, 6)
        self.details_open = False
        self.details = ""
        self.issues: tuple[str, ...] = ()

    def _text_area(self, body_font, height):
        frame = ttk.Frame(self)
        widget = Text(
            frame,
            height=height,
            wrap="word",
            font=body_font,
            state="disabled",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            padx=10,
            pady=8,
            relief="flat",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["primary"],
            takefocus=True,
        )
        scroll = ttk.Scrollbar(frame, command=widget.yview)
        widget.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        widget.pack(fill="both", expand=True)
        return frame, widget

    @staticmethod
    def write(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")
        widget.yview_moveto(0)

    def show(self, title, body="", *, kind="info", issues=(), details=(), metrics=()):
        self.kind, self.issues = kind, tuple(issues)
        self.title.configure(text=title, style=f"{kind}.TLabel")
        self.body.configure(text=body)
        self.metrics_frame.pack_forget()
        for child in self.metrics_frame.winfo_children():
            child.destroy()
        self.metric_summary = " · ".join(f"{label}: {value}" for label, value in metrics)
        if metrics:
            self.metrics_frame.pack(fill="x", after=self.body, pady=(6, 18))
            for column, (label, value) in enumerate(metrics):
                self.metrics_frame.columnconfigure(column, weight=1, uniform="metrics")
                block = ttk.Frame(self.metrics_frame, style="Metric.TFrame", padding=12)
                block.grid(row=0, column=column, sticky="nsew", padx=(0, 12))
                ttk.Label(block, text=str(value), style="MetricValue.TLabel", wraplength=120).pack(
                    anchor="w"
                )
                ttk.Label(
                    block,
                    text=label if label == "PFMIs" else label.capitalize(),
                    style="MetricCaption.TLabel",
                    wraplength=115,
                ).pack(anchor="w", pady=(5, 0))
        self.issue_frame.pack_forget()
        self.details_button.pack_forget()
        self.detail_frame.pack_forget()
        self.details_open = False
        self.details_button.configure(text="Detalhes técnicos")
        self.write(self.issue_text, "\n\n".join(f"{i}. {v}" for i, v in enumerate(issues, 1)))
        if issues:
            self.issue_text.configure(height=min(8, max(3, len(issues) * 2 - 1)))
            self.issue_frame.pack(fill="x", pady=(0, 10))
        self.details = "\n".join(details)
        self.write(self.detail_text, self.details)
        if details:
            self.details_button.pack(anchor="w", pady=(0, 8))

    def toggle_details(self):
        self.details_open = not self.details_open
        if self.details_open:
            self.detail_frame.pack(fill="x", pady=(0, 12))
        else:
            self.detail_frame.pack_forget()
        self.details_button.configure(
            text="Ocultar detalhes técnicos" if self.details_open else "Detalhes técnicos"
        )
