"""Local ttk styling with generated decoration and no extra runtime packages."""

from __future__ import annotations

import struct
import sys
import zlib
from functools import lru_cache
from tkinter import PhotoImage, font, ttk

COLORS = {
    "background": "#F4F7FC",
    "surface": "#FFFFFF",
    "text": "#172B4D",
    "muted": "#526277",
    "header": "#163C75",
    "primary": "#2563EB",
    "hover": "#1D4ED8",
    "border": "#CCD6E6",
    "disabled": "#E5EAF2",
    "info": "#1D4ED8",
    "success": "#166534",
    "warning": "#92400E",
    "error": "#B91C1C",
}


@lru_cache(maxsize=64)
def _rounded_data(fill, border, radius=10, width=32, height=32):
    """Small antialiased nine-slice tile, generated with the standard library."""

    def rgb(value):
        return tuple(int(value[i : i + 2], 16) for i in (1, 3, 5))

    def inside(x, y, inset):
        r = radius - inset
        cx = min(max(x, radius), width - radius)
        cy = min(max(y, radius), height - radius)
        return (x - cx) ** 2 + (y - cy) ** 2 <= r * r

    surface, edge = rgb(fill), rgb(border)
    raw = bytearray()
    for y in range(height):
        raw.append(0)
        for x in range(width):
            # Straight edges/interiors need no supersampling, even in a wide tile.
            if radius <= x < width - radius:
                raw.extend((*(edge if y in (0, height - 1) else surface), 255))
                continue
            if radius <= y < height - radius:
                raw.extend((*(edge if x in (0, width - 1) else surface), 255))
                continue
            count, inner = 0, 0
            for sy in range(4):
                for sx in range(4):
                    px, py = x + (sx + 0.5) / 4, y + (sy + 0.5) / 4
                    count += inside(px, py, 0)
                    inner += inside(px, py, 1)
            color = tuple(
                round((surface[i] * inner + edge[i] * (count - inner)) / count) if count else 0
                for i in range(3)
            )
            raw.extend((*color, round(255 * count / 16)))

    def chunk(kind, data):
        return (
            struct.pack("!I", len(data)) + kind + data + struct.pack("!I", zlib.crc32(kind + data))
        )

    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack("!2I5B", width, height, 8, 6, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(bytes(raw)))
        + chunk(b"IEND", b"")
    )
    return png


def _rounded_image(root, fill, border, radius=10, width=256, height=64):
    return PhotoImage(
        master=root, data=_rounded_data(fill, border, radius, width, height), format="png"
    )


def configure_theme(root):
    style = ttk.Style(root)
    style.theme_use("clam")
    available = set(font.families(root))
    preferred = (
        ("Segoe UI", "Noto Sans", "DejaVu Sans")
        if sys.platform == "win32"
        else ("Noto Sans", "Ubuntu Sans", "DejaVu Sans")
    )
    family = next(
        (name for name in preferred if name in available),
        font.nametofont("TkDefaultFont").actual("family"),
    )
    fonts = {
        "body": font.Font(root=root, family=family, size=11),
        "strong": font.Font(root=root, family=family, size=11, weight="bold"),
        "title": font.Font(root=root, family=family, size=18, weight="bold"),
        "section": font.Font(root=root, family=family, size=13, weight="bold"),
        "small": font.Font(root=root, family=family, size=10),
    }
    c = COLORS
    root.configure(background=c["background"])
    root.option_add("*TCombobox*Listbox.font", fonts["body"])
    style.configure(
        ".",
        font=fonts["body"],
        background=c["surface"],
        foreground=c["text"],
        bordercolor=c["border"],
        lightcolor=c["surface"],
        darkcolor=c["border"],
    )
    style.configure("TFrame", background=c["surface"])
    style.configure("Shell.TFrame", background=c["background"])
    style.configure("Header.TFrame", background=c["header"])
    style.configure(
        "Header.TLabel", background=c["header"], foreground="white", font=fonts["title"]
    )
    style.configure("TLabel", background=c["surface"])
    style.configure("Title.TLabel", font=fonts["title"])
    style.configure("PageTitle.TLabel", font=fonts["title"], background=c["background"])
    style.configure("PageHelp.TLabel", foreground=c["muted"], background=c["background"])
    style.configure(
        "HeaderHelp.TLabel", foreground="#DCEAFF", background=c["header"], font=fonts["small"]
    )
    style.configure("Section.TLabel", font=fonts["section"])
    style.configure("Strong.TLabel", font=fonts["strong"])
    style.configure("Help.TLabel", foreground=c["muted"])
    style.configure("Shell.TLabel", background=c["background"], foreground=c["muted"])
    style.configure("Metric.TFrame", background=c["background"])
    style.configure(
        "MetricValue.TLabel",
        background=c["background"],
        foreground=c["header"],
        font=fonts["section"],
    )
    style.configure("MetricCaption.TLabel", background=c["background"], foreground=c["muted"])
    style.configure("TButton", padding=(14, 9), relief="flat", focusthickness=0)
    style.map(
        "TButton",
        background=[("disabled", c["disabled"]), ("active", "#EAF1FF")],
        foreground=[("disabled", c["muted"])],
        bordercolor=[("focus", c["primary"])],
    )
    for name in ("Primary.TButton", "ActiveStep.TButton"):
        style.configure(name, background=c["primary"], foreground="white")
        style.map(
            name,
            background=[
                ("disabled", c["disabled"]),
                ("pressed", c["header"]),
                ("active", c["hover"]),
            ],
            foreground=[("disabled", c["muted"]), ("!disabled", "white")],
            bordercolor=[("focus", c["header"])],
        )
    style.configure("Step.TButton", padding=(10, 10))
    style.configure("ActiveStep.TButton", padding=(10, 10))
    for widget in ("TEntry", "TCombobox"):
        style.configure(widget, padding=(8, 6), fieldbackground=c["surface"])
        style.map(
            widget,
            fieldbackground=[("disabled", c["disabled"]), ("readonly", c["background"])],
            foreground=[("disabled", c["muted"])],
            bordercolor=[("focus", c["primary"])],
        )
    style.configure(
        "Treeview",
        font=fonts["body"],
        rowheight=fonts["body"].metrics("linespace") + 14,
        background=c["surface"],
        fieldbackground=c["surface"],
        borderwidth=0,
    )
    style.configure("Treeview.Heading", font=fonts["strong"], padding=(8, 8))
    style.configure("Treeview.Heading", background="#F2F6FC", relief="flat", borderwidth=0)
    style.layout(
        "Treeview.Heading",
        [
            (
                "Treeheading.cell",
                {
                    "sticky": "nswe",
                    "children": [
                        (
                            "Treeheading.padding",
                            {
                                "sticky": "nswe",
                                "children": [("Treeheading.text", {"sticky": "we"})],
                            },
                        )
                    ],
                },
            )
        ],
    )
    style.map(
        "Treeview",
        background=[("selected", "#EAF2FF")],
        foreground=[("selected", c["header"])],
    )
    style.configure("Horizontal.TProgressbar", background=c["primary"], borderwidth=0)
    for kind in ("info", "success", "warning", "error"):
        style.configure(f"{kind}.TLabel", foreground=c[kind], font=fonts["strong"])
    # Only decorative backgrounds change: buttons retain ttk keyboard/state behavior.
    images = []
    for name, fill, edge, hover, text in (
        ("TButton", c["surface"], c["border"], "#F0F5FF", c["text"]),
        ("Header.TButton", c["surface"], c["surface"], "#EAF2FF", c["header"]),
        ("Primary.TButton", c["primary"], c["primary"], c["hover"], "#FFFFFF"),
        ("Step.TButton", c["background"], c["background"], "#EAF2FF", c["muted"]),
        ("ActiveStep.TButton", c["background"], c["background"], "#DFEAFF", c["header"]),
        ("Quiet.TButton", c["surface"], c["surface"], "#F0F5FF", c["muted"]),
        ("Footer.TButton", c["background"], c["background"], "#EAF2FF", c["muted"]),
    ):
        normal = _rounded_image(root, fill, edge)
        active = _rounded_image(root, hover, edge)
        focused = _rounded_image(
            root, fill, "#93B4FF" if name == "Primary.TButton" else c["primary"]
        )
        disabled = _rounded_image(root, c["disabled"], c["disabled"])
        images.extend((normal, active, focused, disabled))
        element = f"Nox.{name}.background"
        # A ttk element name is permanent for the interpreter's lifetime; there
        # is no "delete" to undo it. Guarding the (re)creation keeps repeated
        # calls on the same interpreter safe without changing any rendering.
        if element not in style.element_names():
            style.element_create(
                element,
                "image",
                normal,
                ("disabled", disabled),
                ("focus", focused),
                ("active", active),
                border=11,
                # ttk replicates the center/edges: wide tiles reduce native paint work.
                # The decorative image must not enlarge the control's minimum size.
                width=32,
                height=32,
                sticky="nsew",
            )
        style.layout(
            name,
            [
                (
                    element,
                    {
                        "sticky": "nsew",
                        "children": [
                            (
                                "Button.padding",
                                {
                                    "sticky": "nsew",
                                    "children": [("Button.label", {"sticky": "nsew"})],
                                },
                            )
                        ],
                    },
                )
            ],
        )
        style.configure(
            name,
            foreground=text,
            padding=(14, 5),
            background=(
                c["header"]
                if name == "Header.TButton"
                else c["background"]
                if name
                in {"Primary.TButton", "Step.TButton", "ActiveStep.TButton", "Footer.TButton"}
                else c["surface"]
            ),
        )
        style.map(name, background=[])
        style.map(name, foreground=[("disabled", c["muted"]), ("!disabled", text)])
    card = _rounded_image(root, c["surface"], c["border"], radius=14, width=40, height=40)
    images.append(card)
    corners = []
    for x, y in ((0, 0), (25, 0), (0, 25), (25, 25)):
        corner = PhotoImage(master=root, width=15, height=15)
        root.tk.call(str(corner), "copy", str(card), "-from", x, y, x + 15, y + 15)
        corners.append(corner)
    images.extend(corners)
    root._nox_card_corners = corners
    for widget in ("TEntry", "TCombobox"):
        field = _rounded_image(root, "#F8FAFD", c["border"], radius=8)
        focus = _rounded_image(root, "#F8FAFD", c["primary"], radius=8)
        disabled = _rounded_image(root, c["disabled"], c["disabled"], radius=8)
        images.extend((field, focus, disabled))
        element = f"Nox.{widget}.field"
        if element not in style.element_names():
            style.element_create(
                element,
                "image",
                field,
                ("disabled", disabled),
                ("focus", focus),
                border=9,
                width=32,
                height=32,
                sticky="nsew",
            )
        prefix = "Entry" if widget == "TEntry" else "Combobox"
        children = [
            (
                f"{prefix}.padding",
                {"sticky": "nsew", "children": [(f"{prefix}.textarea", {"sticky": "nsew"})]},
            )
        ]
        if widget == "TCombobox":
            children.insert(0, ("Combobox.downarrow", {"side": "right", "sticky": "ns"}))
        style.layout(widget, [(element, {"sticky": "nsew", "children": children})])
        style.configure(widget, padding=(10, 6), background=c["surface"])
    for axis in ("Vertical", "Horizontal"):
        style.layout(
            f"{axis}.TScrollbar",
            [
                (
                    f"{axis}.Scrollbar.trough",
                    {
                        "sticky": "ns" if axis == "Vertical" else "we",
                        "children": [
                            (f"{axis}.Scrollbar.thumb", {"expand": "1", "sticky": "nsew"})
                        ],
                    },
                )
            ],
        )
        style.configure(
            f"{axis}.TScrollbar",
            width=9,
            arrowsize=9,
            borderwidth=0,
            background="#CBD5E1",
            troughcolor=c["background"],
            relief="flat",
        )
    root._nox_theme_images = images
    return fonts
