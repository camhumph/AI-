import argparse
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from PIL import Image, ImageTk


DEFAULT_IMAGE = Path(r"C:\CMS_AI\datasets\major8_full_views\images\train\26036-37_FRAME_ONLY_20260508_FULLTRAIN_LEFT.jpg")
DEFAULT_LABEL = Path(r"C:\CMS_AI\datasets\major8_full_views\labels\train\26036-37_FRAME_ONLY_20260508_FULLTRAIN_LEFT.txt")
DEFAULT_DATA = Path(r"C:\CMS_AI\datasets\major8_full_views\data.yaml")


COLORS = [
    "#1e50ff",
    "#00d2e6",
    "#ffffff",
    "#00dcbe",
    "#142878",
    "#ff5ad2",
    "#ff3c3c",
    "#d2ff00",
]


def parse_names(data_path):
    names = {}
    if not data_path.exists():
        return names
    in_names = False
    for raw in data_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw.rstrip()
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if line.strip() == "names:":
            in_names = True
            continue
        if in_names:
            if not raw.startswith(" ") and not raw.startswith("\t"):
                break
            if ":" in line:
                k, v = line.split(":", 1)
                try:
                    names[int(k.strip())] = v.strip().strip("'\"")
                except ValueError:
                    pass
    return names


def read_labels(path):
    boxes = []
    if not path.exists():
        return boxes
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        parts = line.split()
        if len(parts) < 5:
            continue
        try:
            boxes.append(
                {
                    "cls": int(float(parts[0])),
                    "xc": float(parts[1]),
                    "yc": float(parts[2]),
                    "w": float(parts[3]),
                    "h": float(parts[4]),
                }
            )
        except ValueError:
            continue
    return boxes


def clamp(v, low, high):
    return max(low, min(high, v))


class YoloBoxEditor:
    def __init__(self, root, image_path, label_path, data_path):
        self.root = root
        self.root.title("CMS YOLO Box Editor")
        self.image_path = Path(image_path)
        self.label_path = Path(label_path)
        self.data_path = Path(data_path)
        self.names = parse_names(self.data_path)
        self.boxes = read_labels(self.label_path)
        self.selected = None
        self.mode = None
        self.drag_start = None
        self.original_box = None
        self.handle_size = 10

        self.image = Image.open(self.image_path).convert("RGB")
        self.img_w, self.img_h = self.image.size
        self.scale = min(1180 / self.img_w, 760 / self.img_h, 1.0)
        self.disp_w = int(self.img_w * self.scale)
        self.disp_h = int(self.img_h * self.scale)
        self.tk_image = ImageTk.PhotoImage(self.image.resize((self.disp_w, self.disp_h), Image.LANCZOS))

        self.build_ui()
        self.draw()

    def build_ui(self):
        top = tk.Frame(self.root)
        top.pack(fill="x", padx=8, pady=6)

        tk.Button(top, text="Save", command=self.save).pack(side="left", padx=(0, 6))
        tk.Button(top, text="Save Train + Val", command=self.save_train_val).pack(side="left", padx=(0, 6))
        tk.Button(top, text="Add Box", command=self.add_box).pack(side="left", padx=(0, 6))
        tk.Button(top, text="Add 8 Starters", command=self.add_starters).pack(side="left", padx=(0, 6))
        tk.Button(top, text="Delete Selected", command=self.delete_selected).pack(side="left", padx=(0, 6))
        tk.Button(top, text="Open Image/Label", command=self.open_pair).pack(side="left", padx=(0, 14))

        self.info = tk.Label(
            top,
            text="Drag inside a box to move. Drag a corner to resize. Click a row or box to select.",
            anchor="w",
        )
        self.info.pack(side="left", fill="x", expand=True)

        body = tk.Frame(self.root)
        body.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(body, width=self.disp_w, height=self.disp_h, bg="#222", cursor="crosshair")
        self.canvas.pack(side="left", padx=8, pady=8)
        self.canvas.bind("<Button-1>", self.on_down)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_up)

        side = tk.Frame(body)
        side.pack(side="left", fill="y", padx=(0, 8), pady=8)

        tk.Label(side, text="Boxes").pack(anchor="w")
        self.listbox = tk.Listbox(side, width=34, height=28)
        self.listbox.pack(fill="y", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self.on_list_select)

        edit = tk.Frame(side)
        edit.pack(fill="x", pady=(8, 0))
        tk.Label(edit, text="Class").pack(anchor="w")
        self.class_var = tk.StringVar()
        self.class_entry = tk.Entry(edit, textvariable=self.class_var, width=8)
        self.class_entry.pack(side="left")
        tk.Button(edit, text="Apply", command=self.apply_class).pack(side="left", padx=6)
        tk.Label(side, text="Class IDs:\n0 top clamp\n1 bottom clamp\n2 A cavity\n3 B core\n4 support\n5 rail\n6 ejector\n7 pin", justify="left").pack(anchor="w", pady=(10, 0))

    def box_to_pixels(self, b):
        x1 = (b["xc"] - b["w"] / 2) * self.img_w * self.scale
        y1 = (b["yc"] - b["h"] / 2) * self.img_h * self.scale
        x2 = (b["xc"] + b["w"] / 2) * self.img_w * self.scale
        y2 = (b["yc"] + b["h"] / 2) * self.img_h * self.scale
        return x1, y1, x2, y2

    def pixels_to_box(self, x1, y1, x2, y2):
        x1, x2 = sorted((clamp(x1, 0, self.disp_w), clamp(x2, 0, self.disp_w)))
        y1, y2 = sorted((clamp(y1, 0, self.disp_h), clamp(y2, 0, self.disp_h)))
        return {
            "xc": ((x1 + x2) / 2) / (self.img_w * self.scale),
            "yc": ((y1 + y2) / 2) / (self.img_h * self.scale),
            "w": max(1, x2 - x1) / (self.img_w * self.scale),
            "h": max(1, y2 - y1) / (self.img_h * self.scale),
        }

    def draw(self):
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)
        self.listbox.delete(0, "end")

        for i, b in enumerate(self.boxes):
            name = self.names.get(b["cls"], str(b["cls"]))
            self.listbox.insert("end", f"{i + 1}. {b['cls']} {name}")
            x1, y1, x2, y2 = self.box_to_pixels(b)
            color = COLORS[b["cls"] % len(COLORS)]
            width = 5 if i == self.selected else 3
            self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=width)
            self.canvas.create_rectangle(x1, max(0, y1 - 24), x1 + min(260, 9 * len(name) + 16), y1, fill=color, outline=color)
            self.canvas.create_text(x1 + 4, max(0, y1 - 22), text=name, anchor="nw", fill="#001050", font=("Arial", 14, "bold"))
            if i == self.selected:
                for hx, hy in [(x1, y1), (x2, y1), (x1, y2), (x2, y2)]:
                    s = self.handle_size
                    self.canvas.create_rectangle(hx - s, hy - s, hx + s, hy + s, fill="#ffff00", outline="#000")
                self.class_var.set(str(b["cls"]))

        if self.selected is not None:
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(self.selected)
            self.listbox.see(self.selected)

    def hit_test(self, x, y):
        for i in reversed(range(len(self.boxes))):
            x1, y1, x2, y2 = self.box_to_pixels(self.boxes[i])
            corners = {
                "nw": (x1, y1),
                "ne": (x2, y1),
                "sw": (x1, y2),
                "se": (x2, y2),
            }
            for name, (hx, hy) in corners.items():
                if abs(x - hx) <= self.handle_size + 3 and abs(y - hy) <= self.handle_size + 3:
                    return i, name
            if x1 <= x <= x2 and y1 <= y <= y2:
                return i, "move"
        return None, None

    def on_down(self, event):
        idx, mode = self.hit_test(event.x, event.y)
        if idx is None:
            self.selected = None
            self.draw()
            return
        self.selected = idx
        self.mode = mode
        self.drag_start = (event.x, event.y)
        self.original_box = dict(self.boxes[idx])
        self.draw()

    def on_drag(self, event):
        if self.selected is None or self.drag_start is None:
            return
        b = self.original_box
        x1, y1, x2, y2 = self.box_to_pixels(b)
        dx = event.x - self.drag_start[0]
        dy = event.y - self.drag_start[1]

        if self.mode == "move":
            nx1, ny1, nx2, ny2 = x1 + dx, y1 + dy, x2 + dx, y2 + dy
            if nx1 < 0:
                nx2 -= nx1
                nx1 = 0
            if ny1 < 0:
                ny2 -= ny1
                ny1 = 0
            if nx2 > self.disp_w:
                nx1 -= nx2 - self.disp_w
                nx2 = self.disp_w
            if ny2 > self.disp_h:
                ny1 -= ny2 - self.disp_h
                ny2 = self.disp_h
        else:
            nx1, ny1, nx2, ny2 = x1, y1, x2, y2
            if "w" in self.mode:
                nx1 = event.x
            if "e" in self.mode:
                nx2 = event.x
            if "n" in self.mode:
                ny1 = event.y
            if "s" in self.mode:
                ny2 = event.y

        new_vals = self.pixels_to_box(nx1, ny1, nx2, ny2)
        self.boxes[self.selected].update(new_vals)
        self.draw()

    def on_up(self, _event):
        self.drag_start = None
        self.original_box = None
        self.mode = None

    def on_list_select(self, _event):
        sel = self.listbox.curselection()
        if sel:
            self.selected = sel[0]
            self.draw()

    def apply_class(self):
        if self.selected is None:
            return
        try:
            self.boxes[self.selected]["cls"] = int(self.class_var.get())
        except ValueError:
            messagebox.showerror("Bad class", "Class must be a number.")
            return
        self.draw()

    def delete_selected(self):
        if self.selected is None:
            return
        del self.boxes[self.selected]
        self.selected = None
        self.draw()

    def add_box(self):
        try:
            cls = int(self.class_var.get())
        except ValueError:
            cls = 0
        self.boxes.append({"cls": cls, "xc": 0.5, "yc": 0.5, "w": 0.25, "h": 0.10})
        self.selected = len(self.boxes) - 1
        self.draw()

    def add_starters(self):
        if self.boxes and not messagebox.askyesno("Add starters", "This will add 9 starter boxes. Keep going?"):
            return
        starters = [
            {"cls": 0, "xc": 0.50, "yc": 0.13, "w": 0.50, "h": 0.07},
            {"cls": 2, "xc": 0.50, "yc": 0.25, "w": 0.50, "h": 0.12},
            {"cls": 3, "xc": 0.50, "yc": 0.39, "w": 0.50, "h": 0.10},
            {"cls": 4, "xc": 0.50, "yc": 0.51, "w": 0.50, "h": 0.10},
            {"cls": 5, "xc": 0.25, "yc": 0.72, "w": 0.07, "h": 0.25},
            {"cls": 5, "xc": 0.75, "yc": 0.72, "w": 0.07, "h": 0.25},
            {"cls": 7, "xc": 0.50, "yc": 0.76, "w": 0.35, "h": 0.05},
            {"cls": 6, "xc": 0.50, "yc": 0.83, "w": 0.35, "h": 0.06},
            {"cls": 1, "xc": 0.50, "yc": 0.92, "w": 0.50, "h": 0.07},
        ]
        self.boxes.extend(starters)
        self.selected = len(self.boxes) - len(starters)
        self.draw()

    def label_text(self):
        lines = []
        for b in self.boxes:
            lines.append(f"{b['cls']} {b['xc']:.6f} {b['yc']:.6f} {b['w']:.6f} {b['h']:.6f}")
        return "\n".join(lines) + "\n"

    def save(self):
        self.label_path.write_text(self.label_text(), encoding="ascii")
        messagebox.showinfo("Saved", f"Saved:\n{self.label_path}")

    def save_train_val(self):
        self.save()
        parts = list(self.label_path.parts)
        try:
            split_idx = parts.index("train")
            parts[split_idx] = "val"
            val_path = Path(*parts)
            val_path.parent.mkdir(parents=True, exist_ok=True)
            val_path.write_text(self.label_text(), encoding="ascii")
            for cache in [self.label_path.parent.parent / "train.cache", val_path.parent.parent / "val.cache"]:
                if cache.exists():
                    cache.unlink()
            messagebox.showinfo("Saved", f"Saved train + val labels:\n{self.label_path}\n{val_path}")
        except ValueError:
            messagebox.showinfo("Saved", f"Saved:\n{self.label_path}")

    def open_pair(self):
        img = filedialog.askopenfilename(title="Open image", filetypes=[("Images", "*.jpg *.jpeg *.png"), ("All files", "*.*")])
        if not img:
            return
        lab = filedialog.askopenfilename(title="Open YOLO label", filetypes=[("YOLO labels", "*.txt"), ("All files", "*.*")])
        if not lab:
            return
        self.root.destroy()
        main(img, lab, self.data_path)


def main(image_path, label_path, data_path):
    root = tk.Tk()
    app = YoloBoxEditor(root, image_path, label_path, data_path)
    root.mainloop()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--image", default=str(DEFAULT_IMAGE))
    parser.add_argument("--label", default=str(DEFAULT_LABEL))
    parser.add_argument("--data", default=str(DEFAULT_DATA))
    args = parser.parse_args()
    main(args.image, args.label, args.data)
