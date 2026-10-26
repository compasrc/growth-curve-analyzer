#!/usr/bin/env python3
"""
Growth Curve Analyzer
=====================
Processes OD600 data from 96-well plate Excel files and plots blank-corrected
growth curves with 95% confidence intervals.

Requirements (install once):
    pip install pandas openpyxl scipy matplotlib

Excel layout expected:
    - Column B, starting row 45: time values (HH:MM:SS, supports >24 h)
    - Column D (A1) through CU (H12), starting row 45: OD600 readings
    - Reads down until the first blank cell in column B
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from scipy import stats
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from matplotlib.widgets import SpanSelector
import matplotlib.ticker as mticker
import os, math, datetime, json

# ── Constants ─────────────────────────────────────────────────────────────────
ROWS = list('ABCDEFGH')
COLS = list(range(1, 13))
WELL_NAMES = [f"{r}{c}" for r in ROWS for c in COLS]

PALETTE = [
    '#1565C0', '#2E7D32', '#C62828', '#EF6C00',
    '#6A1B9A', '#00838F', '#4E342E', '#37474F',
    '#AD1457', '#283593', '#00695C', '#558B2F',
    '#F57F17', '#0277BD', '#4527A0',
]
BLANK_COLOR = '#9E9E9E'
TEMPLATE_VERSION = 1


# ── Time helpers ──────────────────────────────────────────────────────────────
def well_to_col(well_name: str) -> int:
    row_idx = ROWS.index(well_name[0])
    col_idx = int(well_name[1:]) - 1
    return 3 + row_idx * 12 + col_idx


def to_hours(val) -> float:
    if isinstance(val, datetime.timedelta):
        return val.total_seconds() / 3600.0
    if isinstance(val, datetime.time):
        return val.hour + val.minute / 60.0 + val.second / 3600.0
    if isinstance(val, (int, float)):
        return float(val) * 24.0
    if isinstance(val, str):
        parts = val.strip().split(':')
        if len(parts) >= 2:
            h, m = int(parts[0]), int(parts[1])
            s = float(parts[2]) if len(parts) > 2 else 0.0
            return h + m / 60.0 + s / 3600.0
    return float(val)


def hours_to_hms(x, _pos=None) -> str:
    if x < 0:
        return ''
    total_s = int(round(x * 3600))
    h = total_s // 3600
    m = (total_s % 3600) // 60
    s = total_s % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def compute_shared_x_limits(times: np.ndarray):
    finite = times[np.isfinite(times)]
    if finite.size == 0:
        return None
    x_min = float(np.min(finite))
    x_max = float(np.max(finite))
    if np.isclose(x_min, x_max):
        pad = max(0.25, abs(x_min) * 0.05)
        return (x_min - pad, x_max + pad)
    pad = (x_max - x_min) * 0.02
    return (x_min - pad, x_max + pad)


def compute_shared_y_limits(results: dict, scale: str):
    y_chunks = []
    for r in results.values():
        y_chunks.append(np.asarray(r['mean'], dtype=float))
        y_chunks.append(np.asarray(r['ci_lo'], dtype=float))
        y_chunks.append(np.asarray(r['ci_hi'], dtype=float))
        y_chunks.extend(np.asarray(trace, dtype=float) for trace in r['raw'])

    if not y_chunks:
        return None

    y = np.concatenate(y_chunks)
    y = y[np.isfinite(y)]
    if y.size == 0:
        return None

    if scale == 'log':
        y = y[y > 0]
        if y.size == 0:
            return None
        y_min = float(np.min(y))
        y_max = float(np.max(y))
        if np.isclose(y_min, y_max):
            return (y_min * 0.8, y_max * 1.2)
        return (y_min * 0.8, y_max * 1.2)

    y_min = float(np.min(y))
    y_max = float(np.max(y))
    if np.isclose(y_min, y_max):
        pad = max(0.05, abs(y_min) * 0.05)
        return (y_min - pad, y_max + pad)
    pad = (y_max - y_min) * 0.05
    return (y_min - pad, y_max + pad)


# ── Chart Configuration Popup ─────────────────────────────────────────────────
class ChartConfigPopup(tk.Toplevel):
    """
    Small floating window for toggling chart elements.
    Pass fig, a list of visible axes, and an optional list of annotation artists.
    """
    def __init__(self, parent, fig, axes: list, annotations: list = None, title='Chart Options'):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.attributes('-topmost', True)
        self.fig = fig
        self.axes = [ax for ax in axes if ax.get_visible()]
        self.annotations = annotations or []

        pad = dict(padx=14, pady=5, anchor='w')

        ttk.Label(self, text='Toggle chart elements', font=('Helvetica', 10, 'bold')
                  ).pack(padx=14, pady=(10, 4), anchor='w')
        ttk.Separator(self, orient='horizontal').pack(fill='x', padx=10, pady=(0, 6))

        self._add_toggle('Show legend',
            getter=lambda: self._any(lambda ax: ax.get_legend() and ax.get_legend().get_visible()),
            setter=lambda v: [ax.get_legend().set_visible(v)
                              for ax in self.axes if ax.get_legend()],
            **pad)

        if self.annotations:
            self._add_toggle('Show well labels',
                getter=lambda: self.annotations[0].get_visible(),
                setter=lambda v: [a.set_visible(v) for a in self.annotations],
                **pad)

        self._add_toggle('Show axis labels',
            getter=lambda: self._any(lambda ax: ax.xaxis.label.get_visible()),
            setter=lambda v: [
                (ax.xaxis.label.set_visible(v), ax.yaxis.label.set_visible(v))
                for ax in self.axes],
            **pad)

        self._add_toggle('Show tick labels',
            getter=lambda: self._any(lambda ax: any(
                t.get_visible() for t in ax.get_xticklabels())),
            setter=lambda v: [
                [(t.set_visible(v) for t in ax.get_xticklabels()),
                 (t.set_visible(v) for t in ax.get_yticklabels())]
                for ax in self.axes],
            **pad)

        self._add_toggle('Show grid',
            getter=lambda: self._any(lambda ax: ax.xaxis.get_gridlines()
                                     and ax.xaxis.get_gridlines()[0].get_visible()),
            setter=lambda v: [ax.grid(v, linestyle='--', alpha=0.3)
                              for ax in self.axes],
            **pad)

        ttk.Button(self, text='Close', command=self.destroy).pack(pady=(6, 10))

    def _any(self, fn):
        try:
            return any(fn(ax) for ax in self.axes)
        except Exception:
            return True

    def _add_toggle(self, label, getter, setter, **pack_kw):
        var = tk.BooleanVar(value=getter())
        def on_change():
            setter(var.get())
            self.fig.canvas.draw_idle()
        ttk.Checkbutton(self, text=label, variable=var, command=on_change).pack(**pack_kw)
        # Keep reference so var isn't GC'd
        setattr(self, f'_var_{label}', var)


# ── Export Single Group Dialog ────────────────────────────────────────────────
class ExportSingleDialog(tk.Toplevel):
    def __init__(self, parent, results_data: dict, times: np.ndarray,
                 log_scale: bool):
        super().__init__(parent)
        self.title('Export Individual Plot')
        self.resizable(False, False)
        self.attributes('-topmost', True)
        self.results_data = results_data
        self.times = times
        self.log_scale = log_scale

        ttk.Label(self, text='Select a group to export:',
                  font=('Helvetica', 10, 'bold')).pack(padx=14, pady=(12, 6))

        self.listbox = tk.Listbox(self, font=('Helvetica', 10),
                                   selectmode='single', height=min(10, len(results_data)),
                                   width=34, activestyle='dotbox')
        self.listbox.pack(padx=14, pady=(0, 8))

        self._keys = []
        for key, r in results_data.items():
            self.listbox.insert('end', f"  {r['name']}  ({r['n']} wells)")
            self._keys.append(key)
        self.listbox.selection_set(0)

        btn_row = ttk.Frame(self)
        btn_row.pack(padx=14, pady=(0, 12))
        ttk.Button(btn_row, text='Save PNG…', command=self._save).pack(side='left', padx=(0, 6))
        ttk.Button(btn_row, text='Cancel', command=self.destroy).pack(side='left')

    def _save(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning('No Selection', 'Select a group first.', parent=self)
            return
        key = self._keys[sel[0]]
        r = self.results_data[key]

        path = filedialog.asksaveasfilename(
            parent=self,
            title=f'Save — {r["name"]}',
            defaultextension='.png',
            initialfile=r['name'].replace(' ', '_'),
            filetypes=[('PNG image', '*.png'), ('PDF', '*.pdf'),
                       ('SVG', '*.svg'), ('All files', '*.*')])
        if not path:
            return

        hms_fmt = mticker.FuncFormatter(hours_to_hms)
        fig = Figure(figsize=(7, 4.8), dpi=150)
        fig.patch.set_facecolor('white')
        ax = fig.add_subplot(111)
        c = r['color']

        for trace in r['raw']:
            ax.plot(self.times, trace, color=c, alpha=0.20, lw=0.9, zorder=1)
        if r['n'] > 1:
            ax.fill_between(self.times, r['ci_lo'], r['ci_hi'],
                            alpha=0.22, color=c, zorder=2, label='95% CI')
        ax.plot(self.times, r['mean'], color=c, lw=2.5,
                label=f"Mean (n={r['n']})", zorder=3)

        ax.set_title(r['name'], fontsize=13, fontweight='bold', pad=10)
        ax.set_xlabel('Time (HH:MM:SS)', fontsize=11)
        ax.set_ylabel('OD₆₀₀ (blank-corrected)', fontsize=11)
        ax.set_yscale('log' if self.log_scale else 'linear')
        x_limits = compute_shared_x_limits(self.times)
        y_limits = compute_shared_y_limits(
            self.results_data,
            'log' if self.log_scale else 'linear')
        if x_limits is not None:
            ax.set_xlim(*x_limits)
        if y_limits is not None:
            ax.set_ylim(*y_limits)
        ax.xaxis.set_major_formatter(hms_fmt)
        for lbl in ax.get_xticklabels():
            lbl.set_rotation(35)
            lbl.set_ha('right')
        ax.legend(fontsize=9, loc='upper left', framealpha=0.7)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, linestyle='--', alpha=0.3)
        ax.annotate(f"Wells: {', '.join(r['wells'])}",
                   xy=(0.5, -0.22), xycoords='axes fraction',
                   ha='center', fontsize=8, color='#888888')
        fig.tight_layout(rect=[0, 0.05, 1, 1])
        fig.savefig(path, dpi=200, bbox_inches='tight', facecolor='white')

        import matplotlib.pyplot as plt
        plt.close(fig)
        messagebox.showinfo('Saved', f'Saved to:\n{path}', parent=self)
        self.destroy()


# ── Main Application ──────────────────────────────────────────────────────────
class GrowthCurveApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Growth Curve Analyzer")
        self.geometry("1220x880")
        self.minsize(960, 700)
        self.configure(bg='#F0F2F5')

        self.groups: dict = {}
        self.group_order: list = []
        self.active_key = None
        self.excel_path = None
        self.color_idx: int = 0

        self.results_fig = None
        self.results_data: dict = {}
        self.times: np.ndarray = np.array([])
        self.log_scale: bool = False

        # Artists tracked for live config toggles
        self.result_axes: list = []           # all subplot Axes in results fig
        self.result_annotations: list = []    # well-label Annotation objects

        # Generation time
        self.gt_fig = None
        self.gt_canvas_widget = None
        self.gt_span: SpanSelector = None
        self.gt_fit_lines: dict = {}
        self.gt_ax = None
        self.gt_axes: list = []               # for config popup

        self._new_group('BLANK', 'BLANK', BLANK_COLOR)
        self._build_ui()
        self._refresh_groups_panel()
        self._refresh_well_colors()

    # ── Group helpers ─────────────────────────────────────────────────────────
    def _new_group(self, key, name, color):
        self.groups[key] = {'name': name, 'color': color, 'wells': set()}
        self.group_order.append(key)

    def _next_color(self):
        c = PALETTE[self.color_idx % len(PALETTE)]
        self.color_idx += 1
        return c

    def add_group(self):
        n = sum(1 for k in self.group_order if k != 'BLANK') + 1
        key = f'g{n}_{id(object())}'
        self._new_group(key, f'Group {n}', self._next_color())
        self._refresh_groups_panel()
        self.set_active(key)

    def set_active(self, key):
        self.active_key = key
        self._refresh_groups_panel()

    def delete_group(self, key):
        if key == 'BLANK':
            messagebox.showwarning('Cannot Delete', 'The BLANK group cannot be deleted.')
            return
        del self.groups[key]
        self.group_order.remove(key)
        if self.active_key == key:
            self.active_key = None
        self._refresh_groups_panel()
        self._refresh_well_colors()

    def clear_group_wells(self, key):
        self.groups[key]['wells'].clear()
        self._refresh_groups_panel()
        self._refresh_well_colors()

    # ── Template save / load ──────────────────────────────────────────────────
    def save_template(self):
        path = filedialog.asksaveasfilename(
            title='Save Well Plate Template',
            defaultextension='.json',
            filetypes=[('Well Plate Template', '*.json'), ('All files', '*.*')])
        if not path:
            return
        payload = {
            'version': TEMPLATE_VERSION,
            'groups': [
                {'name': self.groups[k]['name'],
                 'color': self.groups[k]['color'],
                 'wells': sorted(self.groups[k]['wells'])}
                for k in self.group_order
            ]
        }
        with open(path, 'w') as fh:
            json.dump(payload, fh, indent=2)
        messagebox.showinfo('Saved', f'Template saved to:\n{os.path.basename(path)}')

    def load_template(self):
        path = filedialog.askopenfilename(
            title='Load Well Plate Template',
            filetypes=[('Well Plate Template', '*.json'), ('All files', '*.*')])
        if not path:
            return
        try:
            with open(path) as fh:
                payload = json.load(fh)
            groups_data = payload.get('groups', [])
        except Exception as exc:
            messagebox.showerror('Load Error', f'Could not read template:\n{exc}')
            return

        if not messagebox.askyesno(
                'Load Template',
                'Loading will replace the current well assignments.\nContinue?'):
            return

        self.groups.clear()
        self.group_order.clear()
        self.active_key = None
        self.color_idx = 0

        for i, g in enumerate(groups_data):
            key = 'BLANK' if g['name'] == 'BLANK' else f'g{i}_{id(object())}'
            color = g.get('color', BLANK_COLOR if g['name'] == 'BLANK' else self._next_color())
            self.groups[key] = {
                'name': g['name'],
                'color': color,
                'wells': set(g.get('wells', []))
            }
            self.group_order.append(key)

        # Always ensure BLANK exists
        if 'BLANK' not in self.groups:
            self._new_group('BLANK', 'BLANK', BLANK_COLOR)

        self._refresh_groups_panel()
        self._refresh_well_colors()

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TNotebook.Tab', font=('Helvetica', 11), padding=[14, 5])
        style.configure('TLabelframe.Label', font=('Helvetica', 10, 'bold'))
        style.configure('TButton', font=('Helvetica', 10))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True, padx=10, pady=10)

        self.tab_setup   = ttk.Frame(self.nb)
        self.tab_results = ttk.Frame(self.nb)
        self.tab_gentime = ttk.Frame(self.nb)

        self.nb.add(self.tab_setup,   text='  Well Plate Setup  ')
        self.nb.add(self.tab_results, text='  Growth Curves  ')
        self.nb.add(self.tab_gentime, text='  Generation Time  ')

        self._build_setup_tab()
        self._build_results_tab()
        self._build_gentime_tab()

    # ── Setup tab ─────────────────────────────────────────────────────────────
    def _build_setup_tab(self):
        tab = self.tab_setup

        left = ttk.LabelFrame(tab, text='Groups', padding=8)
        left.pack(side='left', fill='y', padx=(0, 8), pady=6)
        left.configure(width=245)
        left.pack_propagate(False)

        ttk.Button(left, text='＋  New Group', command=self.add_group).pack(
            fill='x', ipady=5, pady=(0, 4))

        # Template buttons
        tmpl_row = ttk.Frame(left)
        tmpl_row.pack(fill='x', pady=(0, 8))
        ttk.Button(tmpl_row, text='💾 Save Template',
                   command=self.save_template).pack(side='left', fill='x', expand=True, padx=(0, 2))
        ttk.Button(tmpl_row, text='📂 Load Template',
                   command=self.load_template).pack(side='left', fill='x', expand=True, padx=(2, 0))

        sf = ttk.Frame(left)
        sf.pack(fill='both', expand=True)
        self.grp_canvas = tk.Canvas(sf, bg='white', bd=0, highlightthickness=0)
        sb = ttk.Scrollbar(sf, orient='vertical', command=self.grp_canvas.yview)
        self.grp_canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.grp_canvas.pack(side='left', fill='both', expand=True)
        self.grp_inner = ttk.Frame(self.grp_canvas)
        self._grp_win = self.grp_canvas.create_window(
            (0, 0), window=self.grp_inner, anchor='nw')
        self.grp_inner.bind('<Configure>', self._on_grp_resize)
        self.grp_canvas.bind('<Configure>', self._on_grp_canvas_resize)

        ttk.Label(left,
                  text='① Add groups  ② Click to select\n③ Click wells to assign  ④ Load & Process',
                  font=('Helvetica', 9), foreground='#777777', justify='left'
                  ).pack(anchor='w', pady=(10, 0))

        right = ttk.Frame(tab)
        right.pack(side='right', fill='both', expand=True, pady=6)

        plate_lf = ttk.LabelFrame(
            right, text='96-Well Plate  —  select a group, then click wells to assign',
            padding=10)
        plate_lf.pack(fill='both', expand=True, pady=(0, 8))
        self._build_plate(plate_lf)

        file_row = ttk.Frame(right)
        file_row.pack(fill='x')

        ttk.Label(file_row, text='Excel file:', font=('Helvetica', 10)).pack(side='left')
        self.file_lbl = ttk.Label(file_row, text='  not selected',
                                   foreground='#AAAAAA', font=('Helvetica', 10))
        self.file_lbl.pack(side='left')
        ttk.Button(file_row, text='Browse…', command=self.browse).pack(
            side='left', padx=(14, 4))

        go_btn = tk.Button(
            file_row, text='  Process & Plot  ▶  ',
            command=self.process,
            bg='#1565C0', fg='white', font=('Helvetica', 11, 'bold'),
            relief='flat', padx=10, pady=6, cursor='hand2',
            activebackground='#0D47A1', activeforeground='white')
        go_btn.pack(side='right')

    def _on_grp_resize(self, _e):
        self.grp_canvas.configure(scrollregion=self.grp_canvas.bbox('all'))

    def _on_grp_canvas_resize(self, e):
        self.grp_canvas.itemconfig(self._grp_win, width=e.width)

    def _build_plate(self, parent):
        R, S, MT, ML = 17, 43, 38, 52
        W = ML + 11 * S + R + 14
        H = MT + 7 * S + R + 14

        container = ttk.Frame(parent)
        container.pack(expand=True)

        c = tk.Canvas(container, width=W, height=H, bg='#F0F4F8',
                      highlightthickness=1, highlightbackground='#C8CDD5')
        c.pack()
        self.plate = c

        c.create_rectangle(ML - R - 8, MT - R - 8,
                           ML + 11 * S + R + 8, MT + 7 * S + R + 8,
                           fill='#E4E8EE', outline='#C0C8D0', width=1)

        for ci in range(12):
            c.create_text(ML + ci * S, MT - R - 16, text=str(ci + 1),
                         font=('Helvetica', 9, 'bold'), fill='#444444')

        self.ovals = {}
        self.oval_texts = {}

        for ri, row in enumerate(ROWS):
            c.create_text(ML - R - 20, MT + ri * S, text=row,
                         font=('Helvetica', 10, 'bold'), fill='#444444')
            for ci in range(12):
                x, y = ML + ci * S, MT + ri * S
                wn = f"{row}{ci + 1}"
                oval = c.create_oval(x - R, y - R, x + R, y + R,
                                    fill='white', outline='#B8B8B8', width=1.5)
                txt = c.create_text(x, y, text='', font=('Helvetica', 7, 'bold'), fill='white')
                self.ovals[wn] = oval
                self.oval_texts[wn] = txt
                for item in [oval, txt]:
                    c.tag_bind(item, '<Button-1>', lambda e, w=wn: self.click_well(w))
                    c.tag_bind(item, '<Enter>', lambda e: c.config(cursor='hand2'))
                    c.tag_bind(item, '<Leave>', lambda e: c.config(cursor=''))

    # ── Growth Curves tab ─────────────────────────────────────────────────────
    def _build_results_tab(self):
        tab = self.tab_results

        bar = ttk.Frame(tab)
        bar.pack(fill='x', padx=8, pady=(8, 4))

        ttk.Button(bar, text='Export All (PNG)',
                   command=self.export_plot).pack(side='left')
        ttk.Button(bar, text='Export Individual…',
                   command=self.export_individual).pack(side='left', padx=(6, 0))
        ttk.Button(bar, text='Save CSV',
                   command=self.save_csv).pack(side='left', padx=(6, 0))

        self.log_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar, text='Log y-axis',
                        variable=self.log_var,
                        command=self._toggle_yscale).pack(side='left', padx=(18, 0))

        self.status_lbl = ttk.Label(
            bar, text='Run Setup → Process & Plot to see results.',
            foreground='#999999', font=('Helvetica', 10))
        self.status_lbl.pack(side='left', padx=18)

        outer = ttk.Frame(tab)
        outer.pack(fill='both', expand=True, padx=8, pady=(0, 8))

        self.res_canvas = tk.Canvas(outer, bg='white', bd=0, highlightthickness=0)
        vs = ttk.Scrollbar(outer, orient='vertical', command=self.res_canvas.yview)
        hs = ttk.Scrollbar(outer, orient='horizontal', command=self.res_canvas.xview)
        self.res_canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        vs.pack(side='right', fill='y')
        hs.pack(side='bottom', fill='x')
        self.res_canvas.pack(side='left', fill='both', expand=True)

        self.res_inner = ttk.Frame(self.res_canvas)
        self._res_win = self.res_canvas.create_window(
            (0, 0), window=self.res_inner, anchor='nw')
        self.res_inner.bind('<Configure>', lambda e: self.res_canvas.configure(
            scrollregion=self.res_canvas.bbox('all')))
        self.res_canvas.bind('<Configure>', lambda e: self.res_canvas.itemconfig(
            self._res_win, width=e.width))

        for widget in [self.res_canvas, self.res_inner]:
            widget.bind('<MouseWheel>',
                        lambda e: self.res_canvas.yview_scroll(-1*(e.delta//120), 'units'))
            widget.bind('<Button-4>',
                        lambda e: self.res_canvas.yview_scroll(-1, 'units'))
            widget.bind('<Button-5>',
                        lambda e: self.res_canvas.yview_scroll(1, 'units'))

    def _toggle_yscale(self):
        self.log_scale = self.log_var.get()
        if self.results_fig is None:
            return
        scale = 'log' if self.log_scale else 'linear'
        x_limits = compute_shared_x_limits(self.times)
        y_limits = compute_shared_y_limits(self.results_data, scale)
        for ax in self.results_fig.get_axes():
            if not ax.get_visible():
                continue
            ax.set_yscale(scale)
            if x_limits is not None:
                ax.set_xlim(*x_limits)
            if y_limits is not None:
                ax.set_ylim(*y_limits)
        self.results_fig.canvas.draw_idle()

    # ── Generation Time tab ───────────────────────────────────────────────────
    def _build_gentime_tab(self):
        tab = self.tab_gentime

        info = ttk.Frame(tab)
        info.pack(fill='x', padx=10, pady=(10, 0))
        ttk.Label(info,
                  text='Drag across the plot to select the exponential phase.  Release to fit.  Drag again to refit.',
                  font=('Helvetica', 10), foreground='#555555').pack(side='left')

        ttk.Button(info, text='Export Plot (PNG)',
                   command=self.export_gt_plot).pack(side='right')
        ttk.Button(info, text='Export Table (CSV)',
                   command=self.export_gentime_csv).pack(side='right', padx=(0, 6))

        self.gt_window_lbl = ttk.Label(
            tab, text='Window: — drag the plot to select',
            font=('Helvetica', 10, 'italic'), foreground='#777777')
        self.gt_window_lbl.pack(anchor='w', padx=12, pady=(4, 0))

        self.gt_plot_outer = ttk.Frame(tab)
        self.gt_plot_outer.pack(fill='both', expand=True, padx=10, pady=(4, 6))

        self.gt_placeholder = ttk.Label(
            self.gt_plot_outer,
            text='Process growth curve data first, then come back here.',
            font=('Helvetica', 11), foreground='#AAAAAA')
        self.gt_placeholder.pack(expand=True)

        tbl_frame = ttk.LabelFrame(tab, text='Results', padding=8)
        tbl_frame.pack(fill='x', padx=10, pady=(0, 10))

        cols = ('Group', 'Wells', 'μ (h⁻¹)', 'Doubling Time (h)',
                'Doubling Time (min)', 'R²', 'Window')
        self.gt_tree = ttk.Treeview(tbl_frame, columns=cols, show='headings', height=5)
        widths = [160, 130, 90, 140, 140, 70, 200]
        for col, w in zip(cols, widths):
            self.gt_tree.heading(col, text=col)
            self.gt_tree.column(col, width=w, anchor='center')

        vsb = ttk.Scrollbar(tbl_frame, orient='vertical', command=self.gt_tree.yview)
        hsb = ttk.Scrollbar(tbl_frame, orient='horizontal', command=self.gt_tree.xview)
        self.gt_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.gt_tree.pack(fill='x')

        self.gt_results: list = []

    def _init_gentime_plot(self):
        for w in self.gt_plot_outer.winfo_children():
            w.destroy()
        if self.gt_fig is not None:
            import matplotlib.pyplot as plt
            plt.close(self.gt_fig)

        times = self.times
        hms_fmt = mticker.FuncFormatter(hours_to_hms)

        fig = Figure(figsize=(10, 5), dpi=96)
        fig.patch.set_facecolor('white')
        self.gt_fig = fig

        ax = fig.add_subplot(111)
        self.gt_ax = ax
        self.gt_axes = [ax]

        self.gt_fit_lines = {}

        for key, r in self.results_data.items():
            mean_od = r['mean']
            pos = mean_od > 0
            if pos.any():
                ax.plot(times[pos], np.log(mean_od[pos]),
                        color=r['color'], lw=2.0, label=r['name'], zorder=2)

            fit_line, = ax.plot([], [], color=r['color'],
                                lw=2.5, ls='--', zorder=3, alpha=0.0)
            self.gt_fit_lines[key] = fit_line

        ax.set_xlabel('Time (HH:MM:SS)', fontsize=11)
        ax.set_ylabel('ln(OD₆₀₀)', fontsize=11)
        ax.set_title('Drag to select the exponential (log-linear) phase', fontsize=11)
        ax.xaxis.set_major_formatter(hms_fmt)
        for lbl in ax.get_xticklabels():
            lbl.set_rotation(35)
            lbl.set_ha('right')
        ax.legend(fontsize=9, loc='upper left', framealpha=0.8)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, linestyle='--', alpha=0.3)

        fig.tight_layout(rect=[0, 0.04, 1, 1])

        canvas = FigureCanvasTkAgg(fig, master=self.gt_plot_outer)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        self.gt_canvas_widget = canvas

        # Toolbar row with ⚙ Configure button
        toolbar_row = ttk.Frame(self.gt_plot_outer)
        toolbar_row.pack(fill='x')
        toolbar = NavigationToolbar2Tk(canvas, toolbar_row)
        toolbar.pack(side='left')
        tk.Button(toolbar_row, text='⚙ Configure',
                  command=lambda: ChartConfigPopup(self, fig, self.gt_axes,
                                                   annotations=None,
                                                   title='Generation Time Chart Options'),
                  relief='flat', bg='#EEEEEE', font=('Helvetica', 9),
                  padx=6, pady=2, cursor='hand2').pack(side='left', padx=(6, 0))
        toolbar.update()

        try:
            self.gt_span = SpanSelector(
                ax, self._on_span_select, 'horizontal',
                useblit=True,
                props=dict(alpha=0.18, facecolor='#1565C0'),
                interactive=True, drag_from_anywhere=True)
        except TypeError:
            self.gt_span = SpanSelector(
                ax, self._on_span_select, 'horizontal',
                useblit=True,
                props=dict(alpha=0.18, facecolor='#1565C0'))

    def _on_span_select(self, t_start: float, t_end: float):
        if t_end <= t_start or not self.results_data:
            return

        times = self.times
        mask = (times >= t_start) & (times <= t_end)
        if mask.sum() < 3:
            self.gt_window_lbl.config(
                text='Window too narrow — fewer than 3 points. Widen the selection.',
                foreground='#C62828')
            return

        t_win = times[mask]
        self.gt_window_lbl.config(
            text=f'Window: {hours_to_hms(t_start)}  →  {hours_to_hms(t_end)}  '
                 f'({mask.sum()} time points)',
            foreground='#1565C0')

        for row in self.gt_tree.get_children():
            self.gt_tree.delete(row)
        self.gt_results.clear()

        win_str = f"{hours_to_hms(t_start)} – {hours_to_hms(t_end)}"

        for key, r in self.results_data.items():
            od_win = r['mean'][mask]
            pos_win = od_win > 0
            t_pos = t_win[pos_win]
            ln_od = np.log(od_win[pos_win])
            fit_line = self.gt_fit_lines[key]

            if pos_win.sum() >= 3:
                reg = stats.linregress(t_pos, ln_od)
                mu, r2 = reg.slope, reg.rvalue ** 2

                if mu > 0:
                    dbl_h = np.log(2) / mu
                    fit_line.set_data([t_start, t_end],
                                      [reg.intercept + mu * t_start,
                                       reg.intercept + mu * t_end])
                    fit_line.set_alpha(1.0)
                    row_vals = (r['name'], ', '.join(r['wells']),
                                f'{mu:.4f}', f'{dbl_h:.3f}',
                                f'{dbl_h * 60:.1f}', f'{r2:.4f}', win_str)
                else:
                    fit_line.set_alpha(0.0)
                    row_vals = (r['name'], ', '.join(r['wells']),
                                '(no growth)', '—', '—', f'{r2:.4f}', win_str)
            else:
                fit_line.set_alpha(0.0)
                row_vals = (r['name'], ', '.join(r['wells']),
                            '(insufficient data)', '—', '—', '—', win_str)

            self.gt_tree.insert('', 'end', values=row_vals)
            self.gt_results.append(row_vals)

        self.gt_canvas_widget.draw_idle()

    # ── Group panel rendering ─────────────────────────────────────────────────
    def _refresh_groups_panel(self):
        for w in self.grp_inner.winfo_children():
            w.destroy()

        for key in self.group_order:
            g = self.groups[key]
            is_active = (key == self.active_key)
            bg = '#E3F2FD' if is_active else 'white'
            border = '#1565C0' if is_active else '#E0E0E0'

            row = tk.Frame(self.grp_inner, bg=bg,
                           highlightbackground=border, highlightthickness=2)
            row.pack(fill='x', padx=3, pady=2)

            swatch = tk.Label(row, bg=g['color'], width=3, cursor='hand2')
            swatch.pack(side='left', fill='y')

            var = tk.StringVar(value=g['name'])
            entry = tk.Entry(row, textvariable=var,
                             font=('Helvetica', 10), relief='flat', bg=bg, bd=0)
            entry.pack(side='left', fill='x', expand=True, padx=6, pady=6)
            entry.bind('<FocusOut>', lambda e, k=key, v=var: self._rename(k, v.get()))
            entry.bind('<Return>',   lambda e, k=key, v=var: self._rename(k, v.get()))

            badge = tk.Label(row, text=f'({len(g["wells"])})',
                             font=('Helvetica', 9), fg='#888888', bg=bg)
            badge.pack(side='left', padx=2)

            clr = tk.Label(row, text='⊘', font=('Helvetica', 10),
                           fg='#CCCCCC', cursor='hand2', bg=bg)
            clr.pack(side='right', padx=2)
            clr.bind('<Button-1>', lambda e, k=key: self.clear_group_wells(k))
            clr.bind('<Enter>',    lambda e, w=clr: w.config(fg='#EF6C00'))
            clr.bind('<Leave>',    lambda e, w=clr: w.config(fg='#CCCCCC'))

            if key != 'BLANK':
                x_lbl = tk.Label(row, text='✕', font=('Helvetica', 10),
                                 fg='#CCCCCC', cursor='hand2', bg=bg)
                x_lbl.pack(side='right', padx=4)
                x_lbl.bind('<Button-1>', lambda e, k=key: self.delete_group(k))
                x_lbl.bind('<Enter>',    lambda e, w=x_lbl: w.config(fg='#E53935'))
                x_lbl.bind('<Leave>',    lambda e, w=x_lbl: w.config(fg='#CCCCCC'))

            for widget in [row, swatch, badge]:
                widget.bind('<Button-1>', lambda e, k=key: self.set_active(k))

    def _rename(self, key, name):
        name = name.strip()
        if name and key in self.groups:
            self.groups[key]['name'] = name

    # ── Well plate interaction ────────────────────────────────────────────────
    def click_well(self, well):
        if not self.active_key:
            messagebox.showinfo('Select a Group',
                                'Create and select a group first, then click wells to assign.')
            return
        g = self.groups[self.active_key]
        if well in g['wells']:
            g['wells'].discard(well)
        else:
            for grp in self.groups.values():
                grp['wells'].discard(well)
            g['wells'].add(well)
        self._refresh_well_colors()
        self._refresh_groups_panel()

    def _refresh_well_colors(self):
        well_map = {}
        for k, g in self.groups.items():
            for w in g['wells']:
                well_map[w] = k
        for wn in WELL_NAMES:
            if wn in well_map:
                g = self.groups[well_map[wn]]
                self.plate.itemconfig(self.ovals[wn],
                                     fill=g['color'], outline='#444444', width=2.0)
                self.plate.itemconfig(self.oval_texts[wn],
                                     text=g['name'][:4], fill='white')
            else:
                self.plate.itemconfig(self.ovals[wn],
                                     fill='white', outline='#B8B8B8', width=1.5)
                self.plate.itemconfig(self.oval_texts[wn], text='')

    # ── File browse ───────────────────────────────────────────────────────────
    def browse(self):
        path = filedialog.askopenfilename(
            title='Select Excel File',
            filetypes=[('Excel files', '*.xlsx *.xls'), ('All files', '*.*')])
        if path:
            self.excel_path = path
            self.file_lbl.config(
                text=f'  {os.path.basename(path)}', foreground='#222222')

    # ── Data processing ───────────────────────────────────────────────────────
    def process(self):
        if not self.excel_path:
            messagebox.showerror('No File', 'Please select an Excel file first.')
            return

        data_groups = {k: g for k, g in self.groups.items()
                       if k != 'BLANK' and g['wells']}
        blank_wells = self.groups['BLANK']['wells']

        if not data_groups:
            messagebox.showerror('No Groups',
                                 'Add at least one non-BLANK group with wells assigned.')
            return

        if not blank_wells:
            if not messagebox.askyesno(
                    'No BLANK Wells',
                    'No wells are assigned to the BLANK group.\n'
                    'Continue without blank subtraction?'):
                return

        try:
            self._analyze(data_groups, blank_wells)
        except Exception as exc:
            messagebox.showerror('Processing Error', f'Failed to process data:\n\n{exc}')
            raise

    def _analyze(self, data_groups, blank_wells):
        df = pd.read_excel(self.excel_path, header=None, engine='openpyxl')

        START = 44
        T_COL = 1

        end = START
        while end < len(df):
            val = df.iat[end, T_COL]
            if val is None or (isinstance(val, float) and np.isnan(val)):
                break
            if isinstance(val, str) and val.strip() == '':
                break
            end += 1

        if end == START:
            raise ValueError(
                'No data found starting at B45.\n'
                'Verify the file has time values in column B beginning at row 45.')

        n_tp = end - START
        raw_times = df.iloc[START:end, T_COL].values

        times_list = []
        for v in raw_times:
            try:
                times_list.append(to_hours(v))
            except Exception:
                times_list.append(float('nan'))
        times = np.array(times_list)

        def well_data(wn):
            col = well_to_col(wn)
            arr = df.iloc[START:end, col].values
            return pd.to_numeric(arr, errors='coerce').astype(float)

        if blank_wells:
            blank_arr = np.vstack([well_data(w) for w in blank_wells])
            blank_mean = np.nanmean(blank_arr, axis=0)
        else:
            blank_mean = np.zeros(n_tp)

        results = {}
        for key, g in data_groups.items():
            wells = sorted(g['wells'])
            corrected = np.vstack([well_data(w) - blank_mean for w in wells])
            n = len(wells)
            mean = np.nanmean(corrected, axis=0)

            if n > 1:
                sem = stats.sem(corrected, axis=0, nan_policy='omit')
                tc = stats.t.ppf(0.975, df=n - 1)
                ci_lo = mean - tc * sem
                ci_hi = mean + tc * sem
            else:
                ci_lo = ci_hi = mean.copy()

            results[key] = {
                'name': g['name'], 'color': g['color'],
                'wells': wells, 'n': n,
                'mean': mean, 'ci_lo': ci_lo, 'ci_hi': ci_hi,
                'raw': corrected,
            }

        self.results_data = results
        self.times = times

        self._render_plots(results, times)
        self._init_gentime_plot()
        self.nb.select(1)
        self.status_lbl.config(
            text=f'✓  {n_tp} time points · {len(results)} group(s)',
            foreground='#2E7D32')

    # ── Plot rendering ────────────────────────────────────────────────────────
    def _render_plots(self, results, times):
        for w in self.res_inner.winfo_children():
            w.destroy()
        if self.results_fig:
            import matplotlib.pyplot as plt
            plt.close(self.results_fig)
            self.results_fig = None

        n = len(results)
        ncols = min(3, n)
        nrows = math.ceil(n / ncols)

        fig = Figure(figsize=(max(6.5 * ncols, 8), max(4.8 * nrows, 5)), dpi=96)
        fig.patch.set_facecolor('white')
        self.results_fig = fig
        self.result_axes = []
        self.result_annotations = []

        hms_fmt = mticker.FuncFormatter(hours_to_hms)
        scale = 'log' if self.log_scale else 'linear'
        x_limits = compute_shared_x_limits(times)
        y_limits = compute_shared_y_limits(results, scale)

        for idx, (key, r) in enumerate(results.items()):
            ax = fig.add_subplot(nrows, ncols, idx + 1)
            self.result_axes.append(ax)
            c = r['color']

            for trace in r['raw']:
                ax.plot(times, trace, color=c, alpha=0.20, lw=0.9, zorder=1)

            if r['n'] > 1:
                ax.fill_between(times, r['ci_lo'], r['ci_hi'],
                                alpha=0.22, color=c, zorder=2, label='95% CI')

            ax.plot(times, r['mean'], color=c, lw=2.5,
                    label=f"Mean (n={r['n']})", zorder=3)

            ax.set_title(r['name'], fontsize=12, fontweight='bold', pad=8)
            ax.set_xlabel('Time (HH:MM:SS)', fontsize=10)
            ax.set_ylabel('OD₆₀₀ (blank-corrected)', fontsize=10)
            ax.set_yscale(scale)
            if x_limits is not None:
                ax.set_xlim(*x_limits)
            if y_limits is not None:
                ax.set_ylim(*y_limits)
            ax.xaxis.set_major_formatter(hms_fmt)
            for lbl in ax.get_xticklabels():
                lbl.set_rotation(35)
                lbl.set_ha('right')
            ax.legend(fontsize=8, loc='upper left', framealpha=0.7)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.grid(True, linestyle='--', alpha=0.3)

            ann = ax.annotate(f"Wells: {', '.join(r['wells'])}",
                             xy=(0.5, -0.26), xycoords='axes fraction',
                             ha='center', fontsize=7.5, color='#888888')
            self.result_annotations.append(ann)

        for idx in range(n, nrows * ncols):
            ax = fig.add_subplot(nrows, ncols, idx + 1)
            ax.set_visible(False)

        fig.tight_layout(rect=[0, 0.02, 1, 1], h_pad=4.0, w_pad=2.5)

        canvas = FigureCanvasTkAgg(fig, master=self.res_inner)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True, padx=4, pady=4)

        # Toolbar row with ⚙ Configure button.
        # pack_toolbar=False is required here because toolbar_row lives inside a
        # scrollable tk.Canvas window: the toolbar's default fill=X/side=BOTTOM
        # auto-pack consumes the full width before the Configure button can be placed.
        toolbar_row = ttk.Frame(self.res_inner)
        toolbar_row.pack(fill='x', padx=4)
        try:
            toolbar = NavigationToolbar2Tk(canvas, toolbar_row, pack_toolbar=False)
        except TypeError:
            # Fallback for matplotlib < 3.3
            toolbar = NavigationToolbar2Tk(canvas, toolbar_row)
        toolbar.pack(side='left', fill='x', expand=True)
        tk.Button(
            toolbar_row, text='⚙ Configure',
            command=lambda: ChartConfigPopup(
                self, fig, self.result_axes,
                annotations=self.result_annotations,
                title='Growth Curves Chart Options'),
            relief='flat', bg='#EEEEEE', font=('Helvetica', 9),
            padx=6, pady=2, cursor='hand2'
        ).pack(side='left', padx=(6, 0))
        toolbar.update()

    # ── Export ────────────────────────────────────────────────────────────────
    def export_plot(self):
        if not self.results_fig:
            messagebox.showinfo('No Plot', 'Process data first to generate plots.')
            return
        path = filedialog.asksaveasfilename(
            title='Save All Plots', defaultextension='.png',
            filetypes=[('PNG', '*.png'), ('PDF', '*.pdf'),
                       ('SVG', '*.svg'), ('All files', '*.*')])
        if path:
            self.results_fig.savefig(path, dpi=200, bbox_inches='tight', facecolor='white')
            messagebox.showinfo('Saved', f'Plot saved to:\n{path}')

    def export_individual(self):
        if not self.results_data:
            messagebox.showinfo('No Data', 'Process data first.')
            return
        ExportSingleDialog(self, self.results_data, self.times, self.log_scale)

    def export_gt_plot(self):
        if not self.gt_fig:
            messagebox.showinfo('No Plot', 'Process data first.')
            return
        path = filedialog.asksaveasfilename(
            title='Save ln(OD) Plot', defaultextension='.png',
            filetypes=[('PNG', '*.png'), ('PDF', '*.pdf'),
                       ('SVG', '*.svg'), ('All files', '*.*')])
        if path:
            self.gt_fig.savefig(path, dpi=200, bbox_inches='tight', facecolor='white')
            messagebox.showinfo('Saved', f'Plot saved to:\n{path}')

    def save_csv(self):
        if not self.results_data:
            messagebox.showinfo('No Data', 'Process data first.')
            return
        path = filedialog.asksaveasfilename(
            defaultextension='.csv', filetypes=[('CSV', '*.csv')])
        if not path:
            return
        rows = []
        for r in self.results_data.values():
            for ti, t in enumerate(self.times):
                rows.append({
                    'Group': r['name'],
                    'Time_HH:MM:SS': hours_to_hms(float(t)),
                    'Time_hours': round(float(t), 6),
                    'Mean_OD600_corrected': round(float(r['mean'][ti]), 6),
                    'CI95_lower': round(float(r['ci_lo'][ti]), 6),
                    'CI95_upper': round(float(r['ci_hi'][ti]), 6),
                    'N_wells': r['n'],
                    'Wells': '; '.join(r['wells']),
                })
        pd.DataFrame(rows).to_csv(path, index=False)
        messagebox.showinfo('Saved', f'Data saved to:\n{path}')

    def export_gentime_csv(self):
        if not self.gt_results:
            messagebox.showinfo('No Results',
                                'Select a window on the Generation Time plot first.')
            return
        path = filedialog.asksaveasfilename(
            defaultextension='.csv', filetypes=[('CSV', '*.csv')])
        if not path:
            return
        cols = ['Group', 'Wells', 'mu_per_h', 'Doubling_time_h',
                'Doubling_time_min', 'R_squared', 'Window']
        pd.DataFrame(self.gt_results, columns=cols).to_csv(path, index=False)
        messagebox.showinfo('Saved', f'Generation time data saved to:\n{path}')


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    app = GrowthCurveApp()
    app.mainloop()


if __name__ == '__main__':
    main()
