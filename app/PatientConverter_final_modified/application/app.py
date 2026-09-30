from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False

from converter import build_all_dataset_from_link

SUPPORTED = {'.xls', '.xlsx', '.ods'}


def base_dir() -> Path:
    """Folder containing Link.xlsx, dict/, working files/, and output/."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    here = Path(__file__).resolve().parent
    return here.parent if here.name.lower() == 'application' else here


def open_path(path: Path):
    if sys.platform.startswith('win'):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == 'darwin':
        subprocess.run(['open', str(path)], check=False)
    else:
        subprocess.run(['xdg-open', str(path)], check=False)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title('Patient Workbook Converter')
        self.root.geometry('900x790')
        self.root.minsize(820, 700)
        self.base = base_dir()

        self.link_path = tk.StringVar(value=str(self.base / 'Link.xlsx'))
        self.template_path = tk.StringVar(value=str(self.base / 'working files' / 'dataset_finale_base.xlsx'))
        self.drugs_path = tk.StringVar(value=str(self.base / 'dict' / 'Drugs_dict.xlsx'))
        self.com_path = tk.StringVar(value=str(self.base / 'dict' / 'Com_dict.xlsx'))
        self.interact_path = tk.StringVar(value=str(self.base / 'dict' / 'Interact_dict.xlsx'))
        self.patient = tk.StringVar()
        self.output = tk.StringVar(value=str(self.base / 'output'))
        self.status = tk.StringVar(value='Select the configuration files and a patient workbook.')

        self.queue = queue.Queue()
        self.output_file: Path | None = None
        self.log_file: Path | None = None
        self._build()
        self.root.after(150, self._poll)

    def _build(self):
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill='both', expand=True)

        ttk.Label(outer, text='Patient Workbook Converter', font=('Segoe UI', 17, 'bold')).pack(anchor='w')
        ttk.Label(
            outer,
            text=(
                'The default LINK, template, and dictionary paths are pre-filled. '
                'Each file can also be selected manually before conversion.'
            ),
            wraplength=850,
        ).pack(anchor='w', pady=(4, 14))

        config = ttk.LabelFrame(outer, text='Configuration files', padding=12)
        config.pack(fill='x', pady=(0, 12))
        self._file_row(config, 'LINK workbook', self.link_path, [('Excel workbook', '*.xlsx')])
        self._file_row(config, 'Output template', self.template_path, [('Excel workbook', '*.xlsx')])
        self._file_row(config, 'Drugs dictionary', self.drugs_path, [('Excel workbook', '*.xlsx')])
        self._file_row(config, 'Comorbidity dictionary', self.com_path, [('Excel workbook', '*.xlsx')])
        self._file_row(config, 'Interaction dictionary', self.interact_path, [('Excel workbook', '*.xlsx')])

        patient_frame = ttk.LabelFrame(outer, text='Patient workbook', padding=12)
        patient_frame.pack(fill='x', pady=(0, 12))
        drop = tk.Label(
            patient_frame,
            text='Drop .xls, .xlsx, or .ods file here',
            relief='groove', borderwidth=2, height=4, cursor='hand2',
        )
        drop.pack(fill='x', pady=(0, 8))
        drop.bind('<Button-1>', lambda _e: self._browse_patient())
        if DND_AVAILABLE:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind('<<Drop>>', self._drop)
        row = ttk.Frame(patient_frame)
        row.pack(fill='x')
        ttk.Entry(row, textvariable=self.patient).pack(side='left', fill='x', expand=True)
        ttk.Button(row, text='Browse', command=self._browse_patient).pack(side='left', padx=(8, 0))

        out = ttk.LabelFrame(outer, text='Output folder', padding=12)
        out.pack(fill='x', pady=(0, 12))
        row = ttk.Frame(out)
        row.pack(fill='x')
        ttk.Entry(row, textvariable=self.output).pack(side='left', fill='x', expand=True)
        ttk.Button(row, text='Choose', command=self._browse_output).pack(side='left', padx=(8, 0))

        actions = ttk.Frame(outer)
        actions.pack(fill='x', pady=(0, 12))
        self.run_btn = ttk.Button(actions, text='Run conversion', command=self._start)
        self.run_btn.pack(side='left')
        self.progress = ttk.Progressbar(actions, mode='indeterminate', length=260)
        self.progress.pack(side='left', padx=12)
        self.open_folder_btn = ttk.Button(actions, text='Open output folder', command=self._open_folder)
        self.open_folder_btn.pack(side='right')
        self.open_btn = ttk.Button(actions, text='Open result', command=self._open, state='disabled')
        self.open_btn.pack(side='right', padx=(0, 8))

        status = ttk.LabelFrame(outer, text='Status', padding=12)
        status.pack(fill='both', expand=True)
        ttk.Label(status, textvariable=self.status, wraplength=820).pack(anchor='w', pady=(0, 8))
        self.log = tk.Text(status, height=10, state='disabled', wrap='word')
        self.log.pack(fill='both', expand=True)

    def _file_row(self, parent, label, variable, filetypes):
        row = ttk.Frame(parent)
        row.pack(fill='x', pady=3)
        ttk.Label(row, text=label, width=23).pack(side='left')
        ttk.Entry(row, textvariable=variable).pack(side='left', fill='x', expand=True)
        ttk.Button(
            row,
            text='Browse',
            command=lambda: self._browse_file(variable, filetypes),
        ).pack(side='left', padx=(8, 0))

    def _browse_file(self, variable, filetypes):
        current = Path(variable.get().strip())
        initial = current.parent if current.parent.is_dir() else self.base
        selected = filedialog.askopenfilename(initialdir=initial, filetypes=filetypes + [('All files', '*.*')])
        if selected:
            variable.set(selected)

    def _browse_patient(self):
        selected = filedialog.askopenfilename(
            filetypes=[('Workbooks', '*.xls *.xlsx *.ods'), ('All files', '*.*')]
        )
        if selected:
            self.patient.set(selected)

    def _drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        if paths:
            self.patient.set(paths[0])

    def _browse_output(self):
        initial = self.output.get().strip() or str(self.base)
        selected = filedialog.askdirectory(initialdir=initial)
        if selected:
            self.output.set(selected)

    def _validate(self):
        named_files = {
            'LINK workbook': Path(self.link_path.get().strip()),
            'output template': Path(self.template_path.get().strip()),
            'drugs dictionary': Path(self.drugs_path.get().strip()),
            'comorbidity dictionary': Path(self.com_path.get().strip()),
            'interaction dictionary': Path(self.interact_path.get().strip()),
        }
        missing = [f'{name}: {path}' for name, path in named_files.items() if not path.is_file()]
        if missing:
            messagebox.showerror('Missing configuration', 'Missing or invalid files:\n\n' + '\n'.join(missing))
            return None

        patient = Path(self.patient.get().strip())
        if not patient.is_file() or patient.suffix.lower() not in SUPPORTED:
            messagebox.showerror('Invalid patient file', 'Select a valid .xls, .xlsx, or .ods workbook.')
            return None

        output_text = self.output.get().strip()
        if not output_text:
            messagebox.showerror('Missing output folder', 'Choose an output folder.')
            return None

        return patient, Path(output_text), named_files

    def _start(self):
        valid = self._validate()
        if not valid:
            return
        patient, output, files = valid
        output.mkdir(parents=True, exist_ok=True)
        self.output_file = output / f'{patient.stem}_mapped.xlsx'
        self.log_file = output / f'{patient.stem}_mapped_log.xlsx'

        self._append(f'Patient: {patient}')
        self._append(f'LINK: {files["LINK workbook"]}')
        self._append(f'Template: {files["output template"]}')
        self._append(f'Drugs dictionary: {files["drugs dictionary"]}')
        self._append(f'Comorbidity dictionary: {files["comorbidity dictionary"]}')
        self._append(f'Interaction dictionary: {files["interaction dictionary"]}')
        self._append(f'Output: {self.output_file}')

        self.status.set('Processing...')
        self.run_btn.config(state='disabled')
        self.open_btn.config(state='disabled')
        self.progress.start(10)
        threading.Thread(
            target=self._worker,
            args=(patient, files),
            daemon=True,
        ).start()

    def _worker(self, patient, files):
        try:
            sheets, log = build_all_dataset_from_link(
                link_path=files['LINK workbook'],
                raw_patient_path=patient,
                output_path=self.output_file,
                log_path=self.log_file,
                drugs_dict_path=files['drugs dictionary'],
                com_dict_path=files['comorbidity dictionary'],
                interact_dict_path=files['interaction dictionary'],
                template_path=files['output template'],
            )
            rows_per_sheet = {name: len(frame) for name, frame in sheets.items()}
            statuses = log['status'].value_counts(dropna=False).to_dict() if not log.empty else {}
            self.queue.put(('ok', {'rows_per_sheet': rows_per_sheet, 'statuses': statuses}))
        except Exception as exc:
            self.queue.put(('error', {'message': str(exc), 'traceback': traceback.format_exc()}))

    def _poll(self):
        try:
            event, data = self.queue.get_nowait()
        except queue.Empty:
            self.root.after(150, self._poll)
            return

        self.progress.stop()
        self.run_btn.config(state='normal')
        if event == 'ok':
            self.status.set('Conversion completed successfully.')
            self._append('Output rows per sheet:')
            for name, count in data['rows_per_sheet'].items():
                self._append(f'  {name}: {count}')
            if data['statuses']:
                self._append('Mapping status frequencies:')
                for key, value in data['statuses'].items():
                    self._append(f'  {key}: {value}')
            self.open_btn.config(state='normal')
            messagebox.showinfo(
                'Completed',
                f'Result:\n{self.output_file}\n\nLog:\n{self.log_file}',
            )
        else:
            self.status.set('Conversion failed.')
            self._append(data['message'])
            self._append(data['traceback'])
            messagebox.showerror('Conversion failed', data['message'])
        self.root.after(150, self._poll)

    def _append(self, text):
        self.log.config(state='normal')
        self.log.insert('end', text.rstrip() + '\n')
        self.log.see('end')
        self.log.config(state='disabled')

    def _open(self):
        if self.output_file and self.output_file.exists():
            open_path(self.output_file)

    def _open_folder(self):
        path = Path(self.output.get().strip())
        path.mkdir(parents=True, exist_ok=True)
        open_path(path)


def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
