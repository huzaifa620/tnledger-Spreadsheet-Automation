"""VA Hub: one window to launch every scraping engine in this repo.

Engines are listed in engines.json. Each one runs as its own process, so the
VA can have several going at once. Output from every run is saved under
logs/ and shown live in the hub.
"""
import json, os, subprocess, sys, webbrowser
from datetime import datetime
import PySimpleGUI as sg

ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(ROOT, 'engines.json')
LOG_DIR = os.path.join(ROOT, 'logs')
KEYS_FILE = os.path.join(ROOT, 'keys.json')


def load_engines():
    with open(CONFIG_FILE) as f:
        return json.load(f)['engines']


def build_command(engine):
    # Packaged .exe builds only run on Windows; everywhere else use the script.
    exe = engine.get('exe')
    if exe and os.name == 'nt' and os.path.isfile(os.path.join(ROOT, exe)):
        return [os.path.join(ROOT, exe)]
    return [sys.executable, '-u', os.path.join(ROOT, engine['script'])]


class Run:
    def __init__(self, engine):
        os.makedirs(LOG_DIR, exist_ok=True)
        stamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        slug = os.path.basename(engine['cwd'].rstrip('/'))
        self.log_path = os.path.join(LOG_DIR, f'{slug}_{stamp}.log')
        self.started = datetime.now()
        env = dict(os.environ, VA_KEYS_FILE=KEYS_FILE, PYTHONIOENCODING='utf-8')
        self.log_file = open(self.log_path, 'w', encoding='utf-8')
        self.proc = subprocess.Popen(build_command(engine), cwd=os.path.join(ROOT, engine['cwd']),
                                     stdout=self.log_file, stderr=subprocess.STDOUT, env=env)

    def status(self):
        code = self.proc.poll()
        if code is None:
            return 'Running'
        self.log_file.close()
        return 'Done' if code == 0 else f'Failed ({code})'

    def stop(self):
        if self.proc.poll() is None:
            self.proc.terminate()

    def tail(self, lines=200):
        try:
            with open(self.log_path, encoding='utf-8', errors='replace') as f:
                return ''.join(f.readlines()[-lines:])
        except OSError:
            return ''


def table_rows(engines, runs):
    rows = []
    for i, engine in enumerate(engines):
        run = runs.get(i)
        status = run.status() if run else 'Idle'
        started = run.started.strftime('%m/%d %H:%M') if run else '-'
        rows.append([engine['name'], engine['category'], status, started])
    return rows


def main():
    engines = load_engines()
    runs = {}
    sg.theme('DarkBlue3')
    layout = [
        [sg.Text('Pick one or more engines (Ctrl/Shift-click), then press Run.', font=('Ubuntu', 11))],
        [sg.Table(values=table_rows(engines, runs), headings=['Engine', 'Type', 'Status', 'Started'],
                  key='-TABLE-', select_mode=sg.TABLE_SELECT_MODE_EXTENDED, enable_events=True,
                  auto_size_columns=False, col_widths=[38, 18, 12, 12], num_rows=len(engines),
                  justification='left', font=('Ubuntu', 11))],
        [sg.Text('', key='-DESC-', size=(90, 2), font=('Ubuntu', 10))],
        [sg.Button('Run', font=('Ubuntu', 12)), sg.Button('Stop', font=('Ubuntu', 12)),
         sg.Button('Open Sheet', font=('Ubuntu', 12)), sg.Button('Open Logs Folder', font=('Ubuntu', 12)),
         sg.Button('Exit', font=('Ubuntu', 12))],
        [sg.Multiline('', key='-LOG-', size=(100, 18), autoscroll=True, disabled=True, font=('Courier', 9))],
    ]
    win = sg.Window('VA Hub - Scraping Engines', layout, finalize=True)
    selected = []

    while True:
        event, values = win.read(timeout=1000)
        if event in (sg.WIN_CLOSED, 'Exit'):
            if any(r.proc.poll() is None for r in runs.values()):
                if sg.popup_yes_no('Engines are still running. Stop them and exit?') != 'Yes':
                    continue
                for r in runs.values():
                    r.stop()
            break

        if event == '-TABLE-':
            selected = values['-TABLE-']
            if selected:
                win['-DESC-'].update(engines[selected[0]]['description'])

        elif event == 'Run':
            if not selected:
                sg.popup('Select an engine first.')
            for i in selected:
                if i in runs and runs[i].proc.poll() is None:
                    sg.popup(f"{engines[i]['name']} is already running.")
                    continue
                try:
                    runs[i] = Run(engines[i])
                except OSError as err:
                    sg.popup_error(f"Could not start {engines[i]['name']}:\n{err}")

        elif event == 'Stop':
            for i in selected:
                if i in runs:
                    runs[i].stop()

        elif event == 'Open Sheet':
            for i in selected:
                webbrowser.open(f"https://docs.google.com/spreadsheets/d/{engines[i]['sheet_id']}")

        elif event == 'Open Logs Folder':
            os.makedirs(LOG_DIR, exist_ok=True)
            webbrowser.open(f'file://{LOG_DIR}')

        # Refresh statuses without losing the VA's row selection.
        win['-TABLE-'].update(values=table_rows(engines, runs), select_rows=selected)
        if selected and selected[0] in runs:
            win['-LOG-'].update(runs[selected[0]].tail())
        elif event == '-TABLE-':
            win['-LOG-'].update('')

    win.close()


if __name__ == '__main__':
    main()
