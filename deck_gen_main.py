import tkinter as tk
from tkinter import messagebox
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv  # python-dotenv
import os
import sys
import time
import win32com.client as win32  # pywin32
from PIL import ImageGrab  # Pillow
from io import BytesIO
from deck_gen_data import DataCleaner
from deck_gen_plots import PlotsGenerator
from deck_gen_ppt import PPTGenerator

start = time.time()

# ----------------------------------------------------------------------------------------------------
# Folder/file paths
# ----------------------------------------------------------------------------------------------------
load_dotenv()

# data used for plotting graphs
all_plot_data = 'Data_1'
os_plot_data = 'Data_2'

# data file for plot & ppt generation
ppt_data = os.getenv('DATA_LOC')

# path for ppt template
template_path = os.getenv('TEMPLATES_LOC')

# path for generated plots
plot_path = os.getenv('PLOTS_LOC')

# path for generated ppts
ppt_path = os.getenv('PPTS_LOC')

# ----------------------------------------------------------------------------------------------------
# Data processing
# ----------------------------------------------------------------------------------------------------
try:
    df_full = pd.read_excel(ppt_data, sheet_name=all_plot_data)
    df_os = pd.read_excel(ppt_data, sheet_name=os_plot_data)
except FileNotFoundError:
    messagebox.showinfo('Missing File!', f'"{ppt_data}" was not found.')
    sys.exit()
except PermissionError:
    messagebox.showinfo('Locked File!', f'"{ppt_data}" is open or locked by another process.')
    sys.exit()
except ValueError as e:
    messagebox.showinfo('File Error!', f'"{ppt_data}" has an error in the file: {e}.')
    sys.exit()

plot_dfs = {'Local': df_full,
            'Foreign': df_os}

cleaner = DataCleaner()
for region, df in plot_dfs.items():
    try:
        if region == 'Foreign':
            foreign_groups, plot_dfs[region] = cleaner.process_data(region, df)
        else:
            groups, plot_dfs[region] = cleaner.process_data(region, df)
    except Exception as e:
        messagebox.showinfo('Invalid data!', f'Ensure "{region}" data is extracted correctly, with the correct fields.\n\n'
                            f'Error: {e}')
        sys.exit()

# ----------------------------------------------------------------------------------------------------
# PPT variables
# ----------------------------------------------------------------------------------------------------
def ordinal(n):  # date suffixes
    if 10 <= n % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return (str(n), suffix)


now = datetime.now()
report_year = now.year
report_month = now.strftime('%b')
report_title = f'Monthly Workforce Report - {now.strftime('%B')}{report_year}'
report_date = ordinal(now.day)

core_groups = {'Local_1': ['Group_1', 'Group_2'],
               'Local_2': 'Group_3',
               'Foreign': foreign_groups}
unique_groups = ['Group_4', 'Group_5']  # unique ppt templates
no_cst_groups = ['Group_2', 'Group_5']  # unique groups without 4th slide
set_2_groups =['Group_3']  # unique group with different tables
set_3_groups = ['Group_4']  # unique group with different tables
missing_groups = []  # exceptions/errors


def set_fy_title():
    if pd.isnull(cleaner.min_historical) or pd.isnull(cleaner.max_historical):
        return f'FY {cleaner.max_forecast.year} Forecast ({cleaner.min_forecast.strftime('%b')} - {cleaner.max_forecast.strftime('%b')})'
    elif pd.isnull(cleaner.min_forecast) or pd.isnull(cleaner.max_forecast):
        return f'FY {cleaner.max_historical.year} Historicals ({cleaner.min_historical.strftime('%b')} - {cleaner.max_historical.strftime('%b')})'
    else:
        return f'FY {cleaner.max_forecast.year} Historicals ({cleaner.min_historical.strftime('%b')} - {cleaner.max_historical.strftime('%b')}) + Forecast ({cleaner.min_forecast.strftime('%b')} - {cleaner.max_forecast.strftime('%b')})'


fy_title = set_fy_title()

xl_update = {'Local': ['Local_Summary', 'A1'],  # sheet for all Local Groups
             'Foreign': ['Foreign_Summary', 'A1']}  # sheet for Foreign Groups
xl_data = {'SET_1_V1': 'A1:M22', 'SET_1_V2': 'A1:M5', 'SET_1_Commentary': 'A1:G7', 'SET_1_V3': 'A1:Q5'}  # General tables
set_2_xl_data = {'SET_2_V1': 'A1:M22', 'SET_2_V2': 'A1:M5', 'SET_2_Commentary': 'A1:G7', 'SET_2_V3': 'A1:Q5'}  # FTM tables
set_3_xl_data = {'SET_3_V1': 'A1:M22', 'SET_3_V2': 'A1:M5', 'SET_3_Commentary': 'A1:G7', 'SET_3_V3': 'A1:Q5'}  # CM&R tables
set_4_xl_data = {'SET_4_V1': 'A1:M22', 'SET_4_V2': 'A1:M5', 'SET_4_Commentary_V1': 'A1:G7', 'SET_4_Commentary_V2': 'A1:G4'}  # Foreign tables
xl_images = []  # table images

# ----------------------------------------------------------------------------------------------------
# Generate all Plots and PPTs
# ----------------------------------------------------------------------------------------------------
try:
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(ppt_data)
except:
    if FileNotFoundError:
        messagebox.showinfo('Missing File!', f'"{ppt_data}" was not found.')
        sys.exit()
    if PermissionError:
        messagebox.showinfo('Locked File!', f'"{ppt_data}" is open or locked by another process.')
        sys.exit()


def update_group(group, target):
    if group == 'Foreign':
        page, cell = xl_update[group][0], xl_update[group][1]
    else:
        page, cell = xl_update['Local'][0], xl_update['Local'][1]
    ws = wb.Sheets(page)
    ws.Range(cell).Value = target


def extract_tables(tables):
    for sheet in tables:
        ws = wb.Sheets(sheet)
        rng = ws.Range(tables[sheet])
        rng.CopyPicture(Appearance=1, Format=2)
        time.sleep(0.8)  # avoid Excel copy error
        img = ImageGrab.grabclipboard()
        if img:
            buffer = BytesIO()
            img.save(buffer, format='PNG')
            buffer.seek(0)
            xl_images.append(buffer)


gen_full_plot = PlotsGenerator(plot_path, plot_dfs['Local'], cleaner.forecast_start_index, cleaner.max_historical)
gen_os_plot = PlotsGenerator(plot_path, plot_dfs['Foreign'], cleaner.forecast_start_index, cleaner.max_historical)

for group in groups:
    targets = foreign_groups if group == 'Foreign' else [group]
    for target in targets:
        gen_plot = gen_os_plot if group == 'Foreign' else gen_full_plot
        try:
            gen_plot.plot_all(group, target, shrink=False if group in no_cst_groups else True)
        except Exception as e:
            missing_groups.append(f'Plot generation error ({target}): {e}')
            continue

        group = None
        for key, values in core_groups.items():
            if target in values:
                group = key
        if group == None:
            missing_groups.append(f'Missing group ({target})')
            continue

        template = None
        if target in unique_groups:
            template = target
        else:
            template = group
        full_template_path = f'{template_path}{template}.pptx'

        update_group(group, target)

        ppt_title = target
        if target in set_3_groups:
            extract_tables(set_3_xl_data)
            ppt_title = 'Local Set 3'
        elif target in set_2_groups:
            extract_tables(set_2_xl_data)
        elif target in foreign_groups:
            extract_tables(set_4_xl_data)
        else:
            extract_tables(xl_data)
        page2_title = f'{ppt_title} - Staffing Levels'
        page3_title = f'{ppt_title} - Capacity Details'
        page4_title = f'{ppt_title} - Shrink Details'

        gen_ppt = PPTGenerator(plot_path, full_template_path)

        # Page 1
        gen_ppt.page1_group(ppt_title)
        gen_ppt.page1_title(report_title)
        gen_ppt.page1_date(report_month, report_date, report_year)
        # Page 2
        gen_ppt.page2_title(page2_title)
        gen_ppt.page2_plot_title(fy_title)
        gen_ppt.page2_plot(target)
        gen_ppt.page2_table(xl_images[1])
        gen_ppt.page2_commentary_1(group, xl_images[2])
        if group == 'Foreign':
            gen_ppt.page2_commentary_2(xl_images[3])
        # Page 3
        gen_ppt.page3_title(page3_title)
        gen_ppt.page3_table(xl_images[0])
        # Page 4
        if group != 'Foreign' and target not in no_cst_groups:
            gen_ppt.page4_title(page4_title)
            gen_ppt.page4_table(xl_images[3])
            gen_ppt.page4_plot_1(target)
            gen_ppt.page4_plot_2(target)
        elif group != 'Foreign' and target in no_cst_groups:
            gen_ppt.page4_delete()

        full_ppt_path = f'{ppt_path}{report_year}\\{report_month}\\{group}'
        try:
            os.makedirs(full_ppt_path, exist_ok=True)
        except OSError as e:
            messagebox.showinfo('Directory Error!', f'Error creating directory: {e}')
            sys.exit()
        gen_ppt.save_ppt(full_ppt_path, report_year, report_month, ppt_title)
        xl_images = []  #clean images from memory

wb.Close(False)
excel.Quit()

# ----------------------------------------------------------------------------------------------------
# Confirmation
# ----------------------------------------------------------------------------------------------------
window = tk.Tk()
window.withdraw()
window.attributes('-topmost', True)

end = time.time()
total_time = (end - start) / 60
total_sec = (total_time - int(total_time)) * 60

if missing_groups:
    messagebox.showinfo(f'Process completed in {int(total_time)} min {int(total_sec)} sec!', 'No deck(s) generated for the following group(s):\n\n'
                        f'{',\n'.join(missing_groups)}')
else:
    messagebox.showinfo(f'Process completed in {int(total_time)} min {int(total_sec)} sec!', 'Process completed with no exceptions.')

window.attributes('-topmost', False)
window.quit()
