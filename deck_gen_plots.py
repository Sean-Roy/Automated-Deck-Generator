import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter
import seaborn as sns


class PlotsGenerator:
    def __init__(self, save_loc, data, forecast_index, recent_month):
        self.plots_folder = save_loc
        self.df = data
        self.forecast_start_index = forecast_index
        self.max_historical = recent_month
        self.period = self.df['Reporting_Period'].unique()  # eg. Nov - Oct
        self.stack_order = ['FTE Required Total', 'Core', 'Agents of the Future', 'Students', 'Centralized Flex', 'IH IS', 'Agency', 'Net Resource Share', 'Overtime']  # comprehensive plot
        self.pie_vars = ['sch_shrnk_hrs', 'unsch_shrnk_hrs', 'cst_hrs']  # shrink pie plot
        self.pie_labels = ['Scheduled Shrinkage (%)', 'Unscheduled Shrinkage (%)', 'CST (%)']  # shrink pie plot
        self.mthly_shrnk = ['Scheduled Shrinkage (%)', 'Unscheduled Shrinkage (%)']  # shrink bar plot
        self.colour_map = {'FTE Required Total': '#003168', 'Core': '#0051a5', 'Agents of the Future': '#fca311', 'Students': '#899299', 'Centralized Flex': '#ffc72c',
                           'IH IS': '#51b5e0', 'Agency': '#588886', 'Net Resource Share': '#87afbf', 'Overtime': '#6f6e6f',
                           'Scheduled Shrinkage (%)': '#0051a5', 'Unscheduled Shrinkage (%)': '#fca311', 'CST (%)': '#87afbf'}
        sns.set_style('whitegrid')


def plot_all(self, group, target, shrink=True):
    self.plot_comp(group, target)
    if group != 'Foreign' and shrink:
        self.plot_shrk_pie(target)
        self.plot_shrk_bar(target)


def plot_comp(self, group, target):
    fig = plt.figure(figsize=(11, 5.09), dpi=100)
    ax = fig.add_axes([0.08, 0.28, 0.88, 0.62])
    pad_ax = fig.add_axes([0.1, 0.06, 0.1, 0.1])
    pad_ax.axis('off')

    bottom = [0] * len(self.period)
    measures = ['FTE Required Total', 'Core'] if group == 'Foreign' else self.stack_order
    filter_group = 'Groups' if group == 'Foreign' else 'Sub_Groups'
    for measure in measures:
        values = self.df[(self.df[filter_group] == target) & (self.df['Measure_Names'] == measure)]['Measure_Values'].values
        if measure == 'FTE Required Total':
            sns.lineplot(x=self.period, y=values, ax=ax, label=measure, color=self.colour_map[measure], markers='o', linewidth=3)
        elif measure == 'Net Resource Share':
            n_bottom = [0 if y < 0 else x for x, y in zip(bottom, values)]
            sns.barplot(x=self.period, y=values, ax=ax, label=measure, bottom=n_bottom, color=self.colour_map[measure], width=0.3)
            bottom = [x if y < 0 else x + y for x, y in zip(bottom, values)]
        else:
            sns.barplot(x=self.period, y=values, ax=ax, label=measure, bottom=bottom, color=self.colour_map[measure], width=0.3)
            bottom = [x + y for x, y in zip(bottom, values)]

    # add a forecast line
    ax.axvline(x=self.forecast_start_index - 0.5, color='grey', linestyle='--')
    ax.text(self.forecast_start_index - 0.5 + 0.1, 1.01, 'Forecast', fontsize=10, fontweight='bold', color='black', transform=ax.get_xaxis_transform(), ha='left', va='bottom', clip_on=False)

    ax.set(ylabel=None)
    sns.despine(ax=ax, top=True, right=True)
    anchor = (0.5, 0.09) if group == 'Foreign' else (0.5, 0.05)
    ax.legend(fontsize=10, ncol=3, frameon=False, loc='lower center', bbox_to_anchor=(anchor), bbox_transform=fig.transFigure)

    plt.savefig(f'{self.plots_folder}/{target}_comp.png', dpi=100)
    plt.close('all')


def plot_shrk_pie(self, target):
    fig = plt.figure(figsize=(6.1, 3.3), dpi=100)
    ax = fig.add_axes([0.00, 0.00, 1, 1])

    pie_colours = [self.colour_map[x] for x in self.colour_map if x in self.pie_labels]
    pie_values = []
    for measure in self.pie_vars:
        pie_values.append(self.df[(self.df['DT_Period'] == self.max_historical) & (self.df['Sub_Groups'] == target) & (self.df['Measure_Names'] == measure)]['Measure_Values'].values[0])

    ax.pie(x=pie_values, colors=pie_colours, startangle=0)
    ax.set_xlim(-1, 3)
    ax.set_ylim(-1.2, 1.4)
    ax.legend(self.pie_labels, fontsize=10, ncol=1, frameon=False, bbox_to_anchor=(0.55, 0.58), bbox_transform=fig.transFigure)
    ax.set_title(f'{self.max_historical.strftime("%b '%y")}: Productivity Breakdown', x=0.5, y=0.88, fontsize=18, fontweight='bold', color='black')

    fig.savefig(f'{self.plots_folder}/{target}_shrk_pie.png', dpi=100)
    plt.close('all')


def plot_shrk_bar(self, target):
    fig = plt.figure(figsize=(6.1, 3.3), dpi=100)
    ax = fig.add_axes([0.085, 0.26, 0.89, 0.565])

    bottom = [0] * len(self.period)
    for measure in self.mthly_shrnk:
        bar_values = self.df[(self.df['Sub_Groups'] == target) & (self.df['Measure_Names'] == measure)]['Measure_Values'].values
        sns.barplot(x=self.period, y=bar_values, ax=ax, label=measure, bottom=bottom, color=self.colour_map[measure], width=0.35)
        bottom = [x + y for x, y in zip (bottom, bar_values)]
    
    # add a forecast line
    ax.axvline(x=self.forecast_start_index - 0.5, color='grey', linestyle='--')
    ax.text(self.forecast_start_index - 0.5 + 0.1, 1.01, 'Forecast', fontsize=10, fontweight='bold', color='black', transform=ax.get_xaxis_transform(), ha='left', va='bottom', clip_on=False)

    ax.set_ylim(0, 1.1)
    ax.yaxis.set_major_formatter(PercentFormatter(xmax=1, decimals=0))
    ax.set_xticks(self.period)
    ax.set_xticklabels(ax.get_xticklabels(), rotation=45, ha='right')
    sns.despine(ax=ax, top=True, right=True)
    ax.legend(fontsize=10, ncol=2, frameon=False, loc='lower center', bbox_to_anchor=(0.5, -0.01), bbox_transform=fig.transFigure)
    ax.set_title('Scheduled and Unscheduled Shrink', x=0.5, y=1.1, fontsize=18, fontweight='bold', color='black')

    plt.savefig(f'{self.plots_folder}/{target}_shrk_bar.png', dpi=100)
    plt.close('all')
