import pandas as pd


class DataCleaner():
    def __init__(self):
        self.key_columns = ['Core_Groups', 'Groups', 'Sub_Groups', 'Value_Type', 'Reporting_Period', 'DT_Period']
        self.floaters = ['Scheduled Shrinkage (%)', 'Unscheduled Shrinkage (%)', 'Total Shrinkage (%)', 'CST (%)',
                         'core_wlh', 'sch_shrnk_hrs', 'unsch_shrnk_hrs', 'cst_hrs']


    def process_data(self, region, data):
        self.prepare_df(data)
        self.calculate_values()
        return self.finalize_df(region)
    

    def prepare_df(self, data):
        df = data.copy()
        df.columns = df.columns.str.strip()
        obj_cols = df.select_dtypes('str').columns
        df.loc[:, obj_cols] = df.loc[:, obj_cols].apply(lambda x: x.str.strip())

        df['DT_Period'] = pd.to_datetime(df['Reporting_Period'], format="%b %y")
        date_range = df.sort_values(by='DT_Period', ascending=True)['DT_Period'].unique()[-12:]
        df = df[df['DT_Period'].isin(date_range)]

        self.df_pivot = df.pivot_table(index=self.key_columns, columns='Measure_Names', values='Measure_Values', sort=False).reset_index()
        self.df_pivot.columns.name = None


    def calculate_values(self):
        self.df_pivot['core_wlh'] = self.df_pivot['Total Workload (hrs)'] / self.df_pivot['Workload FTE'] * self.df_pivot['Core']
        self.df_pivot['sch_shrnk_hrs'] = self.df_pivot['core_wlh'] * self.df_pivot['Scheduled Shrinkage (%)']
        self.df_pivot['unsch_shrnk_hrs'] = self.df_pivot['core_wlh'] * self.df_pivot['Unscheduled Shrinkage (%)']
        self.df_pivot['cst_hrs'] = self.df_pivot['core_wlh'] - (self.df_pivot['sch_shrnk_hrs'] + self.df_pivot['unsch_shrnk_hrs'])

        self.df_pivot['Net Resource Share'] = self.df_pivot['Resource Share In'] + (self.df_pivot['Resource Share Out'] * (-1))

        self.df_pivot['Sub_Groups'] = self.df_pivot['Sub_Groups'].apply(lambda x: x.removesuffix(' (Bilingual)') if 'Bilingual' in x else x)
        self.df_pivot = self.df_pivot.groupby(self.key_columns, as_index=False).sum(numeric_only=True)

        self.df_pivot['Scheduled Shrinkage (%)'] = self.df_pivot['sch_shrnk_hrs'] / self.df_pivot['core_wlh']
        self.df_pivot['Unscheduled Shrinkage (%)'] = self.df_pivot['unsch_shrnk_hrs'] / self.df_pivot['core_wlh']
        self.df_pivot['Total Shrinkage (%)'] = self.df_pivot['Scheduled Shrinkage (%)'] + self.df_pivot['Unscheduled Shrinkage (%)']
        self.df_pivot['CST (%)'] - self.df_pivot['cst_hrs'] / self.df_pivot['core_wlh']


    def finalize_df(self, region):
        df_unpiv = self.df_pivot.melt(id_vars=self.key_columns, var_name='Measure_Names', value_name='Measure_Values')

        if region == 'Foreign':
            new_df = df_unpiv.groupby(['Groups', 'Measure_Names', 'Value_Type', 'Reporting_Period', 'DT_Period'], as_index=False)['Measure_Values'].sum()
            group_list = sorted(new_df['Groups'].unique().tolist())
        else:
            new_df = df_unpiv.groupby(['Sub_Groups', 'Measure_Names', 'Value_Type', 'Reporting_Period', 'DT_Period'], as_index=False)['Measure_Values'].sum()
            group_list = sorted(new_df['Sub_Groups'].unique().tolist())

        new_df.sort_values(by='DT_Period', ascending=True, inplace=True)
        new_df.reset_index(drop=True, inplace=True)

        non_floaters = ~new_df['Measure_Names'].isin(self.floaters)
        new_df.loc[non_floaters, 'Measure_Values'] = new_df.loc[non_floaters, 'Measure_Values'].round().astype(int)

        if region == 'Local':
            self.set_vars(new_df)

        return group_list, new_df
    

    def set_vars(self, data):
        self.min_historical = data[data['Value_Type'] == 'Historical']['DT_Period'].unique().min()
        self.max_historical = data[data['Value_Type'] == 'Historical']['DT_Period'].unique().max()
        self.min_forecast = data[data['Value_Type'] == 'Forecast']['DT_Period'].unique().min()
        self.min_forecast = data[data['Value_Type'] == 'Forecast']['DT_Period'].unique().max()
        forecast_start = data[data['DT_Period'] == self.min_forecast]['Reporting_Period'].unique()
        self.forecast_start_index = data['Reporting_Period'].unique().tolist().index(forecast_start)
