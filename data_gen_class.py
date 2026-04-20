import pandas as pd
import numpy as np
from itertools import product
import openpyxl
from dataclasses import dataclass


# ----------------------------------------------------------------------------------------------------
# Config layer
# ----------------------------------------------------------------------------------------------------
@dataclass
class DatasetConfig:
    sub_groups: list
    measure_names: list
    periods: pd.DatetimeIndex
    value_ranges: dict
    float_measures: set
    cutoff_date: str | pd.Timestamp = "2026-03-01"
    random_seed: int | None = None


    def __post_init__(self):
        if self.cutoff_date is None:
            raise ValueError("cutoff_date cannot be None")
        try:
            self.cutoff_date = pd.Timestamp(self.cutoff_date)
        except Exception as e:
            raise ValueError(f"Invalid cutoff_date: {self.cutoff_date}.") from e

        if self.random_seed is not None:
            self.rng = np.random.default_rng(self.random_seed)
        else:
            self.rng = np.random.default_rng()


# ----------------------------------------------------------------------------------------------------
# Pipeline
# ----------------------------------------------------------------------------------------------------
class DatasetPipeline:
    def __init__(self, config: DatasetConfig):
        self.config = config

    def run(self) -> pd.DataFrame:
        df = self._build_base()
        df = self._generate_values(df)
        df = self._transform(df)
        
        return df


    # Core dataframe
    def _build_base(self) -> pd.DataFrame:
        df = pd.DataFrame(product(
                self.config.sub_groups,
                self.config.measure_names,
                self.config.periods
            ), columns=['Sub_Groups', 'Measure_Names', 'DT_Period']
            )

        df['Value_Type'] = np.where(df['DT_Period'] <= self.config.cutoff_date, 'Historical', 'Forecast')
        df['Reporting_Period'] = df['DT_Period'].dt.strftime("%b '%y")

        return df


    # Randomized values
    def _generate_values(self, df: pd.DataFrame) -> pd.DataFrame:
        bounds = df['Measure_Names'].map(self.config.value_ranges)
        lows = bounds.str[0].to_numpy()
        highs = bounds.str[1].to_numpy()

        values = np.empty(len(df))
        is_float = df['Measure_Names'].isin(self.config.float_measures).to_numpy()
        is_int = ~is_float

        rng = self.config.rng

        values[is_int] = rng.integers(lows[is_int], highs[is_int])
        values[is_float] = rng.uniform(lows[is_float], highs[is_float])

        df['Measure_Values'] = values

        return df


    # Transform dataframe
    def _transform(self, df: pd.DataFrame) -> pd.DataFrame:
        id_cols = ['Sub_Groups', 'Value_Type', 'Reporting_Period', 'DT_Period']

        df_piv = (df.pivot_table(index=id_cols, columns='Measure_Names', values='Measure_Values', sort=False).reset_index())
        df_piv.columns.name = None

        df_piv = self._addn_computation(df_piv)

        df_out = df_piv.melt(id_vars=id_cols, var_name='Measure_Names', value_name='Measure_Values')

        return df_out


    # Additional calculations
    def _addn_computation(self, df: pd.DataFrame) -> pd.DataFrame:
        df['sch_shrnk_hrs'] *= df['core_wlh']
        df['unsch_shrnk_hrs'] *= df['core_wlh']
        df['cst_hrs'] = (df['core_wlh'] - df['sch_shrnk_hrs'] - df['unsch_shrnk_hrs'])
        df['Scheduled Shrinkage (%)'] = (df['sch_shrnk_hrs'] / df['core_wlh'])
        df['Unscheduled Shrinkage (%)'] = (df['unsch_shrnk_hrs'] / df['core_wlh'])

        return df


    # Save to excel template
    def save_to_excel(self, df, save_path):
        with pd.ExcelWriter(save_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name='Data', index=False)
