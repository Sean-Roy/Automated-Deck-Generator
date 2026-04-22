from __future__ import annotations
import pandas as pd
from typing import List, Tuple


# ----------------------------------------------------------------------------------------------------
# Initialize
# ----------------------------------------------------------------------------------------------------
class DataCleaner:
    key_columns: List[str] = [
        "Core_Groups",
        "Groups",
        "Sub_Groups",
        "Value_Type",
        "Reporting_Period",
        "DT_Period",
    ]

    float_columns: List[str] = [
        "Scheduled Shrinkage (%)",
        "Unscheduled Shrinkage (%)",
        "Total Shrinkage (%)",
        "CST (%)",
        "core_wlh",
        "sch_shrnk_hrs",
        "unsch_shrnk_hrs",
        "cst_hrs",
    ]


# ----------------------------------------------------------------------------------------------------
# Pipeline
# ----------------------------------------------------------------------------------------------------
    def process_data(self, region: str, data: pd.DataFrame) -> Tuple[List[str], pd.DataFrame]:
        df_prepared = self._prepare_df(data)
        df_transformed = self._calculate_values(df_prepared)
        return self._finalize_df(df_transformed, region)


    # Prepare Data
    def _prepare_df(self, data: pd.DataFrame) -> pd.DataFrame:
        df = data.copy()

        # Clean column names
        df.columns = df.columns.str.strip()
        str_cols = df.select_dtypes(include="object").columns
        df[str_cols] = df[str_cols].apply(lambda col: col.str.strip())

        df["DT_Period"] = pd.to_datetime(df["Reporting_Period"], format="%b %y")

        # Use only recent 12 months
        latest_periods = (df.sort_values("DT_Period")["DT_Period"].drop_duplicates().tail(12))
        df = df[df["DT_Period"].isin(latest_periods)]

        df_pivot = (df.pivot_table(index=self.key_columns,columns="Measure_Names",values="Measure_Values",sort=False,).reset_index())
        df_pivot.columns.name = None
        return df_pivot


    # Perform calculations and adjustments
    def _calculate_values(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()

        # shrink/cst hours
        df["core_wlh"] = (df["Total Workload (hrs)"] / df["Workload FTE"] * df["Core"])
        df["sch_shrnk_hrs"] = df["core_wlh"] * df["Scheduled Shrinkage (%)"]
        df["unsch_shrnk_hrs"] = df["core_wlh"] * df["Unscheduled Shrinkage (%)"]
        df["cst_hrs"] = df["core_wlh"] - (df["sch_shrnk_hrs"] + df["unsch_shrnk_hrs"])

        # net resource share
        df["Net Resource Share"] = (df["Resource Share In"] - df["Resource Share Out"])

        # clean subgroup names
        df["Sub_Groups"] = df["Sub_Groups"].str.replace(r"\s*\(Bilingual\)", "", regex=True)

        # aggregate
        df = df.groupby(self.key_columns, as_index=False).sum(numeric_only=True)

        # shrink/cst percentages
        df["Scheduled Shrinkage (%)"] = df["sch_shrnk_hrs"] / df["core_wlh"]
        df["Unscheduled Shrinkage (%)"] = df["unsch_shrnk_hrs"] / df["core_wlh"]
        df["Total Shrinkage (%)"] = (df["Scheduled Shrinkage (%)"] + df["Unscheduled Shrinkage (%)"])
        df["CST (%)"] = df["cst_hrs"] / df["core_wlh"]

        return df


    # Finalize output
    def _finalize_df(self, df: pd.DataFrame, region: str) -> Tuple[List[str], pd.DataFrame]:

        df_unpivot = df.melt(id_vars=self.key_columns,var_name="Measure_Names",value_name="Measure_Values",)

        # region-based grouping
        if region == "Foreign":
            group_col = "Groups"
        else:
            group_col = "Sub_Groups"

        grouped = (df_unpivot.groupby([group_col, "Measure_Names", "Value_Type", "Reporting_Period", "DT_Period"],as_index=False,)["Measure_Values"].sum())
        group_list = sorted(grouped[group_col].unique())
        grouped = grouped.sort_values("DT_Period").reset_index(drop=True)

        # format numeric values
        mask = ~grouped["Measure_Names"].isin(self.float_columns)
        grouped.loc[mask, "Measure_Values"] = (grouped.loc[mask, "Measure_Values"].round().astype("Int64"))

        # optional metadata
        if region == "Local":
            self._set_metadata(grouped)

        return group_list, grouped


    # metadata
    def _set_metadata(self, df: pd.DataFrame) -> None:
        historical = df[df["Value_Type"] == "Historical"]
        forecast = df[df["Value_Type"] == "Forecast"]

        self.min_historical = historical["DT_Period"].min()
        self.max_historical = historical["DT_Period"].max()

        self.min_forecast = forecast["DT_Period"].min()
        self.max_forecast = forecast["DT_Period"].max()

        forecast_start_period = (df.loc[df["DT_Period"] == self.min_forecast, "Reporting_Period"].iloc[0])
        self.forecast_start_index = (df["Reporting_Period"].drop_duplicates().tolist().index(forecast_start_period))
