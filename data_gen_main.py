import pandas as pd
from dotenv import load_dotenv
import os
from data_gen_class import DatasetConfig
from data_gen_class import DatasetPipeline


# ----------------------------------------------------------------------------------------------------
# Set config
# ----------------------------------------------------------------------------------------------------
config = DatasetConfig(
    sub_groups = ['Group_1', 'Group_2', 'Group_3', 'Group_4'],
    measure_names = [
        'FTE Required Total', 'Core', 'Agents of the Future', 'Students',
        'Centralized Flex', 'IH IS', 'Agency', 'Net Resource Share',
        'Overtime', 'core_wlh', 'sch_shrnk_hrs', 'unsch_shrnk_hrs'],
    periods = pd.date_range("2026-01-01", freq="MS", periods=12),
    value_ranges = {
        'FTE Required Total': (20, 90),
        'Core': (10, 60),
        'Agents of the Future': (0, 10),
        'Students': (0, 8),
        'Centralized Flex': (0, 5),
        'IH IS': (0, 5),
        'Agency': (0, 5),
        'Net Resource Share': (-10, 15),
        'Overtime': (0, 10),
        'core_wlh': (5000, 17000),
        'sch_shrnk_hrs': (0.2, 0.5),
        'unsch_shrnk_hrs': (0.1, 0.3)},
    float_measures = {'sch_shrnk_hrs', 'unsch_shrnk_hrs'},
    cutoff_date = "2026-05-01",
    random_seed = 125)  # reproducibility

# ----------------------------------------------------------------------------------------------------
# Run the process
# ----------------------------------------------------------------------------------------------------
load_dotenv()
excel_template = os.getenv('EXCEL_LOC')

pipeline = DatasetPipeline(config)
df = pipeline.run()
pipeline.save_to_excel(df, excel_template)
