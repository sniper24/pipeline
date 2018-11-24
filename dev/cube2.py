import pandas as pd
import numpy as np

def build_cube(from_df, to_path, group_by_cols, max_cols, sum_cols):
	df = pd.read_csv("./nba.csv")
	print df
	# group_by_cols
	# max(max_cols)
	# sum(sum_cols)

	pass
"""
Age                    int64
Conference            object
Date                  object
Draft Year             int64
Height                object
Player                object
Position              object
Season                object
Season short           int64
Seasons in league      int64
Team                  object
Weight                object
Real_value           float64
"""
if __name__ == '__main__':
	df = pd.read_csv("./nba.csv")

	dfg = df.groupby(['Conference', 'Team']).agg({'Age':
                                  {'Mean': np.mean, 'Sum': np.sum}})

	dfg = dfg.reset_index()# .columns

	dfg.columns = [' '.join(col).strip() for col in dfg.columns.values]
	dfg.to_csv('week_grouped.csv', index=False)

	#print df.dtypes