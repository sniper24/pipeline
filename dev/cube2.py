import pandas as pd
import numpy as np
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

def build_cube(from_df, to_path, group_by_cols, max_cols, sum_cols, avg_cols):
	df = from_df
	
	agg_dict = {}
	for col in max_cols:
		if col not in agg_dict:
			agg_dict[col] = {}
		agg_dict[col]["MAX"] = np.max
	for col in sum_cols:
		if col not in agg_dict:
			agg_dict[col] = {}
		agg_dict[col]["SUM"] = np.sum
	for col in avg_cols:
		if col not in agg_dict:
			agg_dict[col] = {}
		agg_dict[col]["AVG"] = np.mean

	dfg = df.groupby(group_by_cols).agg(agg_dict)
	dfg = dfg.reset_index()

	dfg.columns = ['_'.join(col).strip("_") for col in dfg.columns.values]
	dfg.to_csv(to_path, index=False)
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

def test():
	df = pd.read_csv("./nba.csv")

	dfg = df.groupby(['Conference', 'Team']).agg({'Age':
                                  {'Mean': np.mean, 'Sum': np.sum}})

	dfg = dfg.reset_index()# .columns

	dfg.columns = [' '.join(col).strip() for col in dfg.columns.values]
	dfg.to_csv('week_grouped.csv', index=False)

def main():
	df = pd.read_csv("./nba.csv")
	build_cube(df, 'week_grouped.csv', ['Conference', 'Team'], ['Age'], ['Real_value'],[])

if __name__ == '__main__':
	main()