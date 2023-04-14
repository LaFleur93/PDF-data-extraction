import pandas as pd
from datetime import *
import numpy as np

def missed_and_extra_planting(d):
	#2023-01-21
	day = date(int(d[0:4]), int(d[5:7]), int(d[8:10]))

	date_14 = day + timedelta(days=-14)
	week_14 = date_14.isocalendar().week

	date_21 = day + timedelta(days=-21)
	week_21 = date_21.isocalendar().week

	week_14 = 'W'+str(week_14).zfill(2)
	week_21 = 'W'+str(week_21).zfill(2)

	weeks = [week_14, week_21]

	missing_herbs = {'pr26': 0, 'pr21': 0, 'tm5': 0, 'cr26': 0, 'rs4': 0, 'rs6': 0, 'mn24': 0, 'sg2': 0, 'bs49': 0, 'dl6': 0, 'di6': 0}
	missing_pots = {'bs49': 0, 'cr26': 0, 'tm5': 0, 'mn24': 0, 'rs4': 0, 'rs6': 0, 'che9': 0, 'wtc2': 0, 'mls5': 0, 'or17': 0, 'pr26': 0, 'pea3': 0, 'bs49 - test': 0}

	extra_herbs = {'pr26': 0, 'pr21': 0, 'tm5': 0, 'cr26': 0, 'rs4': 0, 'rs6': 0, 'mn24': 0, 'sg2': 0, 'bs49': 0, 'dl6': 0, 'di6': 0}
	extra_pots = {'bs49': 0, 'cr26': 0, 'tm5': 0, 'mn24': 0, 'rs4': 0, 'rs6': 0, 'che9': 0, 'wtc2': 0, 'mls5': 0, 'or17': 0, 'pr26': 0, 'pea3': 0, 'bs49 - test': 0}

	for week in weeks:

		missed_planting_id, sheet_name = '1azyJv6WsCrQ5mQ2Z0V9R_JQOWz8wtzWi-cro2nlcYuw', week

		url = 'https://docs.google.com/spreadsheets/d/{0}/gviz/tq?tqx=out:csv&sheet={1}'.format(missed_planting_id, sheet_name)

		# Dataframe Missed Planting
		df_missed_planting = pd.read_csv(url)
		df_missed_planting = pd.DataFrame(df_missed_planting)
		df_missed_planting = df_missed_planting.to_numpy()

		search_date = d[8:10]+'.'+d[5:7]+'.'+d[0:4]
		
		if int(week[1:]) >= 4:
			col_extra = 16
		else:
			col_extra = 15

		for i in range(len(df_missed_planting)):
			if df_missed_planting[i][6] == search_date:
				if df_missed_planting[i][4]*30 == df_missed_planting[i][5]:
					if type(df_missed_planting[i][1]) == float:
						continue
					elif df_missed_planting[i][1].lower() == 'di6':
						missing_herbs['dl6'] += int(df_missed_planting[i][4])
					elif df_missed_planting[i][1].lower() == 'mn25':
						missing_herbs['mn24'] += int(df_missed_planting[i][4])
					elif df_missed_planting[i][1].lower() == 'bs105':
						missing_herbs['bs49'] += int(df_missed_planting[i][4])
					elif df_missed_planting[i][1].lower() == 'rs6':
						missing_herbs['rs4'] += int(df_missed_planting[i][4])
					else:
						missing_herbs[df_missed_planting[i][1].lower()] += int(df_missed_planting[i][4])
				else:
					if type(df_missed_planting[i][1]) == float:
						continue
					elif df_missed_planting[i][1].lower() == 'di6':
						missing_pots['dl6'] += int(df_missed_planting[i][4])
					elif df_missed_planting[i][1].lower() == 'mn25':
						missing_pots['mn24'] += int(df_missed_planting[i][4])
					elif df_missed_planting[i][1].lower() == 'bs105':
						missing_pots['bs49'] += int(df_missed_planting[i][4])
					elif df_missed_planting[i][1].lower() == 'rs6':
						missing_pots['rs4'] += int(df_missed_planting[i][4])
					else:
						missing_pots[df_missed_planting[i][1].lower()] += int(df_missed_planting[i][4])

			if df_missed_planting[i][col_extra] == search_date:
				if df_missed_planting[i][col_extra-2]*30 == df_missed_planting[i][col_extra-1]:
					if type(df_missed_planting[i][col_extra-5]) == float:
						continue
					elif df_missed_planting[i][col_extra-5].lower() == 'di6':
						extra_herbs['dl6'] += int(df_missed_planting[i][col_extra-2])
					elif df_missed_planting[i][col_extra-5].lower() == 'mn25':
						extra_herbs['mn24'] += int(df_missed_planting[i][col_extra-2])
					elif df_missed_planting[i][col_extra-5].lower() == 'bs105':
						extra_herbs['bs49'] += int(df_missed_planting[i][col_extra-2])
					elif df_missed_planting[i][col_extra-5].lower() == 'rs6':
						extra_herbs['rs4'] += int(df_missed_planting[i][col_extra-2])
					else:
						extra_herbs[df_missed_planting[i][col_extra-5].lower()] += int(df_missed_planting[i][col_extra-2])
				else:
					if type(df_missed_planting[i][col_extra-5]) == float:
						continue
					elif df_missed_planting[i][col_extra-5].lower() == 'di6':
						extra_pots['dl6'] += int(df_missed_planting[i][col_extra-2])
					elif df_missed_planting[i][col_extra-5].lower() == 'mn25':
						extra_pots['mn24'] += int(df_missed_planting[i][col_extra-2])
					elif df_missed_planting[i][col_extra-5].lower() == 'bs105':
						extra_pots['bs49'] += int(df_missed_planting[i][col_extra-2])
					elif df_missed_planting[i][col_extra-5].lower() == 'rs6':
						extra_pots['rs4'] += int(df_missed_planting[i][col_extra-2])
					else:
						extra_pots[df_missed_planting[i][col_extra-5].lower()] += int(df_missed_planting[i][col_extra-2])
	
	return missing_herbs, missing_pots, extra_herbs, extra_pots