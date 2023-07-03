# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a python script file.
"""

import pulp
import pandas as pd
import os
import datetime

current_dir = os.getcwd()

#Get data from "Second Round Technical Question - Attachment 1"
attachment1 = 'Second Round Technical Question - Attachment 1.xlsx'
file_path = os.path.join(current_dir, attachment1)
parameters = pd.read_excel(file_path)

# Define the parameters
max_charging_rate = parameters.iloc[0, 1] # MW
max_discharging_rate = parameters.iloc[1, 1] # MW
max_storage_volume = parameters.iloc[2, 1] # MWh
battery_charging_efficiency = parameters.iloc[3, 1]
battery_discharging_efficiency = parameters.iloc[4, 1]

#Get data from "Second Round Technical Question - Attachment 2"
attachment2 = 'Second Round Technical Question - Attachment 2.xlsx'
file_path = os.path.join(current_dir, attachment2)
# Store all sheets in a dictionary
sheets_dict = pd.read_excel(file_path, sheet_name=None)
half_hourly_data = sheets_dict['Half-hourly data']
daily_data = sheets_dict['Daily data']

# Define the time intervals (half-hourly and daily)
half_hour_intervals = len(half_hourly_data)
daily_intervals = len(daily_data)

# Define the market prices for each interval
prices1 = half_hourly_data.iloc[:, 1]  # Market 1 prices (£/MWh)
prices2 = half_hourly_data.iloc[:, 2]  # Market 2 prices (£/MWh)
prices3 = daily_data.iloc[:, 1]  # Market 3 prices (£/MWh)



# Create the LP problem
problem_formulation = pulp.LpProblem("battery_energy_trading", pulp.LpMaximize)

# Define the decision variables
# The power (MW) constantly exported from battery to Market 1 during half-hour period t
P1_out = pulp.LpVariable.dicts("Power_Export_to_Market1", range(half_hour_intervals), lowBound=0, upBound=max_discharging_rate, cat="Continuous")
# The power (MW) constantly exported from battery to Market 2 during half-hour period t
P2_out = pulp.LpVariable.dicts("Power_Export_to_Market2", range(half_hour_intervals), lowBound=0, upBound=max_discharging_rate, cat="Continuous")
# The power (MW) constantly exported from battery to Market 3 on day d
P3_out = pulp.LpVariable.dicts("Power_Export_to_Market3", range(daily_intervals), lowBound=0, upBound=max_discharging_rate, cat="Continuous")
# The power (MW) constantly imported to the battery from Market 1 during half-hour period t
P1_in = pulp.LpVariable.dicts("Power_Import_from_Market1", range(half_hour_intervals), lowBound=0, upBound=max_charging_rate, cat="Continuous")
# The power (MW) constantly imported to the battery from Market 2 during half-hour period t
P2_in = pulp.LpVariable.dicts("Power_Import_from_Market2", range(half_hour_intervals), lowBound=0, upBound=max_charging_rate, cat="Continuous")
# The power (MW) constantly imported to the battery on day d
P3_in = pulp.LpVariable.dicts("Power_Import_from_Market3", range(daily_intervals), lowBound=0, upBound=max_charging_rate, cat="Continuous")
# The energy level of the battery at the beginning of half-hour period t (MWh)
E = pulp.LpVariable.dicts("Energy_Stored", range(half_hour_intervals+1), lowBound=0, upBound=max_storage_volume, cat="Continuous")

# Define the objective function: maximize profit
objective_function = 0
for t in range(half_hour_intervals):
    objective_function += prices1[t] * 0.5 * (P1_out[t]*(1-battery_discharging_efficiency)-P1_in[t]) + prices2[t] * 0.5 * (P2_out[t]*(1-battery_discharging_efficiency)-P2_in[t])
    
for d in range(daily_intervals):
    objective_function += prices3[d] * 24  * (P3_out[d]*(1-battery_discharging_efficiency)-P3_in[d])

problem_formulation += objective_function

# Define the constraints
problem_formulation += E[0] == 0 # Assuming that the battery started with no energy stored
for t in range(half_hour_intervals):
    # Battery energy level constraints
    problem_formulation += E[t+1] == E[t] + 0.5 * ((P1_in[t] + P2_in[t] + P3_in[t//48])*(1-battery_charging_efficiency) - (P1_out[t] + P2_out[t] + P3_out[t//48]))
    problem_formulation += E[t+1] <= max_storage_volume
    #problem_formulation += 0.5 * (P1_out[t] + P2_out[t] + P3_out[t//48]) - 0.5 * (P1_in[t] + P2_in[t] + P3_in[t//48])*(1-battery_charging_efficiency) <= E[t] 
    # Battery charging and discharging constraints
    problem_formulation += P1_in[t] + P2_in[t] + P3_in[t//48] <= max_charging_rate 
    problem_formulation += P1_out[t] + P2_out[t] + P3_out[t//48] <= max_discharging_rate 

# Solve the optimization problem
problem_formulation.solve()

# Print the optimal solution and profit
optimal_profit = pulp.value(problem_formulation.objective)
print("Optimization Status:", pulp.LpStatus[problem_formulation.status])
print("Optimal Profit: £", optimal_profit)

# Save the optimal decisions
optimal_P1_out = pd.DataFrame(pulp.value(P1_out[t]) for t in range(half_hour_intervals))
optimal_P2_out = pd.DataFrame(pulp.value(P2_out[t]) for t in range(half_hour_intervals))
optimal_P3_out_day = [pulp.value(P3_out[d]) for d in range(daily_intervals)]
optimal_P1_in = pd.DataFrame(pulp.value(P1_in[t]) for t in range(half_hour_intervals))
optimal_P2_in = pd.DataFrame(pulp.value(P2_in[t]) for t in range(half_hour_intervals))
optimal_P3_in_day = [pulp.value(P3_in[d]) for d in range(daily_intervals)]

#Organize optimal decisions and optimal profits
optimal_P3_out = []
optimal_P3_in = []
for t in range(half_hour_intervals):
    optimal_P3_out = optimal_P3_out + [optimal_P3_out_day[t//48]]
    optimal_P3_in = optimal_P3_in + [optimal_P3_in_day[t//48]]
    
#Convert Optimal decisions to discharge/charge Market 3 to half-hourly for excel output
optimal_P3_out_df = pd.DataFrame({'Discharge to Market 3 (MW)': optimal_P3_out})
optimal_P3_in_df = pd.DataFrame({'Charge from Market 3 (MW)': optimal_P3_in})

optimal_decisions = pd.concat([half_hourly_data.iloc[:, 0], optimal_P1_out,optimal_P2_out,optimal_P3_out_df, optimal_P1_in,optimal_P2_in,optimal_P3_in_df],axis = 1)
optimal_decisions.columns = ['Time','Discharge to Market 1 (MW)', 'Discharge to Market 2 (MW)', 'Discharge to Market 3 (MW)',
                             'Charge from Market 1 (MW)', 'Charge from Market 2 (MW)', 'Charge from Market 3 (MW)']

optimal_decisions.to_excel("optimal_decisions.xlsx",sheet_name = 'Optimal Decisions', index=False)

half_hourly_data.columns = ["datetime"] + list(half_hourly_data.columns[1:])
half_hourly_data['datetime'] = pd.to_datetime(half_hourly_data['datetime'])
num_rows_2018 = len(half_hourly_data[half_hourly_data['datetime'].dt.year == 2018])
num_rows_2019 = len(half_hourly_data[half_hourly_data['datetime'].dt.year == 2019])

#Calculating optimal
optimal_profit_2018 = 0
for t in range(num_rows_2018):
    optimal_profit_2018 += (  prices1[t] * 0.5 * (optimal_P1_out.iloc[t,0]*(1-battery_discharging_efficiency)-optimal_P1_in.iloc[t,0]) 
                            + prices2[t] * 0.5 * (optimal_P2_out.iloc[t,0]*(1-battery_discharging_efficiency)-optimal_P2_in.iloc[t,0]))  
                         
for d in range(num_rows_2018//48):
    optimal_profit_2018 += prices3[d] * 24  * (optimal_P3_out_day[d]*(1-battery_discharging_efficiency)-optimal_P3_in_day[d])
    
    
optimal_profit_2019 = 0
for t in range(num_rows_2018,(num_rows_2018 + num_rows_2019)):
    optimal_profit_2019 += (  prices1[t] * 0.5 * (optimal_P1_out.iloc[t,0]*(1-battery_discharging_efficiency)-optimal_P1_in.iloc[t,0]) 
                            + prices2[t] * 0.5 * (optimal_P2_out.iloc[t,0]*(1-battery_discharging_efficiency)-optimal_P2_in.iloc[t,0]))  
                         
for d in range(num_rows_2018//48,(num_rows_2018 + num_rows_2019)//48):
    optimal_profit_2019 += prices3[d] * 24  * (optimal_P3_out_day[d]*(1-battery_discharging_efficiency)-optimal_P3_in_day[d])
    
optimal_profit_2020 = optimal_profit - optimal_profit_2018 - optimal_profit_2019

optimal_profits_df = pd.DataFrame({'Year': ['2018','2019','2020','Total'],
                      'Optimal Profit': [optimal_profit_2018,optimal_profit_2019,optimal_profit_2020,optimal_profit]})
optimal_profits_df.to_excel("optimal_profits.xlsx",sheet_name = 'Optimal Profits', index=False)