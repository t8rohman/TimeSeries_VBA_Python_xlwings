'''
Taufiqur Rohman 2022

This is the script from the main file to run all the RunPython sub function
in the parent file, timeseries_forecast.xlsm. In total, there are there are 6 
functions inside this script.

Credits for xlwings team to make such a wonderful module for us to link Python
to Excel users.
'''

import xlwings as xw
import pandas as pd
import numpy as np

from matplotlib import pyplot as plt
from sklearn.metrics import mean_absolute_error, mean_squared_error
from statsmodels.tsa.seasonal import seasonal_decompose
from statsmodels.tsa.holtwinters import SimpleExpSmoothing
from statsmodels.tsa.holtwinters import ExponentialSmoothing


def main():  # template from xlwings quickstart, only used for quick debugging
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"
        
        
def data_read():  # to read the data from the parent file
    wb = xw.Book.caller()
    sheets_data = wb.sheets['data']
    df_data = sheets_data["A1"].expand().options(pd.DataFrame).value
    df_data.set_index("period", inplace=True)
    return df_data
        

def data_diagnose():  # diagnose the components of the time series
    wb = xw.Book.caller()
    sheets_data = wb.sheets['data']
    sheets_diagnose = wb.sheets['diagnose']
    
    df_data = data_read()  # call the data reading function
    
    seasonal_freq = sheets_data['H26'].value  
    seasonal_periods = sheets_data['H27'].value
    df_data.index.freq = seasonal_freq  # to set the index frequency, as the exponential smoothing model needs it
    
    
    def decompose_plot():  # decompose all components in the time series
        decompose_result = seasonal_decompose(df_data['observed'], model='multiplicative')
        plt.style.use("seaborn")
        ax = decompose_result.plot()
        fig = ax.get_figure()
        return fig
    
    
    # create simple exponential smoothing plot 
    fit_hwes1 =  SimpleExpSmoothing(df_data['observed']).fit(optimized=True, use_brute=True)
    df_data['hwes1'] = fit_hwes1.fittedvalues
    
    def hwes1_plot():  
        plt.style.use("seaborn")
        ax = df_data[['observed', 'hwes1']].plot(title='Simple Exponential Smoothing')
        fig = ax.get_figure()  
        return fig
    
    
    # create holt's double exponential smoothing
    fit_hwes2_add =  ExponentialSmoothing(df_data['observed'], trend='add').fit()
    fit_hwes2_mul =  ExponentialSmoothing(df_data['observed'], trend='mul').fit()
    df_data['hwes2_add'] = fit_hwes2_add.fittedvalues
    df_data['hwes2_mul'] = fit_hwes2_mul.fittedvalues
    
    def hwes2_plot():   
        plt.style.use("seaborn")
        ax = df_data[['observed', 'hwes2_add', 'hwes2_mul']].plot(title="Holt's Double Exponential Smoothing: Additive and Multiplicative Trend")
        fig = ax.get_figure()  
        return fig
    
    
    # create holt-winter's triple exponential smoothing
    fit_hwes3_add = ExponentialSmoothing(df_data['observed'], trend='add', seasonal='add', seasonal_periods=seasonal_periods).fit()
    fit_hwes3_mul = ExponentialSmoothing(df_data['observed'], trend='mul', seasonal='mul', seasonal_periods=seasonal_periods).fit()
    df_data['hwes3_add'] = fit_hwes3_add.fittedvalues
    df_data['hwes3_mul'] = fit_hwes3_mul.fittedvalues
    
    def hwes3_plot():
        plt.style.use("seaborn")
        ax = df_data[['observed', 'hwes3_add', 'hwes3_mul']].plot(title="Holt-Winter's Triple Exponential Smoothing: Additive and Multiplicative Seasonality")
        fig = ax.get_figure()  
        return fig
    
    
    # plotting all chart on diagnose sheet
    decompose_plot = decompose_plot()
    hwes1_plot = hwes1_plot()
    hwes2_plot = hwes2_plot()
    hwes3_plot = hwes3_plot()
    
    decompose_chart = sheets_diagnose.pictures.add(decompose_plot, name="decompose_chart",
                                                   top=sheets_diagnose["C5"].top, 
                                                   left=sheets_diagnose["C5"].left)
    decompose_chart.width, decompose_chart.height = decompose_chart.width * 1.5, decompose_chart.height * 1.5

    hwes1_chart = sheets_diagnose.pictures.add(hwes1_plot, name="hwes1_chart",
                                               top=sheets_diagnose["E5"].top, 
                                               left=sheets_diagnose["E5"].left)
    hwes1_chart.width, hwes1_chart.height = hwes1_chart.width * 0.6, hwes1_chart.height * 0.6
    
    hwes2_chart = sheets_diagnose.pictures.add(hwes2_plot, name="hwes2_chart",
                                               top=sheets_diagnose["E20"].top, 
                                               left=sheets_diagnose["E20"].left)
    hwes2_chart.width, hwes2_chart.height = hwes2_chart.width * 0.6, hwes2_chart.height * 0.6
    
    hwes3_chart = sheets_diagnose.pictures.add(hwes3_plot, name="hwes3_chart",
                                               top=sheets_diagnose["E35"].top, 
                                               left=sheets_diagnose["E35"].left)
    hwes3_chart.width, hwes3_chart.height = hwes3_chart.width * 0.6, hwes3_chart.height * 0.6


def forecast_hwes1():  # used to forecast using simple exponential smoothing
    wb = xw.Book.caller()
    sheets_data = wb.sheets['data']
    sheets_forecast = wb.sheets['forecast']
    
    df_data = data_read()  # call the data reading function
    
    df_data.index.freq = sheets_data['H26'].value
    period_forecast = int(sheets_data['H32'].value)
    
    fit_hwes1 =  SimpleExpSmoothing(df_data['observed']).fit(optimized=True, use_brute=True)
    
    df_hwes1 = pd.DataFrame(
        np.c_[df_data['observed'], fit_hwes1.level, fit_hwes1.trend, fit_hwes1.season, fit_hwes1.fittedvalues],
        columns=['observed', 'level', 'trend', 'season', 'forecast'],
        index=df_data.index,
        )
    
    mae = mean_absolute_error(df_hwes1['observed'], df_hwes1['forecast'])
    rms = mean_squared_error(df_hwes1['observed'], df_hwes1['forecast'], squared=False)
    sheets_forecast['J1'].value ="Simple Exponential Smoothing"
    sheets_forecast['J2'].value = mae
    sheets_forecast['J3'].value = rms

    df_hwes1 = df_hwes1.append(fit_hwes1.forecast(period_forecast).rename('forecast').to_frame(), sort=True)
    sheets_forecast['A1'].value = df_hwes1
    
    forecast_val = fit_hwes1.forecast(period_forecast) 


    def plot():  # create plot, but need to nest it in a function to make it callable from a variable
        fig = plt.figure()
        plt.style.use("seaborn")
        df_hwes1['observed'].plot(legend=True, label='actual')
        forecast_val.plot(legend=True, label='forecast')
        plt.title("Actual and Observed Data Using Simple Exponential Smoothing (SES)")
        return fig
        
        
    fig = plot()
    hwes1_chart = sheets_forecast.pictures.add(fig, name="chart_forecast",
                                               top=sheets_forecast["I6"].top, 
                                               left=sheets_forecast["I6"].left)
    hwes1_chart.width, hwes1_chart.height = hwes1_chart.width * 0.7, hwes1_chart.height * 0.8    


def forecast_hwes2():  # used to forecast using holt's double exponential smoothing
    wb = xw.Book.caller()
    sheets_data = wb.sheets['data']
    sheets_forecast = wb.sheets['forecast']
    
    df_data = data_read()  # call the data reading function
    
    df_data.index.freq = sheets_data['H26'].value
    period_forecast = int(sheets_data['H32'].value)
    method = sheets_data['H33'].value
    
    fit_hwes2 =  ExponentialSmoothing(df_data['observed'], trend=method).fit()
    
    df_hwes2 = pd.DataFrame(
        np.c_[df_data['observed'], fit_hwes2.level, fit_hwes2.trend, fit_hwes2.season, fit_hwes2.fittedvalues],
        columns=['observed', 'level', 'trend', 'season', 'forecast'],
        index=df_data.index,
        )
    
    mae = mean_absolute_error(df_hwes2['observed'], df_hwes2['forecast'])
    rms = mean_squared_error(df_hwes2['observed'], df_hwes2['forecast'], squared=False)
    sheets_forecast['J1'].value ="'Holt's 2 Components"
    sheets_forecast['J2'].value = mae
    sheets_forecast['J3'].value = rms

    df_hwes2 = df_hwes2.append(fit_hwes2.forecast(period_forecast).rename('forecast').to_frame(), sort=True)
    sheets_forecast['A1'].value = df_hwes2
    
    forecast_val = fit_hwes2.forecast(period_forecast) 


    def plot():  # create plot, but need to nest it in a function to make it callable from a variable
        fig = plt.figure()
        plt.style.use("seaborn")
        df_hwes2['observed'].plot(legend=True, label='actual')
        forecast_val.plot(legend=True, label='forecast')
        plt.title("Actual and Observed Data Using Holt's 2 Components")
        return fig
        
        
    fig = plot() 
    hwes2_chart = sheets_forecast.pictures.add(fig, name="chart_forecast",
                                               top=sheets_forecast["I6"].top, 
                                               left=sheets_forecast["I6"].left)
    hwes2_chart.width, hwes2_chart.height = hwes2_chart.width * 0.7, hwes2_chart.height * 0.8    


def forecast_hwes3():   # used to forecast using holt-winter's triple exponential smoothing
    wb = xw.Book.caller()
    sheets_data = wb.sheets['data']
    sheets_forecast = wb.sheets['forecast']
    
    df_data = data_read()  # call the data reading function
    
    df_data.index.freq = sheets_data['H26'].value
    seasonal_periods = int(sheets_data['H27'].value)
    period_forecast = int(sheets_data['H32'].value)
    method = sheets_data['H33'].value
    
    fit_hwes3 = ExponentialSmoothing(df_data['observed'], trend=method, seasonal=method, seasonal_periods=seasonal_periods).fit()
    
    df_hwes3 = pd.DataFrame(
        np.c_[df_data['observed'], fit_hwes3.level, fit_hwes3.trend, fit_hwes3.season, fit_hwes3.fittedvalues],
        columns=['observed', 'level', 'trend', 'season', 'forecast'],
        index=df_data.index,
        )
    
    mae = mean_absolute_error(df_hwes3['observed'], df_hwes3['forecast'])
    rms = mean_squared_error(df_hwes3['observed'], df_hwes3['forecast'], squared=False)
    sheets_forecast['J1'].value ="'Holt-Winter's 3 Components"
    sheets_forecast['J2'].value = mae
    sheets_forecast['J3'].value = rms

    df_hwes3 = df_hwes3.append(fit_hwes3.forecast(period_forecast).rename('forecast').to_frame(), sort=True)
    sheets_forecast['A1'].value = df_hwes3
    
    forecast_val = fit_hwes3.forecast(period_forecast) 


    def plot():  # create plot, but need to nest it in a function to make it callable from a variable
        fig = plt.figure()
        plt.style.use("seaborn")
        df_hwes3['observed'].plot(legend=True, label='actual')
        forecast_val.plot(legend=True, label='forecast')
        plt.title("Actual and Observed Data Using Holt-Winter's 3 Components")
        return fig
        
        
    fig = plot()
    hwes3_chart = sheets_forecast.pictures.add(fig, name="chart_forecast",
                                               top=sheets_forecast["I6"].top, 
                                               left=sheets_forecast["I6"].left)
    hwes3_chart.width, hwes3_chart.height = hwes3_chart.width * 0.7, hwes3_chart.height * 0.8
    

if __name__ == "__main__":
    xw.Book("timeseries_forecast.xlsm").set_mock_caller()
    main()
