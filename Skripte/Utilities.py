import pandas as pd
import pylab as plt
import numpy as np
import random
from matplotlib.dates import DateFormatter
import os
import shutil
from pandas.plotting import register_matplotlib_converters

register_matplotlib_converters()
import openpyxl
# import xlrd
import matplotlib.pyplot as plt
from textwrap import wrap
from datetime import datetime
#import seaborn as sns

# ____________________functions for dataframe slicing and data preparation___________________________________________ #


def read_txt(folder_path):
    df = pd.read_csv(folder_path, decimal='.', sep='\n', header=None, encoding='latin-1')
    df = df[0].str.split('[ ]{1,}', expand=True)
    df.drop([0], axis=1, inplace=True)
    df.drop([0, 1])
    return df


def slice_df(df):
    """Slice dataframe based on certain conditions that tell it when a new sample starts and ends,
    returns a list of the dataframe slices"""
    beginning_of_sample = (np.where(df[1] == '[s]')[0]).tolist()
    end_of_sample = (np.where(df[1] == 'Name:')[0]).tolist()
    if len(df) < 8000:
        end_of_sample.append(len(df))
    else:
        end_of_sample.append(len(df)-2)
    end_of_sample = end_of_sample[1:]
    sample_locations = zip(beginning_of_sample, end_of_sample)
    sample_locations_dict = dict(sample_locations)
    appended_data = []
    for sample_start in sample_locations_dict:
        df_slice = df.iloc[sample_start: sample_locations_dict[sample_start]]
        appended_data.append(df_slice)
    return appended_data


def sep_col(dcol, y):
    """takes a dataframe and puts two columns in a new dataframe. Returns a single dataframe."""
    df_col = dcol.iloc[:, y:y + 2]
    # new_col = []
    # for i in df_col.columns:
    #     i = i + str(y)
    #     new_col.append(i)
    # df_col.columns = new_col
    return df_col


def prep_df(mydata):
    """Takes a dataframe and returns 'cleaned' data with Sample_T as index """
    mydata = mydata.copy()
    trial = mydata.iloc[-1, 0] + mydata.iloc[-1, 1]
    mydata = mydata.iloc[1: -3]
    mydata.columns = ['Index', 'Time', 'Sample_T', 'Reference_T', '']
    for (columnName, columnData) in mydata.iteritems():
        mydata[columnName] = pd.to_numeric(columnData, errors='coerce')
    mydata['rounded_Sample_T'] = mydata['Sample_T'].round(decimals=1)
    mydata = mydata.groupby('rounded_Sample_T').mean()
    mydata = mydata[['']]
    new_col = []
    for i in mydata.columns:
        i = i + str(trial)
        new_col.append(i)
    mydata.columns = new_col
    return mydata


def normalise(df_sample):
    df_sample = df_sample.copy()
    # df_sample.interpolate(method='linear', inplace=True)
    df_sample['Sample_T'] = df_sample.index
    df_sample.drop(df_sample[df_sample.Sample_T < 30].index, inplace=True)
    df_sample.drop(df_sample[df_sample.Sample_T > 590].index, inplace=True)
    suggested_sample_weight_temp = [104.9, 105.0, 105.1]
    sample_weight_temp = [value for value in suggested_sample_weight_temp if value in df_sample.Sample_T]
    weight_temp = random.choice(sample_weight_temp)
    df_sample.drop(['Sample_T'], axis=1, inplace=True)
    # sample_T_list = [round(value*0.1, 1) for value in range(300, 5900)]
    # print(sample_T_list)
    # df_sample.interpolate(method='linear', inplace=True)
    df_sample.dropna(inplace=True)
    heatflow = all(value > 0 for value in df_sample.iloc[:, 0])
    if heatflow:
        df_sample['n_weight'] = (df_sample.iloc[:, 0] / df_sample.loc[weight_temp].iloc[0]) * 100
        df_sample['n_HF'] = df_sample.iloc[:, 1] / df_sample.iloc[:, 0]
        sample_weight = df_sample.loc[weight_temp].iloc[0]
    else:
        df_sample['n_weight'] = (df_sample.iloc[:, 1] / df_sample.loc[weight_temp].iloc[1]) * 100
        df_sample['n_HF'] = df_sample.iloc[:, 0] / df_sample.iloc[:, 1]
        sample_weight = df_sample.loc[weight_temp].iloc[1]
    col_list = list(df_sample.columns)
    df_sample.rename(columns={'n_weight': str(col_list[0]) + ' n_weight'}, inplace=True)
    df_sample.rename(columns={'n_HF': str(col_list[0]) + ' n_HF'}, inplace=True)
    df_sample.drop(df_sample.columns[0], axis=1, inplace=True)
    return df_sample, sample_weight


def choose_columns(df, keyword1, keyword2):
    """chooses columns from a dataframe that contain keywords,
    (if only one keyword put it twice) and returns new dataframe"""
    column_list = []
    for col in df.columns:
        if (keyword1 in col) or (keyword1.lower() in col) or (keyword2 in col) or (keyword2.lower() in col):
            column_list.append(col)
    df_new = df[column_list]
    if len(column_list) == 0:
        return
    #df_new.to_excel(r"Csv files/" + str(list(df_new.columns)[0]) + '_' + keyword1 + '.xlsx')
    return df_new


def choose_columns_invert(df, keyword):
    """chooses all columns from a dataframe that DON'T contain the keyword and returns new dataframe"""
    column_list = []
    for col in df.columns:
        if (keyword not in col) and (keyword.lower() not in col):
            column_list.append(col)
    df_new = df[column_list]
    if len(column_list) == 0:
        return
    #df_new.to_excel(r"Csv files/" + str(list(df_new.columns)[0]) + '_' + keyword + '.xlsx')
    return df_new


# build function for calculating mean and stdv
def calc_stats(df):
    """calculates mean, standard deviation and upper and lower limit of a dataframe and returns them as series"""
    mean = df.mean(axis=1)
    stdev = df.std(axis=1)
    upper_std = mean + stdev
    lower_std = mean - stdev
    df_new = pd.concat([mean, stdev, upper_std, lower_std], axis=1)
    # name = df.columns[0][df.columns[0][7:len(df.columns[0])-19]]
    # cols = [f"{name}_mean", f"{name}_stdev", f"{name}_upper_std", f"{name}_lower_std"]
    cols = ["mean", "stdev", "upper_std", "lower_std"]
    df_new.columns = cols
    return df_new


# bulid function for plotting: def plotting(df):
def plot_two_df(df1, df2):
    """plots two dataframes underneath each other and the stats in a next column (average and standard deviation)"""
    df1_stats = calc_stats(df1)
    df2_stats = calc_stats(df2)

    fig3, ax = plt.subplots(nrows=2, ncols=2, figsize=[12, 8], sharex=True)

    # print(ax)
    for col in df1:
        ax[0][0].plot(df1[col])
        # ax[0].set_xlabel('Sample Temperature [°C]')
        ax[0][0].set_ylabel('normalised weight [-]')
    ax[0][0].plot(df1_stats["mean"], color="black")

    if len(df1.columns) <= 8:
        ax[0][0].legend(df1.columns)
    # ax[0][0].set_ylim([0, 1])
    # ax[0].set_xlim([0, 650])

    # start, end = ax[0].get_xlim()
    ax[0][0].xaxis.set_ticks(np.arange(0, 650, 50))
    ax[0][0].set_title('Rohdaten')

    # ax[0][0].grid('on', which='minor', axis='x', linestyle='--')
    # ax[0][0].grid('off', which='major', axis='x', linestyle='--')
    #
    # ax[0][0].grid('on', which='minor', axis='y', linestyle='--')
    # ax[0][0].grid('off', which='major', axis='y', linestyle='--')

    ax[0][1].plot(df1_stats['mean'], color='darkblue')
    ax[0][1].plot(df1_stats['upper_std'], color='dodgerblue')
    ax[0][1].plot(df1_stats['lower_std'], color='dodgerblue')
    ax[0][1].fill_between(df1_stats.index, df1_stats['lower_std'], df1_stats['upper_std'],
                          color='dodgerblue', alpha=0.2)
    ax[0][1].set_title('Mittelwert und Standardabweichung')

    for col in df2:
        ax[1][0].plot(df2[col])
        ax[1][0].set_xlabel('Sample Temperature [°C]')
        ax[1][0].set_ylabel('normalised heatflow [mW/mg]')
    ax[1][0].plot(df2_stats["mean"], color="black")

    if len(df2.columns) <= 8:
        ax[1][0].legend(df2.columns)

    ax[1][1].plot(df2_stats['mean'], color='darkred')
    ax[1][1].plot(df2_stats['upper_std'], color='orangered')
    ax[1][1].plot(df2_stats['lower_std'], color='orangered')
    ax[1][1].fill_between(df2_stats.index, df2_stats['lower_std'], df2_stats['upper_std'], color='orangered', alpha=0.2)
    ax[1][1].set_xlabel('Sample Temperature [°C]')

    axes_list = ax.flatten()

    for i in axes_list:
        i.grid('on', which='minor', axis='x', linestyle='--')
        i.grid('off', which='major', axis='x', linestyle='--')

        i.grid('on', which='minor', axis='y', linestyle='--')
        i.grid('off', which='major', axis='y', linestyle='--')

    # print(os.getcwd())

    sns.despine(left=False, bottom=False, right=True)

    #fig3.show()
    fig3.savefig('C:\\Users\\Juschi\\PycharmProjects\\MP_calc\\TGA_plot/' + str(list(df1.columns)[0]) + '.svg', bbox_inches='tight', dpi=450)

    return fig3


def boxplot(df1, df2, labels):
    """takes a dataframe and plots a boxplot.
    The Dataframe must contain a column named 'stdev', which will be plotted"""
    all_data = [df1['stdev'], df2['stdev']]
    print(type(all_data))
    fig3, ax = plt.subplots(figsize=[5, 8])
    ax.boxplot(all_data, labels=labels)
    ax.set_xticklabels(labels, wrap=True)
    print('done')
    # labels = labels


# dataframe = slice_df(2)
# dataframe = prep_df(dataframe)


