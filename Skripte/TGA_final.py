import pandas as pd
import numpy as np
from matplotlib.dates import DateFormatter
import os
from pandas.plotting import register_matplotlib_converters

register_matplotlib_converters()
import matplotlib.pyplot as plt
import Utilities


def run_tga():
    #try:
        """führt das gesamte Skript zur TGA aus. Raus kommt ein DataFrame mit den eingelesenen Proben,
        TGA-Faktor und Probengewicht"""
        # ____________________________Read in txt-file and store in a dataframe _________________________________________ #
        # TODO Pfad anpassen
        os.chdir('Rohdaten_TGA')

        sample_list = []

    #except ValueError or FileNotFoundError:
    #    print('error!')

        for f in os.listdir():
            data = Utilities.read_txt(f)
            sample_list.append(data)

        # _____________________________ slicing of rows and append in a new dataframe index-sorted________________________ #
        final_sample_list = [] # sammelt alle Dataframes, pro Probe einen, mit den Einzelmessungen und jeweiligen Gewichten

        for df in sample_list:
            appended_data = Utilities.slice_df(df)
            appended_sample_dfs = []
            for sample_df in appended_data:
                new_df = Utilities.prep_df(sample_df)
                appended_sample_dfs.append(new_df)

            df_final_raw = pd.concat(appended_sample_dfs, axis=1)
            df_final_raw.interpolate(method='linear', inplace=True)
            df_final_raw.sort_index(axis=1, inplace=True)

            # TODO evtl anpassen, wenn Daten in einem anderen Verzeichnis gespeichert sind
            os.chdir(os.path.dirname(os.getcwd()))
            os.chdir(os.path.dirname(os.getcwd()))
            os.chdir(os.path.dirname(os.getcwd()))

            # TODO evtl Pfad anpassen oder rausnehmen
            # df_final_raw.to_excel(r"Csv files/" + str(list(df_final_raw.columns)[0]) + '.xlsx')

        # ____________________normalising weight and heatflow_________________________________________________ #

            c = 0
            appended_data_col = []
            sample_weights = []
            while c < len(df_final_raw.columns):
                data = Utilities.sep_col(df_final_raw, c)
                # print(data)
                data, sample_weight = Utilities.normalise(data)
                sample_weights.append(sample_weight)
                appended_data_col.append(data)
                c = c + 2
                #print(c)

            df_final = pd.concat(appended_data_col, axis=1)

            # TODO evtl Pfad anpassen oder rausnehmen
            # df_final.to_excel(fr"Csv files/{str(list(df_final.columns)[0])}.xlsx")

        # _________________ choose everything that has weight in it and HF in it _________________________________________ #

            df_weight = Utilities.choose_columns(df_final, "weight", "weight")

            df_HF = Utilities.choose_columns(df_final, "HF", "HF")

            #Utilities.plot_two_df(df_weight, df_HF)

            df_weight_stats = Utilities.calc_stats(df_weight)
            df_HF_stats = Utilities.calc_stats(df_HF)

            df_single_measurement_list = []
            for col in df_weight:
                new_df = pd.concat([df_weight[col], df_weight_stats["mean"], df_weight_stats["stdev"]], axis=1)
                df_single_measurement_list.append(new_df)

            tga_factor_dict = {}
            stabw_list = []
            for sample in df_single_measurement_list:
                sample["factor"] = (np.sqrt((sample.iloc[:, 0] - sample["mean"])**2))/sample["stdev"]
                sample_name = sample.columns[0].split(",")
                tga_factor_dict[sample_name[0]] = sample["factor"].mean()
                stabw_list.append(sample["stdev"].mean())

            tga_summary = pd.DataFrame(tga_factor_dict.items(), columns=['Sample', 'TGA Faktor'])
            tga_summary = pd.concat([tga_summary, pd.Series(stabw_list), pd.Series(sample_weights)], axis=1)
            tga_summary.columns = ['Standardabweichung [%]' if x == 0 else x for x in tga_summary.columns]
            tga_summary.columns = ['Gewicht' if x == 1 else x for x in tga_summary.columns]

            final_sample_list.append(tga_summary)

            # TODO evtl Pfad anpassen
            os.chdir('C:\\Users\\Juschi\\PycharmProjects\\MP_calc')

        tga_summary_final = pd.concat(final_sample_list)
        tga_summary_final.reset_index(inplace=True)
        tga_summary_final.drop("index", axis=1, inplace=True)

        return tga_summary_final

   # except ValueError or FileNotFoundError:
   #     print('error!')



#tga_final = run_tga()


