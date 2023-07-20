import traceback

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime as dt
import tkinter as tk
import os
from tkinter import filedialog


# Excel format must be the following:
# task  | person  | start  | end
# ________________________________
#  str  |  str    | Date   | Date
#  str  |  str    | Date   | Date
#  ...  |  ...    | ...    | ...
#  str  |  str    | Date   | Date


def main(excel_file):
    try:
        # print(os.path.basename(excel_file))
        excel_file_name = os.path.basename(excel_file).split('.')[0]
        df = pd.read_excel(excel_file)

        pd.set_option('display.max_columns', None)

        df['days_to_start'] = (df['start'] - df['start'].min()).dt.days
        df['days_to_end'] = (df['end'] - df['start'].min()).dt.days
        df['task_duration'] = df['days_to_end'] - df['days_to_start'] + 1  # to include also the end date

        df = df.sort_values(by=['start', 'end'], ascending=False).reset_index()

        # df['start'] = pd.to_datetime(df['start'])

        # Set the figure size
        fig, ax = plt.subplots(figsize=(17, 9))  # Adjust the width (8) as needed

        # Color by person name
        # Create a dictionary using dict comprehension
        color_list = ['lightcoral', 'yellow', 'skyblue', 'lightsalmon', 'forestgreen',
                      'lightseagreen', 'navy', 'crimson']

        people = df['person'].tolist()
        unique_people = list(set(people))

        people_colors = {person: color_list[index] for index, person in enumerate(unique_people)}
        # print(people_colors)

        for index, row in df.iterrows():
            ax.barh(y=index, width=row['task_duration'], left=row['days_to_start'], color=people_colors[row['person']])

        xticks = np.arange(0, df['days_to_end'].max() + 3)

        xticklabels = [""]
        for date in pd.date_range(start=df['start'].min(), end=df['end'].max() + dt.timedelta(days=2)):
            if date.weekday() in [5, 6]:  # Saturday or Sunday
                xticklabels.append("")
            else:
                xticklabels.append(date.strftime("%m/%d"))

        # 5
        ax.set_xticks(xticks)

        # Manually adjust the label positions
        label_positions = np.arange(0, df['days_to_end'].max() + 2)
        label_positions = np.concatenate((label_positions, [df['days_to_end'].max() + 1]))  # Include the last label

        # Extract xticklabels based on label_positions using np.take
        label_ticklabels = np.take(xticklabels, label_positions)

        ax.set_xticks(label_positions)
        ax.set_xticklabels(label_ticklabels, ha='right', rotation=45)

        # Set the Y-axis tick positions and labels starting from 1
        ax.set_yticks(df.index)
        ax.set_yticklabels(df['person'])

        for i, duration in enumerate(df['task_duration']):
            ax.text(df['days_to_start'][i] + duration + 0.3, i, str(df['task'][i]))

        ax.xaxis.grid(True, alpha=0.5)
        plt.title(f'Proposal Project Management {excel_file_name}', fontsize=15)
        plt.subplots_adjust(right=0.85)
        plt.show()

    except Exception:
        print("You had an error:")
        print(traceback.format_exc())

    return


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    main(filedialog.askopenfilename(title="Select an Excel document", filetypes=(("Excel files", "*.xlsx"),
                                                                                 ("All files", "*.*"))))
