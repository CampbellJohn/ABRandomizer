from pandas import *
import pandas as pd
import sys
import numpy as np

class XLS():
    # The XLS class exists to load and save date to & from .xlsx files.
    def __init__(self):
        pass
    # Load XLS file, and return a DataFrame.
    def Load(self, filename, h, i):
        xls = pd.ExcelFile(filename)
        df = xls.parse('Worksheet1', header=h, index_col=i)
        return df
    # Print the DataFrame to an .xlsx sheet.
    def Save(self, data, filename):
        writer = ExcelWriter(filename)
        data.to_excel(writer, 'Worksheet1', index=False)
        writer.save()

# Pull in lead sheet and # of groups.
# Count how many leads we've got.
# Generate a column of random numbers.
# Sort leads.
# Split leads into separate dataframes.
# Export dataframes.

class LeadSplit():
    def __init__(self, dataframe, grpnum):
        self.df = dataframe
        self.num = grpnum
        self.numleads = len(self.df.index)
        # Randomize Data, then spit it out.
        self.df = self.Randomize(self.df)
        self.df = self.Split(self.df)
        self.df = self.Export(self.df)

    def Randomize(self, df):
        # Create new column, 'Rand' here. Populate it with random numbers.
        df['Rand'] = np.random.rand(len(df))
        df = df.sort('Rand')
        return df

    def Split(self, df):
        n = self.num
        df['Grp'] = ''
        g = 1
        for (index, row) in df.iterrows():
            if g > n:
                g = 1
            df.ix[index, 'Grp'] = g
            g += 1
        return df

    def Export(self, df):
        n = self.num + 1
        for i in range(1, n):
            print(i)
            group = df.groupby('Grp').get_group(i)
            group = group.drop('Rand', axis=1)
            group = group.drop('Grp', axis=1)

            XLS.Save(group, "CCS_OnG_EDM3_Test" + str(i) + ".xlsx")
        return df



XLS = XLS()

Data = XLS.Load('Leads.xls', 0, False)

Leads = LeadSplit(Data, 2)

#New = XLS.Save(Leads.df, 'Done.xlsx')
