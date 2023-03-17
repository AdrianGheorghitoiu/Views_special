import pandas as pd

import numpy as np
[2]
#read file from path

viewed_content = pd.read_excel('../Viewed Content/ViewedContent_Jun22.xlsx')

viewed_content.head()

[3]
#dropping unnecessary columns

viewed_content.drop(['Id', 'Url', 'Last Viewed Date (Coordinated Universal Time)', 'Group Id', 'Application Id'], inplace=True, axis=1)

viewed_content.head()

[4]
#renaming column(s)

viewed_content.columns=['Title', 'User Id', 'Username', 'Type', 'date', 'Total View Count', 'Group Name', 'Application Name']

viewed_content.head()

[5]
#finding out the quartiles and interquartile range to eliminate the outliers. In this case, the calculations are as follows:
#Q3(44) - Q1(10)= IQR(34)
#IQR(34) * 1.5 = 51
#51 + Q3(44) = 95 (anything greather than or equal to this value is considered an outlier )

viewed_content.describe()

[8]
#filtering out the outliers

views = viewed_content[(viewed_content['Total View Count']<=95)]

views

[9]
#creating the Excel file

writer = pd.ExcelWriter('Views_Jun2022.xlsx',
                       engine='xlsxwriter',
                       engine_kwargs={'options': {'strings_to_urls':False}}
                       )
views.to_excel(writer)

writer.close()