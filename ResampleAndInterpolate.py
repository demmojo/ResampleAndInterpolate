# Copyright 2019 Blok-Z
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.


"""ResampleAndInterpolate changes the sampling rate of original data from every 1 hour to every 10 seconds and then
   applies linear interpolation to the missing data.
"""

import pandas as pd

# Assign spreadsheet filename to `file`
file = 'borusan.xlsx'

# Load spreadsheet
xl = pd.ExcelFile(file)

index = pd.date_range('2018-10-01', periods=2568, freq='H')  # creates a series of dates in hour intervals

# pandas dataframes are used to store the original and interpolated data
original_data = xl.parse(xl.sheet_names[0], usecols=['Actual Production (MWh)'])  # store Actual Production values
original_data.index = index  # sets index to date series
original_data.index.name = 'Date (Year-Month-Day Hour:Minute:Second'
interpolated_data = original_data.resample('10S').interpolate()  # upsample data to every 10sec and linearly interpolate

# Specify a writer for excel files
writer = pd.ExcelWriter('borusan-interpolated.xlsx',  engine='xlsxwriter')

# Write your DataFrame to a file
interpolated_data.to_excel(writer, 'Sheet1')

# Save the result
writer.save()


