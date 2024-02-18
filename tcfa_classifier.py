import pandas as pd
import xlsxwriter

print("Start read excel file")

target_file = 'matching.xlsx'
output_file = 'matching_output.xlsx'

matching = pd.read_excel(target_file, index_col=False)

# To make the index part of the DataFrame, I used the first row of the DataFrame.
index_df = matching[1:]
index_df = index_df[index_df.columns[:6]]

# Add two rows for each row and find the row numbers to fill.
# There are ivus frames for each row, so I used them as a reference.

# An empty list to store the index numbers of the DataFrame.
index_n = []
# An empty list to store the new index numbers. I will use them when I merge the two DataFrames at the end.
new_index_n = []

count = 0
for n, i in enumerate(matching['Unnamed: 6']):
    if i == 'Ivus frame':
        index_n.append(n - 1)
        new_index_n.append(n + count - 1)
        count += 2

empty_list = [''] * len(index_df.columns)
empty_df = pd.DataFrame([empty_list], columns=index_df.columns)
new_index_df = pd.DataFrame(columns=index_df.columns)

total_rows = len(index_n) * 2 + len(index_df.index)
count = 2

print('Original Rows  : ', len(index_df.index))
print('Total New Rows : ', total_rows)

for i in range(len(index_df.index)):
    if i in index_n:
        new_index_df = new_index_df.append(index_df.iloc[i])
        for step in range(2):
            new_index_df = new_index_df.append(empty_df)
    else:
        new_index_df = new_index_df.append(empty_df)


def T_ROW(t_row):
    """
    To make T65 & T200 row
    t_row: a row of DataFrame
    """
    t65_row = ['TCFA65 0/1']
    t200_row = ['tcfa200 0/1']
    for n, t in enumerate(t_row):
        if t == 0:
            t65_row.append(0.)
            t200_row.append(0.)
        elif t == 1:
            t65_row.append(0.)
            t200_row.append(1.)
        elif t == 2:
            t65_row.append(1.)
            t200_row.append(1.)
        else:
            # Had to fill the NaN with empty space ''.
            t65_row.append('')
            t200_row.append('')
    # Deleted the first element of the list.
    t65_row.remove(t65_row[1])
    t200_row.remove(t200_row[1])

    return t65_row, t200_row

# ### Eliminated the unnecessary columns ###
matching = matching.drop(matching.columns[0:6], axis=1)

matching = matching[1:]
index_col = matching.columns[0]

df = matching
new_df = df

TCFA_n = []
count = 0

for n, i in enumerate(df[df.columns[0]]):
    if i == 'TCFA':
        TCFA_n.append(n + count)
        count += 2

for n in TCFA_n:
    t65_row, t200_row = T_ROW(new_df.iloc[n])
    up_df = new_df.iloc[:n + 1]
    down_df = new_df.iloc[n + 1:]
    new_two_rows = pd.DataFrame([t65_row, t200_row], columns=df.columns)
    merged_df = up_df.append(new_two_rows)
    new_df = merged_df.append(down_df)


# ### Merge two DataFrames ###
# By default, the index is not reset, so I reset the index and deleted the first column.
new_df = new_df.reset_index()
new_df = new_df.drop(new_df.columns[0], axis=1)

new_index_df = new_index_df.reset_index()
new_index_df = new_index_df.drop(new_index_df.columns[0], axis=1)

result = pd.concat([new_index_df, new_df], axis=1)
result.head()


# ### Save the DataFrame to Excel file ###
# If you want to save the contents inside, you can run the code below.
# result.to_excel('matching_tcfa_intermediate_output.xlsx', columns=None, index=None, encoding='utf-8')

# ### Merge the cells of the DataFrame ###
# Because pandas can't merge cells, I used the xlsxwriter module to merge them.
index_list = new_index_n
index_list.append(new_index_n[-1] + 7)

# The NaN values of the two DataFrames were filled with ''.
result = result.fillna('')
new_df = new_df.fillna('')


workbook = xlsxwriter.Workbook(output_file)
worksheet = workbook.add_worksheet()

# ### Merge format ###
merge_format = workbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter'})

# Write the column headers with the defined format.
for i in range(len(index_list) - 1):
    worksheet.merge_range('A' + str(new_index_n[i] + 1) + ':A' + str(new_index_n[i + 1]), result.iloc[new_index_n[i]][0], merge_format)
    worksheet.merge_range('B' + str(new_index_n[i] + 1) + ':B' + str(new_index_n[i + 1]), result.iloc[new_index_n[i]][1], merge_format)
    worksheet.merge_range('C' + str(new_index_n[i] + 1) + ':C' + str(new_index_n[i + 1]), result.iloc[new_index_n[i]][2], merge_format)
    worksheet.merge_range('D' + str(new_index_n[i] + 1) + ':D' + str(new_index_n[i + 1]), result.iloc[new_index_n[i]][3], merge_format)
    worksheet.merge_range('E' + str(new_index_n[i] + 1) + ':E' + str(new_index_n[i + 1]), result.iloc[new_index_n[i]][4], merge_format)
    worksheet.merge_range('F' + str(new_index_n[i] + 1) + ':F' + str(new_index_n[i + 1]), result.iloc[new_index_n[i]][5], merge_format)

# If you want to add the contents of the DataFrame, you can run the code below.
for i in range(len(result.index)):
    for j in range(len(new_df.columns)):
        worksheet.write(i, j + 6, new_df.iloc[i][j])

workbook.close()
print("Done")
