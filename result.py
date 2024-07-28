import pandas as pd


# Gets the excel file and converts it to a CSV file
input_files = ['2024台大電機營P1-電子工程學研究所.xlsx', '2024台大電機營P2-電信工程學研究所生醫電子.xlsx', '2024台大電機營P3-光電工程學研究所 &電機工程學系碩博士班.xlsx']
output_files = ['P1_raw.csv', 'P2_raw.csv', 'P3_raw.csv']

for input_file, output_file in zip(input_files, output_files):
    # Read the input XLSX file
    raw_df = pd.read_excel(input_file, sheet_name='Final Scores', skiprows=2)
    # Save the DataFrame to CSV
    raw_df.to_csv(output_file, index=False)



input_files = ['P1_raw.csv', 'P2_raw.csv', 'P3_raw.csv']
output_files = ['P1_result.csv', 'P2_result.csv', 'P3_result.csv']

# Process each input file, calculate the sectional result to each group, and save the result to the output file
for input_file, output_file in zip(input_files, output_files):
    # Read the input CSV file
    raw_df = pd.read_csv(input_file)
    sorted_df = raw_df.sort_values(by=['Player'], axis=0, ascending=True, inplace=False, kind='quicksort', na_position='last', ignore_index=False)
    grouped_df = pd.DataFrame()
    grouped_df['Group'] = raw_df['Player'].str[:3]
    grouped_df['Total Score (points)'] = raw_df['Total Score (points)']
    grouped_df = grouped_df.groupby('Group')['Total Score (points)'].sum().reset_index()

    grouped_df.to_csv(output_file, index=False)


# For the final result
input_files = ['P1_result.csv', 'P2_result.csv', 'P3_result.csv']
output_files = ['Final_result_group.csv', 'Final_result_rank.csv']


final_df = pd.DataFrame()
final_df['Group'] = ['G01', 'G02', 'G03', 'G04', 'G05', 'G06', 'G07', 'G08', 'G09', 'G10']
final_df['Points'] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]


for input_file in input_files:
    section_df = pd.read_csv(input_file)
    for index, row in section_df.iterrows():
        group = row['Group']
        score = row['Total Score (points)']
        final_df.loc[final_df['Group'] == group, 'Points'] += score


final_df['Rank'] = final_df['Points'].rank(ascending=False, method='min').astype(int)


print('\n\n\n-------------照小隊排--------------')
# prints the sorted group result
print(final_df)
final_df.to_csv('Final_result_group.csv', index=False)

print('\n\n\n-------------照名次排--------------')
final_df = final_df.sort_values(by=['Rank'], axis=0, ascending=True, inplace=False, kind='quicksort', na_position='last', ignore_index=False)
final_df.to_csv('Final_result_rank.csv', index=False)

#prints the sorted rank result
print(final_df)

