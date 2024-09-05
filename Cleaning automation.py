import os
import pandas as pd
import numpy as np

# Set the directory where your Excel files are stored
directory_path = 'C:/Users/QK615NU/OneDrive - EY/Desktop/Dossier de travail/DEV/BEL'

# List to hold data from each file
all_data = []

# Loop through each file in the directory
for filename in os.listdir(directory_path):
    if filename.endswith('.xlsx'):
        try:
            print(f'{filename} processing')
            file_path = os.path.join(directory_path, filename)
            # Your existing code to read and process each Excel file
            df = pd.read_excel(file_path, dtype=str, header=2)
            df_entityA = pd.read_excel(file_path, dtype=str, nrows=2, header=0)
            df_currency = pd.read_excel(file_path, dtype=str, nrows=3, header=0)

            # Extraire le nom et le code de l'entité A
            df['Entity A Name'] = df_entityA.columns[2]
            df['Entity A Code'] = df_entityA.columns[1]
            df['Company Currency'] = df.iloc[1, 2]
            df['Group Currency'] = df.iloc[1, 3]
            print(df[df['Unnamed: 0'] == '6'])
            df['Unnamed: 0'] = df['Unnamed: 0'].str.strip()
            df['Account Family'] = np.where(df['Unnamed: 0'] == '6', '6', np.where(df['Unnamed: 0'] == '7', '7', np.nan))
            df['Account Family'] = df['Account Family'].replace('nan', np.nan)
            df['Account Family'] = df['Account Family'].fillna(method='ffill')
            print(df)
            df['Partie B'] = df['Unnamed: 0'].str.strip()
            df['Unnamed: 0'] = df['Unnamed: 0'].str.strip()
            df['Unnamed: 0'] = np.where(df['Unnamed: 0'].str.startswith('E'), np.nan, df['Unnamed: 0'])
            df['Unnamed: 0'] = df['Unnamed: 0'].fillna(method='ffill')
            df['Partie B'] = np.where(df['Partie B'].str.len() != 4, np.nan, df['Partie B'])
            df = df[df['Fiscal year'].notna()]

            df = df.rename(columns={'Unnamed: 0':'Transactions','Fiscal year':'Entity B name',
                                    'Partie B':'Entity B Code', '2022':'Amount Company Currency 2022',
                                    'Unnamed: 3':'Amount Group Currency 2022','Unnamed: 4': 'Amount Document currency 2022',
                                    '2023':'Amount Company Currency 2023','Unnamed: 6':'Amount Group Currency 2023','Unnamed: 7':'Amount Document currency 2023'})
            
            df['Document Currency'] = np.where(df['Amount Document currency 2022'] == df['Amount Group Currency 2022'],df['Group Currency'],df['Company Currency'])
            df['Document Currency'] = np.where((df['Entity B name'].str.contains('USA')) | (df['Entity B name'].str.contains('Egypt')) |(df['Entity B name'].str.contains('VIETNAM')), 'USD',df['Document Currency'])
           # df['Document Currency_2023'] = np.where(df['Amount Document currency 2023'] == df['Amount Group Currency 2023'],df['Group Currency'],df['Company Currency'])
            #df['Document Currency_2023'] = np.where(df['Entity B name'].str.contains('USA'), 'USD',df['Document Currency_2023'])
            df['Document Currency'] = np.where(df['Entity B name'].str.contains('S.A.'), 'EUR',df['Document Currency'])
            #df['Document Currency_2023'] = np.where(df['Entity B name'].str.contains('S.A.'), 'EUR',df['Document Currency_2023'])
            

            df_2022 = df[['Account Family','Entity A Code','Entity A Name','Entity B Code','Entity B name','Transactions','Amount Company Currency 2022','Company Currency',
            'Amount Group Currency 2022','Amount Document currency 2022','Document Currency']]
        

            df_2023 = df[['Account Family','Entity A Code', 'Entity A Name','Entity B Code','Entity B name','Transactions','Amount Company Currency 2023','Company Currency',
            'Amount Group Currency 2023','Amount Document currency 2023','Document Currency']]

            df = pd.concat([df_2022,df_2023],ignore_index=True)
         

           # Assuming df is your DataFrame and numeric_columns is your list of columns to convert
            numeric_columns = ['Amount Company Currency 2022', 'Amount Group Currency 2022',
                               'Amount Document currency 2022','Amount Company Currency 2023','Amount Group Currency 2023'
                               ,'Amount Document currency 2023']

        # Convert each column in the list to a numeric type

            for col in numeric_columns:
                df[col] = pd.to_numeric(df[col].fillna(0), errors='coerce')
         
            df = df.groupby(['Account Family',
                'Entity A Code',
                'Entity A Name',
                'Entity B Code',
                'Entity B name',
                'Transactions',
                'Company Currency',
                'Document Currency',
               
            ]).agg({
                'Amount Company Currency 2022': 'sum',
                'Amount Group Currency 2022': 'sum',
                'Amount Document currency 2022': 'sum',
                'Amount Company Currency 2023': 'sum',
                'Amount Group Currency 2023': 'sum',
                'Amount Document currency 2023': 'sum'
            }).reset_index()
           
            df = df[['Account Family','Entity A Code', 'Entity A Name','Entity B Code','Entity B name','Transactions','Company Currency','Amount Company Currency 2022',
            'Amount Group Currency 2022','Amount Document currency 2022','Amount Company Currency 2023',
            'Amount Group Currency 2023','Amount Document currency 2023','Document Currency']]
            # Append the processed DataFrame to the list
            all_data.append(df)
            print(f'{filename} has been appended')
        except Exception as e:
            raise Exception(f"An error occurred while processing the file: {filename}") from e

# Concatenate all DataFrames into a single DataFrame
final_data = pd.concat(all_data, ignore_index=True)

# Export the final DataFrame to a new Excel file
final_data.to_excel('C:/Users/QK615NU/OneDrive - EY/Desktop/Dossier de travail/DEV/BEL/output.xlsx', index=False)
