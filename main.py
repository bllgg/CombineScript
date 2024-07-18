import pandas as pd
import glob


# Function to read all Excel files and combine them based on topics
def combine_excel_files(directory_path, output_file):
    # List to hold dataframes
    dataframes = []

    # Read all xlsx files in the directory
    for file in glob.glob(f"{directory_path}/*.xlsx"):
        df = pd.read_excel(file)
        dataframes.append(df)

    # Concatenate all dataframes
    combined_df = pd.concat(dataframes, ignore_index=True)

    # Assuming the topic column is named 'Topic'
    # Adjust as necessary if your column name is different
    combined_df.sort_values(by=['Topic'], inplace=True)

    # Write the combined dataframe to a new Excel file
    combined_df.to_excel(output_file, index=False)


# Example usage
if __name__ == '__main__':
    combine_excel_files('path_to_your_directory', 'combined_output.xlsx')
