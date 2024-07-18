import pandas as pd
import glob


def combine_excel_files(directory_path, output_file):
    # List to hold dataframes
    dataframes = []

    # Read all xlsx files in the directory
    for file in glob.glob(f"{directory_path}/*.xlsx"):
        df = pd.read_excel(file)
        dataframes.append(df)

    # Concatenate all dataframes, using outer join to ensure all topics are included
    combined_df = pd.concat(dataframes, ignore_index=True, sort=False)

    # Assuming the topic column is named 'Topic'
    # Fill NaN values with an appropriate placeholder if needed
    combined_df.fillna('', inplace=True)

    # Write the combined dataframe to a new Excel file
    combined_df.to_excel(output_file, index=False)


# Example usage
if __name__ == '__main__':
    combine_excel_files('path_to_your_directory', 'combined_output.xlsx')
