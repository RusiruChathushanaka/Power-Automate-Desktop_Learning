import pandas as pd
import os

def update_excel_increment(filename="increment.xlsx", column_name="Value"):
    """
    Opens an Excel file (or creates it if it doesn't exist),
    reads the last number in a specified column, adds 1 to it,
    and writes this new number in the next available row.
    The script interacts with the file's data; it doesn't "open"
    the file in an application like Excel for viewing during execution.
    """
    next_value = 1  # Default starting value if the column is empty or file is new

    if os.path.exists(filename):
        try:
            # Attempt to read the Excel file
            # Using engine='openpyxl' is recommended for .xlsx files
            df = pd.read_excel(filename, engine='openpyxl')

            if not df.empty and column_name in df.columns:
                # Convert the specified column to numeric, coercing errors to NaN
                # This handles cases where the column might have non-numeric data or empty cells
                numeric_series = pd.to_numeric(df[column_name], errors='coerce')

                # Filter out NaN values to get only valid numbers
                valid_numbers = numeric_series.dropna()

                if not valid_numbers.empty:
                    last_value = valid_numbers.max()
                    next_value = int(last_value) + 1
                # If valid_numbers is empty (e.g., column exists but has no numeric data),
                # next_value remains 1 (our default)
            # If DataFrame is empty or the specified column doesn't exist,
            # next_value remains 1

        except pd.errors.EmptyDataError:
            # This occurs if the file exists but is completely empty
            df = pd.DataFrame(columns=[column_name])
            print(f"'{filename}' exists but is empty. Starting with value 1 in column '{column_name}'.")
        except ValueError as ve:
            # This can happen if openpyxl has trouble parsing the file (e.g., corrupted)
            print(f"Error reading '{filename}': {ve}. It might be corrupted or not a valid Excel file. A new DataFrame will be used.")
            df = pd.DataFrame(columns=[column_name])
        except Exception as e:
            # Catch other potential errors during file reading
            print(f"An unexpected error occurred while reading '{filename}': {e}. A new DataFrame will be used.")
            df = pd.DataFrame(columns=[column_name])
    else:
        # File does not exist, so we'll create a new DataFrame
        print(f"'{filename}' not found. Creating a new file.")
        df = pd.DataFrame(columns=[column_name])

    # Ensure df is a pandas DataFrame, even if errors occurred or file was new
    if not isinstance(df, pd.DataFrame):
        df = pd.DataFrame(columns=[column_name]) # Fallback

    # If the target column doesn't exist in the DataFrame (e.g., after creating a new df
    # or if df was loaded from an empty file without this column), add it.
    if column_name not in df.columns:
        df[column_name] = pd.Series(dtype='object') # Initialize with object type, will be coerced to numeric later

    # Prepare the new row to be added
    # The new row is a DataFrame itself, which makes concatenation straightforward
    new_row = pd.DataFrame({column_name: [next_value]})

    # Concatenate the existing DataFrame with the new row
    # ignore_index=True ensures that the DataFrame index is reset nicely
    df = pd.concat([df, new_row], ignore_index=True)

    try:
        # Before saving, ensure the target column is treated as numeric.
        # This is important if the column was just added, was empty, or contained mixed data.
        df[column_name] = pd.to_numeric(df[column_name], errors='coerce')

        # Save the updated DataFrame back to the Excel file
        # index=False prevents pandas from writing the DataFrame index as a column in Excel
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"Successfully updated '{filename}'. Value '{next_value}' added to column '{column_name}'.")
        print(f"To view the contents, please open '{filename}' using Microsoft Excel or a similar program.")

    except PermissionError:
        print(f"Error saving data to '{filename}': Permission denied. The file might be open in another program.")
    except Exception as e:
        print(f"An error occurred while saving data to '{filename}': {e}")

if __name__ == "__main__":
    # You can change the filename or column name here if needed
    update_excel_increment(filename="increment.xlsx", column_name="Value")

    # To test, you can run the script multiple times:
    # update_excel_increment(filename="increment.xlsx", column_name="Value")
    # update_excel_increment(filename="increment.xlsx", column_name="Value")