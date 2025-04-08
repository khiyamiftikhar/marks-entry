import sys
import pandas as pd
import os
import keyboard

def save_excel(df, file_path):
    """ Save the DataFrame to the Excel file """
    try:
        temp_file = file_path + "_temp.xlsx"  # Workaround for OneDrive issues
        df.to_excel(temp_file, index=False, engine='openpyxl')
        os.replace(temp_file, file_path)  # Replace original file safely
        print(f"\n✅ Changes saved successfully to {file_path}")
    except Exception as e:
        print(f"\n❌ Error saving file: {e}")

def get_matching_rows(df, reg_number):
    """ Find rows that match the last three digits of the registration number """
    """ Find rows that match the last three digits of the registration number (case insensitive) """
    matches = df[df['Reg Number'].str[-3:].str.upper() == reg_number.upper()]
    return matches

def get_three_digit_input():
    """ Capture exactly 3-digit input without pressing Enter """
    reg_number = ""
    pressed_keys = set()  # Track keys to avoid multiple presses per stroke

    while len(reg_number) < 3:
        event = keyboard.read_event(suppress=True)
        
        if event.event_type == keyboard.KEY_DOWN and event.name.isdigit() and event.name not in pressed_keys:
            pressed_keys.add(event.name)  # Avoid multiple detection
            reg_number += event.name
            print(event.name, end="", flush=True)

        elif event.event_type == keyboard.KEY_UP and event.name in pressed_keys:
            pressed_keys.remove(event.name)  # Allow detection again after release
        
        elif event.name == 'backspace' and len(reg_number) > 0:
            reg_number = reg_number[:-1]
            print("\b \b", end="", flush=True)  # Erase character
        
        elif event.name == 'q':  # Allow quitting mid-entry
            print("\n\nSaving changes before exiting...")
            return 'q'
    
    print("")  # Move to new line
    return reg_number

def get_one_digit_mark():
    """ Capture a single digit without pressing Enter """
    pressed_keys = set()  # Track keys to avoid multiple presses per stroke

    while True:
        event = keyboard.read_event(suppress=True)

        if event.event_type == keyboard.KEY_DOWN and event.name.isdigit() and len(event.name) == 1 and event.name not in pressed_keys:
            pressed_keys.add(event.name)  # Avoid multiple detection
            print(event.name)
            return int(event.name)

        elif event.event_type == keyboard.KEY_UP and event.name in pressed_keys:
            pressed_keys.remove(event.name)  # Allow detection again after release

def main():
    if len(sys.argv) != 2:
        print("Usage: python marks.py <ExcelFileName.xlsx>")
        sys.exit(1)

    file_path = sys.argv[1]

    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        sys.exit(1)

    if 'Reg Number' not in df.columns or 'Marks' not in df.columns:
        print("❌ Excel file must contain 'Reg Number' and 'Marks' columns.")
        sys.exit(1)

    print("\nPress 'q' anytime to quit and save changes.")

    try:
        while True:
            print("\nEnter 3-digit registration number (or 'q' to quit): ", end='', flush=True)
            reg_number = get_three_digit_input()

            if reg_number == 'q':
                save_excel(df, file_path)
                break

            matches = get_matching_rows(df, reg_number)

            if len(matches) == 0:
                print("❌ No matching registration number found.")
                continue
            elif len(matches) > 1:
                print("⚠️ Multiple matches found! Enter full registration number.")
                full_reg = input("Enter full registration number: ").strip().upper()
                matches = df[df['Reg Number'].str.upper() == full_reg]

                if len(matches) != 1:
                    print("❌ No exact match found.")
                    continue

            reg_row = matches.index[0]
            reg_full = df.loc[reg_row, 'Reg Number']
            
            print(f"\nEnter marks for {reg_full} (1-digit only): ", end='', flush=True)
            marks = get_one_digit_mark()

            df.at[reg_row, 'Marks'] = marks
            print(f"✅ Marks updated for {reg_full}: {marks}")

    except KeyboardInterrupt:
        print("\n\n⏳ Exiting program. Saving changes...")
        save_excel(df, file_path)

if __name__ == "__main__":
    main()
