
import sys
import pandas as pd
import os
import keyboard

def save_excel(df, file_path):
    try:
        temp_file = file_path + "_temp.xlsx"
        df.to_excel(temp_file, index=False, engine='openpyxl')
        os.replace(temp_file, file_path)
        print(f"\n✅ Changes saved successfully to {file_path}")
    except Exception as e:
        print(f"\n❌ Error saving file: {e}")

def get_matching_rows(df, reg_number):
    matches = df[df['Reg Number'].str[-3:].str.upper() == reg_number.upper()]
    matches = matches[df['Marks'].isna() | (df['Marks'] == "")]
    return matches

def get_three_digit_input():
    reg_number = ""
    pressed_keys = set()

    while len(reg_number) < 3:
        event = keyboard.read_event(suppress=True)

        if event.event_type == keyboard.KEY_DOWN and event.name.isdigit() and event.name not in pressed_keys:
            pressed_keys.add(event.name)
            reg_number += event.name
            print(event.name, end="", flush=True)

        elif event.event_type == keyboard.KEY_UP and event.name in pressed_keys:
            pressed_keys.remove(event.name)

        elif event.name == 'backspace' and len(reg_number) > 0:
            reg_number = reg_number[:-1]
            print("\b \b", end="", flush=True)

        elif event.name == 'q':
            print("\n\nSaving changes before exiting...")
            return 'q'

    print("")
    return reg_number

def get_one_digit_choice(valid_range):
    pressed_keys = set()
    while True:
        event = keyboard.read_event(suppress=True)
        if event.event_type == keyboard.KEY_DOWN and event.name.isdigit() and event.name not in pressed_keys:
            index = int(event.name)
            if 1 <= index <= valid_range:
                print(index)
                return index - 1
        elif event.event_type == keyboard.KEY_UP and event.name in pressed_keys:
            pressed_keys.remove(event.name)

def get_marks_input():
    pressed_keys = set()
    while True:
        event = keyboard.read_event(suppress=True)
        if event.event_type == keyboard.KEY_DOWN and event.name not in pressed_keys:
            pressed_keys.add(event.name)
            char = event.name.upper()
            if char.isdigit():
                print(char)
                return int(char)
            elif char.isalpha() and len(char) == 1:
                value = ord(char) - ord('A') + 10
                print(char)
                return value
        elif event.event_type == keyboard.KEY_UP and event.name in pressed_keys:
            pressed_keys.remove(event.name)

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
                print("❌ No matching registration number found or all already marked.")
                continue
            elif len(matches) == 1:
                reg_row = matches.index[0]
            else:
                print("⚠️ Multiple unmarked entries found:")
                for idx, (_, row) in enumerate(matches.iterrows(), 1):
                    print(f"{idx}    {row['Reg Number']}")
                print("Select the correct one by number: ", end="", flush=True)
                selected_idx = get_one_digit_choice(len(matches))
                reg_row = matches.index[selected_idx]

            reg_full = df.loc[reg_row, 'Reg Number']
            print(f"\nEnter marks for {reg_full} (0-9 or A=10, B=11, ...): ", end='', flush=True)
            marks = get_marks_input()

            df.at[reg_row, 'Marks'] = marks
            print(f"✅ Marks updated for {reg_full}: {marks}")

    except KeyboardInterrupt:
        print("\n\n⏳ Exiting program. Saving changes...")
        save_excel(df, file_path)

if __name__ == "__main__":
    main()
