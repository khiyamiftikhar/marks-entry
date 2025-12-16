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
    """Return rows where registration number matches (partial match with last 3 digits)."""
    matches = df[df['Reg Number'].str[-3:].str.upper() == reg_number.upper()]
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

def get_total_questions():
    while True:
        print("\nEnter the total number of questions (or 0 for attendance mode): ", end="", flush=True)
        num_questions = get_marks_input()

        if num_questions == 0:
            return 0  # Attendance mode
        elif num_questions > 0:
            return num_questions
        else:
            print("❌ Invalid input. Number of questions must be greater than 0.")

def check_marks_entered(df, reg_row):
    """Check if marks are already entered for the given registration row."""
    return pd.notna(df.at[reg_row, 'Marks']) and df.at[reg_row, 'Marks'] != ""

def get_remaining_matches(df, reg_number, processed_reg_numbers):
    """Get remaining unmarked matches excluding previously processed entries."""
    matches = df[df['Reg Number'].str[-3:].str.upper() == reg_number.upper()]
    # Filter out entries that already have marks or have been processed
    unmarked_matches = matches[matches['Marks'].isna() | (matches['Marks'] == "")]
    unmarked_matches = unmarked_matches[~unmarked_matches['Reg Number'].isin(processed_reg_numbers)]
    return unmarked_matches

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

    # Get the total number of questions for the exam or check for attendance mode
    total_questions = get_total_questions()

    processed_reg_numbers = set()  # To track the processed registration numbers

    try:
        if total_questions == 0:  # Attendance mode
            print("\nYou are in Attendance Mode. Enter registration numbers to mark attendance.")
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

                # --- NEW: Check if already marked ---
                unmarked_matches = matches[matches['Marks'].isna() | (matches['Marks'] == "")]
                if len(unmarked_matches) == 0:
                    print("❌ Attendance already marked for this registration number.")
                    continue
                # -------------------------------------

                # If only one unmarked match found
                if len(unmarked_matches) == 1:
                    row_index = unmarked_matches.index[0]
                    df.at[row_index, 'Marks'] = 1
                    print(f"✅ Attendance marked for {unmarked_matches.iloc[0]['Reg Number']}.")
                    save_excel(df, file_path)
                    continue

                # Multiple unmarked matches
                print("⚠️ Multiple matching unmarked entries:")
                for idx, (_, row) in enumerate(unmarked_matches.iterrows(), 1):
                    print(f"{idx}    {row['Reg Number']}")

                print("Select the correct one by number: ", end='', flush=True)
                selected_idx = get_one_digit_choice(len(unmarked_matches))

                reg_row = unmarked_matches.index[selected_idx]
                df.at[reg_row, 'Marks'] = 1

                print(f"✅ Attendance marked for {unmarked_matches.iloc[selected_idx]['Reg Number']}.")
            save_excel(df, file_path)


        else:  # Regular mode for marks entry
            print("\nYou are in Marks Entry Mode.")
            while True:
                print("\nEnter 3-digit registration number (or 'q' to quit): ", end='', flush=True)
                reg_number = get_three_digit_input()

                if reg_number == 'q':
                    save_excel(df, file_path)
                    break

                # Get the remaining unmarked matches excluding already processed ones
                unmarked_matches = get_remaining_matches(df, reg_number, processed_reg_numbers)

                if len(unmarked_matches) == 0:
                    print("❌ No unmarked matching registration number found or all already marked.")
                    continue

                # If only one match is found, automatically proceed without asking
                if len(unmarked_matches) == 1:
                    reg_row = unmarked_matches.index[0]
                    print(f"\nYou are entering marks for {df.loc[reg_row, 'Reg Number']}.")
                    processed_reg_numbers.add(df.loc[reg_row, 'Reg Number'])

                # If more than one match is found, handle duplicates
                else:
                    print("⚠️ Multiple unmarked entries found:")
                    for idx, (_, row) in enumerate(unmarked_matches.iterrows(), 1):
                        print(f"{idx}    {row['Reg Number']}")
                    print("Select the correct one by number: ", end="", flush=True)
                    selected_idx = get_one_digit_choice(len(unmarked_matches))
                    reg_row = unmarked_matches.index[selected_idx]
                    processed_reg_numbers.add(df.loc[reg_row, 'Reg Number'])

                # Enter marks for each question
                for i in range(total_questions):
                    print(f"\nEnter marks for question {i + 1} (0-9 or A=10, B=11, ...): ", end='', flush=True)
                    marks = get_marks_input()
                    df.at[reg_row, f'Q{i + 1}_Marks'] = marks
                    print(f"✅ Marks updated for question {i + 1}: {marks}")

                # Save the current state after each marks entry
                save_excel(df, file_path)

    except KeyboardInterrupt:
        print("\n\n⏳ Exiting program. Saving changes...")
        save_excel(df, file_path)

if __name__ == "__main__":
    main()
