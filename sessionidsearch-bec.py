import pandas as pd

def clean_path(path):
    return path.strip().strip('"').strip("'")

def search_session_id_excel(input_file, session_id, output_file="filtered_results.xlsx"):
    input_file = clean_path(input_file)
    output_file = clean_path(output_file)

    # Force .xlsx output
    if not output_file.lower().endswith(".xlsx"):
        output_file += ".xlsx"

    xls = pd.ExcelFile(input_file)
    sheets = xls.sheet_names

    print(f"[+] Sheets detected: {sheets}")

    matches_found = False

    # Create Excel writer
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

        for sheet in sheets:
            print(f"[+] Processing sheet: {sheet}")

            try:
                df = pd.read_excel(input_file, sheet_name=sheet, dtype=str)
            except Exception as e:
                print(f"[!] Skipping sheet {sheet}: {e}")
                continue

            df = df.fillna("").astype(str)

            # Search entire row
            mask = df.apply(
                lambda row: row.str.contains(session_id, case=False, na=False)
            ).any(axis=1)

            matches = df[mask]

            if not matches.empty:
                matches_found = True
                matches["__source_sheet"] = sheet

                # Write matches to same-named sheet
                matches.to_excel(writer, sheet_name=sheet, index=False)

                print(f"[+] {len(matches)} matches found in {sheet}")

    if matches_found:
        print(f"[+] Output saved to: {output_file}")
    else:
        print("[!] No matches found — no file created")


if __name__ == "__main__":
    input_file = input("Enter Excel file path (.xlsx): ")
    session_id = input("Enter session ID to search: ")
    output_file = input("Enter output filename (default filtered_results.xlsx): ")

    if not output_file.strip():
        output_file = "filtered_results.xlsx"

    search_session_id_excel(input_file, session_id, output_file)