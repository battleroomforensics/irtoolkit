import csv
import json
import sys
import pandas as pd

def flatten_json(json_obj, prefix=''):
    flat_dict = {}
    for key, value in json_obj.items():
        if isinstance(value, dict):
            flat_dict.update(flatten_json(value, prefix + key + '.'))
        elif isinstance(value, list):
            for i, item in enumerate(value):
                if isinstance(item, dict):
                    flat_dict.update(flatten_json(item, prefix + key + f'[{i}].'))
                else:
                    flat_dict[prefix + key + f'[{i}]'] = item
        else:
            flat_dict[prefix + key] = value
    return flat_dict

def parse_audit_log(csv_file_path, output_excel_path, session_ids=None):
    rows_by_operation = {}
    session_ids_set = set(session_ids) if session_ids else None

    with open(csv_file_path, mode='r', encoding='utf-8') as csv_file:
        csv_reader = csv.DictReader(csv_file)

        for row in csv_reader:
            try:
                audit_data = json.loads(row['AuditData'])
                flat_audit_data = flatten_json(audit_data)
                row.update(flat_audit_data)

                # Reformat CreationDate
                try:
                    original_date = row['CreationDate']
                    trimmed_date = original_date.split('.')[0].replace('T', ' ')
                    row['CreationDate'] = trimmed_date
                except Exception as e:
                    print(f"Error parsing CreationDate in row {row.get('RecordId', 'unknown')}: {e}")

                # Check if we need to filter by session ID
                if session_ids_set:
                    # Look for SessionId in various possible locations
                    session_id = None
                    for key in row:
                        if 'SessionId' in key:
                            session_id = row[key]
                            break
                    
                    # Skip this row if its SessionId doesn't match any in our list
                    if not session_id or session_id not in session_ids_set:
                        continue

                operation = row.get('Operation', 'Unknown')
                if operation not in rows_by_operation:
                    rows_by_operation[operation] = []
                rows_by_operation[operation].append(row)

            except json.JSONDecodeError as e:
                print(f"Error decoding JSON in row {row.get('RecordId', 'unknown')}: {e}")

    # Write to Excel with separate sheets
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        if not rows_by_operation:
            # Create an empty sheet if no matching records found
            pd.DataFrame().to_excel(writer, sheet_name='No Matching Records', index=False)
            print("No records found matching the specified session IDs.")
        else:
            for operation, rows in rows_by_operation.items():
                df = pd.DataFrame(rows)
                # Excel sheet names can't be longer than 31 characters
                safe_sheet_name = operation[:31]
                df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            
            print(f"Successfully processed {sum(len(rows) for rows in rows_by_operation.values())} matching records.")

def get_session_ids():
    """Prompt the user for session IDs and return as a list."""
    print("\nEnter session ID(s) to filter by (comma separated) or leave blank to process all records:")
    session_input = input("> ").strip()
    
    if not session_input:
        return None
    
    # Split by comma and strip whitespace
    session_ids = [sid.strip() for sid in session_input.split(',')]
    return session_ids

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python ualparser.py <input_csv_path> <output_excel_path>")
        sys.exit(1)

    csv_file_path = sys.argv[1]
    output_excel_path = sys.argv[2]
    
    # Get session IDs from user
    session_ids = get_session_ids()
    
    if session_ids:
        print(f"Filtering for session IDs: {', '.join(session_ids)}")
    else:
        print("Processing all records (no session ID filter)")
    
    parse_audit_log(csv_file_path, output_excel_path, session_ids)