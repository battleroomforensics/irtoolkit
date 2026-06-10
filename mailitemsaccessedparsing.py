import pandas as pd
import re

def clean_path(path):
    return path.strip().strip('"').strip("'")

def normalize_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def unique_join(values):
    cleaned = []
    seen = set()

    for v in values:
        v = normalize_value(v)
        if v and v not in seen:
            seen.add(v)
            cleaned.append(v)

    return " | ".join(cleaned)

def extract_message_data(input_file, output_file="mailitems_messageids_unique.xlsx"):
    input_file = clean_path(input_file)
    output_file = clean_path(output_file)

    if not output_file.lower().endswith(".xlsx"):
        output_file += ".xlsx"

    msgid_pattern = re.compile(
        r"^Folders\[(\d+)\]\.FolderItems\[(\d+)\]\.InternetMessageId$",
        re.IGNORECASE
    )

    xls = pd.ExcelFile(input_file, engine="openpyxl")
    sheet_names = xls.sheet_names

    all_rows = []

    print(f"[+] Sheets detected: {sheet_names}")

    for sheet_name in sheet_names:
        print(f"[+] Processing sheet: {sheet_name}")

        try:
            df = pd.read_excel(input_file, sheet_name=sheet_name, dtype=str, engine="openpyxl")
        except Exception as e:
            print(f"[!] Skipping sheet {sheet_name}: {e}")
            continue

        if df.empty:
            print(f"[!] Sheet {sheet_name} is empty, skipping")
            continue

        df.columns = [str(c) for c in df.columns]

        # CreationDate column
        creation_col = None
        for col in df.columns:
            if col.strip().lower() == "creationdate":
                creation_col = col
                break

        # MailboxOwnerUPN column
        mailbox_col = None
        for col in df.columns:
            if col.strip().lower() == "mailboxownerupn":
                mailbox_col = col
                break

        # InternetMessageId columns
        msgid_cols = []
        for col in df.columns:
            m = msgid_pattern.match(col)
            if m:
                folder_idx = m.group(1)
                item_idx = m.group(2)
                msgid_cols.append((col, folder_idx, item_idx))

        if not msgid_cols:
            print(f"[!] No InternetMessageId columns found in sheet: {sheet_name}")
            continue

        for row_idx, row in df.iterrows():
            creation_date = normalize_value(row[creation_col]) if creation_col else ""
            mailbox_owner = normalize_value(row[mailbox_col]) if mailbox_col else ""

            for msgid_col, folder_idx, item_idx in msgid_cols:
                message_id = normalize_value(row.get(msgid_col, ""))

                if not message_id:
                    continue

                subject_col = f"Folders[{folder_idx}].FolderItems[{item_idx}].Subject"
                path_col = f"Folders[{folder_idx}].Path"

                subject = normalize_value(row.get(subject_col, ""))
                path = normalize_value(row.get(path_col, ""))

                all_rows.append({
                    "MessageId": message_id,
                    "MailboxOwnerUPN": mailbox_owner,
                    "Subject": subject,
                    "Path": path,
                    "CreationDate": creation_date,
                    "SourceSheet": sheet_name,
                    "SourceRow": row_idx + 2
                })

    if not all_rows:
        print("[!] No message data found. No output file created.")
        return

    all_df = pd.DataFrame(all_rows)

    # Aggregate by MessageId + MailboxOwnerUPN
    unique_df = (
        all_df.groupby(["MessageId", "MailboxOwnerUPN"], dropna=False)
        .agg({
            "Subject": lambda x: unique_join(x),
            "Path": lambda x: unique_join(x),
            "CreationDate": lambda x: unique_join(x),
            "SourceSheet": lambda x: unique_join(x),
            "SourceRow": lambda x: unique_join(x)
        })
        .reset_index()
    )

    # Write output
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        unique_df.to_excel(writer, sheet_name="UniqueMessageIds", index=False)
        all_df.to_excel(writer, sheet_name="AllOccurrences", index=False)

        # Per-mailbox sheets
        if "MailboxOwnerUPN" in all_df.columns:
            for mailbox, mailbox_df in all_df.groupby("MailboxOwnerUPN"):
                safe_name = "UnknownMailbox" if not mailbox else str(mailbox)
                safe_name = safe_name[:31]  # Excel limit
                mailbox_df.to_excel(writer, sheet_name=safe_name, index=False)

    print(f"[+] Done")
    print(f"[+] Unique MessageId+Mailbox combos: {len(unique_df)}")
    print(f"[+] Total occurrences extracted: {len(all_df)}")
    print(f"[+] Output written to: {output_file}")


if __name__ == "__main__":
    input_file = input("Enter Excel file path (.xlsx): ").strip()
    output_file = input("Enter output filename (default mailitems_messageids_unique.xlsx): ").strip()

    if not output_file:
        output_file = "mailitems_messageids_unique.xlsx"

    extract_message_data(input_file, output_file)
