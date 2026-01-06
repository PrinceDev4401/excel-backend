import pandas as pd
import re

def clean_phone_number(phone):
    if pd.isna(phone):
        return ""
    phone = str(phone)
    phone = re.sub(r"[()\s-]", "", phone)
    if phone.startswith("+233"):
        phone = "0" + phone[4:]
    if len(phone) == 9 and not phone.startswith("0"):
        phone = "0" + phone
    return phone


def process_excel(input_path, output_path):
    df = pd.read_excel(input_path, sheet_name="Orders")
    df.columns = df.columns.str.strip()

    # Auto-detect total column
    subtotal_column = None
    for col in df.columns:
        if "total" in col.lower():
            subtotal_column = col
            break

    if not subtotal_column:
        raise Exception("No total/subtotal column found.")

    df["Total Clean"] = pd.to_numeric(df[subtotal_column], errors="coerce").round()

    df["Full Name"] = (
        df["First Name (Billing)"].astype(str).str.strip()
        + " "
        + df["Last Name (Billing)"].astype(str).str.strip()
    ).str.title()

    df["Phone Clean"] = df["Phone (Billing)"].apply(clean_phone_number)

    output_rows = []

    for _, row in df.iterrows():
        if pd.isna(row["Total Clean"]):
            continue

        repeats = int(row["Total Clean"]) // 38
        if repeats <= 0:
            continue

        for _ in range(repeats):
            output_rows.append([
                row["Full Name"],
                row["Phone Clean"],
                row["Email (Billing)"]
            ])

    result_df = pd.DataFrame(output_rows, columns=["Full Name", "Phone", "Email"])
    result_df.to_excel(output_path, index=False)
