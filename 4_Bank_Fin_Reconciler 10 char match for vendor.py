# Bank-Fin Reconciliation with GUI
import pandas as pd
from datetime import datetime
from tkinter import Tk, filedialog, StringVar, Label, Entry, Button, messagebox
import tkinter as tk
from itertools import combinations
import os

# === Utility ===
def safe_excel_to_date(val):
    try:
        return (pd.to_datetime('1899-12-30') + pd.to_timedelta(float(val), unit='D')).strftime('%Y-%m-%d')
    except:
        try:
            return pd.to_datetime(val).strftime('%Y-%m-%d')
        except:
            return ''

def normalize_and_reconcile(bank_file, finance_file, log_callback=None, log_save_button=None):
    try:
        raw_bank_df = pd.read_excel(bank_file)
        raw_finance_df = pd.read_excel(finance_file)
        bank_df = raw_bank_df.copy()
        finance_df = raw_finance_df.copy()

        bank_df['norm_date'] = bank_df['Date'].apply(safe_excel_to_date)
        bank_df['norm_amount'] = bank_df['Withdrawal (Dr.)'].apply(lambda x: round(float(x), 2) if pd.notna(x) else 0.0)
        bank_df['norm_vendor'] = bank_df['der_bank_ven'].apply(lambda x: str(x).strip().upper()[:10] if pd.notna(x) else '')

        finance_df['norm_date'] = finance_df['Payment Date'].apply(safe_excel_to_date)
        finance_df['norm_amount'] = finance_df['Credit Amount'].apply(lambda x: round(float(x), 2) if pd.notna(x) else 0.0)
        finance_df['norm_vendor'] = finance_df['der_fin_ven'].apply(lambda x: str(x).strip().upper()[:10] if pd.notna(x) else '')

        matched_rows = []
        unmatched_bank = []
        unmatched_finance_indexes = set(finance_df.index)
        match_id_counter = 1
        voucher_exists = 'Voucher No' in finance_df.columns

        start_time = datetime.now()
        log_lines = [f"Bank-Fin Reconciliation Started: {start_time.strftime('%Y-%m-%d %H:%M:%S')}"]

        for i, b_row in bank_df.iterrows():
            if log_callback:
                log_callback(f"Processing Bank UID: {b_row['bank_uid']}")
            match_found = False

            # 1-to-1
            for j, f_row in finance_df.iterrows():
                if j not in unmatched_finance_indexes:
                    continue
                if (
                    b_row['norm_date'] == f_row['norm_date'] and
                    b_row['norm_amount'] == f_row['norm_amount'] and
                    b_row['norm_vendor'] == f_row['norm_vendor']
                ):
                    match_id = f"M{match_id_counter:04}"
                    matched_rows.append({
                        'Match ID': match_id, 'Type': '1-to-1', 'Role': 'Bank', 'UID': b_row['bank_uid'],
                        'Date': b_row['Date'], 'Vendor': b_row['der_bank_ven'], 'Amount': b_row['Withdrawal (Dr.)'], 'Receiver Name': '',
                        'der_bank_ven': b_row['der_bank_ven'], 'der_fin_ven': f_row['der_fin_ven']
                    })
                    matched_rows.append({
                        'Match ID': match_id, 'Type': '1-to-1', 'Role': 'Finance', 'UID': f_row['fin_uid'],
                        'Date': f_row['Payment Date'], 'Vendor': f_row['der_fin_ven'], 'Amount': f_row['Credit Amount'],
                        'Receiver Name': f_row.get('Receiver Name', ''), 'Voucher No': f_row.get('Voucher No', '') if voucher_exists else '',
                        'der_bank_ven': b_row['der_bank_ven'], 'der_fin_ven': f_row['der_fin_ven']
                    })
                    unmatched_finance_indexes.remove(j)
                    match_found = True
                    match_id_counter += 1
                    break
            if match_found:
                continue

            # 1-to-2
            candidate_finance = finance_df[(finance_df['norm_date'] == b_row['norm_date']) & (finance_df['norm_vendor'] == b_row['norm_vendor'])]
            for j1, j2 in combinations(candidate_finance.index.intersection(unmatched_finance_indexes), 2):
                f1, f2 = finance_df.loc[j1], finance_df.loc[j2]
                if abs(b_row['norm_amount'] - (f1['norm_amount'] + f2['norm_amount'])) < 0.01:
                    match_id = f"M{match_id_counter:04}"
                    matched_rows.append({
                        'Match ID': match_id, 'Type': '1-to-2', 'Role': 'Bank', 'UID': b_row['bank_uid'],
                        'Date': b_row['Date'], 'Vendor': b_row['der_bank_ven'], 'Amount': b_row['Withdrawal (Dr.)'], 'Receiver Name': '',
                        'der_bank_ven': b_row['der_bank_ven'], 'der_fin_ven': ''
                    })
                    for f in [f1, f2]:
                        matched_rows.append({
                            'Match ID': match_id, 'Type': '1-to-2', 'Role': 'Finance', 'UID': f['fin_uid'],
                            'Date': f['Payment Date'], 'Vendor': f['der_fin_ven'], 'Amount': f['Credit Amount'],
                            'Receiver Name': f.get('Receiver Name', ''), 'Voucher No': f.get('Voucher No', '') if voucher_exists else '',
                            'der_bank_ven': b_row['der_bank_ven'], 'der_fin_ven': f['der_fin_ven']
                        })
                    unmatched_finance_indexes -= {j1, j2}
                    match_found = True
                    match_id_counter += 1
                    break
            if match_found:
                continue

            # 1-to-N up to 10
            for n in range(3, 11):
                match_found_this_round = False
                for combo in combinations(candidate_finance.index.intersection(unmatched_finance_indexes), n):
                    total = sum(finance_df.loc[j]['norm_amount'] for j in combo)
                    if abs(b_row['norm_amount'] - total) < 0.01:
                        match_id = f"M{match_id_counter:04}"
                        matched_rows.append({
                            'Match ID': match_id, 'Type': f'1-to-{n}', 'Role': 'Bank', 'UID': b_row['bank_uid'],
                            'Date': b_row['Date'], 'Vendor': b_row['der_bank_ven'], 'Amount': b_row['Withdrawal (Dr.)'], 'Receiver Name': '',
                            'der_bank_ven': b_row['der_bank_ven'], 'der_fin_ven': ''
                        })
                        for j in combo:
                            f = finance_df.loc[j]
                            matched_rows.append({
                                'Match ID': match_id, 'Type': f'1-to-{n}', 'Role': 'Finance', 'UID': f['fin_uid'],
                                'Date': f['Payment Date'], 'Vendor': f['der_fin_ven'], 'Amount': f['Credit Amount'],
                                'Receiver Name': f.get('Receiver Name', ''), 'Voucher No': f.get('Voucher No', '') if voucher_exists else '',
                                'der_bank_ven': b_row['der_bank_ven'], 'der_fin_ven': f['der_fin_ven']
                            })
                        unmatched_finance_indexes -= set(combo)
                        match_found = match_found_this_round = True
                        match_id_counter += 1
                        break
                if match_found:
                    break
                if n == 9 and not match_found_this_round:
                    break

            if not match_found:
                unmatched_bank.append({**b_row.to_dict(), 'der_bank_fin_match': False})

        unmatched_finance = finance_df.loc[list(unmatched_finance_indexes)].copy()
        unmatched_finance['der_bank_fin_match'] = False
        if voucher_exists and 'Voucher No' not in unmatched_finance.columns:
            unmatched_finance['Voucher No'] = ''

        norm_cols = ['norm_date', 'norm_amount', 'norm_vendor']
        df_matched = pd.DataFrame(matched_rows)
        df_unmatched_bank = pd.DataFrame(unmatched_bank).drop(columns=norm_cols, errors='ignore')
        df_unmatched_finance = unmatched_finance.drop(columns=norm_cols, errors='ignore')

        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Reconciled File As",
            initialfile=f"Reconciled_Bank_Fin_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not output_file:
            messagebox.showinfo("Cancelled", "💡 Save operation cancelled.")
            return

        with pd.ExcelWriter(output_file) as writer:
            df_matched.to_excel(writer, sheet_name="Matched", index=False)
            df_unmatched_bank.to_excel(writer, sheet_name="Unmatched_Bank", index=False)
            df_unmatched_finance.to_excel(writer, sheet_name="Unmatched_Finance", index=False)
            raw_bank_df.to_excel(writer, sheet_name="Normalized_Bank_Master", index=False)
            raw_finance_df.to_excel(writer, sheet_name="Normalized_Fin_Master", index=False)

        end_time = datetime.now()
        log_lines.append(f"Bank-Fin Reconciliation Completed: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        log_lines.append(f"Matched Rows: {len(df_matched)}")
        log_lines.append(f"Unmatched Bank Rows: {len(df_unmatched_bank)}")
        log_lines.append(f"Unmatched Finance Rows: {len(df_unmatched_finance)}")
        log_text = "\n".join(log_lines)
        if log_callback:
            log_callback("\n" + log_text)
        if log_save_button:
            log_save_button.config(state="normal")

        messagebox.showinfo("Done", f"✅ Reconciliation Complete. File saved as:\n{output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"❌ {str(e)}")

# === GUI ===
def browse_file(var):
    path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        var.set(path)

def run_gui():
    def update_log(message):
        log_textbox.insert('end', message + "\n")
        log_textbox.see('end')

    def save_log():
        log_content = log_textbox.get("1.0", "end").strip()
        if not log_content:
            messagebox.showinfo("Empty", "Log is empty.")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")], title="Save Log As")
        if filepath:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(log_content)
            messagebox.showinfo("Saved", f"Log saved to: {filepath}")

    def start_reconciliation():
        status_label.config(text="🔄 Processing...", fg="blue")
        root.update_idletasks()
        normalize_and_reconcile(bank_path.get(), finance_path.get(), log_callback=update_log, log_save_button=save_log_button)
        status_label.config(text="✅ Done.", fg="green")

    root = Tk()
    root.title("Bank-Fin Reconciler")
    root.geometry("800x400")
    root.resizable(False, False)

    bank_path = StringVar()
    finance_path = StringVar()

    Label(root, text="Normalized Bank File:").grid(row=0, column=0, sticky='w', padx=10, pady=10)
    Entry(root, textvariable=bank_path, width=60).grid(row=0, column=1)
    Button(root, text="Browse", command=lambda: browse_file(bank_path)).grid(row=0, column=2, padx=5)

    Label(root, text="Normalized Finance File:").grid(row=1, column=0, sticky='w', padx=10, pady=10)
    Entry(root, textvariable=finance_path, width=60).grid(row=1, column=1)
    Button(root, text="Browse", command=lambda: browse_file(finance_path)).grid(row=1, column=2, padx=5)

    Button(root, text="Run Reconciliation", command=start_reconciliation,
           bg="#4CAF50", fg="white", padx=20, pady=5).grid(row=3, column=1, pady=20)

    status_label = Label(root, text="", fg="blue")
    status_label.grid(row=4, column=1, pady=5)

    log_textbox = tk.Text(root, height=10, width=90, wrap="word", borderwidth=1, relief="solid")
    log_textbox.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

    save_log_button = Button(root, text="Save Log as .txt", command=save_log, state="disabled")
    save_log_button.grid(row=6, column=1, pady=5)

    root.mainloop()

run_gui()
