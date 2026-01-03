import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd


def show_message_to_user(message, message_type="info"):
    root = tk.Tk()
    root.withdraw()

    if message_type == "info":
        messagebox.showinfo("پیام", message)
    elif message_type == "warning":
        messagebox.showwarning("هشدار", message)
    elif message_type == "error":
        messagebox.showerror("خطا", message)

    root.destroy()


class ExcelMerger:

    def get_folder_path(self):
        """گرفتن مسیر پوشه از کاربر"""

        class FolderSelector:
            def __init__(self):
                self.root = tk.Tk()
                self.root.title("select folder")
                self.root.geometry("400x300")

            def get_folder_path(self):
                """با استفاده از dialog پوشه انتخاب میکند"""
                self.root.withdraw()

                user_folder_path = filedialog.askdirectory(
                    title="پوشه مورد نظر را انتخاب کنید",
                    initialdir=os.path.expanduser("~")  # شروع از پوشه کاربر
                )

                self.root.destroy()

                if user_folder_path and os.path.exists(user_folder_path):
                    return user_folder_path
                else:
                    return None

        # استفاده:
        selector = FolderSelector()
        path = selector.get_folder_path()
        print(f"مسیر انتخاب شده: {path}")
        return path

    @staticmethod
    def is_folder_valid(folder_loc):    # folder_loc = folder_loc
        try:
            folder_files = os.listdir(folder_loc)
            if not folder_files:
                show_message_to_user("⚠️ پوشه خالی است!")
                return False
            return True
        except FileNotFoundError:
            show_message_to_user("❌ پوشه پیدا نشد!")
            return False
        except PermissionError:
            show_message_to_user("🔒 دسترسی به پوشه مجاز نیست!")
            return False

    @staticmethod
    def get_excel_files(folder_loc): # folder_loc = folder_loc
        excel_files = [
            f for f in os.listdir(folder_loc)
            if f.lower().endswith(".xlsx")
        ]
        return excel_files

    @staticmethod
    def read_and_merge_excels(excel_files, folder_path):
        dfs = []

        for file in excel_files:
            file_path = os.path.join(folder_path, file)
            df = pd.read_excel(file_path)
            dfs.append(df)

        if not dfs:
            return pd.DataFrame()  # دیتافریم خالی

        merged_df = pd.concat(dfs, ignore_index=True)
        return merged_df

    @staticmethod
    def clean_merged_data(merged_df):
        if merged_df is None or merged_df.empty:
            print("❌ DataFrame خالی است!")
            return None

        merged_df = merged_df.copy()

        print("🧹 در حال تمیزکاری داده‌ها...")

        rows_before, cols_before = merged_df.shape
        print(f"   قبل: {rows_before} ردیف × {cols_before} ستون")

        merged_df.dropna(how="all", inplace=True)
        rows_removed = rows_before - len(merged_df)

        merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]
        cols_removed = cols_before - len(merged_df.columns)

        merged_df.dropna(axis=1, how="all", inplace=True)

        merged_df.reset_index(drop=True, inplace=True)

        print(f"   بعد: {merged_df.shape[0]} ردیف × {merged_df.shape[1]} ستون")
        print(f"   📊 حذف شد: {rows_removed} ردیف، {cols_removed} ستون")

        return merged_df

    @staticmethod
    def save_final_outputs(final_merged_df, output_folder="output", base_filename="clean_merged"):
        """
        ذخیره DataFrame نهایی به فرمت Excel و CSV
        """
        if final_merged_df is None or final_merged_df.empty:
            print("❌ دیتافریم نهایی خالی است! فایل ذخیره نشد.")
            return

        # ساخت پوشه خروجی در صورت عدم وجود
        os.makedirs(output_folder, exist_ok=True)

        excel_path = os.path.join(output_folder, f"{base_filename}.xlsx")
        csv_path = os.path.join(output_folder, f"{base_filename}.csv")

        final_merged_df.to_excel(excel_path, index=False)
        final_merged_df.to_csv(csv_path, index=False, encoding="utf-8-sig")

        print("💾 فایل‌ها با موفقیت ذخیره شدند:")
        print(f"   📘 Excel: {excel_path}")
        print(f"   📄 CSV:   {csv_path}")


if __name__ == "__main__":

    excel_merger = ExcelMerger()
    folder_path = excel_merger.get_folder_path()

    if excel_merger.is_folder_valid(folder_path):
        files = excel_merger.get_excel_files(folder_path)
        if files:
            print(files)

            merged_df = excel_merger.read_and_merge_excels(files, folder_path)

            final_merged_df = excel_merger.clean_merged_data(merged_df)
            excel_merger.save_final_outputs(final_merged_df)


