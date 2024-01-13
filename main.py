import pandas as pd


class DataTransformer:
    def __init__(self, file_path, start_date, end_date):
        self.df = pd.read_excel(file_path)
        self.start_date = start_date
        self.end_date = end_date
        self.action_priority = ["Revoke", "Softblock", "Warn"]

    def clean_data(self):
        self.df["Action"] = self.df["Action"].replace(
            {
                "Soft1": "Softblock",
                "Soft2": "Softblock",
                "Warn1": "Warn",
                "Warn2": "Warn",
            }
        )

    def pivot_data(self):
        pivot_table = self.df.pivot_table(
            index="Merchant Customer ID",
            columns=["Metrics", "Action"],
            aggfunc="size",
            fill_value=0,
        )
        pivot_table = pivot_table.applymap(lambda x: x if x != 0 else pd.NA)
        pivot_table["Grand Total"] = pivot_table.sum(axis=1).replace({0: pd.NA})
        return pivot_table

    def generate_stats(self, row):
        actions = (
            row.loc[:, self.action_priority].replace({0: pd.NA}).dropna().index.tolist()
        )
        actions.sort(key=lambda x: self.action_priority.index(x[1]))
        stats = "SFP_Enforcement {} {} ".format(self.start_date, self.end_date)
        if actions:
            highest_action = actions[0][1]
            action_list = [
                f"{action[1]} {'GV_' if action[0] in ['OS', 'STD', 'XL'] else ''}{action[0]}"
                for action in actions
            ]
            stats += f"{highest_action} [{', '.join(action_list)}]"
        return highest_action, stats

    def create_stats_sheets(self, pivot_table):
        sheets_data = {
            action: {"Merchant Customer ID": [], "Stats": []}
            for action in self.action_priority
        }

        for index, row in pivot_table.iterrows():
            highest_action, stats = self.generate_stats(row)
            sheets_data[highest_action]["Merchant Customer ID"].append(index)
            sheets_data[highest_action]["Stats"].append(stats)

        stats_sheets = {
            action: pd.DataFrame(data) for action, data in sheets_data.items()
        }
        return stats_sheets

    def save_to_excel(self, pivot_table, stats_sheets, output_file_path):
        with pd.ExcelWriter(output_file_path) as writer:
            pivot_table.to_excel(
                writer, sheet_name="Pivot Table", index=True, na_rep=""
            )
            for sheet_name, data in stats_sheets.items():
                data.to_excel(writer, sheet_name=sheet_name, index=False)


input_file_path = "/content/list.xlsx"
output_file_path = "/content/pivoted_list.xlsx"

start_date = "2023-12-31"
end_date = "2024-01-06"

transformer = DataTransformer(input_file_path, start_date, end_date)
transformer.clean_data()

pivot_table = transformer.pivot_data()
stats_table = transformer.create_stats_sheets(pivot_table)

transformer.save_to_excel(pivot_table, stats_table, output_file_path)
