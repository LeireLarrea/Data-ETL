import pathlib
from collections import ChainMap
from typing import List


# you can call this utility by:
# def transform_function_xlsx_to_csv(source_s3_key: str, f_dest_dir: str) -> Sequence[str]:
#         """Receives a path to an excel file, and returns a list of paths where the csvs for each tab are created"""
#         utils_excel = Utils_Excel()
#         return [str(x) for x in utils_excel.to_csv(Path(source_s3_key), Path(f_dest_dir))]



default_sheet_args = {
    "header_offset": None,
    # "start_index": 1,
    # "end_index": 1
    # Sheet1: {
    # "header_offset": None,
    # }
}


class Utils_Excel:
    """A utility class for methods to deal with xlsx files"""

    def __init__(self, delimiter=",", newline="\r\n", quotechar='"'):
        self.delimiter = delimiter
        self.newline = newline
        self.quotechar = quotechar

    def to_csv(self, input_path: pathlib, output_dir: pathlib, sheet_args={}) -> List[str]:
        """
        Creates a CSV file for each sheet in a xlsx file
        :param input_path: path where xlsx file to transform is located
        :param output_dir: path to directory where to save the converted tabs into CSVs
        :param sheet_args: optional dictionary of names to rename each CSV
        :return: Returns a list of the paths where the CSVs are saved
        """
        import pandas as pd

        xls = pd.ExcelFile(str(input_path))
        results = []
        for sheet_name in xls.sheet_names:

            # Config setup
            if sheet_name in sheet_args:
                per_sheet_args = sheet_args[sheet_name]
            else:
                per_sheet_args = sheet_args

            args = ChainMap(per_sheet_args, default_sheet_args)

            setup_headers = args["header_offset"] is not None

            # Process file
            df = pd.read_excel(input_path, sheet_name=sheet_name, header=args["header_offset"])

            if df.size == 0 or "skip" in args:
                print(f"Ignoring empty sheet {sheet_name}")
                continue

            if "skip" in args:
                print(f"Skipping sheet {sheet_name}")
                continue

            start = int(per_sheet_args.get("start_index", 1)) - 1
            end = int(per_sheet_args.get("end_index", df[df.columns[0]].count())) - 1

            subset_df = df.loc[start:end]
            sheet_name = str(sheet_name).lower().strip().replace(' ', '_').replace('+', '_and_' )
            new_csv_path = (output_dir / sheet_name).with_suffix(".csv")
            subset_df.to_csv(new_csv_path, sep=self.delimiter, lineterminator=self.newline, quotechar=self.quotechar, header=setup_headers, index=False)
            results.append(new_csv_path)

        return results
