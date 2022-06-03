import json
import argparse

import pandas as pd

if __name__ == "__main__":
    
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=str, help="Source .XLSX file")
    parser.add_argument("structure", type=str, help="Structure file destination")
    parser.add_argument("--sheet", "-s", dest="sheet", type=str, help="Name of the sheet")
    args = parser.parse_args()

    print("Loading file ..")
    df = pd.read_excel(args.source, sheet_name=args.sheet)

    print("Loading column names ..")
    if args.sheet is None:
        print(f"Using sheet '{list(df.keys())[0]}'")
        col_names = tuple(list(df.values())[0].columns)
    else:
        col_names = tuple(df.columns)

    cols = [ dict(name=col_name) for col_name in col_names ]

    structure = { "cols" : cols }

    with open(args.structure, "w") as structure_file:
        json.dump(structure, structure_file, indent=4)