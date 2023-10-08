import pandas as pd
from tkinter import filedialog

file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

if not file_path:
    print("No file selected. Exiting.")
else:
    df = pd.read_excel(file_path)

    files = ["french.ts", "arabic.ts"]

    for firstIndex, file in enumerate(files):
        result = {}
        parent_object = {}
        mainkey = None

        for index, row in df.iterrows():
            if index == 0:
                continue
            current_mainkey = row.iloc[0]
            key = row.iloc[1]
            valueFr = row.iloc[3]
            valueAr = row.iloc[4]

            if isinstance(current_mainkey, str):
                mainkey = current_mainkey
                parent_object = result.setdefault(mainkey, {})

            if isinstance(key, str):
                if firstIndex == 0:
                    parent_object[key] = valueFr
                else:
                    parent_object[key] = valueAr

        with open(file, "w", encoding="utf-8") as newFile:
            newFile.write("export const data = {\n")
            for mainkey, inner_object in result.items():
                newFile.write(f"  {mainkey}: {{\n")
                for key, value in inner_object.items():
                    newFile.write(f'    {key}: "{value}",\n')
                newFile.write("  },\n")
            newFile.write("};")

        print(f"TypeScript object has been written to {file}")
