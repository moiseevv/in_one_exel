from pathlib import Path
import pandas as pd
import openpyxl
import os

files = os.listdir()

path = Path("E:\python\openpyxl")
df = pd.DataFrame()

for f in files:
    if "~" not in f:
        if ".xlsx" in f:
            print(f)
            try:
                df = pd.concat([df, pd.read_excel(f,engine='openpyxl',)])
                pass
            except:
                print(f" ОШИБКА!!! В файле  -   {f}   произошла ошибка. Возможно файл поврежден")
                continue



#df = pd.concat([pd.read_excel(f) for f in path.glob("*.xlsx")], ignore_index=True)

df.to_excel("final.xlsx")