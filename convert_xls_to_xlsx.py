import os
import pandas as pd

input_file = "clientes.xls"
output_file = "clientes.xlsx"

# Detectar si es ODS
ext = os.path.splitext(input_file)[-1].lower()

if ext in [".ods", ".xls"]:
    df = pd.read_excel(input_file, engine="odf")
else:
    df = pd.read_excel(input_file)

df.to_excel(output_file, index=False)
print(f"✅ Archivo convertido: {output_file}")
