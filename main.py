from openpyxl import load_workbook
import polars as pl
DOC = '1746105004_СчетаООО  ВНЕШИНТОРГ.xlsx'
def main(filename: str):
    df = pl.read_excel(filename, infer_schema_length=0)
    print(df)
    name = df['Наименование правообладателя']
    

def formar_excel(filename: str):
    wb = load_workbook(filename, read_only=True)
    ws = wb.worksheets[0]

    print("Hello from excel-tables-format!")


if __name__ == "__main__":
    main(DOC)
