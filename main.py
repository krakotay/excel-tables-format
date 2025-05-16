import polars as pl
from excel import create_title_sheet
import datetime

DOC = "1746105004_СчетаООО  ВНЕШИНТОРГ.xlsx"


def main(filename: str):
    df = pl.read_excel(filename, infer_schema_length=0)
    # print(df)
    accounts = []
    for row in df.iter_rows(named=True):
        accounts.append(
            {
                "bank_name": row["КО: Наименование"],
                "bank_details": f"РегНом/НомФ: {row['Счет: Регномер']} ИНН/КПП: {row['КО: ИНН']}/{row['КО: КПП']} БИК(СВИФТ): {row['КО: БИК']}",
                "address": row["КО: Адрес"],
                "account_number": row["Счет: Номер"],
                "open_date": row["Счет: Дата открытия"].removesuffix(" 00:00:00")
                if row["Счет: Дата открытия"]
                else None,
                "status": row["Счет: Состояние"],
                "account_type": row["Счет: Вид"] or row["Счет: Вид (5.12)"],
                "close_date": row["Счет: Дата закрытия"].removesuffix(" 00:00:00")
                if row["Счет: Дата закрытия"]
                else None,
            }
        )
    datetime_now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=3)))
    formation_date = datetime_now.strftime("%d.%m.%Y")
    full_name = df["Наименование правообладателя"].first()
    inn = df["НП: ИНН"].first()
    create_title_sheet(
        output_path=DOC.replace(".xlsx", "_формат.xlsx"),
        form_code="Форма по КНД 1120499",
        form_number="Форма 67ф",
        report_title=(
            "Сведения о банковских счетах (вкладах, электронных средствах "
            "платежа (ЭСП))\nфизического лица, не являющегося индивидуальным предпринимателем"
        ),
        as_of_date=formation_date,
        formation_date=formation_date,
        number_number=" " * 10,
        full_name=full_name,
        inn=inn,
        dob=None,
        pob=None,
        accounts=accounts,
    )


if __name__ == "__main__":
    main(DOC)
