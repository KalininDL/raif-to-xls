from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor, wait
from multiprocessing import cpu_count

from tabula import read_pdf
import pandas as pd
from typing import List, Dict, Tuple, BinaryIO, Any
from pandas import DataFrame
import PyPDF2
from termcolor import colored
import re

from cli import Settings
from util import Currency, to_datetime


@dataclass
class Table:
    dataframe: DataFrame
    currency: Currency


column_names_list: List[str] = [
    "Transaction date",
    "Completion date",
    "Card number",
    "Transaction description",
    "Amount in foreign currency",
    "Amount in original currency",
    "Expense",
    "Income",
    "Balance",
]

areas_eur_usd: Dict[str, List[float]] = {
    "first_page": [366.818, 10.71, 701.123, 594.405],
    "second_and_other": [53.933, 19.0, 693.473, 588.16],
    "last_page": [14.918, 12.24, 636.098, 595.17],
}

areas_rsd: Dict[str, List[float]] = {
    "first_page": [366.818, 10.71, 701.123, 594.405],
    "second_and_other": [53.933, 19.0, 693.473, 588.16],
    "last_page": [56.993, 11.475, 702.653, 596.7],
}


def read_pdf_on_threaded_pool(
    file_name: str, num_of_pages: int, areas: dict[str, list[float]]
) -> List[DataFrame]:
    executor = ThreadPoolExecutor(max_workers=cpu_count())
    read_pdf_tasks: List[Tuple[Dict[str, Any], int]] = [
        (
            {
                "input_path": file_name,
                "pages": 1,
                "area": areas["first_page"],
                "stream": True,
                "pandas_options": {"columns": column_names_list},
            },
            1,
        )
    ]

    if num_of_pages >= 3:
        for page in range(2, num_of_pages):
            read_pdf_tasks.append(
                (
                    {
                        "input_path": file_name,
                        "pages": page,
                        "area": areas["second_and_other"],
                        "stream": True,
                        "pandas_options": {"columns": column_names_list},
                    },
                    page,
                )
            )

    read_pdf_tasks.append(
        (
            {
                "input_path": file_name,
                "pages": num_of_pages,
                "stream": True,
                "pandas_options": {"columns": column_names_list},
            },
            num_of_pages,
        )
    )

    futures = [
        executor.submit(read_page_async, args[1], args[0]) for args in read_pdf_tasks
    ]
    wait(futures)
    results = [f.result() for f in futures]

    results.sort(key=lambda elem: elem[1])
    dataframes_per_file = [result[0][0] for result in results]

    for idx, dataframe in enumerate(dataframes_per_file):
        if dataframe["Expense"][0] == "Isplata":
            dataframes_per_file[idx] = dataframe.tail(-2).reset_index(drop=True)

    return dataframes_per_file


def read_page_async(page_num, kwargs) -> Tuple[List[DataFrame], int]:
    return read_pdf(**kwargs), page_num


def read_pdf_with_log(paths: List[str]) -> List[Tuple[Currency, List[DataFrame]]]:
    dataframes: List[Tuple[Currency, List[DataFrame]]] = []
    for path in paths:
        try:
            file: BinaryIO = open(path, "rb")
            pdf = PyPDF2.PdfReader(file)
            num_of_pages = len(pdf.pages)
            text_of_first_page: str = pdf.pages[0].extract_text()
            currency_str: str = re.findall(
                "Strana:.*$", text_of_first_page, re.MULTILINE
            )[0].split(" ")[1]
            currency: Currency = Currency(currency_str)
            areas = areas_rsd if currency == Currency.RSD else areas_eur_usd
            dataframes_per_file: List[DataFrame] = read_pdf_on_threaded_pool(
                file.name, num_of_pages, areas
            )

            dataframes.append((currency, dataframes_per_file))
        except Exception as e:
            print(
                colored(
                    f"Failed to extract data from file {path}\n{e}",
                    "light_red",
                    force_color=True,
                )
            )

    return dataframes


def preprocess_tables(
    dataframes_per_file: List[Tuple[Currency, List[DataFrame]]], merge: bool
) -> List[Table]:
    all_tables: List[Table] = []
    for dataframes_with_currency in dataframes_per_file:
        currency: Currency = dataframes_with_currency[0]
        dataframes: List[DataFrame] = dataframes_with_currency[1]

        # Concat all the tables into one
        table = pd.concat(dataframes, axis=0, ignore_index=True)

        # Extract exchange rate into separate column

        table.insert(6, "Exchange rate", "")

        for idx, row in table.loc[
            table["Amount in original currency"].str.contains("(?<=Kurs: ).*", na=False)
        ].iterrows():
            table.loc[idx - 1, "Exchange rate"] = row[
                "Amount in original currency"
            ].split(" ")[1]

        # Drop Nan rows
        table.dropna(subset=["Balance"], inplace=True)
        table.reset_index(drop=True, inplace=True)

        # Sometimes "Trsansaction date" may be empty,
        # use "Completion date" as a fallback to not mess with nan rows further
        for idx, row in table.iterrows():
            if pd.isna(row["Transaction date"]):
                table.at[idx, "Transaction date"] = row["Completion date"]
            try:
                table.at[idx, "Amount in foreign currency"] = float(
                    row["Amount in foreign currency"]
                )
            except ValueError:
                pass

        # Settings data types
        table["Expense"] = table["Expense"].replace(",", "", regex=True).astype(float)
        table["Income"] = table["Income"].replace(",", "", regex=True).astype(float)
        table["Balance"] = table["Balance"].replace(",", "", regex=True).astype(float)
        table["Amount in original currency"] = table[
            "Amount in original currency"
        ].replace(",", "", regex=True)
        table["Completion date"] = table["Completion date"].apply(
            lambda date: to_datetime(date)
        )
        table["Transaction date"] = table["Transaction date"].apply(
            lambda date: to_datetime(date)
        )
        table.fillna("No information", inplace=True)

        all_tables.append(Table(table, currency))

    if merge:
        rsd_reports: List[DataFrame] = list(
            filter(
                lambda table: table.currency is Currency.RSD,
                all_tables,
            )
        )
        eur_reports: List[DataFrame] = list(
            filter(
                lambda table: table.currency is Currency.EUR,
                all_tables,
            )
        )
        usd_reports: List[DataFrame] = list(
            filter(
                lambda table: table.currency is Currency.USD,
                all_tables,
            )
        )

        merged_tables: List[Table] = []

        if rsd_reports:
            merged_tables.append(
                Table(
                    pd.concat(map(lambda t: t.dataframe, rsd_reports), axis=0),
                    Currency.RSD,
                )
            )
        if eur_reports:
            merged_tables.append(
                Table(
                    pd.concat(map(lambda t: t.dataframe, eur_reports), axis=0),
                    Currency.RSD,
                )
            )
        if usd_reports:
            merged_tables.append(
                Table(
                    pd.concat(map(lambda t: t.dataframe, usd_reports), axis=0),
                    Currency.RSD,
                )
            )

        for table in merged_tables:
            table.dataframe.reset_index(drop=True, inplace=True)

        return merged_tables

    return all_tables


class PDFReader:
    def __init__(self, settings: Settings):
        self.settings = settings

    def extract_data_from_pdfs(self) -> List[Table]:
        paths: List[str] = self.settings.files
        merge: bool = self.settings.merge

        all_datasets: List[Tuple[Currency, List[DataFrame]]] = read_pdf_with_log(paths)
        all_tables: List[Table] = preprocess_tables(all_datasets, merge)

        return all_tables

    pass
