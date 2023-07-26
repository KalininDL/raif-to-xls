from datetime import datetime
from typing import List, Dict

import pandas as pd
from xlsxwriter import Workbook
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

from util.colors import Colors
from aggregator import (
    Report,
    Income,
    FinOp,
    Expenses,
    Top5Payment,
    CurrencyOperation,
)
from cli import Settings
from util import Currency, try_format_float


class XslxWriter:
    sheet: Worksheet | None
    workbook: Workbook | None
    row_count: int
    settings: Settings

    def __init__(self, settings: Settings):
        self.settings = settings
        self.row_count = 0
        self.workbook: Workbook | None = None
        self.sheet: Worksheet | None = None

    def generate_xlsx(self, reports: List[Report]):
        for report in reports:
            from_date_printable: str = report.from_date.strftime("%d.%b.%Y")
            to_date_printable: str = report.to_date.strftime("%d.%b.%Y")

            file_name: str = (
                f"Report-{report.currency.name}-"
                f"{from_date_printable}-"
                f"{to_date_printable}.xlsx"
            )

            sheet_name: str = f"{from_date_printable}-{to_date_printable}"

            writer = pd.ExcelWriter(
                file_name,
                engine="xlsxwriter",
                datetime_format="dd.mm.yyyy",
            )
            self.workbook = writer.book
            self.workbook.add_worksheet(sheet_name)
            self.sheet: Worksheet = writer.sheets[sheet_name]
            self.row_count = 0

            self.add_section_header(
                self.row_count, 1, 15, height=30, title="General report"
            )

            printable_df = report.table.dataframe

            (max_row, max_col) = printable_df.shape

            self.sheet.add_table(
                1,
                0,
                max_row + 1,
                max_col,
                {"autofilter": True, "style": f"Table Style Medium 9"},
            )

            my_format = self.workbook.add_format()
            my_format.set_align("vcenter")
            self.workbook.add_format({"num_format": "$#,##0.00"})

            self.sheet.autofilter(1, 0, max_row + 1, max_col)

            # Convert the dataframe to an XlsxWriter Excel object.
            printable_df.to_excel(writer, sheet_name=sheet_name, startrow=1, startcol=0)

            for col_num, value in enumerate(printable_df.columns.values):
                self.sheet.write(
                    1, col_num + 1, value, self.workbook.add_format({"font_size": 12})
                )

            for row_num, row in printable_df.iterrows():
                if not pd.isna(row["Card number"]):
                    self.sheet.write(
                        row_num + 2,
                        3,
                        row["Card number"],
                        self.workbook.add_format({"align": "right"}),
                    )

            self.sheet.write(1, 0, "№")

            self.gap(len(printable_df) + 2)

            self.add_report_income_expense_stats(report)

            self.add_income_report(report.income)

            self.add_expenses_report(report.expenses)

            self.sheet.autofit()

            # Close the Pandas Excel writer and output the Excel file.
            writer.close()

    def add_section_header(
        self,
        row: int,
        col: int,
        width: int,
        title: str,
        bg_color: str = Colors.BG_HEADER,
        height: int = 20,
    ) -> None:
        header_format: Format = self.format("merge", bg_color=bg_color)
        self.sheet.merge_range(row, col, row, col + width - 1, title, header_format)
        self.sheet.set_row(row, height)
        self.gap()

    def add_report_income_expense_stats(self, report: Report):
        income: float = report.income.total
        expenses: float = report.expenses.total

        left: Format = self.workbook.add_format({"align": "right", "border": True})
        income_style: Format = self.workbook.add_format(
            {"align": "right", "bg_color": Colors.BG_INCOME, "border": True}
        )
        expense_style: Format = self.workbook.add_format(
            {"align": "right", "bg_color": Colors.BG_EXPENSE, "border": True}
        )

        self.sheet.write(self.row_count, 9, "Total income:", self.format("bold"))
        self.sheet.write(self.row_count, 10, try_format_float(income), income_style)
        self.gap()

        self.sheet.write(self.row_count, 9, "Total expenses:", self.format("bold"))
        self.sheet.write(self.row_count, 10, try_format_float(expenses), expense_style)
        self.gap()

        self.sheet.write(self.row_count, 9, "Left: ", self.format("bold"))
        self.sheet.write(self.row_count, 10, try_format_float(income - expenses), left)
        self.gap()

        self.sheet.write(self.row_count, 8, "You spent", self.format("bold"))
        self.sheet.write(
            self.row_count, 9, str(int((expenses / income) * 100)) + "%", left
        )
        self.sheet.write(self.row_count, 10, "of your income", self.format("bold"))
        self.gap()

    def add_income_report(self, income: Income):
        self.add_section_header(
            self.row_count, 1, 10, "Income statistics", Colors.BG_INCOME, height=25
        )
        self.add_section_header(self.row_count, 1, 3, "Incomes")

        self.add_fin_op_header(self.row_count, 1)
        salaries: List[FinOp] = income.salaries
        for salary in salaries:
            self.add_fin_op(self.row_count, 1, salary)

        meal_allowances: List[FinOp] = income.meal_allowances
        for allowance in meal_allowances:
            self.add_fin_op(self.row_count, 1, allowance)

        other: float = income.other
        self.sheet.write(self.row_count, 2, "Other:", self.format("bold"))
        self.sheet.write(
            self.row_count, 3, try_format_float(other), self.format("right")
        )
        self.gap()

        total: float = income.total
        self.sheet.write(self.row_count, 2, "Total:", self.format("bold"))
        self.sheet.write(
            self.row_count,
            3,
            try_format_float(total),
            self.workbook.add_format(
                {"bg_color": Colors.BG_INCOME, "align": "right", "bold": True}
            ),
        )
        self.gap(2)

    def add_expenses_report(self, expenses: Expenses):
        first_table_start_cell: int = 1
        second_table_start_cell: int = 5
        self.add_section_header(
            self.row_count, 1, 10, "Expenses statistics", Colors.BG_EXPENSE, height=25
        )

        headers_row: int = self.row_count
        self.add_section_header(
            headers_row, first_table_start_cell, 3, "Top-5 biggest purchases"
        )

        self.add_top_5_purchases(expenses, first_table_start_cell)

        self.row_count = headers_row

        start_cell: int = 5
        self.add_top_5_places(
            expenses, headers_row, second_table_start_cell, start_cell
        )

        second_tables_row: int = self.row_count
        self.add_cache_withdraws(expenses, first_table_start_cell)

        self.row_count = second_tables_row
        self.add_currency_operations(expenses, second_table_start_cell)

        self.gap(2)

    def add_top_5_places(
        self, expenses, headers_row, second_table_start_cell, start_cell
    ):
        self.add_section_header(
            headers_row, second_table_start_cell, 4, "Top-5 expenses"
        )
        top_5: List[Top5Payment] = expenses.top_5_places
        self.sheet.write(
            self.row_count, second_table_start_cell, "Description", self.format("label")
        )
        self.sheet.write(
            self.row_count, second_table_start_cell + 1, "Amount", self.format("label")
        )
        self.sheet.write(
            self.row_count, second_table_start_cell + 2, "Times", self.format("label")
        )
        self.sheet.write(
            self.row_count,
            second_table_start_cell + 3,
            "Average bill",
            self.format("label"),
        )
        self.gap()
        for top_5_item in top_5:
            amount: float = top_5_item.amount
            title: str = top_5_item.title
            avg: float = top_5_item.avg_bill
            num: int = top_5_item.num_of_occurrences

            self.sheet.write(self.row_count, start_cell, title, self.format("right"))
            self.sheet.write(
                self.row_count,
                start_cell + 1,
                try_format_float(amount),
                self.format("right"),
            )
            self.sheet.write(self.row_count, start_cell + 2, num, self.format("right"))
            self.sheet.write(
                self.row_count,
                start_cell + 3,
                try_format_float(avg),
                self.format("right"),
            )
            self.gap()
        self.gap()

    def add_top_5_purchases(self, expenses, first_table_start_cell):
        top_5_purchases: List[FinOp] = expenses.top_5_purchases
        self.add_fin_op_header(self.row_count, first_table_start_cell)
        for expense in top_5_purchases:
            self.add_fin_op(self.row_count, first_table_start_cell, expense)

    def add_currency_operations(self, expenses, second_table_start_cell):
        currency_operations: List[CurrencyOperation] = expenses.currency_operations
        self.add_section_header(
            self.row_count, second_table_start_cell, 6, "Currency operations"
        )
        self.sheet.write(
            self.row_count, second_table_start_cell, "Date", self.format("label")
        )
        self.sheet.write(
            self.row_count,
            second_table_start_cell + 1,
            "Description",
            self.format("label"),
        )
        self.sheet.write(
            self.row_count, second_table_start_cell + 2, "Amount", self.format("label")
        )
        self.sheet.write(
            self.row_count,
            second_table_start_cell + 3,
            "Purchased",
            self.format("label"),
        )
        self.sheet.write(
            self.row_count,
            second_table_start_cell + 4,
            "Currency",
            self.format("label"),
        )
        self.sheet.write(
            self.row_count,
            second_table_start_cell + 5,
            "Exchange rate",
            self.format("label"),
        )
        self.gap()
        for curr_op in currency_operations:
            self.add_currency_op(self.row_count, second_table_start_cell, curr_op)

    def add_cache_withdraws(self, expenses, first_table_start_cell):
        cache_withdraws: List[FinOp] = expenses.cache_withdraws
        self.add_section_header(
            self.row_count, first_table_start_cell, 3, "Cache withdraws"
        )
        self.add_fin_op_header(self.row_count, first_table_start_cell)
        for withdraw in cache_withdraws:
            self.add_fin_op(self.row_count, first_table_start_cell, withdraw)

    def add_fin_op(
        self,
        row: int,
        start_cell: int,
        fin_op: FinOp,
    ):
        date: datetime = fin_op.date
        amount: float = fin_op.amount
        title: str = fin_op.title

        self.sheet.write(
            row, start_cell, date.strftime("%d.%m.%Y"), self.format("right")
        )
        self.sheet.write(row, start_cell + 1, title, self.format("right"))
        self.sheet.write(
            row,
            start_cell + 2,
            try_format_float(amount),
            self.workbook.add_format({"align": "right"}),
        )
        self.gap()

    def add_fin_op_header(
        self, row: int, start_cell: int, bg_color: str = Colors.BG_LABELS
    ):
        self.sheet.write(row, start_cell, "Date", self.format("label", bg_color))
        self.sheet.write(
            row, start_cell + 1, "Description", self.format("label", bg_color)
        )
        self.sheet.write(row, start_cell + 2, "Amount", self.format("label", bg_color))
        self.gap()

    def add_currency_op(self, row: int, start_cell: int, curr_op: CurrencyOperation):
        bought_amount: float = curr_op.bought[0]
        currency: Currency = curr_op.bought[1]
        fin_op: FinOp = curr_op.spent
        exchange_rate: float = curr_op.exchange_rate
        self.add_fin_op(row, start_cell, fin_op)
        start_cell += 3
        self.sheet.write(row, start_cell, bought_amount, self.format("right"))
        self.sheet.write(row, start_cell + 1, currency.value)
        self.sheet.write(
            row,
            start_cell + 2,
            try_format_float(exchange_rate),
            self.workbook.add_format({"align": "left"}),
        )

    def gap(self, rows: int = 1):
        self.row_count += rows

    def format(self, name: str, bg_color: str = None) -> Format:
        formats: Dict[str, Format] = {
            "label": self.workbook.add_format(
                {
                    "bold": True,
                    "border": True,
                    "font_size": 12,
                    "bg_color": Colors.BG_LABELS if bg_color is None else bg_color,
                }
            ),
            "merge": self.workbook.add_format(
                {
                    "bold": True,
                    "border": True,
                    "align": "center",
                    "valign": "vcenter",
                    "font_size": 14,
                    "fg_color": Colors.BG_HEADER if bg_color is None else bg_color,
                }
            ),
            "bold": self.workbook.add_format({"bold": True}),
            "right": self.workbook.add_format({"align": "right"}),
        }
        return formats[name]
