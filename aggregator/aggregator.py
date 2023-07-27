from pandas import DataFrame, Series
from aggregator.reportdataclasses import (
    Income,
    FinOp,
    CurrencyOperation,
    Top5Payment,
    Report,
    Expenses,
)
from util import Currency
from datetime import datetime
from typing import List, Tuple
from reader import Table
import tqdm


class Aggregator:
    def generate_reports(self, tables: List[Table]) -> List[Report]:
        progress = tqdm.tqdm(
            total=len(tables), colour="green", desc="Gathering statistics: "
        )
        reports: List[Report] = []

        for table in tables:
            progress.update(1)
            income: Income = self.get_income(table)
            outcome: Expenses = self.get_outcome(table)
            from_date, to_date = self.get_period(table)
            reports.append(
                Report(table, income, outcome, table.currency, from_date, to_date)
            )

        progress.set_description("Gathering statistics complete!")
        progress.close()
        return reports

    def get_outcome(self, table: Table) -> Expenses:
        df: DataFrame = table.dataframe
        total_outcome = df["Expense"].sum()

        top_5_expenses: List[Top5Payment] = self.get_top5_item_stat(df)

        top_5_biggest_purchases: List[FinOp] = [
            FinOp(
                expense["Transaction description"],
                expense["Transaction date"],
                expense["Expense"],
            )
            for index, expense in df.drop(
                df[df["Transaction description"].str.contains(" ATM ", na=False)].index
            )
            .drop(
                df[df["Transaction description"].str.startswith("EB ", na=False)].index
            )
            .nlargest(5, ["Expense"])
            .iterrows()
        ]

        cash_withdraws: List[FinOp] = [
            FinOp("Cash withdraw", withdraw["Transaction date"], withdraw["Expense"])
            for index, withdraw in df[
                df["Transaction description"].str.contains(" ATM ")
            ].iterrows()
        ]

        currency_operations: List[CurrencyOperation] = self.get_currency_operations(df)

        outcome: Expenses = Expenses(
            total_outcome,
            top_5_expenses,
            top_5_biggest_purchases,
            cash_withdraws,
            currency_operations,
        )

        return outcome

    def get_income(self, table: Table) -> Income:
        df: DataFrame = table.dataframe
        income_rows: DataFrame = df.loc[df["Income"] > 0.0]
        salary_rows: DataFrame = income_rows[
            income_rows["Transaction description"].str.match("ZARADA", na=False)
        ]
        meal_allowance_rows: DataFrame = income_rows[
            income_rows["Transaction description"].str.match("Prevoz", na=False)
        ]

        salaries: List[FinOp] = [
            FinOp("Salary", salary["Transaction date"], salary["Income"])
            for idx, salary in salary_rows.iterrows()
        ]

        meal_allowances: List[FinOp] = [
            FinOp("Meal allowance", allowance["Transaction date"], allowance["Income"])
            for idx, allowance in meal_allowance_rows.iterrows()
        ]

        other_incomes: float = income_rows.drop(
            salary_rows.index.append(meal_allowance_rows.index)
        )["Income"].sum()
        total_income: float = df["Income"].sum()

        income: Income = Income(total_income, salaries, meal_allowances, other_incomes)

        return income

    def get_currency_operations(self, df: DataFrame) -> List[CurrencyOperation]:
        currency_operations: List[CurrencyOperation] = []

        for index, operation in df[
            df["Transaction description"].str.startswith("EB ", na=False)
        ].iterrows():
            currency: Currency = Currency(
                operation["Amount in original currency"][-3::]
            )
            currency_amount: float = float(
                operation["Amount in original currency"][0:-4]
            )
            date: datetime = operation["Transaction date"]
            amount: float = operation["Expense"]
            exchange_rate: float = float(operation["Exchange rate"])
            title: str = "Currency operation"
            currency_operations.append(
                CurrencyOperation(
                    (currency_amount, currency),
                    FinOp(title, date, amount),
                    exchange_rate,
                )
            )

        return currency_operations

    def get_period(self, table: Table) -> Tuple[datetime, datetime]:
        df = table.dataframe
        from_date = df["Transaction date"].min()
        to_date = df["Transaction date"].max()

        return from_date, to_date

    def get_top5_item_stat(self, df: DataFrame) -> List[Top5Payment]:
        top_5_payments: List[Top5Payment] = []
        top_5_rows: Series = df["Transaction description"].value_counts().nlargest(5)

        for expense in top_5_rows.index:
            title: str = expense
            num_of_payments: int = top_5_rows[expense]
            payments: Series = df.loc[df["Transaction description"] == expense][
                "Expense"
            ]
            sum: float = payments.sum()
            avg: float = payments.mean()
            top_5_payments.append(Top5Payment(title, None, sum, num_of_payments, avg))

        return top_5_payments
