from dataclasses import dataclass

from datetime import datetime
from typing import List, Tuple

from reader import Table
from util import Currency


@dataclass
class FinOp:
    title: str
    date: datetime | None
    amount: float


@dataclass
class Income:
    total: float
    salaries: List[FinOp]
    meal_allowances: List[FinOp]
    other: float


@dataclass
class CacheWithdraw(FinOp):
    currency: Currency


@dataclass
class CurrencyOperation:
    bought: Tuple[float, Currency]
    spent: FinOp
    exchange_rate: float


@dataclass
class Top5Payment(FinOp):
    num_of_occurrences: int
    avg_bill: float


@dataclass
class Expenses:
    total: float
    top_5_places: List[Top5Payment]
    top_5_purchases: List[FinOp]
    cache_withdraws: List[FinOp]
    currency_operations: List[CurrencyOperation]


@dataclass
class Report:
    table: Table
    income: Income
    expenses: Expenses
    currency: Currency
    from_date: datetime
    to_date: datetime
