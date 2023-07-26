from datetime import datetime
from enum import Enum

import pandas as pd
from termcolor import colored


class Currency(Enum):
    RSD = "RSD"
    EUR = "EUR"
    USD = "USD"


def to_datetime(date: str) -> datetime | str:
    date_format = "%d.%m.%Y"
    return datetime.strptime(date, date_format) if not pd.isna(date) else ""


def print_colored(message: str, color: str):
    print(colored(message, color=color, force_color=True))


def try_format_float(number: float) -> str | float:
    try:
        return f"{number:,.2f}"
    except ValueError:
        return number
