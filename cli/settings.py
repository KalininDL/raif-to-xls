from dataclasses import dataclass
from typing import List


@dataclass
class Settings:
    merge: bool
    files: List[str]
    output: str
    single_file: bool
    yulia_mode: bool
