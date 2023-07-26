import os
from argparse import ArgumentParser, Namespace
from typing import List
from termcolor import colored
from .settings import Settings


class CLI:
    def __init__(self):
        arg_parser: ArgumentParser = ArgumentParser()
        arg_parser.add_argument(
            "-f",
            "--files",
            nargs="+",
            required=True,
            help="Path to your Raiffeisen bank payslip PDF file",
        )
        arg_parser.add_argument(
            "-m",
            "--merge",
            help="Merge all tables into one report",
            default=False,
            action="store_true",
        )
        arg_parser.add_argument(
            "-o", "--output-dir", help="Output directory", default="./"
        )
        arg_parser.add_argument(
            "-s",
            "--single-file",
            help="Merge all reports into one .xslx file with multiple sheets",
            default=False,
            action="store_true",
        )
        arg_parser.add_argument("-u", default=False, action="store_true")
        self.args_parser = arg_parser

    def get_settings(self) -> Settings:
        args: Namespace = self.args_parser.parse_args()
        files: List[str] = args.files
        files_to_process: List[str] = []

        for file in files:
            if not (os.path.exists(file) and os.path.isfile(file)):
                print(
                    colored(f"It seems like file {file} does not exist!", "light_red")
                )
                skip = input("Want to skip this file and continue? (y/n): ")
                if skip.lower() == "y":
                    continue
                else:
                    exit(1)
            else:
                files_to_process.append(file)

        return Settings(
            args.merge, files_to_process, args.output_dir, args.single_file, args.u
        )
