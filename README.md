# raif-to-xls

This tool allows you to convert your Raiffeisen Bank reports into easy-to-work-with xlsx sheets.
This tool supports RSD, USD and EUR reports.

### How to use?
1. Clone this project using git cli tool:
 * `git clone https://github.com/KalininDL/raif-to-xls.git`
2. Install Python > v.3.11 if it is not yet installed
 * You can check your python version using `python3 --version` command in your terminal
 * If it is yet not installed, please visit https://www.python.org/downloads/
3. Install required packages:
 * Go to the directory where this project have been installed
 * Execute `pip install -r requirements.txt`
4. After the required packages are installed, the program is ready to use! Basic usage:
 * Basic usage: `python3 main.py -f *path-to-your-pdf*`
5. See the CLI options description below:

### CLI options:
Firstly, you can always get help by using `python3 main.py -h`

```
options:
  -h, --help            show this help message and exit
  -f FILES [FILES ...], --files FILES [FILES ...]
                        Path to your Raiffeisen bank payslip PDF file
  -m, --merge           Merge all tables into one report
  -o OUTPUT_DIR, --output-dir OUTPUT_DIR
                        Output directory
  -s, --single-file     Merge all reports into one .xslx file with multiple sheets

```

**-f** flag: Allows you to specifies the path to your .pdf reports you want to convert.  
Ways to list the files:
* `-f ./report1.pdf ./report2.pdf ./other_dir/report3.pdf`
* `-f ./reports_dir/*` <- all .pdf will be used

**-m** flag: Merge multiple reports into one table:
For example, you have May, June and July reports. Without this flag, 3 separate monthly tables will be generated.  
Using **-m** flag you can merge them into one huge table with all the states merged

**-o** flag: Allows you to specify output directory for the generated reports. 
* Example: `python3 main.py -f ./report1.pdf ./report2.pdf -o ./output` <- all the generated .xlsx will be written into `./output` directory

**-s** flag: Write multiple reports into the separate sheets of single .xlsx file instead of generating multiple .xlsx file for each .pdf