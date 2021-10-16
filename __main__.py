import os
from excel_handler.excel_splitter import split_excel


file_name = os.path.abspath('report_creator/input/MarketSight Crosstab - Wade_questions.xlsx')


def main():
    split_excel(file_name)


if __name__ == "__main__":
    main()
