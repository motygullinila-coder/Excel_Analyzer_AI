from analyzer_file import MetaAnalyzerFile

path_file = "C:\\Users\\ilyam\\Desktop\\Пр-во+экспрессы ноябр-дек21г.xlsx"

meta_analyzer = MetaAnalyzerFile.MetaAnalyzerFile(path_file)

def main():
    print(meta_analyzer.get_workbook_info())

    print(meta_analyzer.get_profile())

    print(meta_analyzer.get_dataframe_info())

    print(meta_analyzer.get_sheet_preview())

if __name__ == '__main__':
    main()