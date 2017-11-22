from tabula import read_pdf
import pandas as pd
from tkinter import filedialog
import os
import webbrowser

class pdf_table_reader():

    file_link = ""

    def __init__(self, export_format):
        self.export = export_format


    def check_tabular_java_installed(self):
        try:
            read_pdf('dependencies/ast_sci_data_tables_sample.pdf')
            print('All working okay')
        except FileNotFoundError:
            print('JAVA not correctly installed')
            webbrowser.open(r'https://github.com/chezou/tabula-py')


    def get_file_name(self):
        self.file_link = filedialog.askopenfilename()
        if self.file_link.endswith('.pdf'):
            return self.file_link
        else:
            raise ValueError('PDF not supplied')

    def get_output_folder(self):
        return os.path.dirname(self.file_link)

    def pdf_table_to_df(self):
        df = read_pdf(self.get_file_name(), multiple_tables=True, pages='all')
        print(type(df))
        return df

    def df_to_output(self):
        df = self.pdf_table_to_df()

        if self.export == 'excel':
            writer = pd.ExcelWriter(os.path.join(self.get_output_folder(), 'PDF_Output.xlsx'))
            j = 1
            for i in df:
                i.to_excel(writer, 'Page {}'.format(j))
                j += 1
                print(j)

            writer.save()

        elif self.export == 'csv':
            df.to_csv(os.path.join(self.get_output_folder(), 'CSV_Output.csv'), sep='   ')

        elif self.export == 'json':
            df.to_json(os.path.join(self.get_output_folder(), 'JSON_Output.json'))

        else:
            print('Incorrect export format specified')


if __name__ == "__main__":
    a = pdf_table_reader(export_format='excel')
    a.check_tabular_java_installed()
    a.df_to_output()