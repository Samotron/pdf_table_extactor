from tabula import read_pdf
import pandas as pd
from tkinter import filedialog
import os

class pdf_table_reader():

    file_link = ""

    def __init__(self, export_format):
        self.export = export_format


    def get_file_name(self):
        self.file_link = filedialog.askopenfilename()
        if self.file_link.endswith('.pdf'):
            return self.file_link
        else:
            raise ValueError('PDF not supplied')

    def get_output_folder(self):
        return os.path.dirname(self.file_link)

    def pdf_table_to_df(self):
        df = read_pdf(self.get_file_name(), multiple_tables=True, pages=2)
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
    a.df_to_output()