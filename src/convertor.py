import unicodedata
import pandas as pd


class Convertor:
    def __init__(self):
        self.input_path = "data/student_infomation_list.xlsx"
        self.output_path = "data/student_email_list.txt"
        
        self.__convert()
        
    @staticmethod
    def remove_accents(input_str):
        nfkd_form = unicodedata.normalize('NFKD', input_str)
        return "".join([c for c in nfkd_form if not unicodedata.combining(c)])
        
    def __convert(self):
        df = pd.read_excel(self.input_path, header=None)
        emails = []
        
        for _, row in df.iterrows():
            mssv = str(row[0]).strip() 
            ho = str(row[1]).strip()
            dem = str(row[2]).strip()
            ten = str(row[3]).strip()
            
            _ten = self.remove_accents(ten)

            _ho = self.remove_accents(ho)[0].upper()
            _dem = "".join([self.remove_accents(word)[0].upper() for word in dem.split()])
            _mssv = mssv[-7:]

            email = f"{_ten.capitalize()}.{_ho}{_dem}{_mssv}@sis.hust.edu.vn"

            emails.append(email)

        with open(self.output_path, "w", encoding="utf-8") as f:
            for email in emails:
                f.write(email + "\n")

print("Done.")
        