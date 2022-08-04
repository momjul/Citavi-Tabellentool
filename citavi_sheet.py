import pandas as pd
from datetime import date

class sheet_cleaner:
    def __init__(self):
        path = input("Dateipfad angeben: ")
        df = pd.read_excel(path)
        
        df = df.drop(columns=["Unnamed: 0", "Unnamed: 1", "Unnamed: 2"], axis=1)
        df_print = df[df["DOI"].isnull()]
        df_ebook = pd.concat([df, df_print], ignore_index=True, sort=False).drop_duplicates(keep=False)
                
        if not df_print.empty:
            print_order = self.remove_empty_columns(df_print)
            self.printer(print_order)
        
        if not df_ebook.empty:
            ebook_order = self.remove_empty_columns(df_ebook)
            self.printer(ebook_order, category="_E-Book")


    def remove_empty_columns(self, df):
        try:
            df.rename(columns={"Jahr ermittelt" : "Jahr", 
                               "RVK (= Freitext 1)" : "RVK",
                               "Budget (= Freitext 2)" : "Budget",
                               "Anzahl (= Freitext 3)" : "Anzahl",
                               "Standort (= Freitext 4)" : "Standort",
                               "Anmerkung (= Freitext 5)" : "Anmerkung",
                               "Autor, Herausgeber oder Institution" : "Autor/Herausgeber"}, inplace=True)
        except Exception as ex:
            print("Fehler:", ex)
        
        for i in df.columns:
            if i not in ["ISBN", "Titel", "RVK", "Budget", "Anzahl", "Standort"]:
                if df[i].isnull().all():
                    df = df.drop(i, axis=1)
            elif df[i].isnull().any():
                print("Achtung! Feld '" + i + "' ist nicht vollständig ausgefüllt.")
        
        return df


    def printer(self, df, category="_Print"):    
        new_file = input("Zielpfad eingeben: ") + str(date.today()) + category + ".xlsx"
        
        try:
            with pd.ExcelWriter(new_file, engine="xlsxwriter") as writer:
                
                df.to_excel(writer, startrow = 1, sheet_name='Tabelle 1', index=False)
                
                #workbook = writer.book
                worksheet = writer.sheets["Tabelle 1"]
                
                for i, column in enumerate(df.columns):    
                    column_name = df[column].astype(str)
                    column_length = column_name.str.len()
                    column_length = column_length.max()
                        
                    if column_length > 50:
                        column_length = 50
                    else:
                        column_length = max(column_length, len(column)) + 1.5
                    
                    worksheet.set_column(i, i, column_length)
            
            print("Datei erstellt unter", new_file)
            
        except Exception as ex:
            print("Fehler:", ex)


if __name__ == "__main__":
    sheet_cleaner()