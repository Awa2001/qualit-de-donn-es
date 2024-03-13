import os
import sqlite3

from openpyxl import Workbook


dir_path = os.path.dirname(os.path.realpath('C:\\Users\\falil\\PycharmProjects\\donnees'))


conn = sqlite3.connect('assurances.db')
cursor = conn.cursor()


cursor.execute('''CREATE TABLE IF NOT EXISTS Assurances (
                    PermisID INTEGER PRIMARY KEY,
                    Nom TEXT,
                    Prenom TEXT,
                    DateNaissance DATE,
                    Adresse TEXT,
                    Canton TEXT,
                    Assurance TEXT,
                    AssureDepuis DATE
                )''')


cursor.execute('''DELETE FROM Assurances 
                  WHERE PermisID NOT IN 
                  (SELECT MIN(PermisID) FROM Assurances GROUP BY Nom, Prenom, DateNaissance)''')


cursor.execute('''UPDATE Assurances 
                  SET Nom = REPLACE(Nom, '[^a-zA-Z0-9\s]', ''),
                      Prenom = REPLACE(Prenom, '[^a-zA-Z0-9\s]', ''),
                      Adresse = REPLACE(Adresse, '[^a-zA-Z0-9\s]', ''),
                      Canton = REPLACE(Canton, '[^a-zA-Z0-9\s]', ''),
                      Assurance = REPLACE(Assurance, '[^a-zA-Z0-9\s]', '')
               ''')


cursor.execute('''DELETE FROM Assurances 
                  WHERE Nom = '' OR Prenom = '' OR DateNaissance = '' OR Adresse = '' OR Canton = '' OR Assurance = '' OR AssureDepuis = '' 
               ''')

cursor.execute(r'''UPDATE Assurances 
                  SET DateNaissance = NULL 
                  WHERE DateNaissance < AssureDepuis
               ''')



conn.commit()


cursor.execute("SELECT * FROM Assurances")
rows = cursor.fetchall()


wb = Workbook()
ws = wb.active


for row_index, row in enumerate(rows):
    for col_index, cell_value in enumerate(row):
        ws.cell(row=row_index + 1, column=col_index + 1, value=cell_value)


excel_file_path = os.path.join(dir_path, "assurances.xlsx")
wb.save(excel_file_path)


conn.close()

print("Données extraites avec succès et enregistrées dans le fichier assurances.xlsx.")

