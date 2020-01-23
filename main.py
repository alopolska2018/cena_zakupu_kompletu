import win32com.client
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("secret.json", scope)
client = gspread.authorize(creds)
sheet = client.open('KOMPLETY')
worksheet = sheet.worksheet('komplety z niezerowym stanem')

oGT = win32com.client.Dispatch("InsERT.GT")
oGT.Produkt=1
oGT.Autentykacja= 0
oGT.Serwer="localhost\insertgt"
oGT.Uzytkownik="sa"
oGT.UzytkownikHaslo= ""
oGT.Operator="L Mateusz"
oGT.OperatorHaslo=''
oGT.Baza="test"
oSubiekt = oGT.Uruchom(0, 0)

gspread_sku = [item for item in worksheet.col_values(1) if item]
cena_zakupu_list = []
for sku in gspread_sku:
        cena_zakupu = oSubiekt.TowaryManager.WczytajTowar(sku).Zakupy.Element(1)
        cena_zakupu_list.append(str(cena_zakupu))

number_of_rows = len(cena_zakupu_list)
cell_list = worksheet.range('C1:C{}'.format(number_of_rows))

for cell, value in zip(cell_list, cena_zakupu_list):
        cell.value = value

worksheet.update_cells(cell_list)
