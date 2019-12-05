from data import regions, professions_metallurgy, professions_media, professions_food
import functions
import xlwt
import csv

f = open('dataProfessionInFourIndustries.csv', 'r', encoding='UTF-8')
input_file = csv.DictReader(f)
data = list(input_file)

book = xlwt.Workbook('Data')
sheet = book.add_sheet('food')

functions.set_fields(sheet, professions_food, regions)
functions.fill_cells(sheet, data, professions_food, regions)

book.save('example.xls')