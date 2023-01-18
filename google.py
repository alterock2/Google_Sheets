import pandas as pd
import pygsheets
import openpyxl

gc = pygsheets.authorize(service_file=r"C:\Users\user\Desktop\stable-woods-374912-6149e61555af.json")

df = pd.read_excel(r"C:\Users\user\Desktop\Сервис.xlsx")

sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/14clInBRzbD8RzxUqg7aXv-u2eeQw_DrLHaoEn2J8GHc/edit#gid=931961663g')

wks = sh[1]

df['Запчасть'] = df['Запчасть'].fillna('Ждем')

#wks.set_dataframe(df, (1, 1))

color = {
#"blue": 1,
  "red": 2,
  "green": 23,
  "alpha": 1
}

wks2 = sh[0]

wks2.add_conditional_formatting('N', 'N', 'NUMBER_GREATER_THAN_EQ', {'backgroundColor':{'red': 0.8},
                                'textFormat': {'bold': True}}, ['30'])

wks2.add_conditional_formatting('N', 'N', 'NUMBER_BETWEEN', {'backgroundColor':
{"red": 250/255, "green": 77/255, "blue": 77/255, "alpha": 1}}, ['14', '30'])

wks2.add_conditional_formatting('N', 'N', 'NUMBER_LESS_THAN_EQ', {'backgroundColor':
{"red": 255/255, "green": 158/255, "blue": 158/255, "alpha": 1}}, ['14'])

wks2.adjust_column_width(start=1, end=1, pixel_size=30)
#read = wks.get_as_df()

#print(read)

wks2.add_conditional_formatting('K', 'K', 'TEXT_EQ', {'backgroundColor':color}, ['Выдан'])

header = wks2.cell('A1')
header.set_text_format('bold', True)
header.update()

#for i, row in df.iterrows():
#  if row['Товар'] == 'Отпариватель':
#    if row['Статус'] == 'Получено':
#      print("yes of course")
#    else:
#      print('no')
