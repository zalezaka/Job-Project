import pandas as pd

#Выкачка всех файлов из папки (КС)
df1 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/КС1.xlsx')
df2 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/КС2.xlsx')
df3 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/КС3.xlsx')
df4 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/КС4.xlsx')
df5 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/КС5.xlsx')
df6 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/КС6.xlsx')
df7 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/КС7.xlsx')
df8 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/КС8.xlsx')

#Выкачка всех файлов из папки (ДС)
df9  = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/ДС1.xlsx')
df10 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/ДС2.xlsx')
df11 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/ДС3.xlsx')
df12 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/ДС4.xlsx')
df13 = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Входящие данные/ДС5.xlsx')


#Соединение КС и ДС файвлов в единый
frames = [df1, df2, df3,df4, df5, df6,df7,df8,df9, df10,df11,df12,df13 ]
ks_ds = pd.concat(frames)


#Переименовка столбцов 
ks_ds = ks_ds.rename(columns = {ks_ds.columns[0] : 'region', 
                         ks_ds.columns[1] : 'codmo',
                         ks_ds.columns[2] : 'namemo',
                         ks_ds.columns[3] : 'ksg',
                         ks_ds.columns[4] : 'nameksg' ,   
                         ks_ds.columns[5] : 'mkb',    
                         ks_ds.columns[6] : 'namemkb' ,   
                         ks_ds.columns[7] : 'usluga',
                         ks_ds.columns[8] : 'nameusluga',
                         ks_ds.columns[10] : 'cases',
                         ks_ds.columns[11] : 'sums',
                         ks_ds.columns[12] : 'kzsum',
                         ks_ds.columns[13] : 'bssum',
                         ks_ds.columns[14] : 'kusum',
                         ks_ds.columns[17] : 'kslpsu',
                         ks_ds.columns[18] : 'kdsum'
                })


#Обработка типов данных
ks_ds['kusum'] = ks_ds['kusum'].astype('int')
ks_ds['kslpsu'] = ks_ds['kslpsu'].astype('int')
ks_ds['kusum'] = ks_ds['kusum'].replace(0, 1)
ks_ds['kslpsu'] = ks_ds['kslpsu'].replace(0, 1)
ks_ds['ksg'] = ks_ds['ksg'].astype('object')
ks_ds['region'] = ks_ds['region'].astype('object')
ks_ds['mkb'] = ks_ds['mkb'].astype('object')

#Создание столбцы с условиями по типу группы st или ds
ks_ds['usl'] = ks_ds['ksg'].apply(lambda x: 'Круглосуточный стационар' if x.startswith('st') else 'Дневной стационар' )

#Создание столбцов со средними коэфициентами и БС (делим столбцы с суммарными значениями на кол-во случаев)
ks_ds['kz'] =  ks_ds['kzsum']/ks_ds['cases']
ks_ds['bs'] = ks_ds['bssum']/ks_ds['cases']
ks_ds['spk'] = (ks_ds['kusum']*ks_ds['kslpsu'])/ks_ds['cases']

ks_ds['kz'] = ks_ds['kz'].astype('float')
ks_ds['bs'] = ks_ds['bs'].astype('float')
ks_ds['spk'] = ks_ds['spk'].astype('float')
ks_ds['cases'] = ks_ds['cases'].astype('float')
ks_ds['sums'] = ks_ds['sums'].astype('float')

#Индексация столбоцов только по коду ксг и коду диагноза
ks_ds['index_str'] = ks_ds['ksg']+ks_ds['mkb']



#Скачиваем таблицу с группами КСГ и КЗ (создана заранее) и делаем из таблицы обычный словарь для работы в цикле
ksg = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/КСГ.xlsx')
slovar = ksg.set_index('ksg')['kz'].to_dict()


#Создание первой сводной таблицы по Медицинским организациям, субъектам, условиям
Svodnaya_1 = ks_ds.groupby(['region', 'usl','codmo', 'namemo'], as_index = False).agg({'kz':'mean','spk':'mean', 'bs':'mean','cases':'sum','sums':'sum' })
#Индекс для соединения со второй сводной
Svodnaya_1['index_to_merge'] = Svodnaya_1['region'] + Svodnaya_1['usl'] + Svodnaya_1['namemo']


#Скачиваем переходную таблицу из папки
Perehodnaya = pd.read_excel('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на Питоне/Переходная таблица.xlsx', sheet_name ='Переходная')


Perehodnaya = Perehodnaya.rename(columns = {Perehodnaya.columns[0] : 'status',
                                            Perehodnaya.columns[1] : 'oldksg',
                                            Perehodnaya.columns[2] : 'oldkz',
                                            Perehodnaya.columns[3] : 'newksg',
                                            Perehodnaya.columns[4] : 'newkz'
                              
                             })


#Создаем второй виртуальный датафрейм для изменения в нем (это будет прогнозирование после изменений)
ks_ds_transform = ks_ds


#Цикл обработки второго виртуального датафрейма по условиям из таблицы Переходов

for i, row in Perehodnaya.iterrows():
        if row[0] == 'Изменение группы' :
            ks_ds_transform.loc[ks_ds_transform['ksg'] == row[1], 'kz'] = ks_ds_transform.loc[ks_ds_transform['ksg'] == row[1], 'kz'].replace(row[2],row[4])
        elif row[0] == 'Перенос группы' :
            ks_ds_transform.loc[ks_ds_transform['ksg'] == row[1], 'kz'] = ks_ds_transform.loc[ks_ds_transform['ksg'] == row[1], 'kz'].replace(slovar[row[1]],slovar[row[3]])
            ks_ds_transform.loc[ks_ds_transform['ksg'] == row[1], 'ksg'] = row[3]


#Создание прогнозируемой сводной таблицы по Медицинским организациям, субъектам, условиям
Svodnaya_2 = ks_ds_transform.groupby(['region', 'usl','codmo', 'namemo'], as_index = False)                    .agg({'kz':'mean','spk':'mean', 'bs':'mean','cases':'sum'})

Svodnaya_2['new_sum'] = Svodnaya_2['kz']*Svodnaya_2['spk']*Svodnaya_2['bs']*Svodnaya_2['cases']
Svodnaya_2['index_to_merge'] = Svodnaya_2['region'] + Svodnaya_2['usl'] + Svodnaya_2['namemo']


#Создание файла на вывод из скрипта путем соединения двух сводных через индекс
output = Svodnaya_2.merge(Svodnaya_1, on = 'index_to_merge', how = 'left')[['region_x','usl_x', 'codmo_x', 'namemo_x', 'kz_x', 'kz_y','spk_x','bs_x','cases_x','sums','new_sum']]
output['Разница'] = output['new_sum'] - output['sums']


output = output.rename(columns = {'region_x' : 'Субъект_РФ',
                                  'usl_x' : 'Условия_МП',
                                  'codmo_x' : 'Код_МО',
                                  'namemo_x' : 'Наименование_МО',
                                  'kz_y' : 'Коэф_затратоемкости_старый',
                                  'kz_x' : 'Коэф_затратоемкости_новый',
                                  'bs_x' : 'Базовая ставка',
                                  'sums' : 'Сумма_затрат_старая',
                                  'new_sum' : 'Сумма_затрат_новая'
                           
                             })

output.to_csv('//corp.rosmedex.ru/space/obmen/%Отдел анализа ресурсов здравоохранения/_Личное/Шалаева ЕА/Моделирование фед/Решение на питоне/output.csv', encoding = 'windows-1251', sep = ';')



