import sys
import win32com.client
import json
import numpy as np
import csv
import math
import pandas as pd
import mdp_func
# Читаем режим и задаем шаблоны.
rastr = win32com.client.Dispatch('Astra.Rastr')
file_rgm="C:\Users\otrok\Downloads\regime (2).rg2"
shbl_ut = "C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\траектория утяжеления.ut2"
shbl_sech = "C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\сечения.sch"
shbl_reg = "C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2"
rastr.Load(1, file_rgm, shbl_reg)
rastr.NewFile(shbl_ut)
rastr.NewFile(shbl_sech)
ut_table = rastr.Tables('ut_node')
node_table = rastr.Tables('node')
vetv_table = rastr.Tables('vetv')
# Задаем параметры утяжеления.
vector = pd.read_csv(r'C:\Users\otrok\Downloads\vector.csv', delimiter = ',')
mdp_func.set_vector(rastr, vector)
# Задаем сечения.
flowgate = json.load(open(r'C:\Users\otrok\Downloads\flowgate.json',
                          encoding='utf-8'))
sechen_table = rastr.Tables('sechen')
mdp_func.set_flowgate(rastr, flowgate)
print("Начальный переток в КС-", round(sechen_table.Cols('psech').Z(0), 2))
# Утяжеляем.
mdp_func.worsening_norm(rastr)
limit_overflow = abs(sechen_table.Cols('psech').Z(0))
print("----------------КРИТЕРИЙ ПО СТАТИКЕ-------------------- ")
print("Предельный переток в КС-", round(limit_overflow, 2))

# 1 критерий по P в нормальном режиме.
nereg = 30
mdp1 = 0.8*limit_overflow - nereg
print("МДП по 1 критерию", mdp1)

# 2 критерий по U в нормальном режиме.
print("----------------КРИТЕРИЙ ПО U В НОРМАЛЬНОМ РЕЖИМЕ-------------------- ")
rastr.Load(3, file_rgm, shbl_reg)
mdp_func.worsening_U(rastr, 1.15)
mdp2 = abs(sechen_table.Cols('psech').Z(0)) - nereg
print("МДП по 2 критерию-", mdp2)


# 3 критерий по P в послеаварийном режиме.
print("----------------КРИТЕРИЙ ПО СТАТИКЕ В ПАР-------------------- ")
faults = json.load(open(r'C:\Users\otrok\Downloads\faults.json', encoding = 'utf-8'))
doavar_overflow = mdp_func.second_criterion(rastr, faults, shbl_reg)
print("Доаварийный переток в КС-", doavar_overflow)
mdp3 = min(doavar_overflow) - nereg
print("МДП по критерию P в ПАР -", mdp3 )

# 4 критерий по U в ПАР.
print("----------------КРИТЕРИЙ ПО U В ПАР-------------------- ")
rastr.Load(3, file_rgm, shbl_reg)
doavar_overflow2 = mdp_func.fourth_criterion(rastr, faults, shbl_reg)
mdp4 = min(doavar_overflow2) - nereg
print("Доаварийный переток U-", doavar_overflow2)
print("МДП по критерию U в ПАР -", mdp4 )


# 5 критерий по I в нормальном режиме.
print("----------------КРИТЕРИЙ ПО I В НОРМАЛЬНОМ РЕЖИМЕ-------------------- ")
rastr.Load(3, file_rgm, shbl_reg)
mdp_func.worsening_I(rastr,'i_dop_r')
mdp5 = abs(sechen_table.Cols('psech').Z(0)) - nereg
print("МДП по 5 критерию-", mdp5)


# 6 критерий по I в ПАР.
print("----------------КРИТЕРИЙ ПО I В ПАР-------------------- ")
rastr.Load(3, file_rgm, shbl_reg)
doavar_overflow3 = mdp_func.sixth_criterion(rastr, faults, shbl_reg)
mdp6 = min(doavar_overflow3) - nereg
print("Доаварийный переток I-", doavar_overflow3)
print("МДП по критерию I в ПАР -", mdp6)
print("-----------------------РЕЗУЛЬТИРУЮЩИЙ МДП------------------------")
print(round(min(mdp1, mdp2, mdp3, mdp4, mdp5, mdp6)), 2)
