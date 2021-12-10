import sys
import win32com.client
import json
import numpy as np
import csv
import math
import pandas as pd
import mdp_func
from pprint import pprint
# Считываем файл утяжеления и разбиваем по стобцам.
vector = pd.read_csv(r'C:\Users\otrok\Downloads\vector.csv', delimiter = ',')
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
ut_table.size = len(vector)
for index in range(0, len(vector)):
    ut_table.Cols('ny').SetZ(index, vector['node'][index])
    ut_table.Cols('tg').SetZ(index, vector['tg'][index])
    if vector['variable'][index] == 'pn':
        ut_table.Cols('pn').SetZ(index, vector['value'][index])
    else:
        ut_table.Cols('pg').SetZ(index, vector['value'][index])
# Задаем сечения.
flowgate = json.load(open(r'C:\Users\otrok\Downloads\flowgate.json',
                          encoding='utf-8'))
grline_table = rastr.Tables('grline')
sechen_table = rastr.Tables('sechen')
grline_table.size = len(flowgate)
sechen_table.size = 1
sechen_table.Cols('ns').SetZ(0, 1)
for index, k in enumerate(flowgate.keys()):
    grline_table.Cols('ip').SetZ(index, flowgate[k]['ip'])
    grline_table.Cols('iq').SetZ(index, flowgate[k]['iq'])
grline_table.Cols('ns').SetZ(0, 1)
rastr.rgm('')
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
print("МДП по 2 критерию-", abs(sechen_table.Cols('psech').Z(0)) - nereg)


# 3 критерий по P в послеаварийном режиме.
print("----------------КРИТЕРИЙ ПО СТАТИКЕ В ПАР-------------------- ")
faults = json.load(open(r'C:\Users\otrok\Downloads\faults.json', encoding = 'utf-8'))
limit_overflow2 = []
doavar_overflow = []
for k in faults.keys():
    fault = mdp_func.faults(rastr, faults[k],  shbl_reg)
    mdp_func.worsening_norm(rastr)
    limit_overflow2.append(abs(sechen_table.Cols('psech').Z(0)))
    rastr.Load(3, r'C:\Users\otrok\Downloads\regime (2).rg2', shbl_reg)
    vetv_table.Cols('sta').SetZ(fault, faults[k]['sta'])
    rastr.rgm('')
    kd = rastr.step_ut("index")
    while (kd == 0) and abs(sechen_table.Cols('psech').Z(0)) < 0.92*limit_overflow2[-1]:
        kd = rastr.step_ut("z")
    vetv_table.Cols('sta').SetZ(fault, 1-faults[k]['sta'])
    rastr.rgm('')
    doavar_overflow.append(abs(sechen_table.Cols('psech').Z(0)))
limit_overflow2 = [round(v, 2) for v in limit_overflow2]
doavar_overflow = [round(v, 2) for v in doavar_overflow]
print("Послеаварийный переток в КС-", limit_overflow2)
print("Доаварийный переток в КС-", doavar_overflow)
mdp3 = min(doavar_overflow) - nereg
print("МДП по критерию P в ПАР -", mdp3 )

# 4 критерий по U в ПАР.
print("----------------КРИТЕРИЙ ПО U В ПАР-------------------- ")
rastr.Load(3, file_rgm, shbl_reg)
limit_overflow3 = []
doavar_overflow2 = []
for k in faults.keys():
    fault = mdp_func.faults(rastr, faults[k], shbl_reg)
    vetv_table.Cols('sta').SetZ(fault, faults[k]['sta'])
    mdp_func.worsening_U(rastr, 1.1)
    limit_overflow3.append(abs(sechen_table.Cols('psech').Z(0)))
    vetv_table.Cols('sta').SetZ(fault, 1-faults[k]['sta'])
    rastr.rgm('')
    doavar_overflow2.append(abs(sechen_table.Cols('psech').Z(0)))
limit_overflow3 = [round(v, 2) for v in limit_overflow2]
doavar_overflow2 = [round(v, 2) for v in doavar_overflow2]
mdp4 = min(doavar_overflow2) - nereg
print("Послеаварийный переток U-", limit_overflow3)
print("Доаварийный переток U-", doavar_overflow2)
print("МДП по критерию U в ПАР -", min(doavar_overflow2) - nereg )


# 5 критерий по I в нормальном режиме.
print("----------------КРИТЕРИЙ ПО I В НОРМАЛЬНОМ РЕЖИМЕ-------------------- ")
rastr.Load(3, file_rgm, shbl_reg)
mdp_func.worsening_I(rastr,'i_dop_r')
mdp5 = abs(sechen_table.Cols('psech').Z(0)) - nereg
print("МДП по 5 критерию-",abs(sechen_table.Cols('psech').Z(0)) - nereg)


# 6 критерий по I в ПАР.
print("----------------КРИТЕРИЙ ПО I В ПАР-------------------- ")
rastr.Load(3, file_rgm, shbl_reg)
limit_overflow4 = []
doavar_overflow3 = []
for k in faults.keys():
    fault = mdp_func.faults(rastr, faults[k], shbl_reg)
    vetv_table.Cols('sta').SetZ(fault, faults[k]['sta'])
    mdp_func.worsening_I(rastr, 'i_dop_r_av')
    limit_overflow4.append(abs(sechen_table.Cols('psech').Z(0)))
    vetv_table.Cols('sta').SetZ(fault, 1-faults[k]['sta'])
    rastr.rgm('')
    doavar_overflow3.append(abs(sechen_table.Cols('psech').Z(0)))
limit_overflow3 = [round(v, 2) for v in limit_overflow3]
doavar_overflow3 = [round(v, 2) for v in doavar_overflow3]
mdp6 = min(doavar_overflow3) - nereg
print("Послеаварийный переток I-", limit_overflow4)
print("Доаварийный переток I-", doavar_overflow3)
print("МДП по критерию I в ПАР -", min(doavar_overflow3) - nereg )
print("-----------------------РЕЗУЛЬТИРУЮЩИЙ МДП------------------------")
print(round(min(mdp1, mdp2, mdp3, mdp4, mdp5, mdp6)), 2)
