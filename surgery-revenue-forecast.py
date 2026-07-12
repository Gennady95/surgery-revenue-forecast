import numpy as np
import pandas as pd
import re
import os
import sqlalchemy
import requests
import xlsxwriter
import telebot
import getpass
import platform
import time
from dateutil import parser
from datetime import datetime, timedelta
from fast_bitrix24 import Bitrix
from more_itertools.recipes import unique
from sqlalchemy import create_engine
from dotenv import dotenv_values
from mysql.connector import Error
import plotly
import plotly.graph_objs as go
import plotly.express as px
from plotly.subplots import make_subplots
import dash
import dash_bootstrap_components as dbc
from dash import dash_table
from dash import dcc
from dash import html
from dash.dependencies import Input, Output

# Паттерны
re_1 = r'[^0-9,.;/]' # Регулярное выражение для отсева букв и знаков
pd.set_option('display.max_columns', None)
pd.set_option('mode.chained_assignment', None)
pd.options.mode.chained_assignment = None
# Создание коннекторов данных
datename = datetime.now().strftime('%d.%m %H.%M.%S') # Время создания файла
bot = telebot.TeleBot(os.getenv("TELEGRAM_TOKEN"))
engine = create_engine(os.getenv("DB_URL"))
#chid_list = ["Конс каб 72", "Конс каб 77", "Операционная № 1", "Операционная № 2", "Операционная № 3", "Опер 1.1", "Опер 1.2", "Опер 2.1", "Опер 2.2", "Опер 3.2"]
consultations_list = [10001714, 10001715, 10001716, 10003573, 10003574, 10003575, 10003576, 10007819, 10008011, 10008012, 10008133, 10008134, 10008151, 10008795, 10009179, 10009316, 10007306, 10008315, 10008316, 10001959, 10001960, 10008359, 10003829, 10003830]
start_time = time.time()
def GetSQL():
	global doctor_list, start_date, end_date, end_date_plan, start_time
	# SQL запросы
	lightquery_SCHEDULE = "SELECT schedid, dcode, dcode1, workdate, bhour, bmin, fhour, fmin, pcode, treatcode, status, reasid, chid, depnum, createdate, creatorid, clvisit from SCHEDULE WHERE pcode IS NOT NULL" # Расписание
	lightquery_BI_DOCTORS = "SELECT dcode, dname, depnum, lockdate from BI_DOCTORS" # Доктора WHERE lockdate IS NULL - без уволенных
	lightquery_SCHREASONS = "SELECT rid, rname FROM SCHREASONS" # Поводы назначения
	lightquery_SHEDMARKS = "SELECT mrkid, mrktext FROM SHEDMARKS" # Маркеры расписания
	lightquery_CLIENTS = "SELECT pcode, fullname, grtype FROM CLIENTS" # Клиенты
	lightquery_CHAIRS = "SELECT chid, chname FROM CHAIRS" # Рабочие места (кресла)
	lightquery_TREAT = "SELECT treatcode, amountcl_disc, kateg FROM TREAT" # Лечения (отчёт по приёмам) WHERE (depnum = '120363') or (depnum = '118') - только отделения косметологии и хирургии
	lightquery_DEPARTMENTS = "SELECT depnum, depname FROM DEPARTMENTS" # Список отделений
	lightquery_ROOMS = "SELECT rid, rname FROM ROOMS" # Кабинеты
	lightquery_N_NARADTYPES = "SELECT narid, aname FROM N_NARADTYPES" # Типы наряда (продукция, расходники) для отчёта по приёмам
	lightquery_CLGROUP = "SELECT grcod, grname FROM CLGROUP" # Типы наряда (продукция, расходники) для отчёта по приёмам
	lightquery_JPAGREEMENT = "SELECT agrid, agname FROM JPAGREEMENT" # Договоры с клиентами (купоны, сертификаты)
	lightquery_TREATSCH = "SELECT treatcode, schid, scount, workdate FROM TREATSCH" # Состав лечения (список услуг)
	lightquery_WSCHEMA = "SELECT schid, schname FROM WSCHEMA"  # Описание услуг из прайс-листа
	lightquery_DAILYPLAN = "SELECT dcode, pcode, pdate, summarub FROM DAILYPLAN"  # Планы услуги
	# Получение данных
	SCHEDULE = pd.read_sql(lightquery_SCHEDULE, engine) # Расписание
	BI_DOCTORS = pd.read_sql(lightquery_BI_DOCTORS, engine) # dict_BI_DOCTORS_name, dict_BI_DOCTORS_dep
	SCHREASONS = pd.read_sql(lightquery_SCHREASONS, engine) # dict_SCHREASONS
	SHEDMARKS = pd.read_sql(lightquery_SHEDMARKS, engine) # SHEDMARKS
	CLIENTS = pd.read_sql(lightquery_CLIENTS, engine) # dict_CLIENTS
	CHAIRS = pd.read_sql(lightquery_CHAIRS, engine) # dict_CHAIRS
	TREAT = pd.read_sql(lightquery_TREAT, engine) # dict_TREAT_COST
	TREATSCH = pd.read_sql(lightquery_TREATSCH, engine) # dict_TREATSCH
	DEPARTMENTS = pd.read_sql(lightquery_DEPARTMENTS, engine) # dict_DEPARTMENTS
	ROOMS = pd.read_sql(lightquery_ROOMS, engine) # dict_ROOMS
	N_NARADTYPES = pd.read_sql(lightquery_N_NARADTYPES, engine) # dict_N_NARADTYPES
	ARGJ = pd.read_sql(lightquery_CLGROUP, engine) # dict_ARGJ
	JPAGREEMENT = pd.read_sql(lightquery_JPAGREEMENT, engine) # dict_JPAGREEMENT
	WSCHEMA = pd.read_sql(lightquery_WSCHEMA, engine) # dict_WSCHEMA
	DAILYPLAN = pd.read_sql(lightquery_DAILYPLAN, engine) # dict_DAILYPLAN
	print(f"Все базы загружены в буфер за :{(time.time() - start_time):.2f}"); start_time = time.time()
	# Формирование списка докторов
	doctor_list = BI_DOCTORS[(BI_DOCTORS['depnum'] == 120363) & (BI_DOCTORS['lockdate'].isnull())]  # Создаём список действующих врачей хирургов
	doctor_list = list(filter(None, doctor_list['dname'].tolist())); print(doctor_list)
	# Обрезка баз данных для сокращения объёма словарей
	try:
		xlsx = pd.ExcelFile("Обновляемый план операций по расписанию" + ".xlsx")
		SCHED = xlsx.parse('Выполненные услуги')
		start_date = SCHED['Дата назначения'].max()
		end_date = datetime.now() - timedelta(days=1); print("Скрипт работает по алгоритму пополнения. Дата начала " + str(start_date) + ", дата крайних значений " + str(end_date))
	except:
		SCHED = pd.DataFrame()
		start_date = datetime(2017, 1, 1)  # datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1) #datetime(2023, 12, 22)
		end_date = datetime.now() - timedelta(days=1) # datetime(2024, 5, 12)  #datetime.now() #datetime(2023, 12, 4)
		print("Скрипт работает по алгоритму полного обновления базы. Дата начала " + str(start_date) + ", дата крайних значений " + str(end_date))
	end_date_plan = datetime.now() + timedelta(days=365)  # datetime(2024, 5, 12)  #datetime.now() #datetime(2023, 12, 4)
	DataListFirst = pd.date_range(min(start_date, end_date), max(start_date, end_date)).strftime('%Y-%m-%dT%H:%M:%S').tolist()  # Создание списка дат
	print("Дата горизонта планирования: " + str(end_date_plan))
	print("Длина базы SCHEDULE " + str(len(SCHEDULE)))
	print(f"Периоды определены, списки докторов сформированы за {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE = SCHEDULE[((SCHEDULE['workdate'] >= start_date) & (SCHEDULE['workdate'] <= end_date_plan))]
	print("Длина базы SCHEDULE " + str(len(SCHEDULE)))
	SCHEDULE_clients_list = SCHEDULE['pcode'].tolist(); print("словарь собран из " + str(len(SCHEDULE_clients_list)) + " клиентов")
	#CLIENTS = CLIENTS[(CLIENTS['pcode'].isin(list(filter(None, SCHEDULE_clients_list))))]
	SCHEDULE_treats_list = SCHEDULE['treatcode'].tolist(); print("словарь собран из " + str(len(SCHEDULE_treats_list)) + " кодов лечений")
	TREAT = TREAT[(TREAT['treatcode'].isin(list(filter(None, SCHEDULE_treats_list))))]
	TREAT['null'] = 0 # Дополнительный стобец для присвоения нулей
	TREATSCH = TREATSCH[(TREATSCH['treatcode'].isin(list(filter(None, SCHEDULE_treats_list))))]
	DAILYPLAN = DAILYPLAN[(DAILYPLAN['pcode'].isin(list(filter(None, SCHEDULE_clients_list))))]
	print(f"Базы урезаны за {(time.time() - start_time):.2f}"); start_time = time.time()
	# Оформление словарей
	dict_DEPARTMENTS = dict(DEPARTMENTS[['depnum', 'depname']].values)  # Словарь соответствий ID отделений и названий отделений
	dict_BI_DOCTORS_name = dict(BI_DOCTORS[['dcode', 'dname']].values) # Словарь соответствий ID сотрудника и полное имя
	dict_BI_DOCTORS_dep = dict(BI_DOCTORS[['dcode', 'depnum']].values) # Словарь соответствий ID сотрудника и отделение
	dict_SCHREASONS = dict(SCHREASONS[['rid', 'rname']].values) # Словарь соответствий ID повода и наименование повода назначения
	dict_SHEDMARKS = dict(SHEDMARKS[['mrkid', 'mrktext']].values) # Словарь соответствий ID марок расписания и наименований
	dict_CLIENTS = dict(CLIENTS[['pcode', 'fullname']].values) # Словарь соответствий ID клиента и имена клиентов
	dict_CLIENTS_GRTYPE = dict(CLIENTS[['pcode', 'grtype']].values) # Словарь соответствий ID клиента и договор (наличный, не наличный расчёт)
	dict_CHAIRS = dict(CHAIRS[['chid', 'chname']].values) # Словарь соответствий ID рабочего места и его названия
	dict_ROOMS = dict(ROOMS[['rid', 'rname']].values) # Словарь соответствий ID рабочего места и его названия
	dict_N_NARADTYPES = dict(N_NARADTYPES[['narid', 'aname']].values) # Словарь соответствий ID типа наряда и его названия
	dict_TREAT_COST = dict(TREAT[['treatcode', 'amountcl_disc']].values) # Словарь соответствий ID лечения и суммы, начисленной пациенту с учётом скидки
	TREAT['kateg'] = TREAT['kateg'] + TREAT['null']
	TREAT['kateg'].to_excel('kateg ' + '.xlsx', sheet_name='data', index=False) #Пишем результат в файл "TestWestcall.xlsx"
	dict_TREAT_KATEG = dict(TREAT[['treatcode', 'kateg']].values) # Словарь соответствий ID лечения и договора
	dict_ARGJ = dict(ARGJ[['grcod', 'grname']].values) # Словарь соответствий ID категорий расчётов
	dict_JPAGREEMENT = dict(JPAGREEMENT[['agrid', 'agname']].values) # Словарь соответствий ID договоров и наименований
	dict_WSCHEMA =  dict(WSCHEMA[['schid', 'schname']].values) # Словарь соотвествий ID услуги и наименования
	print(f"Словари сформированы за {(time.time() - start_time):.2f}"); start_time = time.time()
	# Произведение замен по словарям и предрасчёт необходимых показателей
	TREATSCH = TREATSCH[(TREATSCH['schid'].isin(list(filter(None, consultations_list))))]
	TREATSCH['Название услуги'] = TREATSCH['schid'].map(dict_WSCHEMA) # Присвоение наименований услуг по ID лечений
	TREATSCH = TREATSCH[['treatcode', 'Название услуги']]
	print(f"Произведены подстановки названий услуг за {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE['dcode'] = SCHEDULE['dcode'].replace(-1, np.nan)
	SCHEDULE['Доктор'] = SCHEDULE[['dcode', 'dcode1']].bfill(axis=1).iloc[:, 0] # Объединение двух полей с кодом доктора
	print(f"Произведены подстановки докторов за {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE['JPAGREEMENT_pcode'] = SCHEDULE['pcode'].map(dict_CLIENTS_GRTYPE)  # Присвоение договора по ID клиента из таблицы приёмов
	print(f"Произведены подстановки договоров клиентам по клиентскому справочнику {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE['JPAGREEMENT_treatcode'] = SCHEDULE['treatcode'].map(dict_TREAT_KATEG) # Присвоение договора по ID клиента из таблицы приёмов
	#SCHEDULE['JPAGREEMENT_treatcode'] = SCHEDULE['JPAGREEMENT_treatcode'].replace(0, np.nan)
	print(f"Произведены подстановки договоров клиентам по фактическим лечениям {(time.time() - start_time):.2f}"); start_time = time.time()
	#SCHEDULE['Категория (наличный расчёт)'] = SCHEDULE['JPAGREEMENT_treatcode'].fillna(TREAT['null'])
	SCHEDULE['Категория (наличный расчёт)'] = SCHEDULE['JPAGREEMENT_treatcode'].fillna(SCHEDULE['JPAGREEMENT_pcode'])
	SCHEDULE['Категория (наличный расчёт)'] = SCHEDULE['Категория (наличный расчёт)'].replace(1, 'Наличный расчёт') # "Наличный расчёт" если = 1
	print(f"Произведены подстановки категории расчёта за {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE['Отделение доктора'] = SCHEDULE['Доктор'].map(dict_BI_DOCTORS_dep) # Отдельная колонка с присвоение доктору кода отделения
	SCHEDULE['Отделение'] = SCHEDULE['depnum'].fillna(SCHEDULE['Отделение доктора']) # Присвоить записи отделение доктора, если не указано в явном виде в расписании
	SCHEDULE['Отделение'] = SCHEDULE['Отделение'].map(dict_DEPARTMENTS) # Присвоить название отделения по коду
	print(f"Произведены подстановки отделений за {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE['Продолжительность процедуры, минуты'] = (SCHEDULE['fhour'] - SCHEDULE['bhour'])*60 + (SCHEDULE['fmin'] - SCHEDULE['bmin']) # Расчёт продолжительности окна приёма в расписании
	print(f"Выполнен расчёт времени продолжительности операций {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE['Имя доктора'] = SCHEDULE['Доктор'].map(dict_BI_DOCTORS_name) # Присвоение имени доктора по ID
	print(f"Привоил имена докторам {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE['Имя пациента'] = SCHEDULE['pcode'].map(dict_CLIENTS) # Присвоение имён пациентам по их ID
	print(f"Привоил имена пациентам {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE['Стоимость лечения'] = SCHEDULE['treatcode'].map(dict_TREAT_COST) # Присвоение стоимости лечения по ID
	print(f"Привоил стоимости лечений {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE = SCHEDULE.merge(TREATSCH, on='treatcode', how='outer')
	print(f"Присвоены имена докторов, пациентов, добавлены названия услуг при помощи merge за {(time.time() - start_time):.2f}"); start_time = time.time()
	SCHEDULE['Статус'] = SCHEDULE['status'].map(dict_SHEDMARKS) # Присвоение статуса по ID
	SCHEDULE['Повод назначения'] = SCHEDULE['reasid'].map(dict_SCHREASONS) # Присвоение повода назначения по ID
	SCHEDULE['Рабочее место'] = SCHEDULE['chid'].map(dict_CHAIRS) # Присвоение рабочего места по ID
	SCHEDULE['Кто создал'] = SCHEDULE['creatorid'].map(dict_BI_DOCTORS_name) # replace() # Присвоение имени создателя записи по ID
	SCHEDULE['Кто создал'] = SCHEDULE['Кто создал'].replace(10000528, "Внешняя система") # Внешняя система
	SCHEDULE = SCHEDULE.sort_values(by=['workdate'], ascending=True) # Сортировка по дате назначения
	print(f"Мелкие присвоения и сортировка по дате {(time.time() - start_time):.2f}"); start_time = time.time()
	# Присвеоение имён
	SCHEDULE = SCHEDULE[['schedid', 'pcode', 'Доктор', 'treatcode', 'workdate', 'createdate', 'Продолжительность процедуры, минуты', 'clvisit', 'Отделение', 'Категория (наличный расчёт)', 'Статус', 'Повод назначения', 'Рабочее место', 'Имя пациента', 'Имя доктора', 'Кто создал', 'Стоимость лечения', 'Название услуги']] # 'Пациент для клиники', 'Пациент для направления', 'Пациент для доктора', 'Первая пластическая операция'
	SCHEDULE.columns = ['Код записи в расписании', 'Код пациента', 'Код доктора', 'Код лечения', 'Дата назначения', 'Дата создания', 'Продолжительность процедуры, минуты', 'Посещение', 'Отделение', 'Категория (наличный расчёт)', 'Статус', 'Повод назначения', 'Рабочее место', 'Имя пациента', 'Имя доктора', 'Кто создал', 'Стоимость лечения', 'Название услуги'] # 'Пациент для клиники', 'Пациент для направления', 'Пациент для доктора', 'Первая пластическая операция'
	VISIT_BASE(SCHEDULE, SCHED, DAILYPLAN)
def VISIT_BASE(base, baseold, plans):
	global basenew, start_time
	basenew = base[(base['Посещение'] == 1) & (base['Дата назначения'] <= end_date)] # Оставляем в визитах только с отметкой о фактическом посещении
	basenew = pd.concat([baseold, basenew], ignore_index=True); print("Формирую список фактически проведённых лечений")
	for dname in list(filter(None, basenew['Имя доктора'].unique().tolist())):
		basenew.loc[basenew[(basenew['Имя доктора'] == dname)].groupby('Код пациента')['Дата назначения'].idxmin(), 'Пациент для доктора'] = 'Первичный'
	for oname in list(filter(None, basenew['Отделение'].unique().tolist())):
		basenew.loc[basenew[(basenew['Отделение'] == oname)].groupby('Код пациента')['Дата назначения'].idxmin(), 'Пациент для направления'] = 'Первичный'
	basenew.loc[basenew.groupby('Код пациента')['Дата назначения'].idxmin(), 'Пациент для клиники'] = 'Первичный'
	basenew.loc[basenew[((basenew['Продолжительность процедуры, минуты'] > 40) & (basenew['Имя доктора'].isin(list(filter(None, doctor_list)))) & ((basenew['Статус'] == 'о/а операция') | (basenew['Статус'] == 'м/а операция')) & ((basenew['Отделение'] == 'Хирургия стационарная') | (basenew['Отделение'] == 'Хирургия амбулаторная') | (basenew['Отделение'] == 'Хирургия амбулаторная,Хирургия стационарная')))].groupby('Код пациента')['Дата назначения'].idxmin(), 'Первая пластическая операция'] = 'Да'
	print(f"Закончено формирование списка выполненных работ {(time.time() - start_time):.2f}"); start_time = time.time()
	OPERATIONS_BASE(base, plans)
def OPERATIONS_BASE(base, plans):
	global operatios_base, start_time
	operatios_base = base[((base['Продолжительность процедуры, минуты'] > 40) & (base['Имя доктора'].isin(list(filter(None, doctor_list)))) & ((base['Статус'] == 'о/а операция') | (base['Статус'] == 'м/а операция')) & (base['Категория (наличный расчёт)'] == 'Наличный расчёт') & (base['Повод назначения'] != 'коррекция') & ((base['Отделение'] == 'Хирургия стационарная') | (base['Отделение'] == 'Хирургия амбулаторная') | (base['Отделение'] == 'Хирургия амбулаторная,Хирургия стационарная')) & (base['Дата назначения'] <= end_date))]
	#consultations_base =
	PLAN_OPERATIONS_BASE(base, plans)
def PLAN_OPERATIONS_BASE(base, plans):
	global plan_base, start_time
	print(base)
	plan_base = base[((base['Продолжительность процедуры, минуты'] > 40) & (base['Имя доктора'].isin(list(filter(None, doctor_list)))) & ((base['Статус'] == 'о/а операция') | (base['Статус'] == 'м/а операция')) & (base['Категория (наличный расчёт)'] == 'Наличный расчёт') & (base['Повод назначения'] != 'коррекция') & ((base['Отделение'] == 'Хирургия стационарная') | (base['Отделение'] == 'Хирургия амбулаторная') | (base['Отделение'] == 'Хирургия амбулаторная,Хирургия стационарная')) & (base['Дата назначения'] > end_date))]
	with pd.ExcelWriter('Обновляемый план операций по расписанию' + '.xlsx') as writer:
		basenew.to_excel(writer, sheet_name='Выполненные услуги', index=False)
		operatios_base.to_excel(writer, sheet_name='Выполненные операции ПХ', index=False)
		plan_base.to_excel(writer, sheet_name='Планируемые операции ПХ', index=False)
	plan_base_schedid = plan_base['Код записи в расписании'].tolist()
	plan_base_pcode = plan_base['Код пациента'].tolist()
	plan_base_dcode = plan_base['Код доктора'].tolist()
	plan_base_date = plan_base['Дата создания'].tolist()
	print(plan_base_schedid)
	print(plan_base_pcode)
	print(plan_base_dcode)
	print(plan_base_date)
	PLANS_TREAT_MIN = pd.DataFrame(); PLANS_TREAT_MAX = pd.DataFrame(); print("Формирую список плановых лечений")
	for schedid, pcode, dcode, date in zip(plan_base_schedid, plan_base_pcode, plan_base_dcode, plan_base_date):  # Цикл проходит по списку SQL n количество раз, равному len(list) базы
		plans_TR = plans[((plans['dcode'] == dcode) & (plans['pcode'] == pcode) & (plans['summarub'].notnull()))]
		# Минимальная сумма выручки
		plans_TR_MIN = plans_TR[((plans_TR['summarub'] > 100000) & (plans_TR['summarub'] == plans_TR['summarub'].min()))]
		if ((plans_TR_MIN['summarub'].sum() <= 0) | (plans_TR_MIN['summarub'].sum() == np.nan)): plans_TR_MIN = plans_TR[(plans_TR['summarub'] == plans_TR['summarub'].max())]
		plans_TR_MIN['Код записи в расписании'] = schedid
		# Максимальная сумма выручки
		plans_TR_MAX = plans_TR[(plans_TR['summarub'] == plans_TR['summarub'].max())]
		plans_TR_MAX['Код записи в расписании'] = schedid
		PLANS_TREAT_MIN = pd.concat([PLANS_TREAT_MIN, plans_TR_MIN], ignore_index=True)
		PLANS_TREAT_MAX = pd.concat([PLANS_TREAT_MAX, plans_TR_MAX], ignore_index=True)

	print(PLANS_TREAT_MIN.columns.tolist())
	PLANS_TREAT_MIN = PLANS_TREAT_MIN[['Код записи в расписании', 'summarub']]; PLANS_TREAT_MIN.columns = ['Код записи в расписании', 'Сумма минимальной планируемой выручки']
	PLANS_TREAT_MAX = PLANS_TREAT_MAX[['Код записи в расписании', 'summarub']]; PLANS_TREAT_MAX.columns = ['Код записи в расписании', 'Сумма максимальной планируемой выручки']

	plan_base = plan_base.merge(PLANS_TREAT_MIN, on='Код записи в расписании', how='outer')
	plan_base = plan_base.merge(PLANS_TREAT_MAX, on='Код записи в расписании', how='outer')
	print(f"Завершено формирование планов работ {(time.time() - start_time):.2f}")
	# Удаление дубликатов операций
	plan_base['Дата назначения'] = pd.to_datetime(plan_base['Дата назначения'], errors='coerce')  # приводим к datetime (на всякий, вдруг там строки)
	plan_base['Дата без времени'] = plan_base['Дата назначения'].dt.normalize()  # дата без времени
	plan_base = (plan_base.sort_values('Дата назначения').drop_duplicates(['Дата без времени', 'Код пациента', 'Код доктора'], keep='first').drop(columns='Дата без времени'))  # САМЫЙ РАННИЙ приём в день
GetSQL()
with pd.ExcelWriter('Обновляемый план операций по расписанию' + '.xlsx') as writer:
	basenew.to_excel(writer, sheet_name='Выполненные услуги', index=False)
	operatios_base.to_excel(writer, sheet_name='Выполненные операции ПХ', index=False)
	plan_base.to_excel(writer, sheet_name='Планируемые операции ПХ', index=False)
print(f"Завершена запись в файл {(time.time() - start_time):.2f}")