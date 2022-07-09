# -*- coding: cp1251 -*-
import time
import csv
import xlwings as xw
import numpy as np
import pandas as pd
from pyxll import xl_macro
from collections import defaultdict
from enpyxll.util.logs import log_error
from sixgill.pipesim import Model
from sixgill.definitions import ModelComponents, Parameters, Units, ALL, ProfileVariables, SystemVariables, Constants
from enpyxll import entry_point

# Открытие модели PIPESIM__________________________________________________________
@entry_point
@xl_macro
@log_error
def Open_Model_UI():
    Model.open_ui(xw.sheets['Инициал.'].range('A21').value)
# __________________________________________________________________________________

# Выгрузка списка трубопроводов______________________________________________________
@entry_point
@xl_macro
@log_error
def Import_Pipes():
    sht = xw.sheets['Инициал.']
    model = Model.open(sht.range('A21').value, Units.METRIC) 
    Pipe_Name = np.array(model.find(Flowline=ALL))[:, None]
    sht.range('I3').value  = Pipe_Name     
    model.save(sht.range('A21').value)
    model.close()
# _________________________________________________________________________________

# Выгрузка списка источников____________________________________________________
@entry_point
@xl_macro
@log_error
def Import_Sources():
    sht = xw.sheets['Инициал.']
    model = Model.open(sht.range('A21').value, Units.METRIC) 
    Source_Name = np.array(model.find(Source=ALL))[:, None]
    sht.range('J3').value = Source_Name 
    model.save(sht.range('A21').value)
    model.close()
# ____________________________________________________________________________

# Выгрузка списка стоков_______________________________________________________
@entry_point
@xl_macro
@log_error
def Import_Sinks():
    sht = xw.sheets['Инициал.']
    model = Model.open(sht.range('A21').value, Units.METRIC) 
    Sink_Name = np.array(model.find(Sink=ALL))[:, None]
    sht.range('L3').value = Sink_Name
    model.save(sht.range('A21').value)
    model.close()
# ____________________________________________________________________________

# Выгрузка списка трубопроводов, источников и стоков___________________________
@entry_point
@xl_macro
@log_error
def Import_All():
    sht = xw.sheets['Инициал.']
    model = Model.open(sht.range('A21').value, Units.METRIC)
    Pipe_Name = np.array(model.find(Flowline=ALL))[:, None]
    sht.range('I3').value  = Pipe_Name
    Source_Name = np.array(model.find(Source=ALL))[:, None]
    sht.range('J3').value = Source_Name
    Sink_Name = np.array(model.find(Sink=ALL))[:, None]
    sht.range('L3').value = Sink_Name
    model.save(sht.range('A21').value)
    model.close()
# ____________________________________________________________________________

# Загрузка данных расхода газа__________________________________________________________________
def Input_GasRate(model,Date_Num):
    shtg = xw.sheets['Газ']
    Source_Num = 2
    while shtg.range(2,Source_Num).value is not None:
        Source_Name = shtg.range(2,Source_Num).value
        Source_Value = shtg.range(Date_Num,Source_Num).value
        if shtg.range(1,1).value == 'Дебит газа, тыс.м3/сут':
            if Source_Value == 0: Source_Value = 0.1
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEFLUIDOVERRIDES, value=False)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.SELECTEDRATETYPE, value=Constants.FlowRateType.GASFLOWRATE)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.GASFLOWRATE, value=Source_Value/1000)
        else:
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEFLUIDOVERRIDES, value=True)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEGASRATIO, value='GOR')
            model.set_value(Source=Source_Name, parameter=Parameters.Source.GOR, value=Source_Value)           
        Source_Num += 1     
# ____________________________________________________________________________________________

# Загрузка данных расхода нефти/конденсата_____________________________________________________
def Input_OilRate(model,Date_Num):
    shto = xw.sheets['Нефть']
    Source_Num = 2
    while shto.range(2,Source_Num).value is not None:
        Source_Name = shto.range(2,Source_Num).value
        Source_Value = shto.range(Date_Num,Source_Num).value
        if shto.range(1,1).value in ['Дебит нефти, м3/сут','Дебит конденсата, м3/сут']:
            if Source_Value == 0: Source_Value = 0.1
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEFLUIDOVERRIDES, value=False)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.SELECTEDRATETYPE, value=Constants.FlowRateType.LIQUIDFLOWRATE)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.LIQUIDFLOWRATE, value=Source_Value)
        else:
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEFLUIDOVERRIDES, value=True)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEGASRATIO, value='OGR')
            model.set_value(Source=Source_Name, parameter=Parameters.Source.OGR, value=Source_Value*1000000)
            
        Source_Num += 1    
# ____________________________________________________________________________________________  

# Загрузка данных расхода воды_____________________________________________________
def Input_WaterRate(model,Date_Num):
    shtw = xw.sheets['Вода']
    Source_Num = 2
    while shtw.range(2,Source_Num).value is not None:
        Source_Name = shtw.range(2,Source_Num).value
        Source_Value = shtw.range(Date_Num,Source_Num).value
        if shtw.range(1,1).value == 'Дебит воды, м3/сут':
            if Source_Value == 0: Source_Value = 0.1
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEFLUIDOVERRIDES, value=False)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.SELECTEDRATETYPE, value=Constants.FlowRateType.LIQUIDFLOWRATE)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.LIQUIDFLOWRATE, value=Source_Value)
        elif shtw.range(1,1).value == 'ВГФ, м3/м3':
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEFLUIDOVERRIDES, value=True)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEWATERRATIO, value='WGR')
            model.set_value(Source=Source_Name, parameter=Parameters.Source.WGR, value=Source_Value*1000000)
        else:
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEFLUIDOVERRIDES, value=True)
            model.set_value(Source=Source_Name, parameter=Parameters.Source.USEWATERRATIO, value='WaterCut')    
            model.set_value(Source=Source_Name, parameter=Parameters.Source.WATERCUT, value=Source_Value)
            
        Source_Num += 1    
# ____________________________________________________________________________________________

# Загрузка данных устьевых температур_________________________________________________________
def Input_SourceTemperature(model,Date_Num):
    shtt = xw.sheets['Темп.Уст.']
    Source_Num = 2
    while shtt.range(2,Source_Num).value is not None:
        Source_Name = shtt.range(2,Source_Num).value
        Source_Value = shtt.range(Date_Num,Source_Num).value
        model.set_value(Source=Source_Name, parameter=Parameters.Source.TEMPERATURE, value=Source_Value)
  
        Source_Num += 1    
# ____________________________________________________________________________________________  

# Загрузка данных давлений точки сбора________________________________________________________
def Input_SinkPressure(model,Date_Num):
    shtp = xw.sheets['Давл.Сеп.']
    Sink_Num = 2
    while shtp.range(2,Sink_Num).value is not None:
        Sink_Name = shtp.range(2,Sink_Num).value
        Sink_Value = shtp.range(Date_Num,Sink_Num).value
        model.set_value(Sink=Sink_Name, parameter=Parameters.Sink.PRESSURE, value=Sink_Value)
        Sink_Num += 1    
# ____________________________________________________________________________________________

# Поиск оптимального диаметра трубопровода на критические даты труб__________________________________________________
def Find_ID(model,Date_Num,Diam_Dict,Pipes_Assortment, Pipes_Thickness, Assortment_Num,Disturbance_Counter,Good_Assortments,Bad_Assortments,Accepted_Diam,MD_Pipe_Dict,Log_File):
    for Iteration in range(100): # взято условно
        model.set_values(Diam_Dict)
        for PipeName in Disturbance_Counter:   # Зануление старых нарушений перед новой проверкой
            Disturbance_Counter[PipeName] = 0
        Disturbance_Counter = Diam_Check(model,MD_Pipe_Dict,Date_Num,Disturbance_Counter,Log_File)
        for PipeName in MD_Pipe_Dict[Date_Num]:
            if Assortment_Num[PipeName] == 0: Bad_Assortments[PipeName] += 1
            if Disturbance_Counter[PipeName] > 0 and Good_Assortments[PipeName] == 0:   # Если превышения на данном диаметре были и не было ранее типоразмеров без превышений
                Assortment_Num[PipeName] += 1
                Bad_Assortments[PipeName] += 1
            elif Disturbance_Counter[PipeName] == 0 and Bad_Assortments[PipeName] == 0:   # Иначе, если типоразмеров с нарушениями не было
                Assortment_Num[PipeName] -= 1
                Good_Assortments[PipeName] += 1
            else:  # Тогда,
                if Disturbance_Counter[PipeName] > 0:   # Если нарушения есть (но были типоразмеры без превышений) - берем на типоразмер выше
                    Bad_Assortments[PipeName] += 1
                    Assortment_Num[PipeName] += 1        
                    Accepted_Diam[PipeName] = True
                else:   # Иначе (превышений нет, но были типоразмеры с превышениями) - берем текущий типоразмер
                    Accepted_Diam[PipeName] = True
            Diam_Dict[PipeName]['InnerDiameter'] = Pipes_Assortment[Assortment_Num[PipeName]]-Pipes_Thickness[Assortment_Num[PipeName]]*2
            Diam_Dict[PipeName]['WallThickness'] = Pipes_Thickness[Assortment_Num[PipeName]]
        print (pd.Series({k:v for k,v in Accepted_Diam.items() if k in MD_Pipe_Dict[Date_Num]}),file=Log_File)
        print (pd.DataFrame({'Diameter':np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[0]*2+np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[1]}, index = Diam_Dict.keys()),file=Log_File)
        if 0 not in [v for k,v in Accepted_Diam.items() if k in MD_Pipe_Dict[Date_Num]]:
            model.set_values(Diam_Dict)
            return Diam_Dict, Assortment_Num
    return Diam_Dict, Assortment_Num
# ______________________________________________________________________________________________________________

# Проверка на критерии применимости____________________________________________________________
def Diam_Check(model,MD_Pipe_Dict,Date_Num,Disturbance_Counter,Log_File):
    model.tasks.networksimulation.reset_conditions()
    if xw.sheets['Spec'].range('D10').value == 'Газ':
        Results = model.tasks.networksimulation.run(profile_variables=[ProfileVariables.VELOCITY_GAS, ProfileVariables.TOTAL_DISTANCE, ProfileVariables.PRESSURE],
        system_variables=[SystemVariables.PRESSURE])
    else:
        Results = model.tasks.networksimulation.run(profile_variables=[ProfileVariables.VELOCITY_LIQUID, ProfileVariables.TOTAL_DISTANCE, ProfileVariables.PRESSURE],
        system_variables=[SystemVariables.PRESSURE])
    for PipeName in MD_Pipe_Dict[Date_Num]:
        for branch, profile in Results.profile.items():
            if PipeName in profile['BranchEquipment']:
                First_Segment = profile['BranchEquipment'].index(PipeName)
                for Last_Segment in range(First_Segment,len(profile['BranchEquipment'])):
                    if profile['BranchEquipment'][Last_Segment] is not None and Last_Segment != First_Segment:  # Если находит другой "объект"
                        Last_Segment-=1
                        break
                if xw.sheets['Spec'].range('D10').value == 'Газ': Last_Velocity = profile['VelocityGas'][Last_Segment]
                else: Last_Velocity = profile['VelocityLiquid'][Last_Segment]
        Crit_Velocity = 20 if xw.sheets['Spec'].range('D10').value == 'Газ' else 3
        try:
            if Last_Velocity > Crit_Velocity:
                Disturbance_Counter[PipeName] += 1
            print (PipeName,Last_Velocity,Crit_Velocity,file=Log_File)
        except UnboundLocalError:
            Disturbance_Counter[PipeName] += 1
    return Disturbance_Counter  
# ____________________________________________________________________________________________

# Корректировка диаметра с учетом устьевых давлений___________________________________________
def Correction_ID(model,Date_Num,Diam_Dict,Pipes_Assortment,Pipes_Thickness,Assortment_Num,Disturbance_Counter,MD_Source_Dict,Dict_Path,Source_Pres_Dict,Log_File):
    Specific_Drop_Pressure = defaultdict(dict)
    for Iteration in range(100): # взято условно
        model.tasks.networksimulation.reset_conditions()
        if xw.sheets['Spec'].range('D10').value == 'Газ':
            Results = model.tasks.networksimulation.run(profile_variables=[ProfileVariables.VELOCITY_GAS, ProfileVariables.TOTAL_DISTANCE, ProfileVariables.PRESSURE],
            system_variables=[SystemVariables.PRESSURE])
        else:
            Results = model.tasks.networksimulation.run(profile_variables=[ProfileVariables.VELOCITY_LIQUID, ProfileVariables.TOTAL_DISTANCE, ProfileVariables.PRESSURE],
            system_variables=[SystemVariables.PRESSURE])
        for Pipe_Name in Diam_Dict:
            for branch, profile in Results.profile.items():
                if Pipe_Name in profile['BranchEquipment']:
                    First_Segment = profile['BranchEquipment'].index(Pipe_Name)
                    for Last_Segment in range(First_Segment,len(profile['BranchEquipment'])):
                        if profile['BranchEquipment'][Last_Segment] is not None and Last_Segment != First_Segment:  # Если находит другой "объект"
                            Last_Segment-=1
                            break
                    if model.get_value(Flowline=Pipe_Name, parameter=Parameters.Flowline.DETAILEDMODEL) == True:
                        Pipe_Geom = model.get_geometry(Flowline=Pipe_Name)
                        Pipe_Length = list(Pipe_Geom['MeasuredDistance'])[-1]/1000
                    else: Pipe_Length = model.get_value(Flowline=Pipe_Name, parameter=Parameters.Flowline.MEASUREDDISTANCE)/1000
                    Specific_Drop_Pressure[Pipe_Name] = (profile['Pressure'][First_Segment]-profile['Pressure'][Last_Segment])/(Pipe_Length)    
        Excess = False       
        for Pipe_Name in Disturbance_Counter:   # Зануление старых нарушений перед новой проверкой
            Disturbance_Counter[Pipe_Name] = 0        
        for Source_Name in MD_Source_Dict[Date_Num]:            # Анализ на превышение линейных давлений над устьевыми
            print (Results.node['Pressure'][Source_Name], Source_Pres_Dict[Source_Name][Date_Num],file=Log_File)
            if Results.node['Pressure'][Source_Name] > Source_Pres_Dict[Source_Name][Date_Num]:
                Excess = True
                MaxSDP = 0
                for Pipe_Name in Dict_Path[Source_Name]:
                    print (Pipe_Name, Assortment_Num[Pipe_Name],file=Log_File)
                    if Specific_Drop_Pressure[Pipe_Name] > MaxSDP and (Assortment_Num[Pipe_Name]+1) < len(Pipes_Assortment):    # Проверяется так же, чтобы труба не была последнего типоразмера
                        MaxSDP = Specific_Drop_Pressure[Pipe_Name]
                        MaxSDP_PipeName = Pipe_Name
                Disturbance_Counter[MaxSDP_PipeName] += 1
                if Disturbance_Counter[MaxSDP_PipeName] < 2: Assortment_Num[MaxSDP_PipeName] += 1
                Diam_Dict[MaxSDP_PipeName]['InnerDiameter'] = Pipes_Assortment[Assortment_Num[MaxSDP_PipeName]] - Pipes_Thickness[Assortment_Num[MaxSDP_PipeName]]*2
                Diam_Dict[MaxSDP_PipeName]['WallThickness'] = Pipes_Thickness[Assortment_Num[MaxSDP_PipeName]]
                model.set_values(Diam_Dict)
                print (pd.DataFrame({'Diameter':np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[0]*2+np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[1]}, index = Diam_Dict.keys()),file=Log_File)
        if Excess == False: return Diam_Dict
# ____________________________________________________________________________________________

# Дополнительная проверка скоростей в трубах (чтобы исключить ошибки из-за перераспределения давлений)   
def Additional_Check(model,Diam_Dict,Pipes_Assortment,Pipes_Thickness,Assortment_Num,Log_File):
    model.tasks.networksimulation.reset_conditions()
    if xw.sheets['Spec'].range('D10').value == 'Газ':
        Results = model.tasks.networksimulation.run(profile_variables=[ProfileVariables.VELOCITY_GAS, ProfileVariables.TOTAL_DISTANCE, ProfileVariables.PRESSURE],
        system_variables=[SystemVariables.PRESSURE])
    else:
        Results = model.tasks.networksimulation.run(profile_variables=[ProfileVariables.VELOCITY_LIQUID, ProfileVariables.TOTAL_DISTANCE, ProfileVariables.PRESSURE],
        system_variables=[SystemVariables.PRESSURE])
    for Pipe_Name in Diam_Dict:
        for branch, profile in Results.profile.items():
            if Pipe_Name in profile['BranchEquipment']:
                First_Segment = profile['BranchEquipment'].index(Pipe_Name)
                for Last_Segment in range(First_Segment,len(profile['BranchEquipment'])):
                    if profile['BranchEquipment'][Last_Segment] is not None and Last_Segment != First_Segment:  # Если находит другой "объект"
                        Last_Segment-=1
                        break
                if xw.sheets['Spec'].range('D10').value == 'Газ': Last_Velocity = profile['VelocityGas'][Last_Segment]
                else: Last_Velocity = profile['VelocityLiquid'][Last_Segment]
        print(Pipe_Name,Diam_Dict[Pipe_Name]['InnerDiameter'],Last_Velocity,file=Log_File)
        Crit_Velocity = 20 if xw.sheets['Spec'].range('D10').value == 'Газ' else 3
        if Last_Velocity > Crit_Velocity:
            Assortment_Num[Pipe_Name] += 1
            Diam_Dict[Pipe_Name]['InnerDiameter'] = Pipes_Assortment[Assortment_Num[Pipe_Name]]-Pipes_Thickness[Assortment_Num[Pipe_Name]]*2
            Diam_Dict[Pipe_Name]['WallThickness'] = Pipes_Thickness[Assortment_Num[Pipe_Name]]
            model.set_values(Diam_Dict)
    return Diam_Dict
#_____________________________________________________________________________________________________ 

# Считывает список труб из экселя и добавляет дефолтный диаметр трубам (200мм)______________________
def PipeList_Counter(Pipes_Assortment,Pipes_Thickness,Init_Assort_Num):
    PipeAmount = 0
    Diam_Dict = defaultdict(dict)
    Assortment_Num = defaultdict(dict)
    Disturbance_Counter = defaultdict(dict)
    Bad_Assortments = defaultdict(dict)
    Good_Assortments = defaultdict(dict)
    Accepted_Diam = {}
    while not xw.sheets['Инициал.'].range(PipeAmount+3,9).value is None:
        Diam_Dict[xw.sheets['Инициал.'].range(PipeAmount+3,9).value]['InnerDiameter'] = Pipes_Assortment[Init_Assort_Num]-Pipes_Thickness[Init_Assort_Num]*2
        Diam_Dict[xw.sheets['Инициал.'].range(PipeAmount+3,9).value]['WallThickness'] = Pipes_Thickness[Init_Assort_Num]
        Assortment_Num[xw.sheets['Инициал.'].range(PipeAmount+3,9).value] = Init_Assort_Num
        Disturbance_Counter[xw.sheets['Инициал.'].range(PipeAmount+3,9).value] = 0
        Good_Assortments[xw.sheets['Инициал.'].range(PipeAmount+3,9).value] = 0
        Bad_Assortments[xw.sheets['Инициал.'].range(PipeAmount+3,9).value] = 0
        Accepted_Diam[xw.sheets['Инициал.'].range(PipeAmount+3,9).value] = False
        PipeAmount += 1
    return Diam_Dict, Assortment_Num, Disturbance_Counter, Good_Assortments, Bad_Assortments, Accepted_Diam
# ____________________________________________________________________________________________________

# Подготовка данных: создание словарей с критическими датами для каждой трубы и источника_________________________
def Dicts_Preparation(model,Log_File):
    shti = xw.sheets['Инициал.']
    shts = xw.sheets['Давл.Сеп.']
    shtr = xw.sheets['Газ'] if xw.sheets['Spec'].range('D10').value == 'Газ' else xw.sheets['Нефть']
    shtu = xw.sheets['Давл.Уст.']
    Pipe_Num = 3
    Sep_Num = 2
    Source_Num = 2
    Dict_Path = defaultdict(dict)
    Pipe_Dict = defaultdict(dict)
    MD_Pipe_Dict = defaultdict(dict)
    Sep_Dict = defaultdict(dict)
    Source_Dict = defaultdict(dict)
    Source_Pres_Dict = defaultdict(dict)
    MD_Source_Dict= defaultdict(dict)
# Создает словарь стоков и источников----------------------------------------------------------------------------------
    while shts.range(2,Sep_Num).value is not None:                      # Создает словарь стоков
        Pres_Num = 3
        Pres_List = []
        while shts.range(Pres_Num,Sep_Num).value is not None:
            Pres_List = Pres_List + [shts.range(Pres_Num,Sep_Num).value]
            Pres_Num += 1
        Sep_Dict[shts.range(2,Sep_Num).value] = Pres_List
        Sep_Num += 1     
    while shtr.range(2,Source_Num).value is not None:                   # Создает словарь источников
        Rate_Num = 3
        Rate_List = []
        Source_Pres_List = []
        Diff = []      
        while shtr.range(Rate_Num,Source_Num).value is not None:
            Rate_List = Rate_List + [shtr.range(Rate_Num,Source_Num).value]
            Source_Pres_List = Source_Pres_List + [shtu.range(Rate_Num,Source_Num).value]
            Pres_Value = 1 if shtu.range(Rate_Num,Source_Num).value == 0 else shtu.range(Rate_Num,Source_Num).value         #Убирает нулевые давления, чтобы потом не ругалось при делении
            Diff = Diff + [shtr.range(Rate_Num,Source_Num).value/Pres_Value]
            Rate_Num += 1
        MaxDiff_Source_Num = int(np.array(Diff).argmax()) 
        if MD_Source_Dict[MaxDiff_Source_Num] == {}: MD_Source_Dict[MaxDiff_Source_Num] = [shtr.range(2,Source_Num).value]
        else: MD_Source_Dict[MaxDiff_Source_Num] += [shtr.range(2,Source_Num).value]              #Словарь в форме "крит.год:источник"                              
        Source_Dict[shtr.range(2,Source_Num).value] = Rate_List
        Source_Pres_Dict[shtr.range(2,Source_Num).value] = Source_Pres_List
        Source_Num += 1 
# Подготовка словаря "труба:влияющие источники, стоки"-------------------------------------------------------------------
    Connections = model.connections()
    while shti.range(Pipe_Num,9).value is not None:
        Pipe_Name = shti.range(Pipe_Num,9).value
        Dep_Source_List = []
        Dep_Sep_List = []
        Element_List = [Pipe_Name]
        while 0 == 0:                                           #Создание бесконечного цикла для нахождения связанных источников
            Source_Element = []
            for Connect in Connections:
                if Connect['Destination'] in Element_List:
                    Source_Element = Source_Element + [Connect['Source']]     
            else:
                    Element_List = Source_Element
                    for Element in Element_List:
                        if Element in Source_Dict.keys():
                            Dep_Source_List += [Element]
            if Element_List == []: break
        Element_List = [Pipe_Name]
        while 0 == 0:                                           #Создание бесконечного цикла для нахождения связанных стоков
            Destination_Element = []
            for Connect in Connections:
                if Connect['Source'] in Element_List:
                    Destination_Element = Destination_Element + [Connect['Destination']]
            else:
                    Element_List = Destination_Element
                    for Element in Element_List:
                        if Element in Sep_Dict.keys():
                            Dep_Sep_List += [Element]
            if Element_List == []: break
        Dep_List = [Dep_Source_List] + [Dep_Sep_List]           #Создание словаря в формате "труба:[источники],[стоки]"
        Pipe_Dict[Pipe_Name] = Dep_List
        Pipe_Num += 1
    print(pd.Series(Pipe_Dict),file=Log_File)
#Выяление критических дат для труб, создание словаря "крит.год:труба:источники,стоки---------------------------------------
    for Pipe_Name in Pipe_Dict:
        for Source_Name in Pipe_Dict[Pipe_Name][0]:
            if Pipe_Dict[Pipe_Name][0].index(Source_Name)==0:                           #Попробовать избавиться от условия "если первый - то создать список, если более - добавить"
                Pipe_Rate = np.array(Source_Dict[Source_Name])
            else:
                Pipe_Rate = Pipe_Rate + np.array(Source_Dict[Source_Name]) 
            for Sep_Name in Pipe_Dict[Pipe_Name][1]:
                if Pipe_Dict[Pipe_Name][1].index(Sep_Name)==0:
                    Pipe_Pres = np.array(Sep_Dict[Sep_Name])
                else:
                    Pipe_Pres = Pipe_Pres + np.array(Sep_Dict[Sep_Name]) 
        MaxDiff_Pipe_Num = (Pipe_Rate/Pipe_Pres).argmax()                                #Выявление критической даты(порядковый номер этой даты)
        MD_Pipe_Dict[MaxDiff_Pipe_Num].update({Pipe_Name:Pipe_Dict[Pipe_Name]})             #Приведение в форму "крит.год:труба:источники,стоки"
#Создает словарь "источник: зависимые трубы"----------------------------------------------------------------------------
    for Element_Name in Source_Dict:
        Source_Name = Element_Name
        Path_List = []
        while model.get_connections(Name = Element_Name)[Element_Name]['Destination'] not in Sep_Dict:
            Element_Name = model.get_connections(Name = Element_Name)[Element_Name]['Destination']
            if Element_Name in Pipe_Dict:
                Path_List = Path_List + [Element_Name]
        Dict_Path[Source_Name] = Path_List  
    return MD_Source_Dict, MD_Pipe_Dict, Dict_Path, Source_Pres_Dict
#_______________________________________________________________________________________________________________

# Подбор диаметров трубопроводов________________________________________________________________________________
@entry_point
@xl_macro
@log_error
def Pipe_Diameters_Selection():
    start_time = time.clock()
    shti = xw.sheets['Инициал.']
    shtd = xw.sheets['Диам.']
    Pipe_Num = 3
    model = Model.open(shti.range('A21').value, Units.METRIC)
    Log_File = open(shti.range('A21').value.rpartition('\\')[0]+'\\Log.csv','w')
    Pipes_Assortment = shtd.range('J3').expand('down').value
    Pipes_Thickness = shtd.range('K3').expand('down').value
    Init_Assort_Num = Pipes_Assortment.index(shtd.range('G5').value)
    Diam_Dict, Assortment_Num, Disturbance_Counter, Good_Assortments, Bad_Assortments, Accepted_Diam = PipeList_Counter(Pipes_Assortment,Pipes_Thickness,Init_Assort_Num)
    MD_Source_Dict, MD_Pipe_Dict, Dict_Path, Source_Pres_Dict = Dicts_Preparation(model,Log_File)
    print (pd.Series(Dict_Path),file=Log_File)
    print ('Критические даты для труб',file=Log_File)
    print (MD_Pipe_Dict,file=Log_File)
    #print ('Дата' + i + ':'\n  for i in MD_Pipe_Dict)
    print ('Критические даты для источников',file=Log_File)
    print (MD_Source_Dict,file=Log_File)
    for Crit_Data in MD_Pipe_Dict:
        Date_Num = int(Crit_Data) + 3
        Input_GasRate(model,Date_Num)
        Input_OilRate(model,Date_Num)
        Input_WaterRate(model,Date_Num)
        Input_SourceTemperature(model,Date_Num)
        Input_SinkPressure(model,Date_Num)
        Diam_Dict, Assortment_Num = Find_ID(model,Date_Num-3,Diam_Dict,Pipes_Assortment,Pipes_Thickness,Assortment_Num,Disturbance_Counter,Good_Assortments, Bad_Assortments, Accepted_Diam, MD_Pipe_Dict,Log_File)
    print ('Диаметры труб при первичном подборе',file=Log_File)
    print (pd.DataFrame({'Diameter':np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[0]*2+np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[1]}, index = Diam_Dict.keys()),file=Log_File)
    for Crit_Data in MD_Source_Dict:
        Date_Num = int(Crit_Data) + 3
        Input_GasRate(model,Date_Num)
        Input_OilRate(model,Date_Num)
        Input_WaterRate(model,Date_Num)
        Input_SourceTemperature(model,Date_Num)
        Input_SinkPressure(model,Date_Num)
        Diam_Dict = Correction_ID(model,Date_Num-3,Diam_Dict,Pipes_Assortment,Pipes_Thickness,Assortment_Num,Disturbance_Counter,MD_Source_Dict,Dict_Path,Source_Pres_Dict,Log_File)
    print ('Диаметры труб после корректировки на устьевые давления',file=Log_File)
    print (pd.DataFrame({'Diameter':np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[0]*2+np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[1]}, index = Diam_Dict.keys()),file=Log_File)
    for Crit_Data in MD_Pipe_Dict:
        Date_Num = int(Crit_Data) + 3
        Input_GasRate(model,Date_Num)
        Input_OilRate(model,Date_Num)
        Input_WaterRate(model,Date_Num)
        Input_SourceTemperature(model,Date_Num)
        Input_SinkPressure(model,Date_Num)
        Diam_Dict = Additional_Check(model,Diam_Dict,Pipes_Assortment,Pipes_Thickness,Assortment_Num,Log_File)
    print ('Диаметры труб после дополнительной проверки',file=Log_File)
    print (pd.DataFrame({'Diameter':np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[0]*2+np.array([list(i.values()) for i in Diam_Dict.values()]).transpose()[1]}, index = Diam_Dict.keys()),file=Log_File) 
    while shtd.range(Pipe_Num,1).value is not None:
        shtd.range(Pipe_Num,2).value = Diam_Dict[shtd.range(Pipe_Num,1).value]['InnerDiameter']+Diam_Dict[shtd.range(Pipe_Num,1).value]['WallThickness']*2
        shtd.range(Pipe_Num,2).number_format = 'Основной'
        Pipe_Num += 1
    model.save(shti.range('A21').value)
    model.close()
    print('Время расчета - ' + str(round(time.clock() - start_time, 2)) + "seconds",file=Log_File)
    Log_File.close()
    # ____________________________________________________________________________________________

# Анализ данных ветвей и выгрузка расчетных данных в Excel___________________________________________________    
def Branch_Analysis(model, Date_Num, shts, Branch_Num, Pipe_Name, Results_In_Results, Branch_In_Results):
    if True in shts.range('D2:D3').value: Velocity_Gas_In_Results = dict(Results_In_Results[Branch_Num]).get('VelocityGas')
    if True in shts.range('D4:D5').value: Velocity_Liq_In_Results = dict(Results_In_Results[Branch_Num]).get('VelocityLiquid')
    if shts.range(6,4).value == True: Pressure_In_Results = dict(Results_In_Results[Branch_Num]).get('Pressure')
    if shts.range(8,4).value == True: EVR_In_Results = dict(Results_In_Results[Branch_Num]).get('ErosionalVelocityRatio')
    for Equipment_Num in range(len(Branch_In_Results)):
        if Branch_In_Results[Equipment_Num] in Pipe_Name:
            Max_EVR = 0
            Branch_Name = Branch_In_Results[Equipment_Num]
            if shts.range(2,4).value == True: Inlet_Velocity_Gas = Velocity_Gas_In_Results[Equipment_Num]
            if shts.range(4,4).value == True: Inlet_Velocity_Liq = Velocity_Liq_In_Results[Equipment_Num]
            if shts.range(6,4).value == True: Inlet_Pressure = Pressure_In_Results[Equipment_Num]
            for Equipment_Num in range(Equipment_Num,len(Branch_In_Results)):
                if Branch_In_Results[Equipment_Num] is not None and Branch_In_Results[Equipment_Num]!=Branch_Name:
                    if shts.range(3,4).value == True: Outlet_Velocity_Gas = Velocity_Gas_In_Results[Equipment_Num-1]
                    if shts.range(5,4).value == True: Outlet_Velocity_Liq = Velocity_Liq_In_Results[Equipment_Num-1]
                    if shts.range(6,4).value == True: Outlet_Pressure = Pressure_In_Results[Equipment_Num-1]
                    break
                elif Equipment_Num+1==len(Branch_In_Results):
                    if shts.range(3,4).value == True: Outlet_Velocity_Gas = Velocity_Gas_In_Results[Equipment_Num]
                    if shts.range(5,4).value == True: Outlet_Velocity_Liq = Velocity_Liq_In_Results[Equipment_Num]
                    if shts.range(6,4).value == True: Outlet_Pressure = Pressure_In_Results[Equipment_Num]
                    if shts.range(8,4).value == True: Max_EVR = max(EVR_In_Results[Equipment_Num],Max_EVR)
                    break
                if shts.range(8,4).value == True: Max_EVR = max(EVR_In_Results[Equipment_Num],Max_EVR)
            for Pipe_Row in range(len(Pipe_Name)+3):
                if xw.sheets['Инициал.'].range(3+Pipe_Row,9).value == Branch_Name:
                    print (Branch_Name)
                    if shts.range(6,4).value == True: 
                        if model.get_value(Flowline=Branch_Name, parameter=Parameters.Flowline.DETAILEDMODEL) == True:
                            Pipe_Geom = model.get_geometry(Flowline=Branch_Name)
                            Pipe_Length = list(Pipe_Geom['MeasuredDistance'])[-1]/1000
                        else: Pipe_Length = model.get_value(Flowline=Branch_Name, parameter=Parameters.Flowline.MEASUREDDISTANCE)/1000
                        print (Inlet_Pressure, Outlet_Pressure, Pipe_Length)
                        Specific_Drop_Pressure = (Inlet_Pressure - Outlet_Pressure)/Pipe_Length
                        xw.sheets['Уд.Пер.Давл.'].range(3+Pipe_Row,Date_Num).value = Specific_Drop_Pressure
                        xw.sheets['Уд.Пер.Давл.'].range(3+Pipe_Row,Date_Num).number_format = '0.00'
                        if Specific_Drop_Pressure > 2: 
                            xw.sheets['Уд.Пер.Давл.'].range(3+Pipe_Row,Date_Num).color = (255,150,150)
                        else: 
                            xw.sheets['Уд.Пер.Давл.'].range(3+Pipe_Row,Date_Num).color = (180,230,180)
                    if shts.range(2,4).value == True:
                        xw.sheets['Нач.Скор.(Газ)'].range(3+Pipe_Row,Date_Num).value = Inlet_Velocity_Gas
                        xw.sheets['Нач.Скор.(Газ)'].range(3+Pipe_Row,Date_Num).number_format = '0.00'
                        if Inlet_Velocity_Gas < 2 or Inlet_Velocity_Gas > 20:
                            xw.sheets['Нач.Скор.(Газ)'].range(3+Pipe_Row,Date_Num).color = (255,150,150)
                        else:
                            xw.sheets['Нач.Скор.(Газ)'].range(3+Pipe_Row,Date_Num).color = (180,230,180)
                    if shts.range(3,4).value == True:
                        xw.sheets['Кон.Скор.(Газ)'].range(3+Pipe_Row,Date_Num).value = Outlet_Velocity_Gas
                        xw.sheets['Кон.Скор.(Газ)'].range(3+Pipe_Row,Date_Num).number_format = '0.00'
                        if Outlet_Velocity_Gas < 2 or Outlet_Velocity_Gas > 20:
                            xw.sheets['Кон.Скор.(Газ)'].range(3+Pipe_Row,Date_Num).color = (255,150,150)
                        else:
                            xw.sheets['Кон.Скор.(Газ)'].range(3+Pipe_Row,Date_Num).color = (180,230,180)
                    if shts.range(4,4).value == True:
                        xw.sheets['Нач.Скор.(Жид.)'].range(3+Pipe_Row,Date_Num).value = Inlet_Velocity_Liq
                        xw.sheets['Нач.Скор.(Жид.)'].range(3+Pipe_Row,Date_Num).number_format = '0.00'
                        if Inlet_Velocity_Liq > 3:
                            xw.sheets['Нач.Скор.(Жид.)'].range(3+Pipe_Row,Date_Num).color = (255,150,150)
                        else:
                            xw.sheets['Нач.Скор.(Жид.)'].range(3+Pipe_Row,Date_Num).color = (180,230,180)
                    if shts.range(5,4).value == True:
                        xw.sheets['Кон.Скор.(Жид.)'].range(3+Pipe_Row,Date_Num).value = Outlet_Velocity_Liq
                        xw.sheets['Кон.Скор.(Жид.)'].range(3+Pipe_Row,Date_Num).number_format = '0.00'
                        if Outlet_Velocity_Liq > 3:
                            xw.sheets['Кон.Скор.(Жид.)'].range(3+Pipe_Row,Date_Num).color = (255,150,150)
                        else:
                            xw.sheets['Кон.Скор.(Жид.)'].range(3+Pipe_Row,Date_Num).color = (180,230,180)
                    if shts.range(8,4).value == True:
                        xw.sheets['Отн.Скор.Эр.'].range(3+Pipe_Row,Date_Num).value = Max_EVR
                        xw.sheets['Отн.Скор.Эр.'].range(3+Pipe_Row,Date_Num).number_format = '0.0'
                        if Max_EVR > 1:
                            xw.sheets['Отн.Скор.Эр.'].range(3+Pipe_Row,Date_Num).color = (255,150,150)
                        else:
                            xw.sheets['Отн.Скор.Эр.'].range(3+Pipe_Row,Date_Num).color = (180,230,180)
                    
                    break    
# ____________________________________________________________________________________________

# Выгрузка прогнозных данных по всем датам__________________________________________________
@entry_point
@xl_macro
@log_error
def Prediction_Calculations_All():
    shti = xw.sheets['Инициал.']
    shts = xw.sheets['Spec']
    shtp = xw.sheets['Разн.Уст.Давл.']
    shtt = xw.sheets['Темп.Точ.']
    model = Model.open(shti.range('A21').value, Units.METRIC)
    Date_Num = 3
    Pipe_Name = np.array(model.find(Flowline=ALL))[:, None]
    xw.sheets['Нач.Скор.(Газ)'].range('B3:XF1048').clear()
    xw.sheets['Кон.Скор.(Газ)'].range('B3:XF1048').clear()
    xw.sheets['Уд.Пер.Давл.'].range('B3:XF1048').clear()
    xw.sheets['Разн.Уст.Давл.'].range('B3:XF1048').clear()
    xw.sheets['Отн.Скор.Эр.'].range('B3:XF1048').clear()
    Profile_Variables = System_Variables = []
    if True in shts.range('D2:D3').value: Profile_Variables = Profile_Variables + [ProfileVariables.VELOCITY_GAS]
    if True in shts.range('D4:D5').value: Profile_Variables = Profile_Variables + [ProfileVariables.VELOCITY_LIQUID]
    if shts.range(6,4).value == True: Profile_Variables = Profile_Variables + [ProfileVariables.PRESSURE]
    if shts.range(8,4).value == True: Profile_Variables = Profile_Variables + [ProfileVariables.EROSIONAL_VELOCITY_RATIO]
    if shts.range(7,4).value == True: System_Variables = System_Variables + [SystemVariables.PRESSURE]
    if shts.range(11, 4).value == True: System_Variables = System_Variables + [SystemVariables.TEMPERATURE]
    while shti.range(Date_Num,8).value is not None:  
        Input_GasRate(model,Date_Num)
        Input_OilRate(model,Date_Num)
        Input_WaterRate(model,Date_Num)
        Input_SourceTemperature(model,Date_Num)
        Input_SinkPressure(model,Date_Num)
        if shts.range(1,4).value == 'ИСТИНА': input_ChokeDiameter(model,Date_Num)
        model.tasks.networksimulation.reset_conditions()
        Results = model.tasks.networksimulation.run(profile_variables=Profile_Variables, system_variables=System_Variables)
        if shts.range(7,4).value == True:
            Nodes_In_Results = Results.node.get('Pressure')
            for Source_Num in range(len(shtp.range('B2').expand('right'))):
                Source_Name = shtp.range(2,Source_Num+2).value
                try:
                    Source_Pres = 0 if Nodes_In_Results.get(Source_Name) == None else Nodes_In_Results.get(Source_Name)
                except AttributeError:
                    Source_Pres = 0
                shtp.range(Date_Num,Source_Num+2).value = xw.sheets['Давл.Уст.'].range(Date_Num,Source_Num+2).value - Source_Pres
                shtp.range(Date_Num,Source_Num+2).number_format = '0.00'
                if shtp.range(Date_Num,Source_Num+2).value < 0:
                    shtp.range(Date_Num,Source_Num+2).color = (255,150,150)
                else: 
                    shtp.range(Date_Num,Source_Num+2).color = (180,230,180)
        if shts.range(11, 4).value == True:
            Nodes_Name = list(Results.node.get('Temperature').keys())[1:]
            Nodes_Temp = []
            Nodes_Name.sort()
            for i in Nodes_Name:
                Nodes_Temp = Nodes_Temp + [Results.node.get('Temperature')[i]]
            shtt.range(Date_Num, 2).value = Nodes_Temp
            shtt.range(2, 2).value = Nodes_Name
        Results_In_Results = list(Results.profile.values())
        for Branch_Num in range(len(Results_In_Results)):
            Branch_In_Results = dict(Results_In_Results[Branch_Num]).get('BranchEquipment')
            Branch_Analysis(model, Date_Num-1, shts, Branch_Num, Pipe_Name, Results_In_Results, Branch_In_Results)
        Date_Num += 1
    model.save(shti.range('A21').value)
    model.close()
# ____________________________________________________________________________

# Выгрузка прогнозных данных по диапазону дат__________________________________________________
@entry_point
@xl_macro
@log_error
def Prediction_Calculations_DateRange():
    shti = xw.sheets['Инициал.']
    shts = xw.sheets['Spec']
    shtp = xw.sheets['Разн.Уст.Давл.']
    shtt = xw.sheets['Темп.Точ.']
    model = Model.open(shti.range('A21').value, Units.METRIC)
    Pipe_Name = np.array(model.find(Flowline=ALL))[:, None]
    xw.sheets['Нач.Скор.(Газ)'].range('B3:XF1048').clear()
    xw.sheets['Кон.Скор.(Газ)'].range('B3:XF1048').clear()
    xw.sheets['Уд.Пер.Давл.'].range('B3:XF1048').clear()
    xw.sheets['Разн.Уст.Давл.'].range('B3:XF1048').clear()
    xw.sheets['Отн.Скор.Эр.'].range('B3:XF1048').clear()
    Profile_Variables = System_Variables = []
    if True in shts.range('D2:D3').value: Profile_Variables = Profile_Variables + [ProfileVariables.VELOCITY_GAS]
    if True in shts.range('D4:D5').value: Profile_Variables = Profile_Variables + [ProfileVariables.VELOCITY_LIQUID]
    if shts.range(6,4).value == True: Profile_Variables = Profile_Variables + [ProfileVariables.PRESSURE]
    if shts.range(8,4).value == True: Profile_Variables = Profile_Variables + [ProfileVariables.EROSIONAL_VELOCITY_RATIO]
    if shts.range(7,4).value == True: System_Variables = System_Variables + [SystemVariables.PRESSURE]
    if shts.range(11, 4).value == True: System_Variables = System_Variables + [SystemVariables.TEMPERATURE]
    First_Date = shti.range('H3').expand('down').value.index(shti.range('C26').value)
    Last_Date = shti.range('H3').expand('down').value.index(shti.range('D26').value)
    for Date_Num in range(First_Date+3, Last_Date+4):
        Input_GasRate(model,Date_Num)
        Input_OilRate(model,Date_Num)
        Input_WaterRate(model,Date_Num)
        Input_SourceTemperature(model,Date_Num)
        Input_SinkPressure(model,Date_Num)
        if shts.range(1,4).value == 'ИСТИНА': input_ChokeDiameter(model,Date_Num)
        model.tasks.networksimulation.reset_conditions()
        Results = model.tasks.networksimulation.run(profile_variables=Profile_Variables, system_variables=System_Variables)
        if shts.range(7,4).value == True:
            Nodes_In_Results = Results.node.get('Pressure')
            for Source_Num in range(len(shtp.range('B2').expand('right'))):
                Source_Name = shtp.range(2,Source_Num+2).value
                try:
                    Source_Pres = 0 if Nodes_In_Results.get(Source_Name) == None else Nodes_In_Results.get(Source_Name)
                except AttributeError:
                    Source_Pres = 0
                shtp.range(Date_Num,Source_Num+2).value = xw.sheets['Давл.Уст.'].range(Date_Num,Source_Num+2).value - Source_Pres
                shtp.range(Date_Num,Source_Num+2).number_format = '0.00'
                if shtp.range(Date_Num,Source_Num+2).value < 0: 
                    shtp.range(Date_Num,Source_Num+2).color = (255,150,150)
                else: 
                    shtp.range(Date_Num,Source_Num+2).color = (180,230,180)
        if shts.range(11, 4).value == True:
            Nodes_Name = list(Results.node.get('Temperature').keys())[1:]
            Nodes_Temp = []
            Nodes_Name.sort()
            for i in Nodes_Name:
                Nodes_Temp = Nodes_Temp + [Results.node.get('Temperature')[i]]
            shtt.range(Date_Num, 2).value = Nodes_Temp
            shtt.range(2, 2).value = Nodes_Name
        Results_In_Results = list(Results.profile.values())
        for Branch_Num in range(len(Results_In_Results)):
            Branch_In_Results = dict(Results_In_Results[Branch_Num]).get('BranchEquipment')
            Branch_Analysis(model, Date_Num-1, shts, Branch_Num, Pipe_Name, Results_In_Results, Branch_In_Results)
    model.save(shti.range('A21').value)
    model.close()
# ______________________________________________________________________________________________________________

# Выгрузка прогнозных данных по одной дате______________________________________________________________________
@entry_point
@xl_macro
@log_error
def Prediction_Calculations_OneDate():
    shti = xw.sheets['Инициал.']
    shts = xw.sheets['Spec']
    shtp = xw.sheets['Разн.Уст.Давл.']
    shtt = xw.sheets['Темп.Точ.']
    model = Model.open(shti.range('A21').value, Units.METRIC)
    Pipe_Name = np.array(model.find(Flowline=ALL))[:, None]
    xw.sheets['Нач.Скор.(Газ)'].range('B3:XF1048').clear()
    xw.sheets['Кон.Скор.(Газ)'].range('B3:XF1048').clear()
    xw.sheets['Уд.Пер.Давл.'].range('B3:XF1048').clear()
    xw.sheets['Разн.Уст.Давл.'].range('B3:XF1048').clear()
    xw.sheets['Отн.Скор.Эр.'].range('B3:XF1048').clear()
    Profile_Variables = System_Variables = []
    if True in shts.range('D2:D3').value: Profile_Variables = Profile_Variables + [ProfileVariables.VELOCITY_GAS]
    if True in shts.range('D4:D5').value: Profile_Variables = Profile_Variables + [ProfileVariables.VELOCITY_LIQUID]
    if shts.range(6,4).value == True: Profile_Variables = Profile_Variables + [ProfileVariables.PRESSURE]
    if shts.range(8,4).value == True: Profile_Variables = Profile_Variables + [ProfileVariables.EROSIONAL_VELOCITY_RATIO]
    if shts.range(7,4).value == True: System_Variables = System_Variables + [SystemVariables.PRESSURE]
    if shts.range(11, 4).value == True: System_Variables = System_Variables + [SystemVariables.TEMPERATURE]
    Date_Num = 3 + shti.range('H3').expand('down').value.index(shti.range('C27').value)
    Input_GasRate(model,Date_Num)
    Input_OilRate(model,Date_Num)
    Input_WaterRate(model,Date_Num)
    Input_SourceTemperature(model,Date_Num)
    Input_SinkPressure(model,Date_Num)
    if shts.range(1,4).value == True: input_ChokeDiameter(model,Date_Num)
    model.tasks.networksimulation.reset_conditions()
    Results = model.tasks.networksimulation.run(profile_variables=Profile_Variables, system_variables=System_Variables)
    if shts.range(7,4).value == True:
        Nodes_In_Results = Results.node.get('Pressure')
        for Source_Num in range(len(shtp.range('B2').expand('right'))):
            Source_Name = shtp.range(2,Source_Num+2).value
            try:
                Source_Pres = 0 if Nodes_In_Results.get(Source_Name) == None else Nodes_In_Results.get(Source_Name)
            except AttributeError:
                Source_Pres = 0
            shtp.range(Date_Num,Source_Num+2).value = xw.sheets['Давл.Уст.'].range(Date_Num,Source_Num+2).value - Source_Pres
            shtp.range(Date_Num,Source_Num+2).number_format = '0.00'
            if shtp.range(Date_Num,Source_Num+2).value < 0: 
                shtp.range(Date_Num,Source_Num+2).color = (255,150,150)
            else: 
                shtp.range(Date_Num,Source_Num+2).color = (180,230,180)
    if shts.range(11, 4).value == True:
        Nodes_Name = list(Results.node.get('Temperature').keys())[1:]
        Nodes_Temp = []
        Nodes_Name.sort()
        for i in Nodes_Name:
            Nodes_Temp = Nodes_Temp + [Results.node.get('Temperature')[i]]
        shtt.range(Date_Num, 2).value = Nodes_Temp
        shtt.range(2, 2).value = Nodes_Name
    Results_In_Results = list(Results.profile.values())
    for Branch_Num in range(len(Results_In_Results)):
        Branch_In_Results = dict(Results_In_Results[Branch_Num]).get('BranchEquipment')
        Branch_Analysis(model, Date_Num-1, shts, Branch_Num, Pipe_Name, Results_In_Results, Branch_In_Results)
    model.save(shti.range('A21').value)
    model.close()
# ______________________________________________________________________________________________________________
Prediction_Calculations_All()
#Pipe_Diameters_Selection()
#Prediction_Calculations_DateRange()
#Prediction_Calculations_OneDate()
#shtg = xw.sheets['Газ']
#model = Model.open(shti.range('A21').value, Units.METRIC)
#Source_Dict = defaultdict(dict)
#Source_List = np.transpose(np.array(shtg.range(2,2).expand().value))
#for Date_Num in range(len(Source_List[0])-1):
#    Source_Num = 0
#    for Source_Name in Source_List[:,0]:
#        Source_Dict[Date_Num] = {Source_Name:{'GasFlowRate':list(Source_List[Source_Num][Date_Num+1])}}
#        Source_Num += 1
#print (Source_Dict)
#Source_Dict = {}
#print (for {Source_Name for Source_Name in shtg.range(2,2).expand('right').value}) 
#print (Source_Dict)
#while shtg.range(2,Source_Num).value is not None:
#    Source_Name = shtg.range(2,Source_Num).value
#    Source_Value = shtg.range(Date_Num,Source_Num).value
#    if shtg.range(1,1).value == 'Дебит газа, тыс.м3/сут':
#        if Source_Value == 0: Source_Value = 0.1
#        model.set_value(Source=Source_Name, parameter=Parameters.Source.USEFLUIDOVERRIDES, value=False)
#        model.set_value(Source=Source_Name, parameter=Parameters.Source.SELECTEDRATETYPE, value=Constants.FlowRateType.GASFLOWRATE)
#        model.set_value(Source=Source_Name, parameter=Parameters.Source.GASFLOWRATE, value=Source_Value/1000)
#    else:
#        model.set_value(Source=Source_Name, parameter=Parameters.Source.USEFLUIDOVERRIDES, value=True)
#        model.set_value(Source=Source_Name, parameter=Parameters.Source.USEGASRATIO, value='GOR')
#        model.set_value(Source=Source_Name, parameter=Parameters.Source.GOR, value=Source_Value)           
#    Source_Num += 1
