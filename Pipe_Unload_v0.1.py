# -*- coding: cp1251 -*-
from matplotlib import pylab
import xlwings as xw
from xlwings.constants import InsertShiftDirection
import numpy as np
import pandas as pd
from pyxll import xl_macro
from collections import defaultdict
from enpyxll.util.logs import log_error
from sixgill.pipesim import Model
from sixgill.definitions import Parameters, Units, ALL
from enpyxll import entry_point

# ______________________________________________________________________________________________________________
# Формирование объемов обустройства____________________________________________________________________________
@entry_point
@xl_macro
@log_error
def Infrastructure_Formation_Volumes():
    sht_foo = xw.sheets['ФОО']
    Prediction_Book = xw.Book(sht_foo.range(1, 2).value)
    Fluid_Type = Prediction_Book.sheets['Spec'].range(10, 4).value
    sht_rate = Prediction_Book.sheets[Fluid_Type]
    # Создание словаря "источник : первый год работы" -----------------------------------------
    Source_Dict = defaultdict(dict)
    Dict_Path = defaultdict(dict)
    Rate_List = np.array(sht_rate.range('B3').expand().value)
    Source_List = sht_rate.range('B2').expand('right').value
    for Source in range(len(Source_List)):
        for Date_Num in range(len(Rate_List[:, 0])):
            if float(Rate_List[:, Source][Date_Num]) > 0:
                Source_Dict[Source_List[Source]] = Date_Num
                break
    # Создание словаря "труба : типоразмер, длина" --------------------------------------------
    model = Model.open(Prediction_Book.sheets['Инициал.'].range('A21').value, Units.METRIC)
    Pipe_List = model.find(Flowline=ALL)
    Pipe_Param = defaultdict(dict)
    for Pipe_Name in Pipe_List:
        Pipe_Diam = model.get_values(Flowline=Pipe_Name, parameters=['InnerDiameter', 'WallThickness'])
        if model.get_value(Flowline=Pipe_Name, parameter=Parameters.Flowline.DETAILEDMODEL) == True:
            Pipe_Geom = model.get_geometry(Flowline=Pipe_Name)
            Pipe_Length = list(Pipe_Geom['MeasuredDistance'])[-1] / 1000
        else:
            Pipe_Length = model.get_value(Flowline=Pipe_Name, parameter=Parameters.Flowline.MEASUREDDISTANCE) / 1000
        Pipe_Param[Pipe_Name][str(round(Pipe_Diam[Pipe_Name]['InnerDiameter']+Pipe_Diam[Pipe_Name]['WallThickness']*2)) + 'x' + str(round(Pipe_Diam[Pipe_Name]['WallThickness']))] = Pipe_Length
    # Создание словаря "номер года ввода : диаметры : протяженности----------------------------
    Pipe_Dict = defaultdict(dict)
    Connections = model.connections()
    for Pipe_Name in Pipe_List:
        Dep_Source_List = []
        Element_List = [Pipe_Name]
        while 0 == 0:  # Создание бесконечного цикла для нахождения связанных источников
            Source_Element = []
            for Connect in Connections:
                if Connect['Destination'] in Element_List:
                    Source_Element = Source_Element + [Connect['Source']]
            else:
                Element_List = Source_Element
                for Element in Element_List:
                    if Element in Source_Dict:
                        Dep_Source_List += [Source_Dict[Element]]
            if Element_List == []: break
        if min(Dep_Source_List) in Pipe_Dict: # Костыль для создания словаря
            Pipe_Diam = list(Pipe_Param[Pipe_Name].keys())[0]
            if Pipe_Diam in Pipe_Dict[min(Dep_Source_List)]:
                Pipe_Dict[min(Dep_Source_List)][Pipe_Diam] += list(Pipe_Param[Pipe_Name].values())[0]
            else:
                Pipe_Dict[min(Dep_Source_List)][Pipe_Diam] = list(Pipe_Param[Pipe_Name].values())[0]
        else:
            Pipe_Dict[min(Dep_Source_List)] = Pipe_Param[Pipe_Name]
    Diam_List = []
    for Date_Num in Pipe_Dict:
        Diam_List += Pipe_Dict[Date_Num]
    Diam_List = sorted(set(Diam_List), key=lambda OD: int(OD[:OD.find('x')]))
    Full_List = np.empty([len(Diam_List), 6 + max(list(Pipe_Dict.keys()))], dtype=np.dtype('U100'))
    First_Row = 3 + sht_foo.range('B3').expand('down').value.index('Промысловые трубопроводы')
    for Diam in range(len(Diam_List)):
        sht_foo.range("A5:CV5").api.Insert(InsertShiftDirection.xlShiftDown)
        Full_List[Diam][0] = '2.' + str(Diam+1)
        Full_List[Diam][1] = 'Газосбор ' + Diam_List[Diam]
        Full_List[Diam][3] = 'км'
        Sum_Length = 0
        for Date_Num in Pipe_Dict:
            if Diam_List[Diam] in list(Pipe_Dict[Date_Num].keys()):
                Full_List[Diam][5+Date_Num] = Pipe_Dict[Date_Num][Diam_List[Diam]]
                Sum_Length += Pipe_Dict[Date_Num][Diam_List[Diam]]
        Full_List[Diam][2] = Sum_Length
    sht_foo.range(First_Row + 1, 1).value = Full_List
    Prediction_Book.app.quit()



# ______________________________________________________________________________________________________________
# if xw.sheets['Spec'].range('D10').value == 'Нефть':
#    Oil_Rate_Massive = shto.range(3,2).expand().value
#    Max_Oil_Rate = max(np.sum(Oil_Rate_Massive,1)*365/1000000)
# xw.sheets['Диам.'].range(1,14).value = Max_Oil_Rate
# xw.sheets['Диам.'].range(2,14).value = Pipe_Dict
