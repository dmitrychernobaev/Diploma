# -*- coding: utf-8 -*-
from ast import Break
from calendar import c
import sys
import os
from xml.etree import ElementTree
import pythoncom
from win32com.client import Dispatch, gencache
import LDefin2D
import MiscellaneousHelpers as MH
import time
import random

def start_kompas():
    Dispatch("KOMPAS.Application.7")
    #  Подключим константы API Компас
    kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

    #  Подключим описание интерфейсов API5
    kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
    MH.iKompasObject  = kompas_object

    #  Подключим описание интерфейсов API7
    kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))
    MH.iApplication  = application

    Documents = application.Documents
    return kompas6_constants_3d, kompas_object, kompas_api7_module, Documents, kompas6_constants, kompas6_api5_module, kompas_api7_module

def open(name_m3d, kompas6_constants_3d, kompas_object, kompas_api7_module, Documents):
    #  Открываем документ
    kompas_document = Documents.Open("X://education//Diploma//Program//source//"+name_m3d+".m3d", True, False)

    kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
    iDocument3D = kompas_object.ActiveDocument3D()
    ##
    iPart7 = kompas_document_3d.TopPart
    iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

    VariableCollection = iPart.VariableCollection() #Получение колекции перменных
    VariableCollection.refresh() #обновление коллекции перменных
    return kompas_document, VariableCollection, iPart, kompas_document_3d

def i_enter_filename():
    flag1 = True
    flag2 = True
    while flag1:
        name_m3d=input('Введите название файла, который вы бы хотели изменить\n')
        if os.path.exists('X:\\education\\Diploma\\Program\\source\\'+name_m3d+'.m3d'):
            flag1 = False
        else:
            print('Выбраного файла нет в папке')

    while flag2:
        name_xml=input('Введите название файла, в котором хранится информация о переменных\n')
        if os.path.exists('X:\\education\\Diploma\\Program\\source_xml\\'+name_xml+'.xml'):
            flag2 = False
        else:
            print('Выбраного файла нет в папке')
    
    return name_m3d, name_xml

def a_enter_filename(m3d, xml):
    flag1 = True
    flag2 = True
    while flag1:
        name_m3d = m3d[1]
        if os.path.exists('X:\\education\\Diploma\\Program\\source\\'+name_m3d+'.m3d'):
            flag1 = False
        else:
            print('Выбраного файла m3d нет в папке')
            time.sleep(5)
            quit()

    while flag2:
        name_xml=xml[1]
        if os.path.exists('X:\\education\\Diploma\\Program\\source_xml\\'+name_xml+'.xml'):
            flag2 = False
        else:
            print('Выбраного файла xml нет в папке')
            time.sleep(5)
            quit()

    
    return name_m3d, name_xml

def elem_is(name): 
    print('В выбранной вами фигуре вам доступны следующие элементы для изменения:')
    tree = ElementTree.parse("X:\\education\\Diploma\\Program\\source_xml\\"+name+".xml")
    root = tree.getroot()
    for element in root.iter('variable'):
            for child in element.iter('name'):
                print(' ', child.text)
    return tree

def a_tree(name):
    tree = ElementTree.parse("X:\\education\\Diploma\\Program\\source_xml\\"+name+".xml")
    return tree 

def show_elem(chosen_elem): #Вызов характеристик переменной на экран
    for element in root.iter('variable'):
        for child in element.iter('name'):
            if child.text == chosen_elem:
                print('\nВыбранный элемент', element[0].text)
                print('Обозначение', element[1].text)
                print('Минимальное значение', element[2].text)
                print('Максимальное значение', element[3].text)
                print('Шаг', element[4].text)

def name_elem(): #Ввод пользователем имени переменной
    flag = False
    while flag == False:
        input_elem = input('\nПожалуйста выберите один из представленных элементов, чтобы изменить его: \n')
        for element in root.iter('name'):
            if element.text != input_elem:
                flag = False
            else:
                flag = True
                return input_elem
        if flag == False:
            print('\nВы ввели неправильную переменную')

def a_name_elem(func, root): #Ввод пользователем имени переменной
    input_elem = func[2]
    flag = False
    while flag == False:
        for element in root.iter('name'):
            if element.text != input_elem:
                flag = False
            else:
                flag = True
                return input_elem
        if flag == False:
            print('\nВы ввели неправильную переменную')
            time.sleep(5)
            quit()



def check_elem(chosen_elem): #Проверка вользователем переменной
    flag = input('Это нужная вам переменная?(Да/Нет)\n')
    if flag == 'Нет':
        chosen_elem = select_elem()
        show_elem(elem)
    return chosen_elem

def select_elem(): #Полный выбор переменной
    elem = name_elem()
    show_elem(elem)
    elem = check_elem(elem)
    return elem

def select_parent(elem, root): #Получение всех данных о выбранной перменной
    for element in root.iter('variable'): 
        for child in element.iter('name'):
            if elem == child.text:
                return element

def change_variable(parent, kompas_document, VariableCollection, iPart):
    Variable = VariableCollection.GetByName(parent[1].text, True, True)
    print('\nСтарое значение переменной: ', Variable.value)
    print('Если введенное вами значение не будет соответствовать шагу, то значение округлится в меньшую сторону')
    new_value = float(input('Введите новое значение для выбранной вами переменной '+ elem + ': '))
    new_value = new_value//float(parent[4].text)*float(parent[4].text)

    if new_value > float(parent[3].text):
        print('Ваше значение было слишком большим')
        new_value = float(parent[3].text)
    elif new_value < float(parent[2].text):
        print('Ваше значение было слишком маленьким')
        new_value = float(parent[2].text)
    Variable.value = new_value
    print('Новое значение переменной: ', Variable.value)
    iPart.RebuildModel() 
    kompas_document.Save()

def a_change_variable_rnd(parent, kompas_document, VariableCollection, iPart):
    Variable = VariableCollection.GetByName(parent[1].text, True, True)
    print('\nСтарое значение переменной: ', Variable.value)
    new_value = float(random.randrange(int(parent[2].text), int(parent[3].text), int(parent[4].text)))

    Variable.value = new_value
    print('Новое значение переменной: ', Variable.value)
    time.sleep(5)
    iPart.RebuildModel() 
    kompas_document.Save()
    return new_value

def a_change_variable_for(parent, kompas_document, VariableCollection, iPart, func,
                          Documents, kompas_document_3d, kompas6_constants, 
                          kompas6_api5_module, kompas_api7_module, kompas_object, name_m3d):
    Variable = VariableCollection.GetByName(parent[1].text, True, True)
    print('\nСтарое значение переменной: ', Variable.value)
    Variable.value = 0
    for i in range(int(func[3]), int(func[4]), int(func[5])):
        Variable.value = float(i)
        print('Новое значение переменной: ', Variable.value)
        iPart.RebuildModel() 
        kompas_document.Save()
        a_create_cdw(Documents, kompas_document_3d, kompas6_constants, 
                     kompas6_api5_module, kompas_api7_module, kompas_object, name_m3d,
                     func[2], Variable.value)
    kompas_document_3d.Close(True)


def create_cdw(Documents, kompas_3d_document, kompas6_constants, 
                kompas6_api5_module, kompas_api7_module, kompas_object, name_m3d):
    #  Создаем новый документ
    kompas_document = Documents.AddWithDefaultSettings(kompas6_constants.ksDocumentDrawing, True)

    kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
    iDocument2D = kompas_object.ActiveDocument2D()

    iAssociationViewParam = kompas6_api5_module.ksAssociationViewParam(kompas_object.GetParamStruct(kompas6_constants.ko_AssociationViewParam))
    iAssociationViewParam.Init()
    iAssociationViewParam.disassembly = False
    iAssociationViewParam.fileName = 'C://Users//1//Desktop//'+name_m3d+'.m3d'
    iAssociationViewParam.hiddenLinesShow = False
    iAssociationViewParam.hiddenLinesStyle = 4
    iAssociationViewParam.projBodies = True
    iAssociationViewParam.projectionLink = False
    iAssociationViewParam.projectionName = "#Сверху"
    iAssociationViewParam.projSurfaces = False
    iAssociationViewParam.projThreads = True
    iAssociationViewParam.sameHatch = False
    iAssociationViewParam.section = False
    iAssociationViewParam.tangentEdgesShow = False
    iAssociationViewParam.tangentEdgesStyle = 2
    iAssociationViewParam.visibleLinesStyle = 1
    iViewParam = kompas6_api5_module.ksViewParam(iAssociationViewParam.GetViewParam())
    iViewParam.Init()
    iViewParam.angle = 0
    iViewParam.color = 0
    iViewParam.name = "Вид 1"
    iViewParam.scale_ = 1
    iViewParam.state = 3
    iViewParam.x = 106.708180693419
    iViewParam.y = 210.703411580158
    iDocument2D.ksCreateSheetArbitraryView(iAssociationViewParam, 0)
    
    new_name = 'X:\\education\\Diploma\\Program\\New_file\\'+input('Введите название файла, в который вы бы хотели сохранить изменения\n')+'.cdw'
    kompas_document.SaveAs(new_name)
    kompas_document.Close(True)
    kompas_3d_document.Close(True)

def a_create_cdw(Documents, kompas_3d_document, kompas6_constants, 
                kompas6_api5_module, kompas_api7_module, kompas_object, name_m3d,
                name_value, new_value):
    #  Создаем новый документ
    kompas_document = Documents.AddWithDefaultSettings(kompas6_constants.ksDocumentDrawing, True)

    kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
    iDocument2D = kompas_object.ActiveDocument2D()

    iAssociationViewParam = kompas6_api5_module.ksAssociationViewParam(kompas_object.GetParamStruct(kompas6_constants.ko_AssociationViewParam))
    iAssociationViewParam.Init()
    iAssociationViewParam.disassembly = False
    iAssociationViewParam.fileName = 'C://Users//1//Desktop//'+name_m3d+'.m3d'
    iAssociationViewParam.hiddenLinesShow = False
    iAssociationViewParam.hiddenLinesStyle = 4
    iAssociationViewParam.projBodies = True
    iAssociationViewParam.projectionLink = False
    iAssociationViewParam.projectionName = "#Сверху"
    iAssociationViewParam.projSurfaces = False
    iAssociationViewParam.projThreads = True
    iAssociationViewParam.sameHatch = False
    iAssociationViewParam.section = False
    iAssociationViewParam.tangentEdgesShow = False
    iAssociationViewParam.tangentEdgesStyle = 2
    iAssociationViewParam.visibleLinesStyle = 1
    iViewParam = kompas6_api5_module.ksViewParam(iAssociationViewParam.GetViewParam())
    iViewParam.Init()
    iViewParam.angle = 0
    iViewParam.color = 0
    iViewParam.name = "Вид 1"
    iViewParam.scale_ = 1
    iViewParam.state = 3
    iViewParam.x = 106.708180693419
    iViewParam.y = 210.703411580158
    iDocument2D.ksCreateSheetArbitraryView(iAssociationViewParam, 0)
    
    new_name = 'X:\\education\\Diploma\\Program\\New_file\\'+name_m3d+'_'+name_value+str(new_value)+'.cdw'
    kompas_document.SaveAs(new_name)
    kompas_document.Close(True)  
    
def auto(a):
    xml = sys.argv[2].split('=')
    m3d = sys.argv[3].split('=')
    func = sys.argv[4].split(':')

    print(xml)
    print(m3d)
    print(func)
    if (xml[0] != 'xml') or (m3d[0] != 'm3d') or (func[0] != 'type'):
        print('Параметры заданы не верно')
        print('Пример: Program.py -a xml=123 m3d=123 type:rnd:Радиус')
        time.sleep(5)
        quit()

    kompas6_constants_3d, kompas_object, kompas_api7_module, Documents, kompas6_constants, kompas6_api5_module, kompas_api7_module = start_kompas()
    name_m3d, name_xml = a_enter_filename(m3d, xml)
    kompas_document, VariableCollection, iPart, kompas_document_3d = open(m3d[1], kompas6_constants_3d, kompas_object, kompas_api7_module, Documents)
    tree = a_tree(name_xml)
    root = tree.getroot()
    elem = a_name_elem(func, root)
    parent = select_parent(elem, root)

    if func[1] == 'rnd':
        new_value = a_change_variable_rnd(parent, kompas_document, VariableCollection, iPart)
        a_create_cdw(Documents, kompas_document_3d, kompas6_constants, 
                kompas6_api5_module, kompas_api7_module, kompas_object, name_m3d,
                func[2], new_value)
        kompas_object.Quit()
        quit()

    elif func[1] == 'for':
        if len(func) < 6:
            print('Количество полученных данных не соответствует нужному')
            time.sleep(5)
            quit()
        a_change_variable_for(parent, kompas_document, VariableCollection, iPart, func,
                             Documents, kompas_document_3d, kompas6_constants, 
                             kompas6_api5_module, kompas_api7_module, kompas_object, name_m3d)
        kompas_object.Quit()
        quit()
    else:  
        print('Введенная вами команда не поддерживается')
        time.sleep(5)
        quit()

    quit()


def h(h):
    for i in dicthelp:
        print(dicthelp[i])
    quit()

dicthelp = {
    '-h':'''Команда -h выводит информацию о параметрах для запуска программы''',
    '-i':'''Команда -i запускает интерактивный режим программы
            пользователь вводит данные в командной строке''',
    '-a':'''Команда -а запускает автоматический режим программы
            c этой командой необхожимо передать следующие значения:''',
    'xml':'Команда xml=[Название файла]',
    'm3d':'Команда m3d=[Название файла]',
    'type':'''Команда type=[rnd[название переменной]/ for:[название переменной]:min:max:step]"
              rnd - Рандомно изменяет указаную переменную,
              for - Изменяет указанную переменную от минимального значения до максимального
              с определенным шагом, и создает сразу несколько копий модели '''
}


elem = ''
parent = ''

if len(sys.argv)>1:
    if sys.argv[1] == '-i':
        kompas6_constants_3d, kompas_object, kompas_api7_module, Documents, kompas6_constants, kompas6_api5_module, kompas_api7_module = start_kompas()
        name_m3d, name_xml = i_enter_filename()
        kompas_document, VariableCollection, iPart, kompas_document_3d = open(name_m3d, kompas6_constants_3d, kompas_object, kompas_api7_module, Documents)
        tree = elem_is(name_xml)
        root = tree.getroot()
    
        flag = 'Да'
        while flag == 'Да':
            elem = select_elem()
            parent = select_parent(elem, root)
            change_variable(parent, kompas_document, VariableCollection, iPart)
            flag = input('\nПродолжить?(Да/Нет) ')
        else: 
            create_cdw(Documents, kompas_document_3d, kompas6_constants, 
                kompas6_api5_module, kompas_api7_module, kompas_object, name_m3d)
            quit()
    elif sys.argv[1] == '-h':
        h(sys.argv[1])
    elif sys.argv[1] == '-a':
        auto(sys.argv[1])
    else:
        print('Заданы неверные параметры')
        quit()
else:
    kompas6_constants_3d, kompas_object, kompas_api7_module, Documents, kompas6_constants, kompas6_api5_module, kompas_api7_module = start_kompas()
    name_m3d, name_xml = i_enter_filename()
    kompas_document, VariableCollection, iPart, kompas_document_3d = open(name_m3d, kompas6_constants_3d, kompas_object, kompas_api7_module, Documents)
    tree = elem_is(name_xml)
    root = tree.getroot()
    flag = 'Да'
    while flag == 'Да':
        elem = select_elem()
        parent = select_parent(elem, root)
        change_variable(parent, kompas_document, VariableCollection, iPart)
        flag = input('\nПродолжить?(Да/Нет) ')
    else: 
        create_cdw(Documents, kompas_document_3d, kompas6_constants, 
                kompas6_api5_module, kompas_api7_module, kompas_object, name_m3d)
        quit()
    







