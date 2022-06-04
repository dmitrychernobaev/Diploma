# -*- coding: utf-8 -*-
import os
from xml.etree import ElementTree
import pythoncom
from win32com.client import Dispatch, gencache
import LDefin2D
import MiscellaneousHelpers as MH

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
#  Получим активный документ
kompas_document = application.ActiveDocument
kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
iDocument3D = kompas_object.ActiveDocument3D()

iPart7 = kompas_document_3d.TopPart
iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

VariableCollection = iPart.VariableCollection() #Получение колекции перменных
VariableCollection.refresh() #обновление коллекции перменных


tree = ElementTree.parse("C:\\Users\\1\\Desktop\\file.xml")
root = tree.getroot()
elem = ''
parent = ''

def elem_is(): 
    print('В выбранной вами фигуре вам доступны следующие элементы для изменения:')
    for element in root.iter('variable'):
            for child in element.iter('name'):
                print(' ', child.text)

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
    return input('\nПожалуйста выберите один из представленных элементов, чтобы изменить его: \n')

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

def select_parent(): #Получение всех данных о выбранной перменной
    for element in root.iter('variable'): 
        for child in element.iter('name'):
            if elem == child.text:
                return element
 

flag = 'Да'
while flag == 'Да':
    elem_is()
    elem = select_elem()
    parent = select_parent()
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
        new_value = float(parent[3].text)
        
    Variable.value = new_value
    print('Новое значение переменной: ', Variable.value)
    iPart.RebuildModel()
    flag = input('\n Продолжить?(Да/Нет) ')
        
    







