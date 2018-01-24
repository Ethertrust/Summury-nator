import xml.etree.ElementTree as ET
import configparser
import json
import math
import xlwt
import collections
import os.path
import sys


class XlsWriter:
    def writetoxls(self, data, semdata, file, inputfile, exlist, settings, zaoch):
        style0 = xlwt.easyxf('font: name Arial, color-index black, bold on; align: vert centre, horiz left; borders: left thin, top thin, bottom thin, right thin',
                             num_format_str='#,##0.00')
        style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on; align: vert centre, horiz center; borders: top thin, bottom thin, right thin',
                             num_format_str='#,##0.00')
        style2 = xlwt.easyxf('font: name Times New Roman, color-index black; align: vert centre, horiz center; borders: top thin, bottom thin, right thin',
                             num_format_str='#,##0.00')
        style3 = xlwt.easyxf('font: name Arial, color-index green, bold on; align: vert centre, horiz center; borders: top thin, bottom thin, right thin',
            num_format_str='#,##0.00')
        style4 = xlwt.easyxf('font: name Arial, color-index orange, bold on; align: vert centre, horiz center; borders: left thin, bottom thin, right thin',
            num_format_str='#,##0.00')
        style5 = xlwt.easyxf('font: name Arial, color-index green, bold on; align: vert centre, horiz center; borders: left thin, top thin, bottom thin, right thin',
            num_format_str='#,##0.00')
        style6 = xlwt.easyxf('font: name Arial, color-index black, bold on; align: vert centre, horiz center; borders: left thin, bottom thin, right thin',
            num_format_str='#,##0.00')
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('Нагрузка')
        #worksheet.col(0).width_mismatch = True
        worksheet.col(0).width = 256 * 25
        semdata2 = collections.OrderedDict(sorted(semdata.items()))
        semdata3 = {}
        for sem in semdata2.items():
            #print(sem)
            semdata3[sem[0]] = collections.OrderedDict(sorted(sem[1].items()))
        semdata2 = semdata3
        #print(semdata2)
        worksheet.write(1, 1, 'Групп: ', style0)
        worksheet.write(1, 2, '1', style1)
        worksheet.write(1, 4, 'Человек: ', style0)
        worksheet.write(1, 5, settings['st']['stpergr'], style1)
        worksheet.write(2, 1, 'Подгрупп: ', style0)
        worksheet.write(2, 2, str(math.ceil(float(settings['st']['stpergr'])/float(settings['st']['stpersubgr']))), style1)
        worksheet.write(2, 4, 'Человек: ', style0)
        worksheet.write(2, 5, settings['st']['stpersubgr'], style1)
        worksheet.write(4, 0, inputfile, style1)

        kurs = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        if not zaoch:
            x = 2
            for sem, value in semdata2.items():
                #print(sem, '', value)
                y = 4

                worksheet.write(y, x, str(sem) + ' семестр', style0)
                worksheet.col(x).width = 256 * 10
                for el, value in value.items():
                    y += 1
                    worksheet.write(y, x, value, style2)
                    #print(y - 3)
                    kurs[y-5] += value

                y = 4
                if (x - 2) % 3 == 1:
                    worksheet.col(x+1).width = 256 * 11
                    worksheet.write(y, x + 1, str(math.ceil(int(sem) / 2)) + ' курс', style3)
                    for y in range(5, 23):
                        worksheet.write(y, x + 1, kurs[y - 5], style1)
                    kurs = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
                    x += 1
                x += 1

            data2 = collections.OrderedDict(sorted(data.items()))
            x = 1
            y = 4
            worksheet.col(x).width = 256 * 11
            worksheet.write(y, x, 'Итого', style3)
            for el, value in data2.items():
                y += 1
                worksheet.write(y, x - 1, str(el), style0)
                worksheet.write(y, x, value, style1)
        else:
            x = 2
            for sem, value in semdata2.items():
                # print(sem, '', value)
                y = 4
                kursstr = math.floor(int(sem)/3)
                ses = int(sem) % 3

                if ses == 1:
                    ses = 'Установочная '
                if ses == 2:
                    ses = 'Зимняя '
                if ses == 0:
                    ses = 'Летняя '

                worksheet.write(y, x, str(ses) + ' сессия', style0)
                worksheet.col(x).width = 256 * 10
                for el, value in value.items():
                    y += 1
                    worksheet.write(y, x, value, style2)
                    # print(y - 3)
                    kurs[y - 5] += value

                y = 4
                if ses == 'Летняя ':
                    worksheet.col(x + 1).width = 256 * 11
                    worksheet.write(y, x + 1, str(kursstr) + ' курс', style3)
                    for y in range(5, 23):
                        worksheet.write(y, x + 1, kurs[y - 5], style1)
                    kurs = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
                    x += 1
                x += 1

            data2 = collections.OrderedDict(sorted(data.items()))
            x = 1
            y = 4
            worksheet.col(x).width = 256 * 11
            worksheet.write(y, x, 'Итого', style3)
            for el, value in data2.items():
                y += 1
                worksheet.write(y, x - 1, str(el), style0)
                worksheet.write(y, x, value, style1)

        c = y
        x = -2

        if len(exlist) > 0:
            x += 2
            y = c
            y += 2
            worksheet.col(x).width = 256 * 30
            worksheet.write(y, x, 'Не найденные исключения:', style5)
            for el in exlist:
                y += 1
                worksheet.write(y, x, el, style6)

        exdictobj = Hours()
        exdictobj.appendexlist(settings)
        if len(exdictobj.exdict) > 0:
            x += 2
            y = c
            y += 2
            worksheet.col(x).width = 256 * 30
            worksheet.write(y, x, 'Исключения:', style5)
            exdict = collections.OrderedDict(sorted(exdictobj.exdict.items()))
            for el, value in exdict.items():
                y += 1
                worksheet.write(y, x, el, style4)
                for elem in value:
                    y += 1
                    worksheet.write(y, x, elem, style6)

        x += 2
        y = c
        y += 2
        if worksheet.col(x).width < 256 * 15:
            worksheet.col(x).width = 256 * 14
        worksheet.write(y, x, 'Настройки:', style5)
        y += 1
        worksheet.write(y, x, 'TimeNormals', style4)
        for el, value in settings['TimeNormals'].items():
            y += 1
            worksheet.write(y, x, el + ': ' + value, style6)

        workbook.save(file)

class XmlReader:
    tree = {}
    root = {}

    def maketree(self, path):
        self.tree = ET.parse(path)
        self.root = self.tree.getroot()

    def childs(self, node):
        for child in list(node.iter()):
            print(child.tag, ' ', child.attrib, ' ', child.keys())

    def discomp(self, settings):
        for neighbor in self.root.iter('Строка'):
            # print(neighbor.attrib)
            if not neighbor.get('ДисциплинаДляРазделов', 0) == 0:
                # print(neighbor.get('Дис', ''), ' ', neighbor.get('ДисциплинаДляРазделов', ''))
                continue
            if 'ФТД' in neighbor.get('НовИдДисциплины', ''):
                # print(neighbor.get('НовИдДисциплины', ''), ' ', 'ФТД')
                continue
            yield neighbor

    def dis(self, settings):
        print(self.root.get('UserName', '1'), 2)
        if self.root.get('UserName', '1') != '1':
            for neighbor in self.root.iter('{http://tempuri.org/dsMMISDB.xsd}ПланыСтроки'):
                #print(neighbor.tag, neighbor.attrib)
                if not neighbor.get('ДисциплинаДляРазделов', 0) == 0:
                    continue
                if 'ФТД' in neighbor.get('ДисциплинаКод', ''):
                    print(neighbor.get('ДисциплинаКод', ''), ' ', 'ФТД')
                    continue

                for neighbor2 in self.root.iter('{http://tempuri.org/dsMMISDB.xsd}ПланыНовыеЧасы'):
                    if neighbor2.get('КодОбъекта', 'err') == neighbor.get('Код', ''):
                        neighbor2.set('Ном', str((int(neighbor2.get('Курс', 0))-1)*2 + int(neighbor2.get('Семестр', 0))))
                        neighbor2.set('Дис', neighbor.get('Дисциплина', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '101':
                            neighbor2.set('Лек', neighbor2.get('Количество', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '103':
                            neighbor2.set('Пр', neighbor2.get('Количество', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '102':
                            neighbor2.set('Лаб', neighbor2.get('Количество', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '6':
                            neighbor2.set('КонтрРаб', neighbor2.get('Количество', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '5':
                            neighbor2.set('КР', neighbor2.get('Количество', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '4':
                            neighbor2.set('КП', neighbor2.get('Количество', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '11':
                            neighbor2.set('РГР', neighbor2.get('Количество', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '3':
                            neighbor2.set('ЗачО', neighbor2.get('Количество', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '2':
                            neighbor2.set('Зач', neighbor2.get('Количество', 0))
                        if neighbor2.get('КодВидаРаботы', 0) == '1':
                            neighbor2.set('Экз', neighbor2.get('Количество', 0))

                        print(neighbor2.tag)
                        neighbor2.tag = 'Сем'

                        print(neighbor2.tag)

                        splits = neighbor.get('ДисциплинаКод', '').split('.')
                        if 'ДВ' in splits:
                            #print(neighbor.attrib['Дис'])
                            #print(neighbor.attrib['Дис'])
                            if int(splits[-1]) <= int(math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr']))):
                                # for x in range(0, ):
                                neighbor2.set('DV', splits[-1])
                                neighbor2.set('Дис', neighbor2.get('Дис', 0))
                                neighbor2.set('Компетенции', neighbor.get('Компетенции', ''))
                                yield neighbor2
                        else:
                            neighbor2.set('Дис', neighbor2.get('Дис', 0))
                            neighbor2.set('Компетенции', neighbor.get('Компетенции', ''))
                            yield neighbor2
#-----
            for neighbor in self.root.iter('СпецВидыРаботНов'):
                #print(neighbor.attrib)
                for nir in neighbor.iter('НИР'):
                    for prpr in nir.iter('ПрочаяПрактика'):
                        for sem in prpr.iter('Семестр'):
                            sem.set('ЗЕТвНеделе', prpr.get('ЗЕТвНеделе', 0))
                            sem.set('Наименование', prpr.get('Наименование', ''))
                            sem.set('Компетенции', prpr.get('Компетенции', ''))
                            yield sem

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for lpr in neighbor.iter('УчебПрактики'):
                    for prpr in lpr.iter('ПрочаяПрактика'):
                        for sem in prpr.iter('Семестр'):
                            sem.set('ЗЕТвНеделе', prpr.get('ЗЕТвНеделе', 0))
                            sem.set('Наименование', prpr.get('Наименование', ''))
                            sem.set('Компетенции', prpr.get('Компетенции', ''))
                            yield sem

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for prprs in neighbor.iter('ПрочиеПрактики'):
                    for prpr in prprs.iter('ПрочаяПрактика'):
                        for sem in prpr.iter('Семестр'):
                            sem.set('ЗЕТвНеделе', prpr.get('ЗЕТвНеделе', 0))
                            sem.set('Наименование', prpr.get('Наименование', ''))
                            sem.set('Компетенции', prpr.get('Компетенции', ''))
                            yield sem

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for dip in neighbor.iter('Диплом'):
                    yield dip

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for vkr in neighbor.iter('ВКР'):
                    yield vkr

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for dip in neighbor.iter('Диплом'):
                    for gak in dip.iter('ГАК'):
                        yield gak

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for vkr in neighbor.iter('ВКР'):
                    for gak in vkr.iter('ГАК'):
                        yield gak

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                ilist = []
                for child in list(neighbor.iter()):
                    if 'ИтоговыйЭкзамен' in child.tag:
                        ilist.append(child)
                for finex in ilist:
                    #print(finex.tag)
                    for finalex in finex.iter(finex.tag):
                        #print(' ', finalex.tag)
                        yield finalex
#-----
        else:
            for neighbor in self.root.iter('Строка'):
                #print(neighbor.attrib)
                if not neighbor.get('ДисциплинаДляРазделов', 0) == 0:
                    #print(neighbor.get('Дис', ''), ' ', neighbor.get('ДисциплинаДляРазделов', ''))
                    continue
                if 'ФТД' in neighbor.get('НовИдДисциплины', ''):
                    #print(neighbor.get('НовИдДисциплины', ''), ' ', 'ФТД')
                    continue

                splits = neighbor.get('НовИдДисциплины', '').split('.')
                if 'ДВ' in splits:
                    #print(neighbor.attrib['Дис'])
                    if int(splits[-1]) <= int(math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr']))):
                        #for x in range(0, ):
                        for sem in neighbor.iter('Сем'):
                            #if int(sem.get('Ном', 0)) == 6 and int(sem.get('КП', 0)) > 0:
                            #    print(neighbor.get('Дис', 0), ' <---Сем. ном: ', sem.get('Ном', 0))
                            sem.set('DV', splits[-1])
                            sem.set('Дис', neighbor.get('Дис', 0))
                            sem.set('Компетенции', neighbor.get('Компетенции', ''))
                            yield sem
                        for kurs in neighbor.iter('Курс'):
                            for ses in kurs.iter('Сессия'):
                                ses.set('Ном', '.'.join([kurs.get('Ном', ''), ses.get('Ном', '')]))
                                print(ses.get('Ном', ''))
                                ses.set('DV', splits[-1])
                                ses.set('Дис', neighbor.get('Дис', 0))
                                ses.set('Компетенции', neighbor.get('Компетенции', ''))
                                yield ses
                else:
                    for sem in neighbor.iter('Сем'):
                        #if int(sem.get('Ном', 0)) == 6 and int(sem.get('КП', 0)) > 0:
                        #    print(neighbor.get('Дис', 0), ' <---Сем. ном: ', sem.get('Ном', 0))
                        sem.set('Дис', neighbor.get('Дис', 0))
                        sem.set('Компетенции', neighbor.get('Компетенции', ''))
                        yield sem
                    for kurs in neighbor.iter('Курс'):
                        for ses in kurs.iter('Сессия'):
                            ses.set('Ном', '.'.join([kurs.get('Ном', ''), ses.get('Ном', '')]))
                            print(ses.get('Ном', ''))
                            ses.set('Дис', neighbor.get('Дис', 0))
                            ses.set('Компетенции', neighbor.get('Компетенции', ''))
                            yield ses

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                #print(neighbor.attrib)
                for nir in neighbor.iter('НИР'):
                    for prpr in nir.iter('ПрочаяПрактика'):
                        for sem in prpr.iter('Семестр'):
                            sem.set('ЗЕТвНеделе', prpr.get('ЗЕТвНеделе', 0))
                            sem.set('Наименование', prpr.get('Наименование', ''))
                            sem.set('Компетенции', prpr.get('Компетенции', ''))
                            yield sem

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for lpr in neighbor.iter('УчебПрактики'):
                    for prpr in lpr.iter('ПрочаяПрактика'):
                        for sem in prpr.iter('Семестр'):
                            sem.set('ЗЕТвНеделе', prpr.get('ЗЕТвНеделе', 0))
                            sem.set('Наименование', prpr.get('Наименование', ''))
                            sem.set('Компетенции', prpr.get('Компетенции', ''))
                            yield sem

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for prprs in neighbor.iter('ПрочиеПрактики'):
                    for prpr in prprs.iter('ПрочаяПрактика'):
                        for sem in prpr.iter('Семестр'):
                            sem.set('ЗЕТвНеделе', prpr.get('ЗЕТвНеделе', 0))
                            sem.set('Наименование', prpr.get('Наименование', ''))
                            sem.set('Компетенции', prpr.get('Компетенции', ''))
                            yield sem

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for dip in neighbor.iter('Диплом'):
                    yield dip

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for vkr in neighbor.iter('ВКР'):
                    yield vkr

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for dip in neighbor.iter('Диплом'):
                    for gak in dip.iter('ГАК'):
                        yield gak

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                for vkr in neighbor.iter('ВКР'):
                    for gak in vkr.iter('ГАК'):
                        yield gak

            for neighbor in self.root.iter('СпецВидыРаботНов'):
                # print(neighbor.attrib)
                ilist = []
                for child in list(neighbor.iter()):
                    if 'ИтоговыйЭкзамен' in child.tag:
                        ilist.append(child)
                for finex in ilist:
                    #print(finex.tag)
                    for finalex in finex.iter(finex.tag):
                        #print(' ', finalex.tag)
                        yield finalex

    def checkedulevel(self, string):
        if self.root[0].attrib['ОбразовательнаяПрограмма'] == string:
            return True
        else:
            return False

    def checkeduform(self, string):
        if self.root[0].get('ФормаОбучения', '') == string:
            return True
        else:
            return False

class Settings:
    config = {}

    def template(self):
        with open(os.path.dirname(os.path.realpath(sys.argv[0])) + '\\' + 'settings.ini', 'w', encoding='utf-8-sig') as configfile:
            configfile.write('[st]\n#количество студентов на группу, изучающих данное направление\n'
                             'stpergr = 25\n#количество студентов на подгруппу, изучающих данное направление\nstpersubgr = 13\n\n[TimeNormals]\n'
                             '#Экзамен\nexam = 0.4\nexcons = 2.0\n#Зачёт\ncredit = 0.3\n#Диф зачёт - "ЗачО"\ndifcredit = 0.35\n'
                             '#РГР\nrgr = 0.75\n# Курсовой проект КП\nKP = 3.0\n#Курсовая работа КР\nKR = 8.0\n#Контрольная работа\n'
                             'Kont = 0.3\n#Лекции\nlec = 1.0\n#Практические занятия\npr = 1.0\n#Руководство магистрантом в год\n'
                             'mag = 12\n\n'
                             '#Указать "Дис" из УП и количество подгрупп на группу (Дис = number)\n[DisExceptionsForSubg]\n'
                             'Иностранный язык = 2\n\n[Практ]\n\n[Лаб]\n\n'
                             '#InFile\n[PathToXMLFile]\npath = \n\n#OutFile\n[PathToResultFile]\n#default is "result.xls"\npath = result.xls\n\n'
                             '#This section is reserved for future use\n[Others]')

    def readsettings(self):
        self.config = configparser.ConfigParser()
        if not os.path.isfile(os.path.dirname(os.path.realpath(sys.argv[0])) + '\\' + 'settings.ini'):
            self.template()
        self.config.read_file(open(os.path.dirname(os.path.realpath(sys.argv[0])) + '\\' + "settings.ini", 'r', encoding='utf-8-sig'))
        if not (self.config.has_section('TimeNormals') and self.config.has_section('st')):
            self.template()
            self.config.read_file(open(os.path.dirname(os.path.realpath(sys.argv[0])) + '\\' + "settings.ini", 'r', encoding='utf-8-sig'))


    def writesettings(self):
        with open(os.path.dirname(os.path.realpath(sys.argv[0])) + '\\' + 'settings.ini', 'w', encoding='utf-8-sig') as configfile:
            self.config.write(configfile)

    def printvalues(self):
        for elem in self.config.sections():
            for key, value in self.config[elem].items():
                print(key, ' = ', value)

class Hours():
    hours = 0
    semsh = {}
    attrs = {}
    plan = ''
    lastsem = 0
    summed = {}
    stnumber = 0
    subgroups = 0
    tempsub = 0
    groups = 0
    settingsset = False
    exlist = {}
    disaud = {}
    exdict = {}
    compdict = {}
    diswithoutcomp = []

    def competenceslist(self, node):
        if not node.get('Индекс', '') == '':
            self.compdict[int(node.get('Код', ''))] = {'Содержание': node.get('Содержание', ''),
                                                  'Индекс': node.get('Индекс', ''),
                                                  'Матрица': {}
                                                  }

    def appendexlist(self, settings):
        list = []
        for sec in settings.items():
            #print(sec[0])
            g = 0
            for el in sec:
                if g%2 == 0:
                    #print(el)
                    list.append(el)
                g += 1
        list.pop(0)
        #print(list)
        for el in ['st', 'TimeNormals', 'PathToXMLFile', 'PathToResultFile', 'Others']:
            list.remove(el)

        #print(list)
        for elem in list:
            for el, val in settings[elem].items():
                self.exlist[el] = True
        #print(self.exlist)

        for elem in list:
            if len(settings[elem].items()) > 0:
                self.exdict[elem] = []
            for el, val in settings[elem].items():
                self.exdict[elem].append(el + ': ' + val)
        #print(self.exdict)

    def removeexitem(self, node):
        if node.get('Дис', '').lower() in self.exlist:
            if not node.get('Дис', '') == '':
                self.exlist.pop(node.get('Дис', '').lower())

    def calcdiv(self, node, settings, div=1):
        for dis, subgroups in settings['DisExceptionsForSubg'].items():
            #print('Перечисляю: ', dis)
            if node.get('Дис', 0).lower() == dis:
                #print('Hey 1, subgr: ', node.get('Дис', ''))
                return math.ceil(self.stnumber/int(subgroups))
        #print('Hey 2, multipler: ', node.get('Дис', ''))
        return int(div)

    def calcmult(self, node, settings, exc, multipler=1):
        #for dis, subgroups in settings['DisExceptionsForSubg'].items():
        #print('Перечисляю: ', dis)

        if settings.has_section(exc):
            if node.get('Дис', '').lower() in settings[exc]:
                return int(settings[exc][node.get('Дис', '').lower()])

        if settings.has_section('DisExceptionsForSubg'):
            if node.get('Дис', '').lower() in settings['DisExceptionsForSubg']:
                #if int(node.get('Ном', 0)) == 2 and int(node.get('Пр', 0)) > 0:
                    #print('Hey 1, subgr: ', node.get('Дис', ''))
                return int(settings['DisExceptionsForSubg'][node.get('Дис', '').lower()])
            #if int(node.get('Ном', 0)) == 2 and int(node.get('Пр', 0)) > 0:
                #print('Hey 2, multipler: ', node.get('Дис', ''))

        return int(multipler)

    def set(self, settings):
        if not self.settingsset:
            self.stnumber = int(settings['st']['stpergr'])
            self.subgroups = math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr']))
            self.tempsub = self.subgroups
            self.groups = 1
            self.settingsset = True

    def getplanopt(self, node):
        for child in list(node.iter()):
            if child.tag == '{http://tempuri.org/dsMMISDB.xsd}Планы':
                node = child
        print(node.tag)
        print(node.findall('{http://tempuri.org/dsMMISDB.xsd}Планы'))
        print(node.findall('Планы'))
        if node.get('ОбразовательнаяПрограмма', '') == 'подготовка магистров' or node.get('КодУровняОбразования', '') == '3':
            self.plan = 'mag'
        else:
            self.plan = node.get('ОбразовательнаяПрограмма', node.get('КодУровняОбразования', ''))

    def checkexcept(self, settings, node, exc):
        if exc in settings:
            if node.get('Дис', '').lower() in settings[exc]:
                #if int(node.get('Ном', '')) == 3:
                #    print(node.get('Дис', '').lower(), settings[exc][node.get('Дис', '').lower()])
                self.subgroups = settings[exc][node.get('Дис', '').lower()]
                #print('subgroups: ', self.subgroups)
            else:
                self.checkdis(settings, node)
                # print('2: ', self.subgroups)
        else:
            self.checkdis(settings, node)

    def checkdis(self, settings, node):
        #print(node.get('Дис', ''), node.tag, node.attrib)
        if node.get('Дис', '').lower() in settings['DisExceptionsForSubg']:
            self.subgroups = settings['DisExceptionsForSubg'][node.get('Дис', '').lower()]
            #print('subgroups: ', self.subgroups)
        else:
            self.subgroups = self.tempsub
            #print('2: ', self.subgroups)

    def parseSemn(self, nom):
        noml = str(nom).split('.')
        if len(noml) > 1:
            #print((int(noml[0]), noml[1], str((int(noml[0]) - 1) * 3 + int(noml[1]))))
            return str((int(noml[0]) - 1) * 3 + int(noml[1]))
        else:
            return nom

    def sum(self, settings, node):
        # print(semn)
        #print(node.attrib)
        self.set(settings)
        self.checkdis(settings, node)
        self.removeexitem(node)
        #print('1: ', int(self.subgroups))
        semn = int(self.parseSemn(node.get('Ном', self.lastsem)))
        if semn > self.lastsem:
            self.lastsem = semn
        #print(self.lastsem)
        #print(semn)
        #print(node.attrib)
        if semn:
            for comp, value in self.compdict.items():
                if value['Индекс'] in node.get('Компетенции', '').split(', '):
                    if semn not in self.compdict[comp]['Матрица']:
                        self.compdict[comp]['Матрица'][semn] = []
                if not node.get('Наименование', '') == '':
                    if str(comp) in node.get('Компетенции', '').split('&amp;'):
                        if semn not in self.compdict[comp]['Матрица']:
                            self.compdict[comp]['Матрица'][semn] = []

            for comp, value in self.compdict.items():
                if not node.get('Дис', '') == '':
                    #print(node.get('Компетенции', '').split(', '))
                    if value['Индекс'] in node.get('Компетенции', '').split(', '):
                        self.compdict[comp]['Матрица'][semn].append(node.get('Дис', ''))
                if not node.get('Наименование', '') == '':
                    #print(node.get('Компетенции', '').split('&amp;'))
                    if str(comp) in node.get('Компетенции', '').split('&amp;'):
                        self.compdict[comp]['Матрица'][semn].append(node.get('Наименование', ''))

            if node.get('Компетенции', '') == '' and (node.get('Дис', '') not in self.diswithoutcomp and node.get('Наименование', '') not in self.diswithoutcomp):
                self.diswithoutcomp.append(node.get('Дис', ''))
                        #print(node.get('Дис', ''))


            if semn not in self.semsh:
                self.semsh[semn] = {'hours': 0,
                                    'Рук. маг': 0,
                                    'Лек': 0,
                                    'Лек. конс': 0,
                                    'Практ': 0,
                                    'Лаб': 0,
                                    'КонтрРаб': 0,
                                    'КурсРаб': 0,
                                    'КурсПр': 0,
                                    'РГР': 0,
                                    'Диф. зач': 0,
                                    'Зач': 0,
                                    'Экз': 0,
                                    'Экз. конс': 0,
                                    'ГЭК': 0,
                                    'ГАК': 0,
                                    'Рук. ВКР': 0,
                                    'Практики и НИР': 0
                                    }
            if semn in self.semsh:
                #print(self.semsh)
                #print(self.semsh[semn])
                if node.get('Наименование', '') not in self.disaud:
                    self.disaud[node.get('Наименование', '')] = {}
                    self.disaud[node.get('Наименование', '')][semn] = {'Лек': 0, 'Лек. конс': 0, 'Практ': 0, 'Лаб': 0, 'Вычесть': False}
                else:
                    if semn not in self.disaud[node.get('Наименование', '')]:
                        self.disaud[node.get('Наименование', '')][semn] = {'Лек': 0, 'Лек. конс': 0, 'Практ': 0, 'Лаб': 0, 'Вычесть': False}

                if node.tag == 'Сем' or node.tag == 'Сессия':
                    #print(node.tag, node.attrib)
                    atdis = node.get('Дис', '')
                    if atdis not in self.disaud:
                        self.disaud[atdis] = {}
                        self.disaud[atdis][semn] = {'Лек': 0, 'Лек. конс': 0, 'Практ': 0, 'Лаб': 0, 'Вычесть': False}
                    else:
                        if semn not in self.disaud[atdis]:
                            self.disaud[atdis][semn] = {'Лек': 0, 'Лек. конс': 0, 'Практ': 0, 'Лаб': 0, 'Вычесть': False}
                    self.checkexcept(settings, node, 'Практ')
                    #print('subgroups: ', self.subgroups)
                    if atdis == 'Прикладная физическая культура' or atdis == 'Физическая культура':
                        self.semsh[semn]['Лек'] += round((float(settings['TimeNormals']['lec'])) * int(node.get('Лек', 0)) * int(self.groups), 2)
                        self.disaud[node.get('Дис', '')][semn]['Лек'] = round((float(settings['TimeNormals']['lec'])) * int(node.get('Лек', 0)) * int(self.groups), 2)
                        #if semn == 12 and round((int(node.get('Лек', 0))) * float(settings['TimeNormals']['lec']), 2) > 0:
                        #    print('sem', semn, 'sum: ', self.semsh[semn]['Лек'], ' || ',
                        #          round((int(node.get('Лек', 0))) * float(settings['TimeNormals']['lec']), 2), node.attrib['Дис'])
                        self.semsh[semn]['Лек. конс'] += round(0.05 * int(node.get('Лек', 0)) * int(self.groups) * float(settings['TimeNormals']['lec']), 2)
                        self.disaud[node.get('Дис', '')][semn]['Лек. конс'] = round(0.05 * int(node.get('Лек', 0)) * int(self.groups) * float(settings['TimeNormals']['lec']), 2)
                        #if semn == 3 and int(node.get('Пр', 0)) > 0:
                        #    print(int(node.get('Пр', 0)), ' ', math.ceil(int(self.stnumber) / 15), ' ', float(settings['TimeNormals']['pr']))
                        #print('Прикладнуха')
                        self.semsh[semn]['Практ'] += int(node.get('Пр', 0)) * self.calcmult(node, settings, 'Практ', math.ceil(int(self.stnumber) / 15)) * float(settings['TimeNormals']['pr'])
                        self.disaud[node.get('Дис', '')][semn]['Практ'] = int(node.get('Пр', 0)) * self.calcmult(node, settings, 'Практ', math.ceil(int(self.stnumber) / 15)) * float(settings['TimeNormals']['pr'])
                        if int(node.get('Пр', 0)) > 0:
                            print(node.get('Дис', 0), ' <---Сем. ном: ', semn)
                            print('Практ: ', int(node.get('Пр', 0)), ' ',
                                  self.calcmult(node, settings, 'Практ', math.ceil(int(self.stnumber) / 15)), ' ',
                                 int(node.get('Пр', 0)) * self.calcmult(node, settings,
                                                                        'Практ', math.ceil(int(self.stnumber) / 15)) * float(
                                      settings['TimeNormals']['pr']),
                                  'Итого: ', self.semsh[semn]['Практ'])
                    else:
                        self.semsh[semn]['Лек'] += round((int(node.get('Лек', 0))) * float(settings['TimeNormals']['lec']), 2)
                        self.disaud[node.get('Дис', '')][semn]['Лек'] = round((int(node.get('Лек', 0))) * float(settings['TimeNormals']['lec']), 2)
                        #if semn == 12 and round((int(node.get('Лек', 0))) * float(settings['TimeNormals']['lec']), 2) > 0:
                        #    print('sem', semn, 'sum: ', self.semsh[semn]['Лек'], ' || ',
                        #          round((int(node.get('Лек', 0))) * float(settings['TimeNormals']['lec']), 2), node.attrib['Дис'])
                        self.semsh[semn]['Лек. конс'] += round(0.05 * int(node.get('Лек', 0)) * int(self.groups) * float(settings['TimeNormals']['lec']), 2)
                        self.disaud[node.get('Дис', '')][semn]['Лек. конс'] = round(0.05 * int(node.get('Лек', 0)) * int(self.groups) * float(settings['TimeNormals']['lec']), 2)
                        #self.summ += int(node.get('Пр', 0))
                        #print(int(node.get('Пр', 0)), ' ', int(self.groups), ' ', float(settings['TimeNormals']['pr']), ' self.sum: ', self.summ)
                        #if semn == 3 and int(node.get('Пр', 0)) > 0:
                        #    print(int(node.get('Пр', 0)), ' ', int(self.groups), ' ', float(settings['TimeNormals']['pr']))
                        #print('Было так: ', self.groups)
                        self.semsh[semn]['Практ'] += int(node.get('Пр', 0)) * self.calcmult(node, settings, 'Практ', self.groups) * float(settings['TimeNormals']['pr'])
                        self.disaud[node.get('Дис', '')][semn]['Практ'] = int(node.get('Пр', 0)) * self.calcmult(node, settings, 'Практ', self.groups) * float(settings['TimeNormals']['pr'])
                        if int(node.get('Пр', 0)) > 0:
                            print(node.get('Дис', 0), ' <---Сем. ном: ', semn)
                            print('Практ: ', int(node.get('Пр', 0)), ' ', self.calcmult(node, settings, 'Практ', self.groups), ' ',
                                 int(node.get('Пр', 0)) * self.calcmult(node, settings, 'Практ', self.groups) * float(
                                     settings['TimeNormals']['pr']),
                                  'Итого: ', self.semsh[semn]['Практ'])
                    self.semsh[semn]['Лек'] = round(self.semsh[semn]['Лек'], 2)
                    self.semsh[semn]['Лек. конс'] = round(self.semsh[semn]['Лек. конс'], 2)
                    self.semsh[semn]['Практ'] = round(self.semsh[semn]['Практ'], 2)
                        #if semn == 12 and int(node.get('Пр', 0)) * int(self.groups) * float(settings['TimeNormals']['pr']) > 0:
                        #    print('sem', semn, 'sum: ', self.semsh[semn]['Практ'], ' || ',
                        #          int(node.get('Пр', 0)) * int(self.groups) * float(settings['TimeNormals']['pr']))
                    #self.summ += int(node.get('Лаб', 0))
                    #print(int(node.get('Лаб', 0)), ' ', int(self.subgroups), ' self.sum: ', self.summ)
                    #print(int(self.subgroups))

                    self.checkexcept(settings, node, 'Лаб')
                    if 'DV' in node.attrib:
                        #self.semsh[semn]['Лаб'] += int(node.get('Лаб', 0))
                        if int(node.attrib['DV']) <= math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr'])):
                            #print('Было так: ', self.groups)
                            self.semsh[semn]['Лаб'] += int(node.get('Лаб', 0)) * self.calcmult(node, settings, 'Лаб')
                            self.disaud[node.get('Дис', '')][semn]['Лаб'] = int(node.get('Лаб', 0)) * self.calcmult(node, settings, 'Лаб')
                            #if semn == 8 and int(node.get('Лаб', 0)) > 0:
                            #    print(node.get('Дис', 0), ' <---Сем. ном: ', semn)
                            #    print('Лаб: ', int(node.get('Лаб', 0)), ' ', self.calcmult(node, settings, 'Лаб'), ' ',
                            #          int(node.get('Лаб', 0)) * self.calcmult(node, settings, 'Лаб'),
                            #          'Итого: ', self.semsh[semn]['Лаб'])
                    else:
                        self.semsh[semn]['Лаб'] += int(node.get('Лаб', 0)) * int(self.subgroups)
                        self.disaud[node.get('Дис', '')][semn]['Лаб'] = int(node.get('Лаб', 0)) * int(self.subgroups)
                        #if semn == 8 and int(node.get('Лаб', 0)) > 0:
                        #    print(node.get('Дис', 0), ' <---Сем. ном: ', semn)
                        #    print('Лаб: ', int(node.get('Лаб', 0)), ' ', int(self.subgroups), ' ',
                        #          int(node.get('Лаб', 0)) * int(self.subgroups),
                        #          'Итого: ', self.semsh[semn]['Лаб'])
                    self.semsh[semn]['Лаб'] = round(self.semsh[semn]['Лаб'], 2)

                    if node.get('Дис', '') in self.disaud and self.disaud[node.get('Дис', '')][semn]['Вычесть']:
                        #print(self.disaud[node.get('Наименование', '')])
                        self.semsh[semn]['Лек'] -= float(self.disaud[node.get('Дис', '')][semn]['Лек'])
                        self.semsh[semn]['Лек. конс'] -= float(self.disaud[node.get('Дис', '')][semn]['Лек. конс'])
                        self.semsh[semn]['Практ'] -= float(self.disaud[node.get('Дис', '')][semn]['Практ'])
                        self.semsh[semn]['Лаб'] -= float(self.disaud[node.get('Дис', '')][semn]['Лаб'])

                    self.checkexcept(settings, node, 'КонтрРаб')
                    if 'DV' in node.attrib:
                        if int(node.attrib['DV']) == math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr'])):
                            #print('Было так: ', self.stnumber - (self.subgroups - 1) * int(settings['st']['stpersubgr']))
                            #print('Стало так: ',
                            #      (self.stnumber % self.calcdiv(node, settings, int(settings['st']['stpersubgr']))))
                            if (self.stnumber % int(settings['st']['stpersubgr'])) == 0:
                                self.semsh[semn]['КонтрРаб'] += int(node.get('КонтрРаб', 0)) * int(
                                    settings['st']['stpersubgr']) * float(settings['TimeNormals']['Kont'])
                            else:
                                self.semsh[semn]['КонтрРаб'] += int(node.get('КонтрРаб', 0)) * (self.stnumber % int(settings['st']['stpersubgr'])) * float(
                                settings['TimeNormals']['Kont'])
                        else:
                            self.semsh[semn]['КонтрРаб'] += int(node.get('КонтрРаб', 0)) * int(settings['st']['stpersubgr']) * float(settings['TimeNormals']['Kont'])
                    else:
                        self.semsh[semn]['КонтрРаб'] += int(node.get('КонтрРаб', 0)) * int(self.stnumber) * float(settings['TimeNormals']['Kont'])
                    self.semsh[semn]['КонтрРаб'] = round(self.semsh[semn]['КонтрРаб'], 2)

                    self.checkexcept(settings, node, 'КурсРаб')
                    if 'DV' in node.attrib:
                        if int(node.attrib['DV']) == math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr'])):
                            if (self.stnumber % int(settings['st']['stpersubgr'])) == 0:
                                self.semsh[semn]['КурсРаб'] += int(node.get('КР', 0)) * int(
                                    settings['st']['stpersubgr']) * float(settings['TimeNormals']['kr'])
                            else:
                                self.semsh[semn]['КурсРаб'] += int(node.get('КР', 0)) * (self.stnumber % int(settings['st']['stpersubgr'])) * float(
                                settings['TimeNormals']['kr'])
                        else:
                            self.semsh[semn]['КурсРаб'] += int(node.get('КР', 0)) * int(settings['st']['stpersubgr']) * float(settings['TimeNormals']['kr'])
                    else:
                        self.semsh[semn]['КурсРаб'] += int(node.get('КР', 0)) * int(self.stnumber) * float(settings['TimeNormals']['kr'])
                    self.semsh[semn]['КурсРаб'] = round(self.semsh[semn]['КурсРаб'], 2)

                    self.checkexcept(settings, node, 'КурсПр')
                    if 'DV' in node.attrib:
                        if int(node.attrib['DV']) == math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr'])):
                            if (self.stnumber % int(settings['st']['stpersubgr'])) == 0:
                                self.semsh[semn]['КурсПр'] += int(node.get('КП', 0)) * int(
                                    settings['st']['stpersubgr']) * float(settings['TimeNormals']['kp'])
                            else:
                                self.semsh[semn]['КурсПр'] += int(node.get('КП', 0)) * (self.stnumber % int(settings['st']['stpersubgr'])) * float(
                                settings['TimeNormals']['kp'])
                        else:
                            self.semsh[semn]['КурсПр'] += int(node.get('КП', 0)) * int(settings['st']['stpersubgr']) * float(settings['TimeNormals']['kp'])
                    else:
                        self.semsh[semn]['КурсПр'] += int(node.get('КП', 0)) * int(self.stnumber) * float(settings['TimeNormals']['kp'])
                    self.semsh[semn]['КурсПр'] = round(self.semsh[semn]['КурсПр'], 2)

                    self.checkexcept(settings, node, 'РГР')
                    if 'DV' in node.attrib:
                        if int(node.attrib['DV']) == math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr'])):
                            if (self.stnumber % int(settings['st']['stpersubgr'])) ==0:
                                self.semsh[semn]['РГР'] += int(node.get('РГР', 0)) * int(
                                    settings['st']['stpersubgr']) * float(settings['TimeNormals']['rgr'])
                            else:
                                self.semsh[semn]['РГР'] += int(node.get('РГР', 0)) * (self.stnumber % int(settings['st']['stpersubgr'])) * float(settings['TimeNormals']['rgr'])
                        else:
                            self.semsh[semn]['РГР'] += int(node.get('РГР', 0)) * int(settings['st']['stpersubgr']) * float(settings['TimeNormals']['rgr'])
                    else:
                        self.semsh[semn]['РГР'] += int(node.get('РГР', 0)) * int(self.stnumber) * float(
                        settings['TimeNormals']['rgr'])
                    self.semsh[semn]['РГР'] = round(self.semsh[semn]['РГР'], 2)

                    self.checkexcept(settings, node, 'Диф.  зач')
                    if 'DV' in node.attrib:
                        if int(node.attrib['DV']) == math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr'])):
                            if (self.stnumber % int(settings['st']['stpersubgr'])) == 0:
                                self.semsh[semn]['Диф. зач'] += int(node.get('ЗачО', 0)) * int(
                                    settings['st']['stpersubgr']) * float(settings['TimeNormals']['difcredit'])
                            else:
                                self.semsh[semn]['Диф. зач'] += int(node.get('ЗачО', 0)) * (self.stnumber % int(settings['st']['stpersubgr'])) * float(settings['TimeNormals']['difcredit'])
                        else:
                            self.semsh[semn]['Диф. зач'] += int(node.get('ЗачО', 0)) * int(settings['st']['stpersubgr']) * float(settings['TimeNormals']['difcredit'])
                    else:
                        self.semsh[semn]['Диф. зач'] += int(node.get('ЗачО', 0)) * int(self.stnumber) * float(
                        settings['TimeNormals']['difcredit'])
                    self.semsh[semn]['Диф. зач'] = round(self.semsh[semn]['Диф. зач'], 2)
                    #if int(node.get('ЗачО', 0)) * int(self.stnumber) * float(settings['TimeNormals']['difcredit']) > 0:
                    #    print('sem', semn, 'sum: ', self.semsh[semn]['Диф. зач'], ' || ',
                    #          int(node.get('ЗачО', 0)) * int(self.stnumber) * float(
                    #              settings['TimeNormals']['difcredit']))

                    self.checkexcept(settings, node, 'Зач')
                    if 'DV' in node.attrib:
                        if int(node.attrib['DV']) == math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr'])):
                            if self.stnumber % int(settings['st']['stpersubgr']) == 0:
                                self.semsh[semn]['Зач'] += int(node.get('Зач', 0)) * int(
                                    settings['st']['stpersubgr']) * float(settings['TimeNormals']['credit'])
                            else:
                                self.semsh[semn]['Зач'] += int(node.get('Зач', 0)) * (self.stnumber % int(settings['st']['stpersubgr'])) * float(settings['TimeNormals']['credit'])
                            #if semn == 12 and int(node.get('Зач', 0)) * (
                            #            self.stnumber - (self.subgroups - 1) * int(settings['st']['stpersubgr'])) * float(
                            #    settings['TimeNormals']['credit']) > 0:
                            #    print(int(node.get('Зач', 0)) * (
                            #    self.stnumber - (self.subgroups - 1) * int(settings['st']['stpersubgr'])) * float(
                            #        settings['TimeNormals']['credit'])
                            #          , node.attrib['Дис'], self.semsh[semn]['Зач'])
                        else:
                            self.semsh[semn]['Зач'] += int(node.get('Зач', 0)) * int(settings['st']['stpersubgr']) * float(settings['TimeNormals']['credit'])
                            #if semn == 12 and int(node.get('Зач', 0)) * int(settings['st']['stpersubgr']) * float(
                            #        settings['TimeNormals']['credit']) > 0:
                            #    print(int(node.get('Зач', 0)) * int(settings['st']['stpersubgr']) * float(
                            #        settings['TimeNormals']['credit']),
                            #          node.attrib['Дис'],
                            #          self.semsh[semn]['Зач'])
                    else:

                        self.semsh[semn]['Зач'] += int(node.get('Зач', 0)) * int(self.stnumber) * float(settings['TimeNormals']['credit'])
                        #if semn == 12 and int(node.get('Зач', 0)) * int(self.stnumber) * float(
                        #        settings['TimeNormals']['credit']) > 0:
                        #    print(int(node.get('Зач', 0)) * int(self.stnumber) * float(settings['TimeNormals']['credit']),
                        #          node.attrib['Дис'],
                        #          self.semsh[semn]['Зач'])
                    self.semsh[semn]['Зач'] = round(self.semsh[semn]['Зач'], 2)
                    #if semn == 1 and int(node.get('Зач', 0)) > 0:
                    #    print(node.get('Дис', 0), ' <---Сем. ном: ', semn)
                    #    print('Зач: ', int(node.get('Зач', 0)), ' ', int(self.stnumber), ' ',
                    #          int(node.get('Зач', 0)) * int(self.stnumber) * float(settings['TimeNormals']['credit']),
                    #          'Итого: ', self.semsh[semn]['Зач'])
                    #if semn == 6 and int(node.get('Зач', 0)) * int(self.stnumber) * float(settings['TimeNormals']['credit']) > 0:
                    #    print('sem', semn, 'sum: ', self.semsh[semn]['Зач'], ' || ',
                    #          int(node.get('Зач', 0)) * int(self.stnumber) * float(settings['TimeNormals']['credit']))

                    self.checkexcept(settings, node, 'Экз')
                    if 'DV' in node.attrib:
                        if int(node.attrib['DV']) == math.ceil(int(settings['st']['stpergr']) / int(settings['st']['stpersubgr'])):
                            #if semn == 3:
                            #    print('Я тут   1',
                            #          (self.stnumber - (self.subgroups - 1) * int(settings['st']['stpersubgr'])))
                            if (self.stnumber % int(settings['st']['stpersubgr'])) == 0:
                                self.semsh[semn]['Экз'] += int(node.get('Экз', 0)) * int(
                                    settings['st']['stpersubgr']) * float(settings['TimeNormals']['exam'])
                            else:
                                self.semsh[semn]['Экз'] += int(node.get('Экз', 0)) * (self.stnumber % int(settings['st']['stpersubgr'])) * float(settings['TimeNormals']['exam'])
                            #if semn == 3 and int(node.get('Экз', 0)) * (
                            #    self.stnumber - (self.subgroups - 1) * int(settings['st']['stpersubgr'])) * float(
                            #        settings['TimeNormals']['exam']) > 0:
                            #    print('sem', semn, 'sum: ', self.semsh[semn]['Экз'], ' || ',
                            #          int(node.get('Экз', 0)) * (
                            #          self.stnumber - (self.subgroups - 1) * int(settings['st']['stpersubgr'])) * float(
                            #              settings['TimeNormals']['exam']))
                        else:
                            #if semn == 3:
                            #    print('Я тут   2', int(settings['st']['stpersubgr']))
                            self.semsh[semn]['Экз'] += int(node.get('Экз', 0)) * int(settings['st']['stpersubgr']) * float(settings['TimeNormals']['exam'])
                            #if semn == 3 and int(node.get('Экз', 0)) * int(settings['st']['stpersubgr']) * float(
                            #        settings['TimeNormals']['exam']) > 0:
                            #    print('sem', semn, 'sum: ', self.semsh[semn]['Экз'], ' || ',
                            #          int(node.get('Экз', 0)) * int(settings['st']['stpersubgr']) * float(
                            #              settings['TimeNormals']['exam']))
                    else:
                        self.semsh[semn]['Экз'] += int(node.get('Экз', 0)) * int(self.stnumber) * float(
                            settings['TimeNormals']['exam'])
                        #if semn == 3 and int(node.get('Экз', 0)) * int(self.stnumber) * float(settings['TimeNormals']['exam']) > 0:
                        #    print('sem', semn, 'sum: ', self.semsh[semn]['Экз'], ' || ',
                        #          int(node.get('Экз', 0)) * int(self.stnumber) * float(settings['TimeNormals']['exam']))
                    self.semsh[semn]['Экз'] = round(self.semsh[semn]['Экз'], 2)
                    #if 'DV' in node.attrib:
                    #    if int(node.attrib['DV']) == self.subgroups:
                    #        self.semsh[semn]['Экз. конс'] += (float(settings['TimeNormals']['excons']) / int(self.subgroups)) * int(node.get('Экз', 0))
                    #    else:
                    #        self.semsh[semn]['Экз. конс'] += (float(settings['TimeNormals']['excons']) / int(self.subgroups)) * int(node.get('Экз', 0))
                    #else:

                    self.checkexcept(settings, node, 'Экз. конс')
                    self.semsh[semn]['Экз. конс'] += (float(settings['TimeNormals']['excons']) * int(self.groups)) * int(node.get('Экз', 0))
                    self.semsh[semn]['Экз. конс'] = round(self.semsh[semn]['Экз. конс'], 2)

                    self.checkdis(settings, node)

                if self.plan == 'mag' and semn not in self.summed:
                    self.summed[semn] = True
                    self.semsh[semn]['Рук. маг'] += float(settings['TimeNormals']['mag'])/2 * int(self.stnumber)
                    self.semsh[semn]['Рук. маг'] = round(self.semsh[semn]['Рук. маг'], 2)

                if node.tag == 'Семестр':
                    for key in node.iter('Кафедра'):
                        #print(node.attrib)
                        #print(key.attrib)
                        opt1 = float(key.get('НормативНаСтуд', 0))
                        opt2 = float(key.get('НормативНаСтудВНед', 0))
                        opt3 = float(key.get('НормативНаПодгр', 0))
                        opt4 = float(key.get('НормативНаПодгрВНед', 0))
                        if not opt1 == 0:
                            self.semsh[semn]['Практики и НИР'] += float(node.get('ПланЗЕТ', 0)) / float(
                                node.get('ЗЕТвНеделе', 0)) * opt1 * int(self.stnumber) / float(key.get('Нед', 1))
                            #print(node.get('Ном', 0), 'НормативНаСтуд: ',
                            #      float(node.get('ПланЗЕТ', 0)) / float(node.get('ЗЕТвНеделе', 0)) * opt1 * int(
                            #          self.stnumber) / float(key.get('Нед', 1)))

                            #opt1 * int(self.stnumber)
                        if not opt2 == 0:
                            #print(int(key.get('Нед', 0)), ' ', opt2,' ', int(self.stnumer))
                            self.semsh[semn]['Практики и НИР'] += float(node.get('ПланЗЕТ', 0)) / float(
                                node.get('ЗЕТвНеделе', 0)) * opt2 * int(self.stnumber)
                            #print(node.get('Ном', 0), 'НормативНаСтудВНед: ',
                            #      float(node.get('ПланЗЕТ', 0)) / float(node.get('ЗЕТвНеделе', 0)) * opt2 * int(
                            #          self.stnumber))

                            #int(key.get('Нед', 0)) * opt2 * int(self.stnumber)
                        if not opt3 == 0:
                            if not node.get('ПланЧасовАуд', '') == '':
                                if node.get('Наименование', '') in self.disaud:
                                #    print(self.disaud[node.get('Наименование', '')])
                                    self.semsh[semn]['Лек'] -= float(
                                        self.disaud[node.get('Наименование', '')][semn]['Лек'])
                                    self.semsh[semn]['Лек. конс'] -= float(
                                        self.disaud[node.get('Наименование', '')][semn]['Лек. конс'])
                                    self.semsh[semn]['Практ'] -= float(
                                        self.disaud[node.get('Наименование', '')][semn]['Практ'])
                                    self.semsh[semn]['Лаб'] -= float(
                                        self.disaud[node.get('Наименование', '')][semn]['Лаб'])
                                    self.disaud[node.get('Наименование', '')][semn]['Вычесть'] = True
                                self.semsh[semn]['Практики и НИР'] += float(node.get('ПланЧасовАуд', ''))
                                #print(node.get('Ном', 0), 'НормативНаПодгр: ', float(node.get('ПланЧасовАуд', '')))
                            else:
                                self.semsh[semn]['Практики и НИР'] += float(node.get('ПланЗЕТ', 0)) / float(
                                    node.get('ЗЕТвНеделе', 0)) * opt3 * int(self.groups) / float(key.get('Нед', 1))
                                #print(node.get('Ном', 0), 'НормативНаПодгр: ',
                                #      float(node.get('ПланЗЕТ', 0)) / float(node.get('ЗЕТвНеделе', 0)) * opt3 * int(
                                #          self.groups) / float(key.get('Нед', 1)))

                            #opt3 * int(self.groups)
                        if not opt4 == 0:
                            self.semsh[semn]['Практики и НИР'] += float(node.get('ПланЗЕТ', 0)) / float(
                                node.get('ЗЕТвНеделе', 0)) * opt4 * int(self.groups)
                            #print(node.get('Ном', 0), 'НормативНаПодгрВНед: ',
                            #      float(node.get('ПланЗЕТ', 0)) / float(node.get('ЗЕТвНеделе', 0)) * opt4 * int(
                            #          self.groups))

                            #int(key.get('Нед', 0)) * opt4 * int(self.groups)

                for ruk in node.iter('Руководство'):

                        for rec in node.iter('Рецензии'):
                            #print(int(ruk[0].get('Часов', 0)), ' ', int(rec[0].get('Часов', 0)))
                            if rec and ruk:
                                self.semsh[semn]['Рук. ВКР'] += (int(ruk[0].get('Часов', 0)) + int(rec[0].get('Часов', 0))) * int(self.stnumber)
                            else:
                                if ruk:
                                    self.semsh[semn]['Рук. ВКР'] += (int(ruk[0].get('Часов', 0))) * int(self.stnumber)
                                else:
                                    self.semsh[semn]['Рук. ВКР'] += 0
                                if rec:
                                    self.semsh[semn]['Рук. ВКР'] += (int(rec[0].get('Часов', 0))) * int(self.stnumber)
                                else:
                                    self.semsh[semn]['Рук. ВКР'] += 0

                if node.tag == 'ГАК':
                    self.semsh[semn]['ГАК'] += int(self.stnumber) * float(node.get('Часов', 0))
                    for member in node.iter('ЧленГАК'):
                        self.semsh[semn]['ГАК'] += int(self.stnumber) * float(member.get('Часов', 0))

                if 'ИтоговыйЭкзамен' in node.tag:
                    self.semsh[semn]['ГЭК'] += int(self.stnumber) * float(node.get('ПредседательЧасов', 0))
                    for member in node.iter('ЧленГЭК'):
                        self.semsh[semn]['ГЭК'] += int(self.stnumber) * float(member.get('Часов', 0))
            else:
                print('trouble ----> ', node.attrib)
            self.semsh[semn]['hours'] = 0
            for key, it in self.semsh[semn].items():
                if not key == 'hours':
                    self.semsh[semn]['hours'] += it

        self.hours = 0
        for key, value in self.semsh.items():
            for key2, value2 in value.items():
                if key2 == 'hours':
                    self.hours = self.hours + value2


hours = Hours()
settings = Settings()
settings.readsettings()
#settings.printvalues()
xml = XmlReader()
if len(sys.argv) == 2:
    settings.config['PathToXMLFile']['path'] = str(sys.argv[1])
    filename, file_extension = os.path.splitext(str(sys.argv[1]))
    print(file_extension)
    if not file_extension == '':
        settings.config['PathToResultFile']['path'] = str.replace(str(sys.argv[1]), file_extension, '.xls')
    else:
        settings.config['PathToResultFile']['path'] += '.xls'
    print(settings.config['PathToResultFile']['path'])
    settings.writesettings()
    settings.config['PathToXMLFile']['path'] = str(sys.argv[1])

if len(sys.argv) == 3:
    settings.config['PathToXMLFile']['path'] = str(sys.argv[1])
    settings.config['PathToResultFile']['path'] = str(sys.argv[2])
    settings.writesettings()
    settings.config['PathToXMLFile']['path'] = str(sys.argv[1])
    settings.config['PathToResultFile']['path'] = str(sys.argv[2])

if settings.config.has_section('PathToXMLFile'):
    if not os.path.isfile(settings.config['PathToXMLFile']['path']):
        sys.exit("No such XML file. At least can't find it...")
else:
    settings.config['PathToXMLFile']['path'] = ''
    settings.writesettings()
    sys.exit("No PathToXMLFile is specified...")

xml.maketree(settings.config['PathToXMLFile']['path'])
#print(xml.root.tag)
#print(xml.root.attrib)
#xml.childs(xml.root[0])
hours.getplanopt(xml.root[0])
hours.appendexlist(settings.config)
for dis in xml.discomp(settings.config):
    hours.competenceslist(dis)
#for el in hours.compdict.items():
#    print(el)
for dis in xml.dis(settings.config):
    hours.sum(settings.config, dis)
#for el in hours.compdict.items():
#    print(el)
    #print(dis.tag, ' ', dis.attrib['Ном'], ' ', dis.keys())
#print('Hours: ', hours.hours)
#print('Semsh: ', hours.semsh)
#xml.childs(xml.root[0][6])
if settings.config.has_section('PathToResultFile'):
    if settings.config['PathToResultFile'].get('path', '') == '':
        path = 'result.xls'
    else:
        path = settings.config['PathToResultFile'].get('path', 'result.xls')
else:
    settings.config['PathToResultFile']['path'] = 'result.xls'
    settings.writesettings()
    path = settings.config['PathToResultFile'].get('path', 'result.xls')

finalsum = {'hours': 0,
            'Рук. маг': 0,
            'Лек': 0,
            'Лек. конс': 0,
            'Практ': 0,
            'Лаб': 0,
            'КонтрРаб': 0,
            'КурсРаб': 0,
            'КурсПр': 0,
            'РГР': 0,
            'Диф. зач': 0,
            'Зач': 0,
            'Экз': 0,
            'Экз. конс': 0,
            'ГЭК': 0,
            'ГАК': 0,
            'Рук. ВКР': 0,
            'Практики и НИР': 0
            }
#print(hours.semsh)
for sem, value in hours.semsh.items():
    #print(sem, '', value)
    for el, value in value.items():
        finalsum[el] += value
        finalsum[el] = round(finalsum[el], 2)
#print(path)
#if path.endswith('.txt'):
#    with open(path, 'w', encoding='utf-8-sig') as fp:
#        fp.write('Hours: ')
#        json.dump(round(hours.hours, 2), fp, sort_keys=True, indent=4, ensure_ascii=False)
#        fp.write('\nSumm: ')
#        json.dump(finalsum, fp, sort_keys=True, indent=4, ensure_ascii=False)
#        fp.write('\nSemsh: ')
#        json.dump(hours.semsh, fp, sort_keys=True, indent=4, ensure_ascii=False)

if not(path.endswith('.xls')) and not path.endswith('.txt'):
    filename, file_extension = os.path.splitext(path)
    #print(file_extension)
    if not file_extension == '':
        path = str.replace(path, file_extension, '.xls')
    else:
        path += '.xls'
    #print(path)
print(hours.semsh)
if path.lower().endswith('.xls') or path.lower().endswith('.xlsx'):
    xlswriter = XlsWriter()
    if not os.path.isdir(os.path.dirname(path)):
        path = str(os.path.dirname(settings.config['PathToXMLFile']['path'])) + '\\' + str(os.path.basename(path))
    xlswriter.writetoxls(finalsum, hours.semsh, path, os.path.splitext(settings.config['PathToXMLFile']['path'])[0].split('\\')[-1], hours.exlist, settings.config, xml.checkeduform('заочная'))

if not path.endswith('.txt'):
    filename, file_extension = os.path.splitext(path)
    #print(file_extension)
    if not file_extension == '':
        path = str.replace(path, file_extension, '.txt')
    else:
        path += '.txt'
    #print(path)
if path.lower().endswith('.txt'):
    with open(path, 'w', encoding='utf-8-sig') as fp:
        missingcomp = True
        for key in collections.OrderedDict(sorted(hours.compdict.items())):
            if not hours.compdict[key]['Матрица']:
                if missingcomp:
                    fp.write('Пустые компетенции:\n')
                    fp.write(hours.compdict[key]['Индекс'])
                    missingcomp = False
                    continue
                fp.write('; ' + hours.compdict[key]['Индекс'])
        if not missingcomp:
            fp.write('\n\n')
        #fp.write('\n')

        for key in collections.OrderedDict(sorted(hours.compdict.items())):
            fp.write('Компетенция: ' + hours.compdict[key]['Индекс'] + '\nСодержание: ' + hours.compdict[key]['Содержание'] + '\n')
            for semn in collections.OrderedDict(sorted(hours.compdict[key]['Матрица'].items())):
                fp.write('Семестр ' + str(semn) + ': ' + '; '.join(hours.compdict[key]['Матрица'][semn]) + '\n')
            fp.write('\n')

        if len(hours.diswithoutcomp) > 0:
            fp.write('Дисциплины без компетенций:\n')
            for el in hours.diswithoutcomp:
                fp.write('   ' + el + '\n')

#for child in xml.root[0][6]:
#    xml.findchilds(child)


