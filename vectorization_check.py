# -*- coding: utf-8 -*-


from dbfpy import dbf
import xlrd
import xlwt
import os
import collections

empty = u'Відсутня'
empty_2 = u'інформація відсутня'
db_files = []
grunt_files = []
grunt_correct_fields = ['ID', 'ND', 'DATESD', 'RDOR', 'VPOR', 'VPOS', 'DATEP',
    'SHAGG', 'NAGG', 'AREAAGG', 'NDPPOG']
prfrm_correct_fields = ['ID', 'KOATUU', 'NATOOBL', 'NATORAY', 'NATORAD',
    'TOPOCODE', 'AREAATORAD', 'NDPVMATO', 'ND', 'DATESD', 'RDOR', 'ORZVMATO',
    'DPR', 'NR', 'VPOR', 'VPOS', 'DATEP']
prfrmnp_correct_fields = ['ID', 'KOATUU', 'NATOOBL', 'NATORAY', 'NATORAD',
    'NATONP', 'NDPVMATO', 'ND', 'DATESD', 'RDOR', 'ORZVMATO', 'NR', 'VPOR',
    'VPOS', 'DATEP', 'AREAATORAY', 'TOPOCODE', 'DPR']
prfrm_files = []
prfrmnp_files = []
user_choise = -1
result_1 = open('./results/fields_check_grunt.txt', 'w')
result_2 = open('./results/fields_check_prfrm.txt', 'w')
result_3 = open('./results/fields_check_prfrmnp.txt', 'w')
result_4 = open('./results/empty_fields_grunt.txt', 'w')
result_5 = open('./results/empty_fields_prfrm.txt', 'w')
result_6 = open('./results/empty_fields_prfrmnp.txt', 'w')
workbook = xlrd.open_workbook('./AGG.xls')
worksheet = workbook.sheet_by_name('agg_codes')
db_files = []
style1 = xlwt.XFStyle()
wb = xlwt.Workbook(encoding='cp1251')
ws = wb.add_sheet('Report')


def generate_report():
    j = -1
    """for prfrm_file in prfrm_files:
        #dbsf
        print(prfrm_file)
        db = dbf.Dbf(prfrm_file)
        i = -1
        for rec in db:
            i += 1
            j += 1
            rec = db[i]
            ws.write(j, 0, prfrm_file)
            ws.write(j, 1, rec['KOATUU'])
            ws.write(j, 2, str(rec['AREAATORAD']))
            ws.write(j, 3, rec['NATORAD'].decode('cp1251'))
            if str(rec['KOATUU']) not in prfrm_file:
                ws.write(j, 4, 'ERROR: Check KOATUU')
        db.close()"""
    for prfrmnp_file in prfrmnp_files:
        #dbsf
        print(prfrmnp_file)
        db = dbf.Dbf(prfrmnp_file)
        i = -1
        for rec in db:
            i += 1
            j += 1
            rec = db[i]
            ws.write(j, 0, prfrmnp_file)
            ws.write(j, 1, rec['KOATUU'])
            ws.write(j, 2, str(rec['AREAATORAY']))
            ws.write(j, 3, rec['NATONP'].decode('cp1251'))
            if str(rec['KOATUU']) not in prfrmnp_file:
                ws.write(j, 4, 'ERROR: Check KOATUU')
        db.close()


def fill_prfrmnp():
    for prfrmnp_file in prfrmnp_files:
        print (prfrmnp_file)
        db = dbf.Dbf(prfrmnp_file)
        i = -1
        for rec in db:
            i = i + 1
            rec = db[i]
            rec['TOPOCODE'] = '81200000'
            rec.store()
            del rec
        db.close()


def fill_prfrm():
    for prfrm_file in prfrm_files:
        #dbsf
        print (prfrm_file)
        db = dbf.Dbf(prfrm_file)
        i = -1
        for rec in db:
            i = i + 1
            rec = db[i]
            rec['TOPOCODE'] = '81200000'
            rec.store()
            del rec
        db.close()


def fill_grunt():
    for grunt_file in grunt_files:
        print(grunt_file)
        db = dbf.Dbf(grunt_file)
        for rec in db:
            shagg_value = (rec['SHAGG'].split('+')[0]).strip()
            for row_num in range(0, worksheet.nrows - 1):
                # quoted in workbook to force string for cell value type
                cell_value = worksheet.cell(row_num, 0).value.strip('"')
                if shagg_value.decode('cp1251').startswith(cell_value):
                    nagg_value = worksheet.cell(row_num, 1).value
                    rec['NAGG'] = nagg_value.encode('cp1251')
                    rec.store()
        db.close()


def check_empty_grunt():
    for grunt_file in grunt_files:
        print(grunt_file)
        result_4.write(grunt_file)
        result_4.write('\n')
        db = dbf.Dbf(grunt_file)
        i = -1

        for rec in db:
            i += 1
            rec = db[i]
            is_mofified = False
            current = 'Line: ' + str(i + 2) + ' => '
            #if rec['ID'] == '':
                #current += ' ID'
            if rec['ND'] == '' or rec['ND'] == '-':
                current += ' ND'
                is_mofified = True
                #rec['ND'] = empty.encode('cp1251')
                #rec.store()
            if rec['DATESD'] == '' or rec['DATESD'] == '-':
                current += ' DATESD'
                is_mofified = True
            if rec['RDOR'] == '' or rec['RDOR'] == '-':
                current += ' RDOR'
                is_mofified = True
            if rec['VPOR'] == '' or rec['VPOR'] == '-':
                current += ' VPOR'
                is_mofified = True
            if rec['VPOS'] == '' or rec['VPOS'] == '-':
                current += ' VPOS'
                is_mofified = True
            if rec['DATEP'] == '' or rec['DATEP'] == '-':
                current += ' DATEP'
                is_mofified = True
            if rec['SHAGG'] == '' or rec['SHAGG'] == '-':
                current += ' SHAGG'
                is_mofified = True
            if rec['NAGG'] == '' or rec['NAGG'] == '-':
                current += ' NAGG (' + rec['SHAGG'] + ')'
                is_mofified = True
                #rec['NAGG'] = empty_2.encode('cp1251')
                #rec.store()
            if rec['AREAAGG'] == '' or rec['AREAAGG'] == '-':
                current += ' AREAAGG'
                is_mofified = True
            if rec['NDPPOG'] == '' or rec['NDPPOG'] == '-':
                current += ' NDPPOG'
                is_mofified = True
            if is_mofified == True:
                result_4.write(current)
                result_4.write('\n')
        db.close()
    #result_4.close()


def check_empty_prfrm():
    for prfrm_file in prfrm_files:
        print(prfrm_file)
        result_5.write(prfrm_file)
        result_5.write('\n')
        db = dbf.Dbf(prfrm_file)
        i = -1
        for rec in db:
            i += 1
            rec = db[i]
            current = 'Line: ' + str(i) + ' => '
            #if rec['ID'] == '':
            #    current += ' ID'
            if rec['KOATUU'] == '':
                current += ' KOATUU'
            if rec['NATOOBL'] == '':
                current += ' NATOOBL'
            if rec['NATORAY'] == '':
                current += ' NATORAY'
            if rec['NATORAD'] == '':
                current += ' NATORAD'
            if rec['TOPOCODE'] == '':
                current += ' TOPOCODE'
            if rec['AREAATORAD'] == '':
                current += ' AREAATORAD'
            if rec['NDPVMATO'] == '':
                current += ' NDPVMATO'
            if rec['ND'] == '':
                current += ' ND'
            if rec['DATESD'] == '':
                current += ' DATESD'
            if rec['RDOR'] == '':
                current += ' RDOR'
            if rec['ORZVMATO'] == '':
                current += ' ORZVMATO'
            if rec['DPR'] == '':
                current += ' DPR'
            if rec['NR'] == '':
                current += ' NR'
            if rec['VPOR'] == '':
                current += ' VPOR'
            if rec['VPOS'] == '':
                current += ' VPOS'
            if rec['DATEP'] == '':
                current += ' DATEP'
            result_5.write(current)
            result_5.write('\n')
        db.close()
    #result_5.close()


def check_empty_prfrmnp():
    for prfrmnp_file in prfrmnp_files:
        print(prfrmnp_file)
        result_6.write(prfrmnp_file)
        result_6.write('\n')
        db = dbf.Dbf(prfrmnp_file)
        i = -1
        for rec in db:
            i += 1
            rec = db[i]
            current = 'Line: ' + str(i) + ' => '
            if rec['ID'] == '':
                current += ' ID'
            if rec['KOATUU'] == '':
                current += ' KOATUU'
            if rec['NATOOBL'] == '':
                current += ' NATOOBL'
            if rec['NATORAY'] == '':
                current += ' NATORAY'
            if rec['NATORAD'] == '':
                current += ' NATORAD'
            if rec['NATONP'] == '':
                current += ' NATONP'
            if rec['NDPVMATO'] == '':
                current += ' NDPVMATO'
            if rec['ND'] == '':
                current += ' ND'
            if rec['DATESD'] == '':
                current += ' DATESD'
            if rec['RDOR'] == '':
                current += ' RDOR'
            if rec['ORZVMATO'] == '':
                current += ' ORZVMATO'
            if rec['NR'] == '':
                current += ' NR'
            if rec['VPOR'] == '':
                current += ' VPOR'
            if rec['VPOS'] == '':
                current += ' VPOS'
            if rec['DATEP'] == '':
                current += ' DATEP'
            if rec['AREAATORAY'] == '':
                current += ' AREAATORAY'
            if rec['TOPOCODE'] == '':
                current += ' TOPOCODE'
            if rec['DPR'] == '':
                current += ' DPR'
            result_6.write(current)
            result_6.write('\n')
        db.close()
    #result_6.close()


def check_grunt_fields():
    for grunt_file in grunt_files:
        print (grunt_file)
        db = dbf.Dbf(grunt_file)
        print((db.fieldNames))
        if collections.Counter(grunt_correct_fields) != collections.Counter(db.fieldNames):
            result_1.write(grunt_file)
            result_1.write('\n')
        #result_1.close()
        db.close()


def check_prfrm_fields():
    for prfrm_file in prfrm_files:
        print (prfrm_file)
        db = dbf.Dbf(prfrm_file)
        print((db.fieldNames))
        if collections.Counter(prfrm_correct_fields) != collections.Counter(db.fieldNames):
            result_2.write(prfrm_file)
            result_2.write('\n')
        #result_2.close()
        db.close()


def check_prfrmnp_fields():
    for prfrmnp_file in prfrmnp_files:
        print (prfrmnp_file)
        db = dbf.Dbf(prfrmnp_file)
        print((db.fieldNames))
        if collections.Counter(prfrmnp_correct_fields) != collections.Counter(db.fieldNames):
            result_3.write(prfrmnp_file)
            result_3.write('\n')
        #result_3.close()
        db.close()

for root, dirs, files in os.walk('./'):
    db_files += [os.path.join(root, name) for name in files if os.path.splitext(name)[1] == '.dbf']
    grunt_files += [os.path.join(root, name) for name in files if 'Grunt' in os.path.splitext(name)[0]]
    prfrm_files += [os.path.join(root, name) for name in files if 'PrFrm' == os.path.splitext(name)[0][-5::]]
    prfrmnp_files += [os.path.join(root, name) for name in files if 'PrFrmNP' in os.path.splitext(name)[0]]

while user_choise != 0:
    print('\n')
    print(' 1. Сheck the correctness of the fields in Grunt tables.')
    print(' 2. Сheck the correctness of the fields in PrFrm tables.')
    print(' 3. Сheck the correctness of the fields in PrFrmNp tables.')
    print(' 4. Сheck for empty fields in Grunt tables.')
    print(' 5. Сheck for empty fields in PrFrm tables.')
    print(' 6. Сheck for empty fields in PrFrmNp tables.')
    print(' 7. Fill \'NAGG\' in Grunt tables.')
    print(' 8. Fill \'Topocode\' in PrFrm tables with \'81200000\'.')
    print(' 9. Fill \'Topocode\' in PrFrmNp tables with \'81200000\'.')
    print('10. Generate PrFrm and PrFrmNp tables report.')
    print('\n')
    print(' 0. Exit.')
    user_choise = int(input('\nYour choise: '))
    if user_choise == 1:
        check_grunt_fields()
    if user_choise == 2:
        check_prfrm_fields()
    if user_choise == 3:
        check_prfrmnp_fields()
    if user_choise == 4:
        check_empty_grunt()
    if user_choise == 5:
        check_empty_prfrm()
    if user_choise == 6:
        check_empty_prfrmnp()
    if user_choise == 7:
        fill_grunt()
    if user_choise == 8:
        fill_prfrm()
    if user_choise == 9:
        fill_prfrmnp()
    if user_choise == 10:
        generate_report()
result_1.close()
result_2.close()
result_3.close()
result_4.close()
result_5.close()
result_6.close()
wb.save('./results/PrFrm_PrFrmNp_report.xls')

