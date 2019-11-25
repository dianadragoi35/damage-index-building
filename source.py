#import matplotlib.pyplot as plt
import os
import xlwt
from xlutils.copy import copy as xl_copy
from xlrd import open_workbook
import math
import shutil
import sys


#fisierul excel cu denumirea etaj_x su un sheet pt fiecare etaj
def output(filename, sheet, list0, list1, list2, list3, list4, list5):
    if os.path.isfile(filename):
        rb = open_workbook(filename, formatting_info=True)
        book = xl_copy(rb)
        sh = book.add_sheet(sheet)

    else:
        book = xlwt.Workbook()
        sh = book.add_sheet(sheet)

    col0_name = 't'
    col1_name = 'Eha(t)'
    col2_name = 'Ehb(t)'
    col3_name = 'Dis_a(t)'
    col4_name = 'Dis_b(t)'
    col5_name = 'Die(t)'
    col6_name = 'Ehe(t)'

    sh.write(0, 0, col0_name)
    sh.write(0, 1, col1_name)
    sh.write(0, 2, col2_name)
    sh.write(0, 3, col3_name)
    sh.write(0, 4, col4_name)
    sh.write(0, 5, col5_name)
    sh.write(0, 6, col6_name)

    for m, e0 in enumerate(list0):
        sh.write(m + 1, 0, e0)

    for m, e1 in enumerate(list1):
        sh.write(m + 1, 1, e1)

    for m, e2 in enumerate(list2):
        sh.write(m + 1, 2, e2)

    for m, e3 in enumerate(list3):
        sh.write(m + 1, 3, e3)

    for m, e4 in enumerate(list4):
        sh.write(m + 1, 4, e4)

    for m, e5 in enumerate(list5):
        sh.write(m + 1, 5, e5)

    for m in range(0, len(list1)):
        sh.write(m + 1, 6, xlwt.Formula("'"+sheet+"'!B"+str(m+2)+"+'"+sheet+"'!C"+str(m+2)))

    book.save(filename)

cols, rows = 31, 6
##  GRINZI  ##
eha = [[[0] for x in range(cols)] for y in range(rows)]
ehb = [[[0] for x in range(cols)] for y in range(rows)]
ehe = [[[0] for x in range(cols)] for y in range(rows)]
dis_a = [[[0] for x in range(cols)] for y in range(rows)]
dis_b = [[[0] for x in range(cols)] for y in range(rows)]
die = [[[0] for x in range(cols)] for y in range(rows)]

##  STALPI  ##
ehx_a = [[[0] for x in range(cols)] for y in range(rows)]
ehx_b = [[[0] for x in range(cols)] for y in range(rows)]
ehy_a = [[[0] for x in range(cols)] for y in range(rows)]
ehy_b = [[[0] for x in range(cols)] for y in range(rows)]
dis_xa = [[[0] for x in range(cols)] for y in range(rows)]
dis_xb = [[[0] for x in range(cols)] for y in range(rows)]
dis_ya = [[[0] for x in range(cols)] for y in range(rows)]
dis_yb = [[[0] for x in range(cols)] for y in range(rows)]
dis_a_st = [[[0] for x in range(cols)] for y in range(rows)]
dis_b_st = [[[0] for x in range(cols)] for y in range(rows)]
die_st = [[[0] for x in range(cols)] for y in range(rows)]

# ------------ GRINZI -------------#

rulari = os.listdir("data")
#pentru fiecare rulare
for rulare in rulari:
    if os.path.isdir("exported\\"+rulare): ##daca fisierul exista deja
        shutil.rmtree("exported\\"+rulare, ignore_errors=True) ##sterg fisierul cu rularea si tot continulul lui
    os.mkdir("exported\\"+rulare) ##creez un folder cu denumirea rularii (RUL 1, RUL 2...)
    etaje = os.listdir("data\\"+rulare) ##iau toate etajele din rulare (1, 2, 3...)

    f = open("data\\"+rulare+"\data_beam.txt", "r")

    #file - fisierul data_beam din fiecare rulare
    file = f.read()

    #parts - impartim fisierul in blocuri pt fiecare grinda in functie de keywordul '--- member properties'
    parts = file.split('--- member properties')
    my = []
    my.append(-1)
    for part in parts:
        top_rebars = part.split('moment from top rebars')
        if len(top_rebars) > 1:
            aux = top_rebars[1].split('My =  ')
            if len(aux) > 1:
                my.append(float(aux[1].split(' ')[0]))

    nr_grinda = 1
    # pentru fiecare etaj cautam fisierele text cu grinzi (response_member21, ...)
    for etaj in range(1, 6):
        grinzi = os.listdir("data\\" + rulare +"\\"+ str(etaj) + "\Grinzi")
        print('Salvare date grinzi pentru etajul ' + str(etaj) + '...')
        k = 0  ## al catelea element grinda este pe etaj
        #pentru fiecare grinda
        for grinda in grinzi:
            filename = "data\\" + rulare + "\\" + str(etaj) + "\\Grinzi\\" + grinda
            f = open(filename, "r")
            g = open(filename, "r")

            lineList = g.readlines()
            last_line = lineList[len(lineList) - 1]
            column = last_line.split()
            t_f = float(column[0])
            rya_f = format(float(column[1]), 'f')
            mya_f = float(column[2])
            uya_f = float(column[3])
            ryb_f = format(float(column[7]), 'f')
            myb_f = float(column[8])
            uyb_f = float(column[9])

            g.close()

            ru = 0.02
            beta = 0.15

            with f as fp:
                # skip primele 2 linii (headere)
                next(fp)
                next(fp)
                line = fp.readline()

                t = []
                rya = []
                mya = []
                uya = []
                ryb = []
                myb = []
                uyb = []

                while line:
                    column = line.split()
                    t.append(float(column[0]))
                    rya.append(float(format(abs(float(column[1])), 'f')))
                    mya.append(abs(float(column[2])))
                    uya.append(abs(float(column[3])))
                    ryb.append(float(format(abs(float(column[7])), 'f')))
                    myb.append(abs(float(column[8])))
                    uyb.append(abs(float(column[9])))

                    # trec la urmatoarea linie
                    line = fp.readline()

            # integrala de la t la tf din M(t)*R(t)
            # Eh(t-tf) = suma de la i=t la i=tf din M(i) * U(i) * (t de i+1 - t de i)
            for index in range(0, len(t) - 1):
                eha_sum = 0
                ehb_sum = 0
                for index2 in range(0, index):
                    if uya[index2] >= 1:
                        eha_sum += mya[index2] * rya[index2] * (t[index2 + 1] - t[index2])
                    if uyb[index2] >= 1:
                        ehb_sum += myb[index2] * ryb[index2] * (t[index2 + 1] - t[index2])
                eha[etaj-1][k].append(eha_sum)
                ehb[etaj-1][k].append(ehb_sum)
                ehe[etaj-1][k].append(eha[etaj-1][k][index] + ehb[etaj-1][k][index])

            #demage index sectiune

            for index in range(0, len(t) - 1):
                dis_a[etaj-1][k].append(max(rya) / ru + beta * (eha[etaj-1][k][index] / (my[nr_grinda] * ru)))
                dis_b[etaj-1][k].append(max(ryb) / ru + beta * (ehb[etaj-1][k][index] / (my[nr_grinda] * ru)))

                if ((eha[etaj-1][k][index] + ehb[etaj-1][k][index]) != 0):
                    aux = dis_a[etaj-1][k][index] * eha[etaj-1][k][index] / (eha[etaj-1][k][index] + ehb[etaj-1][k][index]) + dis_b[etaj-1][k][index] * ehb[etaj-1][k][index] / (
                                eha[etaj-1][k][index] + ehb[etaj-1][k][index])
                    die[etaj-1][k].append(aux)
                else:
                    die[etaj-1][k].append(0)

            print('Salvare date pentru grinda ' + str(nr_grinda) + '...')
            output('exported\\'+rulare+'\grinzi_etaj_' + str(etaj) + '.xls', 'grinda ' + str(nr_grinda), t[:-1], eha[etaj-1][k], ehb[etaj-1][k], dis_a[etaj-1][k], dis_b[etaj-1][k], die[etaj-1][k])
            nr_grinda = nr_grinda + 1

        k = k + 1

    # ------------ END GRINZI -------------#

    # ------------ STALPI -----------------#

    f = open("data\\"+rulare+"\data_column.txt", "r")

    # file - fisierul data_column
    file = f.read()
    beta = 0.15
    # parts - impartim fisierul in blocuri pt fiecare grinda in functie de keywordul '--- member properties'
    parts = file.split('--- member properties')
    myx = []
    myy = []
    rpx = []
    rpy = []

    for part in parts:
        moment = part.split('moment')
        if len(moment) > 1:
            aux = moment[1].split('My_y =  ')
            if len(aux) > 1:
                myy.append(float(aux[1].split(' ')[0]))
            aux = moment[1].split('My_x =  ')
            if len(aux) > 1:
                myx.append(float(aux[1].split(' ')[0]))
            aux = moment[1].split('Rpy_y =  ')
            if len(aux) > 1:
                rpy.append(float(aux[1].split(' ')[0]))
            aux = moment[1].split('Rpy_x =  ')
            if len(aux) > 1:
                rpx.append(float(aux[1].split(' ')[0]))

    stalpi = []
    nr_stalp = 1
    k = 0 ## al catelea element stalp este pe etaj
    for etaj in range(1, 6):
        stalpi = os.listdir("data\\" + rulare + "\\" + str(etaj) + "\Stalpi")  ##fisierele cu stalpi
        print('Salvare date stalpi pentru etajul ' + str(etaj) + '...')
        for stalp in stalpi:
            filename = "data\\" + rulare + "\\" + str(etaj) + "\\Stalpi\\" + stalp
            f = open(filename, "r")
            with f as fp:
                # skip primele 2 linii (headere)
                next(fp)
                next(fp)
                line = fp.readline()

                t = []
                rya = []
                mya = []
                uya = []
                ryb = []
                myb = []
                uyb = []
                rxa = []
                mxa = []
                uxa = []
                rxb = []
                mxb = []
                uxb = []

                while line:
                    column = line.split()
                    t.append(float(column[0]))
                    rya.append(float(format(abs(float(column[1])), 'f')))
                    mya.append(abs(float(column[2])))
                    uya.append(abs(float(column[3])))
                    ryb.append(float(format(abs(float(column[4])), 'f')))
                    myb.append(abs(float(column[5])))
                    uyb.append(abs(float(column[6])))
                    rxa.append(float(format(abs(float(column[7])), 'f')))
                    mxa.append(abs(float(column[8])))
                    uxa.append(abs(float(column[9])))
                    rxb.append(float(format(abs(float(column[10])), 'f')))
                    mxb.append(abs(float(column[11])))
                    uxb.append(abs(float(column[12])))

                    # trec la urmatoarea linie
                    line = fp.readline()

            # integrala de la t la tf din M(t)*R(t)
            # Eh(t-tf) = suma de la i=t la i=tf din M(i) * U(i) * (t de i+1 - t de i)
            for index in range(0, len(t) - 1):
                ehxa_sum = 0
                ehxb_sum = 0
                ehya_sum = 0
                ehyb_sum = 0
                for index2 in range(0, index):
                    if uxa[index2] >= 1:
                        ehxa_sum += mxa[index2] * rxa[index2] * (t[index2 + 1] - t[index2])
                    if uxb[index2] >= 1:
                        ehxb_sum += mxb[index2] * rxb[index2] * (t[index2 + 1] - t[index2])
                    if uya[index2] >= 1:
                        ehya_sum += mya[index2] * rya[index2] * (t[index2 + 1] - t[index2])
                    if uyb[index2] >= 1:
                        ehyb_sum += myb[index2] * ryb[index2] * (t[index2 + 1] - t[index2])
                ehx_a[etaj-1][k].append(ehxa_sum)
                ehx_b[etaj-1][k].append(ehxb_sum)
                ehy_a[etaj-1][k].append(ehya_sum)
                ehy_b[etaj-1][k].append(ehyb_sum)

            # demage index sectiune

            for index in range(0, len(t) - 1):
                rux = 5 * rpx[nr_stalp]
                ruy = 5 * rpy[nr_stalp]

                dis_xa[etaj-1][k].append(max(rxa) / rux + beta * (ehx_a[etaj-1][k][index] / (myx[nr_stalp] * rux)))
                dis_xb[etaj-1][k].append(max(rxb) / rux + beta * (ehx_b[etaj-1][k][index] / (myx[nr_stalp] * rux)))
                dis_ya[etaj-1][k].append(max(rya) / ruy + beta * (ehy_a[etaj-1][k][index] / (myy[nr_stalp] * ruy)))
                dis_yb[etaj-1][k].append(max(ryb) / ruy + beta * (ehy_b[etaj-1][k][index] / (myy[nr_stalp] * ruy)))

                eha.append(math.sqrt(ehx_a[etaj-1][k][index] * ehx_a[etaj-1][k][index] + ehy_a[etaj-1][k][index] * ehy_a[etaj-1][k][index]))
                ehb.append(math.sqrt(ehx_b[etaj-1][k][index] * ehx_b[etaj-1][k][index] + ehy_b[etaj-1][k][index] * ehy_b[etaj-1][k][index]))

                # demage index sectiune stalpi
                dis_a.append(math.sqrt(dis_xa[etaj-1][k][index] * dis_xa[etaj-1][k][index] + dis_ya[etaj-1][k][index] * dis_ya[etaj-1][k][index]))
                dis_b.append(math.sqrt(dis_xb[etaj-1][k][index] * dis_xb[etaj-1][k][index] + dis_yb[etaj-1][k][index] * dis_yb[etaj-1][k][index]))

                if ((eha[etaj-1][k][index] + ehb[etaj-1][k][index]) != 0):
                    aux = dis_a[etaj-1][k][index] * eha[etaj-1][k][index] / (eha[etaj-1][k][index] + ehb[etaj-1][k][index]) + dis_b[etaj-1][k][index] * ehb[etaj-1][k][index] / (
                            eha[etaj-1][k][index] + ehb[etaj-1][k][index])
                    die[etaj-1][k].append(aux)
                else:
                    die[etaj-1][k].append(0)

            print('Salvare date pentru grinda ' + str(nr_stalp) + '...')
            output('exported\\'+rulare+'\stalpi_etaj_' + str(etaj) + '.xls', 'grinda ' + str(nr_stalp), t[:-1], eha, ehb, dis_a, dis_b, die)
            nr_stalp = nr_stalp + 1

        k = k + 1

        #damage index etaj ??


    # plotarea datelor
    ''' 
    fig, axs = plt.subplots(2, 2)
    axs[0, 0].plot(t[:-1], eha)
    axs[0, 0].set_title('Eha(t)')
    axs[0, 1].plot(t[:-1], ehb, 'tab:orange')
    axs[0, 1].set_title('Ehb(t)')
    axs[1, 0].plot(t[:-1], dis_a, 'tab:green')
    axs[1, 0].set_title('Demage_index_a(t)')
    axs[1, 1].plot(t, rya, 'tab:red')
    axs[1, 1].set_title('rya(t)')
    plt.show()
    '''



