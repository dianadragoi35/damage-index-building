import matplotlib.pyplot as plt
import os
import xlwt
from xlutils.copy import copy as xl_copy
from xlrd import open_workbook

def output(filename, sheet, list0, list1, list2, list3, list4):

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

    sh.write(0, 0, col0_name)
    sh.write(0, 1, col1_name)
    sh.write(0, 2, col2_name)
    sh.write(0, 3, col3_name)
    sh.write(0, 4, col4_name)

    for m, e0 in enumerate(list0):
        sh.write(m+1, 0, e0)

    for m, e1 in enumerate(list1):
        sh.write(m+1, 1, e1)

    for m, e2 in enumerate(list2):
        sh.write(m+1, 2, e2)

    for m, e3 in enumerate(list3):
        sh.write(m+1, 3, e3)

    for m, e4 in enumerate(list4):
        sh.write(m+1, 4, e4)

    book.save(filename)


f = open("data\\1\data_beam.txt", "r")

#file - fisierul data_beam
file = f.read()

#parts - impartim fisierul in blocuri pt fiecare grinda in functie de keywordul '--- member properties'
parts = file.split('--- member properties')
my = []
print(len(parts))
for part in parts:
    top_rebars = part.split('moment from top rebars')
    if len(top_rebars) > 1:
        aux = top_rebars[1].split('My =  ')
        if len(aux) > 1:
            my.append(float(aux[1].split(' ')[0]))
print('My(t)=')
print(my)

#pentru fiecare etaj
for etaj in range(1,6):
    grinzi = os.listdir("data\\"+str(etaj)+"\grinzi") ##fisierele cu grinzi
    stalpi = os.listdir("data\\"+str(etaj)+"\stalpi") ##fisierele cu stalpi


grinzi = os.listdir("data\\1\grinzi") ##etaj dinamic
etaj = 1

nr_grinda = 1
for grinda in grinzi:
    filename = "data\\1\\grinzi\\"+grinda
    f = open(filename, "r")
    g = open(filename, "r")

    lineList = g.readlines()
    last_line = lineList[len(lineList)-1]
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
        #skip primele 2 linii (headere)
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

            #trec la urmatoarea linie
            line = fp.readline()

    eha = []
    ehb = []
    # integrala de la t la tf din M(t)*R(t)
    # Eh(t-tf) = suma de la i=t la i=tf din M(i) * U(i) * (t de i+1 - t de i)
    for index in range(0, len(t)-1):
        eha_sum = 0
        ehb_sum = 0
        for index2 in range(0, index):
            if uya[index2] >= 1:
                eha_sum += mya[index2] * rya[index2] * (t[index2+1] - t[index2])
                ehb_sum += myb[index2] * ryb[index2] * (t[index2+1] - t[index2])
        eha.append(eha_sum)
        ehb.append(ehb_sum)

    #demage index sectiune
    dis_a = []
    dis_b = []
    for index in range(0, len(t)-1):
        if uya[index] >= 1:
            dis_a.append(rya[index]/ru + beta * (eha[index]/(my[0] * ru)))
        else:
            dis_a.append(0)
        if uyb[index] >= 1:
            dis_b.append(ryb[index]/ru + beta * (ehb[index] / (my[1] * ru)))
        else:
            dis_b.append(0)
    print('Salvare date pentru grinda '+str(nr_grinda)+'...')
    output('etaj_' + str(etaj) + '.xls', 'grinda ' + str(nr_grinda), t[:-1], eha, ehb, dis_a, dis_b)
    nr_grinda = nr_grinda + 1

    '''
    print('Dis_a('+str(nr_grinda)+')= ')
    print(dis_a)
    print('Dis_b(' + str(nr_grinda) + ')= ')
    print(dis_b)   
    '''

    #plotarea datelor
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

#for ax in axs.flat:
#    ax.set(xlabel='Timp(t)')





