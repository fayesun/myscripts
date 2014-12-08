#!/usr/bin/python
from __future__ import print_function
import xlrd
import sys
import psycopg2
import datetime

reload(sys)
sys.setdefaultencoding('utf-8')

if __name__ == '__main__':
    conn = psycopg2.connect(database='spadm',user='spadm',password='spadm123', host='localhost')
    cur = conn.cursor()
    kvm_xls_file = sys.argv[1]
    data = xlrd.open_workbook(kvm_xls_file)
    table = data.sheet_by_name('Raw Data') 
    rows = table.nrows
    head = 'SPVM00'
    for i in range(rows)[1:]:
        if i == 1:
            if table.row_values(i+1)[0].strip(' '):
                labname = table.row_values(i)[0]
                labtype = 'VMware'

            else:
                labname = '{0}{1}'.format(table.row_values(i)[0], table.row_values(i)[1])
                if not table.row_values(i+2)[0].strip(' '):
                    labtype = 'KVM eSM'

                else:
                    labtype = 'KVM ATCA'

        elif i == (rows-1):
            if table.row_values(i)[0].strip(' '):
                labname = table.row_values(i)[0]
                labtype = 'VMware'

            else:
                labname = '{0}{1}'.format(head, table.row_values(i)[1])
                if not table.row_values(i-1)[0].strip(' '):
                    labtype = 'KVM eSM'

                else:
                    labtype = 'KVM ATCA'

        else:
            if table.row_values(i)[0].strip(' '):
                head = table.row_values(i)[0]
                if table.row_values(i+1)[0].strip(' '):
                    labname = head
                    labtype = 'VMWare'

                else:
                    labname = '{0}{1}'.format(head, table.row_values(i)[1])
                    try:
                        if table.row_values(i+2)[0].strip(' '):
                            labtype = 'KVM ATCA'

                        else:
                            labtype = 'KVM eSM'

                    except IndexError:
                        labtype = 'KVM ATCA'

            else:
                labname = '{0}{1}'.format(head, table.row_values(i)[1])
                if table.row_values(i+1)[0].strip(' ') and table.row_values(i-1)[0].strip(' '):
                    labtype = 'KVM ATCA'

                else:
                    labtype = 'KVM eSM'

        print('{0} {1}'.format(labname, labtype), sep=' ', end='')
        labuser = table.row_values(i)[5]
        if table.row_values(i)[3] == 'NA':
            datefrom = '1/1/2014'
        elif table.row_values(i)[3].strip(' ') == '2/30/2014' or table.row_values(i)[3].strip(' ') == '02/30/2014':
            datefrom = '2/28/2014'
        elif table.row_values(i)[3].strip(' ') == '2/31/2014' or table.row_values(i)[3].strip(' ') == '02/31/2014':
            datefrom = '2/28/2014'
        else:
            datefrom = table.row_values(i)[3]

        if table.row_values(i)[4] == 'NA':
            dateto = '6/30/2015'
        elif table.row_values(i)[4].strip(' ') == 'Long Term':
            dateto = '6/30/2015'
        elif table.row_values(i)[4].strip(' ') == '2/30/2015' or table.row_values(i)[4].strip(' ') == '02/30/2015':
            dateto = '2/28/2015'
        elif table.row_values(i)[4].strip(' ') == '2/31/2015' or table.row_values(i)[4].strip(' ') == '02/31/2015':
            dateto = '2/28/2015'
        else:
            dateto = table.row_values(i)[4]

        extno = str(table.row_values(i)[7]).split('.')[0]
        ptversion = table.row_values(i)[2].split()[0]
        spa = table.row_values(i)[9]
        release = table.row_values(i)[10]
        testtype = table.row_values(i)[8]
        customerdata = table.row_values(i)[11]
        featureid = str(table.row_values(i)[14]).split('.')[0]
        manager = table.row_values(i)[15]
        if table.row_values(i)[6].strip(' ') == 'N':
            free = 'no'
        elif table.row_values(i)[6].strip(' ') == 'Y':
            free = 'yes'

        email = 'test'
        try:
            ss7 = table.row_values(i)[2].split()[1]
        except IndexError:
            ss7 = ''
        print('{0} {1} {2} {3} {4} {5} {6} {7} {8} {9} {10} {11} {12} {13}'.format(labuser, datefrom, dateto, extno,ptversion,spa,release,testtype,customerdata,featureid,manager,free,email,ss7))

        cur.execute("INSERT INTO kvmlabusage (labname,labuser,datefrom,dateto,extno,ptversion,spa,release,testtype,customerdata,featureid,manager,labtype,free,ss7,email) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}')".format(labname,labuser,datefrom,dateto,extno,ptversion,spa,release,testtype,customerdata,featureid,manager,labtype,free,ss7,email))
    conn.commit()
    cur.close()
    conn.close()