# encoding=utf8
# -*- coding: utf-8 -*-

# --------------------------------------------------------------------
#  a crawler to search company information in their 8K form
#
#  first input excel file, use BeautifulSoup to parse and get basic information,
#  get the 10-K file, and filter, then analyize
#
#  input: input.xls
#  output: output.xls
#
#  write by Duowen
#
# --------------------------------------------------------------------


from BeautifulSoup import BeautifulSoup
import urllib2
import re
import string
import xlrd
import xlwt
import sys
import time

'''output_list = ['CIK', 'Filings', 'Period_of_Report', 'Filing_Date', 'Well_Known_Seasoned_Issuer',
               'Not_Required_to_File', 'Large_Accelerated_Filer', 'Accelerated_Filer', 'Non_Accelerated_Filer',
               'Smaller_Reporting_Company', 'Shell_Company', 'Proxy_Incorporated_by_Reference', 'Used_COSO',
               'Version_of_COSO']'''

output_list =['CIK', 'Filings', 'Period_of_Report', 'Filing_Date','Item7.0.1','Item8.0.1']

symbol = 'x'








# initial (load and write)


def progress(width, percent):
    print "%s %d%%\r" % (('%%-%ds' % width) % (width * percent / 100 * '='), percent),
    if percent >= 100:
        print
        sys.stdout.flush()


def open_url(url_address):
    global recursive_count
    try:
        tmp = urllib2.urlopen(url_address, timeout=5)
        result = tmp.read()
    except Exception, e:
        time.sleep(3)
        print url_address+'fail!'
        recursive_count += 1
        if recursive_count == 4:
            recursive_count = 0
            print "fail to recursive"
            return -2
        result = open_url(url_address)
        print str(e)
    return result


def get_cik_list():
    input_file = xlrd.open_workbook('input.xls')
    sheet = input_file.sheets()[0]
    cik_list = []
    for i in range(sheet.nrows - 1):
        cik_list += [int(sheet.cell(i + 1, 0).value)]
    return cik_list
    # tmp_list = [1961]
    # return tmp_list
#def get_date_list():
#    input_file = xlrd.open_workbook('input.xls')
#    sheet = input_file.sheets()[0]
#    date_list = []
#    for i in range(sheet.nrows - 1):
#        date_list += [int(sheet.cell(i + 1, 1).value)]
#    return date_list
    


def get_item_list(cik_num, filing_type):
    aim_web1 = 'https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=' + str(
        cik_num) + '&type=' + filing_type + '&dateb=&owner=exclude&start=0&count=100'

    ct = open_url(aim_web1)


    if ct == -2:
        return -2
    soup = BeautifulSoup(ct)

    aim_web2 = 'https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=' + str(
        cik_num) + '&type=' + filing_type + '&dateb=&owner=exclude&start=100&count=100'
    ct2 = open_url(aim_web2)
    if ct2 == -2:
        return -2
    soup2 = BeautifulSoup(ct2)


    table1 = soup.findAll('tr')

    flag = ct.find('No matching CIK')#check whether this CIK is empty
    if flag > 0 or len(table1) < 4:
        table = []

        return table1

    del table1[table1.__len__()-1]
    for i in range(3):
        del table1[0]
    table2 = soup2.findAll('tr')
    for i in range(3):
        del table2[0]

    table=table1+table2


    return table


def report_error(cik, filing_type):
    output_table.write(row_num, 0, cik)
    output_table.write(row_num, 1, filing_type)
    output_table.write(row_num, 2, 'No matching CIK')


def get_web_address(item, time_threshold):
    base_address = 'https://www.sec.gov'
    filing_date = item.contents[7].contents[0][0:4]+item.contents[7].contents[0][5:7]+item.contents[7].contents[0][8:10]
    #print item.contents[7].contents[0]
    if string.atoi(filing_date) <= time_threshold:
        return 0
    item_address = base_address+item.a['href']
    return item_address


def get_general(web_address, cik, filing_type):
    et = open_url(web_address)
    if et == -2:
        return -2
    soup_each = BeautifulSoup(et)
    form_content = soup_each.findAll('div', {'class': 'formGrouping'})

    # write the cik to excel
    output_table.write(row_num, 0, cik)
    # write the filing_type to excel
    output_table.write(row_num, 1, filing_type)
    # write the filing_date
    filing_date = form_content[0].contents[3].contents[0]
    output_table.write(row_num, 3, filing_date)
    # write the period_of_report
    period_of_report = form_content[1].contents[3].contents[0]
    output_table.write(row_num, 2, period_of_report)
    # print period_of_report

    # search the link of filing_type
    tr_all = soup_each.findAll('tr', {})
    flag = False
    tr = 0
    for tr in tr_all:
        for td in tr:
            if not td.string:
                continue
            if filing_type in td.string:
                flag = True
                break
        if flag:
            break

    # actually I do not consider the case the tr_all has nothing
    file_name = tr.contents[5].a.string
    if file_name is None:
        return -1
    file_address = web_address[0:web_address.rindex("/")+1]+file_name
    return file_address


def clean(document, document_type):
    if document_type == 0:
        soup = BeautifulSoup(document)
        document = soup.getText()
        document = document.replace('&#160;', ' ')
        document = document.replace('&#163;', ' ')
        document = document.replace('&#254;', ' x ')
        document = document.replace('&#253;', ' x ')
        document = document.replace('&#120;', ' x ')
        document = document.replace('&#9746;', ' x ')
        document = document.replace('&nbsp', ' ')
        document = document.replace(';', ' ')
    document = document.replace('\t', ' ')
    document = document.replace('\r\n', ' ')
    document = document.replace('\n', ' ')
    document = document.replace('\r', ' ')
    document = re.sub(' +', ' ', document)
    #document = document.lower()
    return document


def item701801(ft):
    p1 = ft.find('Item 7.01',0)
    p2 = ft.find('Item 8.01',0)
    p3 = ft.find('Item 9.01',0)
    p4 = ft.find('SIGNATURES',0)

    print p1,p2,p3,p4
    if (p1!=-1):#find item 7.01
        if(p2!=-1):
            start = ft.find('Item 7.01',p1)+len('Item 7.01')
            stop = p2
            tmpft = clean(ft[start:stop],0)
            output_table.write(row_num,4, tmpft)
            print(7.01)

        elif(p3!=-1):
            start = ft.find('Item 7.01',p1)+len('Item 7.01')
            stop = p3
            tmpft = clean(ft[start:stop],0)
            output_table.write(row_num,4, tmpft)
            print(7.01)
        else:
            start = ft.find('Item 7.01',p1)+len('Item 7.01')
            stop = p4
            tmpft = clean(ft[start:stop],0)
            output_table.write(row_num,4, tmpft)
            print(7.01)


    if (p2!=-1):#find item 8.01
        if(p3!=-1):
            start = ft.find('Item 8.01',p2)+len('Item 8.01')
            stop = p3
            tmpft = clean(ft[start:stop],0)
            output_table.write(row_num,5, tmpft)

        else:
            start = ft.find('Item 8.01',p2)+len('Item 8.01')
            stop = p4
            tmpft = clean(ft[start:stop],0)
            output_table.write(row_num,5, tmpft)












def well_season_issue(ft):
    file_base = ft.find('well-known seasoned issuer')
    first = ft.find('yes', file_base)
    second = ft.find('no', first)
    third = ft.find('x', first)
    global symbol
    # if not use x as a symbol use s as a symbol
    if third < 0 or third > second + 5:
        third = ft.find('s', first + 3)
        symbol = 's'
    if first <= third <= second:
        output_table.write(row_num, 4, 1)
    elif third >= second:
        output_table.write(row_num, 4, 0)
    else:
        output_table.write(row_num, 4, 'NaN')


def not_require_file(ft):
    file_base = ft.find('not required to file')
    first = ft.find('yes', file_base)
    second = ft.find('no', first)
    third = ft.find(symbol, first+3)
    if first <= third <= second:
        output_table.write(row_num, 5, 1)
    elif third >= second:
        output_table.write(row_num, 5, 0)
    else:
        output_table.write(row_num, 5, 'NaN')


def registrant_type(ft):
    if ft.find('shell company') > ft.find('12b-2'):
        file_base = ft.find('12b-2')
    else:
        file_base = ft.rfind('large accelerated', 0, ft.find('shell company'))

    laf = ft.find('large accelerated filer', file_base)
    laf_x = ft.find(symbol, laf+22)
    # laf_s = ft.find('filter s', laf)
    af = ft.find('accelerated filer', laf + 10)
    af_x = ft.find(symbol, af+16)
    # af_s = ft.find('filter s', af)
    naf = ft.find('non-accelerated filer', af)
    naf_x = ft.find(symbol, naf+19)
    # naf_s = ft.find('filter s', naf)
    sr = ft.find('smaller reporting company', file_base)
    sr_x = ft.find(symbol, sr+24)
    # sr_s = ft.find('company s', sr)
    file_end = ft.find('shell company', file_base)
    filter_flag = 0
    if file_base <= laf <= file_end:
        if file_base <= af <= file_end:
            if laf <= laf_x <= af:
                output_table.write(row_num, 6, 1)
                output_table.write(row_num, 7, 0)
                output_table.write(row_num, 8, 0)
                output_table.write(row_num, 9, 0)
                filter_flag = 1
    if file_base <= af <= file_end:
        if file_base <= naf <= file_end:
            if af <= af_x <= naf:
                output_table.write(row_num, 6, 0)
                output_table.write(row_num, 7, 1)
                output_table.write(row_num, 8, 0)
                output_table.write(row_num, 9, 0)
                filter_flag = 1
    if file_base <= naf <= file_end:
        if file_base <= sr <= file_end:
            if naf <= naf_x <= sr:
                output_table.write(row_num, 6, 0)
                output_table.write(row_num, 7, 0)
                output_table.write(row_num, 8, 1)
                output_table.write(row_num, 9, 0)
                filter_flag = 1
    if file_base <= sr <= file_end:
        if sr <= sr_x <= file_end:
            output_table.write(row_num, 6, 0)
            output_table.write(row_num, 7, 0)
            output_table.write(row_num, 8, 0)
            output_table.write(row_num, 9, 1)
            filter_flag = 1
    if sr < file_base or sr > file_end:
        if file_base <= naf <= file_end:
            if naf <= naf_x <= file_end:
                output_table.write(row_num, 6, 0)
                output_table.write(row_num, 7, 0)
                output_table.write(row_num, 8, 1)
                output_table.write(row_num, 9, 0)
                filter_flag = 1
    if filter_flag == 0:
        output_table.write(row_num, 6, 'NaN')
        output_table.write(row_num, 7, 'NaN')
        output_table.write(row_num, 8, 'NaN')
        output_table.write(row_num, 9, 'NaN')

    file_base = ft.find('shell company')
    first = ft.find('yes', file_base)
    second = ft.find('no', first)
    third = ft.find(symbol, first + 3)
    if (first <= third <= second) == first:
        output_table.write(row_num, 10, 1)
    elif third >= second:
        output_table.write(row_num, 10, 0)
    else:
        output_table.write(row_num, 10, 'NaN')


def incorporated_by_reference(ft):
    file_base = ft.find('documents incorporated by reference')
    if file_base == -1:
        output_table.write(row_num, 11, 'can not find documents incorporated by reference')
        return

    first = ft.find('proxy statement', file_base)
    second = ft.find('definitive proxy statement', file_base)
    third = ft.find('def 14a', file_base)

    if (first > 0 and first - file_base < 500) or \
            (second > 0 and second - file_base < 500) or (third > 0 and third - file_base < 500):
        output_table.write(row_num, 11, 'Yes')
        return

    second = ft.find('none', file_base)
    if 0 < second < file_base + 500:
        output_table.write(row_num, 11, 'No')
        return

    output_table.write(row_num, 11, 'Both yes and no can not be found.')
    return


def coso(ft):
    ft = ft.replace(', 2013', ' ')
    start = ft.rfind('item 9a.')
    if start < 0:
        start = ft.rfind('item9a')
    end = ft.rfind('item 10.')
    if start < 0:
        output_table.write(row_num, 12, 'can not find item 9A.')
        return
    if end < 0:
        output_table.write(row_num, 12, 'can not find item 10.')
        return

    first = ft.find('committee of sponsoring organizations of the treadway commission', start)
    second = ft.find('coso', start)
    third = ft.find('internal control-integrated framework', start)
    fourth = ft.find('internal control integrated framework', start)

    coso_flag = True
    if (0 < first < end) or (0 < second < end) or (0 < third < end) or (0 < fourth < end):
        output_table.write(row_num, 12, 'Yes')
    else:
        output_table.write(row_num, 12, 'No')
        coso_flag = False


    if coso_flag:
        first = ft.find('1992', start)
        second = ft.find('2013', start)

        if 0 < first < end:
            output_table.write(row_num, 13, 1992)
        elif 0 < second < end:
            output_table.write(row_num, 13, 2013)


def main(filing_type, time_threshold):
    # for each cik get the web list
    print "start "+filing_type+"collection"
    cik_list = get_cik_list()
    output_file = xlwt.Workbook()
    global output_table
    global symbol
    output_table = output_file.add_sheet('Sheet1', cell_overwrite_ok=True)
    for i in range(len(output_list)):
        output_table.write(0, i, output_list[i])

    # for each available item get basic information
    global row_num
    global recursive_count
    recursive_count = 0
    row_num = 1
    cik_count = 0
    cik_total = len(cik_list)
    for cik in cik_list:
        #for date in date_list:
            item_list = get_item_list(cik, filing_type)
            if item_list == -2:
                continue

            # check if no filings
            if len(item_list) == 0:
                report_error(cik, filing_type)
                print 'error'
                row_num += 1
                continue

            for item in item_list:


                symbol = 'x'

                if item.contents[1].contents[0]!='8-K':
                    continue
                form_web = get_web_address(item, time_threshold)
                
                # form_web = "https://www.sec.gov/Archives/edgar/data/1800/000104746913003504/0001047469-13-003504-index.htm"
                if form_web == 0:
                    continue
                document_web = get_general(form_web, cik, filing_type)
                print document_web
                if document_web == -1 or document_web == -2:
                    continue
                # get the general information and return if the document is html(0) or txt(1) or other(2)
                if '.htm' in document_web:
                    file_type = 0
                elif '.txt' in document_web:
                    file_type = 1
                else:
                    file_type = 2



                # go to the 8-K file

                ft = open_url(document_web)
                #print ft
                if ft == -2:
                    continue
                document = clean(ft, file_type)
                #print document
                item701801(document)
                #well_season_issue(document)
                #not_require_file(document)
                #registrant_type(document)
                #incorporated_by_reference(document)
                #coso(document)
                row_num += 1
                #if (row_num==10):break
    cik_count += 1
    print "%d of %d" % (cik_count, cik_total)
    #if cik_count % 100 == 0:
    #   output_file.save('output_'+filing_type+'.xls')

    output_file.save('output_'+filing_type+'.xls')

main('8-K', '20040101')
#main('10-KSB', '20000101')
