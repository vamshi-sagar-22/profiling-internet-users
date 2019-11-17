import xlrd
import datetime as dt
import math
import scipy.stats
import os
import time
import xlwt
from tempfile import TemporaryFile

#start a time to get the execution time
start_time = time.time()

#open a workbook to write data into it
book = xlwt.Workbook()
sheet1 = book.add_sheet('sheet1')


#code below splits the time between 8 am and 5pm into 10 second intervals
#we will have 3240 entries in 10 seconds window
intervals = []
t = dt.timedelta(0, 0, 0, 0, 0, 8)
intervals.append(t)
for i in range(108):
    t = t + dt.timedelta(seconds=300)
    intervals.append(t)

#this function is to calculate spearman correlation co-efficient
def spear_cor(la, lb):
    return scipy.stats.spearmanr(la, lb)[0]

#this function calculates the z values for corresponding "r" values
def zvalues(r1a2a, r1a2b, r2a2b, N):
    rm2 = ((r1a2a ** 2) + (r1a2b ** 2)) / 2
    f = (1 - r2a2b) / (2 * (1 - rm2))
    h = (1 - f * rm2) / (1 - rm2)

    z1a2a = 0.5 * (math.log10((1 + r1a2a) / (1 - r1a2a)))
    z1a2b = 0.5 * (math.log10((1 + r1a2b) / (1 - r1a2b)))

    z = (z1a2a - z1a2b) * ((N - 3) ** 0.5) / (2 * (1 - r2a2b) * h)

    return z

#this function calculates p values for the given "z" values
def pvalues(z):
    p = 0.3275911
    a1 = 0.254829592
    a2 = -0.284496736
    a3 = 1.421413741
    a4 = -1.453152027
    a5 = 1.061405429

    sign = None
    if z < 0.01:
        sign = -1
    else:
        sign = 1

    x = abs(z) / (2 ** 0.5)
    t = 1 / (1 + p * x)
    erf = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * math.exp(-x * x)

    return 0.5 * (1 + sign * erf)
#this function takes a list(of packets in a day that are in between 8am and 5pm and returns average values for a day
def findaverage(l,rf):
    w= []
    if len(l)==0:
        for i in range(108):
            w.append(0)
        return w
    temp = []
    for i in range(0, 108):
        for j in range(0, len(l)):
            if rf[j]>=intervals[i] and rf[j]<intervals[i+1]:
                temp.append(l[j])
        if len(temp)==0:
            w.append(0)
        else:
            w.append(sum(temp)/len(temp))
        temp.clear()
    return w

#this function takes a sheet of an excel file and splits the packets based on the given conditions which are:
#1. remove the packets with duration =0
#2. remove the packets that are not in between 8 am and 5pm
#this function checks for each day in a week
#after this all the useful data is stored in lists and calls findaverage() function
#this funtion returns the average for week1
def userweek1avg(s1):
    a1= []
    a2 = []
    a3 = []
    a4 = []
    a5= []
    rf1 = []
    rf2 = []
    rf3 =[]
    rf4 = []
    rf5 = []
    for i in range(1, s1.nrows):
        hr = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).hour
        m = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).minute
        s = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).second
        ms = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).microsecond
        given = dt.timedelta(0, s, ms, 0, m, hr)
        date = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).day
        t1 = dt.timedelta(0, 0, 0, 0, 0, 8)
        t2 = dt.timedelta(0, 0, 0, 0, 0, 17)
        if s1.cell_value(i, 9) != 0:
            #checking whether the date is in week1 starting from 4 and ending at 8
            if given >= t1 and given < t2:
                if date == 4:
                    a1.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf1.append(given)
                if date == 5:
                    a2.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf2.append(given)
                if date == 6:
                    a3.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf3.append(given)
                if date == 7:
                    a4.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf4.append(given)
                if date == 8:
                    a5.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf5.append(given)
    a1 = findaverage(a1,rf1)
    a2 = findaverage(a2,rf2)
    a3 = findaverage(a3,rf3)
    a4 = findaverage(a4,rf4)
    a5 = findaverage(a5,rf5)
    x = a1+a2+a3+a4+a5
    return x

#this function takes a sheet of an excel file and splits the packets based on the given conditions which are:
#1. remove the packets with duration =0
#2. remove the packets that are not in between 8 am and 5pm
#this function checks for each day in a week
#after this all the useful data is stored in lists and calls findaverage() function
#this funtion returns the average for week2
def userweek2avg(s1):
    a1= []
    a2 = []
    a3 = []
    a4 = []
    a5= []
    rf1 = []
    rf2 = []
    rf3 =[]
    rf4 = []
    rf5 = []
    for i in range(1, s1.nrows):
        hr = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).hour
        m = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).minute
        s = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).second
        ms = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).microsecond
        given = dt.timedelta(0, s, ms, 0, m, hr)
        date = dt.datetime.fromtimestamp(s1.cell_value(i, 5) / 1000).day
        t1 = dt.timedelta(0, 0, 0, 0, 0, 8)
        t2 = dt.timedelta(0, 0, 0, 0, 0, 17)
        if s1.cell_value(i, 9) != 0:
            if given >= t1 and given < t2:
                # checking whether the date is in week2 starting from 11 and ending at 15
                if date == 11:
                    a1.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf1.append(given)
                if date == 12:
                    a2.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf2.append(given)
                if date == 13:
                    a3.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf3.append(given)
                if date == 14:
                    a4.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf4.append(given)
                if date == 15:
                    a5.append(s1.cell_value(i, 3) / s1.cell_value(i, 9))
                    rf5.append(given)
    a1 = findaverage(a1,rf1)
    a2 = findaverage(a2,rf2)
    a3 = findaverage(a3,rf3)
    a4 = findaverage(a4,rf4)
    a5 = findaverage(a5,rf5)
    x = a1+a2+a3+a4+a5
    return x

#parse the directory which contains all the excel files
l = os.listdir("U:/info sec/final project/drive-download-20190401T174327Z-001")

#find the combinations between all the files
for a in range(0, 54):
    for b in range(0, 54):

        file1 = "U:/info sec/final project/drive-download-20190401T174327Z-001/"+l[a]
        file2 = "U:/info sec/final project/drive-download-20190401T174327Z-001/"+l[b]

        print(" user {} ---> user {}".format(l[a], l[b]))


        w1 = xlrd.open_workbook(file1)
        w2 = xlrd.open_workbook(file2)
        s1 = w1.sheet_by_index(0)
        s2 = w2.sheet_by_index(0)


        #call the functions to calculate averages
        w1a = userweek1avg(s1)
        w2a = userweek2avg(s1)


        #check if file1 and file2 are same
        #if they are same copy the week1 data of user1 into week1 of user2
        #also copy the week2 data of user1 into week2 of user2
        if file1==file2:
            w1b = w1a
            w2b = w2a
        else:
            #if files are not same calculate the averages
            w1b = userweek1avg(s2)
            w2b = userweek2avg(s2)


        #here we call the functions to calculate spearmann correlation co-efficients
        r1a2a = spear_cor(w1a, w2a)
        r1a2b = spear_cor(w1a, w2b)
        r2a2b = spear_cor(w2a, w2b)

        #if any of the correlation values are same then change their values to 0.99
        if r1a2a == 1:
            r1a2a = 0.99
        if r1a2b == 1:
            r1a2b = 0.99
        if r2a2b == 1:
            r2a2b = 0.99
        #here we call the function to calculate "z value"
        z = zvalues(r1a2a, r1a2b, r2a2b, 16200)
        p1 =pvalues(z)
        print(" p value {}".format(p1))
        #now, write the values to the excel file
        sheet1.write(a,b,p1)


#give the name to the excel file
excel_filename = "p_300.xls"

#save the excel file
book.save(excel_filename)
book.save(TemporaryFile())

#print the execution time
print(time.time()-start_time)
