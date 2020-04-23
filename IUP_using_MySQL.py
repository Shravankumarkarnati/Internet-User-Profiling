#!/usr/bin/python3

import os
import pandas as pd
import MySQLdb
import time
import scipy.stats
import math


def csv_from_excel(xlfile,csvfile):
	data_xls = pd.read_excel(xlfile, 'Sheet1', index_col=None)
	data_xls.to_csv(csvfile, encoding='utf-8')

def csv_to_mysql(load_sql,result = False):
	mydb = MySQLdb.connect(host='localhost',
							database='info',
							user='infosec',
							password='95050')
	cur = mydb.cursor()
	try:
		command = cur.execute(load_sql)
		lis = list(cur.fetchall())
		mydb.commit()
		mydb.close()
		return lis
	except (MySQLdb.Error,TypeError) as e:
		return None

def create_dValues():
	create_query = "create table dValues (id int NOT NULL auto_increment,flag int NULL,primary key(id));"
	csv_to_mysql(create_query)
	for i in range(1,45001):
		insert_query = "insert into dValues (flag) VALUES ({})".format(i)
		csv_to_mysql(insert_query)

def avg(d, num):
	# finds the average doctets/Duration value for a day using all the Doctets/Duration values and time slot
	window = []
	av = []
	array = [x[0] for x in d]
	values = [x[1] for x in d]
	day_times =[80000,170001]
	if num == 10:
		for i in range(day_times[0], day_times[1], 10):
			window.append(i)
	if num == 227:
		for i in range(day_times[0], day_times[1], 227):
			window.append(i)
	if num == 300:
		for i in range(day_times[0], day_times[1], 300):
			window.append(i)

	while (len(window) > 1):
		x = 0
		y = 0
		for l in range(len(array)):
			if window[0] <= float(array[l]%pow(10,6)) < window[1]:
				x = values[l] + x
				y += 1
		if y > 0 and x > 0:
			av.append(x/y)
		else:
			av.append(x)
		window.pop(0)
	return av



def findD(week, num, string, f_day):
	# Finds the average doctets/Duration value for a week data
	d = []
	d1 = []
	d2 = []
	d3 = []
	d4 = []
	d5 = []
	dateArray = week

	for i in week:
		if (int(i[0])//pow(10,6)) == f_day:
			d1.append(list(float(x) for x in i))
		if (int(i[0])//pow(10,6)) == f_day+1:
			d2.append(list(float(x) for x in i))
		if (int(i[0])//pow(10,6)) == f_day+2:
			d3.append(list(float(x) for x in i))
		if (int(i[0])//pow(10,6)) == f_day+3:
			d4.append(list(float(x) for x in i))
		if (int(i[0])//pow(10,6)) == f_day+4:
			d5.append(list(float(x) for x in i))

	# print(d1)
	d1 = avg(d1, num)
	d2 = avg(d2, num)
	d3 = avg(d3, num)
	d4 = avg(d4, num)
	d5 = avg(d5, num)

	d = d1 + d2 + d3 + d4 + d5
	return d
	
def add_column_sql(d,colname):
	start = time.time()
	if week == "week1":
		week = 'w1'
	else:
		week = 'w2'
	col_name = tablename + week
	query_add_col = "ALTER TABLE dValues ADD {} double NOT NULL DEFAULT '0.00';".format(col_name) 
	query_flag = csv_to_mysql(query_add_col)
	if query_flag == None:
		print(col_name,"exists")
		return
	print(col_name,"start")
	if d != None:
		d_new = []
		for i,elem in enumerate(d,0):
			if elem != 0 :
				d_new.append([i,elem])
		for i in d_new:
			query_load_col = "update dValues set {} = {} where id = {}".format(col_name,i[1],i[0])
			csv_to_mysql(query_load_col)
	

def spearmans_rank_correlation(xs, ys):
    # Calculate the rank of x's
    xranks = pd.Series(xs).rank()

    # Caclulate the ranking of the y's
    yranks = pd.Series(ys).rank()

    # Calculate Pearson's correlation coefficient on the ranked versions of the data
    return scipy.stats.pearsonr(xranks, yranks)


def points(a):
    # If the rank is equal to 1, changes it to 0.99
    if a == 1:
        return 0.99
    elif math.isnan(a):
        return 0
    else:
        return a


def findZ(r1a2a, r1a2b, r2a2b, timeSlot):
    # Finding the Z value using the ranks
    N = 0
    rmsq2 = 0
    f = 0
    h = 0
    Z2 = 0
    Z1 = 0
    p = 0
    q = 0
    N = (61200 - 28800) * 5 / timeSlot
    rmsq2 = (r1a2a * r1a2a + r1a2b * r1a2b) / 2
    f = (1 - r2a2b) / 2 * (1 - rmsq2)
    h = (1 - (f * rmsq2)) / (1 - rmsq2)
    Z2 = 0.5 * math.log((1 + r1a2b) / (1 - r1a2b))
    Z1 = 0.5 * math.log((1 + r1a2a) / (1 - r1a2a))
    p = Z1 - Z2
    q = math.sqrt((N - 3)) / 2 * (1 - r2a2b) * h

    value = p * q
    return value


def findP(z):
    # finds the final P value using the Z value
    p = 0.3275911
    a1 = 0.254829592
    a2 = -0.284496736
    a3 = 1.421413741
    a4 = -1.453152027
    a5 = 1.061405429
    sign = 0
    if (z < 0.0):
        sign = -1
    else:
        sign = 1

    x = abs(z) / math.sqrt(2.0)
    t = (1.0) / (1.0 + p * x)
    erf = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * math.exp(-x * x)
    phiZ = 0.5 * (1.0 + sign * erf)
    return 1 - phiZ

def tup_to_list(tup):
	arr = []
	for i in tup:
		arr.append(i[0])
	return arr

def extract(filename1, filename2):
    # Function to extract the week's data and saving the values into a list

    weeks = ['w1','w2']
    col_f1w1,col_f1w2,col_f2w1,col_f2w2 = filename1+weeks[0],filename1+weeks[1],filename2+weeks[0],filename2+weeks[1]
    query_extract = "select {} from dValues LIMIT 45000;"
    aw1 = query_extract.format(col_f1w1)
    aw2 = query_extract.format(col_f1w2)
    bw1 = query_extract.format(col_f2w1)
    bw2 = query_extract.format(col_f2w2)
    aw1 = csv_to_mysql(aw1)
    aw2 = csv_to_mysql(aw2)
    bw1 = csv_to_mysql(bw1)
    bw2 = csv_to_mysql(bw2)
    return tup_to_list(aw1), tup_to_list(aw2), tup_to_list(bw1), tup_to_list(bw2) 


# Excel to CSV-----------------------------

directory = '/mnt/c/Users/shrav/Desktop/Desk/InfoSec/xl/'
directory2 = '/mnt/c/Users/shrav/Desktop/Desk/InfoSec/csv/'


for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):
    	fname = directory+str(filename)
    	tablename = directory2+str((filename.split("."))[0])+".csv"
    	csv_from_excel(fname,tablename)



# Preprocessing------------------------------

directory = '/mnt/c/Users/shrav/Desktop/Desk/InfoSec/csv/'


for filename in os.listdir(directory):
	if filename.endswith(".csv"):
		fname = directory+str(filename)
		tablename = (filename.split("."))[0]
		query_table = "CREATE TABLE {} (id INT NOT NULL AUTO_INCREMENT, rfl double NOT NULL, doctets INT NOT NULL,duration INT NOT NULL, doc_by_dur float,PRIMARY KEY (id));".format(tablename)
		csv_to_mysql(query_table)
		query_load = "LOAD DATA LOCAL INFILE '{}' INTO TABLE {} FIELDS TERMINATED BY ',' ENCLOSED BY '\"' LINES TERMINATED BY '\n' IGNORE 1 ROWS (@col0,@col1,@col2,@col3,@col4,@col5,@col6,@col7,@col8,@col9,@col10) set rfl=@col6,doctets=@col4,duration=@col10 ;".format(fname,tablename)
		csv_to_mysql(query_load)
		query_delete = "Delete from {} where rfl NOT BETWEEN 1359982800000 AND 1360965599599;".format(tablename)
		csv_to_mysql(query_delete)
		query_epoch = "update {}  set rfl = from_unixtime(floor(rfl/1000));".format(tablename)
		csv_to_mysql(query_epoch)
		query_delete_duration = "delete from {} where duration = 0;".format(tablename)
		csv_to_mysql(query_delete_duration)
		query_dbyd = "update {} set doc_by_dur = doctets/duration".format(tablename)
		csv_to_mysql(query_dbyd)
		query_8to5 = "delete from {} where right(rfl,6) not between 080000 and 165959;".format(tablename)
		csv_to_mysql(query_8to5)
		query_mon2fri = "delete from {} where weekday(rfl) not between 0 and 4;".format(tablename)
		csv_to_mysql(query_mon2fri)
		print(filename,"done")


weeks = ["week1","week2"]
all_names = []
for filename in os.listdir(directory):
	if filename.endswith(".csv"):
		tablename = (filename.split("."))[0]
		all_names.append(tablename)

create_dValues()

for filename in os.listdir(directory):
	if filename.endswith(".csv"):
		fname = directory+str(filename)
		tablename = (filename.split("."))[0]
		for week in weeks:
			if week == "week1":
				query_week = "select right(rfl,8),doc_by_dur from {} where day(rfl) < (7-(select weekday(rfl) from ajb9b3 LIMIT 1) + (select day(rfl) from ajb9b3 LIMIT 1));".format(tablename)
				week_temp = 'w1'
			else:
				query_week = "select right(rfl,8),doc_by_dur from {} where day(rfl) > (6-(select weekday(rfl) from ajb9b3 LIMIT 1) + (select day(rfl) from ajb9b3 LIMIT 1));".format(tablename)
				week_temp = 'w2'

			####
			test_query = "select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='dValues';"
			test_results = csv_to_mysql(test_query)
			already_in_db = {}
			for i in test_results:
				already_in_db[i[0]] = 1
			####
			col_name = tablename + week_temp
			print(col_name,"start")
			start = time.time()
			if col_name in already_in_db:
				print(col_name,"exists")
				query_null2zero = "UPDATE dValues SET {fld_name} = 0 WHERE {fld_name} IS NULL;".format(fld_name=col_name)
				csv_to_mysql(query_null2zero)
				end = time.time()
				print((end - start)/60," mins")
				continue;
			week_data = csv_to_mysql(query_week)
			if len(week_data) > 0 :
				f_day = int((week_data[0])[0])//pow(10,6)
				d = findD(week_data,10,week,f_day)
			else:
				d = None	
			query_add_col = "ALTER TABLE dValues ADD {} double NULL;".format(col_name) #double NOT NULL DEFAULT '0.00'
			query_flag = csv_to_mysql(query_add_col)
			if query_flag == None:
				print(col_name,"exists")
			else:
				add_column_sql(d,col_name)

			query_null2zero = "UPDATE dValues SET {fld_name} = 0 WHERE {fld_name} IS NULL;".format(fld_name=col_name)
			csv_to_mysql(query_null2zero)

			end = time.time()
			print((end - start)/60," mins")



# print(all_names)
all_combs = [(x, y) for x in all_names for y in all_names]
p_dic = {}
for x in all_combs:
	p_dic[x] = 0
# print(all_combs)

for x in all_combs:
	# print(x)
	aw1, aw2, bw1, bw2 = extract(x[0],x[1])
	r1a2a = (spearmans_rank_correlation(aw1, aw2)[0])

	r1a2b = (spearmans_rank_correlation(aw1, bw2)[0])

	r2a2b = (spearmans_rank_correlation(aw2, bw2)[0])

	r1a2a = points(r1a2a)

	r1a2b = points(r1a2b)

	r2a2b = points(r2a2b)


	z = findZ(r1a2a, r1a2b, r2a2b, timeSlot = 10)

	p = findP(z)

	p_dic[x] = p 


print(p_dic)

