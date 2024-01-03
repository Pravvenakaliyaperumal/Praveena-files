import xlrd

location="C:\\Users\\CoxPHIT\\OneDrive\\Documents\\try_excel_xlss.xls"
var=xlrd.open_workbook(location)
sht= var.sheet_by_index(0)
print(sht.cell_value(1,0))

for i in range(0,5):
    print(sht.cell_value(1,i))
row_count=1

youth_last_name=sht.cell_value(1,0)
youth_first_name=sht.cell_value(1,1)
Authorization_number=sht.cell_value(1,2)
DOS=sht.cell_value(1,3)
Total_units=sht.cell_value(1,4)
Type_of_service=sht.cell_value(20,16)
madicaid_num=sht.cell_value(1,8)
dgx=str(sht.cell_value(1,9))
dgx=dgx.replace(".","")
DOS_start_date=sht.cell_value(1,11)
number="12"
bill_amount=sht.cell_value(27,17)

print((Type_of_service))

"""sht1= var.sheet_by_index(1)
print(sht1.cell_value(1,1))
unit_rate=int(sht1.cell_value(1,1))
total_BLL=int(unit_rate) * int(Total_units)
print(bill_amount)
bill_amount=int(bill_amount)
if total_BLL== bill_amount:
    print("its equal")
else:
    print("value not matching")
print(total_BLL)"""
DOB=sht.cell_value(29,8)

xl_date_DOB=int(DOB)
unit_rate=1
datetime_date_DOB = xlrd.xldate_as_datetime(xl_date_DOB, 0)
date_object_DOB = datetime_date_DOB.date()
print(date_object_DOB.strftime("%m/%d/%Y"))
DOB_value=date_object_DOB.strftime("%m/%d/%Y")
print(str(DOB_value))
print(type(DOB_value))
xl_date_DOS_start_date=int(DOS_start_date)
datetime_date_DOS_start_date = xlrd.xldate_as_datetime(xl_date_DOS_start_date, 0)
date_object_DOS_start_date = datetime_date_DOS_start_date.date()
print(date_object_DOS_start_date.strftime("%m/%d/%y"))
DOS_start_date_value=date_object_DOS_start_date.strftime("%m/%d/%y")
print(str(DOS_start_date_value))
total_BLL=float(unit_rate) * float(Total_units)
print(bill_amount)
bill_amount=float(bill_amount)
if total_BLL== bill_amount:
    print("its equal")
else:
    print("value not matching")
print(total_BLL)
print(bill_amount)
print("{:.2f}".format(bill_amount))
print("{:.2f}".format(total_BLL))

