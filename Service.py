import xlrd

# session="IIC-L"
global session_code, unit_rate,session_valid
# service_session=try_sesion.service_session
service_session=""

location="C:\\Users\\CoxPHIT\\OneDrive\\Documents\\try_excel_xlss.xls"
var=xlrd.open_workbook(location)
sht= var.sheet_by_index(0)
rowcount=20
Type_of_service=sht.cell_value(rowcount,16)
session_type=Type_of_service
# print(session,"session")

def session(session):
    # session = "IIC-L"
    print(service_session)
    session_code="H360.TJ.U1"
    unit_rate=int(7.06)
    if session=="IIC-L":
        session_code = "H0036.TJ.U1"
        unit_rate = float(31.60)
        print(session_code)
        print("here")
        return session_code, unit_rate
    elif session=="IIC-M":
        session_code = "H0036.TJ.U2"
        unit_rate = float(29.46)
        print(session_code)
        print("here")
        return session_code, unit_rate
    elif session == "BA":
        session_code = "H2014.TJ"
        unit_rate = float(18.64)
        print(session_code)
        print("here")
        return session_code, unit_rate
    elif session == "SHR":
        session_code = "T1005.22.HA"
        unit_rate = float(7.06)
        print(session_code)
        print("here")
        return session_code, unit_rate


    else:
        print("check the Service")
def add(a,b):
    return a+b
session_valid=True

session(str(Type_of_service))