import pandas as pd
import numpy as np
import math
import time
import datetime
from datetime import date
from pandas import ExcelWriter
from pandas import ExcelFile

REGISTRATION_FEE = 40
BOOKS_AND_SUPLIES = 40
TUITION_FEE = 465
SCHOOl_STARTING_DATE = date(2018, 8, 26)

def WriteExcelRecords(FileName, df):
    writer = ExcelWriter(FileName)
    df.to_excel(writer,'Sheet1')
    writer.save()

def IsPaid(payment_due):
    if payment_due < date.today():
        return "yes"
    else:
        return "No"
    
def AddPaymentToStudentRecord(paymentNo ,PaymentData, studentrecord):
    str_c = str(paymentNo)
    datex = "Payment" + str_c+ " date"
    Amount = "Amount" + str_c
    No = "Check" + str_c + " No"
    Paid = "Is Paid" + str_c
    mode = "Payment mode" + str_c

    if PaymentData["PaymentDate"] == None :
        studentrecord[datex] = None
        studentrecord[Paid] = None
    else:
        studentrecord[datex] = PaymentData["PaymentDate"].date()
        studentrecord[Paid] = IsPaid(PaymentData["PaymentDate"].date())

    studentrecord[Amount] = PaymentData["Amount"]
    studentrecord[No] = PaymentData["CheckNumber"]
    studentrecord[mode] = PaymentData["PaymentForm"]
    return studentrecord

        
def complet12(c,student_record):
    c = c+1
    PaymentData = {"PaymentDate": None, "Amount": None, "CheckNumber": None, "PaymentForm": None}
    for i in range(c,13): 
        student_record = AddPaymentToStudentRecord(i, PaymentData, student_record)
    return student_record

def Student_payments(PaymentsExcelFile):
    df = pd.read_excel(PaymentsExcelFile).sort_values(by=["StudentID","PaymentDate"])
    gp = df.groupby("StudentID")
    StudentPayments_df =pd.DataFrame()
    for student, payments in gp:
        student_record = {}
        student_record["StudentID"] = student
        c  = 0
        for row, data in payments.iterrows():
            c = c + 1
            student_record = AddPaymentToStudentRecord(c,df.loc[row],student_record)
        student_record = complet12(c,student_record)
        StudentPayments_df = StudentPayments_df.append(student_record, ignore_index = True)
    return StudentPayments_df

def joinquery(File1,File2, JoinColumn,NewIndex):
    df1 = pd.read_excel(File1)
    df2 = pd.read_excel(File2)
    df3 =  df1.join(df2.set_index(JoinColumn), on = JoinColumn)
    return df3

def ConvertStringToDate(x):
    'return   datetime.datetime.strptime( x, "%Y-%m-%d %H:%M:%S").date()'
    return x.date()
    
def TemplateDataFrame(sourceDf):
    destinationDf = pd.DataFrame()
    destinationDf["Primary contact (PC) First name"] = sourceDf["Father's First Name"]
    destinationDf["PC Middle Name"] = ""
    destinationDf["PC Lastname"] = sourceDf["Father's Last Name"]
    destinationDf["PC Email"] = sourceDf["Father's Email"]
    destinationDf["PC Mobile Number"] = sourceDf["Father's Cell Phone"]
    destinationDf["PC DOB"] = ""
    destinationDf["PC Type"] = ""
    destinationDf["PC Home Phone Number"] = sourceDf["HomePhoneNumber"]
    destinationDf["PC work Number"] = ""
    destinationDf["PC Mailing Street"] = sourceDf["Address"]
    destinationDf["PC Mailing City"] = sourceDf["City"]
    destinationDf["PC Mailing State"] = sourceDf["State"]
    destinationDf["PC Mailing Postal Code"] = sourceDf["ZipCode"]
    destinationDf["PC is Teacher?"] = "No"
    destinationDf["Is ADAMS Member"] = sourceDf["ADAMSID"]
    destinationDf["Secondary contact First name"] = sourceDf["Mother's First Name"]
    destinationDf["SC Middle Name"] = ""
    destinationDf["SC Lastname"] = sourceDf["Mother's Last Name"]
    destinationDf["SC Email"] = sourceDf["Mother's Email"]
    destinationDf["SC Mobile Number"] = sourceDf["Mother's Cell Phone"]
    destinationDf["SC DOB"] = ""
    destinationDf["SC Type"] =""
    destinationDf["SC Home Phone Number"] = ""
    destinationDf["SC work Number"] = ""
    destinationDf["Student First Name"] = sourceDf["FirstName"]
    destinationDf["Student Last Name"] = sourceDf["LastName"]
    destinationDf["Student Birthdate"] = sourceDf["DateOfBirth"].apply(ConvertStringToDate)
    destinationDf["Student Gender"] = sourceDf["Gender"]
    destinationDf["Student Allergies"] = sourceDf["Allergies"]
    destinationDf["Student Emergency Contact Name"] = sourceDf["EmergencyContact"]
    destinationDf["Student Emergency Contact Number"] = sourceDf["EmergencyPhone"]
    destinationDf["Student Enrollment Date"] = ""
    destinationDf["Student Enrollment Status"] = sourceDf["StudentStatus"]
    destinationDf["Student Medication"] = ""
    destinationDf["Student Doctor"] = ""
    destinationDf["Student Med Insur Comp"] = ""
    destinationDf["Student Policy"] = ""
    destinationDf["Student Emergency Phone"] = ""
    destinationDf["Class Name"] = sourceDf["Grade"]
    destinationDf["Class session"] = sourceDf["Period"]
    destinationDf["Student Registration Date"] = sourceDf["EnrollDate"]
    destinationDf["Student Registration Fee"] = REGISTRATION_FEE
    destinationDf["Student Books/supply fee"] = BOOKS_AND_SUPLIES
    destinationDf["Student Tuition Fee"] = TUITION_FEE
    destinationDf["PaymentTermStatus"] = sourceDf["PaymentTermStatus"]
    destinationDf["Discount type"] = ""
    destinationDf["Payment mode1"] = sourceDf["Payment mode1"]
    destinationDf["Payment1 date"] = sourceDf["Payment1 date"]
    destinationDf["Amount1"] = sourceDf["Amount1"]
    destinationDf["Cheque1 No"] = sourceDf["Check1 No"]
    destinationDf["Is Paid1"] = sourceDf["Is Paid1"]
    destinationDf["Payment mode2"] = sourceDf["Payment mode2"]
    destinationDf["Payment2 date"] = sourceDf["Payment2 date"]
    destinationDf["Amount2"] = sourceDf["Amount2"]
    destinationDf["Cheque2 No"] = sourceDf["Check2 No"]
    destinationDf["Is Paid2"] = sourceDf["Is Paid2"]
    destinationDf["Payment mode3"] = sourceDf["Payment mode3"]
    destinationDf["Payment3 date"] = sourceDf["Payment3 date"]
    destinationDf["Amount3"] = sourceDf["Amount3"]
    destinationDf["Cheque3 No"] = sourceDf["Check3 No"]
    destinationDf["Is Paid3"] = sourceDf["Is Paid3"]
    destinationDf["Payment mode4"] = sourceDf["Payment mode4"]
    destinationDf["Payment4 date"] = sourceDf["Payment4 date"]
    destinationDf["Amount4"] = sourceDf["Amount4"]
    destinationDf["Cheque4 No"] = sourceDf["Check4 No"]
    destinationDf["Is Paid4"] = sourceDf["Is Paid4"]
    destinationDf["Payment mode5"] = sourceDf["Payment mode5"]
    destinationDf["Payment5 date"] = sourceDf["Payment5 date"]
    destinationDf["Amount5"] = sourceDf["Amount5"]
    destinationDf["Cheque5 No"] = sourceDf["Check5 No"]
    destinationDf["Is Paid5"] = sourceDf["Is Paid5"]
    destinationDf["Payment mode6"] = sourceDf["Payment mode6"]
    destinationDf["Payment6 date"] = sourceDf["Payment6 date"]
    destinationDf["Amount6"] = sourceDf["Amount6"]
    destinationDf["Cheque6 No"] = sourceDf["Check6 No"]
    destinationDf["Is Paid6"] = sourceDf["Is Paid6"]
    destinationDf["Payment mode7"] = sourceDf["Payment mode7"]
    destinationDf["Payment7 date"] = sourceDf["Payment7 date"]
    destinationDf["Amount7"] = sourceDf["Amount7"]
    destinationDf["Cheque7 No"] = sourceDf["Check7 No"]
    destinationDf["Is Paid7"] = sourceDf["Is Paid7"]
    destinationDf["Payment8 date"] = sourceDf["Payment8 date"]
    destinationDf["Amount8"] = sourceDf["Amount8"]
    destinationDf["Cheque8 No"] = sourceDf["Check8 No"]
    destinationDf["Is Paid8"] = sourceDf["Is Paid8"]
    destinationDf["Payment9 date"] = sourceDf["Payment9 date"]
    destinationDf["Amount9"] = sourceDf["Amount9"]
    destinationDf["Cheque9 No"] = sourceDf["Check9 No"]
    destinationDf["Is Paid9"] = sourceDf["Is Paid9"]
    destinationDf["Payment10 date"] = sourceDf["Payment10 date"]
    destinationDf["Amount10"] = sourceDf["Amount10"]
    destinationDf["Cheque10 No"] = sourceDf["Check10 No"]
    destinationDf["Is Paid10"] = sourceDf["Is Paid10"]
    destinationDf["Payment11 date"] = sourceDf["Payment11 date"]
    destinationDf["Amount11"] = sourceDf["Amount11"]
    destinationDf["Cheque11 No"] = sourceDf["Check11 No"]
    destinationDf["Is Paid11"] = sourceDf["Is Paid11"]
    destinationDf["Payment12 date"] = sourceDf["Payment12 date"]
    destinationDf["Amount12"] = sourceDf["Amount12"]
    destinationDf["Cheque12 No"] = sourceDf["Check12 No"]
    destinationDf["Is Paid12"] = sourceDf["Is Paid12"]
    'WriteExcelRecords("full_sheet_temp.xlsx",destinationDf)'
    return fillPC(destinationDf)

def PC_Type(x):
    if pd.isnull(x):
        return "Mother"
    else:
        return "Father"
 
def fillPC(df):
    
    df["PC Type"]  = df["Primary contact (PC) First name"].apply(PC_Type)
  
    values ={"Primary contact (PC) First name" : df["Secondary contact First name"],
             "PC Lastname": df["SC Lastname"] ,
             "PC Email": df["SC Email"] ,
             "PC Mobile Number": df["SC Mobile Number"]}
    
    df.fillna(value = values, inplace=True)
    
    return df
    
def main():
    'creat student payments records data frame'
    StudentPayments_df = Student_payments("PaymentSubQuery.xlsx")
    'WriteExcelRecords("StudentPayments_df.xlsx",StudentPayments_df)'
    'join stdent payments with students family and student info'
    
    StudenFamily_df = pd.read_excel("Students.xlsx").join(pd.read_excel("Family.xlsx").set_index("FamilyID"), on = "FamilyID")
    'WriteExcelRecords("StudenFamily_df.xlsx",StudenFamily_df)'
    full_df = StudenFamily_df.join(StudentPayments_df.set_index("StudentID"), on="StudentID")
    Temp_df = TemplateDataFrame(full_df).sort_values(by=["PC Email"])
    WriteExcelRecords("template_sheet.xlsx",Temp_df )
    'WriteExcelRecords("full_sheet.xlsx",full_df)'
    
main()
 

