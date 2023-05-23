from Employee import AttendanceDetails, Employee
from EmployeeProfileDAL import EmployeeProfileDAL

from flask import Flask,jsonify,json,redirect,url_for
from flask import request , send_file , after_this_request
from flask import render_template
import os
import datetime as importDateTime
from datetime import timedelta,date, datetime
import calendar
from calendar import monthrange
from calendar import mdays
from shutil import copy
from flask import Response
from flask import session,g
from logging.handlers import RotatingFileHandler
import random
import xlsxwriter


app = Flask(__name__)
app.secret_key = os.urandom(24)

#Login and Logout Page
@app.route('/')
def home():
   return render_template('loginV4.html')

@app.route('/signout', methods=['GET'])
def signout():
    if 'user' not in session:
        return redirect(url_for('home'))
    session.pop('user', None)
    return redirect(url_for('home'))


#Employee Details API's - Create, Update and Delete Employee

@app.route('/profile', methods=['GET', 'POST'])
def login():
    try:
        if request.method == 'GET':
            sb = EmployeeProfileDAL()
            rowReturn = sb.read_employee()
            projectList=get_project_list()
            return render_template("Dashboard.html", rowTable=rowReturn, projectList=projectList)
        if request.method == 'POST':
            corpid=request.form['corpId']
            corppass = request.form['corppass']
            sb = EmployeeProfileDAL()
            EmployeeName = (sb.get_current_employee_Info(corpid))
            rowReturn = sb.read_employee()
            loginfailedmsg = "Invalid credentials"
            projectList=get_project_list()
            if corpid:
                if request.method == 'POST':
                    session.pop('user', None)
                    if EmployeeName:
                        session['user'] = request.form['corpId']
                        app.logger.info('-------------------------------------------------------------------------------------')
                        app.logger.info('Logged in by: %s', corpid)
                        admin_return= Admin()
                        if admin_return=="Yes":
                            return redirect(url_for("viewTeamfun"))
                        else:
                            return render_template("Dashboard.html", rowTable=rowReturn, projectList=projectList)
            app.logger.error('Failed to login for %s',corpid)
            return render_template("loginV4.html", **locals())
    except Exception as e:
        return str(e)

@app.route('/viewteam')
def viewTeamfun():
    if 'user' in session:
        corp_id=session['user']
        # corp_id="conddas"
        sb=EmployeeProfileDAL()
        EmployeeName=sb.get_current_employee_Info(corp_id)[0][0]
        AdminReturn = Admin()
        if AdminReturn == "Yes":
            return render_template("viewteam.html", **locals())
        else:
            return render_template("LeaveAppPart2.html", EmployeeName=EmployeeName,
                                   corpid=corp_id)
    return render_template('loginV4.html', **locals())


@app.route('/employee details')
def list_all_users():
    if 'user' in session:
        sb = EmployeeProfileDAL()
        corpid=session['user']
        EmployeeName = corpid
        row_return = sb.read_employee()
        projectList = get_project_list()
        employeeLevelList = get_employeeLevel_list()
        app.logger.info('Employee Details page viewed by : %s', corpid)
        AdminReturn = Admin()
        if AdminReturn == "Yes":
            return render_template("Dashboard.html", rowTable=row_return, **locals())
        else:
            return render_template("Dashboard.html", rowTable=row_return, EmployeeName=EmployeeName,corpid=corpid, projectList=projectList, employeeLevelList = employeeLevelList)
    return redirect(url_for('home'))


@app.route('/add profile', methods=['POST'])
def add_profile():
    if 'user' in session:
        employee_id = request.form['employeeId']
        employee_name = request.form['employeeName']
        project_name = request.form['ProjectName']
        corpid = session['user']
        email = request.form['Mail']
        corp_idM = request.form['CorpID']
        department = request.form['Department']
        employeeODCStatus="Assigned"
        expertise=request.form['Expertise']
        employeeLevel=request.form['EmployeeLevel']
        sb = EmployeeProfileDAL()
        project_id = sb.get_project_id(project_name=project_name)
        EmployeeName = corpid
        employee = Employee(employee_id, employee_name, project_id, project_name, corp_idM, email, department, employeeODCStatus,expertise, employeeLevel)
        sb.add_employee(employee)
        rowReturn = sb.read_employee()
        sb.c.close()
        projectList = get_project_list()
        employeeLevelList = get_employeeLevel_list()
        app.logger.info('%s added by: %s',employee_id, corpid)
        AdminReturn = Admin()
        if AdminReturn == "Yes":
            return render_template("Dashboard.html", rowTable=rowReturn, **locals())
        else:
            return render_template("Dashboard.html", rowTable=rowReturn, EmployeeName=EmployeeName, employee=employee, corpid=corpid,projectList=projectList, employeeLevelList = employeeLevelList)
    return redirect(url_for('home'))

@app.route('/compare', methods=['POST'])
def compare():
    formElement = request.json
    sb = EmployeeProfileDAL()
    for keyFromDict in formElement:
        key = keyFromDict
    idFromDB = sb.gettingEmployeeDetailsForRepeatedEntries(formElement)
    if idFromDB == 1:
        msg = key + " is already exist in the system.Please try another."
        return jsonify({'error': msg})
    else:
        return jsonify({'success': 'true'})


@app.route('/Update profile/0', methods=['POST'])
def update_profile():
    if 'user' in session:
        sb = EmployeeProfileDAL()
        corpid = session['user']
        EmployeeName = corpid
        employee_id = request.form['employeeId']
        employee_name = request.form['employeeName']
        project_name = request.form['projectNameUpdate']
        corpIdM = request.form['corpIdUpdate']
        email = request.form['emailIdUpdate']
        employeeODCStatus= 'Assigned'
        department = request.form['DepartmentUpdate']
        expertise = request.form['expertiseUpdateName']
        employeeLevelUpdate = request.form['employeeLevelUpdate']
        project_id = sb.get_project_id(project_name=project_name)
        employee = Employee(employee_id, employee_name, project_id, project_name, corpIdM, email, department, employeeODCStatus,expertise, employeeLevelUpdate)
        sb.update_employee(employee)
        rowReturn = sb.read_employee()
        sb.c.close()
        projectList = get_project_list()
        employeeLevelList = get_employeeLevel_list()
        app.logger.info('%s updated profile details', corpid)
        AdminReturn = Admin()
        if AdminReturn == "Yes":
            return render_template("Dashboard.html", rowTable=rowReturn, **locals())
        else:
            return render_template("Dashboard.html", rowTable=rowReturn, EmployeeName=EmployeeName, employee=employee,projectList=projectList, employeeLevelList = employeeLevelList)
    return redirect(url_for('home'))



@app.route('/deleteEmployee',methods=['GET'])
def deleteemp():
    if 'user' in session:
        employeeId = request.args['employeeId']
        sb = EmployeeProfileDAL()
        delete_status = sb.delete_employee(employeeId)
        return delete_status


#Apply Leave API's
@app.route('/personalLeave')
def personalLeave():
    if 'user' in session:
        corpid = session['user']
        sb = EmployeeProfileDAL()
        EmployeeName=(sb.get_current_employee_Info(corpid))[0][0]
        EmployeeName = corpid
        AdminReturn = Admin()
        if AdminReturn == "Yes":
          return render_template('personalCal.html', **locals())
        else:
            return render_template('personalCal.html', EmployeeName=EmployeeName,corpid=corpid)
    return render_template('loginV4.html', **locals())

@app.route('/dj',methods=["GET"])
def jsondata():
    with open("static/json/pi.json",'r', encoding='utf-8-sig') as json_file:
        json_data = json.load(json_file)
        sb=EmployeeProfileDAL()
    return jsonify(json_data)

@app.route('/getCurrentUser', methods=["GET"])
def getCurrentUser():
    if 'user' in session:
        corpid=session['user']
        return jsonify(corpid)
    return jsonify("false")

@app.route('/showPersonalLeave',methods=["POST","GET"])
def showPersonalLeave():
    sb = EmployeeProfileDAL()
    corp_id_org=request.args.get('corpid')
    if corp_id_org is not None:
        rowsForManagerEmployee = sb.readTotalLeavesForAnEmployee(corp_id_org)
        return jsonify(rowsForManagerEmployee)
    return render_template('loginV4.html', **locals())

@app.route('/applyLeave' ,methods=["POST", "GET"])
def applyLeave():
    if 'user' in session:
        date = request.form['Date']
        leaveType=request.form['LeaveType']
        corpid=request.form['CorpID']
        sb = EmployeeProfileDAL()
        sb.submit_leaves(date, corpid,leaveType)
        EmployeeName = (sb.get_current_employee_Info(corpid))[0][0]
        app.logger.info('Leave applied for %s on %s by: %s', corpid, date, corpid)
        return jsonify(success='true')
    return render_template('loginV4.html', **locals())


#Adding All Lab Request API's

@app.route('/labRequest')
def labRequest():
    if 'user' in session:
        corpid = session['user']
        sb = EmployeeProfileDAL()
        EmployeeName=(sb.get_current_employee_Info(corpid))[0][0]
        EmployeeName = corpid
        projectList = get_project_list()
        sb = EmployeeProfileDAL()
        rowTable = sb.read_lab_requests()
        AdminReturn = Admin()
        if AdminReturn == "Yes":
          return render_template('Lab.html', **locals())
        else:
            return render_template('Lab.html', EmployeeName=EmployeeName,corpid=corpid, projectList=projectList, rowTable=rowTable)
    return render_template('loginV4.html', **locals())


@app.route('/add lab request', methods=['POST'])
def add_lab_request():
    if 'user' in session:
        request_description = request.form['description']
        project_name = request.form['ProjectName']
        corpid = session['user']
        sb = EmployeeProfileDAL()
        today = date.today().strftime('%m/%d/%Y')
        EmployeeName = (sb.get_current_employee_Info(corpid))[0][0]
        projectList = get_project_list()
        id = int(random.random() * 100000.0)
        sb.add_lab_request(request_description,EmployeeName,project_name, today, id)
        rowReturn = sb.read_lab_requests()
        rowTable = sb.read_lab_requests()
        AdminReturn = Admin()
        if AdminReturn == "Yes":
          return render_template('Lab.html', **locals())
        else:
            return render_template('Lab.html', EmployeeName=EmployeeName,corpid=corpid, projectList=projectList, rowTable=rowTable)
    return render_template('loginV4.html', **locals())


@app.route('/Delete Request', methods=['POST'])
def delete_lab_request():
    corpid = session['user']
    sb = EmployeeProfileDAL()
    EmployeeName = (sb.get_current_employee_Info(corpid))[0][0]
    request_id = request.form['requestId']
    sb.delete_lab_request(request_id)
    rowTable = sb.read_lab_requests()
    projectList = get_project_list()
    AdminReturn = Admin()
    if AdminReturn == "Yes":
        return render_template('Lab.html', **locals())
    else:
        return render_template('Lab.html', EmployeeName=EmployeeName,corpid=corpid, projectList=projectList, rowTable=rowTable)


#Adding Team Builder API's

@app.route('/org')
def CreateOrg():
    if 'user' in session:
        corpid = session['user']
        sb=EmployeeProfileDAL()
        EmployeeName = (sb.get_current_employee_Info(corpid))[0][0]
        return render_template("CreateOrg.html", **locals())
    return render_template('loginV4.html', **locals())


@app.route('/orgDetails',methods=['GET', 'POST'])
def showdataforManagers():
    if 'user' in session:
        corp_id=session['user']
        sb = EmployeeProfileDAL()
        manager_id = (sb.get_current_employee_Info(corp_id))[0][2]
        if request.method == 'GET':
            EmployeeDetails = sb.read_employee_in_dict()
            return jsonify(EmployeeDetails)
        else:
            formElement = request.json
            sb = EmployeeProfileDAL()
            result=sb.AssiningToManager(manager_id, formElement)
            return jsonify({"result": result})
    return render_template('loginV4.html', **locals())


@app.route('/updateStatus',methods=["GET","POST"])
def updateEmployeeStatus():
    if 'user' in session:
        if request.method=="POST":
            formElement = request.json
            corp_id = session['user']
            sb = EmployeeProfileDAL()
            manager_id = (sb.get_current_employee_Info(corp_id))[0][2]
            status=sb.update_status(formElement)
            return jsonify(status)
    return render_template('loginV4.html', **locals())


@app.route('/getsetviewdata',methods=['GET', 'POST'])
def getsetDataforteam():
    if 'user' in session:
        corp_id = session['user']
        obj = EmployeeProfileDAL()
        EmployeeName=obj.get_current_employee_Info(corp_id)[0][0]
        manager_id = (obj.get_current_employee_Info(corp_id))[0][2]
        if request.method == 'GET':
            orglist = obj.gettingAssignedEmployeeToManager(manager_id=manager_id)
            return jsonify(orglist)
    return render_template('loginV4.html', **locals())



#Month Details API
@app.route('/currentMonth')
def currentMonthDetails():
    if 'user' in session:
        corpid = session['user']
        v = request.args.get('mon')
        if v is not None:
            v = v.split("-")
            sb = EmployeeProfileDAL()
            EmployeeName = (sb.get_current_employee_Info(corpid))[0][0]
            d = importDateTime.date.today()
            month = v[1]
            year = v[0]
            dateArray = []
            dateArray = dateArrayMethod(int(year), int(month))
            employeeStatusListView = []
            employeeStatusListView = gettingInfo(month, int(year))
            AdminReturn = Admin()
            if AdminReturn == "Yes":
                return render_template("LeaveAppPart2.html", **locals())
            else:
                return render_template("LeaveAppPart2.html",dateArray=dateArray,employeeStatusListView=employeeStatusListView, EmployeeName=EmployeeName, corpid=corpid)
        else:
            sb = EmployeeProfileDAL()
            EmployeeName = (sb.get_current_employee_Info(corpid))[0][0]
            d = importDateTime.date.today()
            month = d.strftime('%m')
            year = d.strftime('%Y')
            dateArray = []
            dateArray = dateArrayMethod(int(year), int(month))
            employeeStatusListView = []
            employeeStatusListView = gettingInfo(month, int(year))
            AdminReturn = Admin()
            if AdminReturn == "Yes":
                for employeelist in employeeStatusListView:
                    for value in employeelist:
                        if value is None:
                            print(employeelist[2])

                return render_template("LeaveAppPart2.html", **locals())
            else:
                return render_template("LeaveAppPart2.html", dateArray=dateArray,
                                       employeeStatusListView=employeeStatusListView, EmployeeName=EmployeeName,
                                       corpid=corpid)
    return render_template('loginV4.html', **locals())



#Monthly Other Deductions API

@app.route('/monthlyOtherDeductions')
def monthlyOtherDeductions():
    if 'user' in session:
        corpid = session['user']
        v = request.args.get('mon')
        if v is not None:
            v = v.split("-")
            sb = EmployeeProfileDAL()
            EmployeeName = (sb.get_current_employee_Info(corpid))[0][0]
            d = importDateTime.date.today()
            month = v[1]
            year = v[0]
            dateArray = []
            dateArray = dateArrayMethod(int(year), int(month))
            employeeStatusListView = []
            employeeStatusListView = gettingOtherDeductionsInfo(month, int(year))
            AdminReturn = Admin()
            if AdminReturn == "Yes":
                return render_template("OtherDeductions.html", **locals())
            else:
                return render_template("OtherDeductions.html", dateArray=dateArray,
                                       employeeStatusListView=employeeStatusListView, EmployeeName=EmployeeName,
                                       corpid=corpid)
        else:
            sb = EmployeeProfileDAL()
            EmployeeName = (sb.get_current_employee_Info(corpid))[0][0]
            d = importDateTime.date.today()
            month = d.strftime('%m')
            year = d.strftime('%Y')
            dateArray = []
            dateArray = dateArrayMethod(int(year), int(month))
            employeeStatusListView = []
            employeeStatusListView = gettingOtherDeductionsInfo(month, int(year))
            AdminReturn = Admin()
            if AdminReturn == "Yes":
                return render_template("OtherDeductions.html", **locals())
            else:
                return render_template("OtherDeductions.html", dateArray=dateArray,
                                       employeeStatusListView=employeeStatusListView, EmployeeName=EmployeeName,
                                       corpid=corpid)
    return render_template('loginV4.html', **locals())

#Attendance Portal API
@app.route('/attendance', methods=['GET'])
def attendance():
    if 'user' in session:
        AdminReturn = Admin()
        if AdminReturn == "Yes":
            return render_template("attendance.html", **locals())
        else:
            return render_template("attendance.html")
    else:
        return render_template('loginV4.html', **locals())


@app.route('/attendanceyesterday', methods=['GET'])
def attendanceyesterday():
    if 'user' in session:
        AdminReturn = Admin()
        if AdminReturn == "Yes":
            return render_template("attendanceyesterday.html", **locals())
        else:
            return render_template("attendanceyesterday.html")
    else:
        return render_template('loginV4.html', **locals())


@app.route('/attendanceemployees', methods=['GET'])
def attendanceemployees():
    sb = EmployeeProfileDAL()
    setAttendancetableDates()
    attendanceEmployees = sb.attendance_employees()
    employeeList = []
    for employee in attendanceEmployees:
        employeeDict = {
            "EmployeeId" : employee[0],
            "EmployeeName" : employee[1],
            "ProjectName" : employee[2],
            "AtOffice" : employee[3],
            "SickLeave" : employee[4],
            "CasualLeave" : employee[5],
            "WorkFromHome" : employee[6],
        }
        employeeList.append(employeeDict)
    return jsonify(employeeList)


@app.route('/attendanceemployeesyesterday', methods=['GET'])
def attendanceemployeesyesterday():   
    sb = EmployeeProfileDAL()
    attendanceEmployees = sb.attendance_employees_yesterday()
    employeeList = []
    for employee in attendanceEmployees:
        employeeDict = {
            "EmployeeId" : employee[0],
            "EmployeeName" : employee[1],
            "ProjectName" : employee[2],
            "AtOffice" : employee[3],
            "SickLeave" : employee[4],
            "CasualLeave" : employee[5],
            "WorkFromHome" : employee[6],
        }
        employeeList.append(employeeDict)
    return jsonify(employeeList)

@app.route('/projectnames', methods=['GET'])
def projectnames():
    if request.method =='GET':
        sb = EmployeeProfileDAL()
        return jsonify(sb.getProjects())

@app.route('/saveattendance/<attendanceDay>', methods=['POST'])
def saveattendance(attendanceDay):
    if request.method =='POST':
        sb = EmployeeProfileDAL()
        employeeId = request.form['employeeId']
        atOffice = request.form['atOffice']
        sickLeave = request.form['sickLeave']
        casualLeave = request.form['casualLeave']
        workFromHome = request.form['workFromHome']
        leaveId = getLeaveIdFromLeaveType(sickLeave=sickLeave, casualLeave=casualLeave)
        print(f'Leave id : {leaveId}')
        attendanceDetails = AttendanceDetails(employee_id= employeeId, at_office=atOffice,sick_leave=sickLeave,casual_leave=casualLeave,work_form_home=workFromHome)
        if attendanceDay == 'today':
            sb.update_employee_attendance(attendanceDetails=attendanceDetails)
            if leaveId:                
                sb.insert_into_leave_details_table(attendanceDetails=attendanceDetails, leavetype= leaveId)
        if attendanceDay == 'yesterday':
            sb.update_employee_attendance_yesterday(attendanceDetails=attendanceDetails)
            if leaveId:
                sb.insert_into_leave_details_table_yesterday(attendanceDetails=attendanceDetails, leavetype= leaveId)
        sb.c.close()
        return "Ok"


@app.route('/downloadattendancereport', methods=['GET'])
def downloadattendancereport():
    current_directory = os.getcwd()
    path = os.path.join(current_directory, r'Attendance')
    isExist = os.path.exists(path)
    try:
        if not isExist:
            os.makedirs(path)
    except Exception as e:
        return str(e)
    try:      
        for filename in os.listdir(path):
            file_path = os.path.join(path, filename)            
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.remove(file_path)                
    except Exception as e:
        return str(e)
    try:
        sb = EmployeeProfileDAL()
        #Get the count of all employees for Half Day and Full Day
        attendanceCount = sb.get_count_of_all_attendance_employees()
        attendanceCountHalfDay = sb.get_count_of_all_attendance_employees_halfDay()
        attendanceCountYesterday = sb.get_count_of_all_attendance_employees_yesterday()
        attendanceCountHalfDayYesterday = sb.get_count_of_all_attendance_employees_yesterday_halfday()

        atOfficeCountTotalToday = attendanceCount['AtOfficeCount'] + attendanceCountHalfDay['AtOfficeCountHalfday']
        sickLeaveTotalCountToday = attendanceCount['SickLeaveCount'] + attendanceCountHalfDay['SickLeaveCountHalfDay']
        casualLeaveTotalCountToday = attendanceCount['CasualLeaveCount'] + attendanceCountHalfDay['CasualLeaveCountHalfDay']
        wfhLeaveTotalCountToday = attendanceCount['WorkFromHomeCount'] + attendanceCountHalfDay['WorkFromHomeCountHalfDay']

        atOfficeCountTotalYesterday = attendanceCountYesterday['AtOfficeCount'] + attendanceCountHalfDayYesterday['AtOfficeCountHalfDay']
        wfhCountTotalYesterday = attendanceCountYesterday['WorkFromHomeCount'] + attendanceCountHalfDayYesterday['WorkFromHomeCountHalfDay']
        casualLeaveCountTotalYesterday = attendanceCountYesterday['CasualLeaveCount'] + attendanceCountHalfDayYesterday['CasualLeaveCountHalfDay']
        sickLeaveCountTotalYesterday = attendanceCountYesterday['SickLeaveCount'] + attendanceCountHalfDayYesterday['SickLeaveCountHalfDay']
        
        #Get Employee Names for all half Day and Full day
        attendanceEmployees = sb.get_all_attendance_employees()        
        attendanceEmployeesYesterday = sb.get_all_attendance_employees_yesterday()     
        attendanceEmployeesHalfDay = sb.get_all_attendance_employees_halfday()        
        attendanceEmployeesHalfDayYesterday = sb.get_all_attendance_employees_yesterday_halfday() 

        if attendanceEmployeesHalfDay["AtOfficeEmployeesHalfDay"] != 'None' and attendanceEmployees['AtOfficeEmployees'] != 'None':
            atofficeEmployeesToday = attendanceEmployeesHalfDay["AtOfficeEmployeesHalfDay"] + attendanceEmployees['AtOfficeEmployees']
        if attendanceEmployeesHalfDay["SickLeaveEmployeesHalfDay"] != 'None' and attendanceEmployees['SickLeaveEmployees'] != 'None':
            sickLeaveEmployeesToday = attendanceEmployeesHalfDay["SickLeaveEmployeesHalfDay"] + attendanceEmployees['SickLeaveEmployees']
        if attendanceEmployeesHalfDay["CasualLeaveEmployeesHalfDay"] != 'None' and attendanceEmployees['CasualLeaveEmployees'] != 'None':
            casualLeaveEmployeesToday = attendanceEmployeesHalfDay["CasualLeaveEmployeesHalfDay"] + attendanceEmployees['CasualLeaveEmployees']

        if attendanceEmployeesHalfDay["AtOfficeEmployeesHalfDay"] == 'None' and attendanceEmployees['AtOfficeEmployees'] != 'None':
            atofficeEmployeesToday = attendanceEmployees['AtOfficeEmployees']
        if attendanceEmployeesHalfDay["SickLeaveEmployeesHalfDay"] == 'None' and attendanceEmployees['SickLeaveEmployees'] != 'None':
            sickLeaveEmployeesToday = attendanceEmployees['SickLeaveEmployees']
        if attendanceEmployeesHalfDay["CasualLeaveEmployeesHalfDay"] == 'None' and attendanceEmployees['CasualLeaveEmployees'] != 'None':
            casualLeaveEmployeesToday = attendanceEmployees['CasualLeaveEmployees']

        if attendanceEmployeesHalfDay["AtOfficeEmployeesHalfDay"] != 'None' and attendanceEmployees['AtOfficeEmployees'] == 'None':
            atofficeEmployeesToday = attendanceEmployeesHalfDay["AtOfficeEmployeesHalfDay"]
        if attendanceEmployeesHalfDay["SickLeaveEmployeesHalfDay"] != 'None' and attendanceEmployees['SickLeaveEmployees'] == 'None':
            sickLeaveEmployeesToday = attendanceEmployeesHalfDay["SickLeaveEmployeesHalfDay"]
        if attendanceEmployeesHalfDay["CasualLeaveEmployeesHalfDay"] != 'None' and attendanceEmployees['CasualLeaveEmployees'] == 'None':
            casualLeaveEmployeesToday = attendanceEmployeesHalfDay["CasualLeaveEmployeesHalfDay"]

        print(f'attendanceEmployeesHalfDayYesterday :{attendanceEmployeesHalfDayYesterday}')

        if attendanceEmployeesHalfDayYesterday['AtOfficeEmployeeshalfday'] != 'None' and attendanceEmployeesYesterday['AtOfficeEmployees'] != 'None':
            atOfficeEmployeesYesterday = attendanceEmployeesHalfDayYesterday['AtOfficeEmployeeshalfday'] + attendanceEmployeesYesterday['AtOfficeEmployees']
        if attendanceEmployeesHalfDayYesterday['SickLeaveEmployeeshalfday'] != 'None'  and attendanceEmployeesYesterday['SickLeaveEmployees'] != 'None':
            sickLeaveEmployeesYesterday = attendanceEmployeesHalfDayYesterday['SickLeaveEmployeeshalfday'] + attendanceEmployeesYesterday['SickLeaveEmployees']
        if attendanceEmployeesHalfDayYesterday['CasualLeaveEmployeeshalfday'] != 'None' and attendanceEmployeesYesterday['CasualLeaveEmployees'] != 'None':
            casualLeaveEmployeesYesterday = attendanceEmployeesHalfDayYesterday['CasualLeaveEmployeeshalfday'] + attendanceEmployeesYesterday['CasualLeaveEmployees']

        if attendanceEmployeesHalfDayYesterday['AtOfficeEmployeeshalfday'] != 'None' and attendanceEmployeesYesterday['AtOfficeEmployees'] == 'None':
            atOfficeEmployeesYesterday = attendanceEmployeesHalfDayYesterday['AtOfficeEmployeeshalfday']
        if attendanceEmployeesHalfDayYesterday['SickLeaveEmployeeshalfday'] != 'None'  and attendanceEmployeesYesterday['SickLeaveEmployees'] == 'None':
            sickLeaveEmployeesYesterday = attendanceEmployeesHalfDayYesterday['SickLeaveEmployeeshalfday']
        if attendanceEmployeesHalfDayYesterday['CasualLeaveEmployeeshalfday'] != 'None' and attendanceEmployeesYesterday['CasualLeaveEmployees'] == 'None':
            casualLeaveEmployeesYesterday = attendanceEmployeesHalfDayYesterday['CasualLeaveEmployeeshalfday']

        if attendanceEmployeesHalfDayYesterday['AtOfficeEmployeeshalfday'] == 'None' and attendanceEmployeesYesterday['AtOfficeEmployees'] != 'None':
            atOfficeEmployeesYesterday = attendanceEmployeesYesterday['AtOfficeEmployees']
        if attendanceEmployeesHalfDayYesterday['SickLeaveEmployeeshalfday'] == 'None'  and attendanceEmployeesYesterday['SickLeaveEmployees'] != 'None':
            sickLeaveEmployeesYesterday = attendanceEmployeesYesterday['SickLeaveEmployees']
        if attendanceEmployeesHalfDayYesterday['CasualLeaveEmployeeshalfday'] == 'None' and attendanceEmployeesYesterday['CasualLeaveEmployees'] != 'None':
            casualLeaveEmployeesYesterday = attendanceEmployeesYesterday['CasualLeaveEmployees']

            
        todaysdate = datetime.now().strftime('%d-%m-%Y')
        workbook = xlsxwriter.Workbook(f'{path}\\_{todaysdate}.xlsx')
        worksheet = workbook.add_worksheet(todaysdate)
        #Excel Formatting
        bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter','border':2, 'border_color':'black'})
        bold_border_background_colour = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border':2, 'border_color':'black','bg_color': 'yellow'})
        text_wrap = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border':2, 'border_color':'black'})
        center = workbook.add_format({'align': 'center', 'valign': 'vcenter','border':2, 'border_color':'black'})

        #Setting Column Width
        worksheet.set_column("B:B",10)
        worksheet.set_column("C:C",5)
        worksheet.set_column("D:D",50)
        worksheet.set_column("E:E",15)
        worksheet.set_column("F:F",15)
        worksheet.set_column("G:G",15)
        worksheet.set_column("H:H",20)

        #Adding Data to 3rd Row
        worksheet.write('B3', 'Date', bold_border_background_colour)
        worksheet.write('C3', 'Total', bold_border_background_colour)
        worksheet.write('D3', 'At Office', bold_border_background_colour)
        worksheet.write('E3', 'Work From Home', bold_border_background_colour)
        worksheet.write('F3', 'At Customer Site', bold_border_background_colour)
        worksheet.write('G3', 'On Leave Not Sick', bold_border_background_colour)
        worksheet.write('H3', 'Sick', bold_border_background_colour)

        #adding data for yesterday
        worksheet.write('B9', 'Date', bold_border_background_colour)
        worksheet.write('C9', 'Total', bold_border_background_colour)
        worksheet.write('D9', 'At Office', bold_border_background_colour)
        worksheet.write('E9', 'Work From Home', bold_border_background_colour)
        worksheet.write('F9', 'At Customer Site', bold_border_background_colour)
        worksheet.write('G9', 'On Leave Not Sick', bold_border_background_colour)
        worksheet.write('H9', 'Sick', bold_border_background_colour)
        
        
        #Adding Data to 4th Row
        worksheet.write('B4', "", center)
        worksheet.write('C4',attendanceCount['TotalEmployeeCount'] ,center)
        worksheet.write('D4', atOfficeCountTotalToday ,center)
        worksheet.write("E4",wfhLeaveTotalCountToday,center)
        worksheet.write('F4', "", center)
        worksheet.write('G4',casualLeaveTotalCountToday,center)
        worksheet.write('H4',sickLeaveTotalCountToday,center)

        worksheet.write('B10', "", center)
        worksheet.write('C10',attendanceCountYesterday['TotalEmployeeCount'] ,center)
        worksheet.write('D10',atOfficeCountTotalYesterday ,center)
        worksheet.write("E10",wfhCountTotalYesterday ,center)
        worksheet.write('F10', "", center)
        worksheet.write('G10',casualLeaveCountTotalYesterday ,center)
        worksheet.write('H10',sickLeaveCountTotalYesterday ,center)
        
        #Adding data to 5th row
        worksheet.write('B5', todaysdate , bold)
        worksheet.write('C5', "", center)
        if attendanceEmployees['AtOfficeEmployees'] == 'None' and attendanceEmployeesHalfDay["AtOfficeEmployeesHalfDay"] == 'None':
            worksheet.write('D5',"None",text_wrap) 
        else :
            worksheet.write('D5',' , '.join(atofficeEmployeesToday),text_wrap) 

        worksheet.write('E5', "", center)
        worksheet.write('F5', "", center)
        if attendanceEmployees['CasualLeaveEmployees'] == 'None' and attendanceEmployeesHalfDay["CasualLeaveEmployeesHalfDay"] == 'None':
            worksheet.write('G5',"None",text_wrap) 
        else :
            worksheet.write('G5',' , '.join(casualLeaveEmployeesToday),text_wrap) 
            
        if attendanceEmployees['SickLeaveEmployees'] == 'None' and attendanceEmployeesHalfDay["SickLeaveEmployeesHalfDay"] == 'None':
            worksheet.write('H5',"None",text_wrap) 
        else :
            worksheet.write('H5',' , '.join(sickLeaveEmployeesToday),text_wrap) 
            

        #Adding data to 5th row
        if date.today().weekday() == 0:
            yesterday =  datetime.now() - timedelta(3)
            yesterdaysDate = (datetime.strftime(yesterday, '%d/%m/%Y'))
        else :
            yesterday =  datetime.now() - timedelta(1)
            yesterdaysDate = (datetime.strftime(yesterday, '%d/%m/%Y'))

        worksheet.write('B11', yesterdaysDate , bold)
        worksheet.write('C11', "", center)

        if attendanceEmployeesYesterday['AtOfficeEmployees'] == 'None' and attendanceEmployeesHalfDayYesterday['AtOfficeEmployeeshalfday'] == 'None':
            worksheet.write('D11',"None",text_wrap) 
        else :
            worksheet.write('D11',' , '.join(atOfficeEmployeesYesterday),text_wrap) 
        
        worksheet.write('E11', "", center)
        worksheet.write('F11', "", center)

        if attendanceEmployeesYesterday['CasualLeaveEmployees'] == 'None' and attendanceEmployeesHalfDayYesterday['CasualLeaveEmployeeshalfday'] == 'None':
            worksheet.write('G11',"None",text_wrap) 
        else :
            worksheet.write('G11',' , '.join(casualLeaveEmployeesYesterday),text_wrap)
            
        if attendanceEmployeesYesterday['SickLeaveEmployees'] == 'None' and attendanceEmployeesHalfDayYesterday['SickLeaveEmployeeshalfday'] == 'None':
            worksheet.write('H11',"None",text_wrap) 
        else :
            worksheet.write('H11',' , '.join(sickLeaveEmployeesYesterday),text_wrap) 
           
        #worksheet.write('E5',','.join(attendanceEmployees['WorkFromHomeEmployees']),text_wrap)

        workbook.close()
        file = f'{path}\\_{todaysdate}.xlsx'
        return send_file(file,as_attachment= True)
    except Exception as e:
        print(f'Error when downloading report : {e}')
        return str(e)










#All Private Methods
def get_project_list():
    projectList = []
    for value in ReadJson()['project details']:
        projectList.append(value['projectName'])
    return projectList


def ReadJson():
    with open("static/json/pi.json",'r', encoding='utf-8-sig') as json_file:
        json_data = json.load(json_file)
    return json_data


def Admin():
    managers_corpid = []
    for cid in ReadJson()['ManagersList']:
        managers_corpid.append(cid['CorpID'])
    for value in managers_corpid:
        if session['user'] == value:
            pass
            return "Yes"
    return "No"

def get_employeeLevel_list():
    employeeLevelList = []
    for value in ReadJson()['EmployeeLevelDetails']:
        employeeLevelList.append(value['levelName'])
    return employeeLevelList

def dateArrayMethod(year, month):
    dateArray = []
    dict = {'0': 'Mon', '1': 'Tue', '2': 'Wed', '3': 'Thu', '4': 'Fri', '5': 'Sat', '6': 'Sun'}
    cal = calendar.Calendar()

    for x in cal.itermonthdays2(year, month):
        if x[0] != 0:
            dateArray.append(calendar.month_name[month][:3] + " " + str(x[0]) + " " + dict[str(x[1])])
    return dateArray


def gettingInfo(month,year):
    today = date.today()
    numOfDays = calendar.monthrange(year, int(month))
    numOfDaysCfCurrentMonth = numOfDays[1]
    sb = EmployeeProfileDAL()
    employee_list = sb.read_employee()
    employeeStatusListView = []
    HolidayList = []
    HolidayMonth = []
    HolidayDates = []
    listOfDays = list(range(1, numOfDaysCfCurrentMonth+1))
    dateArray = dateArrayMethod(year, int(month))
    i=0
    for value in ReadJson()['waters holidays']:
        HolidayList.append(value['date'].split("/"))
        HolidayMonth.append(HolidayList[i][1])
        if month in HolidayMonth:
            if HolidayList[i][1] == month:
                HolidayDates.append(int(HolidayList[i][0]))
        i=i+1
    for employee in employee_list:
        employeeWorkStatus = []
        counterForOn = 0
        employeeWorkStatus.append(str(employee[0]))
        employeeWorkStatus.append(str(employee[1]))
        employeeWorkStatus.append(employee[2])
        employeeWorkStatus.append(employee[3])
        employeeWorkStatus.append(employee[4])
        employeeWorkStatus.append(employee[5])
        employeeWorkStatus.append(" ")
        employeeWorkStatus.append(" ")
        employeeWorkStatus.append(" ")
        employeeWorkStatus.append(" ")

        with open("static/json/pi.json", 'r', encoding='utf-8-sig') as json_file:
            json_data = json.load(json_file)

        dateArray = dateArrayMethod(year,int(month))
       
        for i in range(numOfDaysCfCurrentMonth):
            dateloop = date(year,int(month),i+1)
            if dateArray[i][-3:] == 'Sat' or dateArray[i][-3:] == 'Sun' or (i+1 in HolidayDates) or (dateloop > today):
                employeeWorkStatus.append(" ")
            else:
                employeeWorkStatus.append("Present")
                counterForOn += 1

        employee_leave_list = sb.read_leaves_type(employee[7], month, year)
        if employee_leave_list is not None:
            counterForFullDay = 0
            counterForHalfDay = 0
            for leave in employee_leave_list:
                numOfDays = calendar.monthrange(year, int(month))
                numOfDaysCfCurrentMonth = numOfDays[1]
                leave_date = str(leave[0])                
                leave_type = leave[1]                
                leavedate = leave_date.split('/')                
                if leave_type == '1' or leave_type == '4':                    
                    employeeWorkStatus[int(leavedate[0]) + 9] = 'FullDayLeave'
                    counterForFullDay += 1                    
                elif leave_type == '2' or leave_type == '5':
                    employeeWorkStatus[int(leavedate[0]) + 9] = 'HalfDayLeave'
                    counterForHalfDay += 1
                elif leave_type == '3':
                    employeeWorkStatus[int(leavedate[0]) + 9] = 'Non-WIPL'
                    # counterForHalfDay += 1
        totalDayOfFullDays = counterForFullDay
        totalDayOfHalfDays = counterForHalfDay
        totalhoursofWork = (counterForOn*8 - (counterForFullDay*8 + counterForHalfDay*4))
        employeeWorkStatus.append(" ")
        employeeWorkStatus.append(str(totalDayOfFullDays))
        employeeWorkStatus.append(str(totalDayOfHalfDays))
        employeeWorkStatus[7] = str(totalhoursofWork)
        employeeWorkStatus[5] = str(round(21.85 * totalhoursofWork, 1))
        employeeWorkStatus[6] = str(21.85)
        employeeStatusListView.append(employeeWorkStatus)
    return employeeStatusListView



def gettingOtherDeductionsInfo(month, year):
    numOfDays = calendar.monthrange(year, int(month))

    startDate =  "1-" + str(month) +"-" + str(year)
    endDate = str(numOfDays[1]) + "-" + str(month) +"-" + str(year)

    otherDeductions = []
    for value in ReadJson()['OtherDeductions']:
        otherDeductions.append(value['PaymentRecovery'])
        otherDeductions.append(value['Amount'])
        otherDeductions.append(value['PaymentRecoveryTowards'])
        otherDeductions.append(value['LetterToBeIssued'])
        otherDeductions.append(value['ApprovalAttached'])
        otherDeductions.append(value['NameOftheAttachment'])
        otherDeductions.append(value['ApproverName'])
        otherDeductions.append(value['RemarksReason'])
        otherDeductions.append(value['TypeOfDeduction'])
        otherDeductions.append(value['MinimumWorkDays'])


    numOfDaysCfCurrentMonth = numOfDays[1]
    sb = EmployeeProfileDAL()
    
    employee_list = sb.read_employee()
    employeeStatusListView = []
    for employee in employee_list:
        employeeWorkStatus = []
        counterForOn = 0
        dateArray = dateArrayMethod(year, int(month))
        for i in range(numOfDaysCfCurrentMonth):
            if dateArray[i][-3:] == 'Sat' or dateArray[i][-3:] == 'Sun':
                test = " "
            else:
                counterForOn += 1

        employee_leave_list = sb.read_leaves_type(employee[7], month, year)
        if employee_leave_list is not None:
            counterForFullDay = 0
            counterForHalfDay = 0
            for leave in employee_leave_list:
                numOfDays = calendar.monthrange(year, int(month))
                numOfDaysCfCurrentMonth = numOfDays[1]
                leave_date = str(leave[0])
                leave_type = leave[1]
                leavedate = leave_date.split('/')
                if leave_type == '1':
                    counterForFullDay += 1  #full day leave
                else:
                    counterForHalfDay += 1 #half day leave
        totalDayOfFullDays = counterForFullDay
        totalDayOfHalfDays = counterForHalfDay
        totalhoursofWork = 0
        totalhoursofWork = (counterForOn * 8 - (counterForFullDay * 8 + counterForHalfDay * 4))
        if(totalhoursofWork < (int(otherDeductions[9]) * 8)):   #if work days is less than 7 days
            print("inside continue")
            continue
        else:
            employeeWorkStatus.append(str(employee[1]))
            employeeWorkStatus.append(otherDeductions[0])
            employeeWorkStatus.append(employee[2])
            employeeWorkStatus.append(otherDeductions[1])
            employeeWorkStatus.append(startDate)
            employeeWorkStatus.append(endDate)
            employeeWorkStatus.append(otherDeductions[2])
            employeeWorkStatus.append(otherDeductions[3])
            employeeWorkStatus.append(otherDeductions[4])
            employeeWorkStatus.append(otherDeductions[5])
            employeeWorkStatus.append(otherDeductions[6])
            employeeWorkStatus.append(otherDeductions[7])
            employeeStatusListView.append(employeeWorkStatus)
    return employeeStatusListView

def setAttendancetableDates():
    sb = EmployeeProfileDAL()
    
    datesAndCount = sb.get_datesCount_in_attendance_table()
    CurrentDateFromAttendanceTable = datesAndCount['CurrentDateFromAttendanceTable']
    YesterdaysDateFromTable = datesAndCount['YesterdaysDateFromTable']
    today = datetime.now()
    
    if date.today().weekday() == 0:
        yesterday =  datetime.now() - timedelta(3)
        yesterdaysDate = (datetime.strftime(yesterday, '%d/%m/%Y'))
    else:
        yesterday =  datetime.now() - timedelta(1)
        yesterdaysDate = (datetime.strftime(yesterday, '%d/%m/%Y'))
    todaysDate = (datetime.strftime(today, '%d/%m/%Y'))    
    
    weekNumber = datetime.today().weekday()
    if weekNumber < 5:
        if YesterdaysDateFromTable != yesterdaysDate:
            sb.updateAttendanceYesterdayTableDate(yesterdaysDate)            
            sb.UpdatingAttendanceYesterdaysAtOfficeRecords()
            sb.UpdatingAttendanceYesterdaysSickLeaveRecords()
            sb.UpdatingAttendanceYesterdaysWorkFromHomeRecords()
            sb.UpdatingAttendanceYesterdaysCasualLeaveRecords()
        if CurrentDateFromAttendanceTable != todaysDate:
            sb.reset_atOffice()
            sb.reset_sickLeave()
            sb.reset_casualLeave()
            sb.reset_workFromHome()
            sb.updateAttendanceTableDate(todaysDate)
            print('Updating future leaves')
            sb.set_future_leaves_for_today()


def getLeaveIdFromLeaveType(sickLeave, casualLeave):
    leaveId = None
    if sickLeave == 'Full Day':
        leaveId = 1
    if sickLeave == 'Half Day':
        leaveId = 2
    if casualLeave == 'Full Day':
        leaveId = 4
    if casualLeave == 'Half Day':
        leaveId = 5
    return leaveId



















# Port Number Configuration

if __name__ == '__main__':
   app.run(host='0.0.0.0',port=80)