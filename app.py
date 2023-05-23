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



app = Flask(__name__)
app.secret_key = os.urandom(24)

@app.route('/')
def home():
   return render_template('loginv4.html')


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
            print(f'EmployeeName : {EmployeeName}')
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
            return render_template("login.html", **locals())
    except Exception as e:
        return str(e)

@app.route('/viewteam')
def viewTeamfun():
    if 'user' in session:
        corp_id=session['user']
        # corp_id="conddas"
        sb=EmployeeProfileDAL()
        EmployeeName=sb.get_current_employee_Info(corp_id)[0][0]
        # projectName=sb.get_current_employee_Info()[0][]
        print("In view team")
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
    print("inside compare method serverside validation")
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
        print("DataBase is closed")
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
        print(f"Delete Request Initiated for Employee Id : {request.args['employeeId']}")
        employeeId = request.args['employeeId']
        sb = EmployeeProfileDAL()
        delete_status = sb.delete_employee(employeeId)
        return delete_status


#Apply Leave API's
@app.route('/personalLeave')
def personalLeave():
    if 'user' in session:
        print("personalLeave")
        corpid = session['user']
        # corpid="conddas"
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
        print("-----------------------------------")
        print("-----------------------------------")
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
        print("applyLeave")
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
        print("personalLeave")
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
    print(managers_corpid)
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



























# Port Number Configuration

if __name__ == '__main__':
   app.run(host='0.0.0.0',port=80)