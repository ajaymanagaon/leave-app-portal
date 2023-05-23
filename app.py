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


app = Flask(__name__)
app.secret_key = os.urandom(24)

@app.route('/')
def home():
   return render_template('loginv4.html')






























# Port Number Configuration

if __name__ == '__main__':
   app.run(host='0.0.0.0',port=80)