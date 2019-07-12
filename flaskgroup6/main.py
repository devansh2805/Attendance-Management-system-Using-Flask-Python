from flask import Flask,render_template, redirect, url_for, request, flash
from flask_wtf import FlaskForm
from wtforms import StringField,SubmitField,PasswordField
from wtforms.validators import DataRequired,Length
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
import os, sqlite3, csv
from win32com.client import constants, Dispatch
import pandas as pd
import matplotlib.pyplot as plt


app = Flask(__name__)

app.config['SECRET_KEY']='7f371cc08c51329ecccf896cdf66d590'
app.config['SQLALCHEMY_DATABASE_URI']='sqlite:///test.db'
app.config['SQLALCHEMY_BINDS']={'file':'sqlite:///file.db'}
db=SQLAlchemy(app)

#database
class userinfo(db.Model):
    id=db.Column(db.Integer,primary_key=True)
    name = db.Column(db.String(50), nullable=False)
    surname = db.Column(db.String(50), nullable=False)
    username=db.Column(db.String(50),unique=True,nullable=False)
    password=db.Column(db.String(50),nullable=False)
    def __repr__(self):
        return f"User('{self.name}','{self.surname}','{self.username}')"


class filecontent(db.Model):
    __bind_key__='file'
    id=db.Column(db.Integer,primary_key=True)
    name=db.Column(db.String(300))
    data=db.Column(db.LargeBinary)


#login form
class loginform(FlaskForm):
    username=StringField('Username',validators=[DataRequired(),Length(min=2,max=20)])
    password=PasswordField('Password',validators=[DataRequired()])
    submit=SubmitField('Login')


@app.route('/',methods=['GET','POST'])
def login():
    form=loginform(request.form)
    error=None
    if request.method=="POST" and form.validate_on_submit():
        user=userinfo.query.filter_by(username=form.username.data).first()
        #if username entered is in database its record will be stored in user
        if not(user) or user.password!=form.password.data:
            error = 'Login Unsuccessful.Please check username and password'
        elif(user.password==form.password.data):
            #return "Login successful. Welcome "+user.name+" "+user.surname
            return render_template('home.html',title="Home Page",name=user.name,surname=user.surname)
    return render_template('login.html',title="Login",error=error,form=form)


@app.route("/home")
def home():
    return render_template('home.html', title = 'home')

@app.route("/upload",methods=["GET","POST"])
def upload():
    if request.method=="GET":
        return render_template('upload.html')
    elif request.method=="POST":
        file=request.files['inputFile']
        file.save(secure_filename(file.filename))

        with open("sample.csv","w") as copyfile:
            filewriter = csv.writer(copyfile)
            with open(file.filename,"r") as originalfile:
                filereader=csv.reader(originalfile)
                for row in filereader:
                    filewriter.writerow(row)

        return render_template('h1.html')


@app.route("/attpage",methods=["POST"])
def attpage():
    conn = sqlite3.connect("file.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS exceltable")
    df = pd.read_csv("sample.csv")
    df.to_sql('exceltable',conn,index=False)
    cur.execute("SELECT * FROM exceltable")
    row_data = cur.fetchall()

    return render_template('attendance_page.html', title='attendance page',data_in_sheets=row_data)


@app.route("/report",methods=["POST"])
def report():

    conn = sqlite3.connect("file.db")
    cur = conn.cursor()
    plt.figure(figsize=(5,5))

    count0=0
    count1=0
    count2=0
    count3=0
    count4=0

    attendance = ["a","b","c","d","e"]

    cur.execute("SELECT * FROM exceltable")
    row_data = cur.fetchall()

    for row in row_data:
        x=row[2]+row[3]+row[4]+row[5]
        if x==4:
            count4+=1
        elif x==3:
            count3+=1
        elif x==2:
            count2+=1
        elif x==1:
            count1+=1
        elif x==0:
            count0+=1

    att = [count4,count3,count2,count1,count0]
    plt.pie(att,labels = attendance, autopct="%.1f%%")
    plt.show()
    return render_template('report.html', title='report')

if __name__=="__main__":
    app.run(debug=True)