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
    department = db.Column(db.String(50), nullable=False)
    def __repr__(self):
        return f"User('{self.name}','{self.surname}','{self.username}','{self.department}')"

        
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
            return render_template('home.html',title="Home Page",name=user.name,surname=user.surname,department=user.department)
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

        with open("fybC.csv","w") as copyfile:
            filewriter = csv.writer(copyfile)
            with open(file.filename,"r") as originalfile:
                filereader=csv.reader(originalfile)
                for row in filereader:
                    filewriter.writerow(row)

        return render_template('h.html')

@app.route("/upload1",methods=["GET","POST"])
def upload1():
    if request.method=="GET":
        return render_template('upload.html')
    elif request.method=="POST":
        file=request.files['inputFile']
        file.save(secure_filename(file.filename))

        with open("fybLC2.csv","w") as copyfile:
            filewriter = csv.writer(copyfile)
            with open(file.filename,"r") as originalfile:
                filereader=csv.reader(originalfile)
                for row in filereader:
                    filewriter.writerow(row)

        return render_template('h1.html')

@app.route("/upload2",methods=["GET","POST"])
def upload2():
    if request.method=="GET":
        return render_template('upload.html')
    elif request.method=="POST":
        file=request.files['inputFile']
        file.save(secure_filename(file.filename))

        with open("fybLJ1.csv","w") as copyfile:
            filewriter = csv.writer(copyfile)
            with open(file.filename,"r") as originalfile:
                filereader=csv.reader(originalfile)
                for row in filereader:
                    filewriter.writerow(row)

        return render_template('h2.html')

@app.route("/upload3",methods=["GET","POST"])
def upload3():
    if request.method=="GET":
        return render_template('upload.html')
    elif request.method=="POST":
        file=request.files['inputFile']
        file.save(secure_filename(file.filename))

        with open("fybLD2.csv","w") as copyfile:
            filewriter = csv.writer(copyfile)
            with open(file.filename,"r") as originalfile:
                filereader=csv.reader(originalfile)
                for row in filereader:
                    filewriter.writerow(row)

        return render_template('h3.html')

@app.route("/upload4",methods=["GET","POST"])
def upload4():
    if request.method=="GET":
        return render_template('upload.html')
    elif request.method=="POST":
        file=request.files['inputFile']
        file.save(secure_filename(file.filename))

        with open("fybTC2.csv","w") as copyfile:
            filewriter = csv.writer(copyfile)
            with open(file.filename,"r") as originalfile:
                filereader=csv.reader(originalfile)
                for row in filereader:
                    filewriter.writerow(row)

        return render_template('h4.html')

@app.route("/upload5",methods=["GET","POST"])
def upload5():
    if request.method=="GET":
        return render_template('upload.html')
    elif request.method=="POST":
        file=request.files['inputFile']
        file.save(secure_filename(file.filename))

        with open("fybTJ1.csv","w") as copyfile:
            filewriter = csv.writer(copyfile)
            with open(file.filename,"r") as originalfile:
                filereader=csv.reader(originalfile)
                for row in filereader:
                    filewriter.writerow(row)

        return render_template('h5.html')

@app.route("/upload6",methods=["GET","POST"])
def upload6():
    if request.method=="GET":
        return render_template('upload.html')
    elif request.method=="POST":
        file=request.files['inputFile']
        file.save(secure_filename(file.filename))

        with open("fybTD2.csv","w") as copyfile:
            filewriter = csv.writer(copyfile)
            with open(file.filename,"r") as originalfile:
                filereader=csv.reader(originalfile)
                for row in filereader:
                    filewriter.writerow(row)

        return render_template('h6.html')

@app.route("/upload7",methods=["GET","POST"])
def upload7():
    if request.method=="GET":
        return render_template('upload.html')
    elif request.method=="POST":
        file=request.files['inputFile']
        file.save(secure_filename(file.filename))

        with open("tybB1.csv","w") as copyfile:
            filewriter = csv.writer(copyfile)
            with open(file.filename,"r") as originalfile:
                filereader=csv.reader(originalfile)
                for row in filereader:
                    filewriter.writerow(row)

        return render_template('h7.html')


@app.route("/attpage",methods=["POST"])
def attpage():
    conn = sqlite3.connect("file.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS ex")
    df = pd.read_csv("fybC.csv")
    df.to_sql('ex',conn,index=False)
    cur.execute("SELECT * FROM ex")
    row_data = cur.fetchall()

    return render_template('attendance_page.html', title='attendance page',data_in_sheets=row_data)

@app.route("/attpage1",methods=["POST"])
def attpage1():
    conn = sqlite3.connect("file.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS ex1")
    df = pd.read_csv("fybLC2.csv")
    df.to_sql('ex1',conn,index=False)
    cur.execute("SELECT * FROM ex1")
    row_data = cur.fetchall()

    return render_template('attendance_page1.html', title='attendance page',data_in_sheets=row_data)

@app.route("/attpage2",methods=["POST"])
def attpage2():
    conn = sqlite3.connect("file.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS exc2")
    df = pd.read_csv("fybLJ1.csv")
    df.to_sql('ex2',conn,index=False)
    cur.execute("SELECT * FROM ex2")
    row_data = cur.fetchall()

    return render_template('attendance_page2.html', title='attendance page',data_in_sheets=row_data)

@app.route("/attpage3",methods=["POST"])
def attpage3():
    conn = sqlite3.connect("file.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS ex3")
    df = pd.read_csv("fybLD2.csv")
    df.to_sql('ex3',conn,index=False)
    cur.execute("SELECT * FROM ex3")
    row_data = cur.fetchall()

    return render_template('attendance_page3.html', title='attendance page',data_in_sheets=row_data)

@app.route("/attpage4",methods=["POST"])
def attpage4():
    conn = sqlite3.connect("file.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS ex4")
    df = pd.read_csv("fybTC2.csv")
    df.to_sql('ex4',conn,index=False)
    cur.execute("SELECT * FROM ex4")
    row_data = cur.fetchall()

    return render_template('attendance_page4.html', title='attendance page',data_in_sheets=row_data)

@app.route("/attpage5",methods=["POST"])
def attpage5():
    conn = sqlite3.connect("file.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS ex5")
    df = pd.read_csv("fybTJ1.csv")
    df.to_sql('ex5',conn,index=False)
    cur.execute("SELECT * FROM ex5")
    row_data = cur.fetchall()

    return render_template('attendance_page5.html', title='attendance page',data_in_sheets=row_data)

@app.route("/attpage6",methods=["POST"])
def attpage6():
    conn = sqlite3.connect("file.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS ex6")
    df = pd.read_csv("fybTD2.csv")
    df.to_sql('ex6',conn,index=False)
    cur.execute("SELECT * FROM ex6")
    row_data = cur.fetchall()

    return render_template('attendance_page6.html', title='attendance page',data_in_sheets=row_data)

@app.route("/attpage7",methods=["POST"])
def attpage7():
    conn = sqlite3.connect("file.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS ex7")
    df = pd.read_csv("tybB1.csv")
    df.to_sql('ex7',conn,index=False)
    cur.execute("SELECT * FROM ex7")
    row_data = cur.fetchall()

    return render_template('attendance_page7.html', title='attendance page',data_in_sheets=row_data)

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

    attendance = ["Above 75","50 to 75","25 to 50","below 25","Zero"]

    cur.execute("SELECT * FROM ex")
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

@app.route("/report1",methods=["POST"])
def report1():

    conn = sqlite3.connect("file.db")
    cur = conn.cursor()
    plt.figure(figsize=(5,5))

    count0=0
    count1=0
    count2=0
    count3=0
    count4=0

    attendance = ["Above 75","50 to 75","25 to 50","below 25","Zero"]

    cur.execute("SELECT * FROM ex1")
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
    return render_template('report1.html', title='report1')

@app.route("/report2",methods=["POST"])
def report2():

    conn = sqlite3.connect("file.db")
    cur = conn.cursor()
    plt.figure(figsize=(5,5))

    count0=0
    count1=0
    count2=0
    count3=0
    count4=0

    attendance = ["Above 75","50 to 75","25 to 50","below 25","Zero"]

    cur.execute("SELECT * FROM ex2")
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
    return render_template('report2.html', title='report2')

@app.route("/report3",methods=["POST"])
def report3():

    conn = sqlite3.connect("file.db")
    cur = conn.cursor()
    plt.figure(figsize=(5,5))

    count0=0
    count1=0
    count2=0
    count3=0
    count4=0

    attendance = ["Above 75","50 to 75","25 to 50","below 25","Zero"]

    cur.execute("SELECT * FROM ex3")
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
    return render_template('report3.html', title='report3')

@app.route("/report4",methods=["POST"])
def report4():

    conn = sqlite3.connect("file.db")
    cur = conn.cursor()
    plt.figure(figsize=(5,5))

    count0=0
    count1=0
    count2=0
    count3=0
    count4=0

    attendance = ["Above 75","50 to 75","25 to 50","below 25","Zero"]

    cur.execute("SELECT * FROM ex4")
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
    return render_template('report4.html', title='report4')

@app.route("/report5",methods=["POST"])
def report5():

    conn = sqlite3.connect("file.db")
    cur = conn.cursor()
    plt.figure(figsize=(5,5))

    count0=0
    count1=0
    count2=0
    count3=0
    count4=0

    attendance = ["Above 75","50 to 75","25 to 50","below 25","Zero"]

    cur.execute("SELECT * FROM ex5")
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
    return render_template('report5.html', title='report5')

@app.route("/report6",methods=["POST"])
def report6():

    conn = sqlite3.connect("file.db")
    cur = conn.cursor()
    plt.figure(figsize=(5,5))

    count0=0
    count1=0
    count2=0
    count3=0
    count4=0

    attendance = ["Above 75","50 to 75","25 to 50","below 25","Zero"]

    cur.execute("SELECT * FROM ex6")
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
    return render_template('report6.html', title='report6')

@app.route("/report7",methods=["POST"])
def report7():

    conn = sqlite3.connect("file.db")
    cur = conn.cursor()
    plt.figure(figsize=(5,5))

    count0=0
    count1=0
    count2=0
    count3=0
    count4=0

    attendance = ["Above 75","50 to 75","25 to 50","below 25","Zero"]

    cur.execute("SELECT * FROM ex7")
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
    return render_template('report7.html', title='report7')


if __name__=="__main__":
    app.run(debug=True)