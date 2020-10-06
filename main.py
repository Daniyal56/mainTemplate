import os
from werkzeug.utils import secure_filename
from flask import Flask, flash, request, redirect, send_file, render_template, url_for
from docxtpl import DocxTemplate
import requests
import json
from io import StringIO
from docx2pdf import convert
import os
from flask_caching import Cache
import win32com.client as win32
from os import path
import pythoncom
from pprint import pprint
from wtforms import StringField, Form, validators


# word = win32.DispatchEx("Word.Application")
UPLOAD_FOLDER = './'
app = Flask(__name__, template_folder='templates')
app.config['SECRET_KEY'] = '39ee479d4f5e752eb38e4b2cbce2c40b1427b967'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
cache = Cache(app, config={'CACHE_TYPE': 'simple'})


@cache.cached(timeout=3)
@app.route('/garibsons', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            print('no file')
            return redirect(request.url)
        file = request.files['file']
        pi_no = str(request.form['query'])

        if file.filename == '':
            print('no filename')
            return redirect(request.url)
        else:
            filename = secure_filename(file.filename)
            file.save(os.path.join(
                app.config['UPLOAD_FOLDER'], filename))
            print("saved file successfully")
        pythoncom.CoInitialize()
        return redirect('/uploadfile/{}/{}'.format(filename, pi_no))

    return render_template('upload_file.html')

#  @ app.route("/pi_no/<pi_no>", methods=["GET", "POST"])
#  def pi_no():
#      return render_template('display.html', pi_no=pi_no)


@ app.route("/uploadfile/<filename>/<pi_no>", methods=["GET", "POST"])
def download_file(filename, pi_no):

    print(pi_no)
    x = requests.get(
        'http://151.80.237.86:1251/ords/zkt/pi_doc/doc?pi_no={}'.format(str(pi_no)))
    data = x.json()
    doc = DocxTemplate(
        "./{}".format(
            filename))

    pythoncom.CoInitialize()
    for x in data['items']:
        if x['pi_no'].strip() == '{}'.format(pi_no):  # 17865
            # pprint(x)
            doc.render(x)

# # time.sleep(1)

            file_stream = StringIO()
            doc.save('./static/file.docx')
            convert('./static/file.docx',
                    './static/file.pdf')
    flash("Your file has been created, Now you can download!!")

    return render_template('upload_file.html')


@ cache.cached(timeout=3)
@ app.route('/invoice', methods=['GET', 'POST'])
def invoice_upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            print('no file')
            return redirect(request.url)
        file = request.files['file']
        inv = request.form['inv']
        yr = request.form['yr']
        if file.filename == '':
            print('no filename')
            return redirect(request.url)
        else:
            filename = secure_filename(file.filename)
            file.save(os.path.join(
                app.config['UPLOAD_FOLDER'], filename))
            print("saved file successfully")
        pythoncom.CoInitialize()
        # return redirect('/uploadfiles/{}/{}'.format(filename, invoice.translate({ord('/'): None})))
        return redirect('/uploadfiles/{}/{}/{}'.format(filename, inv, yr))
    return render_template('Invoice.html')


@ app.route("/uploadfiles/<filename>/<inv>/<yr>", methods=["GET", "POST"])
def download_inv_file(filename, inv, yr):
    x = requests.get(
        'http://151.80.237.86:1251/ords/zkt/pi_doc/doc?invno={}/{}'.format(inv, yr))
    data = x.json()

    doc = DocxTemplate(
        "./{}".format(
            filename))


#   take_input = int(input('Please enter your invoice: '))
    pythoncom.CoInitialize()
    for x in data['items']:
        if x['invno'].strip() == '{}/{}'.format(str(inv), str(yr)):  # 17865
            # if x['invno'].strip() == '{}'.format(str(inv)):
            doc.render(x)
            file_stream = StringIO()
# time.sleep(1)

            doc.save('./static/static-base/file.docx')
            convert('./static/static-base/file.docx',
                    './static/static-base/file.pdf')
    flash("Your file has been created, Now you can download!!")
    return render_template('Invoice.html')


@ app.route('/docfile', methods=['GET', 'POST'])
def open_word():
    x = os.startfile(
        r"C:\Users\Danyal\Desktop\Arwentech\mainweb\GS-Dictionary.docx")
    return render_template("upload_file.html")


if __name__ == "__main__":
    app.run(port=5000, debug=True, threaded=True)
