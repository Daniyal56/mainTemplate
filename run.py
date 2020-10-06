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
        pi_no = request.form['query']

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
            pprint(x)
            doc.render(x)

# # time.sleep(1)

            file_stream = StringIO()
            doc.save('./static/file.docx')
            convert('./static/file.docx',
                    './static/file.pdf')
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
        if file.filename == '':
            print('no filename')
            return redirect(request.url)
        else:
            filename = secure_filename(file.filename)
            file.save(os.path.join(
                app.config['UPLOAD_FOLDER'], filename))
            print("saved file successfully")
        pythoncom.CoInitialize()
        return redirect('/uploadfile/{}/{}/'.format(filename, inv))
    return render_template('Invoice.html')


@ app.route("/uploadfile/<filename>/<inv>", methods=["GET", "POST"])
def download_inv_file(filename, inv):
    x = requests.get(
        'http://151.80.237.86:1251/ords/zkt/pi_doc/doc?pi_no={}'.format(str(inv)))
    data = data.json()

    doc = DocxTemplate(
        "./{}".format(
            filename))


#   take_input = int(input('Please enter your invoice: '))
    pythoncom.CoInitialize()
    for x in data['items']:
        if x['invno'].strip() == '{}'.format(str(inv)):
            doc.render(x)
            file_stream = StringIO()
# time.sleep(1)

            doc.save('./static/static-base/file.docx')
            convert('./static/static-base/file.docx',
                    './static/static-base/file.pdf')
    return render_template('Invoice.html')


if __name__ == "__main__":
    app.run(port=1252, debug=False, threaded=True)
