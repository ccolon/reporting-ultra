#!/usr/bin/env python
# -*- coding: utf-8 -*-


import os
from flask import Flask, request, redirect, url_for
from werkzeug.utils import secure_filename
from parse import prepare_data, filter_dates, filter_trials, write_xls
from flask import send_from_directory
from datetime import datetime

UPLOAD_FOLDER = os.environ['UPLOAD_FOLDER']
ALLOWED_EXTENSIONS = set(['txt'])

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'],
                               filename)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # if user does not select file, browser also
        # submit a empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            data = prepare_data(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            start_date = request.form.get('startdate')
            
            start_date = datetime.strptime(start_date + '_00:00:00', "%Y-%m-%d_%H:%M:%S")
            end_date = datetime.strptime('2017-12-31_23:59:59', "%Y-%m-%d_%H:%M:%S")
            data = filter_dates(data, start_date, end_date)            
            data = filter_trials(data)
            if not len(data):
                return '''
                <!doctype html>
                <title>Error!</title>
                <h1>Ooops! Problem with the data.</h1>
                '''
            final_filename = filename + '_output.xlsx'
            write_xls(data, os.path.join(app.config['UPLOAD_FOLDER'], final_filename))
            return redirect(url_for('uploaded_file',
                                    filename=final_filename))
    return '''
    <!doctype html>
    <title>Upload new File</title>
    <h1>Upload new File</h1>
    <form method=post enctype=multipart/form-data>
      <p><input type=file name=file>
        Start Date:<input type=text name=startdate value="2017-01-01">
         <input type=submit value=Upload>
    </form>
    '''


