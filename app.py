from flask import Flask
import os
path = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(path,'uploads')
DOWNLOAD_FOLDER = os.path.join(path,'downloads')
TEMP_FOLDER = os.path.join(path,'temp')
app = Flask(__name__)
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['TEMP_FOLDER'] = TEMP_FOLDER

