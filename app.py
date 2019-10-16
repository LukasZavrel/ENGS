from flask import Flask

UPLOAD_FOLDER = '/Users/lukas/ENGS/uploads'
DOWNLOAD_FOLDER = '/Users/lukas/ENGS/downloads'
TEMP_FOLDER = '/Users/lukas/ENGS/temp'
app = Flask(__name__)
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['TEMP_FOLDER'] = TEMP_FOLDER

