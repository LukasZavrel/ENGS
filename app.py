from flask import Flask, flash, request, redirect, render_template, send_file
import os
import urllib.request
from werkzeug.utils import secure_filename
import numpy as np
import pandas as pd
import zipfile
import logging

path = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(path,'uploads')
DOWNLOAD_FOLDER = os.path.join(path,'downloads')
TEMP_FOLDER = os.path.join(path,'temp')
app = Flask(__name__)
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['TEMP_FOLDER'] = TEMP_FOLDER

ALLOWED_EXTENSIONS = set(['xlsx'])
def delete_content(folder):
	for the_file in os.listdir(folder):
		file_path = os.path.join(folder, the_file)
		try:
			if os.path.isfile(file_path):
				os.unlink(file_path)
		except Exception as e:
			print(e)

def allowed_file(filename):
	return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def zpracuj_data(file_path):
	# Mesic pro ktery se operace provadi
	month = file_path[-12:-5]
	xls = pd.ExcelFile(file_path)
	# Lide, pro ktere hledame smlouvy
	people_names = pd.read_excel(xls, 'список МТО')
	# Tabulka smluv prirazenych k lidem
	contract_people = pd.read_excel(xls, 'договора')
	# vsechny tri listy v excelu, zacinajici na tretim radku
	data_tuple = (pd.read_excel(xls, '60.01', header=2), pd.read_excel(xls, '60.21', header=2), pd.read_excel(xls, '60.31', header=2))

	# Vsechny mozne meny
	currencies = ['USD', 'EUR', 'CNY', 'JPY', 'GBP', 'руб']

	# Vsechny mozne jmena firem, prectene z druheho sloupce
	company_names = list(set(contract_people.iloc[:,1].values))
	# Odstraneni prazdnych bunek
	company_names = [name for name in company_names if str(name) != 'nan']

	def add_company(data):
		last_company = ""
		for row in range(len(data)):
			if data.loc[row, "Счет"] in company_names:
				last_company = data.loc[row, "Счет"]
			data.loc[row, "Firma"] = last_company
		return data

	def add_contract_nr(data):  
		last_contract_nr = ""
		for row in range(len(data)):
			if (type(data.loc[row, "Счет"]) is not str and np.isnan(data.loc[row, "Счет"])) or data.loc[row, "Счет"] in currencies:
				data.loc[row, "Smlouva"] = last_contract_nr
			else:
				last_contract_nr = data.loc[row, "Счет"]
		return data

	def filter_by_turnover(data):
		final_rows = []
		for row in range(len(data)):
			if data.iloc[row,1] == 'Оборот':
				if not np.isnan(data.iloc[row,3]):
					final_rows.append(row)
		return data.iloc[final_rows, :]

	def add_currency(data):
		last_currency = ""
		for row in range(len(data)):
			if data.loc[row, "Счет"] in currencies:
				last_currency = data.loc[row, "Счет"]
			data.loc[row, "Mena"] = last_currency    
		return data

	def remove_space(x):
		while type(x) is str and x.endswith(" "):
			x = x[:-1]
		return x

	contracts = contract_people[contract_people['Подготовил'].isin(people_names.values.flatten())].copy()
	contracts.rename(columns={ 'Наименование': "Smlouva", 'Подготовил': "Osoba"}, inplace = True)
	contracts = contracts[['Smlouva','Osoba']]
	contracts = contracts.applymap(remove_space)

	file_names = ['01', '21', '31']
	delete_content(app.config['TEMP_FOLDER'])
	logging.warning('Inside of the function')
	for i in range(3):
		data = data_tuple[i].copy()
		data.rename(columns={ data.columns[3]: "Castka", data.columns[4]: "Firma", data.columns[5]: "Smlouva", data.columns[6]: "Mena"}, inplace = True)
		data = add_company(data)
		data = add_contract_nr(data)
		data = add_currency(data)
		data = filter_by_turnover(data)
		data.loc[:,"Firma":"Mena"] = data.loc[:,"Firma":"Mena"].applymap(remove_space)
		data = data.merge(contracts, on = "Smlouva")
		data = data[["Firma", "Smlouva", "Castka", "Mena", "Osoba"]]
		if file_names[i] == '01':
			data['Mena'] = 'руб'
		last_company = ""
		company_sum = 0
		rows_to_drop = []
		for row in range(len(data)):
			if data.loc[row,"Firma"] != last_company:
				if row > 0 and company_sum/2.01 <= data.loc[row - 1,"Castka"] <= company_sum/1.99:
					rows_to_drop.append(row-1)
				company_sum = data.loc[row,"Castka"]
				last_company = data.loc[row,"Firma"]
			else:
				company_sum += data.loc[row,"Castka"]
			if row == len(data) -1:
				if company_sum/2.01 <= data.loc[row,"Castka"] <= company_sum/1.99:
					rows_to_drop.append(row)
		data_dropped = data.drop(rows_to_drop, axis=0)
		data_dropped.drop_duplicates(inplace=True)
		data_dropped.to_excel(os.path.join(app.config['TEMP_FOLDER'],"ENGS_MTO_"+month+"_"+file_names[i]+".xlsx"), engine='xlsxwriter')

	#zkomprimuj vse do zip archivu
	delete_content(app.config['DOWNLOAD_FOLDER'])
	zf = zipfile.ZipFile(os.path.join(app.config['DOWNLOAD_FOLDER'],"ENGS.zip"), "w")
	for filename in os.listdir(app.config['TEMP_FOLDER']):
		zf.write(os.path.join(app.config['TEMP_FOLDER'], filename), arcname = filename)
	zf.close()

@app.route('/')
def upload_form():
	return render_template('upload.html')


@app.route('/', methods=['POST'])
def upload_file():
	if request.method == 'POST':
        # check if the post request has the file part
		if 'file' not in request.files:
			flash('No file part')
			return redirect(request.url)
		file = request.files['file']
		if file.filename == '':
			flash('No file selected for uploading')
			return redirect(request.url)
		if file and allowed_file(file.filename):
			print(app)
			filename = secure_filename(file.filename)
			delete_content(app.config['UPLOAD_FOLDER'])
			logging.warning('File upload started')
			file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
			logging.warning('File has been uploaded')
			flash('File successfully uploaded')
			try:
				zpracuj_data(os.path.join(app.config['UPLOAD_FOLDER'], filename))
			except Exception as e:
				logging.error(f'{e} raised an error')	
				logging.error("Exception occurred", exc_info=True)	
			logging.warning('After function termination')
			return redirect('/')
		else:
			flash('Allowed file types are xlsx')
			return redirect(request.url)

# @app.route('/file-downloads/')
# def file_downloads():
# 	try:
# 		return render_template('downloads.html')
# 	except Exception as e:
# 		return str(e)

@app.route('/return-files/')
def return_files_tut():
	try:
		return send_file(os.path.join(app.config['DOWNLOAD_FOLDER'],"ENGS.zip"), attachment_filename='ENGS.zip',as_attachment=True)
	except Exception as e:
		return str(e)

if __name__ == "__main__":
    app.run(host='0.0.0.0')
