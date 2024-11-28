import imaplib
import email
from email.header import decode_header
import os
import pdfplumber
import re
import ProductsUtils
from datetime import datetime

# Function to convert DD.MM.YYYY to DD-MMM-YYYY
def format_date_for_imap(date):
	return datetime.strptime(date, "%d.%m.%Y").strftime("%d-%b-%Y")

def scrap_invoices(download_dir, email_address, password, start_date, end_date):
	ProductsUtils.create_or_clear_invoice_dir(download_dir)
	# Connect to Gmail server
	imap = imaplib.IMAP4_SSL("imap.gmail.com")
	# Authentification
	imap.login(email_address, password)
	# Select inbox
	imap.select("inbox")
	# Convert dates to IMAP format
	start_date_imap = format_date_for_imap(start_date)
	end_date_imap = format_date_for_imap(end_date)
	# Search for "Notifications@uba.paris" with object containing "Facture"
	search_criteria = f'FROM "Notifications@uba.paris" SUBJECT "Facture" SINCE "{start_date_imap}" BEFORE "{end_date_imap}"'
	status, messages = imap.search(None, search_criteria)
	# For each email
	for mail_id in messages[0].split():
		# Retrieve email
		status, msg_data = imap.fetch(mail_id, "(RFC822)")
		for response_part in msg_data:
			if isinstance(response_part, tuple):
				# Decode email
				msg = email.message_from_bytes(response_part[1])
				subject = decode_header(msg["Subject"])[0][0]
				encoding = decode_header(msg["Subject"])[0][1]  # Récupérer l'encodage
				if isinstance(subject, bytes):
					try:
						subject = subject.decode(encoding or "utf-8", errors="replace")  # Remplace les caractères invalides
					except Exception as e:
						subject = subject.decode("latin1", errors="replace")  # Essayer un autre encodage (latin1)
				if msg.is_multipart():
					for part in msg.walk():
						# Check attachements
						if part.get_content_disposition() == "attachment":
							filename = part.get_filename()
							if filename and filename.lower().endswith(".pdf"):
								# Enregistrer la pièce jointe
								filepath = os.path.join(download_dir, filename)
								print(f"Downloading invoice : {filename}")
								with open(filepath, "wb") as f:
									f.write(part.get_payload(decode=True))
	# Disconnect
	imap.logout()

def check_row_valid(row):
	return row[1] and row[2] and row[5] and row[12] and row[13]

def extract_invoice_data(pdf_path, product_data):
	print(f"extract data from file: {pdf_path}")
	"""
	([A-Z0-9]+)\s+						-> Code (1)
	(.+?)\s+							-> Désignation (2)
	(-?\d+)\s+(FUT|CAR|CAI|BT|EMB)\s+	-> Qté Livrée (3-4)
	(-?\d+)\s+(L|BT|BOI|EMB)\s+			-> Quantité (L) (5-6)
	(\d+,\d+)\s+						-> Prix Unitaire HTHD (7)
	(-?\d+,\d+)\s+						-> Montant HTHD (8)
	(\d+,\d+\s*\%\s+)?					-> Remise (%) (9) ?
	(\d+,\d+\s+)?						-> Droit Unitaire (10) ?
	(-?\d+,\d+\s+)?						-> Consigne (11) ?
	(-?\d+,\d+\s+)?						-> Deconsigne (12) ?
	(\d+,\d\d\s*\s+)?					-> Alcool (%) (13) ?
	(\d+,\d\d)\s+						-> Cont Unit (14)
	(-?\d+,\d\d\s+)						-> Volume effectif (15)
	(-?\d+,\d\d\s+)?					-> Alcool pur (16) ?
	(-?\d+,\d\d\s+)						-> Poids KG (17)
	([1-3])								-> TVA (18)
	"""
	# First we use a regex to get the TVA value that we cannot get directly in the table:
	tva_index_to_rate = {1: 0.2, 2: 0.055, 3: 0}
	products_tva = {}
	regex = re.compile(
		r"([A-Z0-9]+)\s+(.+?)\s+(-?\d+)\s+(FUT|CAR|CAI|BT|EMB)\s(-?\d+)\s+(L|BT|BOI|EMB)\s+(\d+,\d+)\s+(-?\d+,\d+)\s+(\d+,\d+\s*\%\s+)?(\d+,\d+\s+)?(-?\d+,\d+\s+)?(-?\d+,\d+\s+)?(\d+,\d\d\s*\s+)?(\d+,\d\d)\s+(-?\d+,\d\d\s+)(-?\d+,\d\d\s+)?(-?\d+,\d\d\s+)([1-3])"
	)
	with pdfplumber.open(pdf_path) as pdf:
		for page in pdf.pages:
			text = page.extract_text()
			lines = text.split("\n")
			for line in lines:
				product_match = regex.match(line)
				if product_match:
					product_key = product_match.group(1) + product_match.group(2)
					tva_idx = ProductsUtils.data_to_int(product_match.group(18))
					tva_rate = tva_index_to_rate[tva_idx]
					products_tva[product_key] = tva_rate
	# Then we retrieve the actual products info and match with the tva index:
	tables = extract_all_tables_from_pdf(pdf_path)
	for table in tables:
		# This is the products table:
		if len(table[0]) == 16 and table[0][0] == "CODE":
			# Get all products:
			for rowIdx in range(1, len(table)):
				if not check_row_valid(table[rowIdx]):
					continue
				product_code = table[rowIdx][0] # Product code
				product_name = table[rowIdx][1] # Product name
				product_key = product_code + product_name
				product_tva = 0.2
				if product_key not in products_tva.keys():
					print(f"Could not find the TVA for product {product_name}, fallback to TVA 20%")
				else:
					product_tva = products_tva[product_key]
				match_quantity = re.match(r"(-?\d+)\s+(FUT|CAR|CAI|BT|EMB)", table[rowIdx][2])
				if not match_quantity:
					print("Error: Could not find the product quantity")
					continue
				quantity = ProductsUtils.data_to_int(match_quantity.group(1)) # Quantity
				unit_price = ProductsUtils.data_to_float(table[rowIdx][5]) # Price per cont_unit
				unit_tax = ProductsUtils.data_to_float(table[rowIdx][8]) # Tax per cont_unit
				cont_unit = ProductsUtils.data_to_float(table[rowIdx][12])
				total_volume = ProductsUtils.data_to_float(table[rowIdx][13]) # Total volume
				total_price = total_volume / cont_unit * unit_price # Total price
				total_tax = total_volume / cont_unit * unit_tax # Total tax
				if product_name not in product_data:
					product_data[product_name] = {
						"Quantite": quantity,
						"Poids/Volume": total_volume,
						"Montant HT": total_price,
						"Taxes": total_tax,
						"Promotions": 0,
						"TVA": product_tva,
						"Categorie": "UBA"
					}
				else:
					product_data[product_name]["Quantite"] += quantity
					product_data[product_name]["Poids/Volume"] += total_volume
					product_data[product_name]["Taxes"] += total_tax
					product_data[product_name]["Montant HT"] += total_price

def extract_all_tables_from_pdf(pdf_path):
	tables = []
	with pdfplumber.open(pdf_path) as pdf:
		for page in pdf.pages:
			tables += page.extract_tables()
		return tables

def get_invoices_data(download_dir, product_data):
	# Process PDFs with pdfplumber
	for pdf_file in os.listdir(download_dir):
		if pdf_file.lower().endswith(".pdf"):
			pdf_path = os.path.join(download_dir, pdf_file)
			extract_invoice_data(pdf_path, product_data)
