from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import pdfplumber
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
import re
import ProductsUtils

def scrap_invoices(download_dir, email_address, password, start_date, end_date):
	ProductsUtils.create_or_clear_invoice_dir(download_dir)
	# Set up Selenium WebDriver
	service = Service()
	options = webdriver.ChromeOptions()
	prefs = {
		"download.default_directory": os.path.abspath(download_dir),
		"download.prompt_for_download": False,
		"plugins.always_open_pdf_externally": True
	}
	options.add_experimental_option("prefs", prefs)
	# options.add_argument("--headless")  # Optional: Run in headless mode
	driver = webdriver.Chrome(service=service, options=options)
	try:
		# Open the Metro login page
		driver.get("https://docs.metro.fr/")
		wait = WebDriverWait(driver, 10)  # Timeout after 10 seconds
		# Localize the cookie disclaimer
		cookie_disclaimer = wait.until(
			EC.presence_of_element_located((By.CSS_SELECTOR, "cms-cookie-disclaimer"))
		)
		# Access shadow root
		shadow_root = driver.execute_script("return arguments[0].shadowRoot", cookie_disclaimer)
		# Click the accept all cookies button
		accept_button = shadow_root.find_element(By.CSS_SELECTOR, "button.accept-btn.btn-primary")
		accept_button.click()
		# Wait for the login page to load
		wait.until(
			EC.presence_of_element_located((By.ID, "user_id"))  # Replace with the actual email input ID
		)
		# Enter email and password
		email_input = driver.find_element(By.ID, "user_id")  # Adjust ID or selector as needed
		email_input.send_keys(email_address)
		password_input = driver.find_element(By.ID, "password")  # Adjust ID or selector as needed
		password_input.send_keys(password)
		# Submit the form
		login_button = driver.find_element(By.ID, "submit")  # Replace with the login button ID
		login_button.click()
		# Update the date range for downloading the invoices	
		date_inputs = wait.until(
			EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input[data-testid='DateInputFieldInput']"))
		)
		date_inputs[0].click()
		date_inputs[0].send_keys(Keys.CONTROL + "a")
		date_inputs[0].send_keys(Keys.DELETE)
		date_inputs[0].send_keys(start_date)
		date_inputs[0].send_keys(Keys.RETURN)
		if (end_date):
			date_inputs[1].click()
			date_inputs[1].send_keys(Keys.CONTROL + "a")
			date_inputs[1].send_keys(Keys.DELETE)
			date_inputs[1].send_keys(end_date)
			date_inputs[1].send_keys(Keys.RETURN)
		dropdown = wait.until(
			EC.presence_of_element_located((By.ID, "invoiceLimitId"))
		)
		select = Select(dropdown)
		select.select_by_value("100")
		# Ugly sleep to remove (could not found a better way to make this work for now)
		time.sleep(5)
		# Localize all buttons with data-testid="downloadPdfButton"
		pdf_buttons = wait.until(
			EC.presence_of_all_elements_located((By.CSS_SELECTOR, "button[data-testid='downloadPdfButton']"))
		)
		# Click on each button to download invoices
		for idx, button in enumerate(pdf_buttons):
			wait.until(
				EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-testid='downloadPdfButton']"))
			)
			button.click()
			print(f"Downloading invoice : {idx + 1}")
			# Wait for the file to download (ensure a proper wait if required)
			wait.until(lambda d: len(os.listdir(download_dir)) > idx)
	finally:
		driver.quit()

def extract_invoice_data(pdf_path, product_data):
	print(f"extract data from file: {pdf_path}")
	"""
	Détection des lignes contenant des produits
	Décomposition regex:
	(\\d+)				-> EAN (1)
	\\s+
	(\\d+)				-> N# (2)
	\\s+
	(.+)				-> Nom (3)
	\\s+
	([A-Z]\\s+)?		-> Régie (4)
	(\\d?\\d,\\d\\s+)?	-> Vol % (5)
	(\\d+,\\d+\\s+)?	-> VAP (6)
	(\\d+,\\d+\\s+)?	-> Poids/Volume (7), will be captured in group 6 if VAP is not existing
	(\\d+,\\d+)			-> Prix Unitaire (8)
	\\s+
	(\\d+\\s+)?			-> Colisage (9)
	(\\d+)				-> Qté (10)
	\\s+
	(\\d+,\\d+)			-> Montant (11)
	\\s+
	([A-D])				-> TVA (12)
	"""
	tva_letter_to_rate = {'A': 0, 'B': 0.055, 'C': 0.2, 'D': 0.2}
	regex = re.compile(
		r"(\d+\s+)?(\d+)\s+(.+?)\s+([A-Z]\s+)?(\d?\d,\d\s+)?(\d+,\d+\s+)?(\d+,\d+\s+)?(\d+,\d+)\s+(\d+\s+)?(\d+)\s+(\d+,\d+)\s+([A-D])"
	)
	with pdfplumber.open(pdf_path) as pdf:
		products_in_category = []
		for page in pdf.pages:
			text = page.extract_text(y_tolerance=0)
			lines = text.split("\n")
			current_product = None
			for line in lines:
				product_match = regex.match(line)
				if product_match:
					product_name = product_match.group(3).strip()
					group7 = product_match.group(7)
					if product_match.group(6) and not group7:
						group7 = product_match.group(6)
					weight_or_volume = ProductsUtils.data_to_float(group7)
					colisage = ProductsUtils.data_to_int(product_match.group(9), 1)
					quantity = ProductsUtils.data_to_int(product_match.group(10))
					total_product_count = colisage * quantity
					total_ht = ProductsUtils.data_to_float(product_match.group(11))
					tva = product_match.group(12).strip()
					tva_rate = tva_letter_to_rate[tva]
					if product_name not in product_data:
						product_data[product_name] = {
							"Quantite": total_product_count,
							"Poids/Volume": total_product_count * weight_or_volume,
							"Montant HT": total_ht,
							"Taxes": 0,
							"Promotions": 0,
							"TVA": tva_rate
						}
					else:
						product_data[product_name]["Quantite"] += total_product_count
						product_data[product_name]["Poids/Volume"] += total_product_count * weight_or_volume
						product_data[product_name]["Montant HT"] += total_ht
					current_product = product_name
					products_in_category.append(product_name)
				# Handling additional taxes or promotions
				cotis_sociale = re.match(r"Plus : COTIS\. SECURITE SOCIALE\s+(\d+,\d+)\s+([A-D])", line)
				discount = re.match(r"Offre Achetez Plus Payez Moins\s+(\d+,\d+)-", line)
				if cotis_sociale and current_product:
					to_add = float(cotis_sociale.group(1).replace(",", "."))
					product_data[current_product]["Taxes"] += to_add
				if discount and current_product:
					to_sub = float(discount.group(1).replace(",", "."))
					product_data[current_product]["Promotions"] += to_sub
				category_match = re.match(r"\*\*\*\s+(.+?)\s+Total:\s+(\d+,\d+)", line)
				if category_match:
					for product in products_in_category:
						product_data[product]["Categorie"] = category_match.group(1).strip()
					products_in_category.clear()

def get_invoices_data(download_dir, product_data):
	# Process PDFs with pdfplumber
	for pdf_file in os.listdir(download_dir):
		if pdf_file.lower().endswith(".pdf"):
			pdf_path = os.path.join(download_dir, pdf_file)
			extract_invoice_data(pdf_path, product_data)
