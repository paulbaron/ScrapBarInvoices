import csv
import os
import re

# Product should have:
# - Categorie
# - Quantité
# - Poids/Volume
# - Montant HT
# - Taxes
# - Promotions
# - TVA

# Compute for each product:
# - Montant Total HT
# - Montant Total TTC
def preprocess_and_sort_products(product_data):
	for _, details in product_data.items():
		montant_ht = details["Montant HT"]
		taxes = details["Taxes"]
		promotions = details["Promotions"]
		tva = details["TVA"]
		montant_total_ht = montant_ht - promotions + taxes
		details["Montant Total HT"] = montant_total_ht
		details["Montant Total TTC"] = montant_total_ht * (1.0 + tva)
	return sorted(
		product_data.items(), 
		key=lambda x: x[1]["Montant Total HT"], 
		reverse=True  # Sort by descending order
	)

def replace_cid_sequences(input_string, iso_variant=1):
	def replace_match(match):
		# Extract the numeric index from the match
		index = int(match.group(1))
		if not (0 <= index <= 255):
			raise ValueError(f"Index {index} out of range for ISO-8859 encoding")
		encoding = f'iso8859_{iso_variant}'
		try:
			# Convert the index to the corresponding character
			return bytes([index]).decode(encoding)
		except LookupError:
			raise ValueError(f"ISO-8859-{iso_variant} encoding not recognized")
	# Use regex to find sequences of the form "(cid:###)"
	pattern = r'\(cid:(\d{1,3})\)'
	# Replace all occurrences using the replace_match function
	return re.sub(pattern, replace_match, input_string)

# Write sorted results in csv table
def write_sorted_products_to_csv(sorted_products, output_file):
	"""
	Write the products in a CSV file.

	Args:
		sorted_products (list): List of sorted tuples (product, details).
		output_file (str): Output CSV path.
	"""
	with open(output_file, mode="w", newline="", encoding="utf-8") as csvfile:
		writer = csv.writer(csvfile)
		# Header
		writer.writerow(["Produit", "Catégorie", "Quantité", "Poids/Volume", "Montant HT", "Taxes", "Pomotions", "Montant Total HT", "TVA", "Montant TTC"])
		# Products and details
		for product, details in sorted_products:
			writer.writerow([
				replace_cid_sequences(product),
				details.get("Categorie", "N/A"),
				details.get("Quantite", "N/A"),
				details.get("Poids/Volume", "N/A"),
				details.get("Montant HT", "N/A"),
				details.get("Taxes", "N/A"),
				details.get("Promotions", "N/A"),
				details.get("Montant Total HT", "N/A"),
				details.get("TVA", "N/A"),
				details.get("Montant Total TTC", "N/A")
			])

def MergeProducts(product_data1, product_data2):
	for key, details in product_data2.items():
		if key in product_data1.keys():
			# Check if parameters are the same:
			if product_data1[key]["Categorie"] != details["Categorie"]:
				print(f"Missmatch in product 'Categorie' for {key}")
			if product_data1[key]["TVA"] != details["TVA"]:
				print(f"Missmatch in product 'TVA' for {key}")
			# Update other values:
			product_data1[key]["Quantite"] += details["Quantite"]
			product_data1[key]["Poids/Volume"] += details["Poids/Volume"]
			product_data1[key]["Montant HT"] += details["Montant HT"]
			product_data1[key]["Taxes"] += details["Taxes"]
			product_data1[key]["Promotions"] += details["Promotions"]
		else:
			product_data1[key] = details

def data_to_int(text, default_value=0):
	if text == None:
		return default_value
	int_str = text.strip()
	if not int_str:
		return default_value
	try:
		return int(int_str)
	except ValueError:
		return default_value

def data_to_float(text, default_value=0):
	if text == None:
		return default_value
	float_str = text.strip().replace(",", ".")
	if not float_str:
		return default_value
	try:
		return float(float_str)
	except ValueError:
		return default_value

def create_or_clear_invoice_dir(download_dir):
	os.makedirs(download_dir, exist_ok=True)
	# Remove all pdf files in the folder
	for file_name in os.listdir(download_dir):
		file_path = os.path.join(download_dir, file_name)
		try:
			if os.path.isfile(file_path) and file_path.lower().endswith(".pdf"):
				os.unlink(file_path)
		except Exception as e:
			print(f"Error while deleting {file_name} : {e}")
