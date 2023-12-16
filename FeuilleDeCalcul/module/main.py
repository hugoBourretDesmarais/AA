# main.py

import os
import re
from datetime import datetime
from collections import defaultdict
import pdfplumber
import warnings
from openpyxl import load_workbook

# Ignore specific UserWarning from openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet._reader")

def main_menu():
    print('''    _    _                 _       _                      _                  _   
   / \  | |_   _ _ __ ___ (_)_ __ (_)_   _ _ __ ___      / \   ___  ___ ___ | |_ 
  / _ \ | | | | | '_ ` _ \| | '_ \| | | | | '_ ` _ \    / _ \ / __|/ __/ _ \| __|
 / ___ \| | |_| | | | | | | | | | | | |_| | | | | | |  / ___ \\__ \ (_| (_) | |_ 
/_/   \_\_|\__,_|_| |_| |_|_|_| |_|_|\__,_|_| |_| |_| /_/   \_\___/\___\___/ \__|
                                                                                 
''')
    while True:
        print("\nMenu Principal")
        print("1. Analyse des fichiers Excel")
        print("0. Quitter")
        
        choice = input("Entrez votre choix: ")
        
        if choice == "1":
            process_excel_files()
        elif choice == "3":
            process_pdf_files()
        elif choice == "0":
            print("Fermeture du programme.")
            break
        else:
            print("Choix invalide, veuillez réessayer.")

def print_results(year, excel_count, mobilisation_sum, outils_sum):
    border = "+" + "-" * 60 + "+"
    title = f" RÉSULTATS DE L'ANALYSE {year} "
    formatted_mobilisation_sum = "${:,.2f}".format(mobilisation_sum)  # With comma as thousand separator, 2 decimal places, and a dollar sign
    formatted_outils_sum = "${:,.2f}".format(outils_sum)

    print(border)
    print(f"|{title.center(60)}|")
    print(border)
    print(f"| Année:                 {str(year).rjust(20)}{' ' * 20}|")
    print(f"| Nombre Total de Projet: {str(excel_count).ljust(20)}{' ' * 20}|")
    print(f"| Mobilisation Total:     {formatted_mobilisation_sum.rjust(20)}{' ' * 20}|")
    print(f"| Outils Total:           {formatted_outils_sum.rjust(20)}{' ' * 20}|")
    print(border)


def find_nearest_facturation_client_dir(file_path):
    """
    Walks up the folder hierarchy from the directory of the given file
    to find the nearest /Facturation - Client/ directory.
    """
    current_dir = os.path.dirname(file_path)
    while current_dir != os.path.dirname(current_dir):  # Check until the root directory
        for item in os.listdir(current_dir):
            if item == "Facturation - Client" and os.path.isdir(os.path.join(current_dir, item)):
                return os.path.join(current_dir, item)
        current_dir = os.path.dirname(current_dir)
    return None

def find_earliest_invoice_date(directory):
    """
    Finds the date of the earliest invoice in the specified directory.
    """
    earliest_date = None
    date_pattern = r'\b\d{2}/\d{2}/\d{4}\b'  # Regex for date format DD/MM/YYYY

    for filename in os.listdir(directory):
        if filename.lower().endswith('.pdf'):
            path = os.path.join(directory, filename)
            try:
                with pdfplumber.open(path) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            matches = re.findall(date_pattern, text)
                            for match in matches:
                                date = datetime.strptime(match, "%d/%m/%Y")
                                if earliest_date is None or date < earliest_date:
                                    earliest_date = date
            except Exception as e:
                print(f"Erreur lors du traitement du fichier PDF {filename} : {e}")

    return earliest_date.strftime("%d/%m/%Y") if earliest_date else None


def process_excel_files():
    directory = input("Veuillez saisir le chemin du répertoire: ")
    if not os.path.isdir(directory):
        print("Le répertoire fourni n'existe pas. Veuillez entrer un répertoire valide.")
        return

    year_data = defaultdict(lambda: {'mobilisation_sum': 0, 'outils_sum': 0, 'excel_count': 0})

    for root, dirs, files in os.walk(directory):
        excel_files = [f for f in files if f.lower().endswith('_execute.xlsm')]
        
        for filename in excel_files:
            path = os.path.join(root, filename)
            # Find the nearest /Facturation - Client/ directory for each _EXECUTE file
            facturation_client_dir = find_nearest_facturation_client_dir(path)
            year = 2021
            if facturation_client_dir is None:
                year = 2021
                #print(f"Facturation - Client directory not found for file {filename}")
            else:
                earliest_invoice_date = find_earliest_invoice_date(facturation_client_dir)
                year = datetime.strptime(earliest_invoice_date, "%d/%m/%Y").year
                if year is None:
                    print("Erreur lors de la détermination de l'année pour le fichier {filename}")

            try:
                workbook = load_workbook(filename=path, data_only=True)
                sheet = workbook.active
                mobilisation_value = sheet['K53'].value
                outils_value = sheet['K52'].value

                if mobilisation_value is not None and isinstance(mobilisation_value, (int, float)):
                    year_data[year]['mobilisation_sum'] += mobilisation_value
                if outils_value is not None and isinstance(outils_value, (int, float)):
                    year_data[year]['outils_sum'] += outils_value

                year_data[year]['excel_count'] += 1

            except Exception as e:
                print(f"Erreur lors du traitement du fichier {filename} : {e}")
    
    for year, data in year_data.items():
        print(f"\nRésultats pour l'année {year}:")
        print_results(year, data['excel_count'], data['mobilisation_sum'], data['outils_sum'])

    return

def process_pdf_files():
    directory = input("Veuillez saisir le chemin du répertoire pour les fichiers PDF: ")
    if not os.path.isdir(directory):
        print("Le répertoire fourni n'existe pas. Veuillez entrer un répertoire valide.")
        return

    date_pattern = r'\b\d{2}/\d{2}/\d{4}\b'  # Regex for date format DD/MM/YYYY
    found_dates = []

    # Use os.walk to recursively traverse the directories
    for root, dirs, files in os.walk(directory):
        # Filter files to include only PDF files
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]

        for filename in pdf_files:
            path = os.path.join(root, filename)
            try:
                with pdfplumber.open(path) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            matches = re.findall(date_pattern, text)
                            found_dates.extend(matches)
            except Exception as e:
                print(f"Erreur lors du traitement du fichier PDF {filename} : {e}")

    print("Dates trouvées dans les fichiers PDF:")
    for date in set(found_dates):
        print(date)



if __name__ == "__main__":
    main_menu()
1