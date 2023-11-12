# main.py

import warnings
import os
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from tqdm import tqdm  # Import the tqdm module

# Ignore specific UserWarning from openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet._reader")

def main_menu():
    # Get the directory of the current script
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Construct the full path for banner.txt
    banner_path = os.path.join(script_dir, 'bannerV2.txt')

    try:
        with open(banner_path, 'r') as banner_file:
            banner_contents = banner_file.read()
            print(banner_contents)
    except FileNotFoundError:
        print('''    _    _                 _       _                      _                  _   
   / \  | |_   _ _ __ ___ (_)_ __ (_)_   _ _ __ ___      / \   ___  ___ ___ | |_ 
  / _ \ | | | | | '_ ` _ \| | '_ \| | | | | '_ ` _ \    / _ \ / __|/ __/ _ \| __|
 / ___ \| | |_| | | | | | | | | | | | |_| | | | | | |  / ___ \\__ \ (_| (_) | |_ 
/_/   \_\_|\__,_|_| |_| |_|_|_| |_|_|\__,_|_| |_| |_| /_/   \_\___/\___\___/ \__|
                                                                                 
''')
    while True:
        print("\nMenu Principal")
        print("1. Analyse des fichiers Excel")
        print("2. Générer un document Excel avec les résultats")
        print("0. Quitter")
        
        choice = input("Entrez votre choix: ")
        
        if choice == "1":
            process_excel_files()
        elif choice == "2":
            generate_excel_report()
        elif choice == "0":
            print("Fermeture du programme.")
            break
        else:
            print("Choix invalide, veuillez réessayer.")

def print_results(excel_count, mobilisation_sum, outils_sum):
    border = "+" + "-" * 60 + "+"
    title = " RÉSULTATS DE L'ANALYSE "
    formatted_mobilisation_sum = "${:,.2f}".format(mobilisation_sum)  # With comma as thousand separator, 2 decimal places, and a dollar sign
    formatted_outils_sum = "${:,.2f}".format(outils_sum)

    print(border)
    print(f"|{title.center(60)}|")
    print(border)
    print(f"| Nombre Total de Projet: {str(excel_count).ljust(20)}{' ' * 20}|")
    print(f"| Mobilisation Total:     {formatted_mobilisation_sum.rjust(20)}{' ' * 20}|")
    print(f"| Outils Total:           {formatted_outils_sum.rjust(20)}{' ' * 20}|")
    print(border)

def generate_excel_report():
    excel_count, mobilisation_sum, outils_sum = process_excel_files()

    wb = Workbook()
    ws = wb.active
    ws.title = "Résultats"

    # Define styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    row_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

    # Add headers
    headers = ["Total des Projets", "Mobilisation Total", "Outils Total"]
    ws.append(headers)
    
    # Apply styles to headers
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Add data
    data = [excel_count, mobilisation_sum, outils_sum]
    ws.append(data)

    # Style data rows with alternating colors
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")
            if cell.row % 2 == 0:
                cell.fill = row_fill

    # Set column width
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        adjusted_width = (max_length + 2)
        ws.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width

    output_filename = "Rapport_Analyse.xlsx"
    wb.save(output_filename)
    print(f"Le rapport a été généré: {output_filename}")

def process_excel_files():
    directory = input("Veuillez saisir le chemin du répertoire: ")
    if not os.path.isdir(directory):
        print("Le répertoire fourni n'existe pas. Veuillez entrer un répertoire valide.")
        return

    mobilisation_sum = 0
    outils_sum = 0 
    excel_count = 0

    # Use os.walk to recursively traverse the directories
    for root, dirs, files in os.walk(directory):
        # Filter files to include only those ending with '_EXECUTE.xlsm'
        excel_files = [f for f in files if f.endswith('_EXECUTE.xlsm')]

        # Use tqdm to show the progress bar, iterating over the list of Excel files
        for filename in tqdm(excel_files, desc="Processing Excel files"):
            excel_count += 1
            path = os.path.join(root, filename)
            try:
                workbook = load_workbook(filename=path, data_only=True)
                sheet = workbook.active
                mobilisation_value = sheet['K53'].value
                outils_value = sheet['K52'].value

                if mobilisation_value is not None and isinstance(mobilisation_value, (int, float)):
                    mobilisation_sum += mobilisation_value
                if outils_value is not None and isinstance(outils_value, (int, float)):
                    outils_sum += outils_value

            except Exception as e:
                print(f"Erreur lors du traitement du fichier {filename} : {e}")
    
    print_results(excel_count, mobilisation_sum, outils_sum)

    return excel_count, mobilisation_sum, outils_sum



if __name__ == "__main__":
    main_menu()
1