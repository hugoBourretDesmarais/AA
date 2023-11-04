# main.py

import os
from openpyxl import load_workbook

def main_menu():
    while True:
        print("\nMenu Principal")
        print("1. Analyse des fichiers Excel")
        print("0. Quitter")
        
        choice = input("Entrez votre choix: ")
        
        if choice == "1":
            process_excel_files()
        elif choice == "0":
            print("Fermeture du programme.")
            break
        else:
            print("Choix invalide, veuillez réessayer.")

def process_excel_files():
    directory = input("Please enter the directory path: ")
    if not os.path.isdir(directory):
        print("Le répertoire fourni n'existe pas. Veuillez entrer un répertoire valide.")
        return

    mobilisation_sum = 0
    outils_sum = 0 
    excel_count = 0

    for filename in os.listdir(directory):
        if filename.endswith('.xlsm'):
            excel_count += 1
            path = os.path.join(directory, filename)
            try:
                workbook = load_workbook(filename=path, data_only=True)
                sheet = workbook.active
                mobilisation_value = sheet['K53'].value
                outils_value = sheet['K52'].value  # New line to get the value of "Outils"

                if mobilisation_value is not None and isinstance(mobilisation_value, (int, float)):
                    mobilisation_sum += mobilisation_value
                if outils_value is not None and isinstance(outils_value, (int, float)):
                    outils_sum += outils_value  # Adding the value of "Outils" to its sum

            except Exception as e:
                print(f"Erreur lors du traitement du fichier {filename} : {e}")
    
    print(f"Nombre Total de Projet: {excel_count}")
    print(f"Mobilisation Total: {mobilisation_sum}")
    print(f"Outils Total: {outils_sum}")


if __name__ == "__main__":
    main_menu()
1