#Kurama-90 ( https://github.com/Kurama-90 )

import easyocr
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import os
import cv2  # Importation de OpenCV
import numpy as np  # Importation de NumPy
import pandas as pd  # Pour créer un fichier Excel

# Chemins des dossiers
UPLOADS_DIR = 'uploads'
OUTPUTS_DIR = 'outputs'

def extract_text_from_image(image):
    print("Initialisation de EasyOCR...")
    reader = easyocr.Reader(['fr'])  # Initialiser EasyOCR pour le français
    print("Extraction du texte...")
    result = reader.readtext(image)
    print(f"Texte extrait : {len(result)} éléments.")
    return result

def add_text_layer_to_pdf(pdf_path, output_pdf_path, detections, image_size, offset_y):
    # Ouvrir le PDF original
    pdf_document = fitz.open(pdf_path)
    
    # Parcourir chaque page du PDF
    for page_num in range(len(pdf_document)):
        print(f"Traitement de la page {page_num + 1}...")
        page = pdf_document[page_num]
        
        # Taille de la page PDF en points
        page_width = page.rect.width
        page_height = page.rect.height
        
        # Taille de l'image en pixels
        image_width, image_height = image_size
        
        # Facteur de conversion des coordonnées
        x_scale = page_width / image_width
        y_scale = page_height / image_height
        
        # Ajustement vertical (offset_y)
        offset_y = 7.5  # Ajustez cette valeur si nécessaire
        
        # Ajouter le texte sélectionnable pour chaque détection
        for (bbox, text, prob) in detections[page_num]:
            # Coordonnées de la boîte englobante en pixels
            x0, y0 = bbox[0]  # Coin supérieur gauche
            x1, y1 = bbox[2]  # Coin inférieur droit
            
            # Convertir les coordonnées de pixels en points
            x0_pt = x0 * x_scale
            y0_pt = y0 * y_scale + offset_y  # Ajustement vertical
            x1_pt = x1 * x_scale
            y1_pt = y1 * y_scale + offset_y  # Ajustement vertical
            
            # Calculer la taille de la police en fonction de la hauteur de la boîte englobante
            bbox_height_pixels = y1 - y0  # Hauteur de la boîte englobante en pixels
            fontsize = bbox_height_pixels * y_scale  # Convertir en points
            
            # Ajouter le texte sélectionnable en noir
            page.insert_text((x0_pt, y0_pt), text, fontsize=fontsize, fontname="helv", color=(0, 1, 0))  # Couleur noire
    
    # Sauvegarder le PDF modifié
    pdf_document.save(output_pdf_path)
    pdf_document.close()
    print(f"PDF modifié sauvegardé : {output_pdf_path}")

def export_to_excel(detections, output_excel_path):
    # Créer un DataFrame pour stocker les données
    data = {}

    # Parcourir les détections et les organiser par position
    for page_num, page_detections in enumerate(detections):
        for (bbox, text, prob) in page_detections:
            x0, y0 = bbox[0]  # Coin supérieur gauche
            x1, y1 = bbox[2]  # Coin inférieur droit
            
            # Convertir les coordonnées en indices de cellule Excel
            row = int(y0 // 20)  # Ajustez la hauteur de la cellule si nécessaire
            col = int(x0 // 20)  # Ajustez la largeur de la cellule si nécessaire
            
            # Ajouter le texte à la bonne cellule
            if (row, col) not in data:
                data[(row, col)] = text
            else:
                data[(row, col)] += " " + text  # Concaténer si plusieurs textes dans la même cellule
    
    # Convertir en DataFrame
    df = pd.DataFrame.from_dict(data, orient="index", columns=["Text"])
    df.index = pd.MultiIndex.from_tuples(df.index, names=["Row", "Column"])
    df = df.unstack().fillna("")  # Organiser en tableau 2D
    
    # Exporter en fichier Excel
    df.to_excel(output_excel_path)
    print(f"Fichier Excel généré : {output_excel_path}")

def process_pdf(pdf_path, output_pdf_path, output_excel_path, offset_y):
    # Convertir le PDF en images
    print(f"Conversion du PDF en images : {pdf_path}")
    images = convert_from_path(pdf_path)
    print(f"{len(images)} pages converties en images.")
    
    # Extraire le texte de chaque image
    detections = []
    for i, image in enumerate(images):
        print(f"Extraction du texte de la page {i + 1}...")
        image_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)  # Conversion en format OpenCV
        detections.append(extract_text_from_image(image_cv))
    
    # Taille de l'image (en pixels)
    image_size = images[0].size  # Taille de la première image (supposons que toutes les images ont la même taille)
    
    # Ajouter une couche de texte sélectionnable au PDF original
    add_text_layer_to_pdf(pdf_path, output_pdf_path, detections, image_size, offset_y)
    
    # Exporter les données dans un fichier Excel
    export_to_excel(detections, output_excel_path)

if __name__ == '__main__':
    # Chemin du fichier PDF scanné dans le dossier uploads
    pdf_file = os.path.join(UPLOADS_DIR, 'scanned_document.pdf')  # Remplacez par le nom de votre fichier PDF
    
    # Chemin du fichier PDF de sortie dans le dossier outputs
    output_pdf = os.path.join(OUTPUTS_DIR, 'output.pdf')
    
    # Chemin du fichier Excel de sortie dans le dossier outputs
    output_excel = os.path.join(OUTPUTS_DIR, 'output.xlsx')
    
    # Vérifier si le dossier outputs existe, sinon le créer
    if not os.path.exists(OUTPUTS_DIR):
        print(f"Création du dossier {OUTPUTS_DIR}...")
        os.makedirs(OUTPUTS_DIR)
    
    # Traiter le PDF et générer les résultats
    process_pdf(pdf_file, output_pdf, output_excel, offset_y=7.5)
    print(f"Le PDF et le fichier Excel ont été générés avec succès.")
