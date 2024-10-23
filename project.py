import pytesseract
import cv2
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from PIL import Image, ImageTk
import re
import pandas as pd
import subprocess

# Définir la fonction pour extraire le texte de l'image
def extract_text_from_image(image_path):
    # Charger l'image avec OpenCV
    image = cv2.imread(image_path)
    
    # Utiliser pytesseract pour extraire le texte de l'image
    text = pytesseract.image_to_string(image)
    
    return text

# Définir la fonction pour extraire et enregistrer les données dans un fichier Excel
def conv_im_to_ex(img):
    img = Image.open(img)
    text = pytesseract.image_to_string(img)

    # Split the text into lines
    lines = text.split('\n')

    # Initialize lists to store extracted data
    data = []
    for line in lines:
        if re.match(r"\d{2}/\d{2}", line):  # Check if the line starts with a date-like pattern
            parts = line.split()
            if len(parts) >= 3:
                date = parts[0]
                operation = " ".join(parts[1:-2])
                amount = parts[-1].replace(",", ".")
                if "Débit" in operation:
                    debit = amount
                    credit = "0.00"
                else:
                    debit = "0.00"
                    credit = amount
                data.append([date, operation, debit, credit])

    # Create a DataFrame from the extracted data
    columns = ["Date", "Opérations", "Débit", "Crédit"]
    new_df = pd.DataFrame(data, columns=columns)

    # Read the existing Excel file into a DataFrame (if it exists)
    try:
        existing_df = pd.read_excel("dataset.xlsx")
    except FileNotFoundError:
        existing_df = pd.DataFrame(columns=columns)

    # Concatenate the new data with the existing DataFrame
    combined_df = pd.concat([existing_df, new_df], ignore_index=True)

    # Save the combined DataFrame to the Excel file
    combined_df.to_excel("dataset.xlsx", index=False)
    print("Données ajoutées au fichier Excel avec succès.")

    # Ouvrir le fichier Excel avec l'application par défaut
    try:
        subprocess.Popen(["start", "dataset.xlsx"], shell=True)
    except Exception as e:
        print("Erreur lors de l'ouverture du fichier Excel :", e)

def browse_file():
    file_path = filedialog.askopenfilename()
    file_label.config(text="Selected File: " + file_path)
    display_image(file_path)

def display_image(image_path):
    global img  # Utilisation de la variable img en tant que variable globale
    img = Image.open(image_path)
    img.thumbnail((200, 200))
    img = ImageTk.PhotoImage(img)
    image_label.config(image=img)
    image_label.image = img

def confirm_upload():
    selected_file = file_label.cget("text")[14:].strip()  # Supprimer les espaces autour du nom du fichier
    file_type = file_type_var.get()
    print("File Type:", file_type)
    print("Selected File:", selected_file)
    
    # Extraction du texte de l'image sélectionnée
    if selected_file:
        text = extract_text_from_image(selected_file)
        print("Text from Image:")
        print(text)
        
        # Appeler la fonction conv_im_to_ex pour traiter l'image et l'enregistrer dans Excel
        conv_im_to_ex(selected_file)

root = tk.Tk()
root.title("File Upload Interface")

# Couleurs personnalisées
bg_color = "#6E6E6E"
button_bg_color = "#BAFF39"
button_fg_color = "white"
text_color = "white"
frame_bg_color = "#B0B0B0"  # Couleur de fond du cadre

root.configure(bg=bg_color)

file_type_var = tk.StringVar()

font = ("Helvetica", 12, "bold")

# Créer un cadre principal
main_frame = tk.Frame(root, bg=bg_color)
main_frame.pack(padx=20, pady=20)

# Créer un cadre pour regrouper les éléments liés au fichier
file_frame = tk.Frame(main_frame, bg=frame_bg_color, relief=tk.RAISED, borderwidth=2)
file_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)  # Augmenter le padding ici

# Ajouter un coin arrondi en utilisant un relief circulaire
file_frame.bind("<Enter>", lambda event: file_frame.config(relief=tk.SUNKEN))
file_frame.bind("<Leave>", lambda event: file_frame.config(relief=tk.RAISED))

file_type_label = tk.Label(file_frame, text="Select file type:", bg=frame_bg_color, fg=text_color, font=font)
file_type_label.pack(pady=(10, 0))

file_type_combobox = ttk.Combobox(file_frame, textvariable=file_type_var, values=["CV", "Bank Statement"], font=font)
file_type_combobox.pack(pady=(0, 10))

browse_button = tk.Button(file_frame, text="Browse", command=browse_file, bg=button_bg_color, fg=button_fg_color, font=font)
browse_button.pack()

file_label = tk.Label(file_frame, text="Selected File: None", bg=frame_bg_color, fg=text_color, font=font)
file_label.pack(pady=(10, 0))

image_label = tk.Label(main_frame, bg=bg_color)  # Déplacer l'image_label en dehors de la fonction display_image
image_label.pack()

confirm_button = tk.Button(main_frame, text="Confirm Upload", command=confirm_upload, bg=button_bg_color, fg=button_fg_color, font=font)
confirm_button.pack(pady=(10, 0))

root.geometry("400x500")
root.mainloop()
