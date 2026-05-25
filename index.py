

import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook  

class Stagiaire:
    def __init__(self, matricule, nom, prenom, filiere, groupe, niveau, email, telephone):
        self.matricule = matricule
        self.nom = nom
        self.prenom = prenom
        self.filiere = filiere
        self.groupe = groupe
        self.niveau = niveau
        self.email = email
        self.telephone = telephone

    def afficher_infos(self):
        return f"{self.matricule} | {self.nom} {self.prenom} | {self.filiere} | {self.groupe} | {self.niveau} | {self.email} | {self.telephone}"



# Classe fille: Technicien (avec BAC)
class StagiaireTechnicien(Stagiaire):
    def __init__(self, matricule, nom, prenom, filiere, groupe, niveau, email, telephone, specialite_bac, annee_bac):
        super().__init__(matricule, nom, prenom, filiere, groupe, niveau, email, telephone)
        self.specialite_bac = specialite_bac
        self.annee_bac = annee_bac

    def afficher_infos(self):
        return super().afficher_infos() + f" | Spécialité BAC: {self.specialite_bac} | Année BAC: {self.annee_bac}"


#  Sans BAC
class StagiaireSansBac(Stagiaire):
    def __init__(self, matricule, nom, prenom, filiere, groupe, niveau, email, telephone, niveau_scolaire):
        super().__init__(matricule, nom, prenom, filiere, groupe, niveau, email, telephone)
        self.niveau_scolaire = niveau_scolaire

    def afficher_infos(self):
        return super().afficher_infos() + f" | Niveau scolaire: {self.niveau_scolaire}"


liste_stagiaires = []



# Tkinter Interface
fenetre = tk.Tk()
fenetre.title("Gestion des Stagiaires - Héritage OFPPT")
fenetre.geometry("650x650")


labels = [
    "Matricule", "Nom", "Prénom", "Filière",
    "Groupe", "Niveau", "Email", "Téléphone",
    "Spécialité BAC / Niveau scolaire"
]

entries = {}
for i, label in enumerate(labels):
    tk.Label(fenetre, text=label).grid(row=i, column=0, padx=5, pady=5, sticky="w")
    entry = tk.Entry(fenetre, width=35)
    entry.grid(row=i, column=1, padx=5, pady=5)
    entries[label.lower()] = entry

# Matricule à modifier / supprimer
tk.Label(fenetre, text="Matricule à modifier / supprimer").grid(row=len(labels), column=0, padx=5, pady=5, sticky="w")
entry_action = tk.Entry(fenetre, width=35)
entry_action.grid(row=len(labels), column=1, padx=5, pady=5)

# Text widget 
text_output = tk.Text(fenetre, width=80, height=15)
text_output.grid(row=len(labels)+2, column=0, columnspan=3, pady=10)

def afficher():
    text_output.delete("1.0", tk.END)
    for s in liste_stagiaires:
        text_output.insert(tk.END, s.afficher_infos() + "\n")


def ajouter():
    niveau = entries["niveau"].get().strip().lower()
    spec_input = entries["spécialité bac / niveau scolaire"].get().strip()
    if niveau == "technicien" or niveau == "1ère année":
        s = StagiaireTechnicien(
            entries["matricule"].get(),
            entries["nom"].get(),
            entries["prénom"].get(),
            entries["filière"].get(),
            entries["groupe"].get(),
            entries["niveau"].get(),
            entries["email"].get(),
            entries["téléphone"].get(),
            spec_input,  
            "2025"      
        )
        
    else:
        s = StagiaireSansBac(
            entries["matricule"].get(),
            entries["nom"].get(),
            entries["prénom"].get(),
            entries["filière"].get(),
            entries["groupe"].get(),
            entries["niveau"].get(),
            entries["email"].get(),
            entries["téléphone"].get(),
            spec_input  
        )
    liste_stagiaires.append(s)
    afficher()
    tk.messagebox.showinfo("Succès", "Stagiaire ajouté")


def supprimer():
    mat = entry_action.get().strip()
    for s in liste_stagiaires:
        if s.matricule == mat:
            liste_stagiaires.remove(s)
            afficher()
            tk.messagebox.showinfo("Succès", "Stagiaire supprimé")
            return
    tk.messagebox.showerror("Erreur", "Matricule introuvable")


def modifier_stagiaire():
    mat = entry_action.get().strip()
    for s in liste_stagiaires:
        if s.matricule == mat:
            if entries["nom"].get():
                s.nom = entries["nom"].get()
            if entries["prénom"].get():
                s.prenom = entries["prénom"].get()
            if entries["filière"].get():
                s.filiere = entries["filière"].get()
            if entries["groupe"].get():
                s.groupe = entries["groupe"].get()
            if entries["niveau"].get():
                s.niveau = entries["niveau"].get()
            if entries["email"].get():
                s.email = entries["email"].get()
            if entries["téléphone"].get():
                s.telephone = entries["téléphone"].get()
            if isinstance(s, StagiaireTechnicien):
                s.specialite_bac = entries["spécialité bac / niveau scolaire"].get()
            elif isinstance(s, StagiaireSansBac):
                s.niveau_scolaire = entries["spécialité bac / niveau scolaire"].get()
            afficher()
            tk.messagebox.showinfo("Succès", "Stagiaire modifié")
            return
    tk.messagebox.showerror("Erreur", "Matricule introuvable")



def exporter_excel():
    if not liste_stagiaires:
        tk.messagebox.showwarning("Attention", "Aucun stagiaire à exporter")
        return

    wb = Workbook()  
    ws = wb.active
    ws.title = "Stagiaires"

    
    headers = ["Matricule", "Nom", "Prénom", "Filière", "Groupe", "Niveau", "Email", "Téléphone", "Spécialité BAC / Niveau scolaire"]
    ws.append(headers)

   
    for s in liste_stagiaires:
        if isinstance(s, StagiaireTechnicien):
            spec = f"{s.specialite_bac} ({s.annee_bac})"
        elif isinstance(s, StagiaireSansBac):
            spec = s.niveau_scolaire
        else:
            spec = ""
        ws.append([s.matricule, s.nom, s.prenom, s.filiere, s.groupe, s.niveau, s.email, s.telephone, spec])

      
    fichier = "Stagiaires.xlsx"
    wb.save(fichier)
    tk.messagebox.showinfo("Succès", f"Les stagiaires ont été exportés dans {fichier}")


# Boutons
frame_buttons = tk.Frame(fenetre)
frame_buttons.grid(row=len(labels)+1, column=0, columnspan=2, pady=10)

tk.Button(frame_buttons, text="Ajouter", command=ajouter).pack(side="left", padx=5)
tk.Button(frame_buttons, text="Supprimer", command=supprimer).pack(side="left", padx=5)
tk.Button(frame_buttons, text="Modifier", command=modifier_stagiaire).pack(side="left", padx=5)
tk.Button(frame_buttons, text="Afficher", command=afficher).pack(side="left", padx=5)
tk.Button(frame_buttons, text="Exporter Excel", command=exporter_excel).pack(side="left", padx=5)  


fenetre.mainloop()
