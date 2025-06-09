#!/usr/bin/env python3
# -*- coding: utf-8 -*-


# envejp.py
# Copyright (c) 2025 Stéphane BDC
#
# This software is released under the MIT License.
# https://opensource.org/licenses/MIT
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.



import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
import os
import tempfile
import datetime
import shutil
import re
import sys

def resource_path(relative_path):
    """Obtient le chemin vers les ressources, compatible PyInstaller"""
    try:
        # PyInstaller crée un dossier temporaire et y stocke le chemin dans _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Mode développement - utilise le répertoire du script
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

class EnvelopeGenerator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Générateur d'enveloppes japonaises")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Variables
        self.pdf_temp_path = None
        self.generated_filename = None
        
        # Initialiser la police japonaise
        self.setup_japanese_font()
        
        # Créer l'interface
        self.create_interface()
        
    def setup_japanese_font(self):
        """Configure la police japonaise depuis le dossier fonts/"""
        try:
            # Utiliser resource_path pour obtenir le bon chemin
            font_dir = resource_path("fonts")
            
            print(f"Font directory: {font_dir}")
            print(f"Font directory exists: {os.path.exists(font_dir)}")
            
            available_fonts = []
            if os.path.exists(font_dir):
                files_in_dir = os.listdir(font_dir)
                print(f"Files in fonts directory: {files_in_dir}")
                
                for font_file in files_in_dir:
                    if font_file.endswith(('.ttf', '.ttc')):  # Exclure .otf qui pose problème
                        font_path = os.path.join(font_dir, font_file)
                        print(f"Testing font: {font_path}")
                        
                        # Tester si la police peut être chargée
                        try:
                            # Test avec un nom temporaire
                            test_name = f"TestFont_{len(available_fonts)}"
                            pdfmetrics.registerFont(TTFont(test_name, font_path))
                            available_fonts.append({
                                'name': font_file,
                                'path': font_path,
                                'display_name': font_file.replace('.ttf', '').replace('.ttc', '')
                            })
                            print(f"Successfully tested font: {font_file}")
                        except Exception as font_error:
                            print(f"Failed to load {font_file}: {font_error}")
                            continue
            
            font_found = None
            selected_font = None
            
            if len(available_fonts) == 0:
                print("No valid fonts found in fonts directory")
                # Essayer les polices spécifiques que vous avez mentionnées
                specific_fonts = ["SawarabiMincho-Regular.ttf", "NotoSansJP-VariableFont_wght.ttf"]
                for font_name in specific_fonts:
                    font_path = resource_path(os.path.join("fonts", font_name))
                    if os.path.exists(font_path):
                        try:
                            test_name = f"TestFont_{len(available_fonts)}"
                            pdfmetrics.registerFont(TTFont(test_name, font_path))
                            available_fonts.append({
                                'name': font_name,
                                'path': font_path,
                                'display_name': font_name.replace('.ttf', '')
                            })
                            print(f"Found specific font: {font_name}")
                        except Exception as e:
                            print(f"Failed to load specific font {font_name}: {e}")
            
            if len(available_fonts) == 1:
                # Une seule police, l'utiliser directement
                selected_font = available_fonts[0]
                print(f"Single font found: {selected_font['name']}")
            elif len(available_fonts) > 1:
                # Plusieurs polices, demander à l'utilisateur de choisir
                selected_font = self.show_font_selection_dialog(available_fonts)
            
            if selected_font:
                try:
                    pdfmetrics.registerFont(TTFont("JapaneseFont", selected_font['path']))
                    font_found = selected_font['path']
                    self.selected_font_name = selected_font['display_name']
                    print(f"Successfully loaded selected font: {selected_font['name']}")
                except Exception as font_error:
                    print(f"Failed to load selected font: {font_error}")
            
            # Si aucune police trouvée dans fonts/, chercher les polices système
            if not font_found:
                print("Searching system fonts...")
                possible_fonts = [
                    "NotoSansJP-VariableFont_wght.ttf",
                    "NotoSansJP-Regular.ttf",
                    "NotoSansCJK-Regular.ttc",
                    "fonts-japanese-gothic.ttf"
                ]
                
                system_paths = [
                    "/System/Library/Fonts/",
                    "/usr/share/fonts/",
                    "C:/Windows/Fonts/",
                    os.path.expanduser("~/Library/Fonts/")
                ]
                
                for sys_path in system_paths:
                    if os.path.exists(sys_path):
                        for font_file in possible_fonts:
                            full_path = os.path.join(sys_path, font_file)
                            if os.path.exists(full_path):
                                try:
                                    pdfmetrics.registerFont(TTFont("JapaneseFont", full_path))
                                    font_found = full_path
                                    self.selected_font_name = font_file
                                    print(f"Successfully loaded system font: {full_path}")
                                    break
                                except Exception as font_error:
                                    print(f"Failed to load system font {full_path}: {font_error}")
                                    continue
                    if font_found:
                        break
            
            if font_found:
                self.font_available = True
                self.font_path = font_found
                print(f"Final font selected: {font_found}")
            else:
                self.font_available = False
                self.selected_font_name = "Aucune"
                print("No Japanese font could be loaded")
                
        except Exception as e:
            self.font_available = False
            self.selected_font_name = "Erreur"
            print(f"Error during font setup: {e}")

    def show_font_selection_dialog(self, available_fonts):
        """Affiche une popup pour choisir la police à utiliser"""
        selection_window = tk.Toplevel(self.root)
        selection_window.title("Choix de la police")
        selection_window.geometry("400x300")
        selection_window.resizable(False, False)
        
        # Centrer la fenêtre
        selection_window.transient(self.root)
        selection_window.grab_set()
        
        selected_font = None
        
        # Titre
        title_label = tk.Label(selection_window, text="Plusieurs polices japonaises trouvées", 
                              font=("Arial", 12, "bold"))
        title_label.pack(pady=10)
        
        # Sous-titre
        subtitle_label = tk.Label(selection_window, text="Veuillez choisir la police à utiliser :")
        subtitle_label.pack(pady=5)
        
        # Frame pour la liste
        list_frame = tk.Frame(selection_window)
        list_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Listbox avec scrollbar
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, 
                            font=("Arial", 10), height=8)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Ajouter les polices à la liste
        for font in available_fonts:
            listbox.insert(tk.END, font['display_name'])
        
        # Sélectionner le premier élément par défaut
        listbox.select_set(0)
        
        # Preview du nom de fichier
        preview_frame = tk.Frame(selection_window)
        preview_frame.pack(pady=5)
        
        preview_label = tk.Label(preview_frame, text="Fichier: ", font=("Arial", 9))
        preview_label.pack(side="left")
        
        preview_filename = tk.Label(preview_frame, text=available_fonts[0]['name'], 
                                   font=("Arial", 9, "italic"), fg="gray")
        preview_filename.pack(side="left")
        
        def update_preview(event):
            selection = listbox.curselection()
            if selection:
                selected_index = selection[0]
                preview_filename.config(text=available_fonts[selected_index]['name'])
        
        listbox.bind('<<ListboxSelect>>', update_preview)
        
        # Frame pour les boutons
        button_frame = tk.Frame(selection_window)
        button_frame.pack(pady=15)
        
        def on_ok():
            nonlocal selected_font
            selection = listbox.curselection()
            if selection:
                selected_index = selection[0]
                selected_font = available_fonts[selected_index]
            else:
                selected_font = available_fonts[0]  # Fallback
            selection_window.destroy()
        
        def on_cancel():
            nonlocal selected_font
            selected_font = available_fonts[0]  # Utiliser la première par défaut
            selection_window.destroy()
        
        # Boutons
        ok_btn = tk.Button(button_frame, text="Utiliser cette police", 
                          command=on_ok, bg="#4CAF50", fg="grey",
                          font=("Arial", 10, "bold"), padx=20, pady=5)
        ok_btn.pack(side="left", padx=5)
        
        cancel_btn = tk.Button(button_frame, text="Première par défaut", 
                              command=on_cancel, bg="#6b6b6b", fg="grey",
                              font=("Arial", 10), padx=20, pady=5)
        cancel_btn.pack(side="left", padx=5)
        
        # Permettre la validation avec Entrée
        def on_enter(event):
            on_ok()
        
        selection_window.bind('<Return>', on_enter)
        listbox.focus_set()
        
        # Attendre que la fenêtre soit fermée
        self.root.wait_window(selection_window)
        
        return selected_font
    
    def create_interface(self):
        """Crée l'interface utilisateur"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configuration du redimensionnement
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Titre
        title_label = ttk.Label(main_frame, text="Générateur d'enveloppes japonaises", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Code postal
        ttk.Label(main_frame, text="Code postal (ex: 〒160ｰ0007):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.postal_code_var = tk.StringVar(value="〒160ｰ0007")
        postal_entry = ttk.Entry(main_frame, textvariable=self.postal_code_var, width=20)
        postal_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Adresse ligne 1
        ttk.Label(main_frame, text="Adresse ligne 1:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.address1_var = tk.StringVar(value="東京都新宿区 荒木町11-1")
        address1_entry = ttk.Entry(main_frame, textvariable=self.address1_var, width=50)
        address1_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Adresse ligne 2
        ttk.Label(main_frame, text="Adresse ligne 2:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.address2_var = tk.StringVar(value="ハイム石川8号")
        address2_entry = ttk.Entry(main_frame, textvariable=self.address2_var, width=50)
        address2_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Nom de l'entreprise
        ttk.Label(main_frame, text="Nom de l'entreprise:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.company_var = tk.StringVar(value="ステファン ビーディーシーLTD.")
        company_entry = ttk.Entry(main_frame, textvariable=self.company_var, width=50)
        company_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Destinataire
        ttk.Label(main_frame, text="Destinataire:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.recipient_var = tk.StringVar(value="経理・藤原様")
        recipient_entry = ttk.Entry(main_frame, textvariable=self.recipient_var, width=50)
        recipient_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=5)
        
        
        # Frame pour les boutons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        # Bouton vérifier avec style explicite
        verify_btn = tk.Button(button_frame, text="Vérifier l'adresse", 
                              command=self.verify_address,
                              bg="#6b6b6b", fg="grey", 
                              font=("Arial", 10, "bold"),
                              relief="raised", bd=2,
                              padx=10, pady=5)
        verify_btn.pack(side=tk.LEFT, padx=5)
        
        # Bouton générer PDF
        self.generate_btn = tk.Button(button_frame, text="Générer PDF", 
                                     command=self.generate_pdf, state="disabled",
                                     bg="#6b6b6b", fg="grey",
                                     font=("Arial", 10, "bold"),
                                     relief="raised", bd=2,
                                     padx=10, pady=5,
                                     disabledforeground="gray")
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        
        # Bouton sauvegarder (initialement caché)
        self.save_btn = tk.Button(button_frame, text="Sauvegarder le PDF", 
                                 command=self.save_pdf,
                                 bg="#6b6b6b", fg="grey",
                                 font=("Arial", 10, "bold"),
                                 relief="raised", bd=2,
                                 padx=10, pady=5)
        
        # Zone d'état
        self.status_text = tk.Text(main_frame, height=8, width=70)
        self.status_text.grid(row=7, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

        # Scrollbar pour la zone d'état
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=7, column=2, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)

        # Message initial
        self.add_status("Interface initialisée.")
        if not self.font_available:
            self.add_status("⚠️ Aucune police japonaise trouvée. Vérifiez la console pour plus de détails.")
        else:
            self.add_status(f"✅ Police japonaise chargée: {self.selected_font_name}")

    def add_status(self, message):
        """Ajoute un message dans la zone d'état"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update()

    def verify_address(self):
        """Vérifie la validité de l'adresse saisie"""
        self.add_status("Vérification de l'adresse...")

        # Récupérer les valeurs
        postal = self.postal_code_var.get().strip()
        addr1 = self.address1_var.get().strip()
        addr2 = self.address2_var.get().strip()
        company = self.company_var.get().strip()
        recipient = self.recipient_var.get().strip()

        errors = []
        warnings = []

        # Vérifications
        if not postal:
            errors.append("Le code postal est obligatoire")
        elif not postal.startswith('〒'):
            warnings.append("Le code postal devrait commencer par 〒")

        if not addr1:
            errors.append("L'adresse ligne 1 est obligatoire")
        elif len(addr1) > 40:
            warnings.append(f"Adresse ligne 1 très longue ({len(addr1)} caractères)")

        if not company:
            errors.append("Le nom de l'entreprise est obligatoire")

        if not recipient:
            errors.append("Le destinataire est obligatoire")

        # Vérifier la longueur totale pour chaque ligne
        lines = [addr1, addr2, company, recipient]
        for i, line in enumerate(lines, 1):
            if len(line) > 35:
                warnings.append(f"Ligne {i} risque de déborder ({len(line)} caractères)")

        # Afficher les résultats
        if errors:
            for error in errors:
                self.add_status(f"❌ Erreur: {error}")
            self.generate_btn.configure(state="disabled")
        else:
            self.add_status("✅ Adresse valide")
            if warnings:
                for warning in warnings:
                    self.add_status(f"⚠️ Attention: {warning}")
            self.generate_btn.configure(state="normal")

    def generate_pdf(self):
        """Génère le PDF de l'enveloppe"""
        if not self.font_available:
            messagebox.showerror("Erreur", "Aucune police japonaise disponible")
            return

        self.add_status("Génération du PDF en cours...")

        try:
            # Créer un nom de fichier avec timestamp et portion d'adresse
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            company_clean = re.sub(r'[^a-zA-Z0-9\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]', '', 
                                 self.company_var.get())[:10]
            self.generated_filename = f"envelope_{timestamp}_{company_clean}.pdf"
    
            # Créer un fichier temporaire
            temp_dir = tempfile.gettempdir()
            self.pdf_temp_path = os.path.join(temp_dir, self.generated_filename)
    
            # Générer le PDF
            self.create_pdf(self.pdf_temp_path)
    
            self.add_status(f"✅ PDF généré: {self.generated_filename}")
    
            # Afficher le bouton de sauvegarde
            self.save_btn.pack(side=tk.LEFT, padx=5)
    
        except Exception as e:
            self.add_status(f"❌ Erreur lors de la génération: {str(e)}")
            messagebox.showerror("Erreur", f"Impossible de générer le PDF:\n{str(e)}")

    def create_pdf(self, filepath):
        """Crée le PDF avec l'adresse saisie"""
        # Création du PDF avec taille A5
        c = canvas.Canvas(filepath, pagesize=(A5[1], A5[0]))
        c.setFont("JapaneseFont", 14)

        # Dimensions de la page
        page_width = A5[1]
        page_height = A5[0]

        # Récupérer les données du formulaire
        lines = []
        if self.address1_var.get().strip():
            lines.append(self.address1_var.get().strip())
        if self.address2_var.get().strip():
            lines.append(self.address2_var.get().strip())
        if self.company_var.get().strip():
            lines.append(self.company_var.get().strip())
        if self.recipient_var.get().strip():
            lines.append(self.recipient_var.get().strip())

        postal_code = self.postal_code_var.get().strip()

        # Paramètres de mise en page
        x_start = 150 * mm
        y_start = 140 * mm
        line_spacing = 12 * mm
        char_spacing = 20
        margin_bottom = 10 * mm
        margin_top = 10 * mm

        # Ajustement automatique de la position
        y_start = self.adjust_start_position(lines, postal_code, y_start, char_spacing, 
                                           page_height, margin_bottom)

        # Dessiner le code postal
        x_postal = x_start + line_spacing
        y_postal = y_start

        for i, char in enumerate(postal_code):
            if char == '〒':
                char_width = c.stringWidth(char, "JapaneseFont", 14)
                c.drawString(x_postal - char_width/4, y_postal, char)
            else:
                c.drawString(x_postal, y_postal, char)
            y_postal -= char_spacing

        # Dessiner les autres lignes
        for i, line in enumerate(lines):
            x = x_start - (i * line_spacing)
            y_offset = (i + 1) * 10 * mm
            y = y_start - y_offset
    
            for j, char in enumerate(line):
                current_y = y - (j * char_spacing)
                if current_y < margin_bottom:
                    break
                c.drawString(x, current_y, char)

        c.save()

    def adjust_start_position(self, lines, postal_code, y_start, char_spacing, 
                            page_height, margin_bottom):
        """Ajuste la position de départ pour éviter le débordement"""
        max_height = len(postal_code) * char_spacing

        for i, line in enumerate(lines):
            line_height = len(line) * char_spacing + (i + 1) * 10 * mm
            max_height = max(max_height, line_height)

        margin_top = 10 * mm
        available_height = page_height - margin_bottom - margin_top

        if y_start - max_height < margin_bottom:
            new_y_start = min(y_start, page_height - margin_top - max_height + y_start)
            return new_y_start

        return y_start

    def save_pdf(self):
        """Permet à l'utilisateur de sauvegarder le PDF"""
        if not self.pdf_temp_path or not os.path.exists(self.pdf_temp_path):
            messagebox.showerror("Erreur", "Aucun PDF à sauvegarder")
            return

        # Demander où sauvegarder
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=self.generated_filename
        )

        if save_path:
            try:
                shutil.copy2(self.pdf_temp_path, save_path)
                self.add_status(f"✅ PDF sauvegardé: {save_path}")
                messagebox.showinfo("Succès", f"PDF sauvegardé avec succès:\n{save_path}")
            except Exception as e:
                self.add_status(f"❌ Erreur de sauvegarde: {str(e)}")
                messagebox.showerror("Erreur", f"Impossible de sauvegarder:\n{str(e)}")

    def run(self):
        """Lance l'application"""
        self.root.mainloop()

        # Nettoyer le fichier temporaire à la fermeture
        if self.pdf_temp_path and os.path.exists(self.pdf_temp_path):
            try:
                os.remove(self.pdf_temp_path)
            except:
                pass

if __name__ == "__main__":
    app = EnvelopeGenerator()
    app.run()