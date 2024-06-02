import ttkbootstrap as ttk
import tkinter as tk
from tkinter import filedialog
from ttkbootstrap.style import Style
import os, threading
from threading import Event
import pandas as pd
from datetime import datetime


class Glob_var():
    def __init__(self):

        self.file_path = ""
        self.nom_fichier = ""
        self.var = 1
        self.thread = None
        self.run = True
        self.fichier_app = ""
        self.test = True
        self.download = False
        self.event = Event()
        self.save_path = ""
        self.entry_clicked = False
        self.leave = False
        
gb = Glob_var()

def func_recap():

    label_x.place(relx = 0.005, rely = 0.69, anchor = 'nw')
    label_recap.place(relx = 0.005, rely=0.2, anchor='nw')
    recap.set(f"\nVous avez selectionné le fichier : {os.path.basename(gb.file_path)}, et le filtre suivant : {cb.get()}.")
    check_validation.place(relx = 0.28, rely = 0.72, anchor='nw')
    
def remove_highlight(event):

    if gb.var == 2:
        event.widget.selection_clear()
        cb.config(foreground="black")

    if cb.get() != '...':
        info.set(f"Vous avez selectionné l'option '{cb.get()}'")
        bouton_lb3.state(['!disabled'])
        bouton_lb3.place(relx = 0.999, rely= 0.985, anchor="se")
        sep_type.place(relx = 0.305, rely = 0.1, anchor='nw', height=75)
        
        if cb.get() == "TOPO":
            pre_type.set("Le fichier TOPO possède 13 colonnes, dont voici les libellés dans l'ordre :")
            type_str.set("""
'code topo', 'libelle', 'type commune actuel (R ou N)', 'type commune FIP (RouNFIP)',  
'RUR actuel', 'RUR FIP',  'caractere voie', 'annulation', 'date annulation',  
'date création de article', 'type voie', 'mot classant', 'date derniere transition'.""")
            #Obligé de replacer pour pouvoir modifier la hauteur
            sep_type.place(relx = 0.305, rely = 0.1, anchor='nw', height=60)
            label_pre_type.lift()
            label_info.config(font = ('Arial', 9))
           
        elif cb.get() == 'ACHEMINEMENT' :
            pre_type.set("Le fichier ACHEMINEMENT possède 15 colonnes, dont voici les libellés dans l'ordre :")
            type_str.set(""" 
'code topographique', 'date limite de validite', 'code parite', 'borne superieure',
'borne inferieure', 'date effet', 'code postal', 'libelle de la donnee postale','
indicateur de pluridistribution', 'code localite', 'libelle ligne 5', 'code distribution',
'code type adresse', 'type code postal', 'code ancienne commune'""") 
            label_info.config(font = ('Arial', 8))
        
            label_pre_type.lift()
        elif cb.get() == 'STRUCTURES' :
            pre_type.set("Le fichier STRUCTURES possède 45 colonnes, dont voici un aperçu des libellés dans l'ordre :")
            type_str.set("""
'numero UA', 'date limite validite1', 'code UA', 'date limite validite2', 
'code codique', 'code type structure', 'filler1', 'date effet', 'date previsionnelle', 
...
'codes type liens fonctionnels', 'filler2'""")
            label_pre_type.lift()
            label_info.config(font = ('Arial', 9))

        elif cb.get() =='UAMISSIONS':
            pre_type.set("Le fichier UAMISSIONS possède 14 colonnes, dont voici les libellés dans l'ordre :")
            type_str.set("""
'numero UA', 'date fin mission exercee1', 'code mission', 'date fin mission exercee2', 'code UA', 
'date fin mission exercee3', 'code codique', 'code type de structure', 'date effet mission exercee', 
'libelle court mission', 'indicateur mission heritee', 'indicateur mission obligatoire', 
'indicateur mission competence geographique'""")
            label_pre_type.lift()
            label_info.config(font = ('Arial', 9))
            

        elif cb.get() =='COMPETENCES':
            pre_type.set("Le fichier COMPETENCES possède 16 colonnes, dont voici les libellés dans l'ordre :")
            type_str.set("""
'code mission', 'code topographique', 'type de topo', 'date limite de validite', 'code parite', 
'borne superieure', 'code types de structure', 'identifiant UA', 'borne inferieure', 'code UA', 
'cle UA', 'code codique', 'filler', 'date effet', 'libelle mission', 'filler2'""")
            label_pre_type.lift()
            sep_type.place(relx = 0.305, rely = 0.1, anchor='nw', height=60)
            label_info.config(font = ('Arial', 8))



def active_frame2():

    bouton_lb3.state(['disabled'])
    bouton_lb2.state(['!disabled'])
    bouton_lb2.place(relx = 0.999, rely= 0.985, anchor="se")
    bouton_effacer.state(['!disabled'])
    label_frame2.config(bootstyle = 'primary')
    label_2.config(bootstyle = 'dark')
    label_ask.config(foreground='#0063cb')
    sep2.config(bootstyle ="dark")


def active_frame3():

    cb.bind('<<ComboboxSelected>>', remove_highlight)
    cb.bind('<FocusIn>', remove_highlight)
    cb.bind('<Button-1>', remove_highlight)
    bouton_retour.place(relx = 0.995, rely = 0, anchor='ne')
    label_3.config(bootstyle = 'dark')
    label_info.config(foreground='#0063cb')
    label_type.config(foreground='#2f2f2f')
    label_pre_type.config(foreground='#2f2f2f')
    cb.state(['!disabled'])
    cb.state(['readonly'])
    cb.config(bootstyle = 'secondary', foreground="#2a2a2a")
    bouton_lb3.state(['!disabled'])
    bouton_lb3.place(relx = 0.999, rely= 0.985, anchor="se")
    sep3.config(bootstyle ="info")
    sep_type.config(bootstyle = 'info')
        
def hide_frame3():

    label_frame3.config(bootstyle = 'info')
    label_3.config(bootstyle = 'info')
    cb.state(['disabled'])
    cb.config(bootstyle = 'info')
    cb.set('...')
    cb.place(relx = 2, rely = 2, anchor ='nw')
    info.set('')
    type_str.set('')
    pre_type.set('')
    sep_type.place(relx= 2, rely=2, anchor="center")
    bouton_retour.place(relx =2, rely= 2, anchor = 'nw')
    bouton_lb3.place(relx = 2 , rely = 2, anchor = 'nw')
   
def hide_frame4():

    bouton_retour4.place(relx = 2, rely = 2, anchor = 'nw')
    bouton_lb4.place(relx = 2, rely = 2, anchor = "nw")
    check_validation.place(relx = 2, rely = 2, anchor = 'nw')
    bouton_lb4.state(['disabled'])
    label_4.config(bootstyle = 'info')
    label_frame3.config(bootstyle = 'primary')
    label_frame4.config(bootstyle = 'info')
    recap.set('')
    label_x.place(relx = 2, rely = 2, anchor = 'nw')
    sep4.config(bootstyle ="info")


#Fonction retour : 
def retour():

    if gb.var == 2 : 

        active_frame2()
        hide_frame3()
        gb.var = 1     
    
    if gb.var == 3:

        active_frame3()
        hide_frame4()
        gb.var = 2


#Effacer le fichier selectionné
def supp():

    gb.file_path = ""
    imported_file.set('')
    bouton_lb2.state(['disabled'])
    bouton_ask.state(['!disabled'])
    bouton_effacer.place(relx = 2, rely= 2, anchor='nw')

#Sélectionner un fichier
def file_select():

    try :
    
        gb.file_path = filedialog.askopenfilename(initialdir = "/",
                                                title = "Choisissez votre fichier",
                                                filetypes = (("Fichiers", "*.*"),))
        
        taille_fichier = os.path.getsize(gb.file_path)
        
        if taille_fichier < 1000000000:
            var_o = 'Mo'
            result = taille_fichier/1000000
        elif  1000000000 < taille_fichier:
            var_o = 'Go'
            result = taille_fichier/1000000000
        

        if gb.file_path != '':

            a = 'Vous avez selectionné le fichier suivant : '+gb.file_path+", poids = "+str(round(result,2))+" "+var_o
            length = len(a)

            # Determine the cutoff index based on the length of the save_path
            if length < 129:
                imported_file.set(a)
                
            elif length >= 129:
                cutoff_index = (((length - 55) // 20 + 1) * 18) 
                display_path = a[:60] + "..." + a[cutoff_index:]
                imported_file.set(display_path)
            
            bouton_effacer.state(['!disabled'])
            bouton_ask.state(['disabled'])
            bouton_lb2.state(['!disabled'])
            bouton_effacer.place(relx = 0.7, rely = 0.5, anchor = 'center')
            bouton_lb2.place(relx = 0.999, rely= 0.985, anchor="se")

        else: 
            imported_file.set("Chemin d'accès non valide")

    #Cas où l'utilisateur ouvre l'interface de file_select mais ne choisit rien
    except FileNotFoundError:
        pass


#Application des boutons pour passer d'une frame à une autre.
def apply_func():

    if gb.var == 1:

        #Mise en vert frame 2
        bouton_lb2.state(['disabled'])
        bouton_lb2.place(relx = 2, rely = 2, anchor = 'nw')
        bouton_effacer.state(['disabled'])
        label_frame2.config(bootstyle = 'success')
        label_2.config(bootstyle = 'success')
        label_ask.config(foreground="#4b9f6c")
        label_frame3.config(bootstyle = 'primary')
        sep2.config(bootstyle = "success")

        #Activer frame 3

        cb.place(relx = 0.15, rely=0.5, anchor='center')        
        bouton_retour.place(relx = 0.995, rely = 0, anchor='ne')
        label_3.config(bootstyle = 'dark')
        label_info.config(foreground='#0063cb')
        label_type.config(foreground='#2f2f2f')
        label_pre_type.config(foreground='#2f2f2f')
        cb.state(['!disabled'])
        cb.state(['readonly'])
        cb.bind('<<ComboboxSelected>>', remove_highlight)
        cb.bind('<FocusIn>', remove_highlight)
        cb.bind('<Button-1>', remove_highlight)
        cb.config(bootstyle = 'secondary', foreground="#2a2a2a")
        bouton_lb3.state(['!disabled'])
        
        gb.var = 2
    

    elif gb.var == 2:

        #Mise en vert frame 3
        bouton_lb3.state(['disabled'])
        bouton_lb3.place(relx = 2, rely = 2, anchor = 'nw')
        cb.state(['disabled'])
        cb.config(bootstyle = 'sucess', foreground = "#4b9f6c")
        cb.unbind('<Button-1>')
        label_info.config(foreground="#4b9f6c")  
        label_type.config(foreground="#4b9f6c") 
        label_pre_type.config(foreground="#4b9f6c")
        label_frame3.config(bootstyle = 'success')
        label_3.config(bootstyle = 'success')
        bouton_retour.place(relx = 2, rely = 2, anchor = 'nw')
        sep3.config(bootstyle = "success")
        sep_type.config(bootstyle = "success")
        
        #Activer frame 4
    
        check_var.set(value = 0)
        label_4.config(bootstyle = 'dark')
        label_frame4.config(bootstyle = 'primary')
        func_recap()
        bouton_lb4.state(['!disabled'])
        cancel_button.state(['!disabled'])
        bouton_retour4.place(relx = 0.995, rely = -0.025, anchor='ne')
        bouton_retour4.lift()
        gb.var = 4

        gb.var = 3

        
    elif gb.var == 3:

        #Efface frame 5 
        
        bouton_lb5.state(['disabled'])
        bouton_leave_or_enter.state(['disabled'])
        label_frame5.config(bootstyle = 'info')
        label_desc.config(text='', foreground="#2a2a2a")
        bouton_download.place(relx = 2, rely = 2, anchor = 'nw')
        bouton_download.state(['disabled'])
        bouton_lb5.place(relx = 2, rely = 2, anchor = 'nw')
        bouton_leave_or_enter.place(relx = 2, rely = 2, anchor = 'nw')
        label_5.config(foreground=  "#b2b2b2")

        entry_download.bind("<Button-1>", on_entry_click)
        entry_download.bind('<Return>', enter_pressed)

        entry_download.state(['!disabled'])

        sep5.config(bootstyle = "info")


        #Efface frame 4

        label_frame4.config(bootstyle = 'info')
        label_4.config(bootstyle = 'info')
        recap.set('')
        label_recap.config(foreground="#2f2f2f")

        sep4.config(bootstyle = "info")
        
        #Efface frame 3

        label_frame3.config(bootstyle ='info')
        label_3.config(bootstyle = 'info')
        cb.set('...')
        cb.state(['disabled'])
        cb.config(bootstyle = 'info')
        info.set('')
        cb.place(relx = 2, rely = 2, anchor ='center')
        label_info.config(foreground='#0063cb')
        type_str.set('')
        pre_type.set('')
        label_type.config(foreground='#2f2f2f')
        label_pre_type.config(foreground= "#2f2f2f")
        sep3.config(bootstyle = "info")
        sep_type.place(relx = 2, rely = 2, anchor = 'center')

        #Reset frame 2

        label_frame2.config(bootstyle = 'primary')
        label_2.config(bootstyle = 'dark')
        label_ask.config(foreground='#0063cb')
        imported_file.set('')
        bouton_ask.state(['!disabled'])
        bouton_effacer.place(relx = 2, rely = 2, anchor = 'nw')

        sep2.config(bootstyle = "info")
        sep_type.config(bootstyle ="info")
        
        #Reset de toutes les variables globales
        gb.file_path = ""
        gb.nom_fichier = ""
        gb.var = 1
        gb.thread = None
        gb.run = True
        gb.fichier_app = ""
        gb.save_path = ""
        gb.entry_clicked = False
        gb.leave = False



#Bouton qui gère la confirmation dans le thread 3 
def affichage_():

    if check_var.get() == '0' :

        bouton_lb4.place(relx = 2, rely = 2, anchor = 'nw')
 
    elif check_var.get() == '1' :
        
        bouton_lb4.place(relx = 0.999, rely= 0.985, anchor="se")
        bouton_lb4.state(['!disabled'])
        bouton_lb4.lift()


#Activation des fonctions
def appel_func():

    #Positionnement des widgets lorsque le traitement s'active
    bouton_lb4.state(['disabled'])
    check_var.set(value=0)
    check_validation.place(relx = 2, rely = 2, anchor = 'nw')
    cancel_button.place(relx = 0.999, rely= 0.985, anchor="se")
    cancel_button.config(bootstyle = 'primary')
    cancel_button.state(['!disabled'])
    bouton_lb4.place(relx = 2, rely = 2, anchor = 'nw')
    progress_bar.place(relx=0.5, rely=0.7, anchor='center')
    bouton_retour4.place(relx = 2, rely= 2, anchor = 'nw')
    label_recap.config(foreground="#2f2f2f")
    recap.set('Traitement en cours, veuillez patienter, ce traitement peut prendre plusieurs minutes en fonction de la taille de votre fichier.')
    label_recap.place(relx = 0.5, rely = 0.4, anchor = 'center')
    label_x.place(relx = 2, rely = 2, anchor = 'nw')
    progress_bar.config(mode = 'indeterminate')
    progress_bar.start(120)
    
    #Lancement du thread
    start_thread()


def start_thread():

    gb.event.clear()
    gb.run = True
    label_frame4.config(bootstyle = 'warning')
    #Attribue le thread et et le lance
    gb.thread = threading.Thread(target=run, args=(cb.get(),gb.file_path))
    gb.thread.start()
    
    
def stop_thread():
    #Stop le thread
    gb.event.set()
    progress_bar.stop()

        #Remets en place la frame avant traitement
    check_validation.place(relx = 0.28, rely = 0.70, anchor='nw')
    label_frame4.config(bootstyle = 'primary')
    label_recap.place(relx = 0.005, rely = 0.2, anchor="nw")
    label_x.place(relx = 0.005, rely = 0.69, anchor = 'nw')
    bouton_lb4.place(relx = 0.999, rely= 0.985, anchor="se")
    bouton_lb4.state(['disabled'])
    recap.set(f"\nVous avez selectionné le fichier : {os.path.basename(gb.file_path)}, et le filtre suivant : {cb.get()}.")
    bouton_retour4.place(relx = 0.995, rely = -0.025, anchor='ne')

 
    #Cache ses widgets
    progress_bar.place(relx = 2, rely = 2, anchor = 'nw')
    cancel_button.place(relx = 2, rely = 2, anchor = 'nw')
    cancel_button.state(['disabled'])
    



def run(filtre, file_path):
        
    #On nomme les colonnes en fonction du filtre selectionné
    run = True
    if filtre == 'TOPO' and  file_path!="":
        
        column_names = [
                    'code topo', 
                    'libelle', 
                    'type commune actuel (R ou N)',
                    'type commune FIP (RouNFIP)', 
                    'RUR actuel', 
                    'RUR FIP', 
                    'caractere voie',
                    'annulation', 
                    'date annulation', 
                    'date création de article',
                    'type voie', 
                    'mot classant', 
                    'date derniere transition'
                    ]
                    
        column_positions = [
                    (0, 17), 
                    (18, 87), 
                    (88, 88),
                    (89, 89),
                    (90, 90),
                    (91, 91),
                    (92, 92),
                    (93, 93),
                    (94, 101),
                    (102, 109),
                    (110, 110),
                    (111, 118),
                    (119, 127),
                    ]
    
    elif filtre ==  "ACHEMINEMENT" and file_path!="":
        
        column_names = [
            'code topographique',
            'date limite de validite',
            'code parite',
            'borne superieure',
            'borne inferieure',
            'date effet',
            'code postal',
            'libelle de la donnee postale',
            'indicateur de pluridistribution',
            'code localite',
            'libelle ligne 5',
            'code distribution',
            'code type adresse',
            'type code postal',
            'code ancienne commune'
            ]
        
        column_positions = [
            (0, 8), 
            (9, 16), 
            (17, 17),
            (18, 22),
            (23, 27),
            (28, 35),
            (36, 40),
            (41, 72),
            (73, 73),
            (74, 78),
            (79, 116),
            (117, 117),
            (118, 118),
            (119, 119),
            (120, 122)
            ]
    elif filtre == "COMPETENCES" and file_path!="":
     
        column_names= [
       'code mission', 
       'code topographique', 
       'type de topo',
       'date limite de validite', 
       'code parite', 
       'borne superieure',
       'code types de structure', 
       'identifiant UA', 
       'borne inferieure',
       'code UA', 
       'cle UA', 
       'code codique', 
       'filler', 
       'date effet',
       'libelle mission', 
       'filler2'
        ]
       
        column_positions = [
       (0, 14), 
       (15, 30), 
       (31, 32),
       (33, 40),
       (41, 41),
       (42, 46),
       (47, 51),
       (52, 61),
       (62, 66),
       (67, 76),
       (77, 77),
       (78, 84),
       (85, 89),
       (90, 97),
       (98, 123),
       (124, 194),
       ]

    
    elif filtre == "STRUCTURES" and  file_path!="":

        column_names = [
        'numero UA', 
        'date limite validite1', 
        'code UA', 
        'date limite validite2', 
        'code codique', 
        'code type structure', 
        'filler1', 
        'date effet', 
        'date previsionnelle', 
        'etat', 
        'date effet etat', 
        'service apres recodification', 
        'service avant recodification', 
        'numero UA de la DIR', 
        'numero UA de la CI', 
        'numero UA de la CB', 
        'DIR hierarchique', 
        'CI hierarchique', 
        'CB hierarchique', 
        'libelle long 1', 
        'libelle long 2', 
        'libelle court 1', 
        'libelle court 2', 
        'adresse codifiee', 
        'adresse dans la voie', 
        'adresse complete', 
        'adresse abregee', 
        'prefixe', 
        'telephone 1', 
        'telephone 2', 
        'telephone special', 
        'messagerie electronique', 
        'numero MMA', 
        'jours heures de reception', 
        'reception sur RDV', 
        'horaire telephonique', 
        'code SIRET', 
        'OS', 
        'liens fonctionnels', 
        'CB1', 
        'code flux CB1', 
        'CB2', 
        'code flux CB2', 
        'codes type liens fonctionnels', 
        'filler2'
        ]


        column_positions = [
        (0, 9), 
        (10, 17), 
        (18, 27),
        (28, 35),
        (36, 42),
        (43, 47),
        (48, 52),
        (53, 60),
        (61, 68),
        (69, 69),
        (70, 77),
        (78, 87),
        (88, 97),
        (98, 107),
        (108, 117),
        (118, 127),
        (128, 130),
        (131, 135),
        (136, 142),
        (143, 172),
        (173, 202),
        (203, 217),
        (218, 232),
        (233, 250),
        (251, 255),
        (256, 445),
        (446, 511),
        (512, 521),
        (522, 531),
        (532, 541),
        (542, 551),
        (552, 609),
        (610, 613),
        (614, 673),
        (674, 733),
        (734, 793),
        (794, 807),
        (808, 809),
        (810, 869),
        (870, 907),
        (908, 910),
        (911, 948),
        (949, 951),
        (952, 981),
        (982, 1009)
        ]
    
    elif filtre == "UAMISSIONS" and file_path != "":

        column_names = [
        'numero UA', 
       'date fin mission exercee1', 
       'code mission', 
       'date fin mission exercee2', 
       'code UA', 
       'date fin mission exercee3', 
       'code codique', 
       'code type de structure', 
       'date effet mission exercee', 
       'libelle court mission', 
       'libelle long mission', 
       'indicateur mission heritee', 
       'indicateur mission obligatoire', 
       'indicateur mission competence geographique'
       ]

        column_positions = [
       (0, 9), 
       (10, 17), 
       (18, 32),
       (33, 40),
       (41, 50),
       (51, 58),
       (59, 65),
       (66, 70),
       (71, 78),
       (79, 104),
       (105, 315),
       (316, 316),
       (317, 317),
       (318, 318)
       ]
        
    #Fonction
    while not gb.event.is_set():              
        with open(file_path, 'r', encoding='utf-8') as file:
                    lines = file.readlines()

        
        processed_data = []
        for line in lines:
            row = []             
            for start, end in column_positions:
                row.append(line[start:end].strip())   
            processed_data.append(row)   

            #Pour annuler l'éxécution 
            if gb.event.is_set():
                run = False
                break
        
        break
        
    #Si le traitement n'est pas intérrompu :
    if run == True:         
        df = pd.DataFrame(processed_data, columns=column_names)
        df.drop(index=0, inplace=True)
        gb.fichier_app = df
        file.close()
        go_next()
        gb.event.set()
    


def go_next():
   
    #Fin frame 4
    progress_bar['value'] = 0
    progress_bar.stop()
    cancel_button.place(relx = 2, rely = 2, anchor = 'nw')
    label_frame4.config(bootstyle = 'success')
    label_4.config(bootstyle = 'success')
    progress_bar.place(relx = 2, rely = 2, anchor = 'nw')
    recap.set('Traitement du fichier réussi !')
    cancel_button.state(['disabled'])
    label_recap.config(foreground="#4b9f6c")
    label_recap.place(relx = 0.5, rely = 0.5, anchor ='center')
    sep4.config(bootstyle = "success")

    #Mise en place frame 5
    label_frame5.config(bootstyle = 'primary')
    label_5.config(foreground ="#2f2f2f" )

    entry_string.set('...')
    label_desc.config(text = "Nommez votre fichier ci-dessus et validez, la mention '_gen' avec la date du jour sera ajoutée au nom choisi.")
    label_desc.place(relx = 0.5, rely = 0.8, anchor='center')
    entry_download.place(relx = 0.5, rely = 0.5, anchor='center')

#Lorsque l'utilisateur valide le nom du fichier qu'il souhaite exporter
def enter_pressed(event):

    if entry_string != '':
        last_step()
    else : 
        pass
    

#Fonction pour save
def save_file():

    download_thread = threading.Thread(target=save_file_in_background)
    download_thread.start()
    
#Thread pour save le fichier traité, car sinon l'interface se fige lorsqu'un fichier est de taille trop importante
def save_file_in_background():
    
    
    func_save_csv(label_frame5, label_desc, entry_string.get(),gb.fichier_app)
    
    if gb.download == True : 

        #bouton_download.place(relx = 0.5, rely = 0.35, anchor='center')
        #label_desc.place(relx = 0.5, rely = 0.45, anchor = 'center')
        bouton_lb5.place (relx = 1, rely = 1, anchor = "se", width= 160, height = 30)
        bouton_leave_or_enter.place(relx = 0, rely = 1, anchor = 'sw', width= 160, height = 30)
        bouton_leave_or_enter.state(['!disabled'])
        bouton_lb5.state(['!disabled'])
        label_frame5.config(bootstyle='success')
        bouton_download.config(bootstyle ='success-outline')
        label_5.config(foreground="#4b9f6c")
        sep5.config(bootstyle ="success")


        message = "Fichier exporté avec succès vers :"
        length = len(gb.save_path)

        # Determine the cutoff index based on the length of the save_path
        if length < 55:
            display_path = gb.save_path
        else:
            cutoff_index = (((length - 55) // 20 + 1) * 20)+ 10
            
            display_path = gb.save_path[:9] + "..." + gb.save_path[cutoff_index:]

        # Update the label and button
        label_desc.config(text=f'{message}{display_path}', foreground="#4b9f6c")
        
                
        

def func_save_csv(label_frame5, label_desc, base, fichier_app):
    
    try:

        gb.save_path = filedialog.asksaveasfilename(initialfile=base+"_gen_"+datetime.today().strftime("%d_%m_%Y")+".csv",defaultextension=".csv", filetypes=(("Fichier csv", "*.csv"),))

        #newline ='' pour éviter les sauts de lignes dans le fichier csv
        with open(gb.save_path, "w", newline='') as f :

            #configuration graphique pour prévenir l'utilisateur que le téléchargement est en cours
            
            
            label_desc.config(text = "")
            bouton_download.state(['disabled'])
            label_desc.config(text = "Téléchargement en cours, cette action peut prendre plusieurs secondes en fonction de la taille de votre fichier.")
            label_frame5.config(bootstyle = 'warning')
    
            #Sauvegarde
            fichier_app.to_csv(f, encoding='utf-8', index=False, sep=';')
            gb.download = True
            
    except FileNotFoundError:
        pass
    except PermissionError:
        pass

#Lorsque l'utilisateur valide un nom pour le télécharger
def last_step():

    entry_download.unbind("<Button-1>")
    entry_download.unbind('<Return>')
    entry_download.place(relx =2, rely =2, anchor='center')
    entry_download.state(['disabled'])
    dl_string.set(f'{entry_string.get()}_gen_{datetime.today().strftime("%d_%m_%Y")+".csv"}')
    bouton_download.state(['!disabled'])
    bouton_download.config(bootstyle= 'secondary-outline')
    bouton_download.place(relx = 0.5, rely = 0.5, anchor='center')
    bouton_leave_or_enter.place(relx = 2, rely = 2, anchor='nw')
    bouton_leave_or_enter.config(bootstyle = 'danger', text = 'Quitter')
    bouton_leave_or_enter.state(['disabled'])
    label_desc.config(text = "Cliquez sur le bouton ci-dessus pour télécharger votre fichier.")
    gb.leave = True

def on_entry_click(event):
    
    entry_string.set('')
    entry_download.unbind('<Button-1>')
    

def update_button_state(*args):
    if entry_string.get().strip() and entry_string.get() !='...':
        bouton_leave_or_enter.config(state="normal")
        bouton_leave_or_enter.config(text= 'Valider', bootstyle = 'primary')
        bouton_leave_or_enter.place(relx = 0.65, rely = 0.5, anchor= 'center', height=30, width=60)
    else:
        bouton_leave_or_enter.config(state="disabled")

def leave_or_enter():

    if not gb.leave:
        last_step()
    else : 
        window.destroy()


#Pour éviter qu'un thread continue de tourner en bg alors que la window est fermée.
def fermeture():
    if gb.thread != None:
        stop_thread()
        window.destroy()
    else : 
        window.destroy()


if __name__ == '__main__' : 

    window = ttk.Window()
    window.iconbitmap("LogoAPIDO.ico")
    
    #Image

    tk_image1 = tk.PhotoImage(file = "apido_resized.png")
    label_image1 = tk.Label(window, image=tk_image1)
    label_image1.place(relx=0.945, rely=0, anchor="ne")
    
    tk_image2 = tk.PhotoImage(file = "logo_finance.png")
    label_image2 = tk.Label(window, image=tk_image2)
    label_image2.place(relx = 0.013, rely = 0.012, anchor="nw")



    #Setup fenêtre
    window.geometry('1300x700')
    window.minsize(1300,700)
    window.maxsize(1366,768)
    window.title('Application pour le référentiel TOPAD')
    window.columnconfigure((0,3), weight=1, uniform='a')
    window.columnconfigure((1,2), weight=2, uniform='a')
    window.rowconfigure((6), weight=1, uniform='a')
    window.rowconfigure((0,1), weight=3, uniform='a')
    window.rowconfigure((2,3,4), weight=4, uniform='a')
    window.rowconfigure((5), weight= 4, uniform='a')


    #Chargement du thème personnalisé
    style = Style()
    style.load_user_themes('user.json')
    style.theme_use('ihm_custom') 

    style.configure('Custom.TFrame', background='#f6f6f6')

    frame_title = ttk.Frame(window)
    titre = ttk.Label(frame_title, text="OPEN DATA", font=("Arial", 24, 'bold'), foreground="#21213f")
    titre2 = ttk.Label(frame_title, text="TOPAD", font=("Arial", 24, 'bold'), foreground="#000091")
    titre.place(relx=0.5, rely=0.3, anchor='center')
    titre2.place(relx=0.5, rely=0.7, anchor='center')
    frame_title.grid(column=1, columnspan=2, row = 0,  sticky='nsew')
     
    label_frame = ttk.Frame(window, style = 'Custom.TFrame')
    label_frame.grid(column=1, columnspan=2, row = 1,  sticky='nsew')

    #Disposition de tous les LabelFrame   

    label_frame2 = ttk.LabelFrame(window, text='ÉTAPE 1',bootstyle = 'info', borderwidth= 4)
    label_frame2.grid(column=1, columnspan=2, row=2, sticky='nsew', pady=2)

    label_frame3 = ttk.LabelFrame(window, text='ÉTAPE 2',bootstyle = 'info', borderwidth= 4)
    label_frame3.grid(column=1, columnspan=2, row=3, sticky='nsew', pady=2)

    label_frame4 = ttk.LabelFrame(window, text='ÉTAPE 3',bootstyle = 'info', borderwidth= 4)
    label_frame4.grid(column=1, columnspan=2, row=4, sticky='nsew', pady=2)

    label_frame5 = ttk.LabelFrame(window, text='ÉTAPE 4',bootstyle = 'info', borderwidth= 4)
    label_frame5.grid(column=1, columnspan=2, row=5, sticky='nsew', pady=2)


    #Widgets Frame 1

    label_1 = ttk.Label(label_frame, text = 'Introduction', font=("Arial",10), background="#f6f6f6")
    label_explication = ttk.Label(label_frame, text = "Bienvenue dans cette application destinée à traiter les fichiers TOPAD produits par le bureau DP6 de la DGFiP. Vous serez invités à choisir d'abord le \nfichier que vous souhaitez traiter, puis à indiquer le filtrage souhaité. Veuillez noter que le temps de traitement de vos fichiers dépend de leurs tailles \nrespectives.", font=("Arial",9), background="#f6f6f6")
    sep1 = ttk.Separator(label_frame, bootstyle = 'info')

    label_1.place(relx = 0.005, rely= 0.025, anchor='nw')
    label_explication.place(relx = 0.01, rely= 0.3 )
    sep1.place(relx = 0.01, rely= 0.25, anchor = "nw", width=85)

    #Widgets Frame 2
    
    label_2 = ttk.Label(label_frame2, text = "Importation du fichier source", bootstyle = 'info', font=("Arial",10))
    bouton_ask = ttk.Button(label_frame2, text = 'Choisir son fichier', bootstyle = 'secondary', command = file_select)
    sep2 = ttk.Separator(label_frame2, bootstyle = 'info')
    imported_file = tk.StringVar()
    label_ask = ttk.Label(label_frame2, textvariable=imported_file, foreground='#0063cb', font=("Arial",9))

    sep2.place(relx = 0.01, rely= 0.18, anchor = "nw", width=50)
    label_2.place(relx = 0.005, rely= -0.05, anchor='nw')
    bouton_ask.place(relx = 0.5, rely=0.5, anchor='center')
    label_ask.place(relx = 0.005, rely = 0.75, anchor= 'nw' )

    #Widgets qui seront placés dans le futur
    bouton_lb2 = ttk.Button(label_frame2, text = 'Confirmer', bootstyle = 'primary', command=apply_func)
    bouton_effacer = ttk.Button(label_frame2, text = 'Effacer le fichier selectionné', bootstyle = 'secondary', command=supp)


    #Frame 3 widgets
    label_3 = ttk.Label(label_frame3, text = 'Sélection du type de fichier à traiter', bootstyle = 'info', font=("Arial",10))
    info = tk.StringVar(value='')
    label_info = ttk.Label(label_frame3, textvariable=info, foreground='#0063cb', font=("Arial",9))
    type_str = tk.StringVar(value='')
    sep_type = ttk.Separator(label_frame3, bootstyle = 'info', orient = 'vertical')
    pre_type = tk.StringVar(value ="")
    label_pre_type = ttk.Label(label_frame3, textvariable=pre_type, foreground='#2f2f2f',  font=("Arial",9))
    label_type = ttk.Label(label_frame3, textvariable=type_str, foreground='#2f2f2f',  font=("Arial",9))
    sep3 = ttk.Separator(label_frame3, bootstyle = 'info')
    
    items = ['TOPO', 'UAMISSIONS','STRUCTURES','COMPETENCES','ACHEMINEMENT']
    liste_option = tk.StringVar(value = '...')
    cb = ttk.Combobox(label_frame3, textvariable=liste_option, bootstyle = 'info', font=("Arial",9))
    
    cb['values'] = items
    cb.state(['disabled'])
    cb.bind('<<ComboboxSelected>>', remove_highlight)
    cb.bind('<FocusIn>', remove_highlight)
    cb.bind('<Button-1>', remove_highlight)

    sep3.place(relx = 0.01, rely= 0.18, anchor = "nw", width=50)
    label_3.place(relx = 0.005, rely= -0.05, anchor='nw')
    label_info.place(relx = 0.15, rely = 0.82, anchor= 'center' )
    label_type.place(relx = 0.312, rely = 0.095, anchor='nw')
    label_pre_type.place(relx = 0.312, rely = 0.065, anchor='nw')

    #Widgets qui seront placés dans le futur
    bouton_lb3 = ttk.Button(label_frame3, text = 'Confirmer', bootstyle = 'primary', command=apply_func)
    bouton_retour = ttk.Button(label_frame3, text = 'Retour', command = retour, bootstyle = 'danger')
    
    #Frame 4 widgets
    label_4 = ttk.Label(label_frame4, text = 'Lancement du traitement', bootstyle = 'info', font=("Arial",10))
    sep4 = ttk.Separator(label_frame4, bootstyle = 'info')
    label_x = ttk.Label(label_frame4, text ="Cochez la case suivante pour confirmer :" , foreground="#2f2f2f", font=("Arial",9))
    recap = tk.StringVar(value ='')
    label_recap = ttk.Label(label_frame4, textvariable = recap, foreground="#2f2f2f", font=("Arial",9))
    check_var = tk.StringVar(value = 0)
    check_validation = ttk.Checkbutton(label_frame4, bootstyle = 'light', variable=check_var, command=affichage_)
    bouton_lb4 = ttk.Button(label_frame4, text = 'Commencer le traitement', bootstyle = 'primary', command = appel_func)
    bouton_retour4 = ttk.Button(label_frame4, text = 'Retour', bootstyle = 'danger', command= retour)
    progress_bar = ttk.Progressbar(label_frame4, bootstyle = 'warning')
    cancel_button = ttk.Button(label_frame4, text = 'Annuler', command= stop_thread)

    label_4.place(relx = 0.005, rely= -0.05, anchor='nw')
    sep4.place(relx = 0.01, rely= 0.18, anchor = "nw", width=50)

    
    #Frame 5 widgets
    label_5 = ttk.Label(label_frame5, text ='Exportation du fichier CSV final', bootstyle = 'info', font=("Arial",10))
    sep5 = ttk.Separator(label_frame5, bootstyle = 'info')
    bouton_lb5 = ttk.Button(label_frame5, text = 'Traiter un nouveau fichier', bootstyle = 'primary', command=apply_func)
    
    
    label_desc = ttk.Label(label_frame5, text = "", bootstyle = 'dark', font=("Arial",9))
    dl_string = tk.StringVar(value = '')
    bouton_download = ttk.Button(label_frame5, textvariable=dl_string, bootstyle ='secondary-outline' , command = save_file)
    
    entry_string = tk.StringVar(value='...')
    # Variable associée à l'Entry
    
    entry_string.trace_add("write", update_button_state)


    entry_download = ttk.Entry(label_frame5, textvariable=entry_string)
    entry_download.bind("<Button-1>", on_entry_click)
    entry_download.bind('<Return>', enter_pressed)
    
    bouton_leave_or_enter = ttk.Button(label_frame5, text = 'VALIDER', command= leave_or_enter)

    label_5.place(relx = 0.005, rely= -0.05, anchor='nw')
    sep5.place(relx = 0.01, rely= 0.18, anchor = "nw", width=50)
    
    #Etat initial des boutons (optionnel)
    bouton_lb2.state(['disabled'])
    bouton_lb3.state(['disabled'])
    bouton_lb4.state(['disabled'])
    bouton_lb5.state(['disabled'])
    bouton_download.state(['disabled'])
    
    window.protocol("WM_DELETE_WINDOW", fermeture)

    bouton_ask.config(bootstyle = 'secondary')
    label_frame2.config(bootstyle = 'primary')
    label_2.config(bootstyle = 'dark')
    


    window.mainloop()





