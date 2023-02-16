# -*- coding: utf-8 -*-
"""
Created on Mon May  9 01:05:56 2022

@author: morin
"""
import pandas as pd
import numpy as np
from tkinter import *
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import (Figure)
import matplotlib.pyplot as plt



'___________________Ouverture et récupération données du tableau excel_____________________'

df = pd.read_excel ('Flags.xlsx',"Flags",header=0)
print(df)
# print(df['Années'])

def calculerfcf ():
    global Taux_Inflation, Tarif_VL, cout_fixe, taux_croissance_VL, passage_VL, Tarif_PL, taux_croissance_PL, passage_PL, cout_fixe, marge_cout_var
    global DA, T, taux_BFR, invest_par_an, r
    Taux_Inflation = Tx_I.get()
    Tarif_VL = TVL.get()
    taux_croissance_VL=0.03
    passage_VL= 2000000
    Tarif_PL=3*Tarif_VL
    taux_croissance_PL=0.015
    passage_PL=1000000
    cout_fixe = CF.get()
    marge_cout_var=0.2
    DA=105000000/20
    T=33/100
    taux_BFR=5/100
    invest_par_an=105000000/3
    r=0.1
    
    print(Taux_Inflation, Tarif_VL, cout_fixe, sep="/")
    
    'Calcul de la liste inflation'
    liste_inflation=[]
    base_inflation=1/((1+Taux_Inflation)**4)
    for i in range(len(df['Années'])):    
        Inflation=df['Indicateur exploitation'][i]*base_inflation*(1+Taux_Inflation)
        liste_inflation.append(Inflation)
        base_inflation=base_inflation*(1+Taux_Inflation)
    print(liste_inflation)
    print(len(liste_inflation))
    
    
    
    'Calcul du CA VL'
    
    CA_VL=0
    liste_CA_VL=[]
    for i in range(len(df['Années'])):
        CA_VL=df['Indicateur exploitation'][i]*Tarif_VL*liste_inflation[i]*passage_VL*df['Indexation passage VL'][i]
        # print(i)
        # print(CA_VL)
        liste_CA_VL.append(CA_VL)
    
    ' Calcul du CA PL'

    CA_PL=0
    liste_CA_PL=[]
    for i in range(len(df['Années'])):
        CA_PL=df['Indicateur exploitation'][i]*Tarif_PL*liste_inflation[i]*passage_PL*df['Indexation passage PL'][i]
        # print(i)
        # print(CA_PL)
        liste_CA_PL.append(CA_PL)
        # print(liste_CA_PL)
        
    liste_CA=[]
    for i in range(len(liste_CA_PL)):
        CA=liste_CA_PL[i]+liste_CA_VL[i]
        liste_CA.append(CA)
        # print(liste_CA)
    '_______________________ Tableau des Charges__________________________________'

    'Calcul des couts fixes'
    
    liste_cout_fixe=[]
    for i in range(len(df['Années'])):
        cout_fixe_ind=df['Indicateur exploitation'][i]*liste_inflation[i]*cout_fixe
        liste_cout_fixe.append(cout_fixe_ind)
    
    'Calcul des couts variables'
    
    liste_cout_var=[]
    for i in range(len(df['Années'])):
        cout_var_ind=df['Indicateur exploitation'][i]*liste_CA[i]*marge_cout_var
        liste_cout_var.append(cout_var_ind)
        # print(liste_cout_var)
        
    'Calcul des couts globaux'
    liste_cout=[]
    for i in range(len(liste_cout_fixe)):
        cout=liste_cout_fixe[i]+liste_cout_var[i]
        liste_cout.append(cout)
    
    '_____________________________________________________________________________'
    '_______________ Tableau Flux de Tresorerie nette ____________________________'
    
    'calcul EBE'
    
    liste_EBE=[]
    for i in range(len(liste_cout_fixe)):
        EBE= liste_CA[i]-liste_cout[i]
        liste_EBE.append(EBE)
        # print(liste_EBE)
    'on cherche à remplacer nan par 0'
    df_EBE = pd.DataFrame(liste_EBE)
    df_EBE = df_EBE.fillna(0)
    df_EBE.columns=['EBE']
    # print(df_EBE)
    liste_EBE=[]
    for i in range(len(df_EBE['EBE'])):
        element=df_EBE['EBE'][i]
        # print(element)
        liste_EBE.append(element)
        # print(liste_EBE)
    
    'Calcul BFR'
    
    liste_BFR=[]
    for i in range(len(liste_CA)-1):
        BFR=taux_BFR*liste_CA[i+1]
        liste_BFR.append(BFR)
    liste_BFR.append(0)
    # print(liste_BFR)
    'On cherche à remplacer nan par 0'
    df_BFR = pd.DataFrame(liste_BFR)
    df_BFR = df_BFR.fillna(0)
    df_BFR.columns=['BFR']
    # print(df_BFR)
    liste_BFR=[]
    for i in range(len(df_BFR['BFR'])):
        element=df_BFR['BFR'][i]
        # print(element)
        liste_BFR.append(element)
        # print(liste_BFR)
    
    'Calcul variation du BFR'
    
    liste_var_BFR=[]
    for i in range(len(liste_BFR)):
        var_BFR=liste_BFR[i]-liste_BFR[i-1]
        liste_var_BFR.append(var_BFR)
        # print(liste_BFR[i])
        # print(liste_BFR[i-1])
        # print(var_BFR)
        # print(i)
        # print(liste_var_BFR)
        
    'Calcul de DA'
    
    liste_DA=[]
    for i in range(len(df['Années'])):
        DA_ind=DA*df['Indicateur exploitation'][i]
        liste_DA.append(DA_ind)
        # print(liste_DA)
    
    'Calcul de l investissement initial'
    
    liste_invest=[]
    for i in range(len(df['Années'])):
        invest=invest_par_an*df["Indicateur construction"][i]
        liste_invest.append(invest)
        # print(liste_invest)
    
    'Calcul Flux net de trésorerie'
    
    global liste_FCF
    liste_FCF=[]
    for i in range(len(liste_EBE)):
        FCF=-liste_invest[i]+(liste_EBE[i]*(1-T)+T*liste_DA[i]-liste_var_BFR[i])
        liste_FCF.append(FCF)
    print(liste_FCF)
    
    'Calcul de la van'
    
    # VAN=np.npv(0.1,liste_FCF)
    # print(VAN)
    VAN=0
    liste_calcul_VAN=[]
    for i in range(len(liste_FCF)):
        VAN= liste_FCF[i]/((1+r)**i)
        liste_calcul_VAN.append(VAN)
    VAN=sum(liste_calcul_VAN)
    
    print(VAN)
    
    

def Presentfcf ():
    global window2
    window2 = Tk()
    window2.title("Free-Cash-Flow")
    window2.resizable(height = False, width = False )
    window2.config(background='#009DF7')
    window2.iconbitmap("LogoCY.ico")
    
    lst = [('Année 1', liste_FCF[0]),
            ('Année 2', liste_FCF[1]),
           ('Année 3', liste_FCF[2]),
           ('Année 4', liste_FCF[3]),
           ('Année 5', liste_FCF[4]),
           ('Année 6', liste_FCF[5]),
           ('Année 7', liste_FCF[6]),
           ('Année 8', liste_FCF[7]),
           ('Année 9', liste_FCF[8]),
           ('Année 10', liste_FCF[9]),
           ('Année 11', liste_FCF[10]),
           ('Année 12', liste_FCF[11]),
           ('Année 13', liste_FCF[12]),
           ('Année 14', liste_FCF[13]),
           ('Année 15', liste_FCF[14]),
           ('Année 16', liste_FCF[15]),
           ('Année 17', liste_FCF[16]),
           ('Année 18', liste_FCF[17]),
           ('Année 19', liste_FCF[18]),
           ('Année 20', liste_FCF[19]),
           ('Année 21', liste_FCF[20]),
           ('Année 22', liste_FCF[21]),
           ('Année 23', liste_FCF[22]),
           ('Année 24', liste_FCF[23]),
           ('Année 25', liste_FCF[24]),
           ('Année 26', liste_FCF[25]),
           ('Année 27', liste_FCF[26]),
           ('Année 28', liste_FCF[27]),
           ('Année 29', liste_FCF[28]),
           ('Année 30', liste_FCF[29]),]
           
    total_rows = len(lst)
    total_columns = len (lst[0])
    
    for i in range(total_rows):
        for j in range(total_columns):
            
            Tableau_fcf= Entry(window2, width=20, font=("Calibri", 12, 'bold'), bg='#009DF7',fg='white')
            Tableau_fcf.grid(row=i, column=j)
            Tableau_fcf.insert(END, lst[i][j])
    
    window2.mainloop()

def visualiserHypothèses ():
    global window4
    window4 = Tk()
    window4.title("Hypothèses")
    window4.resizable(height = False, width = False )
    window4.config(background='#009DF7')
    window4.iconbitmap("LogoCY.ico")
    
    Taux_Inflation = Tx_I.get()
    Tarif_VL = TVL.get()
    taux_croissance_VL=0.03
    passage_VL= 2000000
    Tarif_PL=3*Tarif_VL
    taux_croissance_PL=0.015
    passage_PL=1000000
    cout_fixe = CF.get()
    marge_cout_var=0.2
    DA=105000000/20
    T=33/100
    taux_BFR=5/100
    invest_par_an=105000000/3
    r=0.1
    
    Liste_hypothèses = [Tarif_VL, passage_VL, taux_croissance_VL, Tarif_PL,
                        passage_PL, taux_croissance_PL, Taux_Inflation,cout_fixe,
                        marge_cout_var, DA, T, taux_BFR, invest_par_an, r]
    
    LH = [('Tarif Véhicule Léger', Liste_hypothèses[0]),
            ('Nombre Passages VL', Liste_hypothèses[1]),
           ('Taux de croissance VL', Liste_hypothèses[2]),
           ('Tatif Poids Lourd', Liste_hypothèses[3]),
           ('Nombre de Passage PL', Liste_hypothèses[4]),
           ('Taux de croissance PL', Liste_hypothèses[5]),
           ('Taux Inflation', Liste_hypothèses[6]),
           ('Coût Fixe', Liste_hypothèses[7]),
           ('Marge Coût Variable', Liste_hypothèses[8]),
           ('Montant DA', Liste_hypothèses[9]),
           ('Taux Impôt', Liste_hypothèses[10]),
           ('Taux de BFR du CA', Liste_hypothèses[11]),
           ('Investissement par an', Liste_hypothèses[12]),
           ('Taux r', Liste_hypothèses[13])]
    
    total_rows = len(LH)
    total_columns = len (LH[0])
    
    for i in range(total_rows):
        for j in range(total_columns):
            Tableau_HP= Entry(window4, width=20, font=("Calibri", 12, 'bold'), bg='#009DF7',fg='white')
            Tableau_HP.grid(row=i, column=j)
            Tableau_HP.insert(END, LH[i][j])
    
    window4.mainloop()
    
  
    
    
def calculerVAN ():
    global VAN
    VAN=0
    liste_calcul_VAN=[]
    for i in range(len(liste_FCF)):
        VAN= liste_FCF[i]/((1+r)**i)
        liste_calcul_VAN.append(VAN)
    VAN=sum(liste_calcul_VAN)
    
    print(VAN)
    
    VAN_2 = str(VAN)
    with open("liste_VAN.txt", "a+") as fVAN:
        fVAN.write(VAN_2 + "\n")
        fVAN.close()
    
def presentVAN ():
    global window3
    window3 = Tk()
    window3.title("VAN")
    window3.resizable(height = False, width = False )
    window3.config(background='#009DF7')
    window3.iconbitmap("LogoCY.ico")
    
    label_VAN = Label(window3, text = VAN, font=("Calibri", 18), bg='white', fg='#009DF7', justify = 'center' )
    
    label_VAN.pack()
    
    window3.mainloop()
    
def computeVANinfloat ():
    Myfile = open("liste_VAN.txt", 'r')
    Valeur = Myfile.readlines()
    liste_VAN_str = list(Valeur)
    global Liste_VAN
    Liste_VAN = []
    for ITEM in liste_VAN_str:
        VAN_str = ITEM.strip()  #on retire le saut de ligne
        VAN_float = float(VAN_str)  #on convertit en décimal
        Liste_VAN.append(VAN_float)
    nbr_situation = len(Liste_VAN)
    nbr_situation1 = nbr_situation + 1
    global Liste_situation
    Liste_situation = []
    for i in range (1, nbr_situation1):
        Liste_situation.append(i)
    
def analyseVAN ():
    global root
    root = Tk()
    root.title("Analyse des VAN")
    fig = Figure(figsize=(6,4),dpi=100)
    Situation_x = Liste_situation
    VAN_y = Liste_VAN
    fig.add_subplot(111).plot(Situation_x, VAN_y)
       
    canvas = FigureCanvasTkAgg(fig, master = root)
    canvas.draw()
    canvas.get_tk_widget().pack(side= TOP, fill= BOTH, expand=1)
    
    toolbar = NavigationToolbar2Tk(canvas, root)
    toolbar.update()
    canvas.get_tk_widget().pack(side= TOP, fill = BOTH, expand = True)
    
    root.mainloop()

def toutfermer ():
    window.destroy()
    window2.destroy()
    window3.destroy()
    window4.destroy()
    root.destroy()


       
        
'______________________________Création de la fenêtre principal______________________________'

window = Tk()

Tx_I = DoubleVar()
TVL = DoubleVar()
CF = IntVar()

window.title("Investment Modeling")
window.geometry("1080x680")
window.resizable(0,0)
window.config(background='#009DF7')
window.iconbitmap("LogoCY.ico")
    
'______________________________Création des frames______________________________'
frame1 = Frame(window, bg='#009DF7',bd=1, relief=SUNKEN)
frame2 = Frame(window, bg='#009DF7',bd=1, relief=SUNKEN)
frame3 = Frame(window, bg='#009DF7',bd=1, relief=SUNKEN)
frame4 = Frame(window, bg='#009DF7',bd=1, relief=SUNKEN)
frame5 = Frame(window, bg='#009DF7',bd=1, relief=SUNKEN)
frame6 = Frame(window, bg='#009DF7',bd=1, relief=SUNKEN)
frame7 = Frame(window, bg='#009DF7',bd=1, relief=SUNKEN)
frame8 = Frame(window, bg='#009DF7',bd=1, relief=SUNKEN)
frame_button1 = Frame(window, bg='#009DF7', bd=1,relief=SUNKEN)
frame_button2 = Frame(window, bg='#009DF7', bd=1,relief=SUNKEN)
frame_button3 = Frame(window, bg='#009DF7', bd=1,relief=SUNKEN)
frame_button4 = Frame(window,bg='#009DF7', bd=1,relief=SUNKEN)
frame_button_quit = Frame(window,bg='#009DF7', bd=1,relief=SUNKEN)
    
'______________________________Création des Label______________________________'
Label_title = Label(frame1, text="Bienvenue dans notre modélisation d'investissement",font=("Calibri", 26, 'bold'), bg='#009DF7',fg='white', justify = 'center')
Label_Info1 = Label (frame2, text="Taux d'inflation", font=("Calibri", 20), bg='white', fg='#009DF7', justify = 'center', width = 25, height = 2)
Label_Info2 = Label (frame3, text="Prix véhicule léger par passage", font=("Calibri", 20), bg='white', fg='#009DF7', justify = 'center',  width = 25, height = 2)
Label_Info3 = Label (frame4, text="Coût fixe annuel", font=("Calibri", 20), bg='white', fg='#009DF7', justify = 'center',  width = 25, height = 2)
Label_mode_emploi = Label(frame8, text="Etape 1 : Rentrer les valeurs choisies \n Etape 2 : Calculer les Free-Cash-Flow \n Etape 3 : Calculer la Van \n Etape 4 : Après avoir calculé plusieurs VAN, réaliser une analyse des VAN",font=("Calibri", 16), bg='white', fg='#009DF7', justify = 'center',  width = 80, height = 4 )    
'______________________________Création des Entry______________________________'

Entry1 = Entry(frame5, font=("Calibri", 20), width = 12, textvariable = Tx_I, justify = 'center')
Entry2 = Entry(frame6, font=("Calibri", 20),  width = 12, textvariable = TVL, justify = 'center')
Entry3 = Entry(frame7, font=("Calibri", 20), width = 12, textvariable = CF, justify = 'center')

'______________________________Création bouton______________________________'

button1 = Button(frame_button1, text= "Calculer les Free-Cash-Flow", font=("Calibri", 12), bg='white',fg='#009DF7', command= lambda:[calculerfcf(), Presentfcf()], justify ='center', width = 25, height = 5)
button2 = Button(frame_button2, text="Calculer la VAN",font=("Calibri", 12), bg='white',fg='#009DF7', command= lambda:[calculerVAN(),presentVAN()] ,justify ='center', width = 25, height = 5)
button3 = Button(frame_button3, text="Analyse des VAN", font=("Calibri", 12), bg='white',fg='#009DF7', command= lambda:[computeVANinfloat(), analyseVAN()], justify ='center', width = 25, height = 5)
button4 = Button(frame_button4, text="Hypothèses", font=("Calibri", 12), bg='white',fg='#009DF7',justify ='center', width = 25, height = 5, command = visualiserHypothèses)
button_quit = Button(frame_button_quit, text="Appuyez ici pour tout fermer", font=("Calibri", 12), bg='white',fg='#009DF7',justify ='center', width = 25, height = 3, command = toutfermer)

Label_title.pack()
Label_Info1.pack()
Label_Info2.pack()
Label_Info3.pack()
Label_mode_emploi.pack()

Entry1.pack()
Entry2.pack()
Entry3.pack()

button1.pack()
button2.pack()
button3.pack()
button4.pack()
button_quit.pack()

frame1.place(x = '155', y= '0')
frame2.place(x = '90', y= '350')
frame3.place(x = '90', y= '450')
frame4.place(x = '90', y= '550')
frame5.place(x = '550', y= '375')
frame6.place(x = '550', y= '475')
frame7.place(x = '550', y= '575')
frame8.place(x='100', y='70')
frame_button1.place(x = '800', y= '450')
frame_button2.place(x='450', y='200')
frame_button3.place(x='800', y='200')
frame_button4.place(x='100', y='200')
frame_button_quit.place(x='800', y='600')


window.mainloop()



    

    

    
        
    