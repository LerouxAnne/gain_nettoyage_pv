# -*- coding: utf-8 -*-
"""
Created on Mon Jun  7 14:09:03 2021

@author: aleroux
"""

#%% Import librairies

import json
import pandas as pd 
import datetime
from dateutil.relativedelta import relativedelta
import math
import matplotlib.pyplot as plt

from openpyxl import load_workbook

from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu

import matplotlib as mpl

mpl.rcParams.update(mpl.rcParamsDefault)

#%% Functions 

def filter_irr(irr, sat, std):
    """

    This function allows to filter the irradiation column and to replace the
    incorrec values by the satellite's ones.
    ----------
    irr : FLOAT. Represents the measured irradiation
    sat : FLOAT. Represents the satellite irradiation

    Returns
    -------
    FLOAT

    """
    if math.isnan(irr) or irr<15 or std==0:
        return sat
    return irr


#%% Parameters 

# Dossier général
MAIN_DIR = 'Z:\Business\OMEGA\Private\THEMES DIVERS\Photovoltaïque\calculs gains suite aux nettoyages'
print('Code PI du parc à étudier : ')
PARC = input()

# Dossiers du parc
INPUTS_DIR = MAIN_DIR+f"\\{PARC}\inputs\\"
OUTPUTS_DIR = MAIN_DIR+f"\\{PARC}\outputs\\"

with open(INPUTS_DIR+'param.txt') as f:
    param = json.load(f)

# Paramètres des calculs
P_NOM = param['P_nom']
IRR_NOM = param['IRR_nom']

# JOUR = param['JOUR']
JOUR = 28
DIFF_PR = param['DIFF_PR']

SAVEFIG = bool(param['SavePlot'].lower() == 'true')

# Paramètres de data
FORMAT = "%Y-%m-%d"
START_NETT = datetime.datetime.strptime(param['start_nett'], FORMAT)
STOP_NETT = datetime.datetime.strptime(param['stop_nett'], FORMAT)

MIN_DATE = START_NETT - relativedelta(days=JOUR) # Début étude
MAX_DATE = STOP_NETT + relativedelta(days=JOUR) # Fin étude

# Rapport word
table_gain_2 = []
table_gain_3 = []

#%% Import datas

# Import des données de puissance active
df_pwr =  pd.read_csv(INPUTS_DIR+f"{PARC}_Power_bis.csv",
                      index_col=['Date'],
                      parse_dates=['Date'],
                      encoding='utf-8', sep=';')
df_pwr.Power = df_pwr.Power.apply(lambda x:float(str(x).replace(',','.')))

# Import des données d'irradiation
df_irradiation = pd.read_csv(INPUTS_DIR+f"{PARC}_Irradiation.csv",
                     index_col=['Date'],
                     parse_dates=['Date'],
                     encoding='utf-8', sep=';')
df_irradiation.Irradiation = df_irradiation.Irradiation.apply(lambda x:float(str(x).replace(',','.')))

# Ecart type par rapport à la valeur précédente (données figées)
df_irradiation['rolling_std'] = df_irradiation['Irradiation'].rolling(2).std()

#%% Filter irradiation
try:
    # Import des données d'irradiation satellite
    df_sat = pd.read_csv(INPUTS_DIR+f"{PARC}_satellite.csv",
                          index_col=['Date'],
                         parse_dates=['Date'],
                         encoding='utf-8', sep=';')
    df_sat.Irradiation_satellite = df_sat.Irradiation_satellite.apply(lambda x:float(str(x).replace(',','.')))
    
    # On regroupe les données d'irradiation ensemble
    df_irradiation = df_irradiation.join(df_sat, how='outer')
    
    # On remplace les données d'irradiation par les données satellites
    df_irradiation['correct_irr'] = df_irradiation.apply(lambda x: filter_irr(x['Irradiation'],
                                                              x['Irradiation_satellite'],
                                                              x['rolling_std']),
                                      axis=1)
    # On ne garde que la colonne avec les valeurs corrects
    df_irradiation = df_irradiation.drop(['Irradiation','rolling_std', 'Irradiation_satellite'], axis=1)
    
except:
    df_irradiation['correct_irr'] = df_irradiation['Irradiation']
    df_irradiation = df_irradiation[df_irradiation.rolling_std>0]
    df_irradiation = df_irradiation.drop(['Irradiation','rolling_std'], axis=1)

#%% Filtrage de l'irradiation 
# df_irradiation = df_irradiation[(df_irradiation.correct_irr>550)&(df_irradiation.correct_irr<=650)]
df_irradiation = df_irradiation.dropna()
#%% Both dataset have the same number of points

# On vérifie qu'il y a le même nombre de données de puissance et d'irradiation
df_join = df_pwr.join(df_irradiation, how='inner')

df_pwr = df_join.copy()
df_pwr = df_pwr.drop(['correct_irr'], axis=1)
df_pwr.reset_index(inplace=True)

df_irr = df_join.copy()
df_irr = df_irr.drop(['Power'], axis=1)
df_irr.reset_index(inplace=True)


#%% Filter datetime

# Création des dataset avant nettoyage
df_pwr_bfr = df_pwr[(df_pwr.Date>= MIN_DATE) & (df_pwr.Date < START_NETT)]
df_pwr_aft = df_pwr[(df_pwr.Date>= STOP_NETT) & (df_pwr.Date < MAX_DATE)]

# Création des dataset après nettoyage
df_irr_bfr = df_irr[(df_irr.Date>= MIN_DATE) & (df_irr.Date < START_NETT)]
df_irr_aft = df_irr[(df_irr.Date>= STOP_NETT) & (df_irr.Date < MAX_DATE)]

df_pwr_day = df_pwr[(df_pwr.Date>= MIN_DATE) & (df_pwr.Date < MAX_DATE)]

#%% Calcul des PR - part 1

# Calcul intermédiaire : on commence par calculer les sommes de chaque
# indicateurs par jour pour les datasets avant/après nettoyage
df_pwr_bfr_day = df_pwr_bfr.groupby(pd.Grouper(key='Date',freq='D')).sum()/P_NOM
df_irr_bfr_day = df_irr_bfr.groupby(pd.Grouper(key='Date',freq='D')).sum()/IRR_NOM

df_bfr_day = df_pwr_bfr_day.join(df_irr_bfr_day)

df_pwr_aft_day = df_pwr_aft.groupby(pd.Grouper(key='Date',freq='D')).sum()/P_NOM
df_irr_aft_day = df_irr_aft.groupby(pd.Grouper(key='Date',freq='D')).sum()/IRR_NOM

df_aft_day = df_pwr_aft_day.join(df_irr_aft_day)

df_pwr_day = df_pwr_day.groupby(pd.Grouper(key='Date',freq='D')).sum()/P_NOM
df_irr_day = df_irr.groupby(pd.Grouper(key='Date',freq='D')).sum()/IRR_NOM

df_pwr_day = df_pwr_day.join(df_irr_day)

#%% Calcul des PR - part 2

# Calcul des PR en faisant le rapport entre les deux colonnes précédemment 
# calculées
df_bfr_day['PR'] = df_bfr_day.Power/df_bfr_day.correct_irr

df_bfr_filter = df_bfr_day[abs(df_bfr_day.PR-df_bfr_day.PR.mean())<=DIFF_PR]

df_aft_day['PR'] = df_aft_day.Power/df_aft_day.correct_irr

df_aft_filter = df_aft_day[abs(df_aft_day.PR-df_aft_day.PR.mean())<=DIFF_PR]

df_pwr_day['PR'] = df_pwr_day.Power/df_pwr_day.correct_irr

df_pwr_filter = df_pwr_day[abs(df_pwr_day.PR-df_pwr_day.PR.mean())<=DIFF_PR]

#%% Result

# Calcul du gain 
GAIN_1 = (df_aft_filter.PR.mean() / df_bfr_filter.PR.mean() - 1) * 100
print(GAIN_1)

#%% Outputs
FORMAT_PREFIX = "%y%m%d"
OUTPUT_PREFIX = f'{PARC}_'\
    +datetime.datetime.strftime(MIN_DATE,FORMAT_PREFIX)+'_'\
    +datetime.datetime.strftime(MAX_DATE,FORMAT_PREFIX)+'_'

# Plot des PR
plt.figure()
plt.title('PR avant nettoyage')
df_bfr_day.PR.plot(label='brut', linestyle='None', marker='.')
df_bfr_filter.PR.plot(label='filtré', linestyle='None', marker='.')
plt.ylabel("PR")
plt.legend()
if SAVEFIG:
    plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR_avant_nettoyage.png',
                bbox_inches='tight')
plt.show()

plt.figure()
plt.title('PR avant nettoyage')
df_bfr_filter.PR.plot(label='filtré', linestyle='None', marker='.')
plt.ylabel("PR")
plt.legend()
if SAVEFIG:
    plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR_filtré_avant_nettoyage.png',
                bbox_inches='tight')
plt.show()

plt.figure()
plt.title('PR après nettoyage')
df_aft_day.PR.plot(label='brut', linestyle='None', marker='.')
df_aft_filter.PR.plot(label='filtré', linestyle='None', marker='.')
plt.ylabel("PR")
plt.legend()
if SAVEFIG:
    plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR_après_nettoyage.png',
                bbox_inches='tight')
plt.show()

plt.figure()
plt.title('PR après nettoyage')
df_aft_filter.PR.plot(label='filtré', linestyle='None', marker='.')
plt.ylabel("PR")
plt.legend()
if SAVEFIG:
    plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR_filtré_après_nettoyage.png',
                bbox_inches='tight')
plt.show()

plt.figure()
plt.title(f'{PARC} - PR')
df_pwr_day.PR.plot(label='brut', linestyle='None', marker='.')
df_pwr_filter.PR.plot(label='filtré', linestyle='None', marker='.')
plt.ylabel("PR")
plt.legend()
plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR.png',
            bbox_inches='tight')


plt.figure()
plt.title(f'{PARC} - PR filtré')
df_pwr_filter.PR.plot(label='filtré', linestyle='None', marker='.')
plt.ylabel("PR")
plt.legend()
plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR_filtré.png',
            bbox_inches='tight')

#%% Param SJB

SJB_sales = param['SJB_sale']
SJB_propres = param['SJB_propres']
SJB_tot = SJB_sales+SJB_propres
SJB_p_nom = param['SJB_P_nom']
# SJB_p_nom = [1,1,1,1]


#%% Import dirty SJB datas
temp_SJB_sales = SJB_sales[0]
df_pwr_SJB =  pd.read_csv(INPUTS_DIR+f"{PARC}_{temp_SJB_sales}_Power.csv",
                          index_col=['Date'],
                          parse_dates=['Date'],
                          encoding='utf-8', sep=';', decimal=',')

df_pwr_SJB.Power = df_pwr_SJB.Power.apply(lambda x:x if x>0 else 0)
df_pwr_SJB.Power = df_pwr_SJB.Power/SJB_p_nom[0]
df_pwr_SJB = df_pwr_SJB[df_pwr_SJB.Power<=1]

#%% Import clean SJB datas
temp_list = []
ind = 1

for filename in [INPUTS_DIR+f"{PARC}_{x}_Power.csv" for x in SJB_tot[1:]]:
    df = pd.read_csv(filename,
                      index_col=['Date'],
                      parse_dates=['Date'],
                      encoding='utf-8', sep=';', decimal=',')
    df.Power = df.Power.apply(lambda x:x if x>0 else 0)
    df.Power = df.Power/SJB_p_nom[ind]
    df_pwr_SJB = df_pwr_SJB.join(df, rsuffix='-'+SJB_tot[ind])
    ind+=1

df_pwr_SJB = df_pwr_SJB.rename(columns={"Power": 'Power-'+SJB_sales[0]})

#%% Nettoyage des données
# TO DO
for SJB in SJB_tot:
    df_pwr_SJB['rolling_std'] = df_pwr_SJB['Power-'+SJB].rolling(2).std()
    df_pwr_SJB = df_pwr_SJB[df_pwr_SJB.rolling_std>0]
    df_pwr_SJB = df_pwr_SJB.drop(['rolling_std'], axis=1)

#%% Both dataset have the same number of points

# On vérifie qu'il y a le même nombre de données de puissance et d'irradiation
df_join_SJB = df_pwr_SJB.join(df_irradiation, how='inner')

df_SJB = df_join_SJB.copy()
df_SJB = df_SJB.drop(['correct_irr'], axis=1)
df_SJB.reset_index(inplace=True)



df_irr_SJB = df_join_SJB.copy()
df_irr_SJB = df_irr_SJB[['correct_irr']]
df_irr_SJB.reset_index(inplace=True)

df_irr_SJB = df_irr_SJB[~df_irr_SJB.isnull()]


#%% Filter datetime

# Création des dataset avant nettoyage
df_SJB_bfr = df_SJB[(df_SJB.Date>= MIN_DATE) & (df_SJB.Date < START_NETT)]
df_SJB_aft = df_SJB[(df_SJB.Date>= STOP_NETT) & (df_SJB.Date < MAX_DATE)]

# Création des dataset après nettoyage
df_irr_SJB_bfr = df_irr_SJB[(df_irr_SJB.Date>= MIN_DATE) & (df_irr_SJB.Date < START_NETT)]
df_irr_SJB_aft = df_irr_SJB[(df_irr_SJB.Date>= STOP_NETT) & (df_irr_SJB.Date < MAX_DATE)]


#%% Calcul des PR - part 1

# Calcul intermédiaire : on commence par calculer les sommes de chaque
# indicateurs par jour pour les datasets avant/après nettoyage
df_SJB_bfr_day = df_SJB_bfr.groupby(pd.Grouper(key='Date',freq='D')).sum()
df_irr_SJB_bfr_day = df_irr_SJB_bfr.groupby(pd.Grouper(key='Date',freq='D')).sum()/IRR_NOM

df_SJB_bfr_day = df_SJB_bfr_day.join(df_irr_SJB_bfr_day)

df_SJB_aft_day = df_SJB_aft.groupby(pd.Grouper(key='Date',freq='D')).sum()
df_irr_SJB_aft_day = df_irr_SJB_aft.groupby(pd.Grouper(key='Date',freq='D')).sum()/IRR_NOM

df_SJB_aft_day = df_SJB_aft_day.join(df_irr_SJB_aft_day)

# Calcul total pour les graphes
df_SJB_day = df_SJB.groupby(pd.Grouper(key='Date',freq='D')).sum()
df_irr_SJB_day = df_irr_SJB.groupby(pd.Grouper(key='Date',freq='D')).sum()/IRR_NOM

#%% Calcul des PR - part 2

# Calcul des PR en faisant le rapport entre les deux colonnes précédemment 
# calculées
for SJB in SJB_tot:
    df_SJB_bfr_day['PR-'+SJB] = df_SJB_bfr_day['Power-'+SJB]/df_SJB_bfr_day.correct_irr
    df_SJB_bfr_day = df_SJB_bfr_day[abs(df_SJB_bfr_day['PR-'+SJB]\
                                        -df_SJB_bfr_day['PR-'+SJB].mean())<=DIFF_PR]

    df_SJB_aft_day['PR-'+SJB] = df_SJB_aft_day['Power-'+SJB]/df_SJB_aft_day.correct_irr
    df_SJB_aft_day = df_SJB_aft_day[abs(df_SJB_aft_day['PR-'+SJB]\
                                    -df_SJB_aft_day['PR-'+SJB].mean())<=DIFF_PR]
        
    # Calcul total pour les graphes
    df_SJB_day['PR-'+SJB] = df_SJB_day['Power-'+SJB]/df_irr_SJB_day.correct_irr
    df_SJB_day = df_SJB_day[abs(df_SJB_day['PR-'+SJB]\
                                -df_SJB_day['PR-'+SJB].mean())<=DIFF_PR]

#%% Ecart entre propres et sales
for SJB in range(len(SJB_propres)):
    SJB_p = SJB_propres[SJB]
    SJB_s = SJB_sales[SJB]
    gain_2 = 100*((df_SJB_aft_day['PR-'+SJB_p].mean()/df_SJB_bfr_day['PR-'+SJB_p].mean())\
          -(df_SJB_aft_day['PR-'+SJB_s].mean()/df_SJB_bfr_day['PR-'+SJB_s].mean()))
        
    print("GAIN 2 pour "+ SJB_p + ' et ' + SJB_s)
    # print(100*(df_SJB_aft_day['Ecart_'+SJB].mean()/df_SJB_bfr_day['Ecart_'+SJB].mean()-1))
    # print('PR propre bfr : '+str(df_SJB_bfr_day['PR-'+SJB_p].mean()))
    # print('PR propre aft : '+str(df_SJB_aft_day['PR-'+SJB_p].mean()))
    print('Ecart en % (propre): ' +str((df_SJB_aft_day['PR-'+SJB_p].mean()/df_SJB_bfr_day['PR-'+SJB_p].mean())*100-100))
    # print('PR sale bfr : '+str(df_SJB_bfr_day['PR-'+SJB_s].mean()))
    # print('PR sale aft : '+str(df_SJB_aft_day['PR-'+SJB_s].mean()))
    print("Ecart % (sale) : "+str((df_SJB_aft_day['PR-'+SJB_s].mean()/df_SJB_bfr_day['PR-'+SJB_s].mean())*100-100))
    
    print(gain_2)
    
    table_gain_2.append({'SJB_propre':SJB_p,
                        'SJB_sale':SJB_s,
                        'gain2':round(gain_2,2)})
    
print('Gain moyen :')
print(sum([x['gain2'] for x in table_gain_2])/len(table_gain_2))

#%% Calcul du gain 3

col_to_keep = ['PR']+['PR-'+SJB for SJB in SJB_sales]

df_gain3_bfr_day = df_bfr_day.join(df_SJB_bfr_day, rsuffix='_SJB')
df_gain3_bfr_day = df_gain3_bfr_day[col_to_keep]
df_gain3_bfr_day = df_gain3_bfr_day.dropna()

df_gain3_aft_day = df_aft_day.join(df_SJB_aft_day, rsuffix='_SJB')
df_gain3_aft_day = df_gain3_aft_day[col_to_keep]
df_gain3_aft_day = df_gain3_aft_day.dropna()

for SJB in SJB_sales:
    gain_3 = 100*((df_gain3_aft_day['PR'].mean()/df_gain3_bfr_day['PR'].mean())\
          -(df_gain3_aft_day['PR-'+SJB].mean()/df_gain3_bfr_day['PR-'+SJB].mean()))
    print(f'GAIN 3 par rapport au SJB : {SJB}')
    print(gain_3)
    table_gain_3.append({'SJB_sale':SJB,
                         'gain3':round(gain_3,2)})

#%% Outputs SJB

# Plot des PR_SJB avant
plt.figure()
plt.title(f'{PARC} - PR avant nettoyage')
for SJB in SJB_tot:
    if SJB in SJB_sales : 
        df_SJB_bfr_day['PR-'+SJB].plot(label=SJB+' (sale)', linestyle='None', marker='.')
    else:
        df_SJB_bfr_day['PR-'+SJB].plot(label=SJB, linestyle='None', marker='.')
plt.ylabel("PR")
plt.legend()
if SAVEFIG:
    plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR_SJB_avant_nettoyage.png',
                bbox_inches='tight')
plt.show()

# Plot des PR_SJB après
plt.figure()
plt.title(f'{PARC} - PR après nettoyage')
for SJB in SJB_tot:
    if SJB in SJB_sales : 
        df_SJB_aft_day['PR-'+SJB].plot(label=SJB+' (sale)', linestyle='None', marker='.')
    else:
        df_SJB_aft_day['PR-'+SJB].plot(label=SJB, linestyle='None', marker='.')
plt.ylabel("PR")
plt.legend()
if SAVEFIG:
    plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR_SJB_apres_nettoyage.png',
                bbox_inches='tight')
plt.show()

# Plot des PR_SJB totaux
plt.figure()
plt.title(f'{PARC} - PR')
for SJB in SJB_tot:
    if SJB in SJB_sales : 
        df_SJB_day['PR-'+SJB].plot(label=SJB+' (sale)', linestyle='None', marker='.')
    else:
        df_SJB_day['PR-'+SJB].plot(label=SJB, linestyle='None', marker='.')
plt.ylabel("PR")
plt.legend()
if SAVEFIG:
    plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR_SJB_apres_nettoyage.png',
                bbox_inches='tight')
plt.show()

#%% Plot rolling 
df_SJB_day_rolling = df_SJB_day.rolling(14).mean()

plt.figure()
plt.title(f'{PARC} - PR')

plt.axvline(x=STOP_NETT, linestyle='--',color='grey', label='Dates de nettoyage')
for SJB in SJB_tot:
    if SJB in SJB_sales : 
        df_SJB_day_rolling['PR-'+SJB].plot(label=SJB+' (sale)', linestyle='None', marker='.')
    else:
        df_SJB_day_rolling['PR-'+SJB].plot(label=SJB, linestyle='None', marker='.')
plt.axvline(x=START_NETT, linestyle='--',color='grey')
plt.ylabel("PR")
plt.legend(bbox_to_anchor=(0, 1, 1, 0),loc="lower left")
plt.savefig(OUTPUTS_DIR+OUTPUT_PREFIX+'PR_SJB_apres_nettoyage.png',
            bbox_inches='tight')
plt.show()


#%% Plot rolling normalized

# Set the default color cycle
# mpl.rcParams['axes.prop_cycle'] = mpl.cycler(color=['#b2182b','#d6604d','#f4a582','#92c5de','#4393c3','#2166ac']) 

df_SJB_day_rolling = df_SJB_day.rolling(15).mean().dropna()

plt.figure()
plt.title(f'{PARC} - PR')
plt.axvline(x=STOP_NETT, linestyle='--',color='grey', label='Dates de nettoyage')
for SJB in SJB_tot:
    if SJB in SJB_sales : 
        (df_SJB_day_rolling['PR-'+SJB]-df_SJB_day_rolling['PR-'+SJB].iloc[13]).plot(label=SJB+' (sale)', linestyle='None', marker='.')
    else:
        (df_SJB_day_rolling['PR-'+SJB]-df_SJB_day_rolling['PR-'+SJB].iloc[13]).plot(label=SJB, linestyle='None', marker='.')
plt.axvline(x=START_NETT, linestyle='--',color='grey')
plt.axvline(x=STOP_NETT, linestyle='--',color='grey')
plt.ylabel("PR")
plt.legend(bbox_to_anchor=(0, 1, 1, 0),loc="lower left")
plt.show()

#%% Save data to word
#Import template document
template = DocxTemplate('template_nettoyage.docx')

# Mois en français
french_month = {1:'janvier',
                2:'février',
                3:'mars',
                4:'avril',
                5:'mai',
                6:'juin',
                7:'juillet',
                8:'août',
                9:'septembre',
                10:'octobre',
                11:'novembre',
                12:'décembre'}

    
#Import saved figure
image_PR_centrale = InlineImage(template, OUTPUTS_DIR+OUTPUT_PREFIX+
                                       'PR.png',Cm(12))
image_PR_centrale_filtre = InlineImage(template, OUTPUTS_DIR+OUTPUT_PREFIX+
                                       'PR_filtré.png',Cm(12))
image_PR_SJB = InlineImage(template, OUTPUTS_DIR+OUTPUT_PREFIX+
                           'PR_SJB_apres_nettoyage.png',Cm(12))


#Declare template variables
context = {
    'nom_parc':PARC,
    'gain1':round(GAIN_1,2),
    'table_gain2':table_gain_2,
    'table_gain3':table_gain_3,
    'moy_gain2':round(sum([l['gain2'] for l in table_gain_2])/len(table_gain_2),2),
    'moy_gain3':round(sum([l['gain3'] for l in table_gain_3])/len(table_gain_3),2),
    'start_date': french_month[MIN_DATE.month]+MIN_DATE.strftime(' %Y'),
    'stop_date': french_month[MAX_DATE.month]+MAX_DATE.strftime(' %Y'),
    'image_PR_centrale':image_PR_centrale,
    'image_PR_centrale_filtre':image_PR_centrale_filtre,
    'image_PR_SJB':image_PR_SJB
    }

#Render automated report
template.render(context)
template.save(OUTPUTS_DIR+f'{PARC}_rapport_M+10.docx')

#%% Save Data to excel
# try :
#     wb = load_workbook("result_script.xlsx")
#     ws = wb.active
#     ws.append([PARC,
#                GAIN_1/100,
#                df_bfr_filter.PR.mean(),
#                df_aft_filter.PR.mean(),
#                datetime.datetime.strftime(datetime.datetime.now(),FORMAT)])
#     wb.save("result_script.xlsx")
# except:
#     print('Le fichier excel est déjà ouvert')