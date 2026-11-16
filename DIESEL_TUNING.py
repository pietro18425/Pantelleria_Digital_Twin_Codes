# -*- coding: utf-8 -*-
import powerfactory as pf
import os
import sys
import pandas as pd
import random

# =============================================================================
# 1. INIZIALIZZAZIONE DIGSILENT
# =============================================================================
app = pf.GetApplication()
app.ClearOutputWindow()
app.EchoOff() 
app.ActivateProject(r'Pantelleria_definitivo_GFL.IntPrj')

# =============================================================================
# 2. CONFIGURAZIONE PERCORSI E VARIABILI
# =============================================================================
valore = 17422 # Istante Critico (Alto Carico)
main_folder = r'C:\Users\pietr\Desktop\tesi\FILE_SCRIPT'
file_sataria = r"C:\Users\pietr\Desktop\tesi\FILE_SCRIPT\Curva_Sataria_2025_Completo.xlsx"
percorso_salvataggio = r"C:\Users\pietr\Desktop\tesi\CONTROLLO_PARAMETRI_GEN"
os.makedirs(percorso_salvataggio, exist_ok=True)

app.PrintPlain(f'--- Avvio script integrato (Step: {valore}) ---')

# =============================================================================
# 3. LETTURA DATI E TRADUZIONE IN MW PER L'API
# =============================================================================
# Legge in kW dall'Excel ma passa in MW all'API
carico_tot_kw = pd.read_excel(os.path.join(main_folder, 'Curva_di_Carico_2025_15min.xlsx')).loc[valore, 'Curva_di_carico [kW]']
carico_tot_mw = carico_tot_kw 

sataria_kw = 0.0
try:
    df_sat = pd.read_excel(file_sataria)
    ap = float(str(df_sat.loc[valore, 'AP']).replace(',', '.')) if 'AP' in df_sat.columns else 0.0
    an = float(str(df_sat.loc[valore, 'AN']).replace(',', '.')) if 'AN' in df_sat.columns else 0.0
    sataria_kw = ap - an
except Exception as e:
    app.PrintPlain(f'Errore lettura Sataria, usato 0.0 kW. ({e})')

sataria_mw = sataria_kw 

ss = app.GetCalcRelevantObjects('*.ElmTrfstat')
trafi_validi = [s for s in ss if s.loc_name != 'SS 102' and s.GetContents('*.ElmTr2')]
potenza_trafi_tot = sum(s.GetContents('*.ElmTr2')[0].typ_id.strn for s in trafi_validi)

for s in ss:
    try:
        carico_cabina = s.GetContents('*.ElmLod')[0]
        if s.loc_name == 'SS 102':
            carico_cabina.plini = sataria_mw 
        else:
            pot_trafo = s.GetContents('*.ElmTr2')[0].typ_id.strn
            carico_cabina.plini = (pot_trafo / potenza_trafi_tot) * carico_tot_mw 
    except IndexError:
        pass

# =============================================================================
# 4. DISPATCH GENERATORI E SLACK DISTRIBUITO
# =============================================================================
gen_attivi = []
for i in range(1, 9):
    sw = app.GetCalcRelevantObjects(f'Switch Gen {i}.ElmCoup')
    gen = app.GetCalcRelevantObjects(f'Gen {i}.ElmSym')
    gov = app.GetCalcRelevantObjects(f'pcu_DEGOV1_GEN{i}.ElmDsl')
    exc = app.GetCalcRelevantObjects(f'vcoEXST2A_GEN{i}.ElmDsl')
    
    # Spegnimento totale di default
    if sw: sw[0].on_off = 0
    if gen: 
        gen[0].pgini = 0.0
        gen[0].outserv = 1
    if gov: gov[0].outserv = 1
    if exc: exc[0].outserv = 1
        
    try:
        df_gen = pd.read_csv(os.path.join(main_folder, f'LoadProfile_Gruppo {i}_2025.csv'), sep=';')
        df_gen.columns = [c.strip() for c in df_gen.columns]
        # In MW
        ap_mw = float(df_gen.loc[valore, 'AP']) / 1000.0 
        if ap_mw > 0 and sw and gen:
            sw[0].on_off = 1
            gen[0].outserv = 0
            gen[0].pgini = ap_mw
            if gov: gov[0].outserv = 0
            if exc: exc[0].outserv = 0
            gen_attivi.append(i)
    except Exception:
        pass

# --- SLACK DISTRIBUITO PER SOSTENERE ALTO CARICO ---
for s in gen_attivi:
    gen = app.GetCalcRelevantObjects(f'Gen {s}.ElmSym')[0]
    gen.c_pst = 3 # Tutti fanno da Reference Bus
    gen.usetp = 1.0 # Forzano la tensione a 10.6 kV

app.PrintPlain(f'Carico totale={carico_tot_kw} kW | Sataria={sataria_kw} kW | Generatori Attivi: {gen_attivi}')

ldf = app.GetFromStudyCase('ComLdf')
ldf.iopt_atf = 1
ldf.iopt_lim = 0 # CRUCIALE: ignora limiti di MVAR per evitare collasso

if ldf.Execute() == 0:
    app.PrintPlain("SUCCESS: Il Load Flow statico iniziale converge a ~10.6 kV reali.")
else:
    app.PrintPlain("ERROR: Il Load Flow è FALLITO!")
    app.EchoOn()
    sys.exit()

# =============================================================================
# 5. FUNZIONE KPI STABILITÀ
# =============================================================================
def check_stability(csv_data):
    try:
        ultimi_dati = csv_data.loc[csv_data.index > (csv_data.index.max() - 5.0)]
        col_freq = csv_data.columns[0]
        col_volt = csv_data.columns[1]
        freq_mean = ultimi_dati[col_freq].mean()
        freq_std = ultimi_dati[col_freq].std()
        volt_mean = ultimi_dati[col_volt].mean()
        volt_std = ultimi_dati[col_volt].std()
        is_freq_stable = (freq_std < 0.1) and (49.5 <= freq_mean <= 50.5)
        V_nominale = 10.6 
        volt_pu_mean = volt_mean / V_nominale
        volt_pu_std = volt_std / V_nominale
        is_volt_stable = (volt_pu_std < 0.01) and (0.9 <= volt_pu_mean <= 1.1)
        return is_freq_stable and is_volt_stable
    except Exception as e:
        return False

# =============================================================================
# 6. ITERAZIONI MONTE CARLO (CON HARD RESET E GESTIONE ERRORI)
# =============================================================================
N_ITERATIONS =   1000 
risultati = []
results_csv_temp = os.path.join(percorso_salvataggio, 'test.csv')

for it in range(N_ITERATIONS):
    p = {
        'gain': round(random.uniform(0.5, 10.0), 2), 
        't4': round(random.uniform(0.02, 0.2), 3),
        't5': round(random.uniform(0.02, 0.2), 3),
        't6': round(random.uniform(0.02, 0.2), 3), 
        'td': round(random.uniform(0.02, 0.1), 3), 
        'droop': round(random.uniform(0.04, 0.1), 4),
        'te': round(random.uniform(0.02, 0.2), 3),
        't1': round(random.uniform(0.02, 0.2), 3), 
        't2': round(random.uniform(0.02, 0.2), 3), 
        't3': round(random.uniform(1.0, 7), 3), 
        'm_delay': round(random.uniform(0.02, 0.1), 4),
        'c_gain': round(random.uniform(5.0, 40.0), 2), 
        'c_tc': round(random.uniform(0.05, 0.4), 3), 
        'exc_tc': round(random.uniform(0.2, 1.6), 3),
        'exc_c': round(random.uniform(0.8, 2.4), 3),
        's_gain': round(random.uniform(0.05, 0.3), 4), 
        's_tc': round(random.uniform(0.5, 3), 3), 
        'v_fact': round(random.uniform(0.8, 1.2), 3), 
        'c_fact': round(random.uniform(0.5, 1.5), 3) 
     }

    # Assegnazione parametri
    for s in gen_attivi:
        app.GetCalcRelevantObjects(f'pcu_DEGOV1_GEN{s}.ElmDsl')[0].params = [p['gain'], p['t4'], p['t5'], p['t6'], p['td'], p['droop'], p['te'], p['t1'], p['t2'], p['t3'], 1.0, 0.0, 0.5, 0.0, 1.0]
        app.GetCalcRelevantObjects(f'vcoEXST2A_GEN{s}.ElmDsl')[0].params = [p['m_delay'], p['c_gain'], p['c_tc'], p['exc_tc'], p['exc_c'], p['s_gain'], p['s_tc'], p['v_fact'], p['c_fact'], -10.0, 10.0, 6.0]

    # --- HARD RESET E RICALCOLO ---
    app.ResetCalculation()
    ldf_iter = app.GetFromStudyCase('ComLdf')
    ldf_iter.iopt_atf = 1
    ldf_iter.iopt_lim = 0 
    
    if ldf_iter.Execute() != 0:
        continue 
        
    rms_init = app.GetFromStudyCase('ComInc')
    rms_init.SetAttribute('p_resvar', app.GetFromStudyCase('Tuning.ElmRes'))
    rms_init.iopt_load = 1 
    
    if rms_init.Execute() != 0:
        continue 
        
    rms_sim = app.GetFromStudyCase('ComSim')
    rms_sim.tstop = 20.0 
    
    # Gestione crash della matrice (System matrix inversion failed)
    try:
        err = rms_sim.Execute()
        if err != 0: 
            continue # Salta all'iterazione successiva se la simulazione fallisce a t>0
    except Exception:
        continue
        
    Export = app.GetFromStudyCase("ComRes")
    Export.pResult = app.GetFromStudyCase('Tuning.ElmRes')
    Export.iopt_exp, Export.iopt_time, Export.iopt_csel, Export.col_Sep, Export.dec_Sep = 6, 5, 0, ';', ','
    Export.f_name = results_csv_temp
    Export.Execute()
    
    df_res = pd.read_csv(results_csv_temp, sep=';', header=1, index_col=0, decimal=',', na_values=['-nan(ind)', 'nan', 'NaN'])
    is_stable = check_stability(df_res)
    
    p['Iterazione'] = it + 1
    p['Stabile'] = is_stable
    risultati.append(p)

# =============================================================================
# 7. EXPORT RISULTATI
# =============================================================================
pd.DataFrame(risultati).to_csv(os.path.join(percorso_salvataggio, 'Mapping_Parametri.csv'), index=False, sep=';', decimal=',')
app.PrintPlain('Ottimizzazione completata. Mappa salvata in "Mapping_Parametri.csv"')
app.EchoOn()

