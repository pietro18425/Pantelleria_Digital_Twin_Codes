import powerfactory as pf
import os
import pandas as pd

app = pf.GetApplication()
app.ClearOutputWindow()
app.EchoOff()
app.ActivateProject(r'Pantelleria_definitivo_GFL.IntPrj')

valore = 17422

main_folder = r'C:\Users\pietr\Desktop\tesi\FILE_SCRIPT'

# =====================================================
# FILE INPUT
# =====================================================
carico_path = os.path.join(
    main_folder,
    'Curva_di_carico_2025_15min.csv'
)

generatori_path = os.path.join(
    main_folder,
    'Generators_Active_Count_2025.xlsx'
)

file_sataria = os.path.join(
    main_folder,
    'Curva_Sataria_2025_Completo.xlsx'
)

# =====================================================
# LETTURA FILE CARICO
# =====================================================
carico = pd.read_csv(
    carico_path,
    sep=';',
    decimal=',',
    encoding='utf-8',
    engine='python'
)

carico.columns = carico.columns.astype(str).str.strip()
app.PrintPlain(f"Colonne carico: {list(carico.columns)}")

# Nome colonna carico
if 'Curva_di_carico [kW]' in carico.columns:
    nome_colonna_carico = 'Curva_di_carico [kW]'
elif 'Load_kW' in carico.columns:
    nome_colonna_carico = 'Load_kW'
else:
    nome_colonna_carico = carico.columns[-1]
    app.PrintPlain(
        f"ATTENZIONE: uso automaticamente la colonna "
        f"'{nome_colonna_carico}'"
    )

# Controllo indice
if valore < 0 or valore >= len(carico):
    raise IndexError(
        f"Indice valore={valore} fuori intervallo. "
        f"Righe disponibili: {len(carico)}"
    )

# =====================================================
# LETTURA FILE GENERATORI
# =====================================================
generatori = pd.read_excel(generatori_path)

# =====================================================
# OGGETTI POWERFACTORY
# =====================================================
ss = app.GetCalcRelevantObjects('*.ElmTrfstat')

# Cabine da escludere dalla ripartizione
esclusi = {'SS 102', 'BESS'}

# =====================================================
# LETTURA FILE SATARIA (SS 102)
# AP > 0 => prelievo => carico positivo
# AN > 0 => rilascio => carico negativo
# =====================================================
sataria_kw = 0.0

try:
    sataria = pd.read_excel(file_sataria)
    sataria.columns = [str(c).strip() for c in sataria.columns]

    for col in ['AP', 'AN']:
        if col in sataria.columns:
            sataria[col] = (
                sataria[col]
                .astype(str)
                .str.replace(',', '.', regex=False)
                .str.strip()
            )
            sataria[col] = pd.to_numeric(
                sataria[col],
                errors='coerce'
            ).fillna(0.0)
        else:
            sataria[col] = 0.0

    ap_kw = float(sataria.iloc[valore]['AP'])
    an_kw = float(sataria.iloc[valore]['AN'])

    if ap_kw > 0 and an_kw <= 0:
        sataria_kw = ap_kw
    elif an_kw > 0 and ap_kw <= 0:
        sataria_kw = -an_kw
    elif ap_kw > 0 and an_kw > 0:
        sataria_kw = ap_kw - an_kw
    else:
        sataria_kw = 0.0

    app.PrintPlain(
        f'SS 102 da file Sataria: '
        f'AP={ap_kw} kW, AN={an_kw} kW, netto={sataria_kw} kW'
    )

except Exception as e:
    app.PrintPlain(f'Errore lettura file Sataria: {e}')
    sataria_kw = 0.0

# =====================================================
# POTENZA TOTALE TRAFI ESCLUDENDO SS 102 E BESS
# =====================================================
potenza_trafi = 0.0

for s in ss:
    try:
        nome = str(s.loc_name).strip()

        if nome in esclusi:
            app.PrintPlain(f'{nome} esclusa dal calcolo potenza trafi')
            continue

        trasformatore = s.GetContents('*.ElmTr2')[0]
        potenza_apparente_trafo = float(trasformatore.typ_id.strn)

        app.PrintPlain(f'{nome}: {potenza_apparente_trafo}')
        potenza_trafi += potenza_apparente_trafo

    except IndexError:
        app.PrintPlain(f'{s.loc_name}')
        pass

app.PrintPlain(
    f'Potenza totale trafi escluse {sorted(esclusi)} = {potenza_trafi}'
)

# =====================================================
# CARICO TOTALE DI SISTEMA
# =====================================================
carico_totale = float(carico.iloc[valore][nome_colonna_carico])

app.PrintPlain(f'Carico totale = {carico_totale} kW')

# =====================================================
# RIPARTIZIONE CARICO SU TUTTE LE CABINE
# ESCLUSE SS 102 E BESS
# =====================================================
somma_carichi_ripartiti = 0.0

for s in ss:
    try:
        nome = str(s.loc_name).strip()

        if nome in esclusi:
            app.PrintPlain(
                f'{nome} esclusa dalla ripartizione del carico totale'
            )
            continue

        carico_cabina = s.GetContents('*.ElmLod')[0]
        trasformatore = s.GetContents('*.ElmTr2')[0]

        ratio = float(trasformatore.typ_id.strn) / potenza_trafi
        carico_assegnato = ratio * carico_totale

        carico_cabina.plini = carico_assegnato
        somma_carichi_ripartiti += carico_assegnato

        app.PrintPlain(
            f'{nome}: ratio={ratio:.9f}, '
            f'plini={carico_assegnato:.6f} kW'
        )

    except IndexError:
        app.PrintPlain(f'{s.loc_name}')
        pass

app.PrintPlain(
    f'Somma carichi ripartiti escluse {sorted(esclusi)} = '
    f'{somma_carichi_ripartiti:.6f} kW'
)

# =====================================================
# IMPOSTAZIONE SS 102 DA FILE
# positivo = prelievo
# negativo = immissione
# =====================================================
ss102_trovata = False

for s in ss:
    try:
        if str(s.loc_name).strip() == 'SS 102':
            carico_ss102 = s.GetContents('*.ElmLod')[0]
            carico_ss102.plini = sataria_kw
            ss102_trovata = True

            app.PrintPlain(
                f'SS 102 impostata da file con plini = {sataria_kw} kW'
            )
            break

    except IndexError:
        app.PrintPlain('SS 102 trovata ma senza ElmLod')
        pass

if not ss102_trovata:
    app.PrintPlain(
        'ATTENZIONE: SS 102 non trovata oppure nessun ElmLod associato'
    )

# =====================================================
# NUMERO GENERATORI ATTIVI
# =====================================================
if 'Numero Generatori Attivi' in generatori.columns:
    numero_gen_attivi = generatori.iloc[valore]['Numero Generatori Attivi']
    app.PrintPlain(f'Numero generatori attivi da file = {numero_gen_attivi}')
else:
    app.PrintPlain(
        "ATTENZIONE: colonna 'Numero Generatori Attivi' non trovata"
    )
    numero_gen_attivi = None

# =====================================================
# LETTURA PROFILI AP GENERATORI
# =====================================================
profili_ap = {}

for i in range(1, 9):
    file_gen = os.path.join(
        main_folder,
        f'LoadProfile_Gruppo {i}_2025.csv'
    )

    try:
        df = pd.read_csv(
            file_gen,
            sep=';',
            engine='python'
        )
        df.columns = [str(c).strip() for c in df.columns]

        if 'AP' not in df.columns:
            app.PrintPlain(
                f'ATTENZIONE: colonna AP non trovata in {file_gen}'
            )
            app.PrintPlain(f'Colonne trovate: {list(df.columns)}')
            profili_ap[i] = 0.0
            continue

        ap_kw = float(df.iloc[valore]['AP'])
        profili_ap[i] = ap_kw / 1000.0

        app.PrintPlain(f'Gruppo {i}: AP = {profili_ap[i]:.3f} MW')

    except Exception as e:
        app.PrintPlain(f'Errore lettura file gruppo {i}: {e}')
        profili_ap[i] = 0.0

# =====================================================
# RESET SWITCH E GENERATORI
# =====================================================
for i in range(1, 9):
    sw_list = app.GetCalcRelevantObjects(f'Switch Gen {i}.ElmCoup')
    if len(sw_list) > 0:
        sw_list[0].on_off = 0

    gen_list = app.GetCalcRelevantObjects(f'Gen {i}.ElmSym')
    if len(gen_list) > 0:
        gen = gen_list[0]

        try:
            gen.ip_ctrl = 0
        except:
            pass

        try:
            gen.pgini = 0.0
        except:
            pass

# =====================================================
# ATTIVAZIONE IN BASE AD AP
# =====================================================
gen_attivi = []

for i in range(1, 9):
    ap_mw = profili_ap[i]

    sw_list = app.GetCalcRelevantObjects(f'Switch Gen {i}.ElmCoup')
    gen_list = app.GetCalcRelevantObjects(f'Gen {i}.ElmSym')

    if len(sw_list) == 0:
        app.PrintPlain(f'Switch Gen {i} non trovato')
        continue

    if len(gen_list) == 0:
        app.PrintPlain(f'Gen {i} non trovato')
        continue

    sw = sw_list[0]
    gen = gen_list[0]

    if ap_mw > 0:
        sw.on_off = 1
        gen_attivi.append(i)

        try:
            gen.pgini = ap_mw
        except:
            app.PrintPlain(f'Impossibile impostare pgini per Gen {i}')

# =====================================================
# VERIFICA FINALE
# =====================================================
somma_dispatch = sum(profili_ap[i] for i in gen_attivi)

app.PrintPlain(f'Generatori attivi: {gen_attivi}')
app.PrintPlain(f'Somma dispatch iniziale = {somma_dispatch:.3f} MW')

app.EchoOn()