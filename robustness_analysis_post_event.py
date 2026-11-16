

import sys
import os
import csv
import math
import traceback

sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2025 SP4\Python\3.13')
import powerfactory as pf


# =============================================================================
# CONFIGURAZIONE
# =============================================================================

PROJECT_NAME = r'Pantelleria_definitivo_spiller_con_GFL.IntPrj'

CSV_FILE = r"C:\Users\pietr\Desktop\tesi\CONTROLLO_PARAMETRI_GEN\Mapping_Parametri_GFL.csv"
OUTPUT_FOLDER = r"C:\Users\pietr\Desktop\tesi\CONTROLLO_PARAMETRI_GEN"
OUTPUT_CSV = os.path.join(OUTPUT_FOLDER, 'Risultati_Sweep.csv')

T_START = -100.0
EVENT_TIME = 100.0
T_STOP = 300.0

# Usato solamente per ComSim; NON per costruire il vettore del tempo.
DT = 0.05

GENERATOR_NAME = 'Gen 2'
SWITCH_NAME = 'Switch Gen 2'
TERMINAL_NAME = 'CENTRALE'
N_GENERATORS = 8

# =============================================================================
# CRITERIO DI ASSESTAMENTO
# =============================================================================

# Media degli ultimi 10 s per stimare il nuovo equilibrio.
SETTLING_WINDOW = 10.0

# Fascia di assestamento: f_eq +/- 0.01 Hz.
SETTLING_BAND_HZ = 0.01

# Permanenza minima nella fascia.
SETTLING_HOLD = 5.0

# Stampa l'asse temporale letto da ElmRes.
DEBUG_TIME_AXIS = True


# =============================================================================
# PARAMETRI CSV / DSL
# =============================================================================

GOV_COLUMNS = [
    'gain', 't4', 't5', 't6', 'td',
    'droop', 'te', 't1', 't2', 't3'
]

EXC_COLUMNS = [
    'm_delay', 'c_gain', 'c_tc',
    'exc_tc', 'exc_c', 's_gain',
    's_tc', 'v_fact', 'c_fact'
]

GOV_EXTRA = [1.0, 0.0, 0.5, 0.0, 1.0]
EXC_EXTRA = [-10.0, 10.0, 6.0]

RESULT_COLUMNS = [
    'Iterazione',
    'Stabile_CSV',
    'Frequenza_riferimento_Hz',
    'Frequenza_equilibrio_Hz',
    'Delta_f_max_Hz',
    'Delta_f_max_percento',
    'Tempo_assestamento_s',
    'Esito',
    'Errore'
]


# =============================================================================
# FUNZIONI GENERALI
# =============================================================================

def pf_print(app, text):
    app.PrintPlain(str(text))
    print(str(text))


def safe_delete(obj):
    if obj is None:
        return

    try:
        obj.Delete()
    except Exception:
        pass


def parse_number(value):
    """Converte valori con separatore decimale italiano o inglese."""
    text = str(value).strip()

    if text == '':
        raise ValueError('Valore numerico vuoto.')

    # Esempio: 1.234,56 -> 1234.56
    if ',' in text and '.' in text:
        text = text.replace('.', '').replace(',', '.')

    # Esempio: 0,0413 -> 0.0413
    elif ',' in text:
        text = text.replace(',', '.')

    return float(text)


def is_true(value):
    """
    Restituisce True solo quando il campo Stabile del CSV indica
    esplicitamente che l'iterazione deve essere eseguita.

    Sono accettati:
    True, true, TRUE, 1, yes, y, si, sì.
    """
    return str(value).strip().lower() in (
        'true',
        '1',
        'yes',
        'y',
        'si',
        'sì'
    )


def mean(values):
    if not values:
        return float('nan')

    return sum(values) / float(len(values))


def scalar(value):
    """
    ResGetData può restituire float, tuple o list.
    Per tuple/list prende l'ultimo elemento.
    """
    if isinstance(value, tuple):
        value = value[-1]

    if isinstance(value, list):
        value = value[-1]

    return float(value)


def csv_value(value):
    """Converte float in formato europeo per il CSV di output."""
    if value is None:
        return ''

    if isinstance(value, float):
        if not math.isfinite(value):
            return ''

        return '{:.8f}'.format(value).replace('.', ',')

    return str(value)


# =============================================================================
# LETTURA MAPPING PARAMETRI
# =============================================================================

def read_csv_rows(path):
    rows = []

    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f, delimiter=';')

        if not reader.fieldnames:
            raise RuntimeError('CSV senza intestazione.')

        required_columns = GOV_COLUMNS + EXC_COLUMNS + ['Stabile']

        missing = [
            column
            for column in required_columns
            if column not in reader.fieldnames
        ]

        if missing:
            raise RuntimeError(
                'Colonne mancanti nel CSV: {}'.format(
                    ', '.join(missing)
                )
            )

        for line_no, raw in enumerate(reader, start=2):

            if not any((value or '').strip() for value in raw.values()):
                continue

            row = dict(raw)

            for column in GOV_COLUMNS + EXC_COLUMNS:

                raw_value = (row.get(column) or '').strip()

                if raw_value == '':
                    raise RuntimeError(
                        'Riga {}: valore vuoto nella colonna {}.'.format(
                            line_no,
                            column
                        )
                    )

                row[column] = parse_number(raw_value)

            row['_line'] = line_no
            rows.append(row)

    return rows


# =============================================================================
# OGGETTI POWERFACTORY
# =============================================================================

def get_one(app, pattern, label):
    found = app.GetCalcRelevantObjects(pattern)

    if not found:
        raise RuntimeError(
            '{} non trovato: {}'.format(label, pattern)
        )

    return found[0]


def get_study_object(app, name):
    obj = app.GetFromStudyCase(name)

    if obj is None:
        raise RuntimeError(
            'Oggetto Study Case non trovato: {}'.format(name)
        )

    return obj


# =============================================================================
# ASSEGNAZIONE PARAMETRI DSL
# =============================================================================

def set_params_from_row(app, row):
    """
    Governor:
        10 valori dal CSV + 5 valori fissi = 15 parametri.

    Exciter:
        9 valori dal CSV + 3 limiti fissi = 12 parametri.
    """
    gov_values = [row[column] for column in GOV_COLUMNS]
    exc_values = [row[column] for column in EXC_COLUMNS]

    gov_params = gov_values + GOV_EXTRA
    exc_params = exc_values + EXC_EXTRA

    for generator_number in range(1, N_GENERATORS + 1):

        gov_list = app.GetCalcRelevantObjects(
            'pcu_DEGOV1_GEN{}.ElmDsl'.format(generator_number)
        )

        exc_list = app.GetCalcRelevantObjects(
            'vcoEXST2A_GEN{}.ElmDsl'.format(generator_number)
        )

        if not gov_list:
            raise RuntimeError(
                'Governor Gen {} non trovato.'.format(
                    generator_number
                )
            )

        if not exc_list:
            raise RuntimeError(
                'Exciter Gen {} non trovato.'.format(
                    generator_number
                )
            )

        # Stesso metodo del tuo script di assegnazione funzionante.
        gov_list[0].params = gov_params
        exc_list[0].params = exc_params


# =============================================================================
# SWITCH E EVENTI
# =============================================================================

def close_switch(switch_obj):
    """
    Chiude Switch Gen 2.

    Nel modello:
        on_off = 1 -> breaker chiuso
        on_off = 0 -> breaker aperto
    """
    try:
        switch_obj.on_off = 1
        return
    except Exception:
        pass

    try:
        switch_obj.isclosed = 1
        return
    except Exception:
        pass

    raise RuntimeError(
        'Impossibile chiudere {}'.format(
            switch_obj.loc_name
        )
    )


def set_generator_in_service(generator):
    try:
        generator.outserv = 0
    except Exception:
        pass


def reset_switch_and_gen(generator, switch_obj):
    """Ripristina Gen 2 in servizio e il suo breaker chiuso."""
    set_generator_in_service(generator)
    close_switch(switch_obj)


def clear_all_switch_events(app):
    """
    Elimina tutti gli EvtSwitch dalla cartella IntEvt
    per evitare che rimangano eventi duplicati tra iterazioni.
    """
    event_folder = app.GetFromStudyCase('IntEvt')

    if event_folder is None:
        return

    try:
        events = event_folder.GetContents()
    except Exception:
        events = []

    for event in events:
        try:
            if event.GetClassName() == 'EvtSwitch':
                event.Delete()
        except Exception:
            pass


def create_disconnect_event(app, switch_obj):
    """
    Crea EvtSwitch che apre Switch Gen 2 a EVENT_TIME.
    """
    event_folder = app.GetFromStudyCase('IntEvt')

    if event_folder is None:
        raise RuntimeError('Cartella IntEvt non trovata.')

    event = event_folder.CreateObject(
        'EvtSwitch',
        'Trip_Gen2_100s'
    )

    if event is None:
        raise RuntimeError('Impossibile creare EvtSwitch.')

    event.time = EVENT_TIME
    event.p_target = switch_obj
    event.i_switch = 0

    return event


# =============================================================================
# RESULT FILE E RMS
# =============================================================================

def create_result_file(app, terminal):
    """
    Crea un ElmRes temporaneo per l'iterazione corrente
    e registra CENTRALE:m:fehz.
    """
    result_folder = get_study_object(app, 'Results')

    try:
        old_results = result_folder.GetContents(
            'BatchResults.ElmRes'
        )
    except Exception:
        old_results = []

    for old_result in old_results:
        safe_delete(old_result)

    result_file = result_folder.CreateObject(
        'ElmRes',
        'BatchResults'
    )

    if isinstance(result_file, (list, tuple)):
        result_file = result_file[0]

    if result_file is None:
        raise RuntimeError(
            'Creazione BatchResults.ElmRes fallita.'
        )

    result_file.AddVars(terminal, 'm:fehz')

    return result_file


def configure_commands(app, result_file):
    ldf = get_study_object(app, 'ComLdf')
    ini = get_study_object(app, 'ComInc')
    sim = get_study_object(app, 'ComSim')

    try:
        ldf.iopt_net = 0
    except Exception:
        pass

    try:
        ini.iopt_sim = 'rms'
    except Exception:
        pass

    try:
        ini.tstart = T_START
    except Exception:
        pass

    try:
        ini.p_resvar = result_file
    except Exception:
        try:
            ini.p_res = result_file
        except Exception as exc:
            raise RuntimeError(
                'Impossibile collegare ElmRes a ComInc: {}'.format(exc)
            )

    try:
        sim.tstop = T_STOP
    except Exception:
        pass

    try:
        sim.dtgrd = DT
    except Exception:
        pass

    return ldf, ini, sim


def execute_command(command, label):
    try:
        status = command.Execute()
    except Exception as exc:
        raise RuntimeError(
            '{}: {}'.format(label, exc)
        )

    if status not in (None, 0):
        raise RuntimeError(
            '{} fallito, codice {}'.format(label, status)
        )


# =============================================================================
# LETTURA DELLA FREQUENZA E DEL TEMPO REALE DA ELMRES
# =============================================================================

def extract_frequency(app, result_file, terminal):
    """
    Legge:
    - tempo reale dalla colonna ElmRes -1;
    - frequenza da CENTRALE:m:fehz.

    Non ricostruisce il tempo con DT.
    """
    app.ResLoadData(result_file)

    try:
        col_frequency = int(
            app.ResGetIndex(
                result_file,
                terminal,
                'm:fehz'
            )
        )

        if col_frequency < 0:
            raise RuntimeError(
                'm:fehz non presente in BatchResults.ElmRes.'
            )

        n_samples = int(
            app.ResGetValueCount(
                result_file,
                col_frequency
            )
        )

        if n_samples <= 0:
            raise RuntimeError(
                'Result File senza campioni.'
            )

        times = []
        frequencies = []

        for row_index in range(n_samples):

            raw_time = app.ResGetData(
                result_file,
                row_index,
                -1
            )

            raw_frequency = app.ResGetData(
                result_file,
                row_index,
                col_frequency
            )

            times.append(scalar(raw_time))
            frequencies.append(scalar(raw_frequency))

        return times, frequencies

    finally:
        try:
            app.ResReleaseData(result_file)
        except Exception:
            pass


# =============================================================================
# METRICHE DI FREQUENZA E TEMPO DI ASSESTAMENTO
# =============================================================================

def calculate_settling_time(post_pairs, f_equilibrium):
    """
    Definizione:

    - banda: f_equilibrium +/- SETTLING_BAND_HZ;
    - assestamento = primo istante dopo EVENT_TIME in cui:
      1. la frequenza resta nella banda almeno SETTLING_HOLD secondi;
      2. non esce più dalla banda fino alla fine della simulazione.

    Il valore restituito è relativo a EVENT_TIME.
    """
    if not post_pairs:
        return float('nan')

    lower_limit = f_equilibrium - SETTLING_BAND_HZ
    upper_limit = f_equilibrium + SETTLING_BAND_HZ

    final_time = post_pairs[-1][0]

    for start_index, (start_time, _) in enumerate(post_pairs):

        if start_time + SETTLING_HOLD > final_time:
            continue

        hold_end = start_time + SETTLING_HOLD
        hold_ok = True

        for current_time, current_frequency in post_pairs[start_index:]:

            if current_time > hold_end:
                break

            if (
                current_frequency < lower_limit
                or current_frequency > upper_limit
            ):
                hold_ok = False
                break

        if not hold_ok:
            continue

        remains_in_band = True

        for _, current_frequency in post_pairs[start_index:]:

            if (
                current_frequency < lower_limit
                or current_frequency > upper_limit
            ):
                remains_in_band = False
                break

        if remains_in_band:
            return start_time - EVENT_TIME

    return float('nan')


def metrics(times, frequencies):
    """
    Calcola:
    - f_reference: media tra EVENT_TIME - 1 s e EVENT_TIME;
    - f_equilibrium: media ultimi SETTLING_WINDOW secondi;
    - delta_f_max: max |f(t) - f_reference| dopo il distacco;
    - delta_percent: delta_f_max/f_reference * 100;
    - settling_time: tempo relativo all'evento.
    """
    if len(times) != len(frequencies) or len(times) < 2:
        raise RuntimeError(
            'Risultato frequenza vuoto o inconsistente.'
        )

    valid_pairs = [
        (time_value, frequency)
        for time_value, frequency in zip(times, frequencies)
        if math.isfinite(time_value)
        and math.isfinite(frequency)
    ]

    if not valid_pairs:
        raise RuntimeError(
            'Nessun campione valido nel Result File.'
        )

    pre_event = [
        frequency
        for time_value, frequency in valid_pairs
        if EVENT_TIME - 1.0 <= time_value < EVENT_TIME
    ]

    if not pre_event:
        pre_event = [
            frequency
            for time_value, frequency in valid_pairs
            if time_value < EVENT_TIME
        ]

    if not pre_event:
        raise RuntimeError(
            'Nessun campione pre-evento trovato.'
        )

    f_reference = mean(pre_event)

    post_pairs = [
        (time_value, frequency)
        for time_value, frequency in valid_pairs
        if time_value >= EVENT_TIME
    ]

    if not post_pairs:
        raise RuntimeError(
            'Nessun campione post-evento trovato.'
        )

    post_values = [
        frequency
        for _, frequency in post_pairs
    ]

    delta_f_max = max(
        abs(frequency - f_reference)
        for frequency in post_values
    )

    equilibrium_start = T_STOP - SETTLING_WINDOW

    equilibrium_values = [
        frequency
        for time_value, frequency in post_pairs
        if time_value >= equilibrium_start
    ]

    if not equilibrium_values:
        equilibrium_values = [post_pairs[-1][1]]

    f_equilibrium = mean(equilibrium_values)

    settling_time = calculate_settling_time(
        post_pairs,
        f_equilibrium
    )

    if abs(f_reference) > 1e-12:
        delta_percent = (
            100.0 * delta_f_max / abs(f_reference)
        )
    else:
        delta_percent = float('nan')

    return (
        f_reference,
        f_equilibrium,
        delta_f_max,
        delta_percent,
        settling_time
    )


# =============================================================================
# OUTPUT CSV
# =============================================================================

def write_results(path, results):
    folder = os.path.dirname(path)

    if folder:
        os.makedirs(folder, exist_ok=True)

    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(
            f,
            fieldnames=RESULT_COLUMNS,
            delimiter=';'
        )

        writer.writeheader()

        for row in results:
            writer.writerow({
                column: csv_value(row.get(column, ''))
                for column in RESULT_COLUMNS
            })


# =============================================================================
# MAIN
# =============================================================================

def main():

    app = pf.GetApplication()

    if app is None:
        raise RuntimeError('PowerFactory non disponibile.')

    app.ClearOutputWindow()

    activation = app.ActivateProject(PROJECT_NAME)

    if activation not in (None, 0):
        raise RuntimeError(
            'Impossibile attivare il progetto: {}'.format(
                PROJECT_NAME
            )
        )

    pf_print(app, '--- Avvio sweep parametrico ---')

    rows = read_csv_rows(CSV_FILE)

    if not rows:
        raise RuntimeError(
            'Nessuna riga valida trovata nel CSV.'
        )

    generator = get_one(
        app,
        GENERATOR_NAME + '.ElmSym',
        GENERATOR_NAME
    )

    switch_obj = get_one(
        app,
        SWITCH_NAME + '.ElmCoup',
        SWITCH_NAME
    )

    terminal = get_one(
        app,
        TERMINAL_NAME + '.ElmTerm',
        TERMINAL_NAME
    )

    clear_all_switch_events(app)
    reset_switch_and_gen(generator, switch_obj)

    results = []

    for index, row in enumerate(rows, start=1):

        iteration = row.get('Iterazione', index)
        stable_csv = row.get('Stabile', '')

        event = None
        result_file = None

        record = {
            'Iterazione': iteration,
            'Stabile_CSV': stable_csv,
            'Frequenza_riferimento_Hz': '',
            'Frequenza_equilibrio_Hz': '',
            'Delta_f_max_Hz': '',
            'Delta_f_max_percento': '',
            'Tempo_assestamento_s': '',
            'Esito': 'ERRORE',
            'Errore': ''
        }

        # =============================================================
        # FILTRO OBBLIGATORIO:
        # esegue soltanto righe con Stabile=True.
        # =============================================================
        if not is_true(stable_csv):

            record['Esito'] = 'SALTATA'
            record['Errore'] = (
                'Riga esclusa: valore della colonna Stabile diverso da True.'
            )

            results.append(record)
            write_results(OUTPUT_CSV, results)

            pf_print(
                app,
                'Iterazione {} saltata: Stabile={}.'.format(
                    iteration,
                    stable_csv
                )
            )

            continue

        try:
            # 0. Ripristino della configurazione iniziale
            clear_all_switch_events(app)
            reset_switch_and_gen(generator, switch_obj)

            # 1. Assegna i parametri della riga corrente
            set_params_from_row(app, row)

            # 2. Crea evento di apertura switch a t=100 s
            event = create_disconnect_event(app, switch_obj)

            # 3. Crea Result File dedicato con CENTRALE:m:fehz
            result_file = create_result_file(app, terminal)

            # 4. Configura comandi
            ldf, ini, sim = configure_commands(
                app,
                result_file
            )

            # 5. Esegue load flow + inizializzazione + RMS
            execute_command(ldf, 'ComLdf')
            execute_command(ini, 'ComInc')
            execute_command(sim, 'ComSim')

            # 6. Estrae tempo reale e frequenza dall'ElmRes
            times, frequencies = extract_frequency(
                app,
                result_file,
                terminal
            )

            if DEBUG_TIME_AXIS:
                dt_first = (
                    times[1] - times[0]
                    if len(times) > 1
                    else float('nan')
                )

                pf_print(
                    app,
                    (
                        'DEBUG iterazione {}: '
                        't_inizio = {:.6f} s, '
                        't_fine = {:.6f} s, '
                        'dt_primo = {:.6f} s, '
                        'N campioni = {}'
                    ).format(
                        iteration,
                        times[0],
                        times[-1],
                        dt_first,
                        len(times)
                    )
                )

            # 7. Calcola metriche
            (
                f_reference,
                f_equilibrium,
                delta_f_max,
                delta_percent,
                settling_time
            ) = metrics(times, frequencies)

            record.update({
                'Frequenza_riferimento_Hz': f_reference,
                'Frequenza_equilibrio_Hz': f_equilibrium,
                'Delta_f_max_Hz': delta_f_max,
                'Delta_f_max_percento': delta_percent,
                'Tempo_assestamento_s': settling_time,
                'Esito': 'OK'
            })

            if math.isfinite(settling_time):
                settling_text = '{:.3f} s'.format(settling_time)
            else:
                settling_text = 'non raggiunto'

            pf_print(
                app,
                (
                    'Iterazione {} completata: '
                    'Delta f max = {:.6f} Hz, '
                    'f_eq = {:.6f} Hz, '
                    'assestamento = {}'
                ).format(
                    iteration,
                    delta_f_max,
                    f_equilibrium,
                    settling_text
                )
            )

        except Exception as exc_err:

            record['Esito'] = 'ERRORE'
            record['Errore'] = str(exc_err)

            pf_print(
                app,
                'Iterazione {} ERRORE: {}'.format(
                    iteration,
                    exc_err
                )
            )

            # Per debug esteso, scommenta:
            # pf_print(app, traceback.format_exc())

        finally:

            # Elimina evento dell'iterazione corrente.
            safe_delete(event)

            # Elimina eventuali eventi EvtSwitch residui.
            clear_all_switch_events(app)

            # Ripristina Gen 2 connesso e switch chiuso.
            try:
                reset_switch_and_gen(generator, switch_obj)
            except Exception as reset_error:
                pf_print(
                    app,
                    'ATTENZIONE reset iterazione {}: {}'.format(
                        iteration,
                        reset_error
                    )
                )

            # Elimina Result File temporaneo.
            safe_delete(result_file)

        results.append(record)

        # Salvataggio progressivo.
        write_results(OUTPUT_CSV, results)

    # Stato finale della rete.
    clear_all_switch_events(app)
    reset_switch_and_gen(generator, switch_obj)

    write_results(OUTPUT_CSV, results)

    pf_print(
        app,
        '--- Sweep completato: {}/{} righe CSV elaborate ---'.format(
            len(results),
            len(rows)
        )
    )

    pf_print(
        app,
        'Risultati salvati in: {}'.format(OUTPUT_CSV)
    )


if __name__ == '__main__':
    main()
