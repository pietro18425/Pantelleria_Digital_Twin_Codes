

# -*- coding: utf-8 -*-
"""
Sweep PowerFactory: RoCoF con apertura dello switch del generatore.

Il file Classifica_Set_Parametri_Stabili.csv contiene già:
- Iterazione;
- Rank;
- parametri DEGov;
- parametri Exciter;
- solo i set stabili.

Per ogni set:
- applica direttamente i parametri dalla classifica;
- abilita il generatore e chiude lo switch;
- apre lo switch a t = 100 s;
- esegue RMS da -100 s a 300 s;
- legge CENTRALE:m:fehz;
- calcola il massimo RoCoF su finestra mobile di 0,1 s.
"""

import sys
import os
import csv
import math

import pandas as pd

sys.path.append(
    r'C:\Program Files\DIgSILENT\PowerFactory 2025 SP4\Python\3.13'
)
import powerfactory as pf


# =============================================================================
# CONFIGURAZIONE
# =============================================================================

PROJECT_NAME = (
    r'Pantelleria_definitivo_spiller_con_GFM.IntPrj'
)

STABLE_FILE = (
    r'C:\Users\159249\OneDrive - Politecnico di Milano'
    r'\Desktop\TUNING\Analisi_Completa_GFM'
    r'\Classifica_Set_Parametri_Stabili.csv'
)

OUTPUT_FOLDER = (
    r'C:\Users\159249\OneDrive - Politecnico di Milano'
    r'\Desktop\TUNING\Risultati_Batch\GFM\RoCoF'
)

OUTPUT_CSV = os.path.join(
    OUTPUT_FOLDER,
    'Risultati_ROCOF_2_GEN_GFM.csv'
)

T_START = -100.0
EVENT_TIME = 100.0
T_STOP = 300.0
DT = 0.01

ROCOF_WINDOW_S = 0.10

# None = tutto il periodo dopo l'evento.
# Esempio: 5.0 = solo i primi 5 s dopo l'evento.
ROCOF_DURATION_S = None

ROCOF_WINDOW_TOLERANCE = 0.20

GENERATOR_NAME = 'Gen 2'
SWITCH_NAME = 'Switch Gen 2'
TERMINAL_NAME = 'CENTRALE'
N_GENERATORS = 8


# =============================================================================
# PARAMETRI
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
    'Rank',
    'ROCOF_max_abs_Hz_s',
    'ROCOF_picco_con_segno_Hz_s',
    'ROCOF_t_picco_s',
    'ROCOF_finestra_s',
    'Esito',
    'Errore',
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


def scalar(value):
    if isinstance(value, (tuple, list)):
        value = value[-1]

    return float(value)


def csv_value(value):
    if value is None:
        return ''

    if isinstance(value, float):
        if not math.isfinite(value):
            return ''

        return '{:.8f}'.format(value).replace('.', ',')

    return str(value)


def parse_classification_number(value):
    if pd.isna(value):
        return float('nan')

    text = str(value).strip()

    if text == '':
        return float('nan')

    if ',' in text and '.' in text:
        if text.rfind(',') > text.rfind('.'):
            text = text.replace('.', '')
            text = text.replace(',', '.')
        else:
            text = text.replace(',', '')

    elif ',' in text:
        text = text.replace(',', '.')

    return float(text)


# =============================================================================
# LETTURA CLASSIFICA
# =============================================================================

def read_stable_parameter_rows(path):
    if not os.path.isfile(path):
        raise FileNotFoundError(
            'File classifica non trovato:\n{}'.format(path)
        )

    df = pd.read_csv(
        path,
        sep=';',
        decimal=',',
        encoding='utf-8-sig',
        dtype=str,
        low_memory=False,
    )

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    required_columns = (
        GOV_COLUMNS
        + EXC_COLUMNS
        + ['Iterazione', 'Rank']
    )

    missing = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing:
        raise RuntimeError(
            'Colonne mancanti nella classifica: {}'.format(
                ', '.join(missing)
            )
        )

    df['Iterazione'] = pd.to_numeric(
        df['Iterazione'],
        errors='coerce',
    )

    df['Rank'] = pd.to_numeric(
        df['Rank'],
        errors='coerce',
    )

    df = df.dropna(
        subset=['Iterazione', 'Rank']
    ).copy()

    df['Iterazione'] = df['Iterazione'].astype(int)
    df['Rank'] = df['Rank'].astype(int)

    for column in GOV_COLUMNS + EXC_COLUMNS:
        df[column] = df[column].map(
            parse_classification_number
        )

    df = df.dropna(
        subset=GOV_COLUMNS + EXC_COLUMNS
    ).copy()

    df = df.sort_values(
        by='Rank',
        ascending=True,
    ).reset_index(drop=True)

    duplicate_iterations = df.loc[
        df['Iterazione'].duplicated(),
        'Iterazione'
    ].tolist()

    if duplicate_iterations:
        raise RuntimeError(
            'Iterazioni duplicate nella classifica: {}'.format(
                duplicate_iterations[:20]
            )
        )

    return df.to_dict(
        orient='records'
    )


# =============================================================================
# POWERFACTORY
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


def set_params_from_row(app, row):
    gov_params = [
        row[column]
        for column in GOV_COLUMNS
    ] + GOV_EXTRA

    exc_params = [
        row[column]
        for column in EXC_COLUMNS
    ] + EXC_EXTRA

    for generator_number in range(
        1,
        N_GENERATORS + 1
    ):
        gov_list = app.GetCalcRelevantObjects(
            'pcu_DEGOV1_GEN{}.ElmDsl'.format(
                generator_number
            )
        )

        exc_list = app.GetCalcRelevantObjects(
            'vcoEXST2A_GEN{}.ElmDsl'.format(
                generator_number
            )
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

        gov_list[0].params = gov_params
        exc_list[0].params = exc_params


def close_switch(switch_obj):
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


def reset_switch_and_gen(generator, switch_obj):
    try:
        generator.outserv = 0
    except Exception:
        pass

    close_switch(switch_obj)


def clear_all_switch_events(app):
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
    event_folder = app.GetFromStudyCase('IntEvt')

    if event_folder is None:
        raise RuntimeError(
            'Cartella IntEvt non trovata.'
        )

    event = event_folder.CreateObject(
        'EvtSwitch',
        'Trip_{}_{}s'.format(
            GENERATOR_NAME.replace(' ', ''),
            int(EVENT_TIME)
        )
    )

    if event is None:
        raise RuntimeError(
            'Impossibile creare EvtSwitch.'
        )

    event.time = EVENT_TIME
    event.p_target = switch_obj
    event.i_switch = 0

    return event


def create_result_file(app, terminal):
    result_folder = get_study_object(
        app,
        'Results'
    )

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

    result_file.AddVars(
        terminal,
        'm:fehz'
    )

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
                'Impossibile collegare ElmRes a ComInc: {}'.format(
                    exc
                )
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
# LETTURA FREQUENZA
# =============================================================================

def extract_frequency(app, result_file, terminal):
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
            times.append(
                scalar(
                    app.ResGetData(
                        result_file,
                        row_index,
                        -1
                    )
                )
            )

            frequencies.append(
                scalar(
                    app.ResGetData(
                        result_file,
                        row_index,
                        col_frequency
                    )
                )
            )

        return times, frequencies

    finally:
        try:
            app.ResReleaseData(result_file)
        except Exception:
            pass


# =============================================================================
# CALCOLO ROCOF
# =============================================================================

def calculate_rocof(times, frequencies):
    pairs = sorted(
        (float(t), float(f))
        for t, f in zip(times, frequencies)
        if math.isfinite(t)
        and math.isfinite(f)
        and t >= EVENT_TIME
        and (
            ROCOF_DURATION_S is None
            or t <= EVENT_TIME + ROCOF_DURATION_S
        )
    )

    if len(pairs) < 2:
        raise RuntimeError(
            'Campioni post-evento insufficienti.'
        )

    best_abs = None
    best_signed = None
    best_time = None
    start_index = 0

    for end_index in range(1, len(pairs)):
        end_time, end_frequency = pairs[end_index]
        desired_start = end_time - ROCOF_WINDOW_S

        if desired_start < EVENT_TIME:
            continue

        while (
            start_index + 1 < end_index
            and pairs[start_index + 1][0] <= desired_start
        ):
            start_index += 1

        candidates = [start_index]

        if start_index + 1 < end_index:
            candidates.append(start_index + 1)

        candidate_index = min(
            candidates,
            key=lambda idx: abs(
                (end_time - pairs[idx][0])
                - ROCOF_WINDOW_S
            )
        )

        start_time, start_frequency = pairs[
            candidate_index
        ]

        actual_window = end_time - start_time

        if actual_window <= 0:
            continue

        relative_error = abs(
            actual_window - ROCOF_WINDOW_S
        ) / ROCOF_WINDOW_S

        if relative_error > ROCOF_WINDOW_TOLERANCE:
            continue

        signed_rocof = (
            end_frequency - start_frequency
        ) / actual_window

        abs_rocof = abs(signed_rocof)

        if best_abs is None or abs_rocof > best_abs:
            best_abs = abs_rocof
            best_signed = signed_rocof
            best_time = end_time

    if best_abs is None:
        raise RuntimeError(
            'Nessuna finestra RoCoF valida.'
        )

    return best_abs, best_signed, best_time


# =============================================================================
# RISULTATI
# =============================================================================

def write_results(path, results):
    os.makedirs(
        os.path.dirname(path),
        exist_ok=True
    )

    with open(
        path,
        'w',
        encoding='utf-8-sig',
        newline=''
    ) as f:
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
        raise RuntimeError(
            'PowerFactory non disponibile.'
        )

    app.ClearOutputWindow()

    activation = app.ActivateProject(
        PROJECT_NAME
    )

    if activation not in (None, 0):
        raise RuntimeError(
            'Impossibile attivare il progetto: {}'.format(
                PROJECT_NAME
            )
        )

    rows = read_stable_parameter_rows(
        STABLE_FILE
    )

    if not rows:
        raise RuntimeError(
            'La classifica non contiene set stabili.'
        )

    pf_print(
        app,
        'Set stabili da analizzare: {}'.format(
            len(rows)
        )
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
    reset_switch_and_gen(
        generator,
        switch_obj
    )

    results = []

    for position, row in enumerate(rows, start=1):
        iteration = int(row['Iterazione'])
        rank = int(row['Rank'])

        event = None
        result_file = None

        record = {
            'Iterazione': iteration,
            'Rank': rank,
            'ROCOF_max_abs_Hz_s': '',
            'ROCOF_picco_con_segno_Hz_s': '',
            'ROCOF_t_picco_s': '',
            'ROCOF_finestra_s': ROCOF_WINDOW_S,
            'Esito': 'ERRORE',
            'Errore': ''
        }

        try:
            pf_print(
                app,
                '--- Analisi {}/{}: Iterazione {}, Rank {} ---'.format(
                    position,
                    len(rows),
                    iteration,
                    rank
                )
            )

            clear_all_switch_events(app)
            reset_switch_and_gen(
                generator,
                switch_obj
            )

            set_params_from_row(
                app,
                row
            )

            event = create_disconnect_event(
                app,
                switch_obj
            )

            result_file = create_result_file(
                app,
                terminal
            )

            ldf, ini, sim = configure_commands(
                app,
                result_file
            )

            execute_command(ldf, 'ComLdf')
            execute_command(ini, 'ComInc')
            execute_command(sim, 'ComSim')

            times, frequencies = extract_frequency(
                app,
                result_file,
                terminal
            )

            rocof_abs, rocof_signed, rocof_peak_time = (
                calculate_rocof(
                    times,
                    frequencies
                )
            )

            record.update({
                'ROCOF_max_abs_Hz_s': rocof_abs,
                'ROCOF_picco_con_segno_Hz_s': rocof_signed,
                'ROCOF_t_picco_s': rocof_peak_time,
                'Esito': 'OK'
            })

            pf_print(
                app,
                (
                    'Iterazione {} completata: '
                    'RoCoF max = {:.6f} Hz/s; '
                    't picco = {:.3f} s'
                ).format(
                    iteration,
                    rocof_abs,
                    rocof_peak_time
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

        finally:
            safe_delete(event)
            clear_all_switch_events(app)

            try:
                reset_switch_and_gen(
                    generator,
                    switch_obj
                )
            except Exception as reset_error:
                pf_print(
                    app,
                    'ATTENZIONE reset iterazione {}: {}'.format(
                        iteration,
                        reset_error
                    )
                )

            safe_delete(result_file)

        results.append(record)

        write_results(
            OUTPUT_CSV,
            results
        )

    clear_all_switch_events(app)
    reset_switch_and_gen(
        generator,
        switch_obj
    )

    write_results(
        OUTPUT_CSV,
        results
    )

    pf_print(
        app,
        '--- Sweep RoCoF completato: {}/{} ---'.format(
            len(results),
            len(rows)
        )
    )

    pf_print(
        app,
        'Risultati salvati in: {}'.format(
            OUTPUT_CSV
        )
    )


if __name__ == '__main__':
    main()