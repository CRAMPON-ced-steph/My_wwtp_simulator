#!/usr/bin/env python3
"""
Extrait la base climatique du classeur OCEAN vers un JSON exploitable par le
port HVAC : 10 896 stations, chacune avec 12 températures, pressions et
humidités relatives moyennes mensuelles.

    pip install openpyxl
    python3 outils/extraire_meteo.py 140822_OCEAN_CCR.xlsm src/data/meteo.json

La feuille `BD_Climat_Stations` porte les noms définis `idStation` (colonne A),
`nomStation` (C), `AirT_Jan` (J), `Hum_Jan` (V) et `Press_Jan` (AT), chacun
suivi des onze mois suivants.
"""
import json
import sys

import openpyxl

COLONNES = {'id': 1, 'nom': 3, 'T': 10, 'HR': 22, 'P': 46}  # A, C, J, V, AT


def extraire(chemin_xlsm):
    wb = openpyxl.load_workbook(chemin_xlsm, read_only=True, data_only=True)
    ws = wb['BD_Climat_Stations']
    stations = []
    for ligne in ws.iter_rows(values_only=True):
        identifiant = ligne[COLONNES['id'] - 1] if len(ligne) >= COLONNES['id'] else None
        if not isinstance(identifiant, (int, float)):
            continue
        serie = lambda c: [ligne[COLONNES[c] - 1 + i] for i in range(12)]
        T, P, HR = serie('T'), serie('P'), serie('HR')
        if not all(isinstance(x, (int, float)) for x in T):
            continue
        stations.append({
            'id': int(identifiant),
            'nom': str(ligne[COLONNES['nom'] - 1]),
            'T': [round(float(x), 2) for x in T],
            'P': [round(float(x), 2) for x in P],
            'HR': [round(float(x), 1) for x in HR],
        })
    return stations


if __name__ == '__main__':
    source = sys.argv[1] if len(sys.argv) > 1 else '140822_OCEAN_CCR.xlsm'
    cible = sys.argv[2] if len(sys.argv) > 2 else 'src/data/meteo.json'
    stations = extraire(source)
    with open(cible, 'w', encoding='utf-8') as fichier:
        json.dump(stations, fichier, ensure_ascii=False)
    print(f'{len(stations)} stations écrites dans {cible}')
