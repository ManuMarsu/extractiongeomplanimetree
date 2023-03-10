import os
import xlsxwriter
from operator import itemgetter

wb = xlsxwriter.Workbook('geom.xlsx')
ws = {}
min_max_z = {}

with open("mascaret.lis", "r") as f:
    lines = f.readlines()
    section_profil = 0
    ligne_profil = False
    for line in lines:
        # On regarde si la ligne commence par "Profil de donnee numero"
        if line.startswith("Profil de donnee numero"):
            nom_profil = line.split(',')[1].split('\n')[0].split(' : ')[1].replace(' ', '') + '-' + line.split(',')[0][-1]
            print(nom_profil, line)
            min_max_z[nom_profil] = [1000000, -1000000, 0]
            ws[nom_profil] = wb.add_worksheet(nom_profil)
            ws[nom_profil].write(0, 0, "Z")
            ws[nom_profil].write(0, 1, "db1")
            ws[nom_profil].write(0, 2, "db2")
            ws[nom_profil].write(0, 3, "dp1")
            ws[nom_profil].write(0, 4, "dp2")
            ws[nom_profil].write(0, 5, "ds1")
            ws[nom_profil].write(0, 6, "ds2")
            ws[nom_profil].write(0, 7, "dbs")
            ws[nom_profil].write(0, 8, "dss")
            section_profil = 1
            
        if section_profil > 0:
            section_profil += 1
            if line == "\n" and section_profil not in [6, 8]:
                section_profil = 0
            if section_profil == 3:
                abscisse = float(line.split('\n')[0].split('=')[1].replace(' ', '').replace(' ', '').replace(' ', '').replace(' ', ''))
                min_max_z[nom_profil][2] = abscisse
        if section_profil > 8:
            l_simpl = line.split('\n')[0].replace('  ', ' ').replace('  ', ' ').replace('  ', ' ').replace('  ', ' ').replace('  ', ' ').replace('  ', ' ').split(' ')
            if len(l_simpl) == 10:
                t_Z = float(l_simpl[1])
                t_db1 = float(l_simpl[2])
                t_db2 = float(l_simpl[3])
                t_dp1 = float(l_simpl[4])
                t_dp2 = float(l_simpl[5])
                t_ds1 = float(l_simpl[6])
                t_ds2 = float(l_simpl[7])
                t_dbs = float(l_simpl[8])
                t_dss = float(l_simpl[9])
                min_max_z[nom_profil] = [min(t_Z, min_max_z[nom_profil][0]), max(t_Z, min_max_z[nom_profil][1]), min_max_z[nom_profil][2]]
                ws[nom_profil].write(section_profil-8, 0, t_Z)
                ws[nom_profil].write(section_profil-8, 1, t_db1)
                ws[nom_profil].write(section_profil-8, 2, t_db2)
                ws[nom_profil].write(section_profil-8, 3, t_dp1)
                ws[nom_profil].write(section_profil-8, 4, t_dp2)
                ws[nom_profil].write(section_profil-8, 5, t_ds1)
                ws[nom_profil].write(section_profil-8, 6, t_ds2)
                ws[nom_profil].write(section_profil-8, 7, t_dbs)
                ws[nom_profil].write(section_profil-8, 8, t_dss)
                # print(nom_profil, t_Z, t_db1, t_db2, t_dp1, t_dp2, t_ds1, t_ds2, t_dbs, t_dss)
            else:
                section_profil = 0

    # Tri des profils par ordre croissant de l'abscisse
    abs_profils = sorted([(nom_profil, min_max_z[nom_profil][2]) for nom_profil in min_max_z.keys()], key=itemgetter(1,0))
    print(abs_profils)
    # Ajout de feuilles graphiques pour chaque profil
    cs = []
    for i, (prof_cur, absc_cur) in enumerate(abs_profils):
        if i not in [0, len(abs_profils)-1]:
            chart = wb.add_chart({'type': 'scatter', 'subtype': 'straight'})

            (prof_prec, absc_prec) = abs_profils[i-1]
            (prof_suiv, absc_suiv) = abs_profils[i+1]
            d = f'={abs_profils[i-1][0]}!$A$1000+{abs_profils[i+1][0]}!$A$2/2'

            trad = {'F': 'Lit mineur', 'G': 'Lit majeur'}

            # Courbes profil pr??c??dent
            for col in ['F', 'G']:
                chart.add_series({
                    'name': f'Prec_{trad[col]}_{str(absc_cur - absc_prec)}',
                    'categories': f'={prof_prec}!$A$2:$A$1000', 
                    'values': f'={prof_prec}!${col}$2:${col}$1000', 
                })
            # Courbes profil actuel
            for col in ['F', 'G']:
                chart.add_series({
                    'name': f'{prof_cur}_{trad[col]}_{absc_cur}',
                    'categories': f'={prof_cur}!$A$2:$A$1000', 
                    'values': f'={prof_cur}!${col}$2:${col}$1000', 
                })
            # Courbes profil suivant
            for col in ['F', 'G']:
                chart.add_series({
                    'name': f'Suiv_{trad[col]}_{str(absc_suiv - absc_cur)}',
                    'categories': f'={prof_suiv}!$A$2:$A$1000', 
                    'values': f'={prof_suiv}!${col}$2:${col}$1000', 
                })

            chart.set_title({'name': prof_cur})
            chart.set_x_axis({'name': 'Z', 'min': min_max_z[prof_cur][0] - 1, 'max': min_max_z[prof_cur][0] + 4})
            chart.set_y_axis({'name': 'Section hydraulique', 'min': 0, 'max': 100})
            
            cs.append(wb.add_chartsheet(f'={prof_cur}-{absc_cur}'))
            cs[i-1].set_chart(chart)

    # Ajout des graphiques correspondants ?? chaque profil
    for nom_profil, profil in ws.items():
        chart = wb.add_chart({'type': 'scatter', 'subtype': 'straight'})
        # for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
        for col in ['F', 'G']:
            chart.add_series({
                'name': f'={nom_profil}!${col}$1',
                'categories': f'={nom_profil}!$A$2:$A$1000', 
                'values': f'={nom_profil}!${col}$2:${col}$1000', 
            })

        chart.set_title({'name': nom_profil})
        chart.set_size({'width': 1200, 'height': 750})
        ws[nom_profil].insert_chart('D2', chart)
        chart.set_x_axis({'name': 'Z', 'min': min_max_z[nom_profil][0], 'max': min_max_z[nom_profil][0] + 6})
        chart.set_y_axis({'name': 'Section hydraulique', 'min': 0, 'max': 100})


wb.close()