# -*- coding: utf-8 -*-

import numpy as np
from scipy import interpolate
from scipy import integrate
import matplotlib.pyplot as plt
import math
from matplotlib.backends.backend_pdf import PdfPages
from scipy.optimize import curve_fit
from scipy.integrate import quad
import decimal
import pandas as pd
import calendar
import pprint
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.colors import blue
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
import os
from reportlab.pdfbase.pdfmetrics import stringWidth
import xlsxwriter
from openpyxl import load_workbook
from openpyxl import Workbook

hoy=pd.to_datetime('today')
agno=[2020,2021]

def find_missing(arr1, arr2):
    return sum(arr1)-sum(arr2)
    
if hoy.month>2:
    mesdepago=hoy.month-2
elif hoy.month==2:
    mesdepago=12
elif hoy.month==1:
    mesdepago=11

def unique(list1):
  unique_list = []
  for x in list1:
      if x not in unique_list:
          unique_list.append(x)


#Para correr el programa en el compu de Jessica tienes que incluir # en el compienzo de la siguiente linea
PATHES=["Documents/Vitalsur/%i/", "Documents/Vitalsur/%i/Bonos/"]

#Para correr el programa en el compu de Jessica tienes que borrar el # de la siguiente linea
#PATHES=["Desktop\Vitalsur_bonos\%i", "Desktop\Vitalsur_bonos\%i\Bonos"]
    
dfs = pd.read_excel('turnos_nuevos.xlsx')
rut_paciente_f=np.asarray(dfs['Rut paciente'].tolist())
rut_paciente_f=rut_paciente_f.astype(int)
dia=np.asarray(dfs['Dia'].tolist())
mes=np.asarray(dfs['Mes'].tolist())
ano=np.asarray(dfs['Año'].tolist())
rut_profesional=np.asarray(dfs['Rut profesional'].tolist())
profesional=np.asarray(dfs['Profesional'].tolist())
dia_ingreso_datos=np.asarray(dfs['Dia ingreso'].tolist())
mes_ingreso_datos=np.asarray(dfs['Mes ingreso'].tolist())
ano_ingreso_datos=np.asarray(dfs['Año ingreso'].tolist())

dfs2 = pd.read_excel('info_pacientes.xlsx')
nombre_paciente=np.asarray(dfs2['Nombre paciente'].tolist())
rut_paciente=np.asarray(dfs2['Rut paciente'].tolist())
rut_completo=np.asarray(dfs2['Rut completo'].tolist())
dia_ingreso=np.asarray(dfs2['Dia ingreso'].tolist())
mes_ingreso=np.asarray(dfs2['Mes Ingreso'].tolist())
ano_ingreso=np.asarray(dfs2['Año ingreso'].tolist())
isapre=np.asarray(dfs2['Isapre'].tolist())
diagnostico=np.asarray(dfs2['Diagnostico'].tolist())
ciudad=np.asarray(dfs2['Ciudad'].tolist())
sector=np.asarray(dfs2['Sector'].tolist())

dfs3 = pd.read_excel('info_isapres.xlsx')
i_isapre=np.asarray(dfs3['Isapre'].tolist())
i_ciudad=np.asarray(dfs3['Ciudad'].tolist())
i_sector=np.asarray(dfs3['Sector'].tolist())
i_cobropago=np.asarray(dfs3['Cobro/Pago'].tolist())
i_medico=np.asarray(dfs3['Medico'].tolist())
i_enfermera=np.asarray(dfs3['Enfermera'].tolist())
i_nutri=np.asarray(dfs3['Nutricionista'].tolist())
i_kine=np.asarray(dfs3['Kinesiologo'].tolist())
i_psicologo=np.asarray(dfs3['Psicologo'].tolist())
i_fonoaudiologo=np.asarray(dfs3['Fonoaudiologo'].tolist())
i_to=np.asarray(dfs3['Terapeuta Ocupacional'].tolist())
i_tens=np.asarray(dfs3['Tens'].tolist())
i_curacionsimple=np.asarray(dfs3['Curacion Simple'].tolist())
i_educenfermera=np.asarray(dfs3['Educacion Enfermera'].tolist())
i_eductens=np.asarray(dfs3['Educacion Tens'].tolist())
i_intpsicosocial=np.asarray(dfs3['Intervension Psicosocial'].tolist())

dfs4 = pd.read_excel('info_profesionales.xlsx')
i_nombre_profesional=np.asarray(dfs4['Nombre profesional'].tolist())
i_rut_profesional=np.asarray(dfs4['Rut profesional'].tolist())
i_ocupacion=np.asarray(dfs4['Ocupacion'].tolist())
i_ciudadprof=np.asarray(dfs4['Ciudad'].tolist())
i_rutcompleto=np.asarray(dfs4['Rut completo'].tolist())
i_pagodif=np.asarray(dfs4['Pago Diferenciado'].tolist())

dfs7 = pd.read_excel('pago_diferenciado.xlsx')
PD_PROFESIONAL=np.asarray(dfs7['Rut Profesional'].tolist())
PD_PACIENTE=np.asarray(dfs7['Rut Paciente'].tolist())
PD_VALOR=np.asarray(dfs7['Valor'].tolist())

dfs5 = pd.read_excel('info_isapres_bonos.xlsx')
bono_isapre=np.asarray(dfs5['Isapre'].tolist())
bono_medico=np.asarray(dfs5['Medico'].tolist())
bono_enfermera=np.asarray(dfs5['Enfermera'].tolist())
bono_nutri=np.asarray(dfs5['Nutricionista'].tolist())
bono_kine=np.asarray(dfs5['Kinesiologo'].tolist())
bono_psicologo=np.asarray(dfs5['Psicologo'].tolist())
bono_fonoaudiologo=np.asarray(dfs5['Fonoaudiologo'].tolist())
bono_to=np.asarray(dfs5['Terapeuta Ocupacional'].tolist())
bono_tens=np.asarray(dfs5['Tens'].tolist())
bono_curacionsimple=np.asarray(dfs5['Curacion Simple'].tolist())
bono_educenfermera=np.asarray(dfs5['Educacion Enfermera'].tolist())
bono_eductens=np.asarray(dfs5['Educacion Tens'].tolist())
bono_intpsicosocial=np.asarray(dfs5['Intervension Psicosocial'].tolist())

mes_nombre=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

#Tipo de profesionales
rut_profesional_search_type=dfs['Profesional'].tolist()
prof_unique_type = list(set(rut_profesional_search_type))

#-----------------------------------------------------------------------------------------------------------
#Analisis Vitalsur_Pacientes
#-----------------------------------------------------------------------------------------------------------

mes_actual=int(max(mes))
rut_paciente_search=dfs2['Rut paciente'].tolist()
pac_unique = list(set(rut_paciente_search))
for i in range(len(pac_unique)):
    aux_rutpac_infoprof=np.where(pac_unique[i]==rut_paciente)[0]
    if len(aux_rutpac_infoprof)>1:
        ffwarning.write('Hay dos rut iguales en info_pacientes.xlsx')
        ffwarning.write(' Check: ' + rut_paciente[aux_rutpac_infoprof[0]])
        ffwarning.write(' Lineas en Excel: ' +  aux_rutpac_infoprof+2)
        ffwarning.write('\n')
        
rut_profesional_checkdup=dfs4['Rut profesional'].tolist()
prof_unique = list(set(rut_profesional_checkdup))
for i in range(len(prof_unique)):
    aux_rutpac_infoprof=np.where(prof_unique[i]==i_rut_profesional)[0]
    if len(aux_rutpac_infoprof)>1:
        ffwarning.write('Hay dos rut iguales en info_profesionales.xlsx')
        ffwarning.write(' Check: ' + i_rut_profesional[aux_rutpac_infoprof[0]])
        ffwarning.write(' Lineas en Excel: ' + aux_rutpac_infoprof+2)
        ffwarning.write('\n')
        
for gn in range(len(agno)):
    #Loop para años
    #save_name1 = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Vitalsur_Pacientes_%i.txt'%agno[gn])
    #save_name2 = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Jessica_REVISAR_%i.txt'%agno[gn])
    
    save_name1 = os.path.join(os.path.curdir, "%i/"%agno[gn], 'Vitalsur_Pacientes_%i.txt'%agno[gn])
    save_name2 = os.path.join(os.path.curdir, "%i/"%agno[gn], 'Jessica_REVISAR_%i.txt'%agno[gn])

    ff = open(save_name1, 'w')
    ffwarning = open(save_name2, 'w')
    
    #loop para el numero total de pacientes (info_pacientes.xlsx)
    for i in range(len(dfs2)):
        #Guardar informacion paciente (info_pacientes.xlsx)
        ff.write( 'Nombre Paciente: ' + str(format(nombre_paciente[i])) + '\n' )
        ff.write( 'RUT Paciente: ' + str(format(int(rut_paciente[i]))) + '\n' )
        ff.write( 'Fecha ingreso: ' + str(format(int(dia_ingreso[i]))) + '/' + str(format(int(mes_ingreso[i]))) + '/' + str(format(int(ano_ingreso[i])))+ '\n')
        ff.write( 'Isapre: ' + str(format(isapre[i])) + '\n'  )
        ff.write( 'Ciudad: ' + str(format(ciudad[i])) + '\n'  )
        ff.write( 'Diagnostico: ' + str(format(diagnostico[i])) + '\n' )
        ff.write( '\n')

        #Sacar valores de cobros y pagos dependiendo de la isapre usando isapre/ciudad/cobro_pago/profesional (info_isapres.xlsx)
        aux_isapre_infoisapres=np.where(isapre[i]==i_isapre)[0]
        aux_ciudad_dondevive=np.where(ciudad[i]==i_ciudad)[0]
        aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)) #info isapres
        if len(aux_isapreyciudad)==0:
            aux_ciudad_dondevive=np.where('Otros'==i_ciudad)[0]
        aux_sector_dondevive=np.where(sector[i]==i_sector)[0]
        aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)  & set(aux_sector_dondevive)) #info isapres
        aux_cobro_infoisapres=np.where('Cobro'==i_cobropago)[0]
        aux_pago_infoisapres=np.where('Pago'==i_cobropago)[0]
            
        #Posicion de la fila correspondiente al caso del paciente (info_isapres.xlsx)
        aux_poscobro_infoisapres=list(set(aux_isapreyciudad) & set(aux_cobro_infoisapres))
        aux_pospago_infoisapres=list(set(aux_isapreyciudad) & set(aux_pago_infoisapres))


        if len(aux_isapre_infoisapres)==0:
            ffwarning.write('Hay una isapre que no esta especificada en info_isapres.xlsx')
            ffwarning.write(' Check: '+ isapre[i] )
            ffwarning.write( '\n')
            print('WARNING1!')
        
        if len(aux_poscobro_infoisapres)>1:
            ffwarning.write('Hay dos valores para el cobro del turno en info_isapres.xlsx')
            ffwarning.write(' Check: '+ i_isapre[aux_isapre_infoisapres[0]] + ' ' + i_ciudad[aux_ciudad_dondevive[0]] +  ' COBRO')
            ffwarning.write( '\n')
        
        if len(aux_pospago_infoisapres)>1:
            ffwarning.write('Hay dos valores para el pago del turno en info_isapres.xlsx')
            ffwarning.write(' Check: ' +  i_isapre[aux_isapre_infoisapres[0]] + ' '+ i_ciudad[aux_ciudad_dondevive[0]] +  ' PAGO')
            ffwarning.write( '\n')

        if len(aux_pospago_infoisapres)==0 or len(aux_poscobro_infoisapres)==0:
            ffwarning.write('Hay un error en leer el cobro/pago de la isapre - info_isapres.xlsx')
            ffwarning.write(' Check: '+ isapre[i] )
            ffwarning.write( '\n')
            print('WARNING2!')

        if agno[gn]==2021:
            for j in range(mes_actual):
                pagar_mes=0
                cobrar_mes=0
                
                aux_agno_turnos=np.where(ano==agno[gn])[0]
                #Posicion de turnos hecho ese mes en turnos.xlsx
                aux_mes_turnos=np.where(mes==(j+1))[0]
                #Posicion de rud paciente en turnos.xlsx
                aux_rut_turnos=np.where(rut_paciente[i]==rut_paciente_f)[0]
                #Posicion de turnos hecho para ese paciente en ese mes
                aux_rutmes_turnos=list(set(aux_rut_turnos) & set(aux_mes_turnos) & set(aux_agno_turnos))
                
                #Guardar informacion sobre el numero de visitas para ese paciente durante ese mes
                ff.write( '                    _________________________________________________________________________________________________' )
                ff.write( '\n')
                ff.write( '                    ' )
                ff.write('Visitas mes de '  + str(mes_nombre[j]) +  ' = ' + str(format(int(len(aux_rutmes_turnos)))) +  ' ' )
                ff.write( '\n')
                
                #Loop para tipos de profesionales
                for k in range(len(prof_unique_type)):
                    #Posicion de profesional seleccionado en turnos.xlsx
                    aux_profesional_turnos=np.where(prof_unique_type[k]==profesional)[0]
                    #Posicion de profesional seleccionado en turnos para ese paciente durante ese mes en turnos.xlsx
                    aux_rutmesprof_turnos=list(set(aux_rut_turnos) & set(aux_mes_turnos) & set(aux_profesional_turnos) & set(aux_agno_turnos))
                    #Posicion de la columna con valores del profesional en info_isapres.xlsx
                    valores_profesional=np.asarray(dfs3[prof_unique_type[k]].tolist())
                    #Posicion de la columna con numeros de visitas del profesional en info_pacientes.xlsx
                    nvisitas_profesional=np.asarray(dfs2[prof_unique_type[k]].tolist())
                    
                    #Alerta en el caso de que el turno no tenga un valor definido
                    if valores_profesional[aux_poscobro_infoisapres]==0:
                        ffwarning.write('No se ha ingresado el valor del procedimiento/profesional ' + str(prof_unique_type[k]))
                        ffwarning.write(' Paciente: ' + nombre_paciente[i] +  ' / Isapre: ' + isapre[i] + ' / Mes: ' + mes_nombre[j] )
                        ffwarning.write(' Porque el valor no esta especificado en info_isapres.xlsx')
                        ffwarning.write( '\n')
                    
                    #Guardar informacion sobre el numero de visitas de ese profesional, los pagos y cobros equivalentes
                    #ff.write( '                                          .......................................................................................' )
                    ff.write( '\n')
                    ff.write( '                                          ' )
                    ff.write('Visitas de '  + str(prof_unique_type[k]) +  ' = ' + str(format(int(len(aux_rutmesprof_turnos)))) + ' / ' +   str(nvisitas_profesional[i]))
                    if nvisitas_profesional[i]!=0:
                        ff.write(' (' +  str(format(int(len(aux_rutmesprof_turnos) * 100 / nvisitas_profesional[i]))) + '%)')
                    ff.write('   --------------------->   FALTAN =' + str(format(np.abs( int(len(aux_rutmesprof_turnos))   - nvisitas_profesional[i]) )) )
                    if nvisitas_profesional[i]!=0:
                        ff.write(' (' +  str(format(int( 100 - ( len(aux_rutmesprof_turnos) * 100 / nvisitas_profesional[i])))) + '%)')
                    ff.write( '\n')
                    ff.write( '                                          ' )
                    ff.write('Pago equivalente = ' + str(format(int(sum(valores_profesional[aux_pospago_infoisapres]) * len(aux_rutmesprof_turnos)))) +  '   -   ' )
                    ff.write('Cobro equivalente = ' + str(format(int(sum(valores_profesional[aux_poscobro_infoisapres]) * len(aux_rutmesprof_turnos)))) +  ' ' )
                    ff.write( '\n')
                    ff.write( '\n')

                    #Loop para cantidad de visitas de ese rut, mes y profesional en turnos.xlsx
                    for p in range(len(aux_rutmesprof_turnos)):
                        aux_rutprof_infoprof=np.where(rut_profesional[aux_rutmesprof_turnos[p]]==i_rut_profesional)[0]
                        
                        if len(aux_rutprof_infoprof)==0:
                            print('\n \n El rut ', (rut_profesional[aux_rutmesprof_turnos[p]]), 'no esta ingresado en info_profesionales.xlsx \n \n')
                            
                        ff.write( '\n')
                        ff.write( '                                                          ' )
                        ff.write('Visita' + str(format(p+1)) +  '\n' )
                        ff.write( '                                                          ' )
                        ff.write('Profesional = ' + str(format(i_nombre_profesional[aux_rutprof_infoprof[0]])) +  ' - ' )
                        ff.write('Rut = ' + str(format(i_rut_profesional[aux_rutprof_infoprof[0]])) +  '\n' )
                        ff.write( '                                                          ' )
                        ff.write('Fecha = ' + str(format(dia[aux_rutmesprof_turnos[p]])) +  '/' + str(format(mes[aux_rutmesprof_turnos[p]])) +  '/' + str(format(ano[aux_rutmesprof_turnos[p]])) +  '\n' )
                        ff.write( '                                                          ' )
                        ff.write('Pago Turno = ' + str(format(valores_profesional[aux_pospago_infoisapres[0]])))
                        ff.write( '\n')
                        ff.write( '                                                          ' )
                        ff.write('Cobro Turno = ' + str(format(valores_profesional[aux_poscobro_infoisapres[0]])))
                        ff.write( '\n')
                        ff.write( '\n')
                        
                    #Guardar valores a pagar por el mes completo
                    pagar_mes=pagar_mes+(sum(valores_profesional[aux_pospago_infoisapres]) * len(aux_rutmesprof_turnos))
                    cobrar_mes=cobrar_mes+(sum(valores_profesional[aux_poscobro_infoisapres]) * len(aux_rutmesprof_turnos))

                #Guardar informacion sobre pagos y cobros
                ff.write( '\n')
                ff.write( '\n')
                ff.write( '                                          .......................................................................................' )
                ff.write( '\n')
                ff.write( '                                          ' )
                ff.write('Total a pagar =' + str(format(int(pagar_mes))) +  ' ' )
                ff.write( '\n')
                ff.write( '                                          ' )
                ff.write('Cobrar a Isapre =' + str(format(int(cobrar_mes))) +  ' ' )
                ff.write( '\n')
                ff.write( '                                          ' )
                ff.write('Ganancia =' + str(format(int(cobrar_mes-pagar_mes))) +  ' ' )
                ff.write( '\n')
                ff.write( '\n')
            ff.write( '_________________________________________________________________________________')
            ff.write( '\n')
            ff.write( '_________________________________________________________________________________')
            ff.write( '\n')
            ff.write( '\n')
            
        else:
            for j in range(len(mes_nombre)):
            
                pagar_mes=0
                cobrar_mes=0
                
                aux_agno_turnos=np.where(ano==agno[gn])[0]
                #Posicion de turnos hecho ese mes en turnos.xlsx
                aux_mes_turnos=np.where(mes==(j+1))[0]
                #Posicion de rud paciente en turnos.xlsx
                aux_rut_turnos=np.where(rut_paciente[i]==rut_paciente_f)[0]
                #Posicion de turnos hecho para ese paciente en ese mes
                aux_rutmes_turnos=list(set(aux_rut_turnos) & set(aux_mes_turnos) & set(aux_agno_turnos))
                
                #Guardar informacion sobre el numero de visitas para ese paciente durante ese mes
                ff.write( '                    _________________________________________________________________________________________________' )
                ff.write( '\n')
                ff.write( '                    ' )
                ff.write('Visitas mes de '  + str(mes_nombre[j]) +  ' = ' + str(format(int(len(aux_rutmes_turnos)))) +  ' ' )
                ff.write( '\n')
                
                #Loop para tipos de profesionales
                for k in range(len(prof_unique_type)):
                    #Posicion de profesional seleccionado en turnos.xlsx
                    aux_profesional_turnos=np.where(prof_unique_type[k]==profesional)[0]
                    #Posicion de profesional seleccionado en turnos para ese paciente durante ese mes en turnos.xlsx
                    aux_rutmesprof_turnos=list(set(aux_rut_turnos) & set(aux_mes_turnos) & set(aux_profesional_turnos) & set(aux_agno_turnos))
                    #Posicion de la columna con valores del profesional en info_isapres.xlsx
                    valores_profesional=np.asarray(dfs3[prof_unique_type[k]].tolist())
                    #Posicion de la columna con numeros de visitas del profesional en info_pacientes.xlsx
                    nvisitas_profesional=np.asarray(dfs2[prof_unique_type[k]].tolist())
                    
                    #Alerta en el caso de que el turno no tenga un valor definido
                    if valores_profesional[aux_poscobro_infoisapres]==0:
                        ffwarning.write('No se ha ingresado el valor del procedimiento/profesional ' + str(prof_unique_type[k]))
                        ffwarning.write(' Paciente: ' + nombre_paciente[i] +  ' / Isapre: ' + isapre[i] + ' / Mes: ' + mes_nombre[j] )
                        ffwarning.write(' Porque el valor no esta especificado en info_isapres.xlsx')
                        ffwarning.write( '\n')
                    
                    #Guardar informacion sobre el numero de visitas de ese profesional, los pagos y cobros equivalentes
                    #ff.write( '                                          .......................................................................................' )
                    ff.write( '\n')
                    ff.write( '                                          ' )
                    ff.write('Visitas de '  + str(prof_unique_type[k]) +  ' = ' + str(format(int(len(aux_rutmesprof_turnos)))) + ' / ' +   str(nvisitas_profesional[i]))
                    if nvisitas_profesional[i]!=0:
                        ff.write(' (' +  str(format(int(len(aux_rutmesprof_turnos) * 100 / nvisitas_profesional[i]))) + '%)')
                    ff.write('   --------------------->   FALTAN =' + str(format(np.abs( int(len(aux_rutmesprof_turnos))   - nvisitas_profesional[i]) )) )
                    if nvisitas_profesional[i]!=0:
                        ff.write(' (' +  str(format(int( 100 - ( len(aux_rutmesprof_turnos) * 100 / nvisitas_profesional[i])))) + '%)')
                    ff.write( '\n')
                    ff.write( '                                          ' )
                    ff.write('Pago equivalente = ' + str(format(int(sum(valores_profesional[aux_pospago_infoisapres]) * len(aux_rutmesprof_turnos)))) +  '   -   ' )
                    ff.write('Cobro equivalente = ' + str(format(int(sum(valores_profesional[aux_poscobro_infoisapres]) * len(aux_rutmesprof_turnos)))) +  ' ' )
                    ff.write( '\n')
                    ff.write( '\n')
                    
                    #Loop para cantidad de visitas de ese rut, mes y profesional en turnos.xlsx
                    for p in range(len(aux_rutmesprof_turnos)):
                        aux_rutprof_infoprof=np.where(int(rut_profesional[aux_rutmesprof_turnos[p]])==i_rut_profesional)[0]
                        
                        if len(aux_rutprof_infoprof)==0:
                            print('El rut: ', rut_profesional[aux_rutmesprof_turnos[p]], ' no esta especificado en info_profesionales.xlsx.')

                        ff.write( '\n')
                        ff.write( '                                                          ' )
                        ff.write('Visita' + str(format(p+1)) +  '\n' )
                        ff.write( '                                                          ' )
                        ff.write('Profesional = ' + str(format(i_nombre_profesional[aux_rutprof_infoprof[0]])) +  ' - ' )
                        ff.write('Rut = ' + str(format(i_rut_profesional[aux_rutprof_infoprof[0]])) +  '\n' )
                        ff.write( '                                                          ' )
                        ff.write('Fecha = ' + str(format(dia[aux_rutmesprof_turnos[p]])) +  '/' + str(format(mes[aux_rutmesprof_turnos[p]])) +  '/' + str(format(ano[aux_rutmesprof_turnos[p]])) +  '\n' )
                        ff.write( '                                                          ' )

                        ff.write('Pago Turno = ' + str(format(valores_profesional[aux_pospago_infoisapres[0]])))
                        ff.write( '\n')
                        ff.write( '                                                          ' )
                        ff.write('Cobro Turno = ' + str(format(valores_profesional[aux_poscobro_infoisapres[0]])))
                        ff.write( '\n')
                        ff.write( '\n')
                        
                    #Guardar valores a pagar por el mes completo
                    pagar_mes=pagar_mes+(sum(valores_profesional[aux_pospago_infoisapres]) * len(aux_rutmesprof_turnos))
                    cobrar_mes=cobrar_mes+(sum(valores_profesional[aux_poscobro_infoisapres]) * len(aux_rutmesprof_turnos))

                #Guardar informacion sobre pagos y cobros
                ff.write( '\n')
                ff.write( '\n')
                ff.write( '                                          .......................................................................................' )
                ff.write( '\n')
                ff.write( '                                          ' )
                ff.write('Total a pagar =' + str(format(int(pagar_mes))) +  ' ' )
                ff.write( '\n')
                ff.write( '                                          ' )
                ff.write('Cobrar a Isapre =' + str(format(int(cobrar_mes))) +  ' ' )
                ff.write( '\n')
                ff.write( '                                          ' )
                ff.write('Ganancia =' + str(format(int(cobrar_mes-pagar_mes))) +  ' ' )
                ff.write( '\n')
                ff.write( '\n')
            ff.write( '_________________________________________________________________________________')
            ff.write( '\n')
            ff.write( '_________________________________________________________________________________')
            ff.write( '\n')
            ff.write( '\n')

#-----------------------------------------------------------------------------------------------------------
#Analisis Vitalsur_Profesionales
#-----------------------------------------------------------------------------------------------------------

#Identificar profesionales
rut_profesional_search=dfs['Rut profesional'].tolist()
prof_unique = list(set(rut_profesional_search))

for gn in range(len(agno)):
    #Loop para años
    #save_name1 = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Vitalsur_profesionales_%i.txt'%agno[gn])
    #save_name2 = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Yovanka_Whatsapp_profesionales_%i_%i.txt'%(agno[gn],(hoy.month-2)))
    #save_name3 = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'SII_profesionales_%i_%i.txt'%(agno[gn],(hoy.month-2)))

    save_name1 = os.path.join(os.path.curdir, "%i/"%agno[gn], 'Vitalsur_profesionales_%i.txt'%agno[gn])
    save_name2 = os.path.join(os.path.curdir, "%i/"%agno[gn], 'Yovanka_Whatsapp_profesionales_%i_%i.txt'%(agno[gn],(hoy.month-2)))
    save_name3 = os.path.join(os.path.curdir, "%i/"%agno[gn], 'SII_profesionales_%i_%i.txt'%(agno[gn],(hoy.month-2)))

    ff2 = open(save_name1, 'w')
    ff2current = open(save_name2, 'w')
    ff2SII = open(save_name3, 'w')
    
    #Loop para cada profesional
    for i in range(len(prof_unique)):
        #Posicion del rut del profesional en turnos.xlsx
        aux_rutprof_turnos=np.where(prof_unique[i]==rut_profesional)[0]
        #Posicion del rut del profesional en info_profesionales.xlsx
        aux_rutprof_infoprof=np.where(prof_unique[i]==i_rut_profesional)[0]
                
        if len(aux_rutprof_turnos)==0:
            print('El rut: ', prof_unique[i], ' no esta especificado en turnos.xlsx.')
                
        if len(aux_rutprof_infoprof)==0:
            print('El rut: ', prof_unique[i], ' no esta especificado en info_profesionales.xlsx.')
                
        #Guardar informacion del profesional (info_profesionales.xlsx)
        ff2.write( 'Nombre Profesional: ' + str(format(i_nombre_profesional[aux_rutprof_infoprof][0])) + '\n' )
        ff2.write( 'RUT Profesional: ' + str(format(int(i_rut_profesional[aux_rutprof_infoprof]))) + '\n' )
        ff2.write( 'Ciudad: ' + str(format(i_ciudadprof[aux_rutprof_infoprof][0])) + '\n'  )
        ff2.write( 'Ocupacion: ' + str(format(i_ocupacion[aux_rutprof_infoprof][0])) + '\n' )
        ff2.write( '\n')
        
        if agno[gn]==2021:
            #Loop para 12 meses
            for j in range(mes_actual):
                pagar_mes=0
                cobrar_mes=0
                
                aux_agno_turnos=np.where(ano==agno[gn])[0]
                #Posicion de turnos hecho ese mes (turnos.xlsx)
                aux_mes_turnos=np.where(mes==(j+1))[0]
                #Posicion de turnos hecho por ese profesional ese mes (turnos.xlsx)
                aux_rutprofmes_turnos=list(set(aux_rutprof_turnos) & set(aux_mes_turnos)  & set(aux_agno_turnos))
                
                #Guardar informacion sobre la cantidad de visitas que hizo el profesional en ese mes
                ff2.write( '                    _________________________________________________________________________________________________' )
                ff2.write( '\n')
                ff2.write( '                    ' )
                ff2.write('Visitas mes de '  + str(mes_nombre[j]) +  ' = ' + str(format(int(len(aux_rutprofmes_turnos)))) +  ' ' )
                ff2.write( '\n')
                ff2.write( '\n')
                
                if ((j+1) == mesdepago) and  (len(aux_rutprofmes_turnos)!=0):
                
                    ff2current.write( 'Nombre Profesional: ' + str(format(i_nombre_profesional[aux_rutprof_infoprof][0])) + '\n' )
                    ff2current.write( 'RUT Profesional: ' + str(format(int(i_rut_profesional[aux_rutprof_infoprof]))) + '\n' )
                    ff2current.write( 'Ciudad: ' + str(format(i_ciudadprof[aux_rutprof_infoprof][0])) + '\n'  )
                    ff2current.write( 'Ocupacion: ' + str(format(i_ocupacion[aux_rutprof_infoprof][0])) + '\n' )
                    ff2current.write( '\n')
                    ff2SII.write( str(format(int(i_rut_profesional[aux_rutprof_infoprof]))) + ' ' )

                    #Guardar informacion sobre la cantidad de visitas que hizo el profesional en ese mes
                    ff2current.write( '                    _________________________________________________________________________________________________' )
                    ff2current.write( '\n')
                    ff2current.write( '                    ' )
                    ff2current.write('Visitas mes de '  + str(mes_nombre[j]) +  ' = ' + str(format(int(len(aux_rutprofmes_turnos)))) +  ' ' )
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                
                pac_hecho = []
                pac_hecho_nombre = []
                #Loop para la cantidad de visitas hechas por ese profesional ese mes
                for r in range(len(aux_rutprofmes_turnos)):
                    #Posicion del rut del paciente del primer turno hecho por el profesional (info_pacientes.xlsx) para saber la isapre y ciudad
                    aux_rut_infopac=np.where(rut_paciente_f[aux_rutprofmes_turnos[r]]==rut_paciente)[0]
                    #Sacar valores de cobros y pagos dependiendo de la isapre usando isapre/ciudad/cobro_pago/profesional (info_isapres.xlsx)

                    aux_isapre_infoisapres=np.where(isapre[aux_rut_infopac]==i_isapre)[0]
                    aux_ciudad_dondevive=np.where(ciudad[aux_rut_infopac]==i_ciudad)[0]
                    aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)) #info isapres
                    if len(aux_isapreyciudad)==0:
                        aux_ciudad_dondevive=np.where('Otros'==i_ciudad)[0]
                    aux_sector_dondevive=np.where(sector[aux_rut_infopac]==i_sector)[0]
                    aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)  & set(aux_sector_dondevive)) #info isapres
                    
                    aux_cobro_infoisapres=np.where('Cobro'==i_cobropago)[0]
                    aux_pago_infoisapres=np.where('Pago'==i_cobropago)[0]
                        
                    #Posicion de la fila correspondiente al caso del paciente (info_isapres.xlsx)
                    aux_poscobro_infoisapres=list(set(aux_isapreyciudad) & set(aux_cobro_infoisapres))
                    aux_pospago_infoisapres=list(set(aux_isapreyciudad) & set(aux_pago_infoisapres))
                        
                    if len(aux_isapre_infoisapres)==0:
                        ffwarning.write('Hay una isapre que no esta especificada en info_isapres.xlsx')
                        ffwarning.write(' Check: '+ isapre[aux_rut_infopac] )
                        ffwarning.write( '\n')
                        print('WARNING3!')
                    
                    if len(aux_poscobro_infoisapres)>1:
                        ffwarning.write('Hay dos valores para el cobro del turno en info_isapres.xlsx')
                        ffwarning.write(' Check: '+ i_isapre[aux_isapre_infoisapres[0]] + ' ' + i_ciudad[aux_ciudad_dondevive[0]] +  ' COBRO')
                        ffwarning.write( '\n')
                    
                    if len(aux_pospago_infoisapres)>1:
                        ffwarning.write('Hay dos valores para el pago del turno en info_isapres.xlsx')
                        ffwarning.write(' Check: ' +  i_isapre[aux_isapre_infoisapres[0]] + ' '+ i_ciudad[aux_ciudad_dondevive[0]] +  ' PAGO')
                        ffwarning.write( '\n')

                    if len(aux_pospago_infoisapres)==0 or len(aux_poscobro_infoisapres)==0:
                        ffwarning.write('Hay un error en leer el cobro/pago de la isapre - info_isapres.xlsx')
                        ffwarning.write(' Check: '+ isapre[aux_rut_infopac] )
                        ffwarning.write( '\n')
                        print('WARNING4!')
                    
                    #Posicion de la columna con valores del profesional (info_isapres.xlsx)
                    valores_profesional=np.asarray(dfs3[profesional[aux_rutprofmes_turnos[r]]].tolist())
                    cobrar_mes=cobrar_mes+(valores_profesional[aux_poscobro_infoisapres])
                    
                    ff2.write( '                                          .......................................................................................' )
                    ff2.write( '\n')
                    ff2.write( '                                          ' )
                    ff2.write('Visita' + str(format(r+1)) +  '\n' )
                    ff2.write( '                                          ' )
                    ff2.write('Paciente = ' + str(format(nombre_paciente[aux_rut_infopac[0]])) +  ' - ' )
                    ff2.write('Rut = ' + str(format(rut_paciente[aux_rut_infopac[0]])) +  '\n' )
                    ff2.write( '                                          ' )
                    ff2.write('Isapre = ' + str(format(isapre[aux_rut_infopac[0]])) +  '\n' )
                    ff2.write( '                                          ' )
                    ff2.write('Fecha = ' + str(format(dia[aux_rutprofmes_turnos[r]])) +  '/' + str(format(mes[aux_rutprofmes_turnos[r]])) +  '/' + str(format(ano[aux_rutprofmes_turnos[r]])) +  '\n' )
                    ff2.write( '                                          ' )

                    detectar_prof=np.where(i_rut_profesional[aux_rutprof_infoprof]==PD_PROFESIONAL)[0]
                    detectar_pac=np.where(rut_paciente[aux_rut_infopac[0]]==PD_PACIENTE)[0]
                    detectar_prof_pac=list(set(detectar_prof) & set(detectar_pac))

                    if len(detectar_prof_pac)>0:
                        ff2.write('Pago = ' + str(format(PD_VALOR[detectar_prof_pac][0])))
                        pagar_mes=pagar_mes+PD_VALOR[detectar_prof_pac][0]
                        if len(detectar_prof_pac)>1:
                            print('Hay mas de un valor en pagos diferenciados para el profesional ', i_rut_profesional[aux_rutprof_infoprof], ' atendiendo a ', rut_paciente[aux_rut_infopac[0]] )
                    elif i_pagodif[aux_rutprof_infoprof]>10:
                        ff2.write('Pago = ' + str(format(i_pagodif[aux_rutprof_infoprof][0])))
                        pagar_mes=pagar_mes+i_pagodif[aux_rutprof_infoprof]
                    else:
                        ff2.write('Pago = ' + str(format(valores_profesional[aux_pospago_infoisapres[0]])))
                        pagar_mes=pagar_mes+(valores_profesional[aux_pospago_infoisapres])
                    
                    ff2.write( '\n')
                    ff2.write( '\n')
                    
                    if ((j+1) == mesdepago) and  (len(aux_rutprofmes_turnos)!=0):
                        ff2current.write( '\n')
                        ff2current.write( '                                          ' )
                        ff2current.write('Visita' + str(format(r+1)) +  '\n' )
                        ff2current.write( '                                          ' )
                        ff2current.write('Paciente = ' + str(format(nombre_paciente[aux_rut_infopac[0]])) +  ' - ' )
                        ff2current.write('Rut = ' + str(format(rut_paciente[aux_rut_infopac[0]])) +  '\n' )
                        ff2current.write( '                                          ' )
                        ff2current.write('Isapre = ' + str(format(isapre[aux_rut_infopac[0]])) +  '\n' )
                        ff2current.write( '                                          ' )
                        ff2current.write('Fecha = ' + str(format(dia[aux_rutprofmes_turnos[r]])) +  '/' + str(format(mes[aux_rutprofmes_turnos[r]])) +  '/' + str(format(ano[aux_rutprofmes_turnos[r]])) +  '\n' )
                        ff2current.write( '                                          ' )
                        if len(detectar_prof_pac)>0:
                            ff2current.write('Pago = ' + str(format(PD_VALOR[detectar_prof_pac][0])))
                            if len(detectar_prof_pac)>1:
                                print('Hay mas de un valor en pagos diferenciados para el profesional ', i_rut_profesional[aux_rutprof_infoprof], ' atendiendo a ', rut_paciente[aux_rut_infopac[0]] )
                        elif i_pagodif[aux_rutprof_infoprof]>10:
                            ff2current.write('Pago = ' + str(format(i_pagodif[aux_rutprof_infoprof][0])))
                        else:
                            ff2current.write('Pago = ' + str(format(valores_profesional[aux_pospago_infoisapres[0]])))
                        ff2current.write( '\n')
                        ff2current.write( '\n')
                        
                        pac_hecho.append((rut_paciente[aux_rut_infopac[0]]))
                        pac_hecho_nombre.append((nombre_paciente[aux_rut_infopac[0]]))

                ff2.write( '\n')
                ff2.write( '\n')
                ff2.write( '                                          .......................................................................................' )
                ff2.write( '\n')
                ff2.write( '                                          ' )
                ff2.write('Total a pagar = ' + str(format(int(pagar_mes))) +  ' ' )
                ff2.write( '\n')
                ff2.write( '\n')
                

                if ((j+1) == mesdepago) and  (len(aux_rutprofmes_turnos)!=0):
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                    ff2current.write( '                                          .......................................................................................' )
                    ff2current.write( '\n')
                    ff2current.write( '                                          ' )
                    ff2current.write('Total a pagar = ' + str(format(int(pagar_mes))) +  ' ' )
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                
                    ff2current.write( '..................................................................................................' )
                    ff2current.write( '\n')
                    ff2current.write('Whatsapp: ')
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                    ff2current.write('Estimada/o ' + str(format(i_nombre_profesional[aux_rutprof_infoprof][0])) + ', enviar boleta de honorarios de las visitas correspondientes al mes de ' + str(mes_nombre[j]) + ': \n' )
                    
                    for mm in range(len(np.unique(pac_hecho))):
                        pos_pac=np.where(np.unique(pac_hecho)[mm]==pac_hecho)[0]
                        ff2current.write(  str(format(pac_hecho_nombre[pos_pac[0]])) +  ': ' + str(format(len(pos_pac))) + ' visitas \n' )
                    ff2current.write('Saludos  \n' )
                    
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                    ff2current.write( '_________________________________________________________________________________________________' )
                    ff2current.write( '\n')
                    ff2current.write( '_________________________________________________________________________________________________' )
                    ff2current.write( '\n')
                    
                    ff2SII.write( str(format(int(pagar_mes))) + ' '  )
                    ff2SII.write( str(format(int(mesdepago)))  + ' \n'  )


                    
            ff2.write( '\n')
            ff2.write( '\n')
            ff2.write( '_________________________________________________________________________________________________' )
            ff2.write( '\n')
            ff2.write( '_________________________________________________________________________________________________' )
            ff2.write( '\n')
        
        else:
            #Loop para 12 meses
            for j in range(len(mes_nombre)):
                pagar_mes=0
                cobrar_mes=0

                aux_agno_turnos=np.where(ano==agno[gn])[0]
                #Posicion de turnos hecho ese mes (turnos.xlsx)
                aux_mes_turnos=np.where(mes==(j+1))[0]
                #Posicion de turnos hecho por ese profesional ese mes (turnos.xlsx)
                aux_rutprofmes_turnos=list(set(aux_rutprof_turnos) & set(aux_mes_turnos) & set(aux_agno_turnos))
                
                #Guardar informacion sobre la cantidad de visitas que hizo el profesional en ese mes
                ff2.write( '                    _________________________________________________________________________________________________' )
                ff2.write( '\n')
                ff2.write( '                    ' )
                ff2.write('Visitas mes de '  + str(mes_nombre[j]) +  ' = ' + str(format(int(len(aux_rutprofmes_turnos)))) +  ' ' )
                ff2.write( '\n')
                ff2.write( '\n')
                
                if ((j+1) == mesdepago) and  (len(aux_rutprofmes_turnos)!=0):
                    ff2current.write( 'Nombre Profesional: ' + str(format(i_nombre_profesional[aux_rutprof_infoprof][0])) + '\n' )
                    ff2current.write( 'RUT Profesional: ' + str(format(int(i_rut_profesional[aux_rutprof_infoprof]))) + '\n' )
                    ff2current.write( 'Ciudad: ' + str(format(i_ciudadprof[aux_rutprof_infoprof][0])) + '\n'  )
                    ff2current.write( 'Ocupacion: ' + str(format(i_ocupacion[aux_rutprof_infoprof][0])) + '\n' )
                    ff2current.write( '\n')
                    
                    #ff2current.write( str(format(i_nombre_profesional[aux_rutprof_infoprof][0])) + ' ' )
                    ff2SII.write( str(format(int(i_rut_profesional[aux_rutprof_infoprof]))) + ' ' )

                    #Guardar informacion sobre la cantidad de visitas que hizo el profesional en ese mes
                    ff2current.write( '                    _________________________________________________________________________________________________' )
                    ff2current.write( '\n')
                    ff2current.write( '                    ' )
                    ff2current.write('Visitas mes de '  + str(mes_nombre[j]) +  ' = ' + str(format(int(len(aux_rutprofmes_turnos)))) +  ' ' )
                    ff2current.write( '\n')
                    ff2current.write( '\n')

                pac_hecho = []
                pac_hecho_nombre = []
                #Loop para la cantidad de visitas hechas por ese profesional ese mes
                for r in range(len(aux_rutprofmes_turnos)):
                    #Posicion del rut del paciente del primer turno hecho por el profesional (info_pacientes.xlsx) para saber la isapre y ciudad
                    aux_rut_infopac=np.where(rut_paciente_f[aux_rutprofmes_turnos[r]]==rut_paciente)[0]

                    if len(aux_rut_infopac)==0:
                        print('El rut: ', rut_paciente_f[aux_rutprofmes_turnos[r]], ' no esta especificado en info_pacientes.xlsx.')
                    #Sacar valores de cobros y pagos dependiendo de la isapre usando isapre/ciudad/cobro_pago/profesional (info_isapres.xlsx)
                    aux_isapre_infoisapres=np.where(isapre[aux_rut_infopac]==i_isapre)[0]
                    aux_ciudad_dondevive=np.where(ciudad[aux_rut_infopac]==i_ciudad)[0]
                    aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)) #info isapres
                    if len(aux_isapreyciudad)==0:
                        aux_ciudad_dondevive=np.where('Otros'==i_ciudad)[0]
                    aux_sector_dondevive=np.where(sector[aux_rut_infopac]==i_sector)[0]
                    aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)  & set(aux_sector_dondevive)) #info isapres
                    
                    aux_cobro_infoisapres=np.where('Cobro'==i_cobropago)[0]
                    aux_pago_infoisapres=np.where('Pago'==i_cobropago)[0]
                        
                    #Posicion de la fila correspondiente al caso del paciente (info_isapres.xlsx)
                    aux_poscobro_infoisapres=list(set(aux_isapreyciudad) & set(aux_cobro_infoisapres))
                    aux_pospago_infoisapres=list(set(aux_isapreyciudad) & set(aux_pago_infoisapres))
                        
                    if len(aux_isapre_infoisapres)==0:
                        ffwarning.write('Hay una isapre que no esta especificada en info_isapres.xlsx')
                        ffwarning.write(' Check: '+ isapre[aux_rut_infopac] )
                        ffwarning.write( '\n')
                        print('WARNING5!')
                    
                    if len(aux_poscobro_infoisapres)>1:
                        ffwarning.write('Hay dos valores para el cobro del turno en info_isapres.xlsx')
                        ffwarning.write(' Check: '+ i_isapre[aux_isapre_infoisapres[0]] + ' ' + i_ciudad[aux_ciudad_dondevive[0]] +  ' COBRO')
                        ffwarning.write( '\n')
                    
                    if len(aux_pospago_infoisapres)>1:
                        ffwarning.write('Hay dos valores para el pago del turno en info_isapres.xlsx')
                        ffwarning.write(' Check: ' +  i_isapre[aux_isapre_infoisapres[0]] + ' '+ i_ciudad[aux_ciudad_dondevive[0]] +  ' PAGO')
                        ffwarning.write( '\n')

                    if len(aux_pospago_infoisapres)==0 or len(aux_poscobro_infoisapres)==0:
                        ffwarning.write('Hay un error en leer el cobro/pago de la isapre - info_isapres.xlsx')
                        ffwarning.write(' Check: '+ isapre[aux_rut_infopac] )
                        ffwarning.write( '\n')
                        print('WARNING6!')
                    
                    #Posicion de la columna con valores del profesional (info_isapres.xlsx)
                    valores_profesional=np.asarray(dfs3[profesional[aux_rutprofmes_turnos[r]]].tolist())
                    cobrar_mes=cobrar_mes+(valores_profesional[aux_poscobro_infoisapres])
                    
                    ff2.write( '                                          .......................................................................................' )
                    ff2.write( '\n')
                    ff2.write( '                                          ' )
                    ff2.write('Visita' + str(format(r+1)) +  '\n' )
                    ff2.write( '                                          ' )
                    ff2.write('Paciente = ' + str(format(nombre_paciente[aux_rut_infopac[0]])) +  ' - ' )
                    ff2.write('Rut = ' + str(format(rut_paciente[aux_rut_infopac[0]])) +  '\n' )
                    ff2.write( '                                          ' )
                    ff2.write('Isapre = ' + str(format(isapre[aux_rut_infopac[0]])) +  '\n' )
                    ff2.write( '                                          ' )
                    ff2.write('Fecha = ' + str(format(dia[aux_rutprofmes_turnos[r]])) +  '/' + str(format(mes[aux_rutprofmes_turnos[r]])) +  '/' + str(format(ano[aux_rutprofmes_turnos[r]])) +  '\n' )
                    ff2.write( '                                          ' )
                    
                    detectar_prof=np.where(i_rut_profesional[aux_rutprof_infoprof]==PD_PROFESIONAL)[0]
                    detectar_pac=np.where(rut_paciente[aux_rut_infopac[0]]==PD_PACIENTE)[0]
                    detectar_prof_pac=list(set(detectar_prof) & set(detectar_pac))

                    if len(detectar_prof_pac)>0:
                        ff2.write('Pago = ' + str(format(PD_VALOR[detectar_prof_pac][0])))
                        pagar_mes=pagar_mes+PD_VALOR[detectar_prof_pac][0]
                        if len(detectar_prof_pac)>1:
                            print('Hay mas de un valor en pagos diferenciados para el profesional ', i_rut_profesional[aux_rutprof_infoprof], ' atendiendo a ', rut_paciente[aux_rut_infopac[0]] )
                    elif i_pagodif[aux_rutprof_infoprof]>10:
                        ff2.write('Pago = ' + str(format(i_pagodif[aux_rutprof_infoprof][0])))
                        pagar_mes=pagar_mes+i_pagodif[aux_rutprof_infoprof]
                    else:
                        ff2.write('Pago = ' + str(format(valores_profesional[aux_pospago_infoisapres[0]])))
                        pagar_mes=pagar_mes+(valores_profesional[aux_pospago_infoisapres])
                    ff2.write( '\n')
                    ff2.write( '\n')
                        
                    
                    if ((j+1) == mesdepago) and  (len(aux_rutprofmes_turnos)!=0):
                        ff2current.write( '\n')
                        ff2current.write( '                                          ' )
                        ff2current.write('Visita' + str(format(r+1)) +  '\n' )
                        ff2current.write( '                                          ' )
                        ff2current.write('Paciente = ' + str(format(nombre_paciente[aux_rut_infopac[0]])) +  ' - ' )
                        ff2current.write('Rut = ' + str(format(rut_paciente[aux_rut_infopac[0]])) +  '\n' )
                        ff2current.write( '                                          ' )
                        ff2current.write('Isapre = ' + str(format(isapre[aux_rut_infopac[0]])) +  '\n' )
                        ff2current.write( '                                          ' )
                        ff2current.write('Fecha = ' + str(format(dia[aux_rutprofmes_turnos[r]])) +  '/' + str(format(mes[aux_rutprofmes_turnos[r]])) +  '/' + str(format(ano[aux_rutprofmes_turnos[r]])) +  '\n' )
                        ff2current.write( '                                          ' )
                        
                        if len(detectar_prof_pac)>0:
                            ff2current.write('Pago = ' + str(format(PD_VALOR[detectar_prof_pac][0])))
                            if len(detectar_prof_pac)>1:
                                print('Hay mas de un valor en pagos diferenciados para el profesional ', i_rut_profesional[aux_rutprof_infoprof], ' atendiendo a ', rut_paciente[aux_rut_infopac[0]] )
                        elif i_pagodif[aux_rutprof_infoprof]>10:
                            ff2current.write('Pago = ' + str(format(i_pagodif[aux_rutprof_infoprof][0])))
                        else:
                            ff2current.write('Pago = ' + str(format(valores_profesional[aux_pospago_infoisapres[0]])))
                        ff2current.write( '\n')
                        ff2current.write( '\n')
                        
                        pac_hecho.append((rut_paciente[aux_rut_infopac[0]]))
                        pac_hecho_nombre.append((nombre_paciente[aux_rut_infopac[0]]))

                ff2.write( '\n')
                ff2.write( '\n')
                ff2.write( '                                          .......................................................................................' )
                ff2.write( '\n')
                ff2.write( '                                          ' )
                ff2.write('Total a pagar = ' + str(format(int(pagar_mes))) +  ' ' )
                ff2.write( '\n')
                ff2.write( '\n')
                

                if ((j+1) == mesdepago) and  (len(aux_rutprofmes_turnos)!=0):
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                    ff2current.write( '                                          .......................................................................................' )
                    ff2current.write( '\n')
                    ff2current.write( '                                          ' )
                    ff2current.write('Total a pagar = ' + str(format(int(pagar_mes))) +  ' ' )
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                
                    ff2current.write( '..................................................................................................' )
                    ff2current.write( '\n')
                    ff2current.write('Whatsapp: ')
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                    ff2current.write('Estimada/o ' + str(format(i_nombre_profesional[aux_rutprof_infoprof][0])) + ', enviar boleta de honorarios de las visitas correspondientes al mes de ' + str(mes_nombre[j]) + ': \n' )
                    
                    for mm in range(len(np.unique(pac_hecho))):
                        pos_pac=np.where(np.unique(pac_hecho)[mm]==pac_hecho)[0]
                        ff2current.write(  str(format(pac_hecho_nombre[pos_pac[0]])) +  ': ' + str(format(len(pos_pac))) + ' visitas \n' )
                    ff2current.write('Saludos  \n' )
                    
                    ff2current.write( '\n')
                    ff2current.write( '\n')
                    ff2current.write( '_________________________________________________________________________________________________' )
                    ff2current.write( '\n')
                    ff2current.write( '_________________________________________________________________________________________________' )
                    ff2current.write( '\n')
                    
                    ff2SII.write( str(format(int(pagar_mes))) + ' '  )
                    ff2SII.write( str(format(int(mesdepago)))  + ' \n'  )


                    
            ff2.write( '\n')
            ff2.write( '\n')
            ff2.write( '_________________________________________________________________________________________________' )
            ff2.write( '\n')
            ff2.write( '_________________________________________________________________________________________________' )
            ff2.write( '\n')




#-----------------------------------------------------------------------------------------------------------
#Analisis Vitalsur_Isapre
#-----------------------------------------------------------------------------------------------------------

#Identificar isapres
isapre_search=dfs2['Isapre'].tolist()
isapre_unique = list(set(isapre_search))
#Loop para todas las isapres

for gn in range(len(agno)):
    #save_name1 = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Vitalsur_Isapre_Fundacion_%i.txt'%agno[gn])
    #save_name2 = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Jessica_mails_bonos_primeraquincena_%i.txt'%agno[gn])
    #save_name3 = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Jessica_mails_bonos_segundaquincena_%i.txt'%agno[gn])
    
    save_name1 = os.path.join(os.path.curdir, "%i/"%agno[gn], 'Vitalsur_Isapre_Fundacion_%i.txt'%agno[gn])
    save_name2 = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Jessica_mails_bonos_primeraquincena_%i.txt'%agno[gn])
    save_name3 = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Jessica_mails_bonos_segundaquincena_%i.txt'%agno[gn])
    
    save_name2_RES = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Jessica_mails_bonos_atrasados_primeraquincena_%i.txt'%agno[gn])
    save_name3_RES = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Jessica_mails_bonos_atrasados_segundaquincena_%i.txt'%agno[gn])
    
    save_name2_montos = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Valores_bonos_primeraquincena_%i.txt'%agno[gn])
    save_name3_montos = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Valores_bonos_segundaquincena_%i.txt'%agno[gn])
    
    save_name2_montos_RES = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Valores_bonos_atrasados_primeraquincena_%i.txt'%agno[gn])
    save_name3_montos_RES = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Valores_bonos_atrasados_segundaquincena_%i.txt'%agno[gn])
    
    save_name_gesmed = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Planilla_Gesmed_%i.xlsx'%(agno[gn]))
    
    workbook = xlsxwriter.Workbook(save_name_gesmed)
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0

    ff4 = open(save_name1, 'w')
    ff4mail = open(save_name2, 'w')
    ff4mailsq = open(save_name3, 'w')
    
    ff4mail_RES = open(save_name2_RES, 'w')
    ff4mailsq_RES = open(save_name3_RES, 'w')
    
    ff4valores = open(save_name2_montos, 'w')
    ff4valoressq = open(save_name3_montos, 'w')
    
    ff4valores_RES = open(save_name2_montos_RES, 'w')
    ff4valoressq_RES = open(save_name3_montos_RES, 'w')
    
    #ffgesmed = open(save_name_gesmed, 'w')
    
    for i in range(len(isapre_unique)):
        if isapre_unique[i]=='Fundacion':
            #Posicion para pacientes con esa isapre (info_pacientes.xlsx)
            aux_isapre_infopacientes=np.where(isapre_unique[i]==isapre)[0]

            #Guardar informacion sobre isapre
            ff4.write( 'Isapre: ' + str(isapre_unique[i]) + ' ' )
            ff4.write( '\n')
            
            #Loop para cada paciente que se atiende con esa isapre (info_pacientes.xlsx)
            for l in range(len(aux_isapre_infopacientes)):
                #Guardar informacion sobre el paciente (info_pacientes.xlsx)
                ff4.write( '                      .......................................................................................' )
                ff4.write( '\n')
                ff4.write( '                      ' )
                ff4.write('Paciente = ' + str(format(nombre_paciente[aux_isapre_infopacientes[l]])) +  ' - ' )
                ff4.write('Rut = ' + str(format(rut_paciente[aux_isapre_infopacientes[l]])) +  '\n' )
                ff4.write( '                      ' )
                ff4.write('Ciudad = ' + str(format(ciudad[aux_isapre_infopacientes[l]])) +  '\n' )
                ff4.write( '\n')
                
                dd=str(format(dia_ingreso[aux_isapre_infopacientes[l]]))
                mm=str(format(mes_ingreso[aux_isapre_infopacientes[l]]))
                aa=str(format(ano_ingreso[aux_isapre_infopacientes[l]]))

                
                startdate_excel=dd+'/'+mm+'/'+aa
                startdate = pd.to_datetime(startdate_excel)
                enddate = pd.to_datetime(startdate_excel) + pd.DateOffset(days=30)
                
                while enddate<hoy:
                    if enddate.month!=startdate.month:
                        finm1=np.where(mes==startdate.month)[0]
                        finm2=np.where(dia>=startdate.day)[0]
                        finm3=np.where(ano==startdate.year)[0]
                        finm4=np.where(ano==agno[gn])[0]
                        finm_tot=list(set(finm1) & set(finm2)  & set(finm3) & set(finm4))
                        
                        comm1=np.where(mes==enddate.month)[0]
                        comm2=np.where(dia<enddate.day)[0]
                        comm3=np.where(ano==enddate.year)[0]
                        finm5=np.where(ano==agno[gn])[0]
                        comm_tot=list(set(comm1) & set(comm2)  & set(comm3) & set(finm5))
                        
                        aux_mes_turnos= list(set(finm_tot) | set(comm_tot))
                        
                        
                    elif enddate.year!=startdate.year:
                        finm1=np.where(mes==startdate.month)[0]
                        finm2=np.where(dia>=startdate.day)[0]
                        finm3=np.where(ano==startdate.year)[0]
                        finm4=np.where(ano==agno[gn])[0]
                        finm_tot=list(set(finm1) & set(finm2)  & set(finm3) & set(finm4))
                        
                        comm1=np.where(mes==enddate.month)[0]
                        comm2=np.where(dia<enddate.day)[0]
                        comm3=np.where(ano==enddate.year)[0]
                        finm5=np.where(ano==agno[gn]+1)[0]
                        comm_tot=list(set(comm1) & set(comm2)  & set(comm3)  & set(finm5))
                        
                        aux_mes_turnos= list(set(finm_tot) | set(comm_tot))
                        
                    else:
                        finm1=np.where(mes==startdate.month)[0]
                        finm2=np.where(dia>=startdate.day)[0]
                        comm2=np.where(dia<enddate.day)[0]
                        finm3=np.where(ano==startdate.year)[0]
                        finm4=np.where(ano==agno[gn])[0]
                        aux_mes_turnos=list(set(finm1) & set(finm2) & set(finm3) & set(comm2) & set(finm4) )
                
                    cobrar_mes_paciente=0
                    #Turnos que se hicieron para ese paciente (turnos.xlsx)
                    aux_rut_turnos=np.where(rut_paciente_f==  rut_paciente[aux_isapre_infopacientes[l]])[0]
                    #Turnos que se hicieron para ese paciente durante ese mes (turnos.xlsx)
                    aux_mesrut_turnos=list(set(aux_mes_turnos) & set(aux_rut_turnos))
                    #Entrar si hay algun turno hecho para ese paciente

                    if len(aux_mesrut_turnos)>0:
                        ff4.write( '                              ' )
                        ff4.write( '------------------------------------------------------------------ \n')
                        ff4.write( '                              ' )
                        ff4.write('Desde = ' + str(format(startdate)) +  '   -    ' )
                        ff4.write('Hasta = ' + str(format(enddate- pd.DateOffset(days=1))) +  '\n' )
                        ff4.write( '                              ' )
                        ff4.write( '------------------------------------------------------------------ \n')
                        
                        #Loop para la cantidad de turnos hechos para ese paciente durante ese mes
                        for q in range(len(aux_mesrut_turnos)):
                            #Columna de pagos y cobros para el profesional de ese turno
                            valores_profesional=np.asarray(dfs3[profesional[aux_mesrut_turnos[q]]].tolist())
                      
                            #Sacar valores de cobros y pagos dependiendo de la isapre usando isapre/ciudad/cobro_pago/profesional (info_isapres.xlsx)
                            aux_isapre_infoisapre=np.where(isapre_unique[i]==i_isapre)[0]
                            if len(aux_isapre_infoisapre)==0:
                                ffwarning.write('Hay una isapre que no esta especificada en info_isapres.xlsx')
                                ffwarning.write(' Check: '+ isapre[aux_rut_infopac] )
                                ffwarning.write( '\n')
                                
                            aux_ciudad_dondevive=np.where(ciudad[aux_isapre_infopacientes[l]]==i_ciudad)[0]
                            if len(aux_ciudad_dondevive)==0:
                                aux_ciudad_dondevive=np.where('Otros'==i_ciudad)[0]
                            aux_isapreyciudad=list(set(aux_isapre_infoisapre) & set(aux_ciudad_dondevive)) #info isapres
                            
                            #if len(aux_isapreyciudad)==0:
                            #    aux_ciudad_dondevive=np.where('Otros'==i_ciudad)[0]
                            #    aux_isapreyciudad=list(set(aux_isapre_infoisapre) & set(aux_ciudad_dondevive))
                            
                            aux_cobro_infoisapres=np.where('Cobro'==i_cobropago)[0]
                            aux_pago_infoisapres=np.where('Pago'==i_cobropago)[0]
                            #Posicion de la fila correspondiente al caso del paciente (info_isapres.xlsx)
                            aux_poscobro_infoisapres=list(set(aux_isapreyciudad) & set(aux_cobro_infoisapres))
                            aux_pospago_infoisapres=list(set(aux_isapreyciudad) & set(aux_pago_infoisapres))
                            
                            cobrar_mes=cobrar_mes+(valores_profesional[aux_poscobro_infoisapres])
                            pagar_mes=pagar_mes+(valores_profesional[aux_pospago_infoisapres])
                        
                            cobrar_mes_paciente=cobrar_mes_paciente+(valores_profesional[aux_poscobro_infoisapres])
                            cobrar_mes=cobrar_mes+(valores_profesional[aux_poscobro_infoisapres])
                            
                            ff4.write( '                              ' )
                            ff4.write('Visita' + str(format(q+1)) +  '\n' )
                            ff4.write( '                              ' )
                            ff4.write('Tipo Visita = ' + str(format(profesional[aux_mesrut_turnos[q]])) +  '\n' )
                            ff4.write( '                              ' )
                            ff4.write('Profesional = ' + str(format(rut_profesional[aux_mesrut_turnos[q]])) +  '\n' )
                            ff4.write( '                              ' )
                            ff4.write('Fecha = ' + str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])) +  '\n' )
                            ff4.write( '                              ' )
                            ff4.write('Cobro = ' + str(format((valores_profesional[aux_poscobro_infoisapres[0]]))) +  '\n' )
                            ff4.write( '\n')
                        ff4.write( '                              ' )
                        ff4.write('Cobrar a Isapre por paciente =' + str(format(int(cobrar_mes_paciente))) +  ' ' )
                        ff4.write( '\n')
                        ff4.write( '\n')
                        
                        
                        
                    startdate=enddate
                    enddate= startdate + pd.DateOffset(days=30)
    
        else:
            if agno[gn]==2021:
    
                pmes=np.zeros(12)
                pMedico=np.zeros(12)
                pEnfermera=np.zeros(12)
                pNutricionista=np.zeros(12)
                pKinesiologo=np.zeros(12)
                pPsicologo=np.zeros(12)
                pFono=np.zeros(12)
                pTO=np.zeros(12)
                pTens=np.zeros(12)
                pCuracion_Simple=np.zeros(12)
                pEducacion_Enfermera=np.zeros(12)
                pEducacion_Tens=np.zeros(12)
                pIntervension_Psicosocial=np.zeros(12)
                
                #save_name = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Vitalsur_Isapre_%s_%i.txt' %(isapre_unique[i],agno[gn]))
                save_name = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Vitalsur_Isapre_%s_%i.txt' %(isapre_unique[i],agno[gn]))

                ff3 = open(save_name, 'w')
                
                ff4mail.write('_________________________________________________________________________________________________ \n')
                ff4mail.write('_________________________________________________________________________________________________ \n')
                ff4mail.write('Mail isapre '  + str(isapre_unique[i]) + '\n')
                ff4mail.write('_________________________________________________________________________________________________ \n')
                ff4mail.write('_________________________________________________________________________________________________ \n \n')
                ff4mailsq.write('_________________________________________________________________________________________________ \n')
                ff4mailsq.write('_________________________________________________________________________________________________ \n')
                ff4mailsq.write('Mail isapre '  + str(isapre_unique[i]) + '\n')
                ff4mailsq.write('_________________________________________________________________________________________________ \n')
                ff4mailsq.write('_________________________________________________________________________________________________ \n \n')
                
                ff4mail_RES.write('_________________________________________________________________________________________________ \n')
                ff4mail_RES.write('_________________________________________________________________________________________________ \n')
                ff4mail_RES.write('Mail isapre '  + str(isapre_unique[i]) + '\n')
                ff4mail_RES.write('_________________________________________________________________________________________________ \n')
                ff4mail_RES.write('_________________________________________________________________________________________________ \n \n')
                ff4mailsq_RES.write('_________________________________________________________________________________________________ \n')
                ff4mailsq_RES.write('_________________________________________________________________________________________________ \n')
                ff4mailsq_RES.write('Mail isapre '  + str(isapre_unique[i]) + '\n')
                ff4mailsq_RES.write('_________________________________________________________________________________________________ \n')
                ff4mailsq_RES.write('_________________________________________________________________________________________________ \n \n')
                
                ff4valores.write('_________________________________________________________________________________________________ \n')
                ff4valores.write('_________________________________________________________________________________________________ \n')
                ff4valores.write('Valores isapre '  + str(isapre_unique[i]) + '\n')
                ff4valores.write('_________________________________________________________________________________________________ \n')
                ff4valores.write('_________________________________________________________________________________________________ \n \n')
                ff4valoressq.write('_________________________________________________________________________________________________ \n')
                ff4valoressq.write('_________________________________________________________________________________________________ \n')
                ff4valoressq.write('Valores isapre '  + str(isapre_unique[i]) + '\n')
                ff4valoressq.write('_________________________________________________________________________________________________ \n')
                ff4valoressq.write('_________________________________________________________________________________________________ \n \n')
                
                ff4valores_RES.write('_________________________________________________________________________________________________ \n')
                ff4valores_RES.write('_________________________________________________________________________________________________ \n')
                ff4valores_RES.write('Valores isapre '  + str(isapre_unique[i]) + '\n')
                ff4valores_RES.write('_________________________________________________________________________________________________ \n')
                ff4valores_RES.write('_________________________________________________________________________________________________ \n \n')
                ff4valoressq_RES.write('_________________________________________________________________________________________________ \n')
                ff4valoressq_RES.write('_________________________________________________________________________________________________ \n')
                ff4valoressq_RES.write('Valores isapre '  + str(isapre_unique[i]) + '\n')
                ff4valoressq_RES.write('_________________________________________________________________________________________________ \n')
                ff4valoressq_RES.write('_________________________________________________________________________________________________ \n \n')
                
                
                #Posicion para pacientes con esa isapre (info_pacientes.xlsx)
                aux_isapre_infopacientes=np.where(isapre_unique[i]==isapre)[0]

                #Guardar informacion sobre isapre
                ff3.write( 'Isapre: ' + str(isapre_unique[i]) + ' ' )
                ff3.write( '\n')

                #Loop para 12  meses
                for j in range(mes_actual):
                    cobrar_mes=0
                    
                    pmes[j]=j+1
                    aux_agno_turnos=np.where(ano==agno[gn])[0]
                    #Turnos que se hicieron ese mes
                    aux_mes_turnos=np.where(mes==(j+1))[0]
                    
                    #Guardar nombre del mes
                    ff3.write( '_________________________________________________________________________________________________' )
                    ff3.write( '\n')
                    ff3.write( '_________________________________________________________________________________________________' )
                    ff3.write( '\n')
                    ff3.write( '' )
                    ff3.write('Visitas mes de '  + str(mes_nombre[j]) )
                    ff3.write( '\n')
                    ff3.write( '\n')
                    

                    ff4mail.write( '________Primera quincena mes de '  + str(mes_nombre[j]) + '________')
                    ff4mail.write( '\n')
                    ff4mail.write( '\n')
                    ff4mail.write( '\n')
                    ff4mail.write('Estimada/o, \n Junto con saludarla/o, y esperando se encuentre muy bien, me dirijo a usted para solicitar ayuda en el cobro de los bonos de la primera quincena del mes de ' )
                    ff4mail.write(str(mes_nombre[j]) )
                    ff4mail.write(' por la atención de los siguientes pacientes (adjunto planillas): \n')
                    
                    
                    ff4mailsq.write( '________Segunda quincena mes de '  + str(mes_nombre[j]) + '________')
                    ff4mailsq.write( '\n')
                    ff4mailsq.write( '\n')
                    ff4mailsq.write( '\n')
                    ff4mailsq.write('Estimada/o, \n Junto con saludarla/o, y esperando se encuentre muy bien, me dirijo a usted para solicitar ayuda en el cobro de los bonos de la segunda quincena del mes de ' )
                    ff4mailsq.write(str(mes_nombre[j]) )
                    ff4mailsq.write(' por la atención de los siguientes pacientes (adjunto planillas): \n')
                    
                    ff4mail_RES.write( '________Primera quincena mes de '  + str(mes_nombre[j]) + '________')
                    ff4mail_RES.write( '\n')
                    ff4mail_RES.write( '\n')
                    ff4mail_RES.write( '\n')
                    ff4mail_RES.write('Estimada/o, \n Junto con saludarla/o, y esperando se encuentre muy bien, me dirijo a usted para solicitar ayuda en el cobro de los bonos de la primera quincena del mes de ' )
                    ff4mail_RES.write(str(mes_nombre[j]) )
                    ff4mail_RES.write(' por la atención de los siguientes pacientes (adjunto planillas): \n')
                    
                    
                    ff4mailsq_RES.write( '________Segunda quincena mes de '  + str(mes_nombre[j]) + '________')
                    ff4mailsq_RES.write( '\n')
                    ff4mailsq_RES.write( '\n')
                    ff4mailsq_RES.write( '\n')
                    ff4mailsq_RES.write('Estimada/o, \n Junto con saludarla/o, y esperando se encuentre muy bien, me dirijo a usted para solicitar ayuda en el cobro de los bonos de la segunda quincena del mes de ' )
                    ff4mailsq_RES.write(str(mes_nombre[j]) )
                    ff4mailsq_RES.write(' por la atención de los siguientes pacientes (adjunto planillas): \n')
                    
                    
                    #Loop para cada paciente que se atiende con esa isapre (info_pacientes.xlsx)
                    for l in range(len(aux_isapre_infopacientes)):
                        cobrar_mes_paciente=0
                        #Turnos que se hicieron para ese paciente (turnos.xlsx)
                        aux_rut_turnos=np.where(rut_paciente_f==  rut_paciente[aux_isapre_infopacientes[l]])[0]
                        #Turnos que se hicieron para ese paciente durante ese mes (turnos.xlsx)
                        aux_mesrut_turnos=list(set(aux_mes_turnos) & set(aux_rut_turnos) & set(aux_agno_turnos))
                        #Entrar si hay algun turno hecho para ese paciente
                        
                        
                        if len(aux_mesrut_turnos)>0:
                            #Guardar informacion sobre el paciente (info_pacientes.xlsx)
                            ff3.write( '  .......................................................................................' )
                            ff3.write( '\n')
                            ff3.write( '  ' )
                            ff3.write('Paciente = ' + str(format(nombre_paciente[aux_isapre_infopacientes[l]])) +  ' - ' )
                            ff3.write('Rut = ' + str(format(rut_paciente[aux_isapre_infopacientes[l]])) +  '\n' )
                            ff3.write( '  ' )
                            ff3.write('Ciudad = ' + str(format(ciudad[aux_isapre_infopacientes[l]])) +  '\n' )
                            ff3.write( '\n')
         
                            ontime_pq_dia=np.where(dia_ingreso_datos[aux_mesrut_turnos]<20)[0]
                            ontime_pq_mes=np.where(mes_ingreso_datos[aux_mesrut_turnos]==mes[aux_mesrut_turnos])[0]
                            ontime_pq =np.where(dia[aux_mesrut_turnos]<16)[0]
                            aux_pq=list(set(ontime_pq) & set(ontime_pq_dia) & set(ontime_pq_mes))
                            resagados_pq=(list(set(ontime_pq).difference(aux_pq)))
                            
                            late_sq_dia=np.where(dia_ingreso_datos[aux_mesrut_turnos]>5)[0]
                            late_sq_mes=np.where((mes_ingreso_datos[aux_mesrut_turnos]-1)==mes[aux_mesrut_turnos])[0]
                            ontime_sq =np.where(dia[aux_mesrut_turnos]>15)[0]
                            resagados_sq=list(set(ontime_sq) & set(late_sq_dia) & set(late_sq_mes))
                            aux_sq=(list(set(ontime_sq).difference(resagados_sq)))

                            '''
                            ontime_sq=np.where(dia[aux_mesrut_turnos]>15)[0]

                            ontime_sq_dia=np.where(dia_ingreso_datos[aux_mesrut_turnos]>15)[0]
                            ontime_sq_mes=np.where(mes_ingreso_datos[aux_mesrut_turnos]==mes[aux_mesrut_turnos])[0]
                            ontime_sq_pp=list(set(ontime_sq) & set(ontime_sq_dia) & set(ontime_sq_mes))
                            
                            ontime_sq_dia_=np.where(dia_ingreso_datos[aux_mesrut_turnos]<6)[0]
                            ontime_sq_mes_=np.where(mes_ingreso_datos[aux_mesrut_turnos]-1==mes[aux_mesrut_turnos])[0]
                            ontime_sq_sp=list(set(ontime_sq) & set(ontime_sq_dia_) & set(ontime_sq_mes_))

                            aux_sq=list(set(ontime_sq_pp) | set(ontime_sq_sp))
                            resagados_sq=(list(set(ontime_sq).difference(aux_sq)))

                            '''
                            
                            profesionaltodel_pq= list(set(profesional[aux_mesrut_turnos][aux_pq]))
                            profesionaltodel_sq= list(set(profesional[aux_mesrut_turnos][aux_sq]))
                            
                            profesionaltodel_pq_RES= list(set(profesional[aux_mesrut_turnos][resagados_pq]))
                            profesionaltodel_sq_RES= list(set(profesional[aux_mesrut_turnos][resagados_sq]))
                            
                            codigo_bono_pq=[]
                            valor_bono_pq=[]
                            tipo_bono_pq=[]
                            codigo_bono_sq=[]
                            valor_bono_sq=[]
                            tipo_bono_sq=[]
                            
                            codigo_bono_pq_RES=[]
                            valor_bono_pq_RES=[]
                            tipo_bono_pq_RES=[]
                            codigo_bono_sq_RES=[]
                            valor_bono_sq_RES=[]
                            tipo_bono_sq_RES=[]
                                                      
                            codigo_bono_gesmed=[]
                            valor_bono_gesmed=[]
                            tipo_bono_gesmed=[]
                            
                            if len(aux_pq)>0:
                                pdf_name = isapre[aux_isapre_infopacientes[l]]+'_'+ mes_nombre[j] + '_PQ_' + nombre_paciente[aux_isapre_infopacientes[l]]+'.pdf'
                                save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], pdf_name)
                                
                                if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], 'GESMED')

                                canvas1 = Canvas(save_name, pagesize=LETTER)
                                canvas1.drawImage('vitalsur.png',220,650,170,100,anchor='sw',anchorAtXY=True,showBoundary=False)
                                xdata=50
                                deltax=250
                                ydata=550
                                deltay=20
                                canvas1.setFont('Helvetica-Bold', 16)
                                canvas1.drawString(150,625, 'PLANILLA DE ATENCIÓN GES PALIATIVO')
                                canvas1.setFont('Helvetica', 13)
                                canvas1.drawString(xdata,ydata, 'NOMBRE: '+ nombre_paciente[aux_isapre_infopacientes[l]])
                                canvas1.drawString(xdata+deltax,ydata, 'RUT: '+ rut_completo[aux_isapre_infopacientes[l]])
                                canvas1.drawString(xdata,ydata-deltay, 'MES: Primera Quincena '+ mes_nombre[j] )
                                canvas1.drawString(xdata+deltax,ydata-deltay, 'PREVISIÓN: '+ isapre[aux_isapre_infopacientes[l]])
                                canvas1.drawString(xdata,ydata-2*deltay, 'AMB: '+ ' ' )
                                canvas1.drawString(xdata+deltax,ydata-2*deltay, 'PAM: '+ ' ' )
                                canvas1.drawString(xdata,ydata-3*deltay, 'DIAG: '+ diagnostico[aux_isapre_infopacientes[l]])
                                canvas1.drawString(xdata,ydata-4*deltay, 'MÉDICO: '+ ' ')
                                canvas1.drawString(xdata-27,ydata-6*deltay, 'CÓDIGO')
                                canvas1.drawString(xdata+30,ydata-6*deltay, 'GLOSA')
                                #canvas1.drawString(xdata+233,ydata-6*deltay, 'N°')
                                canvas1.drawString(xdata+252,ydata-6*deltay, 'FECHA')
                                canvas1.drawString(xdata+317,ydata-6*deltay, 'PROFESIONAL')
                                canvas1.drawString(xdata+435,ydata-6*deltay, 'RUT')
                                medico_nombre=0
                                canvas1.setFont('Helvetica', 10)
                                
                            if len(resagados_pq)>0:
                                pdf_name = isapre[aux_isapre_infopacientes[l]]+'_'+ mes_nombre[j] + '_PQA_' + nombre_paciente[aux_isapre_infopacientes[l]]+'.pdf'
                                save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], pdf_name)
                                
                                if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], 'GESMED')

                                canvas1r = Canvas(save_name, pagesize=LETTER)
                                canvas1r.drawImage('vitalsur.png',220,650,170,100,anchor='sw',anchorAtXY=True,showBoundary=False)
                                xdata=50
                                deltax=250
                                ydata=550
                                deltay=20
                                canvas1r.setFont('Helvetica-Bold', 16)
                                canvas1r.drawString(150,625, 'PLANILLA DE ATENCIÓN GES PALIATIVO')
                                canvas1r.setFont('Helvetica', 13)
                                canvas1r.drawString(xdata,ydata, 'NOMBRE: '+ nombre_paciente[aux_isapre_infopacientes[l]])
                                canvas1r.drawString(xdata+deltax,ydata, 'RUT: '+ rut_completo[aux_isapre_infopacientes[l]])
                                canvas1r.drawString(xdata,ydata-deltay, 'MES: Primera Quincena '+ mes_nombre[j] )
                                canvas1r.drawString(xdata+deltax,ydata-deltay, 'PREVISIÓN: '+ isapre[aux_isapre_infopacientes[l]])
                                canvas1r.drawString(xdata,ydata-2*deltay, 'AMB: '+ ' ' )
                                canvas1r.drawString(xdata+deltax,ydata-2*deltay, 'PAM: '+ ' ' )
                                canvas1r.drawString(xdata,ydata-3*deltay, 'DIAG: '+ diagnostico[aux_isapre_infopacientes[l]])
                                canvas1r.drawString(xdata,ydata-4*deltay, 'MÉDICO: '+ ' ')
                                canvas1r.drawString(xdata-27,ydata-6*deltay, 'CÓDIGO')
                                canvas1r.drawString(xdata+30,ydata-6*deltay, 'GLOSA')
                                #canvas1r.drawString(xdata+233,ydata-6*deltay, 'N°')
                                canvas1r.drawString(xdata+252,ydata-6*deltay, 'FECHA')
                                canvas1r.drawString(xdata+317,ydata-6*deltay, 'PROFESIONAL')
                                canvas1r.drawString(xdata+435,ydata-6*deltay, 'RUT')
                                medico_nombre_r=0
                                canvas1r.setFont('Helvetica', 10)
                                
                            if len(aux_sq)>0:
                                pdf_name = isapre[aux_isapre_infopacientes[l]]+'_'+ mes_nombre[j] + '_SQ_' + nombre_paciente[aux_isapre_infopacientes[l]]+'.pdf'
                                save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], pdf_name)
                                
                                if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], 'GESMED')
                                    
                                canvas2 = Canvas(save_name, pagesize=LETTER)
                                canvas2.setFont('Helvetica', 12)
                                canvas2.drawImage('vitalsur.png',220,650,170,100,anchor='sw',anchorAtXY=True,showBoundary=False)
                                xdata=50
                                deltax=250
                                ydata=550
                                deltay=20
                                canvas2.setFont('Helvetica-Bold', 16)
                                canvas2.drawString(150,625, 'PLANILLA DE ATENCIÓN GES PALIATIVO')
                                canvas2.setFont('Helvetica', 13)
                                canvas2.drawString(xdata,ydata, 'NOMBRE: '+ nombre_paciente[aux_isapre_infopacientes[l]])
                                canvas2.drawString(xdata+deltax,ydata, 'RUT: '+ rut_completo[aux_isapre_infopacientes[l]])
                                canvas2.drawString(xdata,ydata-deltay, 'MES: Segunda Quincena '+ mes_nombre[j] )
                                canvas2.drawString(xdata+deltax,ydata-deltay, 'PREVISIÓN: '+ isapre[aux_isapre_infopacientes[l]])
                                canvas2.drawString(xdata,ydata-2*deltay, 'AMB: '+ ' ' )
                                canvas2.drawString(xdata+deltax,ydata-2*deltay, 'PAM: '+ ' ' )
                                canvas2.drawString(xdata,ydata-3*deltay, 'DIAG: '+ diagnostico[aux_isapre_infopacientes[l]])
                                canvas2.drawString(xdata,ydata-4*deltay, 'MÉDICO: '+ ' ')
                                canvas2.drawString(xdata-27,ydata-6*deltay, 'CÓDIGO')
                                canvas2.drawString(xdata+30,ydata-6*deltay, 'GLOSA')
                                #canvas2.drawString(xdata+233,ydata-6*deltay, 'N°')
                                canvas2.drawString(xdata+252,ydata-6*deltay, 'FECHA')
                                canvas2.drawString(xdata+317,ydata-6*deltay, 'PROFESIONAL')
                                canvas2.drawString(xdata+435,ydata-6*deltay, 'RUT')
                                medico_nombreb=0
                                canvas2.setFont('Helvetica', 10)
                                
                            if len(resagados_sq)>0:
                                pdf_name = isapre[aux_isapre_infopacientes[l]]+'_'+ mes_nombre[j] + '_SQA_' + nombre_paciente[aux_isapre_infopacientes[l]]+'.pdf'
                                save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], pdf_name)
                                
                                if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], 'GESMED')
                                    
                                canvas2r = Canvas(save_name, pagesize=LETTER)
                                canvas2r.setFont('Helvetica', 12)
                                canvas2r.drawImage('vitalsur.png',220,650,170,100,anchor='sw',anchorAtXY=True,showBoundary=False)
                                xdata=50
                                deltax=250
                                ydata=550
                                deltay=20
                                canvas2r.setFont('Helvetica-Bold', 16)
                                canvas2r.drawString(150,625, 'PLANILLA DE ATENCIÓN GES PALIATIVO')
                                canvas2r.setFont('Helvetica', 13)
                                canvas2r.drawString(xdata,ydata, 'NOMBRE: '+ nombre_paciente[aux_isapre_infopacientes[l]])
                                canvas2r.drawString(xdata+deltax,ydata, 'RUT: '+ rut_completo[aux_isapre_infopacientes[l]])
                                canvas2r.drawString(xdata,ydata-deltay, 'MES: Segunda Quincena '+ mes_nombre[j] )
                                canvas2r.drawString(xdata+deltax,ydata-deltay, 'PREVISIÓN: '+ isapre[aux_isapre_infopacientes[l]])
                                canvas2r.drawString(xdata,ydata-2*deltay, 'AMB: '+ ' ' )
                                canvas2r.drawString(xdata+deltax,ydata-2*deltay, 'PAM: '+ ' ' )
                                canvas2r.drawString(xdata,ydata-3*deltay, 'DIAG: '+ diagnostico[aux_isapre_infopacientes[l]])
                                canvas2r.drawString(xdata,ydata-4*deltay, 'MÉDICO: '+ ' ')
                                canvas2r.drawString(xdata-27,ydata-6*deltay, 'CÓDIGO')
                                canvas2r.drawString(xdata+30,ydata-6*deltay, 'GLOSA')
                                #canvas2r.drawString(xdata+233,ydata-6*deltay, 'N°')
                                canvas2r.drawString(xdata+252,ydata-6*deltay, 'FECHA')
                                canvas2r.drawString(xdata+317,ydata-6*deltay, 'PROFESIONAL')
                                canvas2r.drawString(xdata+435,ydata-6*deltay, 'RUT')
                                medico_nombreb_r=0
                                canvas2r.setFont('Helvetica', 10)
                                
                            linestablea=20
                            linestableb=20
                            visita_a=0
                            visita_b=0
                            
                            linestablea_r=20
                            linestableb_r=20
                            visita_a_r=0
                            visita_b_r=0
                                
                            #Loop para la cantidad de turnos hechos para ese paciente durante ese mes
                            for q in range(len(aux_mesrut_turnos)):
                                #Columna de pagos y cobros para el profesional de ese turno
                                valores_profesional=np.asarray(dfs3[profesional[aux_mesrut_turnos[q]]].tolist())
                        
                                aux_isapre_infoisapres=np.where(isapre_unique[i]==i_isapre)[0]
                                aux_ciudad_dondevive=np.where(ciudad[aux_isapre_infopacientes[l]]==i_ciudad)[0]
                                aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)) #info isapres
                                if len(aux_isapreyciudad)==0:
                                   aux_ciudad_dondevive=np.where('Otros'==i_ciudad)[0]
                                aux_sector_dondevive=np.where(sector[aux_isapre_infopacientes[l]]==i_sector)[0]
                                aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)  & set(aux_sector_dondevive)) #info isapres

                                aux_cobro_infoisapres=np.where('Cobro'==i_cobropago)[0]
                                aux_pago_infoisapres=np.where('Pago'==i_cobropago)[0]
                                  
                                #Posicion de la fila correspondiente al caso del paciente (info_isapres.xlsx)
                                aux_poscobro_infoisapres=list(set(aux_isapreyciudad) & set(aux_cobro_infoisapres))
                                aux_pospago_infoisapres=list(set(aux_isapreyciudad) & set(aux_pago_infoisapres))
                                  
                                if len(aux_isapre_infoisapres)==0:
                                   ffwarning.write('Hay una isapre que no esta especificada en info_isapres.xlsx')
                                   ffwarning.write(' Check: '+ isapre[aux_rut_infopac] )
                                   ffwarning.write( '\n')
                                   print('WARNING7!')

                                if len(aux_poscobro_infoisapres)>1:
                                   ffwarning.write('Hay dos valores para el cobro del turno en info_isapres.xlsx')
                                   ffwarning.write(' Check: '+ i_isapre[aux_isapre_infoisapres[0]] + ' ' + i_ciudad[aux_ciudad_dondevive[0]] +  ' COBRO')
                                   ffwarning.write( '\n')

                                if len(aux_pospago_infoisapres)>1:
                                   ffwarning.write('Hay dos valores para el pago del turno en info_isapres.xlsx')
                                   ffwarning.write(' Check: ' +  i_isapre[aux_isapre_infoisapres[0]] + ' '+ i_ciudad[aux_ciudad_dondevive[0]] +  ' PAGO')
                                   ffwarning.write( '\n')

                                if len(aux_pospago_infoisapres)==0 or len(aux_poscobro_infoisapres)==0:
                                   ffwarning.write('Hay un error en leer el cobro/pago de la isapre - info_isapres.xlsx')
                                   ffwarning.write(' Check: '+ isapre[aux_rut_infopac] )
                                   ffwarning.write( '\n')
                                   print('WARNING7!')

                                cobrar_mes_paciente=cobrar_mes_paciente+(valores_profesional[aux_poscobro_infoisapres])
                                cobrar_mes=cobrar_mes+(valores_profesional[aux_poscobro_infoisapres])
                                 
                                if profesional[aux_mesrut_turnos[q]] == 'Medico':
                                    pMedico[j]=pMedico[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Enfermera':
                                    pEnfermera[j]=pEnfermera[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Nutricionista':
                                    pNutricionista[j]=pNutricionista[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Kinesiologo':
                                    pKinesiologo[j]=pKinesiologo[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Psicologo':
                                    pPsicologo[j]=pPsicologo[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Fonoaudiologo':
                                    pFono[j]=pFono[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Terapeuta Ocupacional':
                                    pTO[j]=pTO[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Tens':
                                    pTens[j]=pTens[j]+1
                                    definicion= 'Tens'
                                elif profesional[aux_mesrut_turnos[q]] == 'Curacion Simple':
                                    pCuracion_Simple[j]=pCuracion_Simple[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Educacion Enfermera':
                                    pEducacion_Enfermera[j]=pEducacion_Enfermera[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Educacion Tens':
                                    pEducacion_Tens[j]=pEducacion_Tens[j]+1
                                elif profesional[aux_mesrut_turnos[q]] == 'Intervension Psicosocial':
                                    pIntervension_Psicosocial[j]=pIntervension_Psicosocial[j]+1

                                ff3.write( '            ' )
                                ff3.write('Visita' + str(format(q+1)) +  '\n' )
                                ff3.write( '            ' )
                                ff3.write('Tipo Visita = ' + str(format(profesional[aux_mesrut_turnos[q]])) +  '\n' )
                                ff3.write( '            ' )
                                ff3.write('Profesional = ' + str(format(rut_profesional[aux_mesrut_turnos[q]])) +  '\n' )
                                ff3.write( '            ' )
                                ff3.write('Fecha = ' + str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])) +  '\n' )
                                ff3.write( '            ' )
                                ff3.write('Cobro = ' + str(format((valores_profesional[aux_poscobro_infoisapres[0]]))) +  '\n' )
                                ff3.write( '\n')

                                data_prof=np.where(rut_profesional[aux_mesrut_turnos[q]]==i_rut_profesional)[0]
                                bono_isapre=np.asarray(dfs5['Isapre'].tolist())
                                bono_cur=np.asarray(dfs5[profesional[aux_mesrut_turnos[q]]].tolist())
                                data_bono=np.where(isapre[aux_isapre_infopacientes[l]]==bono_isapre)[0]

                                if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    codigo_bono_gesmed.append(bono_cur[data_bono[0]])
                                    valor_bono_gesmed.append(valores_profesional[aux_poscobro_infoisapres])
                                    tipo_bono_gesmed.append(profesional[aux_mesrut_turnos[q]])
                                    
                                if dia[aux_mesrut_turnos[q]]<16 and dia_ingreso_datos[aux_mesrut_turnos[q]]<20 and mes_ingreso_datos[aux_mesrut_turnos[q]]==mes[aux_mesrut_turnos[q]]:
                                
                                    
                                    if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                        codigo_bono_pq.append(bono_cur[data_bono[0]])
                                        valor_bono_pq.append(valores_profesional[aux_poscobro_infoisapres])
                                        tipo_bono_pq.append(bono_cur[0])
                                    
                                    canvas1.setFont('Helvetica', 10)
                                    canvas1.drawString(xdata-27,ydata-6*deltay-linestablea, bono_cur[data_bono[0]])
                                    canvas1.drawString(xdata+30,ydata-6*deltay-linestablea, str(bono_cur[0]))
                                    canvas1.drawString(xdata+252,ydata-6*deltay-linestablea,  str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])))
                                    canvas1.drawString(xdata+317,ydata-6*deltay-linestablea, i_nombre_profesional[data_prof][0])
                                    canvas1.drawString(xdata+435,ydata-6*deltay-linestablea, i_rutcompleto[data_prof][0])
                                    thing_list=list(profesionaltodel_pq)
                                    elem=profesional[aux_mesrut_turnos[q]]
                                    aux_del = thing_list.index(elem) if elem in thing_list else -1
                                    cantidad=np.where(profesional[aux_mesrut_turnos[q]]==profesional[aux_mesrut_turnos][aux_pq])[0]
                                    if aux_del>-1:
                                        profesionaltodel_pq = np.delete(profesionaltodel_pq, aux_del)
                                        #canvas1.drawString(xdata+233,ydata-6*deltay-linestablea, '%i' %len(cantidad))
                                    visita_a=visita_a+1
                                    if visita_a>17:
                                        canvas1.showPage()
                                        visita_a=0
                                        linestablea=-260
                                    linestablea=linestablea+20
                                    
                                elif dia[aux_mesrut_turnos[q]]<16 and dia_ingreso_datos[aux_mesrut_turnos[q]]>19:

                                    if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                        codigo_bono_pq_RES.append(bono_cur[data_bono[0]])
                                        valor_bono_pq_RES.append(valores_profesional[aux_poscobro_infoisapres])
                                        tipo_bono_pq_RES.append(bono_cur[0])
                                    
                                    canvas1r.setFont('Helvetica', 10)
                                    canvas1r.drawString(xdata-27,ydata-6*deltay-linestablea_r, bono_cur[data_bono[0]])
                                    canvas1r.drawString(xdata+30,ydata-6*deltay-linestablea_r, str(bono_cur[0]))
                                    canvas1r.drawString(xdata+252,ydata-6*deltay-linestablea_r,  str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])))
                                    canvas1r.drawString(xdata+317,ydata-6*deltay-linestablea_r, i_nombre_profesional[data_prof][0])
                                    canvas1r.drawString(xdata+435,ydata-6*deltay-linestablea_r, i_rutcompleto[data_prof][0])
                                    thing_list=list(profesionaltodel_pq_RES)
                                    elem=profesional[aux_mesrut_turnos[q]]
                                    aux_del = thing_list.index(elem) if elem in thing_list else -1
                                    cantidad=np.where(profesional[aux_mesrut_turnos[q]]==profesional[aux_mesrut_turnos][resagados_pq])[0]
                                    if aux_del>-1:
                                        profesionaltodel_pq_RES = np.delete(profesionaltodel_pq_RES, aux_del)
                                        #canvas1r.drawString(xdata+233,ydata-6*deltay-linestablea_r, '%i' %len(cantidad))
                                    visita_a_r=visita_a_r+1
                                    if visita_a_r>17:
                                        canvas1r.showPage()
                                        visita_a_r=0
                                        linestablea_r=-260
                                    linestablea_r=linestablea_r+20
                                
                                    
                                elif (dia[aux_mesrut_turnos[q]]>15 and dia_ingreso_datos[aux_mesrut_turnos[q]]>15 and (mes_ingreso_datos[aux_mesrut_turnos[q]]==mes[aux_mesrut_turnos[q]]))  or (dia[aux_mesrut_turnos[q]]>15 and dia_ingreso_datos[aux_mesrut_turnos[q]]<6 and (mes_ingreso_datos[aux_mesrut_turnos[q]]-1)==mes[aux_mesrut_turnos[q]]):
                                
                                    if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                        codigo_bono_sq.append(bono_cur[data_bono[0]])
                                        valor_bono_sq.append(valores_profesional[aux_poscobro_infoisapres])
                                        tipo_bono_sq.append(bono_cur[0])
                                    
                                    canvas2.setFont('Helvetica', 10)
                                    canvas2.drawString(xdata-27,ydata-6*deltay-linestableb, bono_cur[data_bono[0]])
                                    canvas2.drawString(xdata+30,ydata-6*deltay-linestableb, str(bono_cur[0]))
                                    canvas2.drawString(xdata+252,ydata-6*deltay-linestableb,  str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])))
                                    canvas2.drawString(xdata+317,ydata-6*deltay-linestableb, i_nombre_profesional[data_prof][0])
                                    canvas2.drawString(xdata+435,ydata-6*deltay-linestableb, i_rutcompleto[data_prof][0])
                                          
                                    thing_list=list(profesionaltodel_sq)
                                    elem=profesional[aux_mesrut_turnos[q]]
                                    aux_del = thing_list.index(elem) if elem in thing_list else -1
                                    cantidad=np.where(profesional[aux_mesrut_turnos[q]]==profesional[aux_mesrut_turnos][aux_sq])[0]
                                    
                                    if aux_del>-1:
                                        profesionaltodel_sq = np.delete(profesionaltodel_sq, aux_del)
                                        #canvas2.drawString(xdata+233,ydata-6*deltay-linestableb, '%i' %len(cantidad))
                                    visita_b=visita_b+1
                                    if visita_b>17:
                                        canvas2.showPage()
                                        visita_b=0
                                        linestableb=-260
                                    linestableb=linestableb+20
                                
                                elif dia[aux_mesrut_turnos[q]]>15 and ((mes_ingreso_datos[aux_mesrut_turnos[q]]-1)==mes[aux_mesrut_turnos[q]]):

                                    if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                        codigo_bono_sq_RES.append(bono_cur[data_bono[0]])
                                        valor_bono_sq_RES.append(valores_profesional[aux_poscobro_infoisapres])
                                        tipo_bono_sq_RES.append(bono_cur[0])

                                    canvas2r.setFont('Helvetica', 10)
                                    canvas2r.drawString(xdata-27,ydata-6*deltay-linestableb_r, bono_cur[data_bono[0]])
                                    canvas2r.drawString(xdata+30,ydata-6*deltay-linestableb_r, str(bono_cur[0]))
                                    canvas2r.drawString(xdata+252,ydata-6*deltay-linestableb_r,  str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])))
                                    canvas2r.drawString(xdata+317,ydata-6*deltay-linestableb_r, i_nombre_profesional[data_prof][0])
                                    canvas2r.drawString(xdata+435,ydata-6*deltay-linestableb_r, i_rutcompleto[data_prof][0])
                                    thing_list=list(profesionaltodel_sq_RES)
                                    elem=profesional[aux_mesrut_turnos[q]]
                                    aux_del = thing_list.index(elem) if elem in thing_list else -1
                                    cantidad=np.where(profesional[aux_mesrut_turnos[q]]==profesional[aux_mesrut_turnos][resagados_sq])[0]
                                    if aux_del>-1:
                                        profesionaltodel_sq_RES = np.delete(profesionaltodel_sq_RES, aux_del)
                                        #canvas2r.drawString(xdata+233,ydata-6*deltay-linestableb_r, '%i' %len(cantidad))
                                    visita_b_r=visita_b_r+1
                                    if visita_b_r>17:
                                        canvas2r.showPage()
                                        visita_b_r=0
                                        linestableb_r=-260
                                    linestableb_r=linestableb_r+20
                                   
                                   
                                   
                            if len(aux_pq)>0:
                               canvas1.drawString(xdata+190,ydata-6*deltay-linestablea-30, 'Total = %i' %len(aux_pq))
                               canvas1.save()
                               if (sector[aux_isapre_infopacientes[l]]>0):
                                   ff4mail.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) + ' - Sector ' +str(sector[aux_isapre_infopacientes[l]]) +') ')
                               else:
                                   ff4mail.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) +') ')
                               ff4mail.write('\n')
                               
                               if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                    ff4valores.write('\n')
                                    ff4valores.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(rut_completo[aux_isapre_infopacientes[l]]) +') \n')
                                    bonos_unique_pq = list(set(codigo_bono_pq))
                                    for iii in range(len(bonos_unique_pq)):
                                        a=np.where(bonos_unique_pq[iii]==np.asarray(codigo_bono_pq))[0]
                                        ff4valores.write( str(bonos_unique_pq[iii]) + '  ')
                                        ff4valores.write( str(tipo_bono_pq[a[0]]) + '  ')
                                        ff4valores.write( str(len(a)) + '  ')
                                        ff4valores.write( str( int(len(a)* valor_bono_pq[a[0]] ) ) + '  \n')
                                        
                            if len(resagados_pq)>0:
                               canvas1r.drawString(xdata+190,ydata-6*deltay-linestablea_r-30, 'Total = %i' %len(resagados_pq))
                               canvas1r.save()
                               if (sector[aux_isapre_infopacientes[l]]>0):
                                   ff4mail_RES.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) + ' - Sector ' +str(sector[aux_isapre_infopacientes[l]]) +') ')
                               else:
                                   ff4mail_RES.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) +') ')
                               ff4mail_RES.write('\n')
                               
                               if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                    ff4valores_RES.write('\n')
                                    ff4valores_RES.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(rut_completo[aux_isapre_infopacientes[l]]) +') \n')
                                    bonos_unique_pq_RES = list(set(codigo_bono_pq_RES))
                                    for iii in range(len(bonos_unique_pq_RES)):
                                        a=np.where(bonos_unique_pq_RES[iii]==np.asarray(codigo_bono_pq_RES))[0]
                                        ff4valores_RES.write( str(bonos_unique_pq_RES[iii]) + '  ')
                                        ff4valores_RES.write( str(tipo_bono_pq_RES[a[0]]) + '  ')
                                        ff4valores_RES.write( str(len(a)) + '  ')
                                        ff4valores_RES.write( str( int(len(a)* valor_bono_pq_RES[a[0]] ) ) + '  \n')
                                    
                            if len(aux_sq)>0:
                                canvas2.drawString(xdata+190,ydata-6*deltay-linestableb-30, 'Total = %i' %len(aux_sq))
                                canvas2.save()
                                if (sector[aux_isapre_infopacientes[l]]>0):
                                    ff4mailsq.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) + ' - Sector ' +str(sector[aux_isapre_infopacientes[l]]) +') ')
                                else:
                                    ff4mailsq.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) +') ')
                                ff4mailsq.write('\n')
                                
                                if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                    ff4valoressq.write('\n')
                                    ff4valoressq.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(rut_completo[aux_isapre_infopacientes[l]]) +') \n ')
                               
                                    bonos_unique_sq = list(set(codigo_bono_sq))
                                    for iii in range(len(bonos_unique_sq)):
                                        a=np.where(bonos_unique_sq[iii]==np.asarray(codigo_bono_sq))[0]
                                        ff4valoressq.write( str(bonos_unique_sq[iii]) + '  ')
                                        ff4valoressq.write( str(tipo_bono_sq[a[0]]) + '  ')
                                        ff4valoressq.write( str(len(a)) + '  ')
                                        ff4valoressq.write( str( int(len(a)* valor_bono_sq[a[0]] ) ) + ' \n')
            
                            if len(resagados_sq)>0:
                               canvas2r.drawString(xdata+190,ydata-6*deltay-linestableb_r-30, 'Total = %i' %len(resagados_sq))
                               canvas2r.save()
                               if (sector[aux_isapre_infopacientes[l]]>0):
                                   ff4mailsq_RES.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) + ' - Sector ' +str(sector[aux_isapre_infopacientes[l]]) +') ')
                               else:
                                   ff4mailsq_RES.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) +') ')
                               ff4mailsq_RES.write('\n')
                               
                               if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                    ff4valoressq_RES.write('\n')
                                    ff4valoressq_RES.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(rut_completo[aux_isapre_infopacientes[l]]) +') \n')
                                    bonos_unique_sq_RES = list(set(codigo_bono_sq_RES))
                                    for iii in range(len(bonos_unique_sq_RES)):
                                        a=np.where(bonos_unique_sq_RES[iii]==np.asarray(codigo_bono_sq_RES))[0]
                                        ff4valoressq_RES.write( str(bonos_unique_sq_RES[iii]) + '  ')
                                        ff4valoressq_RES.write( str(tipo_bono_sq_RES[a[0]]) + '  ')
                                        ff4valoressq_RES.write( str(len(a)) + '  ')
                                        ff4valoressq_RES.write( str( int(len(a)* valor_bono_sq_RES[a[0]] ) ) + '  \n')
            
            
                            if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    bonos_unique_gesmed = list(set(codigo_bono_gesmed))
                                    for iii in range(len(bonos_unique_gesmed)):
                                        a=np.where(bonos_unique_gesmed[iii]==np.asarray(codigo_bono_gesmed))[0]
                                        
                                        worksheet.write(row, 0,     str(j+1) + '/' + str(agno[gn]))
                                        worksheet.write(row, 1,     str(nombre_paciente[aux_isapre_infopacientes[l]]))
                                        worksheet.write(row, 2,     str(rut_completo[aux_isapre_infopacientes[l]]))
                                        worksheet.write(row, 3,     str(format(ciudad[aux_isapre_infopacientes[l]])))
                                        worksheet.write(row, 4,     str(tipo_bono_gesmed[a[0]]) )
                                        worksheet.write(row, 5,     str(len(a)))
                                        worksheet.write(row, 6,     str( int(valor_bono_gesmed[a[0]] ) ) )
                                        worksheet.write(row, 7,     str( int(len(a)* valor_bono_gesmed[a[0]] ) ) )
                                        
                                        if (sector[aux_isapre_infopacientes[l]]>0):
                                            worksheet.write(row, 8,    str(sector[aux_isapre_infopacientes[l]]))
                                            
                                        else:
                                            worksheet.write(row, 8,    ' REGION')
                                        row += 1


                            ff3.write( '\n')
                            ff3.write( '            ' )
                            ff3.write('Cobrar a Isapre por paciente =' + str(format(int(cobrar_mes_paciente))) +  ' ' )
                            ff3.write( '\n')
                            ff3.write( '\n')
                            
                    ff4mail.write('Esperando una buena acogida y su pronta respuesta. \n Saluda cordialmente a usted. \n \n \n')
                    ff4mailsq.write('Esperando una buena acogida y su pronta respuesta. \n Saluda cordialmente a usted. \n \n \n')
                    
                    ff4mail_RES.write('Esperando una buena acogida y su pronta respuesta. \n Saluda cordialmente a usted. \n \n \n')
                    ff4mailsq_RES.write('Esperando una buena acogida y su pronta respuesta. \n Saluda cordialmente a usted. \n \n \n')
                    
                    ff3.write( '\n')
                    ff3.write( '\n')
                    ff3.write( '                      .......................................................................................' )
                    ff3.write( '\n')
                    ff3.write( '\n')
                    ff3.write( '                      ' )
                    ff3.write('Cobrar a Isapre mes de ' + str(mes_nombre[j]) + ' = ' + str(format(int(cobrar_mes))) +  ' ' )
                    ff3.write( '\n')
                    ff3.write( '\n')
                    ff3.write( '\n')

                out_tpl = np.nonzero(pmes)
                            
                fig = plt.figure()
                ax = plt.subplot(111)
                ax.plot(pmes[out_tpl], pMedico[out_tpl], '.-', label='Medico', color='red')
                ax.plot(pmes[out_tpl], pEnfermera[out_tpl], '.-', label='Enfermera', color='blue')
                ax.plot(pmes[out_tpl], pNutricionista[out_tpl], '.-', label='Nutricionista', color='green')
                ax.plot(pmes[out_tpl], pKinesiologo[out_tpl], '.-', label='Kinesiologo', color='magenta')
                ax.plot(pmes[out_tpl], pPsicologo[out_tpl], '.-', label='Psicologo', color='purple')
                ax.plot(pmes[out_tpl], pTens[out_tpl], '.-', label='Tens', color='green')
                ax.plot(pmes[out_tpl], pCuracion_Simple[out_tpl], '.-', label='C Simple', color='cyan')
                ax.plot(pmes[out_tpl], pEducacion_Enfermera[out_tpl], '.-', label='Ed Enfermera', color='black')
                ax.plot(pmes[out_tpl], pEducacion_Tens[out_tpl], '.-', label='Ed Tens', color='yellow')
                ax.plot(pmes[out_tpl], pIntervension_Psicosocial[out_tpl], '.-', label='In Psicosocial', color='gold')
                ax.set_xlabel('Mes')
                ax.set_ylabel('Numero de Visitas')
                ax.set_title('%s' %isapre_unique[i], fontsize=20)
                ax.set_xlim(0.5, max(pmes)+0.5)
                ax.set_xticks(np.arange(1, max(pmes)+0.5, 1))
                box = ax.get_position()
                ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
                ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
                #save_name = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Vitalsur_Isapre_%s_%i.png' %(isapre_unique[i], agno[gn]))
                save_name = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Vitalsur_Isapre_%s_%i.png' %(isapre_unique[i], agno[gn]))
                plt.savefig(save_name)
                
                
                

            else:
               aux_agno_turnos=np.where(ano==agno[gn])[0]
               pmes=np.zeros(12)
               pMedico=np.zeros(12)
               pEnfermera=np.zeros(12)
               pNutricionista=np.zeros(12)
               pKinesiologo=np.zeros(12)
               pPsicologo=np.zeros(12)
               pFono=np.zeros(12)
               pTO=np.zeros(12)
               pTens=np.zeros(12)
               pCuracion_Simple=np.zeros(12)
               pEducacion_Enfermera=np.zeros(12)
               pEducacion_Tens=np.zeros(12)
               pIntervension_Psicosocial=np.zeros(12)
               
               #save_name = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Vitalsur_Isapre_%s_%i.txt' %(isapre_unique[i],agno[gn]))
               save_name = os.path.join(os.path.curdir, "%i/"%agno[gn], 'Vitalsur_Isapre_%s_%i.txt' %(isapre_unique[i],agno[gn]))

               ff3 = open(save_name, 'w')

               
               ff4mail.write('_________________________________________________________________________________________________ \n')
               ff4mail.write('_________________________________________________________________________________________________ \n')
               ff4mail.write('Mail isapre '  + str(isapre_unique[i]) + '\n')
               ff4mail.write('_________________________________________________________________________________________________ \n')
               ff4mail.write('_________________________________________________________________________________________________ \n \n')
               ff4mailsq.write('_________________________________________________________________________________________________ \n')
               ff4mailsq.write('_________________________________________________________________________________________________ \n')
               ff4mailsq.write('Mail isapre '  + str(isapre_unique[i]) + '\n')
               ff4mailsq.write('_________________________________________________________________________________________________ \n')
               ff4mailsq.write('_________________________________________________________________________________________________ \n \n')
               
               
               ff4valores.write('_________________________________________________________________________________________________ \n')
               ff4valores.write('_________________________________________________________________________________________________ \n')
               ff4valores.write('Valores isapre '  + str(isapre_unique[i]) + '\n')
               ff4valores.write('_________________________________________________________________________________________________ \n')
               ff4valores.write('_________________________________________________________________________________________________ \n \n')
               ff4valoressq.write('_________________________________________________________________________________________________ \n')
               ff4valoressq.write('_________________________________________________________________________________________________ \n')
               ff4valoressq.write('Valores isapre '  + str(isapre_unique[i]) + '\n')
               ff4valoressq.write('_________________________________________________________________________________________________ \n')
               ff4valoressq.write('_________________________________________________________________________________________________ \n \n')
               
               
               #Posicion para pacientes con esa isapre (info_pacientes.xlsx)
               aux_isapre_infopacientes=np.where(isapre_unique[i]==isapre)[0]

               #Guardar informacion sobre isapre
               ff3.write( 'Isapre: ' + str(isapre_unique[i]) + ' ' )
               ff3.write( '\n')

               #Loop para 12  meses
               for j in range(len(mes_nombre)):
                   cobrar_mes=0
                   
                   pmes[j]=j+1
                   aux_agno_turnos=np.where(ano==agno[gn])[0]
                   #Turnos que se hicieron ese mes
                   aux_mes_turnos=np.where(mes==(j+1))[0]
                   
                   #Guardar nombre del mes
                   ff3.write( '_________________________________________________________________________________________________' )
                   ff3.write( '\n')
                   ff3.write( '_________________________________________________________________________________________________' )
                   ff3.write( '\n')
                   ff3.write( '' )
                   ff3.write('Visitas mes de '  + str(mes_nombre[j]) )
                   ff3.write( '\n')
                   ff3.write( '\n')
                   

                   ff4mail.write( '________Primera quincena mes de '  + str(mes_nombre[j]) + '________')
                   ff4mail.write( '\n')
                   ff4mail.write( '\n')
                   ff4mail.write( '\n')
                   ff4mail.write('Estimada/o, \n Junto con saludarla/o, y esperando se encuentre muy bien, me dirijo a usted para solicitar ayuda en el cobro de los bonos de la primera quincena del mes de ' )
                   ff4mail.write(str(mes_nombre[j]) )
                   ff4mail.write(' por la atención de los siguientes pacientes (adjunto planillas): \n')
                   
                   
                   ff4mailsq.write( '________Segunda quincena mes de '  + str(mes_nombre[j]) + '________')
                   ff4mailsq.write( '\n')
                   ff4mailsq.write( '\n')
                   ff4mailsq.write( '\n')
                   ff4mailsq.write('Estimada/o, \n Junto con saludarla/o, y esperando se encuentre muy bien, me dirijo a usted para solicitar ayuda en el cobro de los bonos de la segunda quincena del mes de ' )
                   ff4mailsq.write(str(mes_nombre[j]) )
                   ff4mailsq.write(' por la atención de los siguientes pacientes (adjunto planillas): \n')
                   
                   #Loop para cada paciente que se atiende con esa isapre (info_pacientes.xlsx)
                   for l in range(len(aux_isapre_infopacientes)):
                       cobrar_mes_paciente=0
                       #Turnos que se hicieron para ese paciente (turnos.xlsx)
                       aux_rut_turnos=np.where(rut_paciente_f==  rut_paciente[aux_isapre_infopacientes[l]])[0]
                       #Turnos que se hicieron para ese paciente durante ese mes (turnos.xlsx)
                       aux_mesrut_turnos=list(set(aux_mes_turnos) & set(aux_rut_turnos) & set(aux_agno_turnos))
                       #Entrar si hay algun turno hecho para ese paciente
                       
                       
                       if len(aux_mesrut_turnos)>0:
                           #Guardar informacion sobre el paciente (info_pacientes.xlsx)
                           ff3.write( '  .......................................................................................' )
                           ff3.write( '\n')
                           ff3.write( '  ' )
                           ff3.write('Paciente = ' + str(format(nombre_paciente[aux_isapre_infopacientes[l]])) +  ' - ' )
                           ff3.write('Rut = ' + str(format(rut_paciente[aux_isapre_infopacientes[l]])) +  '\n' )
                           ff3.write( '  ' )
                           ff3.write('Ciudad = ' + str(format(ciudad[aux_isapre_infopacientes[l]])) +  '\n' )
                           ff3.write( '\n')
        
                           aux_pq=np.where(dia[aux_mesrut_turnos]<16)[0]
                           aux_sq=np.where(dia[aux_mesrut_turnos]>15)[0]
                           profesionaltodel_pq= list(set(profesional[aux_mesrut_turnos][aux_pq]))
                           profesionaltodel_sq= list(set(profesional[aux_mesrut_turnos][aux_sq]))
                           
                           codigo_bono_pq=[]
                           valor_bono_pq=[]
                           tipo_bono_pq=[]
                           codigo_bono_sq=[]
                           valor_bono_sq=[]
                           tipo_bono_sq=[]
                            
                           codigo_bono_gesmed=[]
                           valor_bono_gesmed=[]
                           tipo_bono_gesmed=[]
                           
                           if len(aux_pq)>0:
                               pdf_name = isapre[aux_isapre_infopacientes[l]]+'_'+ mes_nombre[j] + '_PQ_' + nombre_paciente[aux_isapre_infopacientes[l]]+'.pdf'
                               save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], pdf_name)
                               
                               if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], 'GESMED')
                                    
                               canvas1 = Canvas(save_name, pagesize=LETTER)
                               canvas1.drawImage('vitalsur.png',220,650,170,100,anchor='sw',anchorAtXY=True,showBoundary=False)
                               xdata=50
                               deltax=250
                               ydata=550
                               deltay=20
                               canvas1.setFont('Helvetica-Bold', 16)
                               canvas1.drawString(150,625, 'PLANILLA DE ATENCIÓN GES PALIATIVO')
                               canvas1.setFont('Helvetica', 13)
                               canvas1.drawString(xdata,ydata, 'NOMBRE: '+ nombre_paciente[aux_isapre_infopacientes[l]])
                               canvas1.drawString(xdata+deltax,ydata, 'RUT: '+ rut_completo[aux_isapre_infopacientes[l]])
                               canvas1.drawString(xdata,ydata-deltay, 'MES: Primera Quincena '+ mes_nombre[j] )
                               canvas1.drawString(xdata+deltax,ydata-deltay, 'PREVISIÓN: '+ isapre[aux_isapre_infopacientes[l]])
                               canvas1.drawString(xdata,ydata-2*deltay, 'AMB: '+ ' ' )
                               canvas1.drawString(xdata+deltax,ydata-2*deltay, 'PAM: '+ ' ' )
                               canvas1.drawString(xdata,ydata-3*deltay, 'DIAG: '+ diagnostico[aux_isapre_infopacientes[l]])
                               canvas1.drawString(xdata,ydata-4*deltay, 'MÉDICO: '+ ' ')
                               
                               canvas1.drawString(xdata-27,ydata-6*deltay, 'CÓDIGO')
                               canvas1.drawString(xdata+30,ydata-6*deltay, 'GLOSA')
                               #canvas1.drawString(xdata+233,ydata-6*deltay, 'N°')
                               canvas1.drawString(xdata+252,ydata-6*deltay, 'FECHA')
                               canvas1.drawString(xdata+317,ydata-6*deltay, 'PROFESIONAL')
                               canvas1.drawString(xdata+435,ydata-6*deltay, 'RUT')
                               medico_nombre=0
                               canvas1.setFont('Helvetica', 10)

                               
                           if len(aux_sq)>0:
                               pdf_name = isapre[aux_isapre_infopacientes[l]]+'_'+ mes_nombre[j] + '_SQ_' + nombre_paciente[aux_isapre_infopacientes[l]]+'.pdf'
                               save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], pdf_name)
                               
                               if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    save_name = os.path.join(os.path.expanduser("~"), PATHES[1]%agno[gn], 'GESMED')
                                    
                               canvas2 = Canvas(save_name, pagesize=LETTER)
                               canvas2.setFont('Helvetica', 12)
                               canvas2.drawImage('vitalsur.png',220,650,170,100,anchor='sw',anchorAtXY=True,showBoundary=False)
                               xdata=50
                               deltax=250
                               ydata=550
                               deltay=20
                               canvas2.setFont('Helvetica-Bold', 16)
                               canvas2.drawString(150,625, 'PLANILLA DE ATENCIÓN GES PALIATIVO')
                               canvas2.setFont('Helvetica', 13)
                               canvas2.drawString(xdata,ydata, 'NOMBRE: '+ nombre_paciente[aux_isapre_infopacientes[l]])
                               canvas2.drawString(xdata+deltax,ydata, 'RUT: '+ rut_completo[aux_isapre_infopacientes[l]])
                               canvas2.drawString(xdata,ydata-deltay, 'MES: Segunda Quincena '+ mes_nombre[j] )
                               canvas2.drawString(xdata+deltax,ydata-deltay, 'PREVISIÓN: '+ isapre[aux_isapre_infopacientes[l]])
                               canvas2.drawString(xdata,ydata-2*deltay, 'AMB: '+ ' ' )
                               canvas2.drawString(xdata+deltax,ydata-2*deltay, 'PAM: '+ ' ' )
                               canvas2.drawString(xdata,ydata-3*deltay, 'DIAG: '+ diagnostico[aux_isapre_infopacientes[l]])
                               canvas2.drawString(xdata,ydata-4*deltay, 'MÉDICO: '+ ' ')
                               canvas2.drawString(xdata-27,ydata-6*deltay, 'CÓDIGO')
                               canvas2.drawString(xdata+30,ydata-6*deltay, 'GLOSA')
                               #canvas2.drawString(xdata+233,ydata-6*deltay, 'N°')
                               canvas2.drawString(xdata+252,ydata-6*deltay, 'FECHA')
                               canvas2.drawString(xdata+317,ydata-6*deltay, 'PROFESIONAL')
                               canvas2.drawString(xdata+435,ydata-6*deltay, 'RUT')
                               medico_nombreb=0
                               canvas2.setFont('Helvetica', 10)
                               
                           linestablea=20
                           linestableb=20
                           
                           visita_a=0
                           visita_b=0
                           
                           #Loop para la cantidad de turnos hechos para ese paciente durante ese mes
                           for q in range(len(aux_mesrut_turnos)):
                             #Columna de pagos y cobros para el profesional de ese turno
                               valores_profesional=np.asarray(dfs3[profesional[aux_mesrut_turnos[q]]].tolist())
                         
                               aux_isapre_infoisapres=np.where(isapre_unique[i]==i_isapre)[0]
                               aux_ciudad_dondevive=np.where(ciudad[aux_isapre_infopacientes[l]]==i_ciudad)[0]
                               aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)) #info isapres
                               if len(aux_isapreyciudad)==0:
                                   aux_ciudad_dondevive=np.where('Otros'==i_ciudad)[0]
                               aux_sector_dondevive=np.where(sector[aux_isapre_infopacientes[l]]==i_sector)[0]
                               aux_isapreyciudad=list(set(aux_isapre_infoisapres) & set(aux_ciudad_dondevive)  & set(aux_sector_dondevive)) #info isapres

                               aux_cobro_infoisapres=np.where('Cobro'==i_cobropago)[0]
                               aux_pago_infoisapres=np.where('Pago'==i_cobropago)[0]
                                   
                               #Posicion de la fila correspondiente al caso del paciente (info_isapres.xlsx)
                               aux_poscobro_infoisapres=list(set(aux_isapreyciudad) & set(aux_cobro_infoisapres))
                               aux_pospago_infoisapres=list(set(aux_isapreyciudad) & set(aux_pago_infoisapres))
                                   
                               if len(aux_isapre_infoisapres)==0:
                                   ffwarning.write('Hay una isapre que no esta especificada en info_isapres.xlsx')
                                   ffwarning.write(' Check: '+ isapre[aux_rut_infopac] )
                                   ffwarning.write( '\n')
                                   print('WARNING8!')

                               if len(aux_poscobro_infoisapres)>1:
                                   ffwarning.write('Hay dos valores para el cobro del turno en info_isapres.xlsx')
                                   ffwarning.write(' Check: '+ i_isapre[aux_isapre_infoisapres[0]] + ' ' + i_ciudad[aux_ciudad_dondevive[0]] +  ' COBRO')
                                   ffwarning.write( '\n')

                               if len(aux_pospago_infoisapres)>1:
                                   ffwarning.write('Hay dos valores para el pago del turno en info_isapres.xlsx')
                                   ffwarning.write(' Check: ' +  i_isapre[aux_isapre_infoisapres[0]] + ' '+ i_ciudad[aux_ciudad_dondevive[0]] +  ' PAGO')
                                   ffwarning.write( '\n')

                               if len(aux_pospago_infoisapres)==0 or len(aux_poscobro_infoisapres)==0:
                                   ffwarning.write('Hay un error en leer el cobro/pago de la isapre - info_isapres.xlsx')
                                   ffwarning.write(' Check: '+ isapre[aux_rut_infopac] )
                                   ffwarning.write( '\n')
                                   print('WARNING9!')

                               cobrar_mes_paciente=cobrar_mes_paciente+(valores_profesional[aux_poscobro_infoisapres])
                               cobrar_mes=cobrar_mes+(valores_profesional[aux_poscobro_infoisapres])
                               
                               if profesional[aux_mesrut_turnos[q]] == 'Medico':
                                   pMedico[j]=pMedico[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Enfermera':
                                   pEnfermera[j]=pEnfermera[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Nutricionista':
                                   pNutricionista[j]=pNutricionista[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Kinesiologo':
                                   pKinesiologo[j]=pKinesiologo[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Psicologo':
                                   pPsicologo[j]=pPsicologo[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Fonoaudiologo':
                                   pFono[j]=pFono[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Terapeuta Ocupacional':
                                   pTO[j]=pTO[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Tens':
                                   pTens[j]=pTens[j]+1
                                   definicion= 'Tens'
                               elif profesional[aux_mesrut_turnos[q]] == 'Curacion Simple':
                                   pCuracion_Simple[j]=pCuracion_Simple[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Educacion Enfermera':
                                   pEducacion_Enfermera[j]=pEducacion_Enfermera[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Educacion Tens':
                                   pEducacion_Tens[j]=pEducacion_Tens[j]+1
                               elif profesional[aux_mesrut_turnos[q]] == 'Intervension Psicosocial':
                                   pIntervension_Psicosocial[j]=pIntervension_Psicosocial[j]+1

                               ff3.write( '            ' )
                               ff3.write('Visita' + str(format(q+1)) +  '\n' )
                               ff3.write( '            ' )
                               ff3.write('Tipo Visita = ' + str(format(profesional[aux_mesrut_turnos[q]])) +  '\n' )
                               ff3.write( '            ' )
                               ff3.write('Profesional = ' + str(format(rut_profesional[aux_mesrut_turnos[q]])) +  '\n' )
                               ff3.write( '            ' )
                               ff3.write('Fecha = ' + str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])) +  '\n' )
                               ff3.write( '            ' )
                               ff3.write('Cobro = ' + str(format((valores_profesional[aux_poscobro_infoisapres[0]]))) +  '\n' )
                               ff3.write( '\n')

                               data_prof=np.where(rut_profesional[aux_mesrut_turnos[q]]==i_rut_profesional)[0]
                               bono_isapre=np.asarray(dfs5['Isapre'].tolist())
                               bono_cur=np.asarray(dfs5[profesional[aux_mesrut_turnos[q]]].tolist())
                               data_bono=np.where(isapre[aux_isapre_infopacientes[l]]==bono_isapre)[0]
                                
                               comp1= len(list(set(profesional[aux_mesrut_turnos][aux_pq])))
                               comp2= len(list(set(rut_profesional[aux_mesrut_turnos][aux_pq])))
                               comp1b= len(list(set(profesional[aux_mesrut_turnos][aux_sq])))
                               comp2b= len(list(set(rut_profesional[aux_mesrut_turnos][aux_sq])))
                               
                               comp1=100
                               comp2=200
                               comp1b=300
                               comp2b=400
             
             
                               if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    codigo_bono_gesmed.append(bono_cur[data_bono[0]])
                                    valor_bono_gesmed.append(valores_profesional[aux_poscobro_infoisapres])
                                    tipo_bono_gesmed.append(profesional[aux_mesrut_turnos[q]])
             
                               if dia[aux_mesrut_turnos[q]]<16 and (comp1==comp2):
                                   if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                        codigo_bono_pq.append(bono_cur[data_bono[0]])
                                        valor_bono_pq.append(valores_profesional[aux_poscobro_infoisapres])
                                        tipo_bono_pq.append(bono_cur[0])
                                        
                                    
                                   canvas1.setFont('Helvetica', 10)
                                   canvas1.drawString(xdata+252,ydata-6*deltay-linestablea,  str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])))
                                   
                                   thing_list=list(profesionaltodel_pq)
                                   elem=profesional[aux_mesrut_turnos[q]]
                                   aux_del = thing_list.index(elem) if elem in thing_list else -1
                                   cantidad=np.where(profesional[aux_mesrut_turnos[q]]==profesional[aux_mesrut_turnos][aux_pq])[0]
                                   
                                   if aux_del>-1:
                                       profesionaltodel_pq = np.delete(profesionaltodel_pq, aux_del)
                                       canvas1.drawString(xdata-27,ydata-6*deltay-linestablea, bono_cur[data_bono[0]])
                                       canvas1.drawString(xdata+30,ydata-6*deltay-linestablea, str(bono_cur[0]))
                                       #canvas1.drawString(xdata+233,ydata-6*deltay-linestablea, '%i' %len(cantidad))
                                       canvas1.drawString(xdata+317,ydata-6*deltay-linestablea, i_nombre_profesional[data_prof][0])
                                       canvas1.drawString(xdata+435,ydata-6*deltay-linestablea, i_rutcompleto[data_prof][0])
                                       
                                   if profesional[aux_mesrut_turnos[q]]=='Medico' and medico_nombre==0:
                                       canvas1.setFont('Helvetica', 13)
                                       canvas1.drawString(xdata+60,ydata-4*deltay, i_nombre_profesional[data_prof][0])
                                       medico_nombre=1
                                   visita_a=visita_a+1
                                   if visita_a>17:
                                       canvas1.showPage()
                                       visita_a=0
                                       linestablea=-260
                                   linestablea=linestablea+20
                                       
                               elif dia[aux_mesrut_turnos[q]]<16 and (comp1!=comp2):
                                   if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                        codigo_bono_pq.append(bono_cur[data_bono[0]])
                                        valor_bono_pq.append(valores_profesional[aux_poscobro_infoisapres])
                                        tipo_bono_pq.append(bono_cur[0])
                                        
                                   canvas1.setFont('Helvetica', 10)
                                   canvas1.drawString(xdata-27,ydata-6*deltay-linestablea, bono_cur[data_bono[0]])
                                   canvas1.drawString(xdata+30,ydata-6*deltay-linestablea, str(bono_cur[0]))
                                   canvas1.drawString(xdata+252,ydata-6*deltay-linestablea,  str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])))
                                   canvas1.drawString(xdata+317,ydata-6*deltay-linestablea, i_nombre_profesional[data_prof][0])
                                   canvas1.drawString(xdata+435,ydata-6*deltay-linestablea, i_rutcompleto[data_prof][0])
                                   thing_list=list(profesionaltodel_pq)
                                   elem=profesional[aux_mesrut_turnos[q]]
                                   aux_del = thing_list.index(elem) if elem in thing_list else -1
                                   cantidad=np.where(profesional[aux_mesrut_turnos[q]]==profesional[aux_mesrut_turnos][aux_pq])[0]
                                   if aux_del>-1:
                                       profesionaltodel_pq = np.delete(profesionaltodel_pq, aux_del)
                                       #canvas1.drawString(xdata+233,ydata-6*deltay-linestablea, '%i' %len(cantidad))
                                   visita_a=visita_a+1
                                   if visita_a>17:
                                       canvas1.showPage()
                                       visita_a=0
                                       linestablea=-260
                                   linestablea=linestablea+20
                                   
                               
                               elif dia[aux_mesrut_turnos[q]]>15 and (comp1b==comp2b):
                                   if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                        codigo_bono_sq.append(bono_cur[data_bono[0]])
                                        valor_bono_sq.append(valores_profesional[aux_poscobro_infoisapres])
                                        tipo_bono_sq.append(bono_cur[0])
                               
                                   canvas2.setFont('Helvetica', 10)
                                   canvas2.drawString(xdata+252,ydata-6*deltay-linestableb,  str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])))
                                   
                                   thing_list=list(profesionaltodel_sq)
                                   elem=profesional[aux_mesrut_turnos[q]]
                                   aux_del = thing_list.index(elem) if elem in thing_list else -1
                                   cantidad=np.where(profesional[aux_mesrut_turnos[q]]==profesional[aux_mesrut_turnos][aux_sq])[0]
                                   
                                   if aux_del>-1:
                                       profesionaltodel_sq = np.delete(profesionaltodel_sq, aux_del)
                                       canvas2.drawString(xdata-27,ydata-6*deltay-linestableb, bono_cur[data_bono[0]])
                                       canvas2.drawString(xdata+30,ydata-6*deltay-linestableb, str(bono_cur[0]))
                                       #canvas2.drawString(xdata+233,ydata-6*deltay-linestableb, '%i' %len(cantidad))
                                       canvas2.drawString(xdata+317,ydata-6*deltay-linestableb, i_nombre_profesional[data_prof][0])
                                       canvas2.drawString(xdata+435,ydata-6*deltay-linestableb, i_rutcompleto[data_prof][0])
                                   
                                   if profesional[aux_mesrut_turnos[q]]=='Medico' and medico_nombreb==0:
                                       canvas2.setFont('Helvetica', 13)
                                       canvas2.drawString(xdata+60,ydata-4*deltay, i_nombre_profesional[data_prof][0])
                                       medico_nombre=1
                                   visita_b=visita_b+1
                                   
                                   if visita_b>17:
                                       canvas2.showPage()
                                       visita_b=0
                                       linestableb=-260
                                   linestableb=linestableb+20

                               elif dia[aux_mesrut_turnos[q]]>15 and (comp1b!=comp2b):
                                   if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                        codigo_bono_sq.append(bono_cur[data_bono[0]])
                                        valor_bono_sq.append(valores_profesional[aux_poscobro_infoisapres])
                                        tipo_bono_sq.append(bono_cur[0])
                               
                                   canvas2.setFont('Helvetica', 10)
                                   canvas2.drawString(xdata-27,ydata-6*deltay-linestableb, bono_cur[data_bono[0]])
                                   canvas2.drawString(xdata+30,ydata-6*deltay-linestableb, str(bono_cur[0]))
                                   canvas2.drawString(xdata+252,ydata-6*deltay-linestableb,  str(format(dia[aux_mesrut_turnos[q]])) +  '/' + str(format(mes[aux_mesrut_turnos[q]])) +  '/' + str(format(ano[aux_mesrut_turnos[q]])))
                                   canvas2.drawString(xdata+317,ydata-6*deltay-linestableb, i_nombre_profesional[data_prof][0])
                                   canvas2.drawString(xdata+435,ydata-6*deltay-linestableb, i_rutcompleto[data_prof][0])
                                         
                                   thing_list=list(profesionaltodel_sq)
                                   elem=profesional[aux_mesrut_turnos[q]]
                                   aux_del = thing_list.index(elem) if elem in thing_list else -1
                                   cantidad=np.where(profesional[aux_mesrut_turnos[q]]==profesional[aux_mesrut_turnos][aux_sq])[0]
                                   
                                   if aux_del>-1:
                                       profesionaltodel_sq = np.delete(profesionaltodel_sq, aux_del)
                                       #canvas2.drawString(xdata+233,ydata-6*deltay-linestableb, '%i' %len(cantidad))
                                   visita_b=visita_b+1
                                   if visita_b>17:
                                       canvas2.showPage()
                                       visita_b=0
                                       linestableb=-260
                                   linestableb=linestableb+20


                           if len(aux_pq)>0:
                               canvas1.drawString(xdata+190,ydata-6*deltay-linestablea-30, 'Total = %i' %len(aux_pq))
                               canvas1.save()
                               if (sector[aux_isapre_infopacientes[l]]>0):
                                   ff4mail.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) + ' - Sector ' +str(sector[aux_isapre_infopacientes[l]]) +') ')
                               else:
                                   ff4mail.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) +') ')
                               ff4mail.write('\n')
                               
                               if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                    ff4valores.write('\n')
                                    ff4valores.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(rut_completo[aux_isapre_infopacientes[l]]) +') \n')
                                    bonos_unique_pq = list(set(codigo_bono_pq))
                                    for iii in range(len(bonos_unique_pq)):
                                        a=np.where(bonos_unique_pq[iii]==np.asarray(codigo_bono_pq))[0]
                                        ff4valores.write( str(bonos_unique_pq[iii]) + '  ')
                                        ff4valores.write( str(tipo_bono_pq[a[0]]) + '  ')
                                        ff4valores.write( str(len(a)) + '  ')
                                        ff4valores.write( str( int(len(a)* valor_bono_pq[a[0]] ) ) + '  \n')
                                    
                           if len(aux_sq)>0:
                                canvas2.drawString(xdata+190,ydata-6*deltay-linestableb-30, 'Total = %i' %len(aux_sq))
                                canvas2.save()
                                if (sector[aux_isapre_infopacientes[l]]>0):
                                    ff4mailsq.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) + ' - Sector ' +str(sector[aux_isapre_infopacientes[l]]) +') ')
                                else:
                                    ff4mailsq.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' (' + str(ciudad[aux_isapre_infopacientes[l]]) +') ')
                                ff4mailsq.write('\n')
                                
                                if isapre[aux_isapre_infopacientes[l]]!='Gesmed':
                                    ff4valoressq.write('\n')
                                    ff4valoressq.write( str(nombre_paciente[aux_isapre_infopacientes[l]]) + ' ('  + str(rut_completo[aux_isapre_infopacientes[l]]) +') \n ')
                               
                                    bonos_unique_sq = list(set(codigo_bono_sq))
                                    for iii in range(len(bonos_unique_sq)):
                                        a=np.where(bonos_unique_sq[iii]==np.asarray(codigo_bono_sq))[0]
                                        ff4valoressq.write( str(bonos_unique_sq[iii]) + '  ')
                                        ff4valoressq.write( str(tipo_bono_sq[a[0]]) + '  ')
                                        ff4valoressq.write( str(len(a)) + '  ')
                                        ff4valoressq.write( str( int(len(a)* valor_bono_sq[a[0]] ) ) + ' \n')
                             
                             
                            
                           if isapre[aux_isapre_infopacientes[l]]=='Gesmed':
                                    bonos_unique_gesmed = list(set(codigo_bono_gesmed))
                                    for iii in range(len(bonos_unique_gesmed)):
                                        a=np.where(bonos_unique_gesmed[iii]==np.asarray(codigo_bono_gesmed))[0]

                                        worksheet.write(row, 0,     str(j+1) + '/' + str(agno[gn]))
                                        worksheet.write(row, 1,     str(nombre_paciente[aux_isapre_infopacientes[l]]))
                                        worksheet.write(row, 2,     str(rut_completo[aux_isapre_infopacientes[l]]))
                                        worksheet.write(row, 3,     str(format(ciudad[aux_isapre_infopacientes[l]])))
                                        worksheet.write(row, 4,     str(tipo_bono_gesmed[a[0]]) )
                                        worksheet.write(row, 5,     str(len(a)))
                                        worksheet.write(row, 6,     str( int(valor_bono_gesmed[a[0]] ) ) )
                                        worksheet.write(row, 7,     str( int(len(a)* valor_bono_gesmed[a[0]] ) ) )
                                        
                                        if (sector[aux_isapre_infopacientes[l]]>0):
                                            worksheet.write(row, 8,    str(sector[aux_isapre_infopacientes[l]]))
                                            
                                        else:
                                            worksheet.write(row, 8,    ' REGION')

                                            
                                        row += 1
                            
                           ff3.write( '\n')
                           ff3.write( '            ' )
                           ff3.write('Cobrar a Isapre por paciente =' + str(format(int(cobrar_mes_paciente))) +  ' ' )
                           ff3.write( '\n')
                           ff3.write( '\n')
                           
                   ff4mail.write('Esperando una buena acogida y su pronta respuesta. \n Saluda cordialmente a usted. \n \n \n')
                   ff4mailsq.write('Esperando una buena acogida y su pronta respuesta. \n Saluda cordialmente a usted. \n \n \n')
                   
                   ff3.write( '\n')
                   ff3.write( '\n')
                   ff3.write( '                      .......................................................................................' )
                   ff3.write( '\n')
                   ff3.write( '\n')
                   ff3.write( '                      ' )
                   ff3.write('Cobrar a Isapre mes de ' + str(mes_nombre[j]) + ' = ' + str(format(int(cobrar_mes))) +  ' ' )
                   ff3.write( '\n')
                   ff3.write( '\n')
                   ff3.write( '\n')

               out_tpl = np.nonzero(pmes)

               
               fig = plt.figure()
               ax = plt.subplot(111)
               ax.plot(pmes[out_tpl], pMedico[out_tpl], '.-', label='Medico', color='red')
               ax.plot(pmes[out_tpl], pEnfermera[out_tpl], '.-', label='Enfermera', color='blue')
               ax.plot(pmes[out_tpl], pNutricionista[out_tpl], '.-', label='Nutricionista', color='green')
               ax.plot(pmes[out_tpl], pKinesiologo[out_tpl], '.-', label='Kinesiologo', color='magenta')
               ax.plot(pmes[out_tpl], pPsicologo[out_tpl], '.-', label='Psicologo', color='purple')
               ax.plot(pmes[out_tpl], pTens[out_tpl], '.-', label='Tens', color='green')
               ax.plot(pmes[out_tpl], pCuracion_Simple[out_tpl], '.-', label='C Simple', color='cyan')
               ax.plot(pmes[out_tpl], pEducacion_Enfermera[out_tpl], '.-', label='Ed Enfermera', color='black')
               ax.plot(pmes[out_tpl], pEducacion_Tens[out_tpl], '.-', label='Ed Tens', color='yellow')
               ax.plot(pmes[out_tpl], pIntervension_Psicosocial[out_tpl], '.-', label='In Psicosocial', color='gold')
               ax.set_xlabel('Mes')
               ax.set_ylabel('Numero de Visitas')
               ax.set_title('%s' %isapre_unique[i], fontsize=20)
               ax.set_xlim(0.5, max(pmes)+0.5)
               ax.set_xticks(np.arange(1, max(pmes)+0.5, 1))
               box = ax.get_position()
               ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
               ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
               #save_name = os.path.join(os.path.expanduser("~"), PATHES[0]%agno[gn], 'Vitalsur_Isapre_%s_%i.png' %(isapre_unique[i], agno[gn]))
               save_name = os.path.join(os.path.curdir, "%i/"%agno[gn],  'Vitalsur_Isapre_%s_%i.png' %(isapre_unique[i], agno[gn]))
               plt.savefig(save_name)
                               

    workbook.close()
