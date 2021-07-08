# -*- coding: utf-8 -*-
"""
Construccion del bot de programacion y alertamiento.

"""
import pandas as pd
import pyodbc
import getpass
import openpyxl


#%% Datos eficiencia

efiencias_data = """SELECT    VIPS_AGENDA.Mes_Cita_Id Mes_Cita_Id,
          VIPS_AGENDA.Sede_Id Sede_Id,
          VIPS_SEDE.Codigo_Sede_Op Codigo_Sede_Op,
          VIPS_SEDE.Nombre_Sede Nombre_Sede,
          CASE WHEN VC_SERVICIO_CITA.Codigo_Servicio_Cita_Op  = 94108 THEN '50110' ELSE VC_SERVICIO_CITA.Codigo_Servicio_Cita_Op END AS Codigo_Servicio_Cita_Op,
		  CASE WHEN VC_SERVICIO_CITA.Servicio_Cita_Desc = 'MEDICINA GENERAL' THEN 'CONSULTA MEDICINA GENERAL SALUD (CITA PRESENCIAL)' ELSE VC_SERVICIO_CITA.Servicio_Cita_Desc END AS Servicio_Cita_Desc,
          Avg(VIPS_AGENDA.Numero_Intervalo_Creacioncita) Duracion,
          Sum(VIPS_AGENDA.Numero_Citas_Creadas) Creadas,
          Sum(VIPS_AGENDA.Numero_Citas_Bloqueadas) Bloqueadas,
          Sum(VIPS_AGENDA.Numero_Citas_Ofertadas) Ofertadas,
          Sum(VIPS_AGENDA.Numero_Citas_Disponibles) Disponibles,
          Sum(VIPS_AGENDA.Numero_Citas_Asignadas) Asignadas,
          Sum(VIPS_AGENDA.Numero_Citas_Confirmadas) Confirmadas,
          Sum(VIPS_AGENDA.Numero_Citas_Inasistidas) Inasistidas,
          Sum(VIPS_AGENDA.Numero_Citas_Canceladas) Canceladas,
          Sum(VIPS_AGENDA.Numero_Citas_Confir_sin_aten) Citas_Confir_sin_aten,
          Sum(VIPS_AGENDA.Numero_Citas_Pend_x_asistir) itas_Pend_x_asistir,
          Sum(VIPS_AGENDA.Numero_CItas_Gestion) CItas_Gestion,
          Sum(VIPS_AGENDA.Numero_Tiempo_Admon_Otro) Tiempo_Admon_Otro
          
         FROM      MDB_EPS_IPS_COLOMBIA.VIPS_AGENDA VIPS_AGENDA
          JOIN  MDB_EPS_IPS_COLOMBIA.VIPS_SEDE VIPS_SEDE ON VIPS_SEDE.Sede_Id = VIPS_AGENDA.Sede_Id
          JOIN  MDB_EPS_IPS_COLOMBIA.VM_EPS_CLIENTE_NOMBRE VM_EPS_CLIENTE_NOMBRE  ON VIPS_AGENDA.Cliente_Profesional_Id = VM_EPS_CLIENTE_NOMBRE.Cliente_Id
          JOIN  MDB_EPS_IPS_COLOMBIA.VC_SERVICIO_CITA VC_SERVICIO_CITA  ON VC_SERVICIO_CITA.Servicio_Cita_Id = VIPS_AGENDA.Servicio_Cita_Id AND VC_SERVICIO_CITA.Especialidad_Cita_Id = VIPS_AGENDA.Especialidad_Cita_Id
          WHERE VIPS_AGENDA.Fecha_Cita BETWEEN '2021-01-01'  AND  '2021-06-30' AND VIPS_AGENDA.Sede_Id <> -1
           GROUP  BY 1,2,3,4,5,6"""
           
           

           
#%% Datos Oportunidad
 
Oportunidad_data   = """select    VIPS_SEDE.Codigo_Sede_Op Codigo_Sede_Op,
                  VIPS_SEDE.Nombre_Sede Nombre_Sede,
                  VIPS_AGENDA.Causa_Oportunidad_Id Causa_Oportunidad_Id,
                  VC_CAUSA_OPORTUNIDAD_AGENDA.Causa_Oportunidad_Desc Causa_Oportunidad_Desc,
                  sum(VIPS_AGENDA.Numero_Dias_Habiles) habiles,
                  sum(VIPS_AGENDA.Numero_Dias_Calendario) calendario,
                  sum(VIPS_AGENDA.Numero_Citas_Asignadas) asignadas, 
                  (cast(habiles as float)/asignadas) as oport,
                  VC_SERVICIO_CITA.Codigo_Servicio_Cita_Op Codigo_Servicio_Cita_Op,
                  VC_SERVICIO_CITA.Servicio_Cita_Desc Servicio_Cita_Desc,
                  VIPS_AGENDA.Mes_Cita_Id Mes_Cita_Id 
            from   MDB_EPS_IPS_COLOMBIA.VC_SERVICIO_CITA VC_SERVICIO_CITA 
                  join (MDB_EPS_IPS_COLOMBIA.VC_CAUSA_OPORTUNIDAD_AGENDA VC_CAUSA_OPORTUNIDAD_AGENDA join (MDB_EPS_IPS_COLOMBIA.VIPS_SEDE VIPS_SEDE join MDB_EPS_IPS_COLOMBIA.VIPS_AGENDA VIPS_AGENDA 
                  on (VIPS_SEDE.Sede_Id = VIPS_AGENDA.Sede_Id)) 
                  on (VC_CAUSA_OPORTUNIDAD_AGENDA.Causa_Oportunidad_Id = VIPS_AGENDA.Causa_Oportunidad_Id)) 
                  on ((VC_SERVICIO_CITA.Servicio_Cita_Id = VIPS_AGENDA.Servicio_Cita_Id and VC_SERVICIO_CITA.Especialidad_Cita_Id = VIPS_AGENDA.Especialidad_Cita_Id)) 
            where ((VIPS_AGENDA.Fecha_Cita >= DATE '2021-01-01' and VIPS_SEDE.Codigo_Sede_Op <> '-1') and VIPS_AGENDA.Causa_Oportunidad_Id <> 15491 and VIPS_AGENDA.Fecha_Cita <= DATE '2021-06-30' ) 
            group  by VIPS_SEDE.Codigo_Sede_Op,
                  VIPS_SEDE.Nombre_Sede,
                  VIPS_AGENDA.Causa_Oportunidad_Id,
                  VC_CAUSA_OPORTUNIDAD_AGENDA.Causa_Oportunidad_Desc,
                  VC_SERVICIO_CITA.Codigo_Servicio_Cita_Op,
                  VC_SERVICIO_CITA.Servicio_Cita_Desc,
                  VIPS_AGENDA.Mes_Cita_Id"""
                  
                  
#%% Datos Capacidad 

Capacidad_data = """SELECT * from user_stage.wilmagju_resumen_capacidad
                    WHERE EspacioFisico = 'Consultorio'"""
                    
                    
#serv_consul_data = """SELECT * FROM user_stage.wilmagju_maestra_prestacion"""

#%%Solicitudes de citas
solicitudescitas_data = """SELECT  VIPS_AGENDA.Mes_Solicitud_Id Mes_Solicitud_Id,
		--VIPS_AGENDA.Fecha_Solicitud_Cita Fecha_Solicitud_Cita,
        VIPS_SEDE.Codigo_Sede_Op Codigo_Sede_Op,
        VIPS_SEDE.Nombre_Sede Nombre_Sede,
		VC_SERVICIO_CITA.Codigo_Servicio_Cita_Op Codigo_Servicio_Cita_Op ,
        VC_SERVICIO_CITA.Servicio_Cita_Desc Servicio_Cita_Desc,
        
		Sum(VIPS_AGENDA.Numero_Citas_Asignadas) AS Solicitud
		
		FROM VIPS_AGENDA 
		LEFT JOIN VM_EPS_CLIENTE_NOMBRE VM_EPS_CLIENTE_NOMBRE ON VIPS_AGENDA.Cliente_Profesional_Id = VM_EPS_CLIENTE_NOMBRE.Cliente_Id
		LEFT JOIN VIPS_SEDE VIPS_SEDE ON VIPS_SEDE.Sede_Id = VIPS_AGENDA.Sede_Id 
		LEFT JOIN VC_SERVICIO_CITA VC_SERVICIO_CITA ON VC_SERVICIO_CITA.Servicio_Cita_Id = VIPS_AGENDA.Servicio_Cita_Id
		WHERE VIPS_AGENDA.Numero_Citas_Asignadas > 0 AND VIPS_AGENDA.Sede_Id <> -1 AND VIPS_AGENDA.Fecha_Solicitud_Cita BETWEEN '2021-01-01' AND '2021-06-30'
		GROUP BY 1,2,3,4,{}""".format(5)

#%% Conectar y ejectuar modelo de datos
# Probando conexion a traves del modulo de pyodbc
#password = getpass.getpass(prompt='Ingrese contraseña ')
print('Conectado con la base de datos')

connection = pyodbc.connect('DSN=TERADATA2;UID=wilmagju;PWD=Magangue123#')

print('Se realiza la conexion correctamente')

# Construccion de los dataframes
print('Consultando datos.....')
Eficiencia  = pd.read_sql(efiencias_data,connection)
Oportunidad  = pd.read_sql(Oportunidad_data,connection)
Capacidad = pd.read_sql(Capacidad_data,connection)
#serv_consul_data = pd.read_sql(serv_consul_data,connection)
solicitudes = pd.read_sql(solicitudescitas_data,connection)
print('Datos consultados con exito')

Eficiencia1  = Eficiencia
Oportunidad1  = Oportunidad
Capacidad1 = Capacidad
#serv_consul_data1 = serv_consul_data 
solicitudes1 = solicitudes


Eficiencia1['Inasi'] = round(Eficiencia1.Inasistidas/Eficiencia1.Asignadas,2)
Eficiencia1['Dispo'] = round(Eficiencia1.Disponibles/Eficiencia1.Ofertadas,2)
Eficiencia1['Bloq'] = round(Eficiencia1.Bloqueadas/Eficiencia1.Creadas,2)
Eficiencia1['apro'] = round(Eficiencia1.Inasistidas/Eficiencia1.Creadas,2)

connection.close()  

#%% Modelo de variables y diccionario vacio



Sedes = [['42','54','93','2218','80','1788','20','51','2660','1709','2187','79','2666','78','1789','2694','2703'],
          ['IPS SURA ALMACENTRO','IPS SURA ALTOS DEL PRADO','IPS SURA BOSTON','IPS SURA BUCARAMANGA','IPS SURA CENTRO','IPS SURA CHAPINERO ','IPS SURA CORDOBA ','IPS SURA LA FLORA','IPS SURA LOS MOLINOS','IPS SURA MONTERREY ','IPS SURA MURILLO','IPS SURA OLAYA','IPS SURA PASO ANCHO','IPS SURA RIONEGRO','IPS SURA SAMAN','IPS SURA SAN DIEGO','IPS SURA TEQUENDAMA'],
          ['wagamez@sura.com.co','wagamez@sura.com.co','wagamez@sura.com.co','wagamez@sura.com.co','wilmer070@gmail.com','wagamez@sura.com.co','wilmer070@gmail.com','wilmer070@gmail.com','wilmer070@gmail.com','wilmer070@gmail.com','wilmer070@gmail.com','wilmer070@gmail.com','wilmer070@gmail.com','wilmer070@gmail.com','wagamez@unal.edu.co','wagamez@unal.edu.co','salcaraz@sura.com.co']]

Codigos= [['50110','50114','50120','50130','50140','50190','2','22','50270','18','16','70004','70005'],
                        ['CONSULTA MEDICO GENERAL','CONSULTA MEDICO GENERAL NO PROGRAMADA','CONSULTA MEDICINA INTERNA','CONTROL MEDICINA INTERNA','CONSULTA GINECOLOGO','CONSULTA DERMATOLOGIA','CONSULTA CARDIOLOGO ADULTO','CONSULTA ENDOCRINOLOGO','CONSULTA FISIATRA','CONTROL NEFROLOGO (A)','CONSULTA UROLOGIA','INGRESO PROGRAMA ASMA','INGRESO PROGAMA EPOC'],
                        [25,20,25,25,25,25,25,25,25,25,25,25,25],[3,1,7,7,15,20,20,20,7,20,20,3,3]] 

Meses = [202101,202102,202103,202104,202105,202106]

# Ejemplo para seleccionar m = Eficiencia1.loc[(Eficiencia1.Codigo_Servicio_Cita_Op=='50110'),['%Ind_aprovec']]
# Generacion de diccionario de datos con los datos principales a evaluar

dicopor = {}

for mt in Codigos[1]:
    dicopor[mt] = []
    
def actualizacion():
    confi_sedes2 = {}
    for i in Sedes[1]:
        confi_sedes2[i] = {'Afiliados':[],'Ocupacion_fisica':[],'Consultorios':[],'Medicos':[],'Codigos': 
                      {'CONSULTA MEDICO GENERAL': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONSULTA MEDICO GENERAL NO PROGRAMADA': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONSULTA MEDICINA INTERNA': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONTROL MEDICINA INTERNA': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONSULTA GINECOLOGO': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONSULTA DERMATOLOGIA': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONSULTA CARDIOLOGO ADULTO':{'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONSULTA ENDOCRINOLOGO': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONSULTA FISIATRA':{'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONTROL NEFROLOGO (A)': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'CONSULTA UROLOGIA': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'INGRESO PROGRAMA ASMA': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]},
                    'INGRESO PROGAMA EPOC': {'Oportunidad':[],'Solicitudes':[],'Bloqueadas':[],'Inasistencia':[],'Aprovechamiento':[],'Ofertadas':[],'Disponibles':[]}}}
    return confi_sedes2
    
        
dictio = {'Meses': Meses, 'Sedes': actualizacion()}


#%%Poblar afiliaciones y oportunidad

for i, t in enumerate(Sedes[0]):    
    consul = round(Capacidad1.loc[(Capacidad1.mes == 202105) & (Capacidad1.Cod_codigo == t)].Area.sum(),0)
    dictio['Sedes'][Sedes[1][i]]['Consultorios'] = consul
    for m in Meses:                   
        ocupacion = Capacidad1.loc[(Capacidad1.mes ==  m) & (Capacidad1.Cod_codigo == t)]
        afili= ocupacion.Sum_Afiliados.unique()[0]
        fisicos = round(ocupacion.capacidad_mes.unique()[0],2)
        dictio['Sedes'][Sedes[1][i]]['Afiliados'].append(afili)
        dictio['Sedes'][Sedes[1][i]]['Ocupacion_fisica'].append(fisicos)
        
        
#%% Poblar Conjunto de codigos

for i in range(len(Sedes[0])):
    for j in range(len(Codigos[0])):
        for m in Meses:
             Habiles = Oportunidad1[(Oportunidad1.Mes_Cita_Id == m) & (Oportunidad1.Codigo_Sede_Op == Sedes[0][i]) & (Oportunidad1.Codigo_Servicio_Cita_Op == Codigos[0][j])].habiles.sum()
             Asignadas = Oportunidad1[(Oportunidad1.Mes_Cita_Id == m) & (Oportunidad1.Codigo_Sede_Op == Sedes[0][i]) & (Oportunidad1.Codigo_Servicio_Cita_Op == Codigos[0][j])].asignadas.sum()
             bloq = round(Eficiencia1[(Eficiencia1.Mes_Cita_Id == m) & (Eficiencia1.Codigo_Sede_Op == Sedes[0][i]) & (Eficiencia1.Codigo_Servicio_Cita_Op == Codigos[0][j])].Bloq.sum(),2)
             ina = round(Eficiencia1[(Eficiencia1.Mes_Cita_Id == m) & (Eficiencia1.Codigo_Sede_Op == Sedes[0][i]) & (Eficiencia1.Codigo_Servicio_Cita_Op == Codigos[0][j])].Inasi.sum(),2)
             apro = round(Eficiencia1[(Eficiencia1.Mes_Cita_Id == m) & (Eficiencia1.Codigo_Sede_Op == Sedes[0][i]) & (Eficiencia1.Codigo_Servicio_Cita_Op == Codigos[0][j])].apro.sum(),2)
             ofer = round(Eficiencia1[(Eficiencia1.Mes_Cita_Id == m) & (Eficiencia1.Codigo_Sede_Op == Sedes[0][i]) & (Eficiencia1.Codigo_Servicio_Cita_Op == Codigos[0][j])].Ofertadas.sum(),2)
             Oport = round(Habiles/Asignadas,2) if Asignadas!= 0 else 0
             Sol  = solicitudes1[(solicitudes1.Mes_Solicitud_Id == m) & (solicitudes1.Codigo_Sede_Op == Sedes[0][i]) & (solicitudes1.Codigo_Servicio_Cita_Op == Codigos[0][j])].Solicitud.sum()
             dispo = round(Eficiencia1[(Eficiencia1.Mes_Cita_Id == m) & (Eficiencia1.Codigo_Sede_Op == Sedes[0][i]) & (Eficiencia1.Codigo_Servicio_Cita_Op == Codigos[0][j])].Disponibles.sum(),2)
             dictio['Sedes'][Sedes[1][i]]['Codigos'][Codigos[1][j]]['Oportunidad'].append(Oport)
             dictio['Sedes'][Sedes[1][i]]['Codigos'][Codigos[1][j]]['Solicitudes'].append(Sol)
             dictio['Sedes'][Sedes[1][i]]['Codigos'][Codigos[1][j]]['Bloqueadas'].append(bloq)
             dictio['Sedes'][Sedes[1][i]]['Codigos'][Codigos[1][j]]['Inasistencia'].append(ina)
             dictio['Sedes'][Sedes[1][i]]['Codigos'][Codigos[1][j]]['Aprovechamiento'].append(apro)
             dictio['Sedes'][Sedes[1][i]]['Codigos'][Codigos[1][j]]['Ofertadas'].append(ofer)
             dictio['Sedes'][Sedes[1][i]]['Codigos'][Codigos[1][j]]['Disponibles'].append(dispo)
             
             

                 
 #%% Anexo de medicos
 
for i in range(len(Sedes[0])):
    for y in range(len(Meses)):
        gene = (dictio['Sedes'][Sedes[1][i]]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][y])*(Codigos[2][0])
        priori = (dictio['Sedes'][Sedes[1][i]]['Codigos']['CONSULTA MEDICO GENERAL NO PROGRAMADA']['Ofertadas'][y])*(Codigos[2][1])
        med = round(((gene+priori)/60)/(197*(1-0.13)),2)
        dictio['Sedes'][Sedes[1][i]]['Medicos'].append(med)
        
        


#%%
def inasistencia(x):
    if dictio['Sedes'][x]['Codigos']['CONSULTA MEDICO GENERAL']['Inasistencia'][len(Meses)-1]>0.15:
        m5='hay que revisar como se consiguen eficiencias en el servicio desde la inasistencia ya que se encuentra en {}%'.format(dictio['Sedes'][x]['Codigos']['CONSULTA MEDICO GENERAL']['Inasistencia'][len(Meses)-1]*100)
    else:
        m5= 'las inasistencias al servicio es {}%  dentro de el rango normal'.format(dictio['Sedes'][x]['Codigos']['CONSULTA MEDICO GENERAL']['Inasistencia'][len(Meses)-1]*100)
    return m5

#%%
control = pd.DataFrame({'Id': [],
                       'Email':[],
                       'Asunto':[],
                       'Cuerpo':[],
                       'Date':[]})
           
for k,h in enumerate(dictio['Sedes']):
     # Revisar la oportunidad ultimo mes 
     if dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Oportunidad'][len(Meses)-1] == 0:
         m1  = 'la sede presenta una oportunidad de {} probablemente estuvo cerrada'.format(dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Oportunidad'][len(Meses)-1] )

         if dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][len(Meses)-1]==0:
             m2='y adicionalmente la oferta de citas para el servicio fue 0 durante el mes'
             m3=''
             
             
            
         elif dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Solicitudes'][len(Meses)-1]>dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][len(Meses)-1]:             
             m2='aunque presenta un numero de solicitudes de citas {} mayor a la oferta que tuve de {} se pueden presentar anomalias en el dato'.format(dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Solicitudes'][len(Meses)-1],dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][len(Meses)-1]) 
             m3=''
             
         elif dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Solicitudes'][len(Meses)-1]<dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][len(Meses)-1]:             
             m2='tal vez pueda deberse a la amplia oferta de citas {} con respecto a las solicitudes {} hay que revisar si es necesaria una reducción'.format(dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][len(Meses)-1],dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Solicitudes'][len(Meses)-1]) 
             m3=''
             
        
        
         
         
         
         

     elif dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Oportunidad'][len(Meses)-1]> 3:
         m1= 'por encima del rango permitido de 3'
         
         
         if dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Solicitudes'][len(Meses)-1]>dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][len(Meses)-1]:
             m2= 'Esto puede deberse a que las solicitud de citas fue {} mayor que la oferta de la sede que esta actualmente en {}'.format(dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Solicitudes'][len(Meses)-1],dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][len(Meses)-1])
             if dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Bloqueadas'][len(Meses)-1]>0.10:
                 m3= 'sin embargo {}% de las citas para este servicio estuvieron bloqueadas '.format(dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Bloqueadas'][len(Meses)-1]*100)
             else:
                 m3= 'no se presenteron bloqueos representativos este fue de {}%'.format(dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Bloqueadas'][len(Meses)-1]*100)
                 
     
        
     else:
         
         m1 = 'dentro del rango normal'
         if dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Solicitudes'][len(Meses)-1]>dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][len(Meses)-1]:
             m2 = 'aun así estamos generando una mayor cantidad de solicitudes {} en comparacion a la oferta {} lo que significa que debemos estar en permanente vigilancia '.format(dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Solicitudes'][len(Meses)-1],dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Ofertadas'][len(Meses)-1])
             if dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Bloqueadas'][len(Meses)-1]>0.10:
                 m3= 'ademas de prestarle atencion a la alta cantidad de bloqueos que representan un {}%'.format(dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Bloqueadas'][len(Meses)-1]*100)
         
         
    
    
     mensaje = """
     
Para la Sede de {} tenemos el ultimo mes una afiliacion de {:,.0f} esto es un {} % con respecto al mes anterior \
logrando un uso de sus capacidades al {}% el servicio de CONSULTA MEDICO GENERAL tuvo oportunidad de {} {} {} {} {}
""".format(
     h,
     dictio['Sedes'][h]['Afiliados'][len(Meses)-1],
     round(1-dictio['Sedes'][h]['Afiliados'][len(Meses)-1]/dictio['Sedes'][h]['Afiliados'][len(Meses)-2],2),
     dictio['Sedes'][h]['Ocupacion_fisica'][len(Meses)-1]*100,
     dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL']['Oportunidad'][len(Meses)-1],m1,m2,m3,inasistencia(h))
    
    # Mensaje 2
     dispo = dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL NO PROGRAMADA']['Disponibles'][len(Meses)-1]
     ofert = dictio['Sedes'][h]['Codigos']['CONSULTA MEDICO GENERAL NO PROGRAMADA']['Ofertadas'][len(Meses)-1]
     rel = round((dispo/ofert if ofert!= 0 else 0), 2)
     if ofert==0:
         m7='Para medicina prioritaria la oferta fue 0 debido a que la sede probablemente estuvo cerrada'
     elif rel<0.10:
         m7='Para servicios como medicina prioritaria de la oferta que generamos en el ultimo mes de '+str(ofert)+' aproximandamente el '+str(rel)+'% aprovechandose correctamente las citas'
     elif rel >0.10:
         m7='Para servicios como medicina prioritaria de la oferta que generamos en el ultimo mes de '+str(ofert)+' aproximandamente el '+str(rel)+'% es un poco mas alto de lo permitido surgiendo la necesidad de hacer un mejor aprovechamiento de las citas'
    
    
     print(mensaje)
     print(m7)
    
     
    # Agregar a la base de datos
     control=control.append({'Id': k,
                       'Email':Sedes[2][k],
                       'Asunto':'Datos de la IPS {}'.format(h),
                       'Cuerpo':mensaje,
                       'Date':[]},ignore_index=True)
   
     
     

control.to_excel(r"D:\Usuarios\wilmagju\OneDrive - Seguros Suramericana, S.A\25.Modelo completo\correosmasivos.xlsx", index = False)
control.to_csv(r"D:\Usuarios\wilmagju\OneDrive - Seguros Suramericana, S.A\25.Modelo completo\correosmasivos.csv", index=False)   
     
#%%graficas
import pickle
a_file = open("dictio.pkl", "rb")
dictio = pickle.load(a_file)

Meses = ['enero','febrero','marzo','abril','mayo']
import plotly.graph_objects as go


fig = go.Figure(data=go.Scatter(x=Meses, y=dictio['Sedes']['IPS SURA ALMACENTRO']['Codigos']['CONSULTA MEDICO GENERAL']['Oportunidad']))
fig.write_image("oportunidad2.png")
fig.show()




#%% Prueba envios de correo electronico
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

msg = MIMEMultipart()
mensajex = "Este es un mensaje con header {}".format(mensaje)

password = ""
msg['From'] = "wagamez@sura.com.co"
msg['To'] = "wilmer070@gmail.com"
msg['Subject'] = "Prueba para la IPS Almacentro"


msg.attach(MIMEText(mensajex, 'plain'))

#create server
#server = smtplib.SMTP('smtp.gmail.com: 587')
server = smtplib.SMTP('smtp.office365.com: 587')
server.starttls()

# Login Credentials for sending the mail
server.login(msg['From'], password)

# send the message via the server.
server.sendmail(msg['From'], msg['To'], msg.as_string())


print("successfully sent email to %s:" % (msg['To']))

server.quit()





#%% Guardar diccionario e importar datos
# Guardar en formato pickle
import pickle
f = open(r"D:\Usuarios\wilmagju\OneDrive - Seguros Suramericana, S.A\25.Modelo completo\dictio.pkl","wb")
pickle.dump(dictio,f)
f.close()
# Leer formato pickle
a_file = open("dictio.pkl", "rb")
dictio2 = pickle.load(a_file)
print(dictio2)



import json
# Guardar en formato json
a_file = open("person2.json", "w")
json.dump(dictio, a_file)
a_file.close()





 









 

     

     
     
     




        
    

    
    
               
 






#%%


#%% Utilizando SQalchemy

import pandas as pd
import re
import io
from unicodedata import normalize
from sqlalchemy import create_engine

user = 'wilmagju'
pasw= 'Magangue123#' #getpass.getpass(prompt='Ingrese contraseña')
host = 'teradata2.suranet.com'
# connect
td_engine = create_engine('teradata://'+ user +':' + pasw + '@'+ host + ':22/')


#conn = td_engine.connect()
conn = td_engine.raw_connection()
cur = conn.cursor()
#sql = 'select * from user_stage.wilmagju_resultado_hada'
result = td_engine.execute(sql)
Data = pd.DataFrame(data)

len(Data)






