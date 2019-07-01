# -*- coding: utf-8 -*-
"""
Created on Thu Jun 20 23:49:53 2019

@author: ALEXIS GUAMAN
"""

from tkinter import *
import tkinter.ttk as ttk
from tkinter import messagebox
from tkinter import filedialog


import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import style
style.use('ggplot')
import numpy as np


from pandas import ExcelWriter
from xlrd import open_workbook
from xlutils.copy import copy


import calendar
from datetime import date

from time import time


from sklearn import preprocessing, cross_validation, svm
from sklearn.linear_model import LinearRegression, LogisticRegression
import math

import numpy
import matplotlib.pyplot as plt, mpld3
import pandas

from keras.models import Sequential
from keras.layers import Dense
#==============================================================================
#========================= COMPRUEBA AÑO ACTUAL ===============================
#======================= DIRECCION BASE DE DATOS ==============================
import datetime
# Obtiene el año actual
anio_actual = (datetime.datetime.now()).year
# Abre archivo de texto
archivo = open("anios.txt", "r")
# Lee archivo de texto y lo almacena como lista
anios_disponibles = archivo.readlines()
# Elimina primer elemento de la lista (la direccion de la base de datos) y la almacena en la otra variable
# Éstas son las direcciones de la base de datos y de los resultados
direccion_base_datos = (anios_disponibles.pop(0))[:-1] # Elimina el último carácter (\n)
direccion_resultados = (anios_disponibles.pop(0))[:-1] # Elimina el último carácter (\n)
# Cierra archivo de texto
archivo.close()

# Si el año actual no está en lista, lo agrega (se compara con el último valor)      
if anio_actual != int(anios_disponibles[-1]):
    # Abre archivo de texto como escritura
    archivo = open("anios.txt", "a")
    # Se guarda el nuevo año en el archivo de texto
    archivo.write('\n'+str(anio_actual))
    # Se cierra el archivo de texto
    archivo.close()
    # Se abre nuevamente el archivo de texto como lectura
    archivo = open("anios.txt", "r")
    # SE ALMACENA LISTA DE AÑOS EN ESTA VARIABLE
    anios_disponibles = archivo.readlines()
    anios_disponibles.pop(0) # Se elimina el primer dato de la lista (dirección de la base de datos)
    anios_disponibles.pop(0) # Se elimina el segundo dato de la lista (dirección destino resultados)
    # Se cierra el archivo de texto
    archivo.close()

#anios_disponibles.pop(0)
anios_prov = []
for i in range (len(anios_disponibles)):
    anios_prov.append(str(int(anios_disponibles[i])))
anios_disponibles = anios_prov
del anios_prov

from datetime import datetime,timedelta
#==============================================================================
#==============================================================================        
def proyeccion_mlp():
    # -*- coding: utf-8 -*-
    """
    Created on Tue Apr  9 12:03:53 2019
    ESCUELA POLITÉCNICA NACIONAL
    AUTOR: ALEXIS GUAMAN FIGUEROA
    8_MÓDULO DE PROYECCIÓN CON ALGORITMO MLP
    """
    
#    from time import time
    def test():
        #===================== MODELO DE PERCEPTRÓN MULTICAPA =========================
        
        #===> PREDICCIÓN DE CARGA (t+1, dado t), Exactitud del 97%
        
#        import numpy
#        import matplotlib.pyplot as plt, mpld3
#        import pandas
#        import math
#        from keras.models import Sequential
#        from keras.layers import Dense
#        
#        
#        import pandas as pd
#        from pandas import ExcelWriter
#        
#        from xlrd import open_workbook
#        from xlutils.copy import copy
#        
#        from datetime import date
#        from datetime import datetime,timedelta
        
        #=============================== VAR GLOBALES =================================
        global pos_Metodologia, pos_Realizado, pos_Comparacion
        global pos_anio_p, pos_mes_p, pos_dia_p
        global pos_anio_c, pos_mes_c, pos_dia_c
        #******************************************************************************
        global po_Metodologia, po_Realizado, po_Comparacion
        global po_anio_p, po_mes_p, po_dia_p
        global po_anio_c, po_mes_c, po_dia_c
        global direccion_base_datos
        #==============================================================================
        
        doc_Proy       = direccion_base_datos+'\\HIS_POT_' + po_anio_p + '.xls'
        doc_Proy_extra = direccion_base_datos+'\\HIS_POT_' + str(int(po_anio_p)-1) + '.xls'
        doc_Comp       = direccion_base_datos+'\\HIS_POT_' + po_anio_c + '.xls'
        doc_Comp_extra = direccion_base_datos+'\\HIS_POT_' + str(int(po_anio_c)-1) + '.xls'
        
        # Arreglo de semilla para generar números pseaudo aleatorios para reproducibilidad
        #numpy.random.seed(7)
        
        # Genera el DataFrame a partir del archivo de datos
        df_Proy       = pd.read_excel(doc_Proy      , sheetname='Hoja1',  header=None)
        df_Proy_extra = pd.read_excel(doc_Proy_extra, sheetname='Hoja1',  header=None)
        df_Comp       = pd.read_excel(doc_Comp      , sheetname='Hoja1',  header=None)
        df_Comp_extra = pd.read_excel(doc_Comp_extra, sheetname='Hoja1',  header=None)
        
        # Numero total de datos a proyectar, se toma 24hrs x 7 días
        forecast_out = 24*7
        #******************************************************************************
        #******************************************************************************
        #******************************************************************************
        #******************************************************************************
        #******************************************************************************
        # Lista de datos de históricos de proyección
        amba_1 = (df_Proy_extra.iloc[:,4].values.tolist())
        toto_1 = (df_Proy_extra.iloc[:,5].values.tolist())
        puyo_1 = (df_Proy_extra.iloc[:,6].values.tolist())
        tena_1 = (df_Proy_extra.iloc[:,7].values.tolist())
        bani_1 = (df_Proy_extra.iloc[:,8].values.tolist())
        
        amba_2 = (df_Proy.iloc[:,4].values.tolist())
        toto_2 = (df_Proy.iloc[:,5].values.tolist())
        puyo_2 = (df_Proy.iloc[:,6].values.tolist())
        tena_2 = (df_Proy.iloc[:,7].values.tolist())
        bani_2 = (df_Proy.iloc[:,8].values.tolist())
        
        # Elimino la primera fila de cada DF (TEXTOS TÍTULOS)
        amba_1.pop(0)
        toto_1.pop(0)
        puyo_1.pop(0)
        tena_1.pop(0)
        bani_1.pop(0)
        
        amba_2.pop(0)
        toto_2.pop(0)
        puyo_2.pop(0)
        tena_2.pop(0)
        bani_2.pop(0)
        
        # Crea nueva lsita con los datos de los dos años el actual y el anterior
        amba_p = amba_1 + amba_2
        toto_p = toto_1 + toto_2
        puyo_p = puyo_1 + puyo_2
        tena_p = tena_1 + tena_2
        bani_p = bani_1 + bani_2
        # Transforma la lista a un DF
        amba_p = pd.DataFrame(amba_p)
        toto_p = pd.DataFrame(toto_p)
        puyo_p = pd.DataFrame(puyo_p)
        tena_p = pd.DataFrame(tena_p)
        bani_p = pd.DataFrame(bani_p)
        
        #==============================================================================
        #==============================================================================
        #===>UBICACIONES GENERALES de las posiciones para la proyección
        
        #===> Variables internas
        # FECHA SELECCIONADA
        fehca_selec = datetime(int(po_anio_p),(pos_mes_p+1),int(po_dia_p))
        # DÍAS DE PROYECCIÓN
        dias_proy = timedelta(days = 5952/24 )   # 5952 * 24 = 248 días mas un día de salto
        
        # FECHA DE INICIO
        fecha_inicio = (fehca_selec + timedelta(days = 7)) - dias_proy
        # FECHA DE FIN
        fecha_fin = fehca_selec + timedelta(days = 7) # 24 horas de un día
        # FECHAS DE INICIO DEl AÑO INICIAL
        enero_1 = datetime(fecha_inicio.year,1,1)
        
        # POSICIÓN INICIAL & FINAL
        val_ini = (abs(fecha_inicio - enero_1).days) * 24 # Posición del valor inicial 5952 DATOS
        val_fin = 3+(abs(fecha_fin    - enero_1).days) * 24 -169# Posición del valor final 5952 + 24 = 5976 DATOS
        
        # ELIMINA LOS ÚLTIMOS VALORES DEL DF
        amba_p = amba_p[:val_fin]
        toto_p = toto_p[:val_fin]
        puyo_p = puyo_p[:val_fin]
        tena_p = tena_p[:val_fin]
        bani_p = bani_p[:val_fin]
        # ELIMINA LOS PRIMEROS VALORES DEL DF ** Queda un DF con 5952 datos
        amba_p = amba_p[val_ini:]
        toto_p = toto_p[val_ini:]
        puyo_p = puyo_p[val_ini:]
        tena_p = tena_p[val_ini:]
        bani_p = bani_p[val_ini:]
        # REINICIA EL INDEX DE CADA DF   ***** matriz datos de entrenamiento
        amba_p = amba_p.reset_index(drop=True)
        toto_p = toto_p.reset_index(drop=True)
        puyo_p = puyo_p.reset_index(drop=True)
        tena_p = tena_p.reset_index(drop=True)
        bani_p = bani_p.reset_index(drop=True)
        
        #******************************************************************************
        # Hatas aquí, amba_p y las demás son los DF utilizados para la proyección
        # tienen 5952 datos cada DF
        #******************************************************************************
        #******************************************************************************
        #******************************************************************************
        #******************************************************************************
        
        # Lista de datos de históricos de COMPARACIÓN
        amba_1 = (df_Comp_extra.iloc[:,4].values.tolist())
        toto_1 = (df_Comp_extra.iloc[:,5].values.tolist())
        puyo_1 = (df_Comp_extra.iloc[:,6].values.tolist())
        tena_1 = (df_Comp_extra.iloc[:,7].values.tolist())
        bani_1 = (df_Comp_extra.iloc[:,8].values.tolist())
        
        amba_2 = (df_Comp.iloc[:,4].values.tolist())
        toto_2 = (df_Comp.iloc[:,5].values.tolist())
        puyo_2 = (df_Comp.iloc[:,6].values.tolist())
        tena_2 = (df_Comp.iloc[:,7].values.tolist())
        bani_2 = (df_Comp.iloc[:,8].values.tolist())
        
        # Elimino la primera fila de cada DF (TEXTOS TÍTULOS)
        amba_1.pop(0)
        toto_1.pop(0)
        puyo_1.pop(0)
        tena_1.pop(0)
        bani_1.pop(0)
        
        amba_2.pop(0)
        toto_2.pop(0)
        puyo_2.pop(0)
        tena_2.pop(0)
        bani_2.pop(0)
        # Une las dos listas en una sola
        amba_c = amba_1 + amba_2
        toto_c = toto_1 + toto_2
        puyo_c = puyo_1 + puyo_2
        tena_c = tena_1 + tena_2
        bani_c = bani_1 + bani_2
        # Transforma la lista al tipo DF
        #amba_c = pd.DataFrame(amba_c)
        #toto_c = pd.DataFrame(toto_c)
        #puyo_c = pd.DataFrame(puyo_c)
        #tena_c = pd.DataFrame(tena_c)
        #bani_c = pd.DataFrame(bani_c)
        #==============================================================================
        #==============================================================================
        #===>UBICACIONES GENERALES de las posiciones para la COMPARACIÓN
        
        
        #===> Variables internas
        # FECHA SELECCIONADA
        fehca_selec = datetime(int(po_anio_c),(pos_mes_c+1),int(po_dia_c))
        # DÍAS DE PROYECCIÓN
        dias_proy_c = timedelta(days = 7)
        
        # FECHA DE INICIO
        fecha_inicio = (fehca_selec - dias_proy_c)
        # FECHA DE FIN
        fecha_fin = fehca_selec
        # FECHAS DE INICIO DEl AÑO INICIAL
        enero_1 = datetime((int(po_anio_c)-1),1,1)
        
        # POSICIÓN INICIAL & FINAL
        val_ini = (abs(fecha_inicio - enero_1).days) * 24 # Posición del valor inicial 5952 DATOS
        val_fin = (abs(fecha_fin    - enero_1).days) * 24 # Posición del valor final 5952 + 24 = 5976 DATOS
        
        # ELIMINA LOS ÚLTIMOS VALORES DEL DF
        amba_c = amba_c[:val_fin]
        toto_c = toto_c[:val_fin]
        puyo_c = puyo_c[:val_fin]
        tena_c = tena_c[:val_fin]
        bani_c = bani_c[:val_fin]
        # ELIMINA LOS PRIMEROS VALORES DEL DF ** Queda un DF con 5952 datos
        amba_c = amba_c[val_ini:]
        toto_c = toto_c[val_ini:]
        puyo_c = puyo_c[val_ini:]
        tena_c = tena_c[val_ini:]
        bani_c = bani_c[val_ini:]
        # REINICIA EL INDEX DE CADA DF   ***** matriz datos de comparación
        CompAmba = amba_c
        CompToto = toto_c
        CompPuyo = puyo_c
        CompTena = tena_c
        CompBani = bani_c
        
        
        # Datos de comparación
        #CompAmba = (amba_p[0][-forecast_out:]).reset_index(drop=True)
        #CompToto = (toto_p[0][-forecast_out:]).reset_index(drop=True)
        #CompPuyo = (puyo_p[0][-forecast_out:]).reset_index(drop=True)
        #CompTena = (tena_p[0][-forecast_out:]).reset_index(drop=True)
        #CompBani = (bani_p[0][-forecast_out:]).reset_index(drop=True)
        
        
        #******************************************************************************
        #******************************************************************************
        #******************************************************************************
        #******************************************************************************
        #******************************************************************************
        
        #==============================================================================
        #==============================================================================
        
        # Obtiene los valores del DF
        dataset_amba = amba_p.values
        dataset_toto = toto_p.values
        dataset_puyo = puyo_p.values
        dataset_tena = tena_p.values
        dataset_bani = bani_p.values
        
        # Transforma los valores del DS al tipo float32
        dataset_amba = dataset_amba.astype('float32')
        dataset_toto = dataset_toto.astype('float32')
        dataset_puyo = dataset_puyo.astype('float32')
        dataset_tena = dataset_tena.astype('float32')
        dataset_bani = dataset_bani.astype('float32')
        
        
        # Obtiene la media del conjunto de datos del DF
        avgv_amba = amba_p.mean()
        avgv_toto = toto_p.mean()
        avgv_puyo = puyo_p.mean()
        avgv_tena = tena_p.mean()
        avgv_bani = bani_p.mean()
        
        # Se dividide en datos de entrenamiento y de prueba:
        # Tamaño del conjunto de datos de entrenamiento
        #train_size = int(len(dataset) * 0.67) # Usa el 67% del total de los datos
        train_size = int(len(dataset_amba) - 179) #168
        # Tamaño del conjunto de datos de prueba
        test_size = len(dataset_amba) - train_size # Resulta el 33% de datos restantes
        
        # Se genera las columnas de datos de entrenamiento y de prueba
        train_amba, test_amba = dataset_amba[0:train_size,:], dataset_amba[train_size:len(dataset_amba),:]
        train_toto, test_toto = dataset_toto[0:train_size,:], dataset_toto[train_size:len(dataset_toto),:]
        train_puyo, test_puyo = dataset_puyo[0:train_size,:], dataset_puyo[train_size:len(dataset_puyo),:]
        train_tena, test_tena = dataset_tena[0:train_size,:], dataset_tena[train_size:len(dataset_tena),:]
        train_bani, test_bani = dataset_bani[0:train_size,:], dataset_bani[train_size:len(dataset_bani),:]
        #==============================================================================
        
        # Convierte una matriz de valores en una matriz de conjunto de datos
        def create_dataset_amba(dataset_amba, look_back=1):
                dataX_amba, dataY_amba = [], []
                for i in range(len(dataset_amba)-look_back-1):
                        a_amba = dataset_amba[i:(i+look_back), 0]
                        dataX_amba.append(a_amba)
                        dataY_amba.append(dataset_amba[i + look_back, 0])
                return numpy.array(dataX_amba), numpy.array(dataY_amba)
        
        # reshape into X=t and Y=t+1
        look_back = 10
        
        trainX_amba, trainY_amba = create_dataset_amba(train_amba, look_back)
        testX_amba, testY_amba = create_dataset_amba(test_amba, look_back)
        #==============================================================================
        
        # Convierte una matriz de valores en una matriz de conjunto de datos
        def create_dataset_toto(dataset_toto, look_back=1):
                dataX_toto, dataY_toto = [], []
                for i in range(len(dataset_toto)-look_back-1):
                        a_toto = dataset_toto[i:(i+look_back), 0]
                        dataX_toto.append(a_toto)
                        dataY_toto.append(dataset_toto[i + look_back, 0])
                return numpy.array(dataX_toto), numpy.array(dataY_toto)
        
        # reshape into X=t and Y=t+1
        look_back = 10
        
        trainX_toto, trainY_toto = create_dataset_toto(train_toto, look_back)
        testX_toto, testY_toto = create_dataset_toto(test_toto, look_back)
        #==============================================================================
        
        # Convierte una matriz de valores en una matriz de conjunto de datos
        def create_dataset_puyo(dataset_puyo, look_back=1):
                dataX_puyo, dataY_puyo = [], []
                for i in range(len(dataset_puyo)-look_back-1):
                        a_puyo = dataset_puyo[i:(i+look_back), 0]
                        dataX_puyo.append(a_puyo)
                        dataY_puyo.append(dataset_puyo[i + look_back, 0])
                return numpy.array(dataX_puyo), numpy.array(dataY_puyo)
        
        # reshape into X=t and Y=t+1
        look_back = 10
        
        trainX_puyo, trainY_puyo = create_dataset_puyo(train_puyo, look_back)
        testX_puyo, testY_puyo = create_dataset_puyo(test_puyo, look_back)
        #==============================================================================
        
        # Convierte una matriz de valores en una matriz de conjunto de datos
        def create_dataset_tena(dataset_tena, look_back=1):
                dataX_tena, dataY_tena = [], []
                for i in range(len(dataset_tena)-look_back-1):
                        a_tena = dataset_tena[i:(i+look_back), 0]
                        dataX_tena.append(a_tena)
                        dataY_tena.append(dataset_tena[i + look_back, 0])
                return numpy.array(dataX_tena), numpy.array(dataY_tena)
        
        # reshape into X=t and Y=t+1
        look_back = 10
        
        trainX_tena, trainY_tena = create_dataset_tena(train_tena, look_back)
        testX_tena, testY_tena = create_dataset_tena(test_tena, look_back)
        #==============================================================================
        
        # Convierte una matriz de valores en una matriz de conjunto de datos
        def create_dataset_bani(dataset_bani, look_back=1):
                dataX_bani, dataY_bani = [], []
                for i in range(len(dataset_bani)-look_back-1):
                        a_bani = dataset_bani[i:(i+look_back), 0]
                        dataX_bani.append(a_bani)
                        dataY_bani.append(dataset_bani[i + look_back, 0])
                return numpy.array(dataX_bani), numpy.array(dataY_bani)
        
        # reshape into X=t and Y=t+1
        look_back = 10
        
        trainX_bani, trainY_bani = create_dataset_bani(train_bani, look_back)
        testX_bani, testY_bani = create_dataset_bani(test_bani, look_back)
        #==============================================================================
        
        
        # Crea y ajusta el modelo de perceptron multicapa
        model = Sequential()
        model.add(Dense(8, input_dim=look_back, activation='relu'))
        model.add(Dense(1))
        model.compile(loss='mean_squared_error', optimizer='adam')
        
        #===> VARIAR EL NÚMERO DE CAPAS PARA MEJORAR PRESICIÓN, LAYER CAPAS
        #/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
        #/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
        #/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
        epoca = 10
        #/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
        #/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
        #/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
        print('=======>  Iteraciones para la S/E AMBATO: ')
        model.fit(trainX_amba, trainY_amba, epochs=epoca, batch_size=2, verbose=2)
        print('=======>  Iteraciones para la S/E TOTORAS: ')
        model.fit(trainX_toto, trainY_toto, epochs=epoca, batch_size=2, verbose=2)
        print('=======>  Iteraciones para la S/E PUYO: ')
        model.fit(trainX_puyo, trainY_puyo, epochs=epoca, batch_size=2, verbose=2)
        print('=======>  Iteraciones para la S/E TENA: ')
        model.fit(trainX_tena, trainY_tena, epochs=epoca, batch_size=2, verbose=2)
        print('=======>  Iteraciones para la S/E BAÑOS: ')
        model.fit(trainX_bani, trainY_bani, epochs=epoca, batch_size=2, verbose=2)
        
        # Estimar el rendimiento del modelo
        trainScore_amba = model.evaluate(trainX_amba, trainY_amba, verbose=0)
        trainScore_toto = model.evaluate(trainX_toto, trainY_toto, verbose=0)
        trainScore_puyo = model.evaluate(trainX_puyo, trainY_puyo, verbose=0)
        trainScore_tena = model.evaluate(trainX_tena, trainY_tena, verbose=0)
        trainScore_bani = model.evaluate(trainX_bani, trainY_bani, verbose=0)
        #print(' ')
        #print(' ')
        #print('Train Score: %.2f MSE (%.2f RMSE)' % (trainScore, math.sqrt(trainScore)))
        testScore_amba = model.evaluate(testX_amba, testY_amba, verbose=0)
        testScore_toto = model.evaluate(testX_toto, testY_toto, verbose=0)
        testScore_puyo = model.evaluate(testX_puyo, testY_puyo, verbose=0)
        testScore_tena = model.evaluate(testX_tena, testY_tena, verbose=0)
        testScore_bani = model.evaluate(testX_bani, testY_bani, verbose=0)
        #print('Test Score: %.2f MSE (%.2f RMSE)' % (testScore, math.sqrt(testScore)))
        print(' ')
        print(' ')
        newV_amba = avgv_amba + math.sqrt(trainScore_amba)
        newV_toto = avgv_toto + math.sqrt(trainScore_toto)
        newV_puyo = avgv_puyo + math.sqrt(trainScore_puyo)
        newV_tena = avgv_tena + math.sqrt(trainScore_tena)
        newV_bani = avgv_bani + math.sqrt(trainScore_bani)
        
        acc_amba = 100-(((newV_amba - avgv_amba)/(avgv_amba))*100)
        acc_toto = 100-(((newV_toto - avgv_toto)/(avgv_toto))*100)
        acc_puyo = 100-(((newV_puyo - avgv_puyo)/(avgv_puyo))*100)
        acc_tena = 100-(((newV_tena - avgv_tena)/(avgv_tena))*100)
        acc_bani = 100-(((newV_bani - avgv_bani)/(avgv_bani))*100)
        #print(avgv)
        #print('Mean Absolute Error of',newV)
        print('* Porcentaje de Exactitud de la predicción por MLP : ')
        print('=> S/E AMBATO  :' ,acc_amba[0])
        print('=> S/E TOTORAS :' ,acc_toto[0])
        print('=> S/E PUYO    :' ,acc_puyo[0])
        print('=> S/E TENA    :' ,acc_tena[0])
        print('=> S/E BAÑOS   :' ,acc_bani[0])
        
        # Generar predicciones para el entrenamiento.
        #trainPredict_amba = model.predict(trainX_amba)
        ProyAmba  = model.predict(testX_amba)
        
        #trainPredict_toto = model.predict(trainX_toto)
        ProyToto  = model.predict(testX_toto)
        
        #trainPredict_puyo = model.predict(trainX_puyo)
        ProyPuyo  = model.predict(testX_puyo)
        
        #trainPredict_tena = model.predict(trainX_tena)
        ProyTena  = model.predict(testX_tena)
        
        #trainPredict_bani = model.predict(trainX_bani)
        ProyBani  = model.predict(testX_bani)
        
        
        #==============================================================================
        #================================ GRÁFICAS ====================================
        #==============================================================================
        if po_Comparacion == 'SI':
        #===> Creamos una lista tipo entero para relacionar con las etiquetas
            can_datos = []
            for i in range(7*24):
                if i%2!=1:
                    can_datos.append(i)     
        #===> Creamos una lista con los números de horas del día seleccionado (etiquetas)
            horas_dia = []
            horas_str = []
            for i in range (7):
                for i in range (1,25):
                    if i%2!=0:
                        horas_dia.append(i)
            for i in range (len(horas_dia)):
                horas_str.append(str(horas_dia[i]))
                
        #===> Tamaño de la ventana de la gráfica
            plt.subplots(figsize=(15, 8))
            
        #===> Título general superior
            plt.suptitle(u' PROYECCIÓN SEMANAL DE CARGA\n MULTILAYER PERCEPTRON ',fontsize=14, fontweight='bold') 
            
            plt.subplot(5,1,1)
            plt.plot(CompAmba,'blue', label = 'Comparación')
            plt.plot(ProyAmba,'#DCD037', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E AMBATO\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,2)
            plt.plot(CompToto,'blue', label = 'Comparación')
            plt.plot(ProyToto,'#CD336F', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E TOTORAS\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,3)  
            plt.plot(CompPuyo,'blue', label = 'Comparación')
            plt.plot(ProyPuyo,'#349A9D', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E PUYO\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,4)  
            plt.plot(CompTena,'blue', label = 'Comparación')
            plt.plot(ProyTena,'#CC8634', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E TENA\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,5)  
            plt.plot(CompBani,'blue', label = 'Comparación')
            plt.plot(ProyBani,'#4ACB71', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E BAÑOS\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
        
        #=========> Fechas de salida:
        sem_ini = datetime(int(po_anio_p),(pos_mes_p+1),int(po_dia_p))
        sem_fin = sem_ini + timedelta(days = 6)
        
        #==============================================================================
        #======================== RUTINA GENERA EXCEL =================================
        #==============================================================================
            
        ProyAmba = ProyAmba.astype('float64')
        ProyToto = ProyToto.astype('float64')
        ProyPuyo = ProyPuyo.astype('float64')
        ProyTena = ProyTena.astype('float64')
        ProyBani = ProyBani.astype('float64')
        
        df = pd.DataFrame([' '])
        
        writer = ExcelWriter('04_PROYECCION_MLP.xls')
        df.to_excel(writer, 'Salida_Proyección_CECON', index=False)
        df.to_excel(writer, 'Salida_Comparación_CECON', index=False)
        df.to_excel(writer, 'Salida_ERRORES_CECON', index=False)
        writer.save()
        
        #abre el archivo de excel plantilla
        rb = open_workbook('04_PROYECCION_MLP.xls')
        #crea una copia del archivo plantilla
        wb = copy(rb)
        #se ingresa a la hoja 1 de la copia del archivo excel
        ws = wb.get_sheet(0)
        ws_c = wb.get_sheet(1)
        ws_e = wb.get_sheet(2)
        #===========================Hoja 1 Proyección =================================
        #ws.write(0,0,'MES')
        #ws.write(0,1,'DÍA')
        #ws.write(0,2,'#')
        ws.write(0,0,'METODOLOGÍA PERCEPTRÓN MULTI - CAPA  (M.L.P.)')
        ws.write(1,0,'PROYECCIÓN SEMANAL DE CARGA')
        ws.write(2,0,'Semana de Proyección: del '+str(sem_ini.day)+'/'
                 +str(sem_ini.month) +'/'+str(sem_ini.year)+'\t al '+
                 str(sem_fin.day)+'/'+str(sem_fin.month)+'/'+str(sem_fin.year))
        
        ws.write(3,0,'Realizado por: '+str(po_Realizado))
        
        # Define lista de títulos
        titulos = ['HORA','S/E AMBATO','S/E TOTORAS','S/E PUYO','S/E TENA','S/E BAÑOS']
        
        for i in range(len(titulos)):
            ws.write(9,i,titulos[i])
            
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws.write(Aux,0,j+1)
                ws.write(Aux,1,ProyAmba[j + Aux2][0])
                ws.write(Aux,2,ProyToto[j + Aux2][0])
                ws.write(Aux,3,ProyPuyo[j + Aux2][0])
                ws.write(Aux,4,ProyTena[j + Aux2][0])
                ws.write(Aux,5,ProyBani[j + Aux2][0])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
        #===========================Hoja 2 Comparación ================================
        
        #===> Fecha
        ws_c.write(0,0,'Semana de Proyección: del '+str(sem_ini.day)+'/'
                   +str(sem_ini.month) +'/'+str(sem_ini.year)+'\t al '+
                   str(sem_fin.day)+'/'+str(sem_fin.month)+'/'+str(sem_fin.year))
        ws_c.write(2,0,'PROYECCIÓN SEMANAL')
        
        for i in range(len(titulos)):
            ws_c.write(9,i,titulos[i])
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,0,j+1)
                ws_c.write(Aux,1,ProyAmba[j + Aux2][0])
                ws_c.write(Aux,2,ProyToto[j + Aux2][0])
                ws_c.write(Aux,3,ProyPuyo[j + Aux2][0])
                ws_c.write(Aux,4,ProyTena[j + Aux2][0])
                ws_c.write(Aux,5,ProyBani[j + Aux2][0])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
            
        #ws_c.write(0,6,po_dia_c+'/'+po_mes_c+'/'+po_anio_c)
        ws_c.write(2,6,'SEMANA DE COMPARACIÓN')
        
        for i in range(len(titulos)):
            ws_c.write(9,i+6,titulos[i])
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,6,j+1)
                ws_c.write(Aux,7,CompAmba[j + Aux2])
                ws_c.write(Aux,8,CompToto[j + Aux2])
                ws_c.write(Aux,9,CompPuyo[j + Aux2])
                ws_c.write(Aux,10,CompTena[j + Aux2])
                ws_c.write(Aux,11,CompBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
            
            
        # ==========================   Cálculo de errores    ==========================
            
        ws_c.write(2,12,'CÁLCULO DE ERRORES')
        ws_c.write(3,12,'PORCENTAJE DE ERROR MEDIO ABSOLUTO (PEMA)')
        
        for i in range(len(titulos)):
            ws_c.write(9,i+12,titulos[i])
        
        
        errAmba = []
        errToto = []
        errPuyo = []
        errTena = []
        errBani = []
        
        for i in range(len(ProyAmba)):
            errAmba.append((abs( ProyAmba[i][0] - CompAmba[i] ) / CompAmba[i] )*100)
            errToto.append((abs( ProyToto[i][0] - CompToto[i] ) / CompToto[i] )*100)
            errPuyo.append((abs( ProyPuyo[i][0] - CompPuyo[i] ) / CompPuyo[i] )*100)
            errTena.append((abs( ProyTena[i][0] - CompTena[i] ) / CompTena[i] )*100)
            errBani.append((abs( ProyBani[i][0] - CompBani[i] ) / CompBani[i] )*100)
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,12,j+1)
                ws_c.write(Aux,13,errAmba[j + Aux2])
                ws_c.write(Aux,14,errToto[j + Aux2])
                ws_c.write(Aux,15,errPuyo[j + Aux2])
                ws_c.write(Aux,16,errTena[j + Aux2])
                ws_c.write(Aux,17,errBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
        # Suma de los valores de las listas
        SumAmba = 0
        SumToto = 0
        SumPuyo = 0
        SumTena = 0
        SumBani = 0
        for i in range (len(ProyAmba)):
            SumAmba = errAmba[i] + SumAmba
            SumToto = errToto[i] + SumToto
            SumPuyo = errPuyo[i] + SumPuyo
            SumTena = errTena[i] + SumTena
            SumBani = errBani[i] + SumBani
        # Almaceno en una lista los resultados de las sumas
        Sumas = [SumAmba, SumToto, SumPuyo, SumTena, SumBani]
        
        # Imprime los resultados en la fila correspondiente
        ws_c.write(Aux+1,11,'SUMATORIA TOTAL')
        for i in range (len(Sumas)):
             ws_c.write(Aux+1,i+13,Sumas[i])     
        
        # Cálculo de los promedios de las sumas
        ws_c.write(Aux+2,11,'ERROR MEDIO ABSOLUTO')
        for i in range(len(Sumas)):
            ws_c.write(Aux+2,i+13,(Sumas[i]/len(errAmba)))
        
        # Cálculo de la exactitud de la proyección
        ws_c.write(Aux+3,11,'EXACTITUD DE LA PROYECCIÓN')
        for i in range(len(Sumas)):
            ws_c.write(Aux+3,i+13,(100-(Sumas[i]/len(errAmba))))
        
        
        for i in range(len(titulos)-1):
            ws_c.write(Aux,i+13,titulos[i+1])
        
        # ==========================   SOLO   errores    ==========================
        
        # ==================================> cálculo de et = Comp - Proy
        for i in range(len(titulos)-1):
            ws_e.write(9,i,titulos[i+1])
        ws_e.write(7,0,'et = Comp - Proy')
        
        et_Amba = []
        et_Toto = []
        et_Puyo = []
        et_Tena = []
        et_Bani = []
        
        for i in range (len(ProyAmba)):
            et_Amba.append(CompAmba[i] - ProyAmba[i][0])
            et_Toto.append(CompToto[i] - ProyToto[i][0])
            et_Puyo.append(CompPuyo[i] - ProyPuyo[i][0])
            et_Tena.append(CompTena[i] - ProyTena[i][0])
            et_Bani.append(CompBani[i] - ProyBani[i][0])
            
            ws_e.write(i+10,0, (et_Amba[i]))
            ws_e.write(i+10,1, (et_Toto[i]))
            ws_e.write(i+10,2, (et_Puyo[i]))
            ws_e.write(i+10,3, (et_Tena[i]))
            ws_e.write(i+10,4, (et_Bani[i]))
        
        # ==================================> cálculo de abs(et) = abs(Comp - Proy)
        for i in range(len(titulos)-1):
            ws_e.write(9,i+6,titulos[i+1])
        ws_e.write(7,6,'abs(et) = abs(Comp - Proy)')
        
        abs_et_Amba = []
        abs_et_Toto = []
        abs_et_Puyo = []
        abs_et_Tena = []
        abs_et_Bani = []
        
        for i in range (len(ProyAmba)):
            abs_et_Amba.append(abs(et_Amba[i]))
            abs_et_Toto.append(abs(et_Toto[i]))
            abs_et_Puyo.append(abs(et_Puyo[i]))
            abs_et_Tena.append(abs(et_Tena[i]))
            abs_et_Bani.append(abs(et_Bani[i]))
            
            ws_e.write(i+10,6,  (abs_et_Amba[i]))
            ws_e.write(i+10,7,  (abs_et_Toto[i]))
            ws_e.write(i+10,8,  (abs_et_Puyo[i]))
            ws_e.write(i+10,9,  (abs_et_Tena[i]))
            ws_e.write(i+10,10, (abs_et_Bani[i]))
            
        # ==================================> cálculo de et^2
        for i in range(len(titulos)-1):
            ws_e.write(9,i+12,titulos[i+1])
        ws_e.write(7,12,'et^2')
        
        et_Amba2 = []
        et_Toto2 = []
        et_Puyo2 = []
        et_Tena2 = []
        et_Bani2 = []
        
        for i in range (len(ProyAmba)):
            et_Amba2.append((et_Amba[i])**2)
            et_Toto2.append((et_Toto[i])**2)
            et_Puyo2.append((et_Puyo[i])**2)
            et_Tena2.append((et_Tena[i])**2)
            et_Bani2.append((et_Bani[i])**2)
            
            ws_e.write(i+10,12, (et_Amba2[i]))
            ws_e.write(i+10,13, (et_Toto2[i]))
            ws_e.write(i+10,14, (et_Puyo2[i]))
            ws_e.write(i+10,15, (et_Tena2[i]))
            ws_e.write(i+10,16, (et_Bani2[i]))
            
        # ==================================> cálculo de abs(et) / Comp
        for i in range(len(titulos)-1):
            ws_e.write(9,i+18,titulos[i+1])
        ws_e.write(7,18,'abs(et) / Comp')
        
        d1_Amba = []
        d1_Toto = []
        d1_Puyo = []
        d1_Tena = []
        d1_Bani = []
        
        for i in range (len(ProyAmba)):
            d1_Amba.append(abs_et_Amba[i] / CompAmba[i])
            d1_Toto.append(abs_et_Toto[i] / CompToto[i])
            d1_Puyo.append(abs_et_Puyo[i] / CompPuyo[i])
            d1_Tena.append(abs_et_Tena[i] / CompTena[i])
            d1_Bani.append(abs_et_Bani[i] / CompBani[i])
            
            ws_e.write(i+10,18, (d1_Amba[i]))
            ws_e.write(i+10,19, (d1_Toto[i]))
            ws_e.write(i+10,20, (d1_Puyo[i]))
            ws_e.write(i+10,21, (d1_Tena[i]))
            ws_e.write(i+10,22, (d1_Bani[i]))   
        
        # ==================================> cálculo de et / Comp
        for i in range(len(titulos)-1):
            ws_e.write(9,i+24,titulos[i+1])
        ws_e.write(7,24,'et / Comp')
        
        d2_Amba = []
        d2_Toto = []
        d2_Puyo = []
        d2_Tena = []
        d2_Bani = []
        
        for i in range (len(ProyAmba)):
            d2_Amba.append(et_Amba[i] / CompAmba[i])
            d2_Toto.append(et_Toto[i] / CompToto[i])
            d2_Puyo.append(et_Puyo[i] / CompPuyo[i])
            d2_Tena.append(et_Tena[i] / CompTena[i])
            d2_Bani.append(et_Bani[i] / CompBani[i])
            
            ws_e.write(i+10,24, (d2_Amba[i]))
            ws_e.write(i+10,25, (d2_Toto[i]))
            ws_e.write(i+10,26, (d2_Puyo[i]))
            ws_e.write(i+10,27, (d2_Tena[i]))
            ws_e.write(i+10,28, (d2_Bani[i]))   
            
        ws_e.write(0,0, 'INDICADORES') 
        # ==================================> Cálculo DAM
        DAM_Amba = 0
        DAM_Toto = 0
        DAM_Puyo = 0
        DAM_Tena = 0
        DAM_Bani = 0
        
        for i in range (len(ProyAmba)):
            DAM_Amba = abs_et_Amba[i] + DAM_Amba
            DAM_Toto = abs_et_Toto[i] + DAM_Toto
            DAM_Puyo = abs_et_Puyo[i] + DAM_Puyo
            DAM_Tena = abs_et_Tena[i] + DAM_Tena
            DAM_Bani = abs_et_Bani[i] + DAM_Bani
            
        DAM_Amba = DAM_Amba / (len(ProyAmba))
        DAM_Toto = DAM_Toto / (len(ProyAmba))
        DAM_Puyo = DAM_Puyo / (len(ProyAmba))
        DAM_Tena = DAM_Tena / (len(ProyAmba))
        DAM_Bani = DAM_Bani / (len(ProyAmba))
        
        ws_e.write(1,0, ('DAM'))
        DAM = [DAM_Amba,DAM_Toto,DAM_Puyo,DAM_Tena,DAM_Bani]
        for i in range(len(DAM)):
            ws_e.write(1,i+1,DAM[i])
        
        
        # ==================================> Cálculo EMC 
        EMC_Amba = 0
        EMC_Toto = 0
        EMC_Puyo = 0
        EMC_Tena = 0
        EMC_Bani = 0
        
        for i in range (len(ProyAmba)):
            EMC_Amba = et_Amba2[i] + EMC_Amba
            EMC_Toto = et_Toto2[i] + EMC_Toto
            EMC_Puyo = et_Puyo2[i] + EMC_Puyo
            EMC_Tena = et_Tena2[i] + EMC_Tena
            EMC_Bani = et_Bani2[i] + EMC_Bani
            
        EMC_Amba = EMC_Amba / (len(ProyAmba))
        EMC_Toto = EMC_Toto / (len(ProyAmba))
        EMC_Puyo = EMC_Puyo / (len(ProyAmba))
        EMC_Tena = EMC_Tena / (len(ProyAmba))
        EMC_Bani = EMC_Bani / (len(ProyAmba))
        
        ws_e.write(2,0, ('EMC'))
        EMC = [EMC_Amba,EMC_Toto,EMC_Puyo,EMC_Tena,EMC_Bani]
        
        for i in range(len(EMC)):
            ws_e.write(2,i+1,EMC[i])
            
        # ==================================> Cálculo PEMA
        PEMA_Amba = 0
        PEMA_Toto = 0
        PEMA_Puyo = 0
        PEMA_Tena = 0
        PEMA_Bani = 0
        
        for i in range (len(ProyAmba)):
            PEMA_Amba = (abs_et_Amba[i] / CompAmba [i]) + PEMA_Amba
            PEMA_Toto = (abs_et_Toto[i] / CompToto [i]) + PEMA_Toto
            PEMA_Puyo = (abs_et_Puyo[i] / CompPuyo [i]) + PEMA_Puyo
            PEMA_Tena = (abs_et_Tena[i] / CompTena [i]) + PEMA_Tena
            PEMA_Bani = (abs_et_Bani[i] / CompBani [i]) + PEMA_Bani
            
        PEMA_Amba = (PEMA_Amba / (len(ProyAmba))) *100
        PEMA_Toto = (PEMA_Toto / (len(ProyAmba))) *100
        PEMA_Puyo = (PEMA_Puyo / (len(ProyAmba))) *100
        PEMA_Tena = (PEMA_Tena / (len(ProyAmba))) *100
        PEMA_Bani = (PEMA_Bani / (len(ProyAmba))) *100
        
        ws_e.write(3,0, ('PEMA'))
        PEMA = [PEMA_Amba,PEMA_Toto,PEMA_Puyo,PEMA_Tena,PEMA_Bani]
        
        for i in range(len(PEMA)):
            ws_e.write(3,i+1,PEMA[i])
            
        # ==================================> Cálculo PME
        PME_Amba = 0
        PME_Toto = 0
        PME_Puyo = 0
        PME_Tena = 0
        PME_Bani = 0
        
        for i in range (len(ProyAmba)):
            PME_Amba = (et_Amba[i] / CompAmba [i]) + PME_Amba
            PME_Toto = (et_Toto[i] / CompToto [i]) + PME_Toto
            PME_Puyo = (et_Puyo[i] / CompPuyo [i]) + PME_Puyo
            PME_Tena = (et_Tena[i] / CompTena [i]) + PME_Tena
            PME_Bani = (et_Bani[i] / CompBani [i]) + PME_Bani
            
        PME_Amba = (PME_Amba / (len(ProyAmba))) *100
        PME_Toto = (PME_Toto / (len(ProyAmba))) *100
        PME_Puyo = (PME_Puyo / (len(ProyAmba))) *100
        PME_Tena = (PME_Tena / (len(ProyAmba))) *100
        PME_Bani = (PME_Bani / (len(ProyAmba))) *100
        
        ws_e.write(4,0, ('PME'))
        PME = [PME_Amba,PME_Toto,PME_Puyo,PME_Tena,PME_Bani]
        
        for i in range(len(PME)):
            ws_e.write(4,i+1,PME[i])
            
        # ==================================> Cálculo EXACTITUD DE LA PROYECCION
        ws_e.write(5,0,'EP')
        for i in range(len(PEMA)):
            ws_e.write(5,i+1,(100-PEMA[i]))
        
        
        for i in range(len(titulos)-1):
            ws_e.write(0,i+1,titulos[i+1])
        global direccion_resultados
        wb.save(direccion_resultados + '\\04_PROYECCION_MLP.xls')
        #==============================================================================
        #==============================================================================
        
        
        print(' ')
        print(' ')
        print('* La fecha seleccionada es:',int(po_anio_p),(pos_mes_p+1),int(po_dia_p))
        print(' ')
        print(' ')
        print(' ')
        print('*** MLP Completed...!!!')
        print(' ')
        print(' ')
        
    start_time = time()
    test()
    elapsed_time = time() - start_time
    
    if (elapsed_time >= 60):
        minutos = elapsed_time / 60
        segundos = elapsed_time - int (minutos * 60)
        print(' ')
        print(' ')
        print('/////////////////////////////////////////')
        print(' ')
        print('ALGORITMO UTILIZADO: MLP')
        print('Tiempo transcurrido: %.0f minutos '%minutos +'y %.2f segundos' %segundos)            
        print(' ') 
        print('/////////////////////////////////////////') 
    else:
        
        print(' ')
        print(' ')
        print('/////////////////////////////////////////')
        print(' ')
        print('ALGORITMO UTILIZADO: MLP')
        print("Tiempo transcurrido: %.10f segundos." % elapsed_time)            
        print(' ') 
        print('/////////////////////////////////////////')  

        


def proyeccion_svm():
    # -*- coding: utf-8 -*-
    """
    Created on Fri Apr  5 11:17:08 2019
    ESCUELA POLITÉCNICA NACIONAL
    AUTOR: ALEXIS GUAMAN FIGUEROA
    7_MÓDULO DE PROYECCIÓN CON ALGORITMO SVM
    """
#    from time import time
    def test():
        
        #========================= MODELO DE REGRESIÓN LINEAL =========================
        
        #===> LA MUESTRA USARÁ 5952 DATOS HISTÓRICOS
        #===> La variable forecast_out controla el tiempo de predicción a realizar
        
#        import pandas as pd
#        #import csv, math , datetime
#        #import time
#        import numpy as np
#        from sklearn import preprocessing, cross_validation, svm
#        from sklearn.linear_model import LinearRegression, LogisticRegression
#        import math
#        import matplotlib.pyplot as plt
#        from matplotlib import style
#        style.use('ggplot')
#        
#        
#        from pandas import ExcelWriter
#        
#        from xlrd import open_workbook
#        from xlutils.copy import copy
#        
#        from datetime import date
#        from datetime import datetime,timedelta
        
        #=============================== VAR GLOBALES =================================
        global pos_Metodologia, pos_Realizado, pos_Comparacion
        global pos_anio_p, pos_mes_p, pos_dia_p
        global pos_anio_c, pos_mes_c, pos_dia_c
        #******************************************************************************
        global po_Metodologia, po_Realizado, po_Comparacion
        global po_anio_p, po_mes_p, po_dia_p
        global po_anio_c, po_mes_c, po_dia_c
        global direccion_base_datos
        #==============================================================================
        
        doc_Proy       = direccion_base_datos+'\\HIS_POT_' + po_anio_p + '.xls'
        doc_Proy_extra = direccion_base_datos+'\\HIS_POT_' + str(int(po_anio_p)-1) + '.xls'
        doc_Comp       = direccion_base_datos+'\\HIS_POT_' + po_anio_c + '.xls'
        
        
        # Genera el DataFrame a partir del archivo de datos
        df_Proy       = pd.read_excel(doc_Proy      , sheetname='Hoja1',  header=None)
        df_Proy_extra = pd.read_excel(doc_Proy_extra, sheetname='Hoja1',  header=None)
        df_Comp       = pd.read_excel(doc_Comp      , sheetname='Hoja1',  header=None)
        
        ## Se escoge la columna que se desea realizar la prediccion
        #forcast_col = 4
        ## Llena con -99999 las celdas o espacios vacíos del DF
        #df_Proy.fillna(-99999, inplace=True)
        
        # Numero total de datos a proyectar, se toma 24hrs x 7 días
        forecast_out = 24*7
        
        # Lista de datos de históricos de proyección
        amba_1 = (df_Proy_extra.iloc[:,4].values.tolist())
        toto_1 = (df_Proy_extra.iloc[:,5].values.tolist())
        puyo_1 = (df_Proy_extra.iloc[:,6].values.tolist())
        tena_1 = (df_Proy_extra.iloc[:,7].values.tolist())
        bani_1 = (df_Proy_extra.iloc[:,8].values.tolist())
        
        amba_2 = (df_Proy.iloc[:,4].values.tolist())
        toto_2 = (df_Proy.iloc[:,5].values.tolist())
        puyo_2 = (df_Proy.iloc[:,6].values.tolist())
        tena_2 = (df_Proy.iloc[:,7].values.tolist())
        bani_2 = (df_Proy.iloc[:,8].values.tolist())
        
        # Elimino la primera fila de cada DF (TEXTOS TÍTULOS)
        amba_1.pop(0)
        toto_1.pop(0)
        puyo_1.pop(0)
        tena_1.pop(0)
        bani_1.pop(0)
        
        amba_2.pop(0)
        toto_2.pop(0)
        puyo_2.pop(0)
        tena_2.pop(0)
        bani_2.pop(0)
        
        amba_p = amba_1 + amba_2
        toto_p = toto_1 + toto_2
        puyo_p = puyo_1 + puyo_2
        tena_p = tena_1 + tena_2
        bani_p = bani_1 + bani_2
        
        amba_p = pd.DataFrame(amba_p)
        toto_p = pd.DataFrame(toto_p)
        puyo_p = pd.DataFrame(puyo_p)
        tena_p = pd.DataFrame(tena_p)
        bani_p = pd.DataFrame(bani_p)
        
        #==============================================================================
        #==============================================================================
        #===>UBICACIONES GENERALES de las posiciones para la proyección
        
        
        #===> Variables internas
        # FECHA SELECCIONADA
        fehca_selec = datetime(int(po_anio_p),(pos_mes_p+1),int(po_dia_p))
        # DÍAS DE PROYECCIÓN
        dias_proy = timedelta(days = 5952/24 + 1)   # 5952 * 24 = 248 días mas un día de salto
        
        # FECHA DE INICIO
        fecha_inicio = (fehca_selec + timedelta(days = 7)) - dias_proy
        # FECHA DE FIN
        fecha_fin = fehca_selec
        #fecha_fin = fehca_selec + timedelta(days = 7) # 24 horas de un día
        # FECHAS DE INICIO DE CADA AÑO POSICIÓN INICIAL
        enero_1 = datetime(fecha_inicio.year,1,1)
        #enero_2 = datetime(fecha_fin.year   ,1,1)
        
        ## POSICIÓN INICIAL
        val_ini = (abs(fecha_inicio - enero_1).days) * 24 # Posición del valor inicial 5952 DATOS
        val_fin = (abs(fecha_fin    - enero_1).days) * 24 # Posición del valor final 5952 + 24 = 5976 DATOS
        
        #print(val_ini)
        #print(val_fin)
        
        # ELIMINA LOS ÚLTIMOS VALORES DEL DF
        amba_p = amba_p[:val_fin]
        toto_p = toto_p[:val_fin]
        puyo_p = puyo_p[:val_fin]
        tena_p = tena_p[:val_fin]
        bani_p = bani_p[:val_fin]
        # ELIMINA LOS PRIMEROS VALORES DEL DF ** Queda un DF con 5952 datos
        amba_p = amba_p[val_ini:]
        toto_p = toto_p[val_ini:]
        puyo_p = puyo_p[val_ini:]
        tena_p = tena_p[val_ini:]
        bani_p = bani_p[val_ini:]
        # REINICIA EL INDEX DE CADA DF
        amba_p = amba_p.reset_index(drop=True)
        toto_p = toto_p.reset_index(drop=True)
        puyo_p = puyo_p.reset_index(drop=True)
        tena_p = tena_p.reset_index(drop=True)
        bani_p = bani_p.reset_index(drop=True)
        
        #==============================================================================
        #==============================================================================
        
        # Crea una nueva columna con la regresion apuntada del forecast_out
        # La nueva columna es una copia y desplazo de valores de la columna volume
        # La copia y desplaza los valores desde la posicion forecast_out.
        amba_p['Prediccion'] = amba_p[0].shift(-forecast_out)
        toto_p['Prediccion'] = toto_p[0].shift(-forecast_out)
        puyo_p['Prediccion'] = puyo_p[0].shift(-forecast_out)
        tena_p['Prediccion'] = tena_p[0].shift(-forecast_out)
        bani_p['Prediccion'] = bani_p[0].shift(-forecast_out)
        
        #Se crea un arreglo a partir del DF eliminando la columna de Prediccion
        X_amba = np.array(amba_p.drop(['Prediccion'],1))
        X_toto = np.array(toto_p.drop(['Prediccion'],1))
        X_puyo = np.array(puyo_p.drop(['Prediccion'],1))
        X_tena = np.array(tena_p.drop(['Prediccion'],1))
        X_bani = np.array(bani_p.drop(['Prediccion'],1))
        
        # Preprocesa y escala los datos del areglo X
        X_amba = preprocessing.scale(X_amba)
        X_toto = preprocessing.scale(X_toto)
        X_puyo = preprocessing.scale(X_puyo)
        X_tena = preprocessing.scale(X_tena)
        X_bani = preprocessing.scale(X_bani)
        
        # Crea nuevo arreglo tomando los (forecast_out = 7*24= 168) Ultimos datos de X
        X_lately_amba = X_amba[-forecast_out:]
        X_lately_toto = X_toto[-forecast_out:]
        X_lately_puyo = X_puyo[-forecast_out:]
        X_lately_tena = X_tena[-forecast_out:]
        X_lately_bani = X_bani[-forecast_out:]
        
        #El nuevo arreglo contiende todos los datos, incluso el valor de (forecast_out)
        #Se queda con los primeros datos excepto los 168 últimos
        X_amba = X_amba[:-forecast_out:]
        X_toto = X_toto[:-forecast_out:]
        X_puyo = X_puyo[:-forecast_out:]
        X_tena = X_tena[:-forecast_out:]
        X_bani = X_bani[:-forecast_out:]
        
        # Elimina todas las filas que tienen datos vaci­os NAN  o sea los (forecast_out) ultimos datos
        amba_p.dropna(inplace=True)
        toto_p.dropna(inplace=True)
        puyo_p.dropna(inplace=True)
        tena_p.dropna(inplace=True)
        bani_p.dropna(inplace=True)
        
        
        # Se crea nuevo arreglo que contiene la columna de Prediccion
        y_amba = np.array(amba_p['Prediccion'])
        y_toto = np.array(toto_p['Prediccion'])
        y_puyo = np.array(puyo_p['Prediccion'])
        y_tena = np.array(tena_p['Prediccion'])
        y_bani = np.array(bani_p['Prediccion'])
        
        #==============================================================================
        #=======================   INICIO ALGORITMO SVM    ============================
        #==============================================================================
        
        model = svm.SVR(kernel='rbf',gamma='auto',C=10)
        
        # Metodologi­a de validacion cruzada para los datos obtenidos
        X_train_amba, X_test_amba, y_train_amba, y_test_amba = cross_validation.train_test_split(X_amba, y_amba, test_size=0.3)
        model.fit(X_train_amba,y_train_amba)
        accuracy_amba = model.score(X_test_amba,y_test_amba)
        ProyAmba = model.predict(X_lately_amba)# ============>>>   SALIDA PROYECCIÓN
        
        X_train_toto, X_test_toto, y_train_toto, y_test_toto = cross_validation.train_test_split(X_toto, y_toto, test_size=0.3)
        model.fit(X_train_toto,y_train_toto)
        accuracy_toto = model.score(X_test_toto,y_test_toto)
        ProyToto = model.predict(X_lately_toto)# ============>>>   SALIDA PROYECCIÓN            forecast_set
        
        X_train_puyo, X_test_puyo, y_train_puyo, y_test_puyo = cross_validation.train_test_split(X_puyo, y_puyo, test_size=0.3)
        model.fit(X_train_puyo,y_train_puyo)
        accuracy_puyo = model.score(X_test_puyo,y_test_puyo)
        ProyPuyo = model.predict(X_lately_puyo)# ============>>>   SALIDA PROYECCIÓN            forecast_set
        
        X_train_tena, X_test_tena, y_train_tena, y_test_tena = cross_validation.train_test_split(X_tena, y_tena, test_size=0.3)
        model.fit(X_train_tena,y_train_tena)
        accuracy_tena = model.score(X_test_tena,y_test_tena)
        ProyTena = model.predict(X_lately_tena)# ============>>>   SALIDA PROYECCIÓN            forecast_set
        
        X_train_bani, X_test_bani, y_train_bani, y_test_bani = cross_validation.train_test_split(X_bani, y_bani, test_size=0.3)
        model.fit(X_train_bani,y_train_bani)
        accuracy_bani = model.score(X_test_bani,y_test_bani)
        ProyBani = model.predict(X_lately_bani)# ============>>>   SALIDA PROYECCIÓN            forecast_set
        
        # Datos de comparación
        #CompAmba = (amba_p[0][-forecast_out:]).reset_index(drop=True)
        #CompToto = (toto_p[0][-forecast_out:]).reset_index(drop=True)
        #CompPuyo = (puyo_p[0][-forecast_out:]).reset_index(drop=True)
        #CompTena = (tena_p[0][-forecast_out:]).reset_index(drop=True)
        #CompBani = (bani_p[0][-forecast_out:]).reset_index(drop=True)
        #==============================================================================
        #========================== SEMANA DE COMPARACIÓN =============================
        #==============================================================================
        
        mes_c = df_Comp.iloc[:,0].values.tolist()
        mes_c.pop(0)
        
        dia_c = df_Comp.iloc[:,1].values.tolist()
        dia_c.pop(0)
        
        hora_c = df_Comp.iloc[:,2].values.tolist()
        hora_c.pop(0)
        
        amba_c = df_Comp.iloc[:,4].values.tolist()
        amba_c.pop(0)
        
        toto_c = df_Comp.iloc[:,5].values.tolist()
        toto_c.pop(0)
        
        puyo_c = df_Comp.iloc[:,6].values.tolist()
        puyo_c.pop(0)
        
        tena_c = df_Comp.iloc[:,7].values.tolist()
        tena_c.pop(0)
        
        bani_c = df_Comp.iloc[:,8].values.tolist()
        bani_c.pop(0)
        
        #===> Reemplaza valores (0) con (nan)
        for i in range(len(dia_c)):
            if amba_c[i] == 0:
                amba_c[i] = float('nan')
            if toto_c[i] == 0:
                toto_c[i] = float('nan')
            if puyo_c[i] == 0:
                puyo_c[i] = float('nan')
            if tena_c[i] == 0:
                tena_c[i] = float('nan')
            if bani_c[i] == 0:
                bani_c[i] = float('nan')
        
        #===> Se establece una matriz con los datos importados
        #data_c = np.column_stack((amba_c, toto_c, puyo_c, tena_c, bani_c))
        
        #==============================================================================
        #==============================================================================
        if po_mes_c == 'ENERO':
            if int(po_dia_c) < 8:
                
                doc_Proy_EXTRA_c = direccion_base_datos+'\\HIS_POT_' + str(int(po_anio_c)-1) + '.xls'
                df_Proy_EXTRA_c = pd.read_excel(doc_Proy_EXTRA_c, sheetname='Hoja1',  header=None)
                df_Proy_EXTRA_c = df_Proy_EXTRA_c[-(24*31):]
                
                mess_extra_c = df_Proy_EXTRA_c.iloc[:,0].values.tolist()
                amba_extra_c = df_Proy_EXTRA_c.iloc[:,4].values.tolist()
                toto_extra_c = df_Proy_EXTRA_c.iloc[:,5].values.tolist()
                puyo_extra_c = df_Proy_EXTRA_c.iloc[:,6].values.tolist()
                tena_extra_c = df_Proy_EXTRA_c.iloc[:,7].values.tolist()
                bani_extra_c = df_Proy_EXTRA_c.iloc[:,8].values.tolist()
                #===> Obtengo únicamente 7 días de diciembre del año pasado
                mess_extra_c = mess_extra_c[-(24*31):]
                #===> Elimino el mes de diciembre del año actual
                mes_c = mes_c[:-(24*31)]
                #===> Nueva lista con 7 días de Dic del año pasado mas resto de datos año actual
                mes_c = mess_extra_c + mes_c
                
                amba_c = amba_extra_c + amba_c
                toto_c = toto_extra_c + toto_c
                puyo_c = puyo_extra_c + puyo_c
                tena_c = tena_extra_c + tena_c
                bani_c = bani_extra_c + bani_c
        
        #===> APUNTA AL MES SELECCIONADO DE COMPARACIÓN
        x_meses_c = [mes_c.index('ENERO'),   mes_c.index('FEBRERO'),   mes_c.index('MARZO'),
                     mes_c.index('ABRIL'),   mes_c.index('MAYO'),      mes_c.index('JUNIO'), 
                     mes_c.index('JULIO'),   mes_c.index('AGOSTO'),    mes_c.index('SEPTIEMBRE'),
                     mes_c.index('OCTUBRE'), mes_c.index('NOVIEMBRE'), mes_c.index('DICIEMBRE')]
        
        for i in range (12):
            if pos_mes_c == i:
                if pos_mes_c == 11:
                    ubic_mes_c = x_meses_c[i]
                else:
                    ubic_mes_c = x_meses_c[i] # VARIABLE Posición del mes seleccionado
                    
        
        #===> DEFINE LA UBICACIÓN DEL DÍA SELECCIONADO DE COMPARACIÓN
    #    print(' ')
    #    print ('Fecha Comparación: ',po_dia_c,' de ',mes_c[ubic_mes_c],' de ',po_anio_c)
        if int(po_dia_c) == 1:
            ubicacion_c = ubic_mes_c
        else:
            ubicacion_c = (int(po_dia_c)-1) * 24 + ubic_mes_c
            
    #    print(' ')  
    #    print (ubicacion_c) #==========================================================>>> VARIABLE DE UBICACIÓN EXACTA EN LISTA DE DATOS de comparación
        
        amba_c = amba_c[:ubicacion_c]
        toto_c = toto_c[:ubicacion_c]
        puyo_c = puyo_c[:ubicacion_c]
        tena_c = tena_c[:ubicacion_c]
        bani_c = bani_c[:ubicacion_c]
        
        #===> Valores inicio y fin para lista de datos a ser usada
        val_ini_c = ubicacion_c - 24 * 7 #Posición del valor inicial
        val_fin_c = ubicacion_c - 1      #Posición del valor final
        
        #===> Comparación semananal S/E Ambato
        CompAmba = []
        for i in range (24*7):
            CompAmba.append(amba_c[val_ini_c + i ])  
        #===> Comparación semananal S/E Totoras
        CompToto = []
        for i in range (24*7):
            CompToto.append(toto_c[val_ini_c + i ])
        #===> Comparación semananal S/E Puyo
        CompPuyo = []
        for i in range (24*7):
            CompPuyo.append(puyo_c[val_ini_c + i  ])
        #===> Comparación semananal S/E Tena
        CompTena = []
        for i in range (24*7):
            CompTena.append(tena_c[val_ini_c + i ])
        #===> Comparación semananal S/E Baños
        CompBani = []
        for i in range (24*7):
            CompBani.append(bani_c[val_ini_c + i ])
        
        #==============================================================================
        #================================ GRÁFICAS ====================================
        #==============================================================================
        if po_Comparacion == 'SI':
        #===> Creamos una lista tipo entero para relacionar con las etiquetas
            can_datos = []
            for i in range(7*24):
                if i%2!=1:
                    can_datos.append(i)     
        #===> Creamos una lista con los números de horas del día seleccionado (etiquetas)
            horas_dia = []
            horas_str = []
            for i in range (7):
                for i in range (1,25):
                    if i%2!=0:
                        horas_dia.append(i)
            for i in range (len(horas_dia)):
                horas_str.append(str(horas_dia[i]))
                
        #===> Tamaño de la ventana de la gráfica
            plt.subplots(figsize=(15, 8))
            
        #===> Título general superior
            plt.suptitle(u' PROYECCIÓN SEMANAL DE CARGA\n SUPORT VECTOR MACHINE ',fontsize=14, fontweight='bold') 
            
            plt.subplot(5,1,1)
            plt.plot(CompAmba,'blue', label = 'Comparación')
            plt.plot(ProyAmba,'#DCD037', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E AMBATO\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,2)
            plt.plot(CompToto,'blue', label = 'Comparación')
            plt.plot(ProyToto,'#CD336F', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E TOTORAS\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,3)  
            plt.plot(CompPuyo,'blue', label = 'Comparación')
            plt.plot(ProyPuyo,'#349A9D', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E PUYO\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,4)  
            plt.plot(CompTena,'blue', label = 'Comparación')
            plt.plot(ProyTena,'#CC8634', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E TENA\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,5)  
            plt.plot(CompBani,'blue', label = 'Comparación')
            plt.plot(ProyBani,'#4ACB71', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E BAÑOS\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
        
        #graf = []
        #for i in range (52248, 52415):
        #    graf.append(xp[i])
        #plt.legend(loc=1)
        
        #=========> Fechas de salida:
        sem_ini = datetime(int(po_anio_p),(pos_mes_p+1),int(po_dia_p))
        sem_fin = sem_ini + timedelta(days = 6)
        
        #==============================================================================
        #======================== RUTINA GENERA EXCEL =================================
        #==============================================================================
        df = pd.DataFrame([' '])
        writer = ExcelWriter('03_PROYECCION_SVM.xls')
        df.to_excel(writer, 'Salida_Proyección_CECON', index=False)
        df.to_excel(writer, 'Salida_Comparación_CECON', index=False)
        df.to_excel(writer, 'Salida_ERRORES_CECON', index=False)
        writer.save()
        
        #abre el archivo de excel plantilla
        rb = open_workbook('03_PROYECCION_SVM.xls')
        #crea una copia del archivo plantilla
        wb = copy(rb)
        #se ingresa a la hoja 1 de la copia del archivo excel
        ws = wb.get_sheet(0)
        ws_c = wb.get_sheet(1)
        ws_e = wb.get_sheet(2)
        #===========================Hoja 1 Proyección =================================
        #ws.write(0,0,'MES')
        #ws.write(0,1,'DÍA')
        #ws.write(0,2,'#')
        ws.write(0,0,'METODOLOGÍA MÁQUINA DE VECTOR SOPORTE  (S.V.M.)')
        ws.write(1,0,'PROYECCIÓN SEMANAL DE CARGA')
        ws.write(2,0,'Semana de Proyección: del '+str(sem_ini.day)+'/'
                 +str(sem_ini.month) +'/'+str(sem_ini.year)+'\t al '+
                 str(sem_fin.day)+'/'+str(sem_fin.month)+'/'+str(sem_fin.year))
        
        ws.write(3,0,'Realizado por: '+str(po_Realizado))
        
        # Define lista de títulos
        titulos = ['HORA','S/E AMBATO','S/E TOTORAS','S/E PUYO','S/E TENA','S/E BAÑOS']
        
        for i in range(len(titulos)):
            ws.write(9,i,titulos[i])
            
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws.write(Aux,0,j+1)
                ws.write(Aux,1,ProyAmba[j + Aux2])
                ws.write(Aux,2,ProyToto[j + Aux2])
                ws.write(Aux,3,ProyPuyo[j + Aux2])
                ws.write(Aux,4,ProyTena[j + Aux2])
                ws.write(Aux,5,ProyBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
        #===========================Hoja 2 Comparación ================================
        
        #===> Fecha
        ws_c.write(0,0,'Semana de Proyección: del '+str(sem_ini.day)+'/'
                   +str(sem_ini.month) +'/'+str(sem_ini.year)+'\t al '+
                   str(sem_fin.day)+'/'+str(sem_fin.month)+'/'+str(sem_fin.year))
        ws_c.write(2,0,'PROYECCIÓN SEMANAL')
        
        for i in range(len(titulos)):
            ws_c.write(9,i,titulos[i])
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,0,j+1)
                ws_c.write(Aux,1,ProyAmba[j + Aux2])
                ws_c.write(Aux,2,ProyToto[j + Aux2])
                ws_c.write(Aux,3,ProyPuyo[j + Aux2])
                ws_c.write(Aux,4,ProyTena[j + Aux2])
                ws_c.write(Aux,5,ProyBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
            
        #ws_c.write(0,6,po_dia_c+'/'+po_mes_c+'/'+po_anio_c)
        ws_c.write(2,6,'SEMANA DE COMPARACIÓN')
        
        for i in range(len(titulos)):
            ws_c.write(9,i+6,titulos[i])
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,6,j+1)
                ws_c.write(Aux,7,CompAmba[j + Aux2])
                ws_c.write(Aux,8,CompToto[j + Aux2])
                ws_c.write(Aux,9,CompPuyo[j + Aux2])
                ws_c.write(Aux,10,CompTena[j + Aux2])
                ws_c.write(Aux,11,CompBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
            
            
        # ==========================   Cálculo de errores    ==========================
            
        ws_c.write(2,12,'CÁLCULO DE ERRORES')
        ws_c.write(3,12,'PORCENTAJE DE ERROR MEDIO ABSOLUTO (PEMA)')
        
        for i in range(len(titulos)):
            ws_c.write(9,i+12,titulos[i])
        
        
        errAmba = []
        errToto = []
        errPuyo = []
        errTena = []
        errBani = []
        
        for i in range(len(ProyAmba)):
            errAmba.append((abs( ProyAmba[i] - CompAmba[i] ) / CompAmba[i] )*100)
            errToto.append((abs( ProyToto[i] - CompToto[i] ) / CompToto[i] )*100)
            errPuyo.append((abs( ProyPuyo[i] - CompPuyo[i] ) / CompPuyo[i] )*100)
            errTena.append((abs( ProyTena[i] - CompTena[i] ) / CompTena[i] )*100)
            errBani.append((abs( ProyBani[i] - CompBani[i] ) / CompBani[i] )*100)
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,12,j+1)
                ws_c.write(Aux,13,errAmba[j + Aux2])
                ws_c.write(Aux,14,errToto[j + Aux2])
                ws_c.write(Aux,15,errPuyo[j + Aux2])
                ws_c.write(Aux,16,errTena[j + Aux2])
                ws_c.write(Aux,17,errBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
        # Suma de los valores de las listas
        SumAmba = 0
        SumToto = 0
        SumPuyo = 0
        SumTena = 0
        SumBani = 0
        for i in range (len(ProyAmba)):
            SumAmba = errAmba[i] + SumAmba
            SumToto = errToto[i] + SumToto
            SumPuyo = errPuyo[i] + SumPuyo
            SumTena = errTena[i] + SumTena
            SumBani = errBani[i] + SumBani
        # Almaceno en una lista los resultados de las sumas
        Sumas = [SumAmba, SumToto, SumPuyo, SumTena, SumBani]
        
        # Imprime los resultados en la fila correspondiente
        ws_c.write(Aux+1,11,'SUMATORIA TOTAL')
        for i in range (len(Sumas)):
             ws_c.write(Aux+1,i+13,Sumas[i])     
        
        # Cálculo de los promedios de las sumas
        ws_c.write(Aux+2,11,'ERROR MEDIO ABSOLUTO')
        for i in range(len(Sumas)):
            ws_c.write(Aux+2,i+13,(Sumas[i]/len(errAmba)))
        
        # Cálculo de la exactitud de la proyección
        ws_c.write(Aux+3,11,'EXACTITUD DE LA PROYECCIÓN')
        for i in range(len(Sumas)):
            ws_c.write(Aux+3,i+13,(100-(Sumas[i]/len(errAmba))))
        
        
        for i in range(len(titulos)-1):
            ws_c.write(Aux,i+13,titulos[i+1])
        
        # ==========================   SOLO   errores    ==========================
        
        # ==================================> cálculo de et = Comp - Proy
        for i in range(len(titulos)-1):
            ws_e.write(9,i,titulos[i+1])
        ws_e.write(7,0,'et = Comp - Proy')
        
        et_Amba = []
        et_Toto = []
        et_Puyo = []
        et_Tena = []
        et_Bani = []
        
        for i in range (len(ProyAmba)):
            et_Amba.append(CompAmba[i] - ProyAmba[i])
            et_Toto.append(CompToto[i] - ProyToto[i])
            et_Puyo.append(CompPuyo[i] - ProyPuyo[i])
            et_Tena.append(CompTena[i] - ProyTena[i])
            et_Bani.append(CompBani[i] - ProyBani[i])
            
            ws_e.write(i+10,0, (et_Amba[i]))
            ws_e.write(i+10,1, (et_Toto[i]))
            ws_e.write(i+10,2, (et_Puyo[i]))
            ws_e.write(i+10,3, (et_Tena[i]))
            ws_e.write(i+10,4, (et_Bani[i]))
        
        # ==================================> cálculo de abs(et) = abs(Comp - Proy)
        for i in range(len(titulos)-1):
            ws_e.write(9,i+6,titulos[i+1])
        ws_e.write(7,6,'abs(et) = abs(Comp - Proy)')
        
        abs_et_Amba = []
        abs_et_Toto = []
        abs_et_Puyo = []
        abs_et_Tena = []
        abs_et_Bani = []
        
        for i in range (len(ProyAmba)):
            abs_et_Amba.append(abs(et_Amba[i]))
            abs_et_Toto.append(abs(et_Toto[i]))
            abs_et_Puyo.append(abs(et_Puyo[i]))
            abs_et_Tena.append(abs(et_Tena[i]))
            abs_et_Bani.append(abs(et_Bani[i]))
            
            ws_e.write(i+10,6,  (abs_et_Amba[i]))
            ws_e.write(i+10,7,  (abs_et_Toto[i]))
            ws_e.write(i+10,8,  (abs_et_Puyo[i]))
            ws_e.write(i+10,9,  (abs_et_Tena[i]))
            ws_e.write(i+10,10, (abs_et_Bani[i]))
            
        # ==================================> cálculo de et^2
        for i in range(len(titulos)-1):
            ws_e.write(9,i+12,titulos[i+1])
        ws_e.write(7,12,'et^2')
        
        et_Amba2 = []
        et_Toto2 = []
        et_Puyo2 = []
        et_Tena2 = []
        et_Bani2 = []
        
        for i in range (len(ProyAmba)):
            et_Amba2.append((et_Amba[i])**2)
            et_Toto2.append((et_Toto[i])**2)
            et_Puyo2.append((et_Puyo[i])**2)
            et_Tena2.append((et_Tena[i])**2)
            et_Bani2.append((et_Bani[i])**2)
            
            ws_e.write(i+10,12, (et_Amba2[i]))
            ws_e.write(i+10,13, (et_Toto2[i]))
            ws_e.write(i+10,14, (et_Puyo2[i]))
            ws_e.write(i+10,15, (et_Tena2[i]))
            ws_e.write(i+10,16, (et_Bani2[i]))
            
        # ==================================> cálculo de abs(et) / Comp
        for i in range(len(titulos)-1):
            ws_e.write(9,i+18,titulos[i+1])
        ws_e.write(7,18,'abs(et) / Comp')
        
        d1_Amba = []
        d1_Toto = []
        d1_Puyo = []
        d1_Tena = []
        d1_Bani = []
        
        for i in range (len(ProyAmba)):
            d1_Amba.append(abs_et_Amba[i] / CompAmba[i])
            d1_Toto.append(abs_et_Toto[i] / CompToto[i])
            d1_Puyo.append(abs_et_Puyo[i] / CompPuyo[i])
            d1_Tena.append(abs_et_Tena[i] / CompTena[i])
            d1_Bani.append(abs_et_Bani[i] / CompBani[i])
            
            ws_e.write(i+10,18, (d1_Amba[i]))
            ws_e.write(i+10,19, (d1_Toto[i]))
            ws_e.write(i+10,20, (d1_Puyo[i]))
            ws_e.write(i+10,21, (d1_Tena[i]))
            ws_e.write(i+10,22, (d1_Bani[i]))   
        
        # ==================================> cálculo de et / Comp
        for i in range(len(titulos)-1):
            ws_e.write(9,i+24,titulos[i+1])
        ws_e.write(7,24,'et / Comp')
        
        d2_Amba = []
        d2_Toto = []
        d2_Puyo = []
        d2_Tena = []
        d2_Bani = []
        
        for i in range (len(ProyAmba)):
            d2_Amba.append(et_Amba[i] / CompAmba[i])
            d2_Toto.append(et_Toto[i] / CompToto[i])
            d2_Puyo.append(et_Puyo[i] / CompPuyo[i])
            d2_Tena.append(et_Tena[i] / CompTena[i])
            d2_Bani.append(et_Bani[i] / CompBani[i])
            
            ws_e.write(i+10,24, (d2_Amba[i]))
            ws_e.write(i+10,25, (d2_Toto[i]))
            ws_e.write(i+10,26, (d2_Puyo[i]))
            ws_e.write(i+10,27, (d2_Tena[i]))
            ws_e.write(i+10,28, (d2_Bani[i]))   
            
        ws_e.write(0,0, 'INDICADORES') 
        # ==================================> Cálculo DAM
        DAM_Amba = 0
        DAM_Toto = 0
        DAM_Puyo = 0
        DAM_Tena = 0
        DAM_Bani = 0
        
        for i in range (len(ProyAmba)):
            DAM_Amba = abs_et_Amba[i] + DAM_Amba
            DAM_Toto = abs_et_Toto[i] + DAM_Toto
            DAM_Puyo = abs_et_Puyo[i] + DAM_Puyo
            DAM_Tena = abs_et_Tena[i] + DAM_Tena
            DAM_Bani = abs_et_Bani[i] + DAM_Bani
            
        DAM_Amba = DAM_Amba / (len(ProyAmba))
        DAM_Toto = DAM_Toto / (len(ProyAmba))
        DAM_Puyo = DAM_Puyo / (len(ProyAmba))
        DAM_Tena = DAM_Tena / (len(ProyAmba))
        DAM_Bani = DAM_Bani / (len(ProyAmba))
        
        ws_e.write(1,0, ('DAM'))
        DAM = [DAM_Amba,DAM_Toto,DAM_Puyo,DAM_Tena,DAM_Bani]
        for i in range(len(DAM)):
            ws_e.write(1,i+1,DAM[i])
        
        
        # ==================================> Cálculo EMC 
        EMC_Amba = 0
        EMC_Toto = 0
        EMC_Puyo = 0
        EMC_Tena = 0
        EMC_Bani = 0
        
        for i in range (len(ProyAmba)):
            EMC_Amba = et_Amba2[i] + EMC_Amba
            EMC_Toto = et_Toto2[i] + EMC_Toto
            EMC_Puyo = et_Puyo2[i] + EMC_Puyo
            EMC_Tena = et_Tena2[i] + EMC_Tena
            EMC_Bani = et_Bani2[i] + EMC_Bani
            
        EMC_Amba = EMC_Amba / (len(ProyAmba))
        EMC_Toto = EMC_Toto / (len(ProyAmba))
        EMC_Puyo = EMC_Puyo / (len(ProyAmba))
        EMC_Tena = EMC_Tena / (len(ProyAmba))
        EMC_Bani = EMC_Bani / (len(ProyAmba))
        
        ws_e.write(2,0, ('EMC'))
        EMC = [EMC_Amba,EMC_Toto,EMC_Puyo,EMC_Tena,EMC_Bani]
        
        for i in range(len(EMC)):
            ws_e.write(2,i+1,EMC[i])
            
        # ==================================> Cálculo PEMA
        PEMA_Amba = 0
        PEMA_Toto = 0
        PEMA_Puyo = 0
        PEMA_Tena = 0
        PEMA_Bani = 0
        
        for i in range (len(ProyAmba)):
            PEMA_Amba = (abs_et_Amba[i] / CompAmba [i]) + PEMA_Amba
            PEMA_Toto = (abs_et_Toto[i] / CompToto [i]) + PEMA_Toto
            PEMA_Puyo = (abs_et_Puyo[i] / CompPuyo [i]) + PEMA_Puyo
            PEMA_Tena = (abs_et_Tena[i] / CompTena [i]) + PEMA_Tena
            PEMA_Bani = (abs_et_Bani[i] / CompBani [i]) + PEMA_Bani
            
        PEMA_Amba = (PEMA_Amba / (len(ProyAmba))) *100
        PEMA_Toto = (PEMA_Toto / (len(ProyAmba))) *100
        PEMA_Puyo = (PEMA_Puyo / (len(ProyAmba))) *100
        PEMA_Tena = (PEMA_Tena / (len(ProyAmba))) *100
        PEMA_Bani = (PEMA_Bani / (len(ProyAmba))) *100
        
        ws_e.write(3,0, ('PEMA'))
        PEMA = [PEMA_Amba,PEMA_Toto,PEMA_Puyo,PEMA_Tena,PEMA_Bani]
        
        for i in range(len(PEMA)):
            ws_e.write(3,i+1,PEMA[i])
            
        # ==================================> Cálculo PME
        PME_Amba = 0
        PME_Toto = 0
        PME_Puyo = 0
        PME_Tena = 0
        PME_Bani = 0
        
        for i in range (len(ProyAmba)):
            PME_Amba = (et_Amba[i] / CompAmba [i]) + PME_Amba
            PME_Toto = (et_Toto[i] / CompToto [i]) + PME_Toto
            PME_Puyo = (et_Puyo[i] / CompPuyo [i]) + PME_Puyo
            PME_Tena = (et_Tena[i] / CompTena [i]) + PME_Tena
            PME_Bani = (et_Bani[i] / CompBani [i]) + PME_Bani
            
        PME_Amba = (PME_Amba / (len(ProyAmba))) *100
        PME_Toto = (PME_Toto / (len(ProyAmba))) *100
        PME_Puyo = (PME_Puyo / (len(ProyAmba))) *100
        PME_Tena = (PME_Tena / (len(ProyAmba))) *100
        PME_Bani = (PME_Bani / (len(ProyAmba))) *100
        
        ws_e.write(4,0, ('PME'))
        PME = [PME_Amba,PME_Toto,PME_Puyo,PME_Tena,PME_Bani]
        
        for i in range(len(PME)):
            ws_e.write(4,i+1,PME[i])
            
        # ==================================> Cálculo EXACTITUD DE LA PROYECCION
        ws_e.write(5,0,'EP')
        for i in range(len(PEMA)):
            ws_e.write(5,i+1,(100-PEMA[i]))
        
        
        for i in range(len(titulos)-1):
            ws_e.write(0,i+1,titulos[i+1])
        global direccion_resultados
        wb.save(direccion_resultados + '\\03_PROYECCION_SVM.xls')
        #==============================================================================
        #==============================================================================
        
        print(' ')
        print(' ')
        print('* La fecha seleccionada es:',int(po_anio_p),'/',(pos_mes_p+1),'/',int(po_dia_p))
        print(' ')
        print(' ')
        print('* Porcentaje de Exactitud de la predicción por LR :')
        print(' ')
        print('==> S/E Ambato :',accuracy_amba*100,'%')
        print('==> S/E Totoras:',accuracy_toto*100,'%')
        print('==> S/E Puyo   :',accuracy_puyo*100,'%')
        print('==> S/E Tena   :',accuracy_tena*100,'%')
        print('==> S/E Baños  :',accuracy_bani*100,'%')
        print(' ')
        print(' ')
        print('*** SVM Completed...!!!')
        print(' ')
        print(' ')
        
    start_time = time()
    test()
    elapsed_time = time() - start_time
    print(' ')
    print(' ')
    print('/////////////////////////////////////////')
    print(' ')
    print('ALGORITMO UTILIZADO: SVM')
    print("Tiempo transcurrido: %.10f segundos." % elapsed_time)            
    print(' ') 
    print('/////////////////////////////////////////')    



def proyeccion_rl():
    # -*- coding: utf-8 -*-
    """
    Created on Tue Mar 26 08:51:09 2019
    ESCUELA POLITÉCNICA NACIONAL
    AUTOR: ALEXIS GUAMAN FIGUEROA
    6_MÓDULO DE PROYECCIÓN CON ALGORITMO LR
    """
#    from time import time
    def test():
        #========================= MODELO DE REGRESIÓN LINEAL =========================
        
        #===> LA MUESTRA USARÁ 5952 DATOS HISTÓRICOS
        #===> La variable forecast_out controla el tiempo de predicción a realizar
        
#        import pandas as pd
#        #import csv, math , datetime
#        #import time
#        import numpy as np
#        from sklearn import preprocessing, cross_validation, svm
#        from sklearn.linear_model import LinearRegression, LogisticRegression
#        import math
#        import matplotlib.pyplot as plt
#        from matplotlib import style
#        style.use('ggplot')
#        
#        
#        from pandas import ExcelWriter
#        
#        from xlrd import open_workbook
#        from xlutils.copy import copy
#        
#        from datetime import date
#        from datetime import datetime,timedelta
        
        #=============================== VAR GLOBALES =================================
        global pos_Metodologia, pos_Realizado, pos_Comparacion
        global pos_anio_p, pos_mes_p, pos_dia_p
        global pos_anio_c, pos_mes_c, pos_dia_c
        #******************************************************************************
        global po_Metodologia, po_Realizado, po_Comparacion
        global po_anio_p, po_mes_p, po_dia_p
        global po_anio_c, po_mes_c, po_dia_c
        #==============================================================================
        
        doc_Proy       = direccion_base_datos+'\\HIS_POT_' + po_anio_p + '.xls'
        doc_Proy_extra = direccion_base_datos+'\\HIS_POT_' + str(int(po_anio_p)-1) + '.xls'
        doc_Comp       = direccion_base_datos+'\\HIS_POT_' + po_anio_c + '.xls'
        
        
        # Genera el DataFrame a partir del archivo de datos
        df_Proy       = pd.read_excel(doc_Proy      , sheetname='Hoja1',  header=None)
        df_Proy_extra = pd.read_excel(doc_Proy_extra, sheetname='Hoja1',  header=None)
        df_Comp       = pd.read_excel(doc_Comp      , sheetname='Hoja1',  header=None)
        
        ## Se escoge la columna que se desea realizar la prediccion
        #forcast_col = 4
        ## Llena con -99999 las celdas o espacios vacíos del DF
        #df_Proy.fillna(-99999, inplace=True)
        
        # Numero total de datos a proyectar, se toma 24hrs x 7 días
        forecast_out = 24*7
        
        # Lista de datos de históricos de proyección
        amba_1 = (df_Proy_extra.iloc[:,4].values.tolist())
        toto_1 = (df_Proy_extra.iloc[:,5].values.tolist())
        puyo_1 = (df_Proy_extra.iloc[:,6].values.tolist())
        tena_1 = (df_Proy_extra.iloc[:,7].values.tolist())
        bani_1 = (df_Proy_extra.iloc[:,8].values.tolist())
        
        amba_2 = (df_Proy.iloc[:,4].values.tolist())
        toto_2 = (df_Proy.iloc[:,5].values.tolist())
        puyo_2 = (df_Proy.iloc[:,6].values.tolist())
        tena_2 = (df_Proy.iloc[:,7].values.tolist())
        bani_2 = (df_Proy.iloc[:,8].values.tolist())
        
        # Elimino la primera fila de cada DF (TEXTOS TÍTULOS)
        amba_1.pop(0)
        toto_1.pop(0)
        puyo_1.pop(0)
        tena_1.pop(0)
        bani_1.pop(0)
        
        amba_2.pop(0)
        toto_2.pop(0)
        puyo_2.pop(0)
        tena_2.pop(0)
        bani_2.pop(0)
        
        amba_p = amba_1 + amba_2
        toto_p = toto_1 + toto_2
        puyo_p = puyo_1 + puyo_2
        tena_p = tena_1 + tena_2
        bani_p = bani_1 + bani_2
        
        amba_p = pd.DataFrame(amba_p)
        toto_p = pd.DataFrame(toto_p)
        puyo_p = pd.DataFrame(puyo_p)
        tena_p = pd.DataFrame(tena_p)
        bani_p = pd.DataFrame(bani_p)
        
        #==============================================================================
        #==============================================================================
        #===>UBICACIONES GENERALES de las posiciones para la proyección
        
        
        #===> Variables internas
        # FECHA SELECCIONADA
        fehca_selec = datetime(int(po_anio_p),(pos_mes_p+1),int(po_dia_p))
        # DÍAS DE PROYECCIÓN
        dias_proy = timedelta(days = 5952/24 + 1)   # 5952 * 24 = 248 días mas un día de salto
        #dias_proy = timedelta(days = 10000/24 + 1)   # 5952 * 24 = 248 días mas un día de salto
        
        # FECHA DE INICIO
        fecha_inicio = (fehca_selec + timedelta(days = 7)) - dias_proy
        # FECHA DE FIN
        fecha_fin = fehca_selec
        #fecha_fin = fehca_selec + timedelta(days = 7) # 24 horas de un día
        # FECHAS DE INICIO DE CADA AÑO POSICIÓN INICIAL
        enero_1 = datetime(fecha_inicio.year,1,1)
        #enero_2 = datetime(fecha_fin.year   ,1,1)
        
        ## POSICIÓN INICIAL
        val_ini = (abs(fecha_inicio - enero_1).days) * 24 # Posición del valor inicial 5952 DATOS
        val_fin = (abs(fecha_fin    - enero_1).days) * 24 # Posición del valor final 5952 + 24 = 5976 DATOS
        
        #print(val_ini)
        #print(val_fin)
        
        # ELIMINA LOS ÚLTIMOS VALORES DEL DF
        amba_p = amba_p[:val_fin]
        toto_p = toto_p[:val_fin]
        puyo_p = puyo_p[:val_fin]
        tena_p = tena_p[:val_fin]
        bani_p = bani_p[:val_fin]
        # ELIMINA LOS PRIMEROS VALORES DEL DF ** Queda un DF con 5952 datos
        amba_p = amba_p[val_ini:]
        toto_p = toto_p[val_ini:]
        puyo_p = puyo_p[val_ini:]
        tena_p = tena_p[val_ini:]
        bani_p = bani_p[val_ini:]
        # REINICIA EL INDEX DE CADA DF
        amba_p = amba_p.reset_index(drop=True)
        toto_p = toto_p.reset_index(drop=True)
        puyo_p = puyo_p.reset_index(drop=True)
        tena_p = tena_p.reset_index(drop=True)
        bani_p = bani_p.reset_index(drop=True)
        
        #==============================================================================
        #==============================================================================
        
        # Crea una nueva columna con la regresion apuntada del forecast_out
        # La nueva columna es una copia y desplazo de valores de la columna volume
        # La copia y desplaza los valores desde la posicion forecast_out.
        amba_p['Prediccion'] = amba_p[0].shift(-forecast_out)
        toto_p['Prediccion'] = toto_p[0].shift(-forecast_out)
        puyo_p['Prediccion'] = puyo_p[0].shift(-forecast_out)
        tena_p['Prediccion'] = tena_p[0].shift(-forecast_out)
        bani_p['Prediccion'] = bani_p[0].shift(-forecast_out)
        
        #Se crea un arreglo a partir del DF eliminando la columna de Prediccion
        X_amba = np.array(amba_p.drop(['Prediccion'],1))
        X_toto = np.array(toto_p.drop(['Prediccion'],1))
        X_puyo = np.array(puyo_p.drop(['Prediccion'],1))
        X_tena = np.array(tena_p.drop(['Prediccion'],1))
        X_bani = np.array(bani_p.drop(['Prediccion'],1))
        
        # Preprocesa y escala los datos del areglo X
        X_amba = preprocessing.scale(X_amba)
        X_toto = preprocessing.scale(X_toto)
        X_puyo = preprocessing.scale(X_puyo)
        X_tena = preprocessing.scale(X_tena)
        X_bani = preprocessing.scale(X_bani)
        
        # Crea nuevo arreglo tomando los (forecast_out = 7*24= 168) Ultimos datos de X
        X_lately_amba = X_amba[-forecast_out:]
        X_lately_toto = X_toto[-forecast_out:]
        X_lately_puyo = X_puyo[-forecast_out:]
        X_lately_tena = X_tena[-forecast_out:]
        X_lately_bani = X_bani[-forecast_out:]
        
        #El nuevo arreglo contiende todos los datos, incluso el valor de (forecast_out)
        #Se queda con los primeros datos excepto los 168 últimos
        X_amba = X_amba[:-forecast_out:]
        X_toto = X_toto[:-forecast_out:]
        X_puyo = X_puyo[:-forecast_out:]
        X_tena = X_tena[:-forecast_out:]
        X_bani = X_bani[:-forecast_out:]
        
        # Elimina todas las filas que tienen datos vaci­os NAN  o sea los (forecast_out) ultimos datos
        amba_p.dropna(inplace=True)
        toto_p.dropna(inplace=True)
        puyo_p.dropna(inplace=True)
        tena_p.dropna(inplace=True)
        bani_p.dropna(inplace=True)
        
        
        # Se crea nuevo arreglo que contiene la columna de Prediccion
        y_amba = np.array(amba_p['Prediccion'])
        y_toto = np.array(toto_p['Prediccion'])
        y_puyo = np.array(puyo_p['Prediccion'])
        y_tena = np.array(tena_p['Prediccion'])
        y_bani = np.array(bani_p['Prediccion'])
        
        #==============================================================================
        #=======================   INICIO ALGORITMO RL    =============================
        #==============================================================================
        
        clf = LinearRegression()
        # Metodologi­a de validacion cruzada para los datos obtenidos
        X_train_amba, X_test_amba, y_train_amba, y_test_amba = cross_validation.train_test_split(X_amba, y_amba, test_size=0.3)
        clf.fit(X_train_amba,y_train_amba)
        accuracy_amba = clf.score(X_test_amba,y_test_amba)
        ProyAmba = clf.predict(X_lately_amba)# ============>>>   SALIDA PROYECCIÓN            forecast_set
        
        X_train_toto, X_test_toto, y_train_toto, y_test_toto = cross_validation.train_test_split(X_toto, y_toto, test_size=0.3)
        clf.fit(X_train_toto,y_train_toto)
        accuracy_toto = clf.score(X_test_toto,y_test_toto)
        ProyToto = clf.predict(X_lately_toto)# ============>>>   SALIDA PROYECCIÓN            forecast_set
        
        X_train_puyo, X_test_puyo, y_train_puyo, y_test_puyo = cross_validation.train_test_split(X_puyo, y_puyo, test_size=0.3)
        clf.fit(X_train_puyo,y_train_puyo)
        accuracy_puyo = clf.score(X_test_puyo,y_test_puyo)
        ProyPuyo = clf.predict(X_lately_puyo)# ============>>>   SALIDA PROYECCIÓN            forecast_set
        
        X_train_tena, X_test_tena, y_train_tena, y_test_tena = cross_validation.train_test_split(X_tena, y_tena, test_size=0.3)
        clf.fit(X_train_tena,y_train_tena)
        accuracy_tena = clf.score(X_test_tena,y_test_tena)
        ProyTena = clf.predict(X_lately_tena)# ============>>>   SALIDA PROYECCIÓN            forecast_set
        
        X_train_bani, X_test_bani, y_train_bani, y_test_bani = cross_validation.train_test_split(X_bani, y_bani, test_size=0.3)
        clf.fit(X_train_bani,y_train_bani)
        accuracy_bani = clf.score(X_test_bani,y_test_bani)
        ProyBani = clf.predict(X_lately_bani)# ============>>>   SALIDA PROYECCIÓN            forecast_set
        
        # Datos de comparación
        #CompAmba = (amba_p[0][-forecast_out:]).reset_index(drop=True)
        #CompToto = (toto_p[0][-forecast_out:]).reset_index(drop=True)
        #CompPuyo = (puyo_p[0][-forecast_out:]).reset_index(drop=True)
        #CompTena = (tena_p[0][-forecast_out:]).reset_index(drop=True)
        #CompBani = (bani_p[0][-forecast_out:]).reset_index(drop=True)
        #==============================================================================
        #========================== SEMANA DE COMPARACIÓN =============================
        #==============================================================================
        
        mes_c = df_Comp.iloc[:,0].values.tolist()
        mes_c.pop(0)
        
        dia_c = df_Comp.iloc[:,1].values.tolist()
        dia_c.pop(0)
        
        hora_c = df_Comp.iloc[:,2].values.tolist()
        hora_c.pop(0)
        
        amba_c = df_Comp.iloc[:,4].values.tolist()
        amba_c.pop(0)
        
        toto_c = df_Comp.iloc[:,5].values.tolist()
        toto_c.pop(0)
        
        puyo_c = df_Comp.iloc[:,6].values.tolist()
        puyo_c.pop(0)
        
        tena_c = df_Comp.iloc[:,7].values.tolist()
        tena_c.pop(0)
        
        bani_c = df_Comp.iloc[:,8].values.tolist()
        bani_c.pop(0)
        
        #===> Reemplaza valores (0) con (nan)
        for i in range(len(dia_c)):
            if amba_c[i] == 0:
                amba_c[i] = float('nan')
            if toto_c[i] == 0:
                toto_c[i] = float('nan')
            if puyo_c[i] == 0:
                puyo_c[i] = float('nan')
            if tena_c[i] == 0:
                tena_c[i] = float('nan')
            if bani_c[i] == 0:
                bani_c[i] = float('nan')
        
        #===> Se establece una matriz con los datos importados
        #data_c = np.column_stack((amba_c, toto_c, puyo_c, tena_c, bani_c))
        
        #==============================================================================
        #==============================================================================
        if po_mes_c == 'ENERO':
            if int(po_dia_c) < 8:
                
                doc_Proy_EXTRA_c = direccion_base_datos+'\\HIS_POT_' + str(int(po_anio_c)-1) + '.xls'
                df_Proy_EXTRA_c = pd.read_excel(doc_Proy_EXTRA_c, sheetname='Hoja1',  header=None)
                df_Proy_EXTRA_c = df_Proy_EXTRA_c[-(24*31):]
                
                mess_extra_c = df_Proy_EXTRA_c.iloc[:,0].values.tolist()
                amba_extra_c = df_Proy_EXTRA_c.iloc[:,4].values.tolist()
                toto_extra_c = df_Proy_EXTRA_c.iloc[:,5].values.tolist()
                puyo_extra_c = df_Proy_EXTRA_c.iloc[:,6].values.tolist()
                tena_extra_c = df_Proy_EXTRA_c.iloc[:,7].values.tolist()
                bani_extra_c = df_Proy_EXTRA_c.iloc[:,8].values.tolist()
                #===> Obtengo únicamente 7 días de diciembre del año pasado
                mess_extra_c = mess_extra_c[-(24*31):]
                #===> Elimino el mes de diciembre del año actual
                mes_c = mes_c[:-(24*31)]
                #===> Nueva lista con 7 días de Dic del año pasado mas resto de datos año actual
                mes_c = mess_extra_c + mes_c
                
                amba_c = amba_extra_c + amba_c
                toto_c = toto_extra_c + toto_c
                puyo_c = puyo_extra_c + puyo_c
                tena_c = tena_extra_c + tena_c
                bani_c = bani_extra_c + bani_c
        
        #===> APUNTA AL MES SELECCIONADO DE COMPARACIÓN
        x_meses_c = [mes_c.index('ENERO'),   mes_c.index('FEBRERO'),   mes_c.index('MARZO'),
                     mes_c.index('ABRIL'),   mes_c.index('MAYO'),      mes_c.index('JUNIO'), 
                     mes_c.index('JULIO'),   mes_c.index('AGOSTO'),    mes_c.index('SEPTIEMBRE'),
                     mes_c.index('OCTUBRE'), mes_c.index('NOVIEMBRE'), mes_c.index('DICIEMBRE')]
        
        for i in range (12):
            if pos_mes_c == i:
                if pos_mes_c == 11:
                    ubic_mes_c = x_meses_c[i]
                else:
                    ubic_mes_c = x_meses_c[i] # VARIABLE Posición del mes seleccionado
                    
        
        #===> DEFINE LA UBICACIÓN DEL DÍA SELECCIONADO DE COMPARACIÓN
        print(' ')
        print ('Fecha Comparación: ',po_dia_c,' de ',mes_c[ubic_mes_c],' de ',po_anio_c)
        if int(po_dia_c) == 1:
            ubicacion_c = ubic_mes_c
        else:
            ubicacion_c = (int(po_dia_c)-1) * 24 + ubic_mes_c
            
    #    print(' ')  
    #    print (ubicacion_c) #==========================================================>>> VARIABLE DE UBICACIÓN EXACTA EN LISTA DE DATOS de comparación
        
        amba_c = amba_c[:ubicacion_c]
        toto_c = toto_c[:ubicacion_c]
        puyo_c = puyo_c[:ubicacion_c]
        tena_c = tena_c[:ubicacion_c]
        bani_c = bani_c[:ubicacion_c]
        
        #===> Valores inicio y fin para lista de datos a ser usada
        val_ini_c = ubicacion_c - 24 * 7 #Posición del valor inicial
        val_fin_c = ubicacion_c - 1      #Posición del valor final
        
        #===> Comparación semananal S/E Ambato
        CompAmba = []
        for i in range (24*7):
            CompAmba.append(amba_c[val_ini_c + i ])  
        #===> Comparación semananal S/E Totoras
        CompToto = []
        for i in range (24*7):
            CompToto.append(toto_c[val_ini_c + i ])
        #===> Comparación semananal S/E Puyo
        CompPuyo = []
        for i in range (24*7):
            CompPuyo.append(puyo_c[val_ini_c + i  ])
        #===> Comparación semananal S/E Tena
        CompTena = []
        for i in range (24*7):
            CompTena.append(tena_c[val_ini_c + i ])
        #===> Comparación semananal S/E Baños
        CompBani = []
        for i in range (24*7):
            CompBani.append(bani_c[val_ini_c + i ])
        
        
        #==============================================================================
        #================================ GRÁFICAS ====================================
        #==============================================================================
        if po_Comparacion == 'SI':
        #===> Creamos una lista tipo entero para relacionar con las etiquetas
            can_datos = []
            for i in range(7*24):
                if i%2!=1:
                    can_datos.append(i)     
        #===> Creamos una lista con los números de horas del día seleccionado (etiquetas)
            horas_dia = []
            horas_str = []
            for i in range (7):
                for i in range (1,25):
                    if i%2!=0:
                        horas_dia.append(i)
            for i in range (len(horas_dia)):
                horas_str.append(str(horas_dia[i]))
                
        #===> Tamaño de la ventana de la gráfica
            plt.subplots(figsize=(15, 8))
            
        #===> Título general superior
            plt.suptitle(u' PROYECCIÓN SEMANAL DE CARGA\n REGRESIÓN LINEAL ',fontsize=14, fontweight='bold') 
            
            plt.subplot(5,1,1)
            plt.plot(CompAmba,'blue', label = 'Comparación')
            plt.plot(ProyAmba,'#DCD037', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E AMBATO\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,2)
            plt.plot(CompToto,'blue', label = 'Comparación')
            plt.plot(ProyToto,'#CD336F', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E TOTORAS\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,3)  
            plt.plot(CompPuyo,'blue', label = 'Comparación')
            plt.plot(ProyPuyo,'#349A9D', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E PUYO\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,4)  
            plt.plot(CompTena,'blue', label = 'Comparación')
            plt.plot(ProyTena,'#CC8634', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E TENA\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,5)  
            plt.plot(CompBani,'blue', label = 'Comparación')
            plt.plot(ProyBani,'#4ACB71', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E BAÑOS\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
        
        #graf = []
        #for i in range (52248, 52415):
        #    graf.append(xp[i])
        #plt.legend(loc=1)
        
        #=========> Fechas de salida:
        sem_ini = datetime(int(po_anio_p),(pos_mes_p+1),int(po_dia_p))
        sem_fin = sem_ini + timedelta(days = 6)
        
        #==============================================================================
        #======================== RUTINA GENERA EXCEL =================================
        #==============================================================================
        df = pd.DataFrame([' '])
        writer = ExcelWriter('02_PROYECCION_RL.xls')
        df.to_excel(writer, 'Salida_Proyección_CECON', index=False)
        df.to_excel(writer, 'Salida_Comparación_CECON', index=False)
        df.to_excel(writer, 'Salida_ERRORES_CECON', index=False)
        writer.save()
        
        #abre el archivo de excel plantilla
        rb = open_workbook('02_PROYECCION_RL.xls')
        #crea una copia del archivo plantilla
        wb = copy(rb)
        #se ingresa a la hoja 1 de la copia del archivo excel
        ws = wb.get_sheet(0)
        ws_c = wb.get_sheet(1)
        ws_e = wb.get_sheet(2)
        #===========================Hoja 1 Proyección =================================
        #ws.write(0,0,'MES')
        #ws.write(0,1,'DÍA')
        #ws.write(0,2,'#')
        ws.write(0,0,'METODOLOGÍA REGRESIÓN LINEAL  (L.R.)')
        ws.write(1,0,'PROYECCIÓN SEMANAL DE CARGA')
        ws.write(2,0,'Semana de Proyección: del '+str(sem_ini.day)+'/'
                 +str(sem_ini.month) +'/'+str(sem_ini.year)+'\t al '+
                 str(sem_fin.day)+'/'+str(sem_fin.month)+'/'+str(sem_fin.year))
        
        ws.write(3,0,'Realizado por: '+str(po_Realizado))
        
        # Define lista de títulos
        titulos = ['HORA','S/E AMBATO','S/E TOTORAS','S/E PUYO','S/E TENA','S/E BAÑOS']
        
        for i in range(len(titulos)):
            ws.write(9,i,titulos[i])
            
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws.write(Aux,0,j+1)
                ws.write(Aux,1,ProyAmba[j + Aux2])
                ws.write(Aux,2,ProyToto[j + Aux2])
                ws.write(Aux,3,ProyPuyo[j + Aux2])
                ws.write(Aux,4,ProyTena[j + Aux2])
                ws.write(Aux,5,ProyBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
        #===========================Hoja 2 Comparación ================================
        
        #===> Fecha
        ws_c.write(0,0,'Semana de Proyección: del '+str(sem_ini.day)+'/'
                   +str(sem_ini.month) +'/'+str(sem_ini.year)+'\t al '+
                   str(sem_fin.day)+'/'+str(sem_fin.month)+'/'+str(sem_fin.year))
        ws_c.write(2,0,'PROYECCIÓN SEMANAL')
        
        for i in range(len(titulos)):
            ws_c.write(9,i,titulos[i])
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,0,j+1)
                ws_c.write(Aux,1,ProyAmba[j + Aux2])
                ws_c.write(Aux,2,ProyToto[j + Aux2])
                ws_c.write(Aux,3,ProyPuyo[j + Aux2])
                ws_c.write(Aux,4,ProyTena[j + Aux2])
                ws_c.write(Aux,5,ProyBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
            
        #ws_c.write(0,6,po_dia_c+'/'+po_mes_c+'/'+po_anio_c)
        ws_c.write(2,6,'SEMANA DE COMPARACIÓN')
        
        for i in range(len(titulos)):
            ws_c.write(9,i+6,titulos[i])
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,6,j+1)
                ws_c.write(Aux,7,CompAmba[j + Aux2])
                ws_c.write(Aux,8,CompToto[j + Aux2])
                ws_c.write(Aux,9,CompPuyo[j + Aux2])
                ws_c.write(Aux,10,CompTena[j + Aux2])
                ws_c.write(Aux,11,CompBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
            
            
        # ==========================   Cálculo de errores    ==========================
            
        ws_c.write(2,12,'CÁLCULO DE ERRORES')
        ws_c.write(3,12,'PORCENTAJE DE ERROR MEDIO ABSOLUTO (PEMA)')
        
        for i in range(len(titulos)):
            ws_c.write(9,i+12,titulos[i])
        
        
        errAmba = []
        errToto = []
        errPuyo = []
        errTena = []
        errBani = []
        
        for i in range(len(ProyAmba)):
            errAmba.append((abs( ProyAmba[i] - CompAmba[i] ) / CompAmba[i] )*100)
            errToto.append((abs( ProyToto[i] - CompToto[i] ) / CompToto[i] )*100)
            errPuyo.append((abs( ProyPuyo[i] - CompPuyo[i] ) / CompPuyo[i] )*100)
            errTena.append((abs( ProyTena[i] - CompTena[i] ) / CompTena[i] )*100)
            errBani.append((abs( ProyBani[i] - CompBani[i] ) / CompBani[i] )*100)
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,12,j+1)
                ws_c.write(Aux,13,errAmba[j + Aux2])
                ws_c.write(Aux,14,errToto[j + Aux2])
                ws_c.write(Aux,15,errPuyo[j + Aux2])
                ws_c.write(Aux,16,errTena[j + Aux2])
                ws_c.write(Aux,17,errBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
        # Suma de los valores de las listas
        SumAmba = 0
        SumToto = 0
        SumPuyo = 0
        SumTena = 0
        SumBani = 0
        for i in range (len(ProyAmba)):
            SumAmba = errAmba[i] + SumAmba
            SumToto = errToto[i] + SumToto
            SumPuyo = errPuyo[i] + SumPuyo
            SumTena = errTena[i] + SumTena
            SumBani = errBani[i] + SumBani
        # Almaceno en una lista los resultados de las sumas
        Sumas = [SumAmba, SumToto, SumPuyo, SumTena, SumBani]
        
        # Imprime los resultados en la fila correspondiente
        ws_c.write(Aux+1,11,'SUMATORIA TOTAL')
        for i in range (len(Sumas)):
             ws_c.write(Aux+1,i+13,Sumas[i])     
        
        # Cálculo de los promedios de las sumas
        ws_c.write(Aux+2,11,'ERROR MEDIO ABSOLUTO')
        for i in range(len(Sumas)):
            ws_c.write(Aux+2,i+13,(Sumas[i]/len(errAmba)))
        
        # Cálculo de la exactitud de la proyección
        ws_c.write(Aux+3,11,'EXACTITUD DE LA PROYECCIÓN')
        for i in range(len(Sumas)):
            ws_c.write(Aux+3,i+13,(100-(Sumas[i]/len(errAmba))))
        
        
        for i in range(len(titulos)-1):
            ws_c.write(Aux,i+13,titulos[i+1])
        
        # ==========================   SOLO   errores    ==========================
        
        # ==================================> cálculo de et = Comp - Proy
        for i in range(len(titulos)-1):
            ws_e.write(9,i,titulos[i+1])
        ws_e.write(7,0,'et = Comp - Proy')
        
        et_Amba = []
        et_Toto = []
        et_Puyo = []
        et_Tena = []
        et_Bani = []
        
        for i in range (len(ProyAmba)):
            et_Amba.append(CompAmba[i] - ProyAmba[i])
            et_Toto.append(CompToto[i] - ProyToto[i])
            et_Puyo.append(CompPuyo[i] - ProyPuyo[i])
            et_Tena.append(CompTena[i] - ProyTena[i])
            et_Bani.append(CompBani[i] - ProyBani[i])
            
            ws_e.write(i+10,0, (et_Amba[i]))
            ws_e.write(i+10,1, (et_Toto[i]))
            ws_e.write(i+10,2, (et_Puyo[i]))
            ws_e.write(i+10,3, (et_Tena[i]))
            ws_e.write(i+10,4, (et_Bani[i]))
        
        # ==================================> cálculo de abs(et) = abs(Comp - Proy)
        for i in range(len(titulos)-1):
            ws_e.write(9,i+6,titulos[i+1])
        ws_e.write(7,6,'abs(et) = abs(Comp - Proy)')
        
        abs_et_Amba = []
        abs_et_Toto = []
        abs_et_Puyo = []
        abs_et_Tena = []
        abs_et_Bani = []
        
        for i in range (len(ProyAmba)):
            abs_et_Amba.append(abs(et_Amba[i]))
            abs_et_Toto.append(abs(et_Toto[i]))
            abs_et_Puyo.append(abs(et_Puyo[i]))
            abs_et_Tena.append(abs(et_Tena[i]))
            abs_et_Bani.append(abs(et_Bani[i]))
            
            ws_e.write(i+10,6,  (abs_et_Amba[i]))
            ws_e.write(i+10,7,  (abs_et_Toto[i]))
            ws_e.write(i+10,8,  (abs_et_Puyo[i]))
            ws_e.write(i+10,9,  (abs_et_Tena[i]))
            ws_e.write(i+10,10, (abs_et_Bani[i]))
            
        # ==================================> cálculo de et^2
        for i in range(len(titulos)-1):
            ws_e.write(9,i+12,titulos[i+1])
        ws_e.write(7,12,'et^2')
        
        et_Amba2 = []
        et_Toto2 = []
        et_Puyo2 = []
        et_Tena2 = []
        et_Bani2 = []
        
        for i in range (len(ProyAmba)):
            et_Amba2.append((et_Amba[i])**2)
            et_Toto2.append((et_Toto[i])**2)
            et_Puyo2.append((et_Puyo[i])**2)
            et_Tena2.append((et_Tena[i])**2)
            et_Bani2.append((et_Bani[i])**2)
            
            ws_e.write(i+10,12, (et_Amba2[i]))
            ws_e.write(i+10,13, (et_Toto2[i]))
            ws_e.write(i+10,14, (et_Puyo2[i]))
            ws_e.write(i+10,15, (et_Tena2[i]))
            ws_e.write(i+10,16, (et_Bani2[i]))
            
        # ==================================> cálculo de abs(et) / Comp
        for i in range(len(titulos)-1):
            ws_e.write(9,i+18,titulos[i+1])
        ws_e.write(7,18,'abs(et) / Comp')
        
        d1_Amba = []
        d1_Toto = []
        d1_Puyo = []
        d1_Tena = []
        d1_Bani = []
        
        for i in range (len(ProyAmba)):
            d1_Amba.append(abs_et_Amba[i] / CompAmba[i])
            d1_Toto.append(abs_et_Toto[i] / CompToto[i])
            d1_Puyo.append(abs_et_Puyo[i] / CompPuyo[i])
            d1_Tena.append(abs_et_Tena[i] / CompTena[i])
            d1_Bani.append(abs_et_Bani[i] / CompBani[i])
            
            ws_e.write(i+10,18, (d1_Amba[i]))
            ws_e.write(i+10,19, (d1_Toto[i]))
            ws_e.write(i+10,20, (d1_Puyo[i]))
            ws_e.write(i+10,21, (d1_Tena[i]))
            ws_e.write(i+10,22, (d1_Bani[i]))   
        
        # ==================================> cálculo de et / Comp
        for i in range(len(titulos)-1):
            ws_e.write(9,i+24,titulos[i+1])
        ws_e.write(7,24,'et / Comp')
        
        d2_Amba = []
        d2_Toto = []
        d2_Puyo = []
        d2_Tena = []
        d2_Bani = []
        
        for i in range (len(ProyAmba)):
            d2_Amba.append(et_Amba[i] / CompAmba[i])
            d2_Toto.append(et_Toto[i] / CompToto[i])
            d2_Puyo.append(et_Puyo[i] / CompPuyo[i])
            d2_Tena.append(et_Tena[i] / CompTena[i])
            d2_Bani.append(et_Bani[i] / CompBani[i])
            
            ws_e.write(i+10,24, (d2_Amba[i]))
            ws_e.write(i+10,25, (d2_Toto[i]))
            ws_e.write(i+10,26, (d2_Puyo[i]))
            ws_e.write(i+10,27, (d2_Tena[i]))
            ws_e.write(i+10,28, (d2_Bani[i]))   
            
        ws_e.write(0,0, 'INDICADORES') 
        # ==================================> Cálculo DAM
        DAM_Amba = 0
        DAM_Toto = 0
        DAM_Puyo = 0
        DAM_Tena = 0
        DAM_Bani = 0
        
        for i in range (len(ProyAmba)):
            DAM_Amba = abs_et_Amba[i] + DAM_Amba
            DAM_Toto = abs_et_Toto[i] + DAM_Toto
            DAM_Puyo = abs_et_Puyo[i] + DAM_Puyo
            DAM_Tena = abs_et_Tena[i] + DAM_Tena
            DAM_Bani = abs_et_Bani[i] + DAM_Bani
            
        DAM_Amba = DAM_Amba / (len(ProyAmba))
        DAM_Toto = DAM_Toto / (len(ProyAmba))
        DAM_Puyo = DAM_Puyo / (len(ProyAmba))
        DAM_Tena = DAM_Tena / (len(ProyAmba))
        DAM_Bani = DAM_Bani / (len(ProyAmba))
        
        ws_e.write(1,0, ('DAM'))
        DAM = [DAM_Amba,DAM_Toto,DAM_Puyo,DAM_Tena,DAM_Bani]
        for i in range(len(DAM)):
            ws_e.write(1,i+1,DAM[i])
        
        
        # ==================================> Cálculo EMC 
        EMC_Amba = 0
        EMC_Toto = 0
        EMC_Puyo = 0
        EMC_Tena = 0
        EMC_Bani = 0
        
        for i in range (len(ProyAmba)):
            EMC_Amba = et_Amba2[i] + EMC_Amba
            EMC_Toto = et_Toto2[i] + EMC_Toto
            EMC_Puyo = et_Puyo2[i] + EMC_Puyo
            EMC_Tena = et_Tena2[i] + EMC_Tena
            EMC_Bani = et_Bani2[i] + EMC_Bani
            
        EMC_Amba = EMC_Amba / (len(ProyAmba))
        EMC_Toto = EMC_Toto / (len(ProyAmba))
        EMC_Puyo = EMC_Puyo / (len(ProyAmba))
        EMC_Tena = EMC_Tena / (len(ProyAmba))
        EMC_Bani = EMC_Bani / (len(ProyAmba))
        
        ws_e.write(2,0, ('EMC'))
        EMC = [EMC_Amba,EMC_Toto,EMC_Puyo,EMC_Tena,EMC_Bani]
        
        for i in range(len(EMC)):
            ws_e.write(2,i+1,EMC[i])
            
        # ==================================> Cálculo PEMA
        PEMA_Amba = 0
        PEMA_Toto = 0
        PEMA_Puyo = 0
        PEMA_Tena = 0
        PEMA_Bani = 0
        
        for i in range (len(ProyAmba)):
            PEMA_Amba = (abs_et_Amba[i] / CompAmba [i]) + PEMA_Amba
            PEMA_Toto = (abs_et_Toto[i] / CompToto [i]) + PEMA_Toto
            PEMA_Puyo = (abs_et_Puyo[i] / CompPuyo [i]) + PEMA_Puyo
            PEMA_Tena = (abs_et_Tena[i] / CompTena [i]) + PEMA_Tena
            PEMA_Bani = (abs_et_Bani[i] / CompBani [i]) + PEMA_Bani
            
        PEMA_Amba = (PEMA_Amba / (len(ProyAmba))) *100
        PEMA_Toto = (PEMA_Toto / (len(ProyAmba))) *100
        PEMA_Puyo = (PEMA_Puyo / (len(ProyAmba))) *100
        PEMA_Tena = (PEMA_Tena / (len(ProyAmba))) *100
        PEMA_Bani = (PEMA_Bani / (len(ProyAmba))) *100
        
        ws_e.write(3,0, ('PEMA'))
        PEMA = [PEMA_Amba,PEMA_Toto,PEMA_Puyo,PEMA_Tena,PEMA_Bani]
        
        for i in range(len(PEMA)):
            ws_e.write(3,i+1,PEMA[i])
            
        # ==================================> Cálculo PME
        PME_Amba = 0
        PME_Toto = 0
        PME_Puyo = 0
        PME_Tena = 0
        PME_Bani = 0
        
        for i in range (len(ProyAmba)):
            PME_Amba = (et_Amba[i] / CompAmba [i]) + PME_Amba
            PME_Toto = (et_Toto[i] / CompToto [i]) + PME_Toto
            PME_Puyo = (et_Puyo[i] / CompPuyo [i]) + PME_Puyo
            PME_Tena = (et_Tena[i] / CompTena [i]) + PME_Tena
            PME_Bani = (et_Bani[i] / CompBani [i]) + PME_Bani
            
        PME_Amba = (PME_Amba / (len(ProyAmba))) *100
        PME_Toto = (PME_Toto / (len(ProyAmba))) *100
        PME_Puyo = (PME_Puyo / (len(ProyAmba))) *100
        PME_Tena = (PME_Tena / (len(ProyAmba))) *100
        PME_Bani = (PME_Bani / (len(ProyAmba))) *100
        
        ws_e.write(4,0, ('PME'))
        PME = [PME_Amba,PME_Toto,PME_Puyo,PME_Tena,PME_Bani]
        
        for i in range(len(PME)):
            ws_e.write(4,i+1,PME[i])
            
        # ==================================> Cálculo EXACTITUD DE LA PROYECCION
        ws_e.write(5,0,'EP')
        for i in range(len(PEMA)):
            ws_e.write(5,i+1,(100-PEMA[i]))
        
        
        for i in range(len(titulos)-1):
            ws_e.write(0,i+1,titulos[i+1])
            
        global direccion_resultados
        wb.save(direccion_resultados + '\\02_PROYECCION_RL.xls')
        #==============================================================================
        #==============================================================================
        
        print(' ')
        print(' ')
        
        print('* Porcentaje de Exactitud de la Regresión Lineal:')
        print(' ')
        print('==> S/E Ambato :',accuracy_amba*100,'%')
        print('==> S/E Totoras:',accuracy_toto*100,'%')
        print('==> S/E Puyo   :',accuracy_puyo*100,'%')
        print('==> S/E Tena   :',accuracy_tena*100,'%')
        print('==> S/E Baños  :',accuracy_bani*100,'%')
        print(' ')
        print(' ')
        print('*** RL Completed...!!!')
        print(' ')
        print(' ')
    
    
    start_time = time()
    test()
    elapsed_time = time() - start_time
    print(' ')
    print(' ')
    print('/////////////////////////////////////////')
    print(' ')
    print('ALGORITMO UTILIZADO: LR')
    print("Tiempo transcurrido: %.10f segundos." % elapsed_time)            
    print(' ') 
    print('/////////////////////////////////////////')
    



def proyeccion_arima():
    # -*- coding: utf-8 -*-
    """
    Created on Wed Mar 13 23:50:06 2019
    ESCUELA POLITÉCNICA NACIONAL
    AUTOR: ALEXIS GUAMAN FIGUEROA
    5_MÓDULO DE PROYECCIÓN CON ALGORITMO ARIMA
    """
#    from time import time
    def test():
        
    
#        #importa librerías
#        import pandas as pd
#        import matplotlib.pyplot as plt
#        import numpy as np
#        
#        
#        from pandas import ExcelWriter
#        
#        from xlrd import open_workbook
#        from xlutils.copy import copy
#        
#        
#        from tkinter import messagebox
#        
#        import calendar
#        from datetime import datetime,timedelta
        #=============================== VAR GLOBALES =================================
        global pos_Metodologia, pos_Realizado, pos_Comparacion
        global pos_anio_p, pos_mes_p, pos_dia_p
        global pos_anio_c, pos_mes_c, pos_dia_c
        #******************************************************************************
        global po_Metodologia, po_Realizado, po_Comparacion
        global po_anio_p, po_mes_p, po_dia_p
        global po_anio_c, po_mes_c, po_dia_c
        global direccion_base_datos
        #==============================================================================
        
        doc_Proy = direccion_base_datos+'\\HIS_POT_' + po_anio_p + '.xls'
        doc_Comp = direccion_base_datos+'\\HIS_POT_' + po_anio_c + '.xls'
        
        
        df_Proy = pd.read_excel(doc_Proy, sheetname='Hoja1',  header=None)
        df_Comp = pd.read_excel(doc_Comp, sheetname='Hoja1',  header=None)
        
        #==============================================================================
        #========================== SEMANA DE PROYECCIÓN ==============================
        #==============================================================================
        
        mes_p = df_Proy.iloc[:,0].values.tolist()
        mes_p.pop(0)
        
        dia_p = df_Proy.iloc[:,1].values.tolist()
        dia_p.pop(0)
        
        hora_p = df_Proy.iloc[:,2].values.tolist()
        hora_p.pop(0)
        
        amba_p = df_Proy.iloc[:,4].values.tolist()
        amba_p.pop(0)
        
        toto_p = df_Proy.iloc[:,5].values.tolist()
        toto_p.pop(0)
        
        puyo_p = df_Proy.iloc[:,6].values.tolist()
        puyo_p.pop(0)
        
        tena_p = df_Proy.iloc[:,7].values.tolist()
        tena_p.pop(0)
        
        bani_p = df_Proy.iloc[:,8].values.tolist()
        bani_p.pop(0)
        
        ##===> Reemplaza valores (0) con (nan)
        #for i in range(len(mes_p)):
        #    if amba_p[i] == 0:
        #        amba_p[i] = float('nan')
        #    if toto_p[i] == 0:
        #        toto_p[i] = float('nan')
        #    if puyo_p[i] == 0:
        #        puyo_p[i] = float('nan')
        #    if tena_p[i] == 0:
        #        tena_p[i] = float('nan')
        #    if bani_p[i] == 0:
        #        bani_p[i] = float('nan')
        
        #===> Reemplaza valores (0) con (el inmediato siguiente)
        for i in range(len(mes_p)):
            if amba_p[i] == 0:
                amba_p[i] = amba_p[i+1]
            if toto_p[i] == 0:
                toto_p[i] = toto_p[i+1]
            if puyo_p[i] == 0:
                puyo_p[i] = puyo_p[i+1]
            if tena_p[i] == 0:
                tena_p[i] = tena_p[i+1]
            if bani_p[i] == 0:
                bani_p[i] = bani_p[i+1]
        
        
        
        
        
        #===> Se establece una matriz con los datos importados
        #data_p = np.column_stack((amba_p, toto_p, puyo_p, tena_p, bani_p))
                
        #==============================================================================
        #==============================================================================
        #===> Día de inicio de la semana de proyección
        
        if po_mes_p == 'ENERO':
            if int(po_dia_p) < 28:
                
                doc_Proy_EXTRA_p = direccion_base_datos+'\\HIS_POT_' + str(int(po_anio_p)-1) + '.xls'
                df_Proy_EXTRA_p = pd.read_excel(doc_Proy_EXTRA_p, sheetname='Hoja1',  header=None)
                df_Proy_Extra_p = df_Proy_EXTRA_p[-(24*31):]
                
                mess_extra = df_Proy_Extra_p.iloc[:,0].values.tolist()
                amba_extra = df_Proy_Extra_p.iloc[:,4].values.tolist()
                toto_extra = df_Proy_Extra_p.iloc[:,5].values.tolist()
                puyo_extra = df_Proy_Extra_p.iloc[:,6].values.tolist()
                tena_extra = df_Proy_Extra_p.iloc[:,7].values.tolist()
                bani_extra = df_Proy_Extra_p.iloc[:,8].values.tolist()
                #===> Obtengo únicamente el mes de diciembre del año pasado
                mess_extra = mess_extra[-(24*31):]
                #===> Elimino el mes de diciembre del año actual
                mes_p = mes_p[:-(24*31)]
                #===> Nueva lista con el Dic año pasado mas resto de datos año actual
                mes_p = mess_extra + mes_p
                
                amba_p = amba_extra + amba_p
                toto_p = toto_extra + toto_p
                puyo_p = puyo_extra + puyo_p
                tena_p = tena_extra + tena_p
                bani_p = bani_extra + bani_p
        
        
        #===> APUNTA AL MES SELECCIONADO DE PROYECCIÓN
        x_meses_p = [mes_p.index('ENERO'),   mes_p.index('FEBRERO'),   mes_p.index('MARZO'),
                     mes_p.index('ABRIL'),   mes_p.index('MAYO'),      mes_p.index('JUNIO'), 
                     mes_p.index('JULIO'),   mes_p.index('AGOSTO'),    mes_p.index('SEPTIEMBRE'),
                     mes_p.index('OCTUBRE'), mes_p.index('NOVIEMBRE'), mes_p.index('DICIEMBRE')]
        
        for i in range (12):
            if pos_mes_p == i:
                if pos_mes_p == 11:
                    ubic_mes = x_meses_p[i]
                else:
                    ubic_mes = x_meses_p[i] # VARIABLE Posición del mes seleccionado
                    
        #===> DEFINE LA UBICACIÓN DEL DÍA SELECCIONADO DE PROYECCIÓN
        print(' ')
        print ('Fecha de Proyección: ',po_dia_p,' de ',mes_p[ubic_mes],' de ',po_anio_p)
        if int(po_dia_p) == 1:
            ubicacion = ubic_mes
        else:
            ubicacion = (int(po_dia_p)-1) * 24 + ubic_mes
    #    print(' ')  
    #    print (ubicacion) #==========================================================>>> VARIABLE DE UBICACIÓN EXACTA EN LISTA DE DATOS
        
        amba_p = amba_p[:ubicacion]
        toto_p = toto_p[:ubicacion]
        puyo_p = puyo_p[:ubicacion]
        tena_p = tena_p[:ubicacion]
        bani_p = bani_p[:ubicacion]
        
              
        #===> Valores inicio y fin para lista de datos a ser usada
        val_ini = ubicacion - 24 * 7 * 4 #Posición del valor inicial de 4 semanas
        val_fin = ubicacion - 1          #Posición del valor final
        #==============================================================================
        #==============================================================================
        
        #===> Proyección semananal S/E Ambato
        ProyAmba = []
        for i in range (24*7):
            ProyAmba.append((amba_p[val_ini + i + 7*24*0] + 
                             amba_p[val_ini + i + 7*24*1] +
                             amba_p[val_ini + i + 7*24*2] + 
                             amba_p[val_ini + i + 7*24*3])/4)
        #===> Proyección semananal S/E Totoras
        ProyToto = []
        for i in range (24*7):
            ProyToto.append((toto_p[val_ini + i + 7*24*0] + 
                             toto_p[val_ini + i + 7*24*1] +
                             toto_p[val_ini + i + 7*24*2] + 
                             toto_p[val_ini + i + 7*24*3])/4)
        #===> Proyección semananal S/E Puyo
        ProyPuyo = []
        for i in range (24*7):
            ProyPuyo.append((puyo_p[val_ini + i + 7*24*0] + 
                             puyo_p[val_ini + i + 7*24*1] +
                             puyo_p[val_ini + i + 7*24*2] + 
                             puyo_p[val_ini + i + 7*24*3])/4)
        #===> Proyección semananal S/E Tena
        ProyTena = []
        for i in range (24*7):
            ProyTena.append((tena_p[val_ini + i + 7*24*0] + 
                             tena_p[val_ini + i + 7*24*1] +
                             tena_p[val_ini + i + 7*24*2] + 
                             tena_p[val_ini + i + 7*24*3])/4)
        #===> Proyección semananal S/E Baños
        ProyBani = []
        for i in range (24*7):
            ProyBani.append((bani_p[val_ini + i + 7*24*0] + 
                             bani_p[val_ini + i + 7*24*1] +
                             bani_p[val_ini + i + 7*24*2] + 
                             bani_p[val_ini + i + 7*24*3])/4)
        
        #==============================================================================
        #========================== SEMANA DE COMPARACIÓN =============================
        #==============================================================================
        
        mes_c = df_Comp.iloc[:,0].values.tolist()
        mes_c.pop(0)
        
        dia_c = df_Comp.iloc[:,1].values.tolist()
        dia_c.pop(0)
        
        hora_c = df_Comp.iloc[:,2].values.tolist()
        hora_c.pop(0)
        
        amba_c = df_Comp.iloc[:,4].values.tolist()
        amba_c.pop(0)
        
        toto_c = df_Comp.iloc[:,5].values.tolist()
        toto_c.pop(0)
        
        puyo_c = df_Comp.iloc[:,6].values.tolist()
        puyo_c.pop(0)
        
        tena_c = df_Comp.iloc[:,7].values.tolist()
        tena_c.pop(0)
        
        bani_c = df_Comp.iloc[:,8].values.tolist()
        bani_c.pop(0)
        
        #===> Reemplaza valores (0) con (nan)
        for i in range(len(mes_p)):
            if amba_c[i] == 0:
                amba_c[i] = float('nan')
            if toto_c[i] == 0:
                toto_c[i] = float('nan')
            if puyo_c[i] == 0:
                puyo_c[i] = float('nan')
            if tena_c[i] == 0:
                tena_c[i] = float('nan')
            if bani_c[i] == 0:
                bani_c[i] = float('nan')
        
        #===> Se establece una matriz con los datos importados
        #data_c = np.column_stack((amba_c, toto_c, puyo_c, tena_c, bani_c))
        
        #==============================================================================
        #==============================================================================
        if po_mes_c == 'ENERO':
            if int(po_dia_c) < 8:
                
                doc_Proy_EXTRA_c = direccion_base_datos+'\\HIS_POT_' + str(int(po_anio_c)-1) + '.xls'
                df_Proy_EXTRA_c = pd.read_excel(doc_Proy_EXTRA_c, sheetname='Hoja1',  header=None)
                df_Proy_EXTRA_c = df_Proy_EXTRA_c[-(24*31):]
                
                mess_extra_c = df_Proy_EXTRA_c.iloc[:,0].values.tolist()
                amba_extra_c = df_Proy_EXTRA_c.iloc[:,4].values.tolist()
                toto_extra_c = df_Proy_EXTRA_c.iloc[:,5].values.tolist()
                puyo_extra_c = df_Proy_EXTRA_c.iloc[:,6].values.tolist()
                tena_extra_c = df_Proy_EXTRA_c.iloc[:,7].values.tolist()
                bani_extra_c = df_Proy_EXTRA_c.iloc[:,8].values.tolist()
                #===> Obtengo únicamente 7 días de diciembre del año pasado
                mess_extra_c = mess_extra_c[-(24*31):]
                #===> Elimino el mes de diciembre del año actual
                mes_c = mes_c[:-(24*31)]
                #===> Nueva lista con 7 días de Dic del año pasado mas resto de datos año actual
                mes_c = mess_extra_c + mes_c
                
                amba_c = amba_extra_c + amba_c
                toto_c = toto_extra_c + toto_c
                puyo_c = puyo_extra_c + puyo_c
                tena_c = tena_extra_c + tena_c
                bani_c = bani_extra_c + bani_c
        
        #===> APUNTA AL MES SELECCIONADO DE COMPARACIÓN
        x_meses_c = [mes_c.index('ENERO'),   mes_c.index('FEBRERO'),   mes_c.index('MARZO'),
                     mes_c.index('ABRIL'),   mes_c.index('MAYO'),      mes_c.index('JUNIO'), 
                     mes_c.index('JULIO'),   mes_c.index('AGOSTO'),    mes_c.index('SEPTIEMBRE'),
                     mes_c.index('OCTUBRE'), mes_c.index('NOVIEMBRE'), mes_c.index('DICIEMBRE')]
        
        for i in range (12):
            if pos_mes_c == i:
                if pos_mes_c == 11:
                    ubic_mes_c = x_meses_c[i]
                else:
                    ubic_mes_c = x_meses_c[i] # VARIABLE Posición del mes seleccionado
                    
        
        #===> DEFINE LA UBICACIÓN DEL DÍA SELECCIONADO DE COMPARACIÓN
        print(' ')
        print ('Fecha Comparación: ',po_dia_c,' de ',mes_c[ubic_mes_c],' de ',po_anio_c)
        if int(po_dia_c) == 1:
            ubicacion_c = ubic_mes_c
        else:
            ubicacion_c = (int(po_dia_c)-1) * 24 + ubic_mes_c
            
    #    print(' ')  
    #    print (ubicacion_c) #==========================================================>>> VARIABLE DE UBICACIÓN EXACTA EN LISTA DE DATOS de comparación
        
        amba_c = amba_c[:ubicacion_c]
        toto_c = toto_c[:ubicacion_c]
        puyo_c = puyo_c[:ubicacion_c]
        tena_c = tena_c[:ubicacion_c]
        bani_c = bani_c[:ubicacion_c]
        
        #===> Valores inicio y fin para lista de datos a ser usada
        val_ini_c = ubicacion_c - 24 * 7 #Posición del valor inicial
        val_fin_c = ubicacion_c - 1      #Posición del valor final
        
        #===> Comparación semananal S/E Ambato
        CompAmba = []
        for i in range (24*7):
            CompAmba.append(amba_c[val_ini_c + i ])  
        #===> Comparación semananal S/E Totoras
        CompToto = []
        for i in range (24*7):
            CompToto.append(toto_c[val_ini_c + i ])
        #===> Comparación semananal S/E Puyo
        CompPuyo = []
        for i in range (24*7):
            CompPuyo.append(puyo_c[val_ini_c + i  ])
        #===> Comparación semananal S/E Tena
        CompTena = []
        for i in range (24*7):
            CompTena.append(tena_c[val_ini_c + i ])
        #===> Comparación semananal S/E Baños
        CompBani = []
        for i in range (24*7):
            CompBani.append(bani_c[val_ini_c + i ])
        
         #===> Cuadro informativo de la semana proyectada y comparativa
            
        x_proy = datetime(int(po_anio_p),(pos_mes_p+1),int(po_dia_p))
        x_comp = datetime(int(po_anio_c),(pos_mes_c+1),int(po_dia_c))
        
        dias_proy = timedelta(days = 28)
        dias_comp = timedelta(days =  7)
        
        fecha_ini_proy = x_proy - dias_proy
        fecha_fin_proy = x_proy - timedelta(days = 1)
        
        fecha_ini_comp = x_comp - dias_comp
        fecha_fin_comp = x_comp - timedelta(days = 1)
        
        sem_ini = x_proy
        sem_fin = x_proy + timedelta(days = 6)
        
        s_i_1 = fecha_ini_proy
        s_i_2 = s_i_1 + timedelta(days = 7)
        s_i_3 = s_i_2 + timedelta(days = 7)
        s_i_4 = s_i_3 + timedelta(days = 7)
        
        s_f_1 = fecha_ini_proy + timedelta(days = 6)
        s_f_2 = s_f_1 + timedelta(days = 7)
        s_f_3 = s_f_2 + timedelta(days = 7)
        s_f_4 = s_f_3 + timedelta(days = 7)
        
        
        #===> Salida de cuadro de texto informativo
        
        messagebox.showinfo(message='Semana de Proyección:\n   del '+str(sem_ini.day)+' / '+
                            str(sem_ini.month) +' / '+str(sem_ini.year)+'\tal '+
                            str(sem_fin.day)+' / '+str(sem_fin.month)+' / '+str(sem_fin.year)+
                            '\n\nSemana de Comparación:\n   del '+ str(fecha_ini_comp.day)+' / '+
                            str(fecha_ini_comp.month)+' / '+str(fecha_ini_comp.year)+'\t al '+
                            str(fecha_fin_comp.day)+' / '+str(fecha_fin_comp.month)+' / '+
                            str(fecha_fin_comp.year)+
                            '\n\nSemanas Promediadas:\n\n   Semana 1 : del '+str(s_i_1.day)+' / '+
                            str(s_i_1.month)+ ' / '+str(s_i_1.year)+' al '+str(s_f_1.day)+ ' / '+
                            str(s_f_1.month)+ ' / '+str(s_f_1.year)+
                            '\n   Semana 2 : del '+str(s_i_2.day)+' / '+str(s_i_2.month)+ ' / '+
                            str(s_i_2.year)+' al '+str(s_f_2.day)+ ' / '+str(s_f_2.month)+ ' / '+
                            str(s_f_2.year)+
                            '\n   Semana 3 : del '+str(s_i_3.day)+' / '+str(s_i_3.month)+ ' / '+
                            str(s_i_3.year)+' al '+str(s_f_3.day)+ ' / '+str(s_f_3.month)+ ' / '+
                            str(s_f_3.year)+
                            '\n   Semana 4 : del '+str(s_i_4.day)+' / '+str(s_i_4.month)+ ' / '+
                            str(s_i_4.year)+' al '+str(s_f_4.day)+ ' / '+str(s_f_4.month)+ ' / '+
                            str(s_f_4.year)
                            ,title="Información")
        
        #==============================================================================
        #======================== RUTINA GENERA EXCEL =================================
        #==============================================================================
        df = pd.DataFrame([' '])
        writer = ExcelWriter('01_PROYECCION_PROMEDIOS.xls')
        df.to_excel(writer, 'Salida_Proyección_CECON', index=False)
        df.to_excel(writer, 'Salida_Comparación_CECON', index=False)
        df.to_excel(writer, 'Salida_ERRORES_CECON', index=False)
        writer.save()
        
        #abre el archivo de excel plantilla
        rb = open_workbook('01_PROYECCION_PROMEDIOS.xls')
        #crea una copia del archivo plantilla
        wb = copy(rb)
        #se ingresa a la hoja 1 de la copia del archivo excel
        ws = wb.get_sheet(0)
        ws_c = wb.get_sheet(1)
        ws_e = wb.get_sheet(2)
        #===========================Hoja 1 Proyección =================================
        #ws.write(0,0,'MES')
        #ws.write(0,1,'DÍA')
        #ws.write(0,2,'#')
        ws.write(0,0,'METODOLOGÍA DE PROMEDIOS  (A.R.I.M.A.)')
        ws.write(1,0,'PROYECCIÓN SEMANAL DE CARGA')
        ws.write(2,0,'Semana de Proyección: del '+str(sem_ini.day)+'/'
                 +str(sem_ini.month) +'/'+str(sem_ini.year)+'\t al '+
                 str(sem_fin.day)+'/'+str(sem_fin.month)+'/'+str(sem_fin.year))
        
        ws.write(3,0,'Realizado por: '+str(po_Realizado))
        
        # Define lista de títulos
        titulos = ['HORA','S/E AMBATO','S/E TOTORAS','S/E PUYO','S/E TENA','S/E BAÑOS']
        
        for i in range(len(titulos)):
            ws.write(9,i,titulos[i])
            
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws.write(Aux,0,j+1)
                ws.write(Aux,1,ProyAmba[j + Aux2])
                ws.write(Aux,2,ProyToto[j + Aux2])
                ws.write(Aux,3,ProyPuyo[j + Aux2])
                ws.write(Aux,4,ProyTena[j + Aux2])
                ws.write(Aux,5,ProyBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
        #===========================Hoja 2 Comparación ================================
        
        #===> Fecha
        ws_c.write(0,0,'Semana de Proyección: del '+str(sem_ini.day)+'/'
                   +str(sem_ini.month) +'/'+str(sem_ini.year)+'\t al '+
                   str(sem_fin.day)+'/'+str(sem_fin.month)+'/'+str(sem_fin.year))
        ws_c.write(2,0,'PROYECCIÓN SEMANAL')
        
        for i in range(len(titulos)):
            ws_c.write(9,i,titulos[i])
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,0,j+1)
                ws_c.write(Aux,1,ProyAmba[j + Aux2])
                ws_c.write(Aux,2,ProyToto[j + Aux2])
                ws_c.write(Aux,3,ProyPuyo[j + Aux2])
                ws_c.write(Aux,4,ProyTena[j + Aux2])
                ws_c.write(Aux,5,ProyBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
            
        #ws_c.write(0,6,po_dia_c+'/'+po_mes_c+'/'+po_anio_c)
        ws_c.write(2,6,'SEMANA DE COMPARACIÓN')
        
        for i in range(len(titulos)):
            ws_c.write(9,i+6,titulos[i])
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,6,j+1)
                ws_c.write(Aux,7,CompAmba[j + Aux2])
                ws_c.write(Aux,8,CompToto[j + Aux2])
                ws_c.write(Aux,9,CompPuyo[j + Aux2])
                ws_c.write(Aux,10,CompTena[j + Aux2])
                ws_c.write(Aux,11,CompBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
            
            
        # ==========================   Cálculo de errores    ==========================
            
        ws_c.write(2,12,'CÁLCULO DE ERRORES')
        ws_c.write(3,12,'PORCENTAJE DE ERROR MEDIO ABSOLUTO (PEMA)')
        
        for i in range(len(titulos)):
            ws_c.write(9,i+12,titulos[i])
        
        
        errAmba = []
        errToto = []
        errPuyo = []
        errTena = []
        errBani = []
        
        for i in range(len(ProyAmba)):
            errAmba.append((abs( ProyAmba[i] - CompAmba[i] ) / CompAmba[i] )*100)
            errToto.append((abs( ProyToto[i] - CompToto[i] ) / CompToto[i] )*100)
            errPuyo.append((abs( ProyPuyo[i] - CompPuyo[i] ) / CompPuyo[i] )*100)
            errTena.append((abs( ProyTena[i] - CompTena[i] ) / CompTena[i] )*100)
            errBani.append((abs( ProyBani[i] - CompBani[i] ) / CompBani[i] )*100)
        
        Aux = 10
        Aux2 = 0
        for i in range (7):
            for j in range (24):
                ws_c.write(Aux,12,j+1)
                ws_c.write(Aux,13,errAmba[j + Aux2])
                ws_c.write(Aux,14,errToto[j + Aux2])
                ws_c.write(Aux,15,errPuyo[j + Aux2])
                ws_c.write(Aux,16,errTena[j + Aux2])
                ws_c.write(Aux,17,errBani[j + Aux2])
                Aux = Aux + 1
            Aux2 = 24 * (i+1)
            Aux = Aux + 1
        # Suma de los valores de las listas
        SumAmba = 0
        SumToto = 0
        SumPuyo = 0
        SumTena = 0
        SumBani = 0
        for i in range (len(ProyAmba)):
            SumAmba = errAmba[i] + SumAmba
            SumToto = errToto[i] + SumToto
            SumPuyo = errPuyo[i] + SumPuyo
            SumTena = errTena[i] + SumTena
            SumBani = errBani[i] + SumBani
        # Almaceno en una lista los resultados de las sumas
        Sumas = [SumAmba, SumToto, SumPuyo, SumTena, SumBani]
        
        # Imprime los resultados en la fila correspondiente
        ws_c.write(Aux+1,11,'SUMATORIA TOTAL')
        for i in range (len(Sumas)):
             ws_c.write(Aux+1,i+13,Sumas[i])     
        
        # Cálculo de los promedios de las sumas
        ws_c.write(Aux+2,11,'ERROR MEDIO ABSOLUTO')
        for i in range(len(Sumas)):
            ws_c.write(Aux+2,i+13,(Sumas[i]/len(errAmba)))
        
        # Cálculo de la exactitud de la proyección
        ws_c.write(Aux+3,11,'EXACTITUD DE LA PROYECCIÓN')
        for i in range(len(Sumas)):
            ws_c.write(Aux+3,i+13,(100-(Sumas[i]/len(errAmba))))
        
        
        for i in range(len(titulos)-1):
            ws_c.write(Aux,i+13,titulos[i+1])
        
        # ==========================   SOLO   errores    ==========================
        
        # ==================================> cálculo de et = Comp - Proy
        for i in range(len(titulos)-1):
            ws_e.write(9,i,titulos[i+1])
        ws_e.write(7,0,'et = Comp - Proy')
        
        et_Amba = []
        et_Toto = []
        et_Puyo = []
        et_Tena = []
        et_Bani = []
        
        for i in range (len(ProyAmba)):
            et_Amba.append(CompAmba[i] - ProyAmba[i])
            et_Toto.append(CompToto[i] - ProyToto[i])
            et_Puyo.append(CompPuyo[i] - ProyPuyo[i])
            et_Tena.append(CompTena[i] - ProyTena[i])
            et_Bani.append(CompBani[i] - ProyBani[i])
            
            ws_e.write(i+10,0, (et_Amba[i]))
            ws_e.write(i+10,1, (et_Toto[i]))
            ws_e.write(i+10,2, (et_Puyo[i]))
            ws_e.write(i+10,3, (et_Tena[i]))
            ws_e.write(i+10,4, (et_Bani[i]))
        
        # ==================================> cálculo de abs(et) = abs(Comp - Proy)
        for i in range(len(titulos)-1):
            ws_e.write(9,i+6,titulos[i+1])
        ws_e.write(7,6,'abs(et) = abs(Comp - Proy)')
        
        abs_et_Amba = []
        abs_et_Toto = []
        abs_et_Puyo = []
        abs_et_Tena = []
        abs_et_Bani = []
        
        for i in range (len(ProyAmba)):
            abs_et_Amba.append(abs(et_Amba[i]))
            abs_et_Toto.append(abs(et_Toto[i]))
            abs_et_Puyo.append(abs(et_Puyo[i]))
            abs_et_Tena.append(abs(et_Tena[i]))
            abs_et_Bani.append(abs(et_Bani[i]))
            
            ws_e.write(i+10,6,  (abs_et_Amba[i]))
            ws_e.write(i+10,7,  (abs_et_Toto[i]))
            ws_e.write(i+10,8,  (abs_et_Puyo[i]))
            ws_e.write(i+10,9,  (abs_et_Tena[i]))
            ws_e.write(i+10,10, (abs_et_Bani[i]))
            
        # ==================================> cálculo de et^2
        for i in range(len(titulos)-1):
            ws_e.write(9,i+12,titulos[i+1])
        ws_e.write(7,12,'et^2')
        
        et_Amba2 = []
        et_Toto2 = []
        et_Puyo2 = []
        et_Tena2 = []
        et_Bani2 = []
        
        for i in range (len(ProyAmba)):
            et_Amba2.append((et_Amba[i])**2)
            et_Toto2.append((et_Toto[i])**2)
            et_Puyo2.append((et_Puyo[i])**2)
            et_Tena2.append((et_Tena[i])**2)
            et_Bani2.append((et_Bani[i])**2)
            
            ws_e.write(i+10,12, (et_Amba2[i]))
            ws_e.write(i+10,13, (et_Toto2[i]))
            ws_e.write(i+10,14, (et_Puyo2[i]))
            ws_e.write(i+10,15, (et_Tena2[i]))
            ws_e.write(i+10,16, (et_Bani2[i]))
            
        # ==================================> cálculo de abs(et) / Comp
        for i in range(len(titulos)-1):
            ws_e.write(9,i+18,titulos[i+1])
        ws_e.write(7,18,'abs(et) / Comp')
        
        d1_Amba = []
        d1_Toto = []
        d1_Puyo = []
        d1_Tena = []
        d1_Bani = []
        
        for i in range (len(ProyAmba)):
            d1_Amba.append(abs_et_Amba[i] / CompAmba[i])
            d1_Toto.append(abs_et_Toto[i] / CompToto[i])
            d1_Puyo.append(abs_et_Puyo[i] / CompPuyo[i])
            d1_Tena.append(abs_et_Tena[i] / CompTena[i])
            d1_Bani.append(abs_et_Bani[i] / CompBani[i])
            
            ws_e.write(i+10,18, (d1_Amba[i]))
            ws_e.write(i+10,19, (d1_Toto[i]))
            ws_e.write(i+10,20, (d1_Puyo[i]))
            ws_e.write(i+10,21, (d1_Tena[i]))
            ws_e.write(i+10,22, (d1_Bani[i]))   
        
        # ==================================> cálculo de et / Comp
        for i in range(len(titulos)-1):
            ws_e.write(9,i+24,titulos[i+1])
        ws_e.write(7,24,'et / Comp')
        
        d2_Amba = []
        d2_Toto = []
        d2_Puyo = []
        d2_Tena = []
        d2_Bani = []
        
        for i in range (len(ProyAmba)):
            d2_Amba.append(et_Amba[i] / CompAmba[i])
            d2_Toto.append(et_Toto[i] / CompToto[i])
            d2_Puyo.append(et_Puyo[i] / CompPuyo[i])
            d2_Tena.append(et_Tena[i] / CompTena[i])
            d2_Bani.append(et_Bani[i] / CompBani[i])
            
            ws_e.write(i+10,24, (d2_Amba[i]))
            ws_e.write(i+10,25, (d2_Toto[i]))
            ws_e.write(i+10,26, (d2_Puyo[i]))
            ws_e.write(i+10,27, (d2_Tena[i]))
            ws_e.write(i+10,28, (d2_Bani[i]))   
            
        ws_e.write(0,0, 'INDICADORES') 
        # ==================================> Cálculo DAM
        DAM_Amba = 0
        DAM_Toto = 0
        DAM_Puyo = 0
        DAM_Tena = 0
        DAM_Bani = 0
        
        for i in range (len(ProyAmba)):
            DAM_Amba = abs_et_Amba[i] + DAM_Amba
            DAM_Toto = abs_et_Toto[i] + DAM_Toto
            DAM_Puyo = abs_et_Puyo[i] + DAM_Puyo
            DAM_Tena = abs_et_Tena[i] + DAM_Tena
            DAM_Bani = abs_et_Bani[i] + DAM_Bani
            
        DAM_Amba = DAM_Amba / (len(ProyAmba))
        DAM_Toto = DAM_Toto / (len(ProyAmba))
        DAM_Puyo = DAM_Puyo / (len(ProyAmba))
        DAM_Tena = DAM_Tena / (len(ProyAmba))
        DAM_Bani = DAM_Bani / (len(ProyAmba))
        
        ws_e.write(1,0, ('DAM'))
        DAM = [DAM_Amba,DAM_Toto,DAM_Puyo,DAM_Tena,DAM_Bani]
        for i in range(len(DAM)):
            ws_e.write(1,i+1,DAM[i])
    
        
        # ==================================> Cálculo EMC 
        EMC_Amba = 0
        EMC_Toto = 0
        EMC_Puyo = 0
        EMC_Tena = 0
        EMC_Bani = 0
        
        for i in range (len(ProyAmba)):
            EMC_Amba = et_Amba2[i] + EMC_Amba
            EMC_Toto = et_Toto2[i] + EMC_Toto
            EMC_Puyo = et_Puyo2[i] + EMC_Puyo
            EMC_Tena = et_Tena2[i] + EMC_Tena
            EMC_Bani = et_Bani2[i] + EMC_Bani
            
        EMC_Amba = EMC_Amba / (len(ProyAmba))
        EMC_Toto = EMC_Toto / (len(ProyAmba))
        EMC_Puyo = EMC_Puyo / (len(ProyAmba))
        EMC_Tena = EMC_Tena / (len(ProyAmba))
        EMC_Bani = EMC_Bani / (len(ProyAmba))
        
        ws_e.write(2,0, ('EMC'))
        EMC = [EMC_Amba,EMC_Toto,EMC_Puyo,EMC_Tena,EMC_Bani]
        
        for i in range(len(EMC)):
            ws_e.write(2,i+1,EMC[i])
            
        # ==================================> Cálculo PEMA
        PEMA_Amba = 0
        PEMA_Toto = 0
        PEMA_Puyo = 0
        PEMA_Tena = 0
        PEMA_Bani = 0
        
        for i in range (len(ProyAmba)):
            PEMA_Amba = (abs_et_Amba[i] / CompAmba [i]) + PEMA_Amba
            PEMA_Toto = (abs_et_Toto[i] / CompToto [i]) + PEMA_Toto
            PEMA_Puyo = (abs_et_Puyo[i] / CompPuyo [i]) + PEMA_Puyo
            PEMA_Tena = (abs_et_Tena[i] / CompTena [i]) + PEMA_Tena
            PEMA_Bani = (abs_et_Bani[i] / CompBani [i]) + PEMA_Bani
            
        PEMA_Amba = (PEMA_Amba / (len(ProyAmba))) *100
        PEMA_Toto = (PEMA_Toto / (len(ProyAmba))) *100
        PEMA_Puyo = (PEMA_Puyo / (len(ProyAmba))) *100
        PEMA_Tena = (PEMA_Tena / (len(ProyAmba))) *100
        PEMA_Bani = (PEMA_Bani / (len(ProyAmba))) *100
        
        ws_e.write(3,0, ('PEMA'))
        PEMA = [PEMA_Amba,PEMA_Toto,PEMA_Puyo,PEMA_Tena,PEMA_Bani]
        
        for i in range(len(PEMA)):
            ws_e.write(3,i+1,PEMA[i])
            
        # ==================================> Cálculo PME
        PME_Amba = 0
        PME_Toto = 0
        PME_Puyo = 0
        PME_Tena = 0
        PME_Bani = 0
        
        for i in range (len(ProyAmba)):
            PME_Amba = (et_Amba[i] / CompAmba [i]) + PME_Amba
            PME_Toto = (et_Toto[i] / CompToto [i]) + PME_Toto
            PME_Puyo = (et_Puyo[i] / CompPuyo [i]) + PME_Puyo
            PME_Tena = (et_Tena[i] / CompTena [i]) + PME_Tena
            PME_Bani = (et_Bani[i] / CompBani [i]) + PME_Bani
            
        PME_Amba = (PME_Amba / (len(ProyAmba))) *100
        PME_Toto = (PME_Toto / (len(ProyAmba))) *100
        PME_Puyo = (PME_Puyo / (len(ProyAmba))) *100
        PME_Tena = (PME_Tena / (len(ProyAmba))) *100
        PME_Bani = (PME_Bani / (len(ProyAmba))) *100
        
        ws_e.write(4,0, ('PME'))
        PME = [PME_Amba,PME_Toto,PME_Puyo,PME_Tena,PME_Bani]
        
        for i in range(len(PME)):
            ws_e.write(4,i+1,PME[i])
            
        # ==================================> Cálculo EXACTITUD DE LA PROYECCION
        ws_e.write(5,0,'EP')
        for i in range(len(PEMA)):
            ws_e.write(5,i+1,(100-PEMA[i]))
        
        
        for i in range(len(titulos)-1):
            ws_e.write(0,i+1,titulos[i+1])
            
        global direccion_resultados
        wb.save(direccion_resultados + '\\01_PROYECCION_PROMEDIOS.xls')
        
        #==============================================================================
        #================================ GRÁFICAS ====================================
        #==============================================================================
        
        if po_Comparacion == 'SI':
            
        #===> Creamos una lista tipo entero para relacionar con las etiquetas
            can_datos = []
            for i in range(7*24):
                if i%2!=1:
                    can_datos.append(i)
    #        print(len(can_datos))  
            
        #===> Creamos una lista con los números de horas del día seleccionado (etiquetas)
            
            horas_dia = []
            horas_str = []
            for i in range (7):
                for i in range (1,25):
                    if i%2!=0:
                        horas_dia.append(i)
            for i in range (len(horas_dia)):
                horas_str.append(str(horas_dia[i]))
                
        #===> Tamaño de la ventana de la gráfica
            plt.subplots(figsize=(15, 8))
            
        #===> Título general superior
            plt.suptitle(u' PROYECCIÓN SEMANAL DE CARGA\nMETODOLOGÍA DE PROMEDIOS ',fontsize=14, fontweight='bold') 
            
            plt.subplot(5,1,1)
            plt.plot(CompAmba,'blue', label = 'Comparación')
            plt.plot(ProyAmba,'#DCD037', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E AMBATO\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,2)
            plt.plot(CompToto,'blue', label = 'Comparación')
            plt.plot(ProyToto,'#CD336F', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E TOTORAS\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,3)  
            plt.plot(CompPuyo,'blue', label = 'Comparación')
            plt.plot(ProyPuyo,'#349A9D', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E PUYO\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,4)  
            plt.plot(CompTena,'blue', label = 'Comparación')
            plt.plot(ProyTena,'#CC8634', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E TENA\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
            
            plt.subplot(5,1,5)  
            plt.plot(CompBani,'blue', label = 'Comparación')
            plt.plot(ProyBani,'#4ACB71', label = 'Proyección')
            plt.legend(loc='upper left')
            plt.xlabel(u'TIEMPO [ días ]')
            plt.ylabel(u'S/E BAÑOS\n\nCARGA  [ kW ]') 
            plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
            plt.xticks(can_datos, horas_str, size = 5, color = 'b', rotation = 90)
                     
            
                 
       
        
        #==============================================================================
        #==============================================================================
        
        print(' ')
        print(' ')
        print('*** Archivo Excel de Proyección semanal ha sido generado')
        print(' ')
        print('*** Completado !...')
        
    start_time = time()
    test()
    elapsed_time = time() - start_time
    print(' ')
    print(' ')
    print('/////////////////////////////////////////')
    print(' ')
    print('ALGORITMO UTILIZADO: ARIMA')
    print("Tiempo transcurrido: %.10f segundos." % elapsed_time)            
    print(' ') 
    print('/////////////////////////////////////////')




























def graficas_rapidas():
    
    # -*- coding: utf-8 -*-
    """
    Created on Thu Feb 21 08:38:10 2019
    ESCUELA POLITÉCNICA NACIONAL
    AUTOR: ALEXIS GUAMAN FIGUEROA
    1_MÓDULO DE GRÁFICAS RÁPIDAS
    """    
    print (' ')
    print (' ')
    print ('**  ANALIZADOR GRÁFICO... ')
    print ('**  Compilando... ')
    print (' ')
    print (' ')
    
    #---------- PARÁMETROS INICIALES, CREACIÓN DE LA VENTANA RAIZ -------------
    
    raiz = Tk()
    raiz.geometry('600x520+100+100')
    raiz.title('Análisis de datos')
    raiz.resizable(0,0)
    raiz.iconbitmap('epn.ico')

    def infoAdicional():
        messagebox.showinfo('Acerca de este programa','PROYECTO PREVIO A LA OBTENCIÓN DEL TÍTULO DE INGENIERO ELÉCTRICO\n\nEste Programa fue realizado por Alexis Guamán Figueroa\n\n\n\nVersión BETA 1.0')
    
    #===> Ejecuta proyección semanal
    def proySem(): 
        raiz.destroy()
        proyeccion_semanal()

    #===> Cierra el proyecto
    def salirProyecto():
        PregCierre = messagebox.askquestion('Cerrar programa','Desea cerrar el programa ?')
        if PregCierre == 'yes':
            raiz.destroy()
    
    #=========================SUBRUTINA PARA RESET ============================
    def boton_reset():
        raiz.destroy()    
        graficas_rapidas()
    b_res = Button(raiz, text = '  RESET  ',bg='#A81C1C',fg='#FFFFFF',command = boton_reset)
    b_res.place(x=445,y=445)
    
    #==========================================================================
    #============================= Menú de la raíz ============================
    #==========================================================================
    
    barraMenu = Menu(raiz)
    raiz.config(menu=barraMenu)
    
    archivoMenu = Menu(barraMenu, tearoff = 0)
    
    archivoMenu.add_command(label = 'Nuevo')
    archivoMenu.add_command(label = 'Guardar')
    archivoMenu.add_command(label = 'Guardar Como')
    archivoMenu.add_separator()
    archivoMenu.add_command(label = 'Nueva Proyección Semanal',command = proySem)
    archivoMenu.add_separator()
    archivoMenu.add_command(label = 'Reiniciar', command = boton_reset)
    archivoMenu.add_command(label = 'Salir', command = salirProyecto)
    
    archivoEdicion = Menu(barraMenu, tearoff = 0)
    archivoEdicion.add_command(label = 'Copiar')
    archivoEdicion.add_command(label = 'Cortar')
    archivoEdicion.add_command(label = 'Pegar')
    
#    archivoConfiguraciones = Menu(barraMenu, tearoff = 0)
#    archivoConfiguraciones.add_command(label = 'Configuraciones por defecto')
#    archivoConfiguraciones.add_separator()
#    archivoConfiguraciones.add_command(label = 'Definir Base de Datos', command = base_datos)
    
    archivoAyuda = Menu(barraMenu, tearoff = 0)
    archivoAyuda.add_command(label = 'Acerca de...', command = infoAdicional)
    
    barraMenu.add_cascade(label = 'Archivo', menu = archivoMenu)
    barraMenu.add_cascade(label = 'Edición', menu = archivoEdicion)
#    barraMenu.add_cascade(label = 'Configuraciones', menu = archivoConfiguraciones)
    barraMenu.add_cascade(label = 'Ayuda', menu = archivoAyuda)
    
    #==============================================================================
    #==============================================================================
    
    
    
    #------------------------- Ingreso de imágenes --------------------------------
    
    imag_eeasa = PhotoImage(file='eeasa_peq.png')
    Label(raiz, image = imag_eeasa).place(x=140,y=20) 
     
    imag_pronost = PhotoImage(file='e_pronostico_peq.png')       
    Label(raiz, image = imag_pronost).place(x=10,y=390) 
    
    
    #------------------------- Ingreso de textos ----------------------------------
    
    txq_1 = Label(raiz, text = 'Seleccione el tipo de gráficas a generar: ', fg='#616A6B',font=('Arial',12)).place(x=20,y=150)  
    Label(raiz, text = 'AGF ®', fg='#616A6B',font=('Tw Cen MT',9)).place(x=545,y=475) 
    
    #-------------------------Ingreso de MENÚ 1------------------------------------
    
    cbx = ttk.Combobox(values=['RÁPIDA','ACUMULATIVA'],state="readonly")
    cbx.place(x=330,y=150)

    #------------------------- Ingreso de BOTON CHECK 1 ---------------------------
    
    b1 = Button(raiz, text = '  Ok  ',command = lambda:boton_1())
    b1.place(x=500,y=150)
    
    #------------------------- LISTA DE FUNCIONES ----------------------
    
    #-----------------------Función 1, PARA EL BOTON CHECK 1----------------------
    
    def boton_1():
        
        pos_menu_1 = cbx.current( )

        if pos_menu_1 == 0:
            txt_2 = Label(raiz, text = 'GRÁFICA RÁPIDA', fg='#616A6B',font=('Arial',12)).place(x=20,y=200)
            txt_3 = Label(raiz, text = 'Seleccione el tipo de gráfica: ', fg='#616A6B',font=('Arial',10)).place(x=20,y=230)
            
            def boton_2():
                global pos_menu_2
                
                pos_menu_2 = cbx_tipograf.current()
                po_menu_2 = cbx_tipograf.get()
                
                varOpcion = IntVar()
                
                if pos_menu_2 == -1:
                    messagebox.showinfo(message="  Debe seleccionar el tipo de gráficas a generar ! ", title="Faltan datos")
                    print('  Debe seleccionar el tipo de gráficas a generar ! ')
                else:
                    print('**  Gráfica del tipo: ',po_menu_2)                                                  
                
                
                
                def imp1():
                    global val_check
                    val_check = varOpcion.get()
                    if val_check == 1:
                        print('**  Ha seleccionado SCATTER')
                    else:
                        print('**  Sin SCATTER')
    
    #==============================================================================                
    #================================= GRÁFICA ANUAL ==============================
    #==============================================================================
                if   pos_menu_2==0:
                    b2.place_forget()
                    
                    Label(raiz, text = 'AÑO: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=230)
#                    cbx_anio = ttk.Combobox(values=['2013','2014','2015','2016','2017','2018'],state="readonly")
                    cbx_anio = ttk.Combobox(values=anios_disponibles,state="readonly")
                    cbx_anio.place(x=400,y=230)
                    
                    Label(raiz, text = 'S / E: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=260)
                    cbx_se = ttk.Combobox(values=['AMBATO','BAÑOS','PUYO','TENA','TOTORAS'],state="readonly")
                    cbx_se.place(x=400,y=260)
                    
                    #Radiobutton(raiz,text = 'GENERAR SCATTER', fg='#616A6B',font=('Arial',10), variable = varOpcion, value=1, command=imp1).place(x=350,y=290)
                    Checkbutton(raiz, text = 'GENERAR SCATTER', fg='#616A6B',font=('Arial',10), variable = varOpcion, onvalue = 1, offvalue = 0, command=imp1).place(x=350,y=290)
                                
                    b_go = Button(raiz, text = '  GRAFICAR  ',bg='#2ECC71',fg='#1A5276',command = lambda:boton_go())
                    b_go.place(x=505,y=445)
                    
                    def boton_go():
                        
                        global pos_se, po_anio, po_se
    
                        pos_anio = cbx_anio.current()
                        pos_se  = cbx_se.current()
                        
                        po_anio = cbx_anio.get()
                        po_se  = cbx_se.get()
                        
                        if pos_anio == -1:
                            messagebox.showinfo(message="Ingrese Año", title="Faltan datos")
                        elif pos_se == -1:
                            messagebox.showinfo(message="Ingrese la Subestación", title="Faltan datos")
                        else:
                            print ('****    Se ha seleccionado el año', po_anio, 'Para la subestación', po_se)
                            graficas_rapidas_2()
    
                          
    #==============================================================================                
    #================================= GRÁFICA MENSUAL ============================
    #==============================================================================                           
                elif pos_menu_2==1:
                    b2.place_forget()
                    
                    Label(raiz, text = 'AÑO: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=230)
#                    cbx_anio = ttk.Combobox(values=['2013','2014','2015','2016','2017','2018'],state="readonly")
                    cbx_anio = ttk.Combobox(values=anios_disponibles,state="readonly")
                    cbx_anio.place(x=400,y=230)
                    
                    Label(raiz, text = 'MES: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=260)
                    cbx_mes = ttk.Combobox(values=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'],state="readonly")
                    cbx_mes.place(x=400,y=260)
                    
                    Label(raiz, text = 'S / E: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=290)
                    cbx_se = ttk.Combobox(values=['AMBATO','BAÑOS','PUYO','TENA','TOTORAS'],state="readonly")
                    cbx_se.place(x=400,y=290)
                    
                    Checkbutton(raiz, text = 'GENERAR SCATTER', fg='#616A6B',font=('Arial',10), variable = varOpcion, onvalue = 1, offvalue = 0, command=imp1).place(x=350,y=320)
                     
                    b_go = Button(raiz, text = '  GRAFICAR  ',bg='#2ECC71',fg='#1A5276',command = lambda:boton_go())
                    b_go.place(x=505,y=445)
                    
                    def boton_go():
                        
                        global pos_se, pos_mes, po_anio, po_se, po_mes
    
                        pos_anio = cbx_anio.current()
                        pos_mes = cbx_mes.current()
                        pos_se  = cbx_se.current()
                        
                        po_anio = cbx_anio.get()
                        po_mes = cbx_mes.get()
                        po_se  = cbx_se.get()
                        
                        if pos_anio == -1:
                            messagebox.showinfo(message="Ingrese Año", title="Faltan datos")
                        elif pos_mes == -1:
                            messagebox.showinfo(message="Ingrese Mes", title="Faltan datos")
                        elif pos_se == -1:
                            messagebox.showinfo(message="Ingrese la Subestación", title="Faltan datos")
                        else:
                            print ('****    Se ha seleccionado el año', po_anio, ',con mes', po_mes,'para la subestación',po_se)
                            graficas_rapidas_2()
                            
    #==============================================================================                
    #================================= GRÁFICA SEMANAL ============================
    #==============================================================================                       
                elif pos_menu_2==2:
                    b2.place_forget()
                    
                    Label(raiz, text = 'AÑO: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=230)
#                    cbx_anio = ttk.Combobox(values=['2013','2014','2015','2016','2017','2018'],state="readonly")
                    cbx_anio = ttk.Combobox(values=anios_disponibles,state="readonly")
                    cbx_anio.place(x=400,y=230)
                    
                    Label(raiz, text = 'MES: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=260)
                    cbx_mes = ttk.Combobox(values=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'],state="readonly")
                    cbx_mes.place(x=400,y=260)
                    
                    Label(raiz, text = 'SEM: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=290)
                    cbx_sem = ttk.Combobox(values=['1','2','3','4','5','6'],state="readonly")
                    cbx_sem.place(x=400,y=290)
                    
                    Label(raiz, text = 'S / E: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=320)
                    cbx_se = ttk.Combobox(values=['AMBATO','BAÑOS','PUYO','TENA','TOTORAS'],state="readonly")
                    cbx_se.place(x=400,y=320)
                    
                    Checkbutton(raiz, text = 'GENERAR SCATTER', fg='#616A6B',font=('Arial',10), variable = varOpcion, onvalue = 1, offvalue = 0, command=imp1).place(x=350,y=350)
                     
                    b_go = Button(raiz, text = '  GRAFICAR  ',bg='#2ECC71',fg='#1A5276',command = lambda:boton_go())
                    b_go.place(x=505,y=445)
                    
                    def boton_go():
                        
                        global pos_se, pos_mes, pos_sem, po_anio, po_mes, po_sem, po_se
                        
                        
                        pos_anio = cbx_anio.current()
                        pos_mes = cbx_mes.current()
                        pos_sem = cbx_sem.current()
                        pos_se  = cbx_se.current()
                        
                        po_anio = cbx_anio.get()
                        po_mes = cbx_mes.get()
                        po_sem = cbx_sem.get()
                        po_se  = cbx_se.get()
                        
                        def verif_semana():
                            import calendar
                            num_semanas = len(calendar.monthcalendar(int(po_anio),pos_mes+1))

                            if int(po_sem) <= num_semanas:
                                
                                if (calendar.monthcalendar(int(po_anio),pos_mes+1))[pos_sem][0] == 0:
                                    messagebox.showinfo(message="La semana "+po_sem+' para el mes ' +po_mes + ' de '+po_anio + ' No posee suficientes días para el análisis', title="Seleccione otra semana")
                       
                                elif (calendar.monthcalendar(int(po_anio),pos_mes+1))[pos_sem][-1] == 0:
                                    messagebox.showinfo(message="La semana "+po_sem+' para el mes ' +po_mes + ' de '+po_anio + ' No posee suficientes días para el análisis', title="Seleccione otra semana")
        
                                else:
                                    global dias_restantes
                                    inicio_semana = 0
                                    for j in range (7):
                                        if calendar.monthcalendar(int(po_anio),pos_mes+1)[0][j] == 0:
                                            inicio_semana = inicio_semana + 1
                                    dias_restantes = 7 - inicio_semana
                                    print('días restantes son: ',dias_restantes)
                                    graficas_rapidas_2()
                                    
                                
                            else:
                                messagebox.showinfo(message="No existe semana "+po_sem+' para el mes ' +po_mes + ' de '+po_anio, title="Seleccione otra semana")
                       
                        
                        if pos_anio == -1:
                            messagebox.showinfo(message="Ingrese Año", title="Faltan datos")
                        elif pos_mes == -1:
                            messagebox.showinfo(message="Ingrese Mes", title="Faltan datos")
                        elif pos_sem == -1:
                            messagebox.showinfo(message="Ingrese Semana", title="Faltan datos")
                        elif pos_se == -1:
                            messagebox.showinfo(message="Ingrese la Subestación", title="Faltan datos")
                        else:
                            print ('****    Se ha seleccionado el año', po_anio, ',con mes', po_mes, ',semana N°', po_sem,'para la subestación',po_se)
                            verif_semana()
                            
                        
    #==============================================================================                
    #================================= GRÁFICA DIARIA =============================
    #============================================================================== 
                
                elif pos_menu_2==3:
                    b2.place_forget()
                    
                    Label(raiz, text = 'AÑO: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=230)
#                    cbx_anio = ttk.Combobox(values=['2013','2014','2015','2016','2017','2018'],state="readonly")
                    cbx_anio = ttk.Combobox(values=anios_disponibles,state="readonly")
                    cbx_anio.place(x=400,y=230)
                    
                    Label(raiz, text = 'MES: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=260)
                    cbx_mes = ttk.Combobox(values=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'],state="readonly")
                    cbx_mes.place(x=400,y=260)
                    
                    Label(raiz, text = 'DÍA: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=290)
                    cbx_dia = ttk.Combobox(values=['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31'],state="readonly")
                    cbx_dia.place(x=400,y=290)
                    
                    Label(raiz, text = 'S / E: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=320)
                    cbx_se = ttk.Combobox(values=['AMBATO','BAÑOS','PUYO','TENA','TOTORAS'],state="readonly")
                    cbx_se.place(x=400,y=320)
                    
                    Checkbutton(raiz, text = 'GENERAR SCATTER', fg='#616A6B',font=('Arial',10), variable = varOpcion, onvalue = 1, offvalue = 0, command=imp1).place(x=350,y=350)
                     
                    b_go = Button(raiz, text = '  GRAFICAR  ',bg='#2ECC71',fg='#1A5276',command = lambda:boton_go())
                    b_go.place(x=505,y=445)
                    
                    def boton_go():
                        
                        global pos_se, pos_mes, pos_sem, po_anio, po_mes, po_sem, po_se, pos_dia, po_dia
    
                        pos_anio = cbx_anio.current()
                        pos_mes = cbx_mes.current()
                        pos_dia = cbx_dia.current()
                        pos_se  = cbx_se.current()
                        
                        po_anio = cbx_anio.get()
                        po_mes = cbx_mes.get()
                        po_dia = cbx_dia.get()
                        po_se  = cbx_se.get()
                        
                        from calendar import monthrange
                        
                        if pos_anio == -1:
                            messagebox.showinfo(message="Ingrese Año", title="Faltan datos")
                        elif pos_mes == -1:
                            messagebox.showinfo(message="Ingrese Mes", title="Faltan datos")
                        elif pos_dia == -1:
                            messagebox.showinfo(message="Ingrese Día", title="Faltan datos")
                        elif pos_se == -1:
                            messagebox.showinfo(message="Ingrese la Subestación", title="Faltan datos")
                        elif int(po_dia) > monthrange(int(po_anio), pos_mes+1)[1]:
                            messagebox.showinfo(message='No existe día '+po_dia +' para el mes ' +po_mes + ' de '+ po_anio, title="Dato erróneo")
                        else:
                            graficas_rapidas_2()
                            print ('****    Se ha seleccionado el año', po_anio, ',con mes', po_mes, ',día', po_dia,'para la subestación',po_se)

            cbx_tipograf = ttk.Combobox(values=['Anual','Mensual','Semanal','Diario'],state="readonly")
            cbx_tipograf.place(x=200,y=230)
            
    #------------------------- Ingreso de BOTON CHECK 2 ------------------------------------------------------------
    
            b2 = Button(raiz, text = '  Ok  ',command = lambda:boton_2())
            b2.place(x=350,y=230)
    
            b1.place_forget()       #   Elimina el botón pulsado
            print('**  Ha seleccionado gráfica RÁPIDA...')
            
        
        elif  pos_menu_1 == 1:
            txt_2 = Label(raiz, text = 'GRÁFICAS ACUMULTIVAS ', fg='#616A6B',font=('Arial',12)).place(x=20,y=200)
            b1.place_forget()
            print('**  Ha seleccionado gráfica ACUMULATIVA...')
            
    #==============================================================================                
    #================================= GRÁFICA ANUAL ==============================
    #==============================================================================
            
            Label(raiz, text = 'AÑO: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=230)
#            cbx_anio = ttk.Combobox(values=['2013','2014','2015','2016','2017','2018'],state="readonly")
            cbx_anio = ttk.Combobox(values=anios_disponibles,state="readonly")
            cbx_anio.place(x=400,y=230)
            
            Label(raiz, text = 'S / E: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=260)
            cbx_se = ttk.Combobox(values=['AMBATO','BAÑOS','PUYO','TENA','TOTORAS'],state="readonly")
            cbx_se.place(x=400,y=260)
            
                     
            b_go = Button(raiz, text = '  GRAFICAR  ',bg='#2ECC71',fg='#1A5276',command = lambda:boton_go())
            b_go.place(x=505,y=445)
            
            def boton_go():
                
                global pos_se, po_anio, po_se
    
                pos_anio = cbx_anio.current()
                pos_se  = cbx_se.current()
                
                po_anio = cbx_anio.get()
                po_se  = cbx_se.get()
                
                if pos_anio == -1:
                    messagebox.showinfo(message="Ingrese Año", title="Faltan datos")
                elif pos_se == -1:
                    messagebox.showinfo(message="Ingrese la Subestación", title="Faltan datos")
                else:
                    print ('****    Se ha seleccionado el año', po_anio, 'Para la subestación', po_se)
                    graficas_acumuladas()
            
        elif  pos_menu_1 == 2:
            txt_2 = Label(raiz, text = 'GRÁFICAS ACUMULATIVA HISTÓRICA', fg='#616A6B',font=('Arial',12)).place(x=20,y=200)
            b1.place_forget()
            print('**  Ha seleccionado gráfica HISTÓRICA...')
        
        
                
    #==========================================================================               
    #================================= GRÁFICA HISTÓRICA ======================
    #==========================================================================
            
            Label(raiz, text = 'AÑO: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=230)
#                cbx_anio = ttk.Combobox(values=['2013','2014','2015','2016','2017','2018'],state="readonly")
            cbx_anio = ttk.Combobox(values=anios_disponibles,state="readonly")
            cbx_anio.place(x=400,y=230)
            
            Label(raiz, text = 'S / E: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=260)
            cbx_se = ttk.Combobox(values=['AMBATO','BAÑOS','PUYO','TENA','TOTORAS'],state="readonly")
            cbx_se.place(x=400,y=260)
            
                     
            b_go = Button(raiz, text = '  GRAFICAR  ',bg='#2ECC71',fg='#1A5276',command = lambda:boton_go())
            b_go.place(x=505,y=445)
            
            def boton_go():
                
                global pos_se, po_anio, po_se
    
                pos_anio = cbx_anio.current()
                pos_se  = cbx_se.current()
                
                po_anio = cbx_anio.get()
                po_se  = cbx_se.get()
                
                if pos_anio == -1:
                    messagebox.showinfo(message="Ingrese Año", title="Faltan datos")
                if pos_se == -1:
                    messagebox.showinfo(message="Ingrese la Subestación", title="Faltan datos")
                else:
                    print ('****    Se ha seleccionado el año', po_anio, 'Para la subestación', po_se)
                    graficas_acumuladas()
                    
        else:
                print('  Debe seleccionar el tipo de gráficas a generar ! ')
                
                messagebox.showinfo(message="  Debe seleccionar el tipo de gráficas a generar ! ", title="Faltan datos")
             
    #-------------------------------Encierra en la ventana principal a la raiz---------------------------------------------------
    raiz.mainloop()
    
    
    print (' ')
    print (' ')
    print ('**  Módulo de Gráficas Rápidas COMPLETADO...!')
    print (' ')
    print (' ')

    
    
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#******************************************************************************
#==============================================================================    

def proyeccion_semanal():
#==============================================================================    
# INGRESO DE LA DIRECCIÓN DE LA BASE DE DATOS
#==============================================================================    
    def base_datos():
        def base_datos_ok():
            global direccion_base_datos
            
            direccion_base_datos = tx.get() + '\n'
            b_datos.destroy()
            archivo = open("anios.txt", "r")
            anios_disponibles = archivo.readlines()
            anios_disponibles[0] = direccion_base_datos
            archivo = open("anios.txt", "w")
            for i in range (len(anios_disponibles)):
                archivo.write(anios_disponibles[i])
            archivo.close()
            direccion_base_datos = (anios_disponibles.pop(0))[:-1] # Elimina el último carácter (\n)

            print('Se ha establecido la siguiente dirección para importar los datos:')
            print(direccion_base_datos)
        
        b_datos = Toplevel(raiz)
        b_datos.geometry('600x220+100+100')
        b_datos.title('Base de datos')
        b_datos.resizable(0,0)
        b_datos.iconbitmap('epn.ico')
        Label(b_datos, text = 'AGF ®', fg='#616A6B',font=('Tw Cen MT',9)).place(x=550,y=190) 
        Label(b_datos, text = 'Copie y pegue aquí la dirección de la base de datos', fg='#616A6B',font=('Arial',12)).place(x=10,y=10)  
        tx = Entry(b_datos, width=70, fg='#616A6B', bg='#D7DBDD', font=('Arial',11))
        tx.place(x=10,y= 50)
        tx.insert(END, direccion_base_datos )
#        tx.insert(END, 'C:\\Users\\ALEXIS GUAMAN\\Desktop\\BASE DE DATOS')
        Button(b_datos, text = '  OK...!  ',width=20, height=5,bg='#2ECC71',fg='#1A5276',
        command = lambda:base_datos_ok()).place(x=50,y=100)
#==============================================================================
# INGRESO DE LA DIRECCIÓN DE RESULTADOS
#==============================================================================    
    def destino_resultados():
        def result_datos_ok():
            global direccion_resultados
        
            direccion_resultados = tx.get() + '\n'
            r_datos.destroy()
            archivo = open("anios.txt", "r")
            anios_disponibles = archivo.readlines()
            anios_disponibles[1] = direccion_resultados
            archivo = open("anios.txt", "w")
            for i in range (len(anios_disponibles)):
                archivo.write(anios_disponibles[i])
            archivo.close()
            direccion_resultados = (anios_disponibles.pop(1))[:-1] # Elimina el último carácter (\n)

            print('Se ha establecido la siguiente dirección para almacenar los resultados:')
            print(direccion_base_datos)
            
        
        r_datos = Toplevel(raiz)
        r_datos.geometry('600x220+100+100')
        r_datos.title('Dirección Resultados')
        r_datos.resizable(0,0)
        r_datos.iconbitmap('epn.ico')
        Label(r_datos, text = 'AGF ®', fg='#616A6B',font=('Tw Cen MT',9)).place(x=550,y=190) 
        Label(r_datos, text = 'Copie y pegue aquí la dirección de los resultados', fg='#616A6B',font=('Arial',12)).place(x=10,y=10)  
        tx = Entry(r_datos, width=70, fg='#616A6B', bg='#D7DBDD', font=('Arial',11))
        tx.place(x=10,y= 50)
        tx.insert(END, direccion_resultados)
#        tx.insert(END, 'C:\\Users\\ALEXIS GUAMAN\\Desktop')
        Button(r_datos, text = '  OK...!  ',width=20, height=5,bg='#2ECC71',fg='#1A5276',
        command = lambda:result_datos_ok()).place(x=50,y=100)
        




#==============================================================================
    def his_carga_template():
#        import pandas as pd
#        from pandas import ExcelWriter
#        
#        from xlrd import open_workbook
#        from xlutils.copy import copy
        
        
        
        df = pd.DataFrame([' '])       
        writer = ExcelWriter('FINAL_TEMPLATE.xls')
        df.to_excel(writer, 'Hoja1', index=False)
        writer.save()
        
        
        #abre el archivo de excel plantilla
        rb = open_workbook('FINAL_TEMPLATE.xls')
        #crea una copia del archivo plantilla
        wb = copy(rb)
        #se ingresa a la hoja 1 de la copia del archivo excel
        ws = wb.get_sheet(0)
        
        ws.write(0,0,'MES')
        ws.write(0,1,'DÍA')
        ws.write(0,2,'#')
        ws.write(0,3,'HORA')
        ws.write(0,4,'S/E AMBATO')
        ws.write(0,5,'S/E TOTORAS')
        ws.write(0,6,'S/E PUYO')
        ws.write(0,7,'S/E TENA')
        ws.write(0,8,'S/E BAÑOS')
        
        Aux = 1
        
        for i in range (1,32):
            for j in range (1, 25):
                ws.write(Aux,0,'ENERO')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
        
        for i in range (1,29):
            for j in range (1, 25):
                ws.write(Aux,0,'FEBRERO')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
        
        for i in range (1,32):
            for j in range (1, 25):
                ws.write(Aux,0,'MARZO')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
                
        for i in range (1,31):
            for j in range (1, 25):
                ws.write(Aux,0,'ABRIL')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
        
        for i in range (1,32):
            for j in range (1, 25):
                ws.write(Aux,0,'MAYO')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
        
        for i in range (1,31):
            for j in range (1, 25):
                ws.write(Aux,0,'JUNIO')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
        
        for i in range (1,32):
            for j in range (1, 25):
                ws.write(Aux,0,'JULIO')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
                
        for i in range (1,32):
            for j in range (1, 25):
                ws.write(Aux,0,'AGOSTO')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
        
        for i in range (1,31):
            for j in range (1, 25):
                ws.write(Aux,0,'SEPTIEMBRE')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
        
        for i in range (1,32):
            for j in range (1, 25):
                ws.write(Aux,0,'OCTUBRE')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
        
        for i in range (1,31):
            for j in range (1, 25):
                ws.write(Aux,0,'NOVIEMBRE')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
                
        for i in range (1,32):
            for j in range (1, 25):
                ws.write(Aux,0,'DICIEMBRE')
                ws.write(Aux,1,i)
                ws.write(Aux,2,j)
                Aux = Aux + 1
            
            
            
            
        wb.save('FINAL_TEMPLATE.xls')
        
    def his_carga_final():
        
        
#========================================================================================================================
#========================================================================================================================
#========================================================================================================================
#========================================================================================================================
        def anio_ok():
            anio_seleccionado = tx.get() + '\n'
            a_his.destroy()
            anio_seleccionado = anio_seleccionado[:-1] # Elimina el último carácter (\n)
            
            print(' ')
            print(' ')
            print(' ')
            print(' ')
            print('****     Compilando... ')
            print(' ')
            print(' ')
            print(' ')
            print(' ')
            #Introducir el Nombre del doucmento
            name = 'HOJA_TOTAL_'
            #Introducir el año__________*******MODIFICAR EL AÑO SEGÚN HOJA A EVALUAR******
            anio = anio_seleccionado
            #Formato de archivos de excel
            formato = '.xls'
            #nombre del documento a leer
            doc_to_read = name + anio + formato
            
            #Abre y da lectura a la hoja de excel con los datos de las hojas totales
            df = pd.read_excel(doc_to_read, sheetname='Historial_datos',  header=None)
            
            # Se reemplaza datos NaN por cero (0)
            df = df.fillna(0.0)
            
            
            #S/E TOTORAS salida Ambato
            s1 = df.iloc[:,4].values.tolist()
            #S/E TOTORAS salida Montalvo
            s2 = df.iloc[:,8].values.tolist()
            #S/E TOTORAS salida Baños
            s3 = df.iloc[:,12].values.tolist()
            
            #S/E AMBATO salida Ambato 1
            s4 = df.iloc[:,16].values.tolist()
            #S/E AMBATO salida Ambato 2
            s5 = df.iloc[:,20].values.tolist()
            
            #S/E BAÑOS salida Baños 2
            s6 = df.iloc[:,24].values.tolist()
            
            #S/E PUYO salida Puyo
            s7 = df.iloc[:,28].values.tolist()
            
            #S/E TENA salida Tena
            s8 = df.iloc[:,32].values.tolist()
            #S/E TENA salida Tena Norte
            s9 = df.iloc[:,36].values.tolist()
            
            
            #Se eliminan las 3 primeras celdas
            for i in range(3):
                s1.pop(0)
                s2.pop(0)
                s3.pop(0)
                s4.pop(0)
                s5.pop(0)
                s6.pop(0)
                s7.pop(0)
                s8.pop(0)
                s9.pop(0)
            
            #Se inicializa las variablespara las columnas de las sumas
            p1 = []
            p2 = []
            p3 = []
            p4 = []
            p5 = []
            
            # Se realiza las sumas de las columnas de las salidas para obtener el
            # valor total de las subestaciones Transelectric
            
            for i in range(len(s1)):
                p1.append(s1[i]+s2[i]+s3[i])
                p2.append(s4[i]+s5[i])
                p3.append(s6[i])
                p4.append(s7[i])
                p5.append(s8[i]+s9[i])
            
            
            #Se inicializa las variables para almacenar los resultados
            Tot = []
            Amb = []
            Ban = []
            Puy = []
            Ten = []
            
            #Lectura de todos los datos de la columna *****poner 8784 para año bisiesto ****
            for i in range(8760):
                Tot.append(p1[0]+p1[1]+p1[2]+p1[3])
                Amb.append(p2[0]+p2[1]+p2[2]+p2[3])
                Ban.append(p3[0]+p3[1]+p3[2]+p3[3])
                Puy.append(p4[0]+p4[1]+p4[2]+p4[3])
                Ten.append(p5[0]+p5[1]+p5[2]+p5[3])
                for k in range(4):
                    p1.pop(0)
                    p2.pop(0)
                    p3.pop(0)
                    p4.pop(0)
                    p5.pop(0)
            
            #abre el archivo de excel plantilla final
            rb = open_workbook('FINAL_TEMPLATE.xls')
            #crea una copia del archivo plantilla
            wb = copy(rb)
            #se ingresa a la hoja 1 de la copia del archivo excel
            ws = wb.get_sheet(0)
            
            #se realiza la escritura de las celdas en el archivo excel *****poner 8784 para año bisiesto ****
            for i in range (8760):
                #Escritura en Excel
                ws.write(i+1,4,Amb[i])
                ws.write(i+1,5,Tot[i])
                ws.write(i+1,6,Puy[i])
                ws.write(i+1,7,Ten[i])
                ws.write(i+1,8,Ban[i])
            
            
            
            
            wb.save('HIS_POT_'+anio+'.xls')
            
            print(' ')
            print(' ')
            print(' ')
            print(' ')
            print('**** COMPLETED ****')
            print(' ')
            print(' ')
            print(' ')
            print(' ')
                        
        
        a_his = Toplevel(raiz)
        a_his.geometry('270x150+100+100')
        a_his.title('Año de selección')
        a_his.resizable(0,0)
        a_his.iconbitmap('epn.ico')
        Label(a_his, text = 'AGF ®', fg='#616A6B',font=('Tw Cen MT',9)).place(x=220,y=120) 
        Label(a_his, text = 'Ingrese el año a evaluar', fg='#616A6B',font=('Arial',12)).place(x=10,y=10)  
        tx = Entry(a_his, width=6, fg='#616A6B', bg='#D7DBDD', font=('Arial',11))
        tx.place(x=50,y= 50)
#        tx.insert(END, 'Pegue aquí la dirección de la base de datos y presione OK')
        tx.insert(END, '2019')
        Button(a_his, text = '  OK...!  ',width=15, height=3,bg='#2ECC71',fg='#1A5276',
        command = lambda:anio_ok()).place(x=40,y=80)
#========================================================================================================================
#========================================================================================================================
#========================================================================================================================
#========================================================================================================================
        

    
    # -*- coding: utf-8 -*-
    """
    Created on Wed Mar 13 12:17:56 2019
    ESCUELA POLITÉCNICA NACIONAL
    AUTOR: ALEXIS GUAMAN FIGUEROA
    4_MÓDULO DE PROYECCIÓN SEMANAL
    """
    #******************************************************************************
    #***************************GENERADOR DE PROYECCIÓN ***************************
    #******************************************************************************
    
    print (' ')
    print (' ')
    print ('**  Programa de proyección semanal... ')
    print ('**  Compilando...  ')
    print (' ')
    print (' ')
    
#    from tkinter import *
#    import tkinter.ttk as ttk
#    from tkinter import messagebox
    
    #--------------- PARÁMETROS INICIALES, CREACIÓN DE LA VENTANA RAIZ ------------
    
    # Se genera la raiz o ventana principal
    raiz = Tk()
    # Se establece el tamaño y posición de la ventana principal raiz
    raiz.geometry('600x520+100+100') #+600+200
    # TÍTULO DE LA VENTANA PRINCIPAL
    raiz.title('Proyección Semanal')
    # impide redimensionar la ventana principal raiz
    raiz.resizable(0,0)
    # Se estable un ícono para la ventana principal
    raiz.iconbitmap('epn.ico')

    
    
    #===>Texto de ayuda (Acerca de)
    def infoAdicional():
        messagebox.showinfo('Acerca de este programa','PROYECTO PREVIO A LA OBTENCIÓN DEL TÍTULO DE INGENIERO ELÉCTRICO\n\nEste Programa fue realizado por Alexis Guamán Figueroa\n\n\n\nVersión BETA 1.0')
    #===>Ejecuta archivo de proyección semanal
    def proySem(): 
        raiz.destroy() 
        graficas_rapidas()
        
    #===>Cierra el proyecto
    def salirProyecto():
        PregCierre = messagebox.askquestion('Cerrar programa','Desea cerrar el programa ?')
        if PregCierre == 'yes':
            raiz.destroy()
    
    #===>Reinicia el proyecto
    def boton_reset():
        raiz.destroy()    
        proyeccion_semanal()
        
    
    #==============================================================================
    #===================== Se establece el menú de la raíz ========================
    #==============================================================================
    
    barraMenu = Menu(raiz)
    raiz.config(menu=barraMenu)
    
    archivoMenu = Menu(barraMenu, tearoff = 0)
    
    archivoMenu.add_command(label = 'Nuevo')
    archivoMenu.add_command(label = 'Guardar')
    archivoMenu.add_command(label = 'Guardar Como')
    archivoMenu.add_separator()
    archivoMenu.add_command(label = 'Nuevo Análiss de Datos',command = proySem)
    archivoMenu.add_separator()
    archivoMenu.add_command(label = 'Reiniciar', command = boton_reset)
    archivoMenu.add_command(label = 'Salir', command = salirProyecto)
    
    archivoEdicion = Menu(barraMenu, tearoff = 0)
    archivoEdicion.add_command(label = 'Copiar')
    archivoEdicion.add_command(label = 'Cortar')
    archivoEdicion.add_command(label = 'Pegar')
    
    archivoHerramientas = Menu(barraMenu, tearoff = 0)
    archivoHerramientas.add_separator()
    archivoHerramientas.add_command(label = 'Generar Plantilla Para Hstóricos de Carga', command = his_carga_template)
    archivoHerramientas.add_command(label = 'Generar Hstóricos de Carga', command = his_carga_final)
    
    archivoConfiguraciones = Menu(barraMenu, tearoff = 0)
#    archivoConfiguraciones.add_command(label = 'Configuraciones por defecto')
    archivoConfiguraciones.add_command(label = 'Definir Base de Datos', command = base_datos)
    archivoConfiguraciones.add_separator()
    archivoConfiguraciones.add_command(label = 'Definir Destino de Resultados', command = destino_resultados)
    
    archivoAyuda = Menu(barraMenu, tearoff = 0)
    archivoAyuda.add_command(label = 'Acerca de...', command = infoAdicional)
    
    barraMenu.add_cascade(label = 'Archivo', menu = archivoMenu)
    barraMenu.add_cascade(label = 'Edición', menu = archivoEdicion)
    barraMenu.add_cascade(label = 'Herramientas', menu = archivoHerramientas)
    barraMenu.add_cascade(label = 'Configuraciones', menu = archivoConfiguraciones)
    barraMenu.add_cascade(label = 'Ayuda', menu = archivoAyuda)
    
    #==============================================================================
    #==============================================================================
    
    
    #------------------------- Ingreso de imágenes --------------------------------
    
    imag_eeasa_2 = PhotoImage(file='eeasa_peq.png')    #eeasa_peq_2
    Label(raiz, image = imag_eeasa_2).place(x=140,y=20) 
     
    imag_pronost_2 = PhotoImage(file='energy.png')       
    Label(raiz, image = imag_pronost_2).place(x=10,y=390) 
    
    #------------------------- Ingreso de textos ----------------------------------
    
    txt_1 = Label(raiz, text = 'LOAD FORECASTING', fg='#616A6B',
                  font=('Arial',12)).place(x=200,y=150)
            
    txq_2 = Label(raiz, text = 'Seleccione la metodología: ',fg='#616A6B',
                  font=('Arial',12)).place(x=30,y=200)
    txq_3 = Label(raiz, text = 'Realizado por: ',fg='#616A6B',
                  font=('Arial',12)).place(x=30,y=230) 
    txq_4 = Label(raiz, text = 'Generar gráficas: ',fg='#616A6B',
                  font=('Arial',12)).place(x=30,y=260)  
    
    txt_5 = Label(raiz, text = 'SELECCIONE LAS FECHAS: ', fg='#616A6B',
                  font=('Arial',12)).place(x=20,y=310)
    #-------------------------- DATOS DE PROYECCIÓN -------------------------------
    txq_6 = Label(raiz, text = 'Datos\nProyección: ',fg='#616A6B',
                  font=('Arial',12)).place(x=215,y=340) 
    
    Label(raiz, text = 'AÑO: ', fg='#616A6B',font=('Arial',10)).place(x=140,y=390)
#    cbx_anio_p = ttk.Combobox(values=['2013','2014','2015','2016','2017','2018'],state="readonly")
    cbx_anio_p = ttk.Combobox(values=anios_disponibles,state="readonly")
    cbx_anio_p.place(x=190,y=390)
    
    Label(raiz, text = 'MES: ', fg='#616A6B',font=('Arial',10)).place(x=140,y=420)
    cbx_mes_p = ttk.Combobox(values=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'],state="readonly")
    cbx_mes_p.place(x=190,y=420)
    
    Label(raiz, text = 'DÍA: ', fg='#616A6B',font=('Arial',10)).place(x=140,y=450)
    cbx_dia_p = ttk.Combobox(values=['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31'],state="readonly")
    cbx_dia_p.place(x=190,y=450)
    
    #-------------------------- DATOS DE COMPARACIÓN ------------------------------
    txq_7 = Label(raiz, text = 'Datos\nComparación: ',fg='#616A6B',
                  font=('Arial',12)).place(x=410,y=340) 
    
    Label(raiz, text = 'AÑO: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=390)
#    cbx_anio_c = ttk.Combobox(values=['2013','2014','2015','2016','2017','2018'],state="readonly")
    cbx_anio_c = ttk.Combobox(values=anios_disponibles,state="readonly")
    cbx_anio_c.place(x=400,y=390)
    
    Label(raiz, text = 'MES: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=420)
    cbx_mes_c = ttk.Combobox(values=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'],state="readonly")
    cbx_mes_c.place(x=400,y=420)
    
    Label(raiz, text = 'DÍA: ', fg='#616A6B',font=('Arial',10)).place(x=350,y=450)
    cbx_dia_c = ttk.Combobox(values=['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31'],state="readonly")
    cbx_dia_c.place(x=400,y=450)
    
    
    Label(raiz, text = 'AGF ®', fg='#616A6B',font=('Tw Cen MT',9)).place(x=545,y=475) 
    
    #----------------------Ingreso de MENÚ txq_2 METODOLOGÍA ----------------------
    cbx_txq_2 = ttk.Combobox(values=['Promedios (CECON)', 'REGRESIÓN LINEAL (LR)','S.V.M.','M.L.P.'],state="readonly")
    cbx_txq_2.place(x=250,y=200)
    
    #----------------------Ingreso de MENÚ txq_3 REALIZD POR ----------------------
    cbx_txq_3 = ttk.Combobox(values=['A.G','B.M','A.T','C.G','L.T','E.Ñ','E.T','Otro'],state="readonly")
    cbx_txq_3.place(x=250,y=230)
    
    #----------------------Ingreso de MENÚ txq_4 COMPARACIÓN ----------------------
    cbx_txq_4 = ttk.Combobox(values=['SI','NO'],state="readonly")
    cbx_txq_4.place(x=250,y=260)
    
    
    #==============================================================================
    #============================== BOTÓN GENERAR =================================
    #==============================================================================
    b_go = Button(raiz, text = '  GENERAR PROYECCIÓN  ',bg='#2ECC71',fg='#1A5276',
                  width=20, height=5, command = lambda:boton_Proyeccion())
    b_go.place(x=420,y=200)
    
    def boton_Proyeccion():
       
    #==============================================================================
    #======================== DECLARACIÓN DE VARIABLES ============================
    #==============================================================================    
        global pos_Metodologia, pos_Realizado, pos_Comparacion
        pos_Metodologia = cbx_txq_2.current()
        pos_Realizado   = cbx_txq_3.current()
        pos_Comparacion = cbx_txq_4.current()
        
        global pos_anio_p, pos_mes_p, pos_dia_p
        pos_anio_p = cbx_anio_p.current()
        pos_mes_p  = cbx_mes_p. current()
        pos_dia_p  = cbx_dia_p. current()
        
        global pos_anio_c, pos_mes_c, pos_dia_c
        pos_anio_c = cbx_anio_c.current()
        pos_mes_c  = cbx_mes_c. current()
        pos_dia_c  = cbx_dia_c. current()
    #******************************************************************************
        global po_Metodologia, po_Realizado, po_Comparacion
        po_Metodologia = cbx_txq_2.get()
        po_Realizado   = cbx_txq_3.get()
        po_Comparacion = cbx_txq_4.get()
        
        global po_anio_p, po_mes_p, po_dia_p
        po_anio_p = cbx_anio_p.get()
        po_mes_p  = cbx_mes_p. get()
        po_dia_p  = cbx_dia_p. get()
        
        global po_anio_c, po_mes_c, po_dia_c
        po_anio_c = cbx_anio_c.get()
        po_mes_c  = cbx_mes_c. get()
        po_dia_c  = cbx_dia_c. get()
    #==============================================================================
    #==============================================================================
    
       
    #==============================================================================
    #======================== VERIFICA DATOS FALTANTES ============================
    #==============================================================================   
        from calendar import monthrange
        
        if pos_Metodologia == -1:
            messagebox.showinfo(message="Ingrese la Metodología", title="Faltan datos")
        elif pos_Realizado == -1:
            messagebox.showinfo(message="Ingrese Responsable", title="Faltan datos")
        elif pos_Comparacion == -1:
            messagebox.showinfo(message="Ingrese Comparación", title="Faltan datos")
    #******************************************************************************        
        elif pos_anio_p == -1:
            messagebox.showinfo(message="Ingrese Año de Proyección", title="Faltan datos")
        elif pos_mes_p == -1:
            messagebox.showinfo(message="Ingrese Mes de Proyección", title="Faltan datos")
        elif pos_dia_p == -1:
            messagebox.showinfo(message="Ingrese Día de Proyección", title="Faltan datos")
            
        elif int(po_dia_p) > monthrange(int(po_anio_p), pos_mes_p+1)[1]:
            messagebox.showinfo(message='No existe día '+po_dia_p +' para el mes ' +po_mes_p + ' de '+ po_anio_p, title="Dato erróneo")
    #******************************************************************************    
        elif pos_anio_c == -1:
            messagebox.showinfo(message="Ingrese Año de Comparación", title="Faltan datos")
        elif pos_mes_c == -1:
            messagebox.showinfo(message="Ingrese Mes de Comparación", title="Faltan datos")
        elif pos_dia_c == -1:
            messagebox.showinfo(message="Ingrese Día de Comparación", title="Faltan datos")
            
        elif int(po_dia_c) > monthrange(int(po_anio_c), pos_mes_c+1)[1]:
            messagebox.showinfo(message='No existe día '+po_dia_c +' para el mes ' +po_mes_c + ' de '+ po_anio_c, title="Dato erróneo")
    #******************************************************************************    
        
        else:
            if pos_Metodologia == 0:
                print(' ')
                print ('EJECUTADO PROYECCIÓN POR METOLOGÍA PROMEDIOS...')
                print(' ')
                proyeccion_arima()
            elif pos_Metodologia == 1:
                print(' ')
                print ('EJECUTADO PROYECCIÓN POR METOLOGÍA REGRESIÓN LINEAL...')
                print(' ')
                proyeccion_rl()
            elif pos_Metodologia == 2:
                print(' ')
                print ('EJECUTADO PROYECCIÓN POR METOLOGÍA MÁQUINA DE VECTOR SOPORTE...')
                print(' ')
                proyeccion_svm()
            elif pos_Metodologia == 3:
                print(' ')
                print ('EJECUTADO PROYECCIÓN POR METOLOGÍA PERCEPTRÓN MULTICAPA...')
                print(' ')
                proyeccion_mlp()
                
    
    #==============================================================================
    #==============================================================================
    raiz.mainloop()
    
    print (' ')
    print (' ')
    print ('**  Módulo de Proyección Semanal COMPLETADO...!')
    print (' ')
    print (' ')
   
     
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#******************************************************************************
#==============================================================================
    
    
def graficas_rapidas_2():
    # -*- coding: utf-8 -*-
    """
    Created on Tue Feb 26 14:03:20 2019
    ESCUELA POLITÉCNICA NACIONAL
    AUTOR: ALEXIS GUAMAN FIGUEROA
    2_MÓDULO DE GRÁFICAS RÁPIDAS
    """
    
    # -*- coding: utf-8 -*-
    """
    Created on Tue Feb 19 12:15:09 2019
    
    @author: ALEXIS GUAMAN
    """
    
    #importa librería pandas
#    import pandas as pd
#    import matplotlib.pyplot as plt
#    import numpy as np
    
    
    #=============================== VAR GLOBALES =================================
    global pos_menu_2
    global pos_se, po_anio, po_se, val_check,po_dia
    global direccion_base_datos
    anio = po_anio
    sube = po_se + ' '
    
    #==============================================================================
    
    print(' ')
    print(' ')
    print('Procesing Plots...')
    print(' ')
#    global file_path
#    print (file_path)
    
    doc_clima = direccion_base_datos+'\\HIS_CLIMA_' + anio + '.xls'
    doc_demand = direccion_base_datos+'\\HIS_POT_' + anio + '.xls'
    
    
    df1 = pd.read_excel(doc_clima, sheetname='Hoja1',  header=None)
    
    hora = df1.iloc[:,2].values.tolist()
    hora.pop(0)
    
    temp = df1.iloc[:,4].values.tolist()
    temp.pop(0)
    
    mes = df1.iloc[:,0].values.tolist()
    mes.pop(0)
    
    dia = df1.iloc[:,1].values.tolist()
    dia.pop(0)
    
    #===> Reemplaza valores (0) con (nan)
    for i in range(len(temp)):
        if temp[i] == 0:
            temp[i] = float('nan')
    #==============================================================================
    df2 = pd.read_excel(doc_demand, sheetname='Hoja1',  header=None)
    
    amba = df2.iloc[:,4].values.tolist()
    amba.pop(0)
    
    toto = df2.iloc[:,5].values.tolist()
    toto.pop(0)
    
    puyo = df2.iloc[:,6].values.tolist()
    puyo.pop(0)
    
    tena = df2.iloc[:,7].values.tolist()
    tena.pop(0)
    
    bani = df2.iloc[:,8].values.tolist()
    bani.pop(0)
    
    #===> Reemplaza valores (0) con (nan)
    for i in range(len(temp)):
        if amba[i] == 0:
            amba[i] = float('nan')
        if toto[i] == 0:
            toto[i] = float('nan')
        if puyo[i] == 0:
            puyo[i] = float('nan')
        if tena[i] == 0:
            tena[i] = float('nan')
        if bani[i] == 0:
            bani[i] = float('nan')
    #===> Se establece una matriz con los datos importados
    data = np.column_stack((temp,amba,toto, puyo, tena, bani))
    
    #==============================================================================
    #=============================== GRÁFICA ANUAL ================================
    #==============================================================================
    if   pos_menu_2==0:
        
    #===> Tamaño de la ventana de la gráfica
        plt.subplots(figsize=(15, 8))
    #===> Título general superior
        plt.suptitle(u' GRÁFICA ANUAL ', fontsize=16, fontweight='bold')
        
        
        import calendar
    #===> Creamos una lista con las posiciones del primer día de cada mes
        dias = [mes.index('ENERO'), mes.index('FEBRERO'), mes.index('MARZO'),
                mes.index('ABRIL'), mes.index('MAYO'),    mes.index('JUNIO'), 
                mes.index('JULIO'), mes.index('AGOSTO'),  mes.index('SEPTIEMBRE'),
                mes.index('OCTUBRE'), mes.index('NOVIEMBRE'), mes.index('DICIEMBRE')]
        
    #===> Creamos una lista con los nombres de los meses
        meses = calendar.month_name[1:13]
        
        plt.subplot(1,1,1)
    #===> Tiempo para la gráfica, considera un año al tomar todos los datos
        plt.xlim(0,8759)
        
    #===> Subtítulo de la gráfica
        plt.title(u'Carga vs. Temperatura S/E '+sube +anio, fontweight='bold',
                  style='italic')
        
    #=== > Etiq. del eje x (meses del año) se marcan según el primer dia de c/mes
        plt.xticks(dias, meses, size = 'small', color = 'b', rotation = 45) 
        
    #===> Condicional de gráficas, var generales: se, col
        if pos_se == 0:
            se = amba
            col = '#DCD037'
        elif pos_se == 1:
            se = bani
            col = '#4ACB71'
        elif pos_se == 2:
            se = puyo
            col = '#349A9D'
        elif pos_se == 3:
            se = tena
            col = '#CC8634'
        elif pos_se == 4:
            se = toto
            col = '#CD336F'
        
        plt.plot(se, col, label = 'Potencia')
        plt.legend()
        plt.xlabel(u'TIEMPO [ meses ]')
        plt.ylabel(u'CARGA  [ kW ]') 
        plt.twinx()  
        g2, = plt.plot(temp, 'b', label = 'Temperatura en °C')
        plt.ylabel(u'TEMPERATURA  [ °C ]') 
        plt.legend(handles=[g2], loc=4)
        plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
                
    #=============================== SCATTER ANUAL ================================
    
    #===> Si se activa el check, genera el scatter
        if val_check == 1:
            
            fig2, (ax1) = plt.subplots(nrows=1,ncols=1,figsize=(15, 8))      
            plt.suptitle(u'SCATTER ANUAL',fontweight='bold',
                      style='italic')
            ax1.scatter(x=temp, y=se, marker='.', c=col)
            ax1.set_title('$Temperatura$ vs. $Potencia$ $S/E$ '+sube+'$año $: ' +anio)
            ax1.set_ylabel('$Carga$  $[$ $kW$ $]$')
            ax1.set_xlabel('$Temperatura$ $en$ $°C$')
    
    #==============================================================================
    #==============================================================================
            
            
    
    #==============================================================================
    #=============================== GRÁFICA MENSUAL ==============================
    #==============================================================================
    if   pos_menu_2==1:
        
        global pos_mes, po_mes
        mesi = ' / '+po_mes+ ' '
        
        import calendar
        
    #===> Tamaño de la ventana de la gráfica
        plt.subplots(figsize=(15, 8))
    #===> Título general superior
        plt.suptitle(u' GRÁFICA MENSUAL ', fontsize=16, fontweight='bold') 
                       
    #===> Subrutina, genera posiciones de los meses según el mes seleccionado
        x_meses = [mes.index('ENERO'), mes.index('FEBRERO'), mes.index('MARZO'),
                   mes.index('ABRIL'), mes.index('MAYO'),    mes.index('JUNIO'), 
                   mes.index('JULIO'), mes.index('AGOSTO'),  mes.index('SEPTIEMBRE'),
                   mes.index('OCTUBRE'), mes.index('NOVIEMBRE'), mes.index('DICIEMBRE')]
        for i in range (12):
            if pos_mes == i:
                if pos_mes == 11:
                    val_ini = x_meses[i]
                    val_fin = 8759
                else:
                    val_ini = x_meses[i]
                    val_fin = x_meses[i+1]-1
                    
        cant_datos_mes = val_fin - val_ini
        
    #===> Creamos una lista con los números de día del mes seleccionado (etiquetas)
        x_dias = calendar.monthcalendar(int(anio),pos_mes+1)
        dias = []
        for i in range (len(x_dias)):
            for j in range (len(x_dias[i])):
                if x_dias[i][j] != 0:
                    dias.append(x_dias[i][j])
        dias_str = []
        for i in range (len(dias)):
            dias_str.append(str(dias[i]))
                    
    #===> Creamos una lista con las escalas correspondientes a los num de días
        i = val_ini
        can_dias = []
        while i <= val_fin:
            can_dias.append(i)
            i = i + (cant_datos_mes / len(dias))
    
        dias_int = []
        for i in range (len(can_dias)):
            dias_int.append(int(can_dias[i])+1)
        
        plt.subplot(1,1,1)
    #===> Tiempo para la gráfica, considera posiciones del mes seleccionado
        plt.xlim(val_ini,val_fin)
        
    #===> Subtítulo de la gráfica
        plt.title(u'Carga vs. Temperatura S/E '+sube +mesi +anio, fontweight='bold',
                  style='italic')
        
    #=== > Etiq. del eje x (días del mes seleccionado)
        plt.xticks(dias_int, dias_str, size = 'small', color = 'b', rotation = 45)
        
    #===> Condicional de gráficas, var generales: se, col
        if pos_se == 0:
            se = amba
            col = '#DCD037'
        elif pos_se == 1:
            se = bani
            col = '#4ACB71'
        elif pos_se == 2:
            se = puyo
            col = '#349A9D'
        elif pos_se == 3:
            se = tena
            col = '#CC8634'
        elif pos_se == 4:
            se = toto
            col = '#CD336F'
        
        plt.plot(se, col, label = 'Potencia')
        plt.legend()
        plt.xlabel(u'TIEMPO [ días ]')
        plt.ylabel(u'CARGA  [ kW ]') 
        plt.twinx()  
        g2, = plt.plot(temp, 'b', label = 'Temperatura en °C')
        plt.ylabel(u'TEMPERATURA  [ °C ]') 
        plt.legend(handles=[g2], loc=4)
        plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
                
    #=============================== SCATTER MENSUAL ================================
    
    #===> Si se activa el check, genera el scatter
        if val_check == 1:
            
            for i in range (val_ini):
                se.pop(0)
                temp.pop(0)
                
            i = cant_datos_mes + 1
            inicio = len(se)
            for i in range (inicio - cant_datos_mes - 1):
                se.pop(cant_datos_mes + 1)
                temp.pop(cant_datos_mes + 1)
            
            
            fig2, (ax1) = plt.subplots(nrows=1,ncols=1,figsize=(15, 8))      
             
            ax1.scatter(x=temp, y=se, marker='.',linewidth=3, c=col)
            ax1.set_title('Scatter MENSUAL: $Temperatura$ vs. $Potencia$ $S/E$ '+sube)
            ax1.set_ylabel('$Carga$  $[$ $kW$ $]$')
            ax1.set_xlabel('$Temperatura$ $en$ $°C$')
    
    #==============================================================================
    #==============================================================================
    
    
    
    
    #==============================================================================
    #=============================== GRÁFICA SEMANAL ==============================
    #==============================================================================
    if   pos_menu_2==2:
        
        #global pos_mes, po_mes
        global po_sem, dias_restantes
        mesi = ' / '+po_mes+ ' '
        
        import calendar
        
    #===> Tamaño de la ventana de la gráfica
        plt.subplots(figsize=(15, 8))
    #===> Título general superior
        plt.suptitle(u' GRÁFICA SEMANAL ', fontsize=16, fontweight='bold') 
                       
    #===> Subrutina, genera posiciones de los días según el mes y semana seleccionada
        x_meses = [mes.index('ENERO'), mes.index('FEBRERO'), mes.index('MARZO'),
                   mes.index('ABRIL'), mes.index('MAYO'),    mes.index('JUNIO'), 
                   mes.index('JULIO'), mes.index('AGOSTO'),  mes.index('SEPTIEMBRE'),
                   mes.index('OCTUBRE'), mes.index('NOVIEMBRE'), mes.index('DICIEMBRE')]
        for i in range (12):
            if pos_mes == i:
                if pos_mes == 11:
                    val_ini = x_meses[i]
                    #val_fin = 8759
                else:
                    val_ini = x_meses[i]
                    #val_fin = x_meses[i+1]-1
            
        if int(po_sem) == 1:
            val_ini_sem = val_ini
        elif int(po_sem) == 2:
            val_ini_sem = val_ini + 24 * dias_restantes
        elif int(po_sem) == 3:
            val_ini_sem = val_ini + 24 * dias_restantes + 24 * 7
        elif int(po_sem) == 4:
            val_ini_sem = val_ini + 24 * dias_restantes + 24 * 14 
        elif int(po_sem) == 5:
            val_ini_sem = val_ini + 24 * dias_restantes + 24 * 21
        elif int(po_sem) == 6:
            val_ini_sem = val_ini + 24 * dias_restantes + 24 * 28        
            
    
        val_fin = val_ini_sem + 167
        cant_datos_mes = val_fin - val_ini_sem
        
    #===> Creamos una lista con los números de día del mes seleccionado (etiquetas)
        
        dias_semana =['Lunes ','Martes ','Miércoles ','Jueves ','Viernes ','Sabado ','Domingo ']
        numero_dias = []
        for i in range (dia[val_ini_sem], dia[val_ini_sem] + 7):
            numero_dias.append(i)
        dias_str = []
        for i in range (7):
            dias_str.append(dias_semana[i]+str(numero_dias[i]))
                    
    #===> Creamos una lista con las escalas correspondientes a los num de días
        i = val_ini_sem
        can_dias = []
        while i <= val_fin:
            can_dias.append(i)
            i = i + (cant_datos_mes / len(dias_str))
    
        dias_int = []
        for i in range (len(can_dias)):
            dias_int.append(int(can_dias[i])+1)
        
        plt.subplot(1,1,1)
    #===> Tiempo para la gráfica, considera posiciones del mes y SEMANA seleccionada
        plt.xlim(val_ini_sem,val_fin)
        
    #===> Subtítulo de la gráfica
        plt.title(u'Carga vs. Temperatura S/E '+sube +' / Semana N° '+po_sem +mesi +anio, fontweight='bold',
                  style='italic')
        
    #=== > Etiq. del eje x (días del mes seleccionado)
        plt.xticks(dias_int, dias_str, size = 'small', color = 'b', rotation = 45)
        
    #===> Condicional de gráficas, var generales: se, col
        if pos_se == 0:
            se = amba
            col = '#DCD037'
        elif pos_se == 1:
            se = bani
            col = '#4ACB71'
        elif pos_se == 2:
            se = puyo
            col = '#349A9D'
        elif pos_se == 3:
            se = tena
            col = '#CC8634'
        elif pos_se == 4:
            se = toto
            col = '#CD336F'
        
        plt.plot(se, col, label = 'Potencia')
        plt.legend()
        plt.xlabel(u'TIEMPO [ días ]')
        plt.ylabel(u'CARGA  [ kW ]') 
        plt.twinx()  
        g2, = plt.plot(temp, 'b', label = 'Temperatura en °C')
        plt.ylabel(u'TEMPERATURA  [ °C ]') 
        plt.legend(handles=[g2], loc=4)
        plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
                
    #=============================== SCATTER SEMANAL ================================
    
    #===> Si se activa el check, genera el scatter
        if val_check == 1:
            
            for i in range (val_ini_sem):
                se.pop(0)
                temp.pop(0)
                
            i = cant_datos_mes + 1
            inicio = len(se)
            for i in range (inicio - cant_datos_mes - 1):
                se.pop(cant_datos_mes + 1)
                temp.pop(cant_datos_mes + 1)
            
            
            fig2, (ax1) = plt.subplots(nrows=1,ncols=1,figsize=(15, 8))      
             
            ax1.scatter(x=temp, y=se, marker='.',linewidth=3, c=col)
            ax1.set_title('Scatter SEMANAL: $Temperatura$ vs. $Potencia$ $S/E$ '+sube)
            ax1.set_ylabel('$Carga$  $[$ $kW$ $]$')
            ax1.set_xlabel('$Temperatura$ $en$ $°C$')
    

    #==============================================================================
    #=============================== GRÁFICA DIARIA ==============================
    #==============================================================================
    if   pos_menu_2==3:
        
        #global pos_mes, po_mes
        global pos_dia, po_dia
        
    #    mesi = ' / '+po_mes+ ' '
        
        import calendar
        
    #===> Tamaño de la ventana de la gráfica
        plt.subplots(figsize=(15, 8))
        
    #===> Título general superior
        plt.suptitle(u' GRÁFICA DIARIA ', fontsize=16, fontweight='bold') 
                       
    #===> Subrutina, genera posiciones de los días según el mes y semana seleccionada
        x_meses = [mes.index('ENERO'), mes.index('FEBRERO'), mes.index('MARZO'),
                   mes.index('ABRIL'), mes.index('MAYO'),    mes.index('JUNIO'), 
                   mes.index('JULIO'), mes.index('AGOSTO'),  mes.index('SEPTIEMBRE'),
                   mes.index('OCTUBRE'), mes.index('NOVIEMBRE'), mes.index('DICIEMBRE')]
        for i in range (12):
            if pos_mes == i:
                if pos_mes == 11:
                    val_ini = x_meses[i]
                else:
                    val_ini = x_meses[i]
        
        val_ini = val_ini + pos_dia * 24
        val_fin = val_ini + 24
        
        print(val_ini)
        print(val_fin)
        
        
    #===> Creamos una lista con los números de horas del día seleccionado (etiquetas)
        
        horas_dia = []
        horas_str = []
        for i in range (1,25):
            horas_dia.append(i)
            horas_str.append(str(horas_dia[i-1])+':00:00')
                    
    #===> Creamos una lista con las escalas correspondientes a los num de horas
            
        i = val_ini
        can_horas = []
        while i <= val_fin:
            can_horas.append(i)
            i = i + (24 / len(horas_str))
        
    
        plt.subplot(1,1,1)
    #===> Tiempo para la gráfica, considera posiciones del mes y SEMANA seleccionada
        plt.xlim(val_ini,val_fin)
        
    #===> Subtítulo de la gráfica
        plt.title(u'Carga vs. Temperatura S/E '+sube +' / '+po_dia +' de '+po_mes +' de '+anio, fontweight='bold',
                  style='italic')
        
    #=== > Etiq. del eje x (días del mes seleccionado)
        plt.xticks(can_horas, horas_str, size = 'small', color = 'b', rotation = 45)
        
    #===> Condicional de gráficas, var generales: se, col
        if pos_se == 0:
            se = amba
            col = '#DCD037'
        elif pos_se == 1:
            se = bani
            col = '#4ACB71'
        elif pos_se == 2:
            se = puyo
            col = '#349A9D'
        elif pos_se == 3:
            se = tena
            col = '#CC8634'
        elif pos_se == 4:
            se = toto
            col = '#CD336F'
        
        plt.plot(se, col, label = 'Potencia')
        plt.legend()
        plt.xlabel(u'TIEMPO [ Horas ]')
        plt.ylabel(u'CARGA  [ kW ]') 
        plt.twinx()  
        g2, = plt.plot(temp, 'b', label = 'Temperatura en °C')
        plt.ylabel(u'TEMPERATURA  [ °C ]') 
        plt.legend(handles=[g2], loc=4)
        plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
                
    #=============================== SCATTER DIARIO ================================
    
    #===> Si se activa el check, genera el scatter
        if val_check == 1:
            
            for i in range (val_ini):
                se.pop(0)
                temp.pop(0)
                
            i = 24 + 1
            inicio = len(se)
            for i in range (inicio - 24 - 1):
                se.pop(24 + 1)
                temp.pop(24 + 1)
            
            
            fig2, (ax1) = plt.subplots(nrows=1,ncols=1,figsize=(15, 8))      
             
            ax1.scatter(x=temp, y=se, marker='.',linewidth=3, c=col)
            ax1.set_title('Scatter DIARIO: $Temperatura$ vs. $Potencia$ $S/E$ '+sube)
            ax1.set_ylabel('$Carga$  $[$ $kW$ $]$')
            ax1.set_xlabel('$Temperatura$ $en$ $°C$')
    
    #==============================================================================
    #==============================================================================
    
    
    print(' ')
    print('Completed...                 OK...!')
    print(' ')

#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#******************************************************************************
#==============================================================================     
    
    
    
def graficas_acumuladas():
    # -*- coding: utf-8 -*-
    """
    Created on Tue Mar 12 11:21:42 2019
    
    ESCUELA POLITÉCNICA NACIONAL
    AUTOR: ALEXIS GUAMAN FIGUEROA
    3_MÓDULO DE GRÁFICAS ACUMULADAS
    """
    
    #importa librería pandas
#    import pandas as pd
#    import matplotlib.pyplot as plt
#    import numpy as np
    
    
    #=============================== VAR GLOBALES =================================
    global pos_menu_2
    global pos_se, po_anio, po_se, val_check,po_dia
    anio = po_anio
    sube = po_se + ' '
    
    #==============================================================================
    
    print(' ')
    print(' ')
    print('Procesing Plots...')
    print(' ')
    
    doc_clima = direccion_base_datos+'\\HIS_CLIMA_' + anio + '.xls'
    doc_demand = direccion_base_datos+'\\HIS_POT_' + anio + '.xls'
    
    df1 = pd.read_excel(doc_clima, sheetname='Hoja1',  header=None)
    
    mes = df1.iloc[:,0].values.tolist()
    mes.pop(0)
    
    dia = df1.iloc[:,1].values.tolist()
    dia.pop(0)
    
    hora = df1.iloc[:,2].values.tolist()
    hora.pop(0)
    
    temp = df1.iloc[:,4].values.tolist()
    temp.pop(0)
    
    
    
    
    #===> Reemplaza valores (0) con (nan)
    for i in range(len(temp)):
        if temp[i] == 0:
            temp[i] = float('nan')
    #==============================================================================
    df2 = pd.read_excel(doc_demand, sheetname='Hoja1',  header=None)
    
    amba = df2.iloc[:,4].values.tolist()
    amba.pop(0)
    
    toto = df2.iloc[:,5].values.tolist()
    toto.pop(0)
    
    puyo = df2.iloc[:,6].values.tolist()
    puyo.pop(0)
    
    tena = df2.iloc[:,7].values.tolist()
    tena.pop(0)
    
    bani = df2.iloc[:,8].values.tolist()
    bani.pop(0)
    
    #===> Reemplaza valores (0) con (nan)
    for i in range(len(temp)):
        if amba[i] == 0:
            amba[i] = float('nan')
        if toto[i] == 0:
            toto[i] = float('nan')
        if puyo[i] == 0:
            puyo[i] = float('nan')
        if tena[i] == 0:
            tena[i] = float('nan')
        if bani[i] == 0:
            bani[i] = float('nan')
    #===> Se establece una matriz con los datos importados
    data = np.column_stack((temp,amba,toto, puyo, tena, bani))
    
    
    
    #===> Condicional de gráficas, var generales: se, col
    if pos_se == 0:
        se = amba
        col = '#DCD037'
    elif pos_se == 1:
        se = bani
        col = '#4ACB71'
    elif pos_se == 2:
        se = puyo
        col = '#349A9D'
    elif pos_se == 3:
        se = tena
        col = '#CC8634'
    elif pos_se == 4:
        se = toto
        col = '#CD336F'
    
    
    
    for i in range (len(dia)):
        if mes[i] == 'ENERO':
            mes[i] = '01'
        elif mes [i] == 'FEBRERO':
            mes[i] = '02'
        elif mes [i] == 'MARZO':
            mes[i] = '03'
        elif mes [i] == 'ABRIL':
            mes[i] = '04'
        elif mes [i] == 'MAYO':
            mes[i] = '05'
        elif mes [i] == 'JUNIO':
            mes[i] = '06'
        elif mes [i] == 'JULIO':
            mes[i] = '07'
        elif mes [i] == 'AGOSTO':
            mes[i] = '08'
        elif mes [i] == 'SEPTIEMBRE':
            mes[i] = '09'
        elif mes [i] == 'OCTUBRE':
            mes[i] = '10'
        elif mes [i] == 'NOVIEMBRE':
            mes[i] = '11'
        elif mes [i] == 'DICIEMBRE':
            mes[i] = '12'
            
    dia_str = []
    for i in range (len(dia)):
        dia_str.append( str (dia[i]))
        if dia[i] < 10:
            dia_str[i] = '0'+dia_str[i]
    
    data_2 = np.column_stack((mes,dia_str,se,hora))
    
    df_general = pd.DataFrame(np.array(data_2))
    
    df_general = df_general.assign(idx = (anio + '-' + df_general[0] + '-' + 
                                          df_general[1]))
    
    df_general = df_general.set_index('idx')
    
    del df_general[0]
    del df_general[1]
    
    df_general = df_general.astype(np.float)
    
    df_general_pivot = df_general.pivot(columns=3)
    
    
    df_general_pivot.T.plot(figsize=(15,8), legend=False, color=col, alpha=0.02)
    
    #===> Título general superior
    plt.suptitle(u' GRÁFICA DIARIA ACUMULADA ', fontsize=16, fontweight='bold') 
    
    #===> Subtítulo de la gráfica
    plt.title(u'Comparativa para la S/E '+sube +', año '+anio, fontweight='bold',
              style='italic')
    
    
    
    plt.ylabel(u'CARGA  [ kW ]')
    plt.xlabel(u'TIEMPO [ Horas ]')
    
    plt.grid(color='#C8C8C8', linestyle='-', linewidth=0.5)
    
             
     
             
    #===> Creamos una lista con los números de horas del día seleccionado (etiquetas)
        
    horas_dia = []
    horas_str = []
    for i in range (1,25):
        horas_dia.append(i)
        horas_str.append(str(horas_dia[i-1])+':00:00')
          
    #===> Creamos una lista con las escalas correspondientes a los num de horas    
    
    can_horas = []
    for i in range (24):
        can_horas.append(i)
        
    #=== > Etiq. del eje x (días del mes seleccionado)
    plt.xticks(can_horas, horas_str, size = 'small', color = 'b', rotation = 45)

#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#//////////////////////////////////////////////////////////////////////////////
#******************************************************************************
#==============================================================================
#******************************************************************************
#==============================================================================    

root = Tk()

#miFrame = Frame(root)
#miFrame.pack()

root.geometry('350x230+700+300')
root.title('LOGIN')
root.resizable(0,0)
root.iconbitmap('epn.ico')

txq_1 = Label(root, text = 'Inicio de sesión ', fg='#616A6B',font=('Arial',12)).place(x=0,y=0)  
Label(root, text = 'AGF ®', fg='#616A6B',font=('Tw Cen MT',9)).place(x=300,y=200)
      
Label(root, text = 'Usuario: '   , fg='#616A6B',font=('Arial',12)).place(x=10,y= 50)
Label(root, text = 'Contraseña: ', fg='#616A6B',font=('Arial',12)).place(x=10,y=100)  

# Usuario por defecto      
default_user = StringVar(root, value='1720282571')
US = Entry(root,textvariable=default_user, fg='#616A6B',font=('Arial',11)).place(x=150,y=  50)
PS = Entry(root,show ='*',fg='#616A6B',font=('Arial',11))
PS.place(x=150,y= 100)             
           
def ingreso(self):
    pasw = PS.get()
    
    if pasw == '152125':     #================ CONTRASEÑA =====================
        root.destroy()
        print(' ')
        print(' ')
        print('**   Contraseña correcta...     EJECUTANDO....')
        print(' ')
        print(' ')
        #graficas_rapidas()
        proyeccion_semanal()
    else:
        incorr_ps = messagebox.askretrycancel('ERROR','Contraseña incorrecta...!')
        if incorr_ps == False:
            root.destroy()
            
#Button(root, text = '  INGRESAR AL SISTEMA  ',bg='#2ECC71',fg='#1A5276',command = lambda:ingreso(self)).place(x=110,y=150)
Label(root, text = 'Presione ENTER para INGRESAR: ', bg='#2ECC71',fg='#1A5276',font=('Arial',12)).place(x=50,y=150) 
PS.bind('<Return>', ingreso)

      
root.mainloop()