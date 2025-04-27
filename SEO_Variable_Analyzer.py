# Nombre:
# Ariel Martínez Birlanga


###############################
###                         ###
###        Librerias        ###
###                         ###
###############################

from sklearn.ensemble import GradientBoostingClassifier
from sklearn.model_selection import cross_val_score, RepeatedStratifiedKFold

from bs4 import BeautifulSoup
from urllib.parse import urlparse, urlsplit
from urllib.request import urlopen
import numpy as np
import re
import requests
import sys
import os
import tldextract
import whois
import datetime
import json
import warnings
import time
import xlsxwriter
import threading
#import gc

import tkinter as tk
from tkinter import ttk,messagebox



###############################
###                         ###
###   Elementos Generales   ###
###                         ###
###############################

#No mostrar warnings en la terminal
warnings.filterwarnings("ignore")

#Tokenizador y StopWords
tokenizer = re.compile(r"\W+")
swEn = ['i', 'me', 'my', 'myself', 'we', 'our', 'ours', 'ourselves', 'you', 'your', 'yours', 'yourself', 'yourselves', 'he', 'him', 'his', 'himself', 'she', 'her', 'hers', 'herself', 'it', 'its', 'itself', 'they', 'them', 'their', 'theirs', 'themselves', 'what', 'which', 'who', 'whom', 'this', 'that', 'these', 'those', 'am', 'is', 'are', 'was', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'having', 'do', 'does', 'did', 'doing', 'a', 'an', 'the', 'and', 'but', 'if', 'or', 'because', 'as', 'until', 'while', 'of', 'at', 'by', 'for', 'with', 'about', 'against', 'between', 'into', 'through', 'during', 'before', 'after', 'above', 'below', 'to', 'from', 'up', 'down', 'in', 'out', 'on', 'off', 'over', 'under', 'again', 'further', 'then', 'once', 'here', 'there', 'when', 'where', 'why', 'how', 'all', 'any', 'both', 'each', 'few', 'more', 'most', 'other', 'some', 'such', 'no', 'nor', 'not', 'only', 'own', 'same', 'so', 'than', 'too', 'very', 's', 't', 'can', 'will', 'just', 'don', 'should', 'now', 'd', 'll', 'm', 'o', 're', 've', 'y', 'ain', 'aren', 'couldn', 'didn', 'doesn', 'hadn', 'hasn', 'haven', 'isn', 'ma', 'mightn', 'mustn', 'needn', 'shan', 'shouldn', 'wasn', 'weren', 'won', 'wouldn']
swEs = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'sí', 'porque', 'esta', 'entre', 'cuando', 'muy', 'sin', 'sobre', 'también', 'me', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'durante', 'todos', 'uno', 'les', 'ni', 'contra', 'otros', 'ese', 'eso', 'ante', 'ellos', 'e', 'esto', 'mí', 'antes', 'algunos', 'qué', 'unos', 'yo', 'otro', 'otras', 'otra', 'él', 'tanto', 'esa', 'estos', 'mucho', 'quienes', 'nada', 'muchos', 'cual', 'poco', 'ella', 'estar', 'estas', 'algunas', 'algo', 'nosotros', 'mi', 'mis', 'tú', 'te', 'ti', 'tu', 'tus', 'ellas', 'nosotras', 'vosostros', 'vosostras', 'os', 'mío', 'mía', 'míos', 'mías', 'tuyo', 'tuya', 'tuyos', 'tuyas', 'suyo', 'suya', 'suyos', 'suyas', 'nuestro', 'nuestra', 'nuestros', 'nuestras', 'vuestro', 'vuestra', 'vuestros', 'vuestras', 'esos', 'esas', 'estoy', 'estás', 'está', 'estamos', 'estáis', 'están', 'esté', 'estés', 'estemos', 'estéis', 'estén', 'estaré', 'estarás', 'estará', 'estaremos', 'estaréis', 'estarán', 'estaría', 'estarías', 'estaríamos', 'estaríais', 'estarían', 'estaba', 'estabas', 'estábamos', 'estabais', 'estaban', 'estuve', 'estuviste', 'estuvo', 'estuvimos', 'estuvisteis', 'estuvieron', 'estuviera', 'estuvieras', 'estuviéramos', 'estuvierais', 'estuvieran', 'estuviese', 'estuvieses', 'estuviésemos', 'estuvieseis', 'estuviesen', 'estando', 'estado', 'estada', 'estados', 'estadas', 'estad', 'he', 'has', 'ha', 'hemos', 'habéis', 'han', 'haya', 'hayas', 'hayamos', 'hayáis', 'hayan', 'habré', 'habrás', 'habrá', 'habremos', 'habréis', 'habrán', 'habría', 'habrías', 'habríamos', 'habríais', 'habrían', 'había', 'habías', 'habíamos', 'habíais', 'habían', 'hube', 'hubiste', 'hubo', 'hubimos', 'hubisteis', 'hubieron', 'hubiera', 'hubieras', 'hubiéramos', 'hubierais', 'hubieran', 'hubiese', 'hubieses', 'hubiésemos', 'hubieseis', 'hubiesen', 'habiendo', 'habido', 'habida', 'habidos', 'habidas', 'soy', 'eres', 'es', 'somos', 'sois', 'son', 'sea', 'seas', 'seamos', 'seáis', 'sean', 'seré', 'serás', 'será', 'seremos', 'seréis', 'serán', 'sería', 'serías', 'seríamos', 'seríais', 'serían', 'era', 'eras', 'éramos', 'erais', 'eran', 'fui', 'fuiste', 'fue', 'fuimos', 'fuisteis', 'fueron', 'fuera', 'fueras', 'fuéramos', 'fuerais', 'fueran', 'fuese', 'fueses', 'fuésemos', 'fueseis', 'fuesen', 'sintiendo', 'sentido', 'sentida', 'sentidos', 'sentidas', 'siente', 'sentid', 'tengo', 'tienes', 'tiene', 'tenemos', 'tenéis', 'tienen', 'tenga', 'tengas', 'tengamos', 'tengáis', 'tengan', 'tendré', 'tendrás', 'tendrá', 'tendremos', 'tendréis', 'tendrán', 'tendría', 'tendrías', 'tendríamos', 'tendríais', 'tendrían', 'tenía', 'tenías', 'teníamos', 'teníais', 'tenían', 'tuve', 'tuviste', 'tuvo', 'tuvimos', 'tuvisteis', 'tuvieron', 'tuviera', 'tuvieras', 'tuviéramos', 'tuvierais', 'tuvieran', 'tuviese', 'tuvieses', 'tuviésemos', 'tuvieseis', 'tuviesen', 'teniendo', 'tenido', 'tenida', 'tenidos', 'tenidas', 'tened']


# Disable Print
def blockPrint():
    sys.stdout = open(os.devnull, 'w')

# Restore Print
def enablePrint():
    sys.stdout = sys.__stdout__






























############################################################################
###                                                                      ###
###             PARTE 1: Obtención de Variables para una URL             ###
###                                                                      ###
############################################################################





###################################################################################################################################################

#############################################
###                                       ###
###       Parte 1.1 Función General       ###
###                                       ###
#############################################



def crearTabla(n,aTable,cTable,ListGdes,ListGdes1,headers,keywords,id):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Obtiene los valores de las variables SEO para una URL específica.

    param:  "n": Posición de la URL en la Query.  
            "aTable": Tabla Variables SEO.
            "cTable": Tabla Datos Keywords.
            "ListGdes": Listado de URLs de Google Search donde se encontraron las URLs de ListaGdes1.
            "ListGdes1": Listado de las primeras URLs para una query.
            "headers": Encabezados utilizados para que los requests no den problemas.
            "keywords": Keywords para la query a la que pertence la URL.
            "id": Identificador del Thread, util para comporbar como se van ejecutando los resultados de las queries (Desarrollador).

    """

    #Variables locales a la Iteracion
    numg=n
    gdes=ListGdes[n]
    gdes1=ListGdes1[n]

    #Posicionamiento
    aTable[numg][0]=numg+1
    cTable[numg][0]=numg+1

    #URL & Longitud
    #print(gdes1)
    #print("Longitud Url: %d" %(len(gdes1)))
    aTable[numg][1]=gdes1
    cTable[numg][1]=gdes1
    aTable[numg][2]=len(gdes1)


    #Fecha de Creacion (&Dias)
    CDate, DAlive = creacion(gdes1)
    aTable[numg][7]=CDate
    aTable[numg][8]=DAlive

    #Scheme
    gdesScheme = urlparse(gdes1).scheme
    #print("Protocol: %s" %(gdesScheme))
    aTable[numg][9]=gdesScheme


    #Comprbación de URL Amigable
    FriUrl = urlAmigable(gdes1)
    aTable[numg][10]=FriUrl


    #Busqueda Robots.txt y Sitemap.xml
    IsRob, IsSite = robotsNsitemaps(gdes1, gdesScheme, headers)
    aTable[numg][11]=IsRob
    aTable[numg][12]=IsSite

    #Compatibilidad con Dispositivos Móviles
    compM=compMovil(gdes1)
    aTable[numg][13]=compM

    #Overall Performance Score & LCP Load Speed
    mop, mls, dop, dls = score(gdes1)
    aTable[numg][14]=mop
    aTable[numg][15]=mls
    aTable[numg][16]=dop
    aTable[numg][17]=dls


    #Probar conexión a la URL
    try:
        responseg = requests.get(gdes1, headers=headers, timeout=60)
        status_code = responseg.status_code
        if status_code < 400:

            #Creamos la Sopa
            respDes = responseg
            soupDes = BeautifulSoup(respDes.content, "html.parser")


            #Titulo/Descripcion & Longitud
            titulo, tituloL, descripcion, descripcionL = MetaTD(gdes, gdes1, headers)
            aTable[numg][3]=titulo
            aTable[numg][4]=tituloL
            aTable[numg][5]=descripcion
            aTable[numg][6]=descripcionL
            #print("")


            #Existencia Encabezados h1-h6
            encabezado = encabezados(soupDes)
            aTable[numg][18]=encabezado


            #Tipo de Datos Estructurados y Marcado Schema
            markup, mschema = DatosEstructurados(soupDes)
            aTable[numg][19]=markup
            aTable[numg][20]=mschema


            #Cantidad de Imagenes y Videos
            images = soupDes.find_all('img')
            videos = soupDes.find_all('video')
            #print("Numero Imagenes: %d" %(len(images)))
            #print("Numero Videos: %d" %(len(videos)))
            aTable[numg][21]=len(images)
            aTable[numg][22]=len(videos)


            #Links Internos/Externos
            linksT, linksI, linksE, linksIF, linksEF, linksIN, linksEN, error404, errorC = linkCount(soupDes, gdes1, headers)
            aTable[numg][23]=linksT
            aTable[numg][24]=linksI
            aTable[numg][25]=linksE
            aTable[numg][26]=linksIF
            aTable[numg][27]=linksEF
            aTable[numg][28]=linksIN
            aTable[numg][29]=linksEN
            aTable[numg][30]=error404
            aTable[numg][31]=errorC

            #Coincidencia/Posicionamiento Keywords
            #print("Keyword Coincidencias")
            TOcurrences, POcurrences = KeywordSearch(soupDes, "title", keywords, str(titulo),cTable,numg)
            aTable[numg][32]=TOcurrences
            aTable[numg][33]=POcurrences
            TOcurrences, POcurrences =KeywordSearch(soupDes, "descripcion", keywords, str(descripcion),cTable,numg)
            aTable[numg][34]=TOcurrences
            aTable[numg][35]=POcurrences
            TOcurrences, POcurrences =KeywordSearch(soupDes, "h1", keywords, "",cTable,numg)
            aTable[numg][36]=TOcurrences
            aTable[numg][37]=POcurrences
            TOcurrences, POcurrences =KeywordSearch(soupDes, "alt", keywords, "",cTable,numg)
            aTable[numg][38]=TOcurrences
            aTable[numg][39]=POcurrences
            TOcurrences, POcurrences =KeywordSearch(soupDes, "src", keywords, "",cTable,numg)
            aTable[numg][40]=TOcurrences
            aTable[numg][41]=POcurrences
            TOcurrences, POcurrences =KeywordSearch(soupDes, "body", keywords, "",cTable,numg)
            aTable[numg][42]=TOcurrences
            aTable[numg][43]=POcurrences



    except requests.exceptions.RequestException:
        pass
        #print("Unable to Handle... Timeout error")

    return


    #Saber por terminal la busqueda terminada
    #print("%d (%d)" %(numg+1, id+1))










###################################################################################################################################################

#############################################
###                                       ###
###    Parte 1.2: Funciónes auxiliares    ###
###                                       ###
#############################################



def creacion(link):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Obtiene los Days Alive y Fecha de Creacción de una URL específica a partir de la API de whois.

    param:  "link": URL a analizar en crearTabla.

    return: "CDate": Fecha de Creación de la URL.
            "CAlive": Dias desde la creación de la URL.

    """

    hoy=datetime.datetime.now().date()

    try:
        wis = whois.whois(link)
        if isinstance(wis.creation_date, list):
            wisd=wis.creation_date[0].date()
            wisdold=hoy-wisd
        elif wis.creation_date is None:
            #print("Creacion: Desconocido")
            #print("Dias desde creacion: Desconocido")
            return "Unknown", "Unknown"
        else:
            wisd=wis.creation_date.date()
            wisdold=hoy-wisd
    except Exception:
        return "Unknown", "Unknown"
    #print("Creacion: %s" %(wisd))
    #print("Dias desde creacion: %s" %(str(wisdold.days)))

    return str(wisd), str(wisdold.days)





def urlAmigable(link):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Decide a partir de un listado basico de simbolos y de si incorpora segmentos concretos de una URL el si la URL especificada
    es amigable o no,

    param:  "link": URL a analizar en crearTabla.

    return: "friURL": Indica si la URL es amigable en formato string (Yes/No).

    """

    noAmigable=['_', '&', '.','=','@','#','$','%','?','!',';',':','+','*']
    Path=str(urlparse(link).path)
    if any([x in Path for x in noAmigable]):
        #print("Url: No amigable")
        return "No"
    else:
        if(not(urlparse(link).query == "" and urlparse(link).fragment == "")):
            #print("Url: No amigable")
            return "No"
        else:
            #print("Url: Amigable")
            return "Yes"





def robotsNsitemaps(link, scheme, headers):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Busca en la raiz de la URL si existe robots.txt y busca sitmap.xml tanto en robots.txt como en la raiz.

    param:  "link": URL a analizar en crearTabla.
            "scheme": Protocolo de la URL a analizar (http/https normalmente).
            "headers": Encabezados utilizados para que los requests no den problemas.
            
    return: "IsRob": Indica si se ha encontrado robots.txt en la url en formato string (Yes/No/Unknown).
            "IsSite": Indica si se ha encontrado sitemap.xml en la url en formato string (Yes/No/Unknown).

    """

    #Busqueda Robots.txt
    robStr=""
    IsRob="No"
    try:
        gdesRob= scheme + "://" + urlparse(link).netloc + "/robots.txt"
        responseg = requests.get(gdesRob, headers=headers, timeout=10)
        status_codeR = responseg.status_code

        if status_codeR >= 400:
            isRob="No"
            #print("Robots.txt: No")
        else:
            #print("Robots.txt: Si")
            IsRob="Yes"
            robStr=str(BeautifulSoup(responseg.content, "html.parser"))
    except requests.exceptions.RequestException:
        #print("Robots.txt: Timeout error")
        IsRob="Unknown"


    #Busqueda Sitemap.xml
    try:
        gdesSmap= scheme + "://" + urlparse(link).netloc + "/sitemap.xml"
        responseg = requests.get(gdesSmap, headers=headers, timeout=10)
        status_codeS = responseg.status_code
    except requests.exceptions.RequestException:
        status_codeS=0

    if "Sitemap:" in robStr:
        robStr=robStr.count("Sitemap:")
        #print("Sitemap.xml: Si (%d)" %(robStr))
        return IsRob, "Yes (" + str(robStr) + ")"
    elif status_codeS < 400:
        #print("Sitemap.xml: Si (1)")
        return IsRob, "Yes (1)"
    else:
        #print("Sitemap.xml: No")
        return IsRob, "No"





def compMovil(link):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Comprueba si la web de la URL es amigable para un entorno movil a partir de googleapis.

    param:  "link": URL a analizar en crearTabla.
            
    return: "compM": Indica si la URL es amigable para moviles en formato string (MOBILE_FIRENDLY/No_MOBILE_FRIENDLY/Unknown).

    """

    MFurl = 'https://searchconsole.googleapis.com/v1/urlTestingTools/mobileFriendlyTest:run'
    MFparams = {

                'url': link,

                'key': "GKEY_PLACEHOLDER"

            }
    try:
        MFx = requests.post(MFurl, data = MFparams)
        MFdata = json.loads(MFx.text)
        if MFdata["testStatus"]["status"] == "PAGE_UNREACHABLE":
            #print("Error Compatibilidad Moviles: %s" %(MFdata["testStatus"]["status"]))
            return "Unknown"
        else:
            #print("Compatibilidad Moviles: %s" %(MFdata["mobileFriendliness"]))
            return MFdata["mobileFriendliness"]
    except Exception:
        return "Unknown"





def score(link):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Se obtiene la Velocidad de Ejecucion (LCP) y el Puntuación de Rendimiento a paritr de la informacion obtenida de googleapis runPageSpeed.

    param:  "link": URL a analizar en crearTabla.
            
    return: "mop": Indica en un valor del 0-100 la Puntuación de Rendimiento en dispositivos moviles de la web de la URL especificada.
            "mls": Velocidad de Ejecucion en segundos en dispositivos moviles para la web de la URL especificada.
            "dop": Indica en un valor del 0-100 la Puntuación de Rendimiento en dispositivos sobremesa de la web de la URL especificada.
            "dls": Velocidad de Ejecucion en segundos en dispositivos sobremesa para la web de la URL especificada.

    """

    #Google Api runspeed Links
    urlMobil = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=" + link  + "&strategy=mobile&locale=es&key=GKEY_PLACEHOLDER"
    urlDesktop = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=" + link  + "&strategy=desktop&locale=es&key=GKEY_PLACEHOLDER"

    #Datos Mobil
    #print("Movil:")
    try:
        responseMobil = urlopen(urlMobil)
        dataMobil = json.loads(responseMobil.read().decode('utf-8'))
        overall_scoreMobil = dataMobil["lighthouseResult"]["categories"]["performance"]["score"] * 100
        overall_scoreMobil=int(overall_scoreMobil)
        lcpMobil = dataMobil["lighthouseResult"]["audits"]["largest-contentful-paint"]["displayValue"]
        #print("Overall Performance %d/100" %(overall_scoreMobil))
        #print("Load Speed (LCP): %s" %(lcpMobil))
    except Exception:
        overall_scoreMobil = "Unknown"
        lcpMobil = "Unknown"


    #Datos Desktop
    #print("Desktop:")
    try:
        responseDesktop = urlopen(urlDesktop)
        dataDesktop = json.loads(responseDesktop.read().decode('utf-8'))
        overall_scoreDesktop = dataDesktop["lighthouseResult"]["categories"]["performance"]["score"] * 100
        overall_scoreDesktop = int(overall_scoreDesktop)
        lcpDesktop = dataDesktop["lighthouseResult"]["audits"]["largest-contentful-paint"]["displayValue"]
        #print("Overall Performance %d/100" %(overall_scoreDesktop))
        #print("Load Speed (LCP): %s" %(lcpDesktop))
    except Exception:
        overall_scoreDesktop = "Unknown"
        lcpDesktop = "Unknown"

    return overall_scoreMobil, lcpMobil, overall_scoreDesktop, lcpDesktop





def MetaTD(glink, link, headers):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Se obtiene el Meta Titulo y Meta Descripción mostrados en Google Search ademas de sus longitudes para una URL especificada.

    param:  "glink": URL de la pagina de Google search en donde aparece la URL a analizar.
            "link": URL a analizar en crearTabla.
            "headers": Encabezados utilizados para que los requests no den problemas.
            
    return: "titulo": Muestra en formato String el Meta Titulo de la URL especificada.
            "tituloL": Muestra la longitud en carácteres del Meta Titulo.
            "descripcion, descripcionL": Muestra en formato String la Meta Descripción de la URL especificada.
            "descripcionL": Muestra la longitud en carácteres de la Meta Descripción.

    """

    #Meta Titulo & Logitud
    #print("Titulo Longitud")
    #print("Longitud: %d   Palabras %d" %(len(tit),len(tit.split())))
    tit=""
    gdes2=glink.find_all("h3", class_="LC20lb MBeuO DKV0Md")
    if(len(gdes2) >=1):
        tit=gdes2[-1].get_text()


    #Meta Descripcion & Logitud
    text=""
    gbool=True
    gdes2=glink.find_all("div", class_="VwiC3b yXK7lf MUxGbd yDYNvb lyLwlc lEBKkf")
    if(len(gdes2) >=1):
        gdes3=gdes2[-1].find_all("span")
        if(len(gdes3) >=1):
            gdes3 = str(gdes3[-1])
            if not gdes3.startswith("<span class"):
                gbool=False
                gdes3=gdes3.replace('<span>','')
                gdes3=gdes3.replace('</span>','')
                gdes3=gdes3.replace('<em>','')
                gdes3=gdes3.replace('</em>','')
                text=gdes3
                #print("Longitud: %d   Palabras %d" %(len(gdes3),len(gdes3.split())))
        else:
            gbool=False
            gdes3=gdes2[-1].get_text()
            text=gdes3
            #print("Longitud: %d   Palabras %d" %(len(gdes3),len(gdes3.split())))

    #Busqueda a la fuerza en caso extremo
    if(gbool):
        try:
            responseg = requests.head(link, headers=headers, timeout=10)
            status_code = responseg.status_code
            if status_code < 400:
                htmlg =  requests.get(link, headers=headers, timeout=10)
                soupg1 = BeautifulSoup(htmlg.content, features="html.parser")
                metas = soupg1.find_all('meta') #Get Meta Description
                for m in metas:
                    if m.get ('name') == 'description':
                        text = m.get('content')
                        #print("Longitud: %d   Palabras %d" %(len(text),len(text.split())))
        except Exception:
            pass
            #print("Fallo en la obtencion de la Descripcion")

    return str(tit), len(tit), str(text), len(text)





def encabezados(soup):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Obtiene los headers (h1-h6) utilizados para la web de la URL específicada.

    param:  "soup": Sopa en la que se gurda la información HTML de la URL.

    return: "encabezado": Lista de los encabezados usados.

    """

    #Comprobacion Exisistencia h1-h6
    encabezados = []
    if(not(soup.h1 == None)):
        encabezados.append("h1")
    if(not(soup.h2 == None)):
        encabezados.append("h2")
    if(not(soup.h3 == None)):
        encabezados.append("h3")
    if(not(soup.h4 == None)):
        encabezados.append("h4")
    if(not(soup.h5 == None)):
        encabezados.append("h5")
    if(not(soup.h6 == None)):
        encabezados.append("h6")

    #print("Encabezados usados: %s" %(str(encabezados)))
    return str(encabezados)





def DatosEstructurados(soup):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Busca el tipo de Datos Estructurados de la web de la URL y si utiliza Markado Schema.

    param:  "soup": Sopa en la que se gurda la información HTML de la URL.
            
    return: "markup": Indica que tipo de Datos Estructurados tienen la web de la URL especificada (JSON-LD/Microdata/RDFa/No).
            "mschema": Indica si se utiliza Markado Schema en la web en formato string (Yes/No).

    """

    #Tipo de Datos Estructurados
    strSoup=str(soup).lower()
    strDS = "Datos estructurados:"
    strDSB=True
    if "application/ld+json" in strSoup:
        strDs= strDS + " JSON-LD"
    elif "itemscope itemtype" in strSoup:
        strDs= strDS + " Microdata"
    elif "vocab=" in strSoup:
        strDs= strDS + " RDFa"
    else:
        strDs= strDS + " No"
        strDSB=False

    #Marcado Schema
    if strDSB and "schema.org" in strSoup:
        #print("%s (Marcado Schema)" %(strDs))
        return strDs, "Yes"
    else:
        #print(strDs)
        return strDs, "No"





def linkCount(soup, link, headers):

    """
    Num ejecuciones: Num Queries * Num Resultados + 1

    Se obtiene el numero de links Totales/Internos(NoFollow y Follow)/Externos(NoFollow y Follow)
    y el numero de Errores (404 y por conexión) encontrados en la web de la URL especificada.

    param:  "soup": Sopa en la que se gurda la información HTML de la URL.
            "link": URL a analizar en crearTabla.
            "headers": Encabezados utilizados para que los requests no den problemas.
            
    return: "linksT": Número de links Totales encontrados. 
            "linksI": Número de links Internos encontrados. 
            "linksE": Número de links Externos encontrados. 
            "linksIF": Número de links Internos con etiqueta Follow encontrados. 
            "linksEF": Número de links Externos con etiqueta Follow encontrados. 
            "linksIN": Número de links Internos con etiqueta NoFollow encontrados. 
            "linksEN": Número de links Externos con etiqueta NoFollow encontrados. 
            "error404": Número de Errores 404 encontrados en los links.
            "errorC": Número de Errores por Conexión encontrados en los links.

    """

    #Parametros iniciales
    DesIF=0
    DesEF=0
    DesIN=0
    DesEN=0

    extracted = tldextract.extract(link)
    extracted = "{}.{}".format(extracted.domain, extracted.suffix)

    urlPart = urlsplit(link)
    netl = urlPart.netloc
    schem = urlPart.scheme

    numerror=0
    timero=0


    for a in soup.find_all("a"):
        href=a.get('href')

        if(href == None): continue

        if(bool(urlparse(href).netloc)):
            extracted2 = tldextract.extract(href)
            extracted2 = "{}.{}".format(extracted2.domain, extracted2.suffix)

        #Calificación Externo/Interno Follow/NoFollow
        if(not bool(urlparse(href).netloc) or extracted == extracted2):
            try:
                if(a.get('rel') is None or a.get('rel')[0] != 'nofollow'):
                    DesIF=DesIF+1
                else:
                    DesIN=DesIN+1
            except IndexError:
                DesIF=DesIF+1
        else:
            try:
                if(a.get('rel') is None or a.get('rel')[0] != 'nofollow'):
                    DesEF=DesEF+1
                else:
                    DesEN=DesEN+1
            except IndexError:
                DesEF=DesEF+1


        if(urlparse(href).netloc == ''): href = schem + '://' + netl +href
        elif(urlparse(href).scheme == ''): href = schem + ':'  +href


        #Calificación Dead Links
        try:
            response = requests.head(href, headers=headers, timeout=10)
            status_code = response.status_code
            #print(status_code)
            if 'redlink=1' in href or status_code == 404 :
                if 'redlink=1' in href or requests.get(href).status_code == 404:
                    numerror = numerror + 1
            elif status_code >= 400:
                timero = timero + 1
        except Exception:
            timero = timero + 1


    #Muestreo
    DesTot = DesEF+DesEN+DesIF+DesIN

    return DesTot, DesIF+DesIN, DesEF+DesEN, DesIF, DesEF, DesIN, DesEN, numerror, timero





def KeywordSearch(soup, tipo, keywords, aux, cTable, numg):

    """
    Num ejecuciones: (Num Queries * Num Resultados + 1) * 6

    Busca las Keywords seleccionadas (de manera Parcial y Total) dentro de un segmento específico de la web de la URL especificada.

    param:  "soup": Sopa en la que se gurda la información HTML de la URL.
            "tipo": Segmento en el que se buscaran las Keywords (Body/Title/H1/alt/src/descripción)
            "keywords": Keywords para la query a la que pertence la URL especificada.
            "aux": En caso de que seleccionemos Titulo o Descripción, pasaremos su texto que ya habiamos obtenido previamente (por conveniencia).
            "cTable": Tabla Datos Keywords.
            "numg": Posición de la URL en la Query.

    return: "TOcurrences": Ocurrencias Totales de las Keywords encontradas en el segmento especificado. 
            "POcurrences": Ocurrencias Parciales de las Keywords encontradas en el segmento especificado (Usando Levenshtein). 

    """

    #Seleccion de Texto
    CTotales = 0
    CParciales = 0
    tex=""
    try:
        if tipo == "body":
            tex = soup.body.get_text()
            cId=len(keywords)*4
        elif tipo == "title":
            tex= aux
            cId=0
        elif tipo == "h1":
            if not(soup.h1 == None):
                tex = soup.h1.get_text()
                cId=len(keywords)*6
            else:
                #print("H1 No Existe \n\n")
                return CTotales, CParciales
        elif tipo == "alt":
            for img in soup.find_all('img', alt=True):
                try:
                    alts=img['alt']
                    tex=tex+" "+alts
                except KeyError:
                    tex=tex
            cId=len(keywords)*8
        elif tipo == "src":
            for img in soup.find_all('img', alt=True):
                try:
                    srcs=img['src']
                    tex=tex+" "+srcs
                except KeyError:
                    tex=tex
            cId=len(keywords)*10
        else:
            tex = aux
            cId=len(keywords)*2
    except AttributeError:
        return CTotales, CParciales


    text = tokenizer.sub(' ', tex.lower()).split()
    for word in keywords:
        indices=[]
        indices2=[]
        indices3=[]

        #Coincidencias Totales
        if word in text:
            indices = [i for i, x in enumerate(text) if x == word]
        if(len(word)==4):
            indices2 = [i for i, x in enumerate(text) if dp_levenshtein_threshold(word,x,3) == 1]
        elif(len(word)==5):
            indices2 = [i for i, x in enumerate(text) if dp_levenshtein_threshold(word,x,3) in range(1,3)]
        elif(len(word)>5):
            indices2 = [i for i, x in enumerate(text) if dp_levenshtein_threshold(word,x,3) in range(1,4)]


        #Datos de Tabla de Keywords
        cId=cId+2
        cStrT= "(" + str(len(indices)) + ") " + str(indices)
        indicesAux=indices2+indices3
        indicesAux.sort()
        cStrP= "(" + str(len(indices2)+len(indices3)) + ") " + str(indicesAux)
        cTable[numg][cId]= cStrT
        cTable[numg][cId+1]= cStrP

        #Cuantificación Coincidencias
        CTotales = CTotales + len(indices)
        CParciales = CParciales + len(indices2)+len(indices3)
    lt=len(text)
    if lt==0: lt=1
    return (CTotales/lt)*100, (CParciales/lt)*100





def dp_levenshtein_threshold(x, y, th):

    """
    Num ejecuciones: (Num Queries * Num Resultados + 1) * 6 * Num Keywords * (Num palabras de cada segmento)

    Calcula la Distancia de Levenshtein de dos palabras a partir de un threshold especifico.

    param:  "x": Palabra 1.
            "y": Palabra 2.
            "th": Threshold límite.

    return: Distancia obtenida

    """

    current_row = [None] * (1+len(x))
    previous_row = [None] * (1+len(x))
    current_row[0] = 0
    for i in range(1, len(x)+1):
        current_row[i] = current_row[i-1] + 1
    for j in range(1, len(y)+1):
        previous_row, current_row = current_row, previous_row
        current_row[0] = previous_row[0] + 1
        for i in range(1, len(x)+1):
            current_row[i] = min( current_row[i-1] + 1,
                previous_row[i] + 1,
                previous_row[i-1] + (x[i-1] != y[j-1]))
        if(min(current_row) > th):
            return th+1
    return min(current_row[len(x)],th+1)



###################################################################################################################################################



































############################################################################
###                                                                      ###
###             PARTE 2: Creación de Tablas para las Queries             ###
###                                                                      ###
############################################################################





###################################################################################################################################################

#############################################
###                                       ###
###   Parte 2.1 Crador Tablas Generales   ###
###                                       ###
#############################################



def principal(DataTable,DataTY,DataTY10,q,num,keywords,workbook,limitThread,id,bTable,dTable,limiteInf,idioma):

    """
    Num ejecuciones: Num Queries 

    Función desde donde se lanzan los threads para obtener las variables SEO de una query especifica 
    y crear las tablas en xmlx correspondientes de dicha query.

    param:  "DataTable": Tabla de datos para el clasificador.
            "DataTY": Etiquetas para el clasificador.
            "DataTY10": Etiquetas para el clasificador simple.
            "q": Query a analizar.
            "num": Numero de resultados que buscar en la Query.
            "keywords": Keywords de la Query.
            "workbook": Woorkbook xlsx en el que mostrar los datos.
            "limitThread": Numero de threads simultáneos límite.
            "id": identificador del thread principal.
            "bTable": Tabla de Medias
            "dTable": Tabla de Medias limitado a limiteInf
            "limiteInf": Posicionamiento inferior limite para la Table Media de mejores resultados
            "idioma": Idioma seleccionado (Ingles/Español).
            

    """

    #columns = ['PageRank', 'Url', 'Url Length', 'Title', 'Title Length', 'Description', 'Description Length', 'Creation Date', 'Days Alive', 'Schema', 'Friendly Url', 'Robots.txt', 'Sitemap.xml', 'Mobile Compatibility', 'Mobile Overall Performance', 'Mobile Load Speed', 'Desktop Oveall Performance', 'Desktop Load Speed', 'Headers', 'Structed Data', 'Schema Markup', 'Total Images', 'Total Videos', 'Total Links', 'Internal Links', 'External Links', 'Internal Follow', 'External Follow', 'Internal NoFollow', 'External NoFollow', 'Errors 404', 'Conection Errors', 'Title Total Occurrences (Keywords)', 'Title Partial Occurrences (Keywords)', 'Description Total Occurrences (Keywords)', 'Description Partial Occurrences (Keywords)', 'H1 Total Occurrences (Keywords)', 'H1 Partial Occurrences (Keywords)', 'Alt Total Occurrences (Keywords)', 'Alt Partial Occurrences (Keywords)', 'Src Total Occurrences (Keywords)', 'Src Partial Occurrences (Keywords)', 'Body Total Occurrences (Keywords)', 'Body Partial Occurrences (Keywords)']
    aTable = list(map(lambda _: ["Unknown"]*44, range(num)))
    clength=2 + len(keywords)*12
    cTable = list(map(lambda _: ["Unknown"]*clength, range(num)))

    #Creamos los Headers
    q = q.replace(' ', '+')
    URL = "https://google.com/search?q="+q  + "&num=10"
    USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.14; rv:65.0) Gecko/20100101 Firefox/65.0"
    headers = {"user-agent" : USER_AGENT, "Cache-Control": "no-cache", "Pragma": "no-cache"}
    print("")


    #Primera página de google search
    prevURL=""
    numg=0
    nplus=0
    ListGdes=[]
    ListGdes1=[]
    ListNumg=[]

    try:
        #Sacamos el HTML de google search
        resp = requests.get(URL, headers=headers, timeout=60)
        if resp.status_code == 200:
            soupg = BeautifulSoup(resp.content, "html.parser")
        else:
            num=0
            print("429")


        #Busqueda URLs
        #print("\nPagina: %s" %(URL))
        URL2=URL
        while(numg<num):
            #Buqueda de URLs
            count=0
            for gdes in soupg.find_all("div", {"class":["tF2Cxc","jtfYYd"]}):
                if(numg==num):
                    break
                gdes1=gdes.find_all("a")
                gdes1=gdes1[0].get("href")

                #Evitar Sitelinks
                if(gdes1==prevURL or urlparse(gdes1).netloc == urlparse(prevURL).netloc):
                    continue
                prevURL=gdes1

                ListGdes.append(gdes)
                ListGdes1.append(gdes1)
                ListNumg.append(numg)

                numg=numg+1
                count=count+1
                #print(gdes1)
            if count==0:
                num=numg
            #Pasamos de Pagina en Google Search si es necesario
            if(numg<num):
                nplus=nplus+10
                URL2=URL + "&start=" + str(nplus)
                #print("Pagina: %s" %(URL2))
                resp = requests.get(URL2, headers=headers, timeout=60)
                if resp.status_code == 200:
                    soupg = BeautifulSoup(resp.content, "html.parser")
                else:
                    numg=num
        if(idioma):
            print("URLs Found (%d)" %(id+1))
        else:
            print("URLs Encontradas (%d)" %(id+1))

    except Exception:
        #print("Some URLs Not Found (%d)" %(id+1))
        pass



    #Lanzamos Threads CrearTabla
    try:
        lnumgLen=len(ListNumg)
        lnum1=0
        lnum2=limitThread
        while(lnumgLen>limitThread):
            threads = [threading.Thread(target=crearTabla, daemon=True, args=(n,aTable,cTable,ListGdes,ListGdes1,headers,keywords,id,)) for n in range(lnum1,lnum2)]
            for thread in threads:
                thread.start()
            t2=time.time()
            t1=600
            for thread in threads:
                if t1>5:
                    #print("Timeout: %d (%d)" %(t1, id+1))
                    thread.join(t1)
                    t1=600+t2-time.time()
                else:
                    #print("Timeout: 5 (%d)" %(t1, id+1))
                    thread.join(5)
            lnumgLen = lnumgLen - limitThread
            lnum1 = lnum2
            lnum2=lnum2+limitThread
            #gc.collect()
        if(lnumgLen>=1):
            lnum2=lnum1+lnumgLen
            threads = [threading.Thread(target=crearTabla, daemon=True, args=(n,aTable,cTable,ListGdes,ListGdes1,headers,keywords,id,)) for n in range(lnum1,lnum2)]
            for thread in threads:
                thread.start()
            t2=time.time()
            t1=600
            for thread in threads:
                if t1>5:
                    #print("Timeout: %d (%d)" %(t1, id+1))
                    thread.join(t1)
                    t1=600+t2-time.time()
                else:
                    #print("Timeout: 5 (%d)" %(t1, id+1))
                    thread.join(5)
            #gc.collect()
    except KeyboardInterrupt:
        sys.exit("User Quited")

    #for n in ListNumg: crearTabla(n)
    #print(aTable)




    #Datos worksheet excel local
    try:
        worksheet = workbook.add_worksheet(q)
    except Exception:
        worksheet = workbook.add_worksheet(str(id+1))

    worksheet.set_column('A:A', 19)
    worksheet.set_column('C:AR', 15)
    worksheet.set_column('B:B', 50)
    worksheet.set_column('D:D', 50)
    worksheet.set_column('G:G', 18)
    worksheet.set_column('F:F', 50)
    worksheet.set_column('H:H', 14)
    worksheet.set_column('N:N', 20)
    worksheet.set_column('O:T', 26)
    worksheet.set_column('U:U', 16)
    worksheet.set_column('V:Z', 14)
    worksheet.set_column('AA:AF', 17)
    worksheet.set_column('AG:AR', 38)

    #Creamos Tabla Principal para esta Busqueda
    #keywords3 = '+'.join(str(e) for e in keywords)
    if(idioma):columns = [{'header': 'Ranking'}, {'header': 'Url'}, {'header': 'Url Length'}, {'header': 'Title'}, {'header': 'Title Length'}, {'header': 'Description'}, {'header': 'Description Length'}, {'header': 'Creation Date'}, {'header': 'Days Alive'}, {'header': 'Protocol'}, {'header': 'Friendly Url'}, {'header': 'Robots.txt'}, {'header': 'Sitemap.xml'}, {'header': 'Mobile Compatibility'}, {'header': 'Mobile Overall Performance'}, {'header': 'Mobile Load Speed'}, {'header': 'Desktop Oveall Performance'}, {'header': 'Desktop Load Speed'}, {'header': 'Headers'}, {'header': 'Structed Data'}, {'header': 'Schema Markup'}, {'header': 'Total Images'}, {'header': 'Total Videos'}, {'header': 'Total Links'}, {'header': 'Internal Links'}, {'header': 'External Links'}, {'header': 'Internal Follow'}, {'header': 'Internal NoFollow'}, {'header': 'External Follow'}, {'header': 'External NoFollow'}, {'header': 'Errors 404'}, {'header': 'Conection Errors'}, {'header': 'Title Total Occurrences (Keywords%)'}, {'header': 'Title Partial Occurrences (Keywords%)'}, {'header': 'Description Total Occurrences (Keywords%)'}, {'header': 'Description Partial Occurrences (Keywords%)'}, {'header': 'H1 Total Occurrences (Keywords%)'}, {'header': 'H1 Partial Occurrences (Keywords%)'}, {'header': 'Alt Total Occurrences (Keywords%)'}, {'header': 'Alt Partial Occurrences (Keywords%)'}, {'header': 'Src Total Occurrences (Keywords%)'}, {'header': 'Src Partial Occurrences (Keywords%)'}, {'header': 'Body Total Occurrences (Keywords%)'}, {'header': 'Body Partial Occurrences (Keywords%)'}]
    else:columns = [{'header': 'Ranking'}, {'header': 'Url'}, {'header': 'Url Longitud'}, {'header': 'Título'}, {'header': 'Título Longitud'}, {'header': 'Descripción'}, {'header': 'Descripción Longitud'}, {'header': 'Fecha Creación'}, {'header': 'Dias Vivo'}, {'header': 'Protocolo'}, {'header': 'Url Amigable'}, {'header': 'Robots.txt'}, {'header': 'Sitemap.xml'}, {'header': 'Compatibilidad Móvil'}, {'header': 'Rendimiento Móvil'}, {'header': 'Velocidad Carga Móvil'}, {'header': 'Rendimiento Sobremesa'}, {'header': 'Velocidad Carga Sobremesa'}, {'header': 'Cabeceras'}, {'header': 'Datos Estructurados'}, {'header': 'Schema Markup'}, {'header': 'Imágenes Totales'}, {'header': 'Videos Totales'}, {'header': 'Links Totales'}, {'header': 'Links Internos'}, {'header': 'Links Externos'}, {'header': 'Follow Internos'}, {'header': 'NoFollow Internos'}, {'header': 'Follow Externos'}, {'header': 'NoFollow Externos'}, {'header': 'Errores 404'}, {'header': 'Errores Conexión'}, {'header': 'Título Ocurrencias Totales (Keywords%)'}, {'header': 'Título Ocurrencias Parciales (Keywords%)'}, {'header': 'Descripción Ocurrencias Totales (Keywords%)'}, {'header': 'Descripción Ocurrencias Parciales (Keywords%)'}, {'header': 'H1 Ocurrencias Totales (Keywords%)'}, {'header': 'H1 Ocurrencias Parciales (Keywords%)'}, {'header': 'Alt Ocurrencias Totales (Keywords%)'}, {'header': 'Alt Ocurrencias Parciales (Keywords%)'}, {'header': 'Src Ocurrencias Totales (Keywords%)'}, {'header': 'Src Ocurrencias Parciales (Keywords%)'}, {'header': 'Cuerpo Ocurrencias Totales (Keywords%)'}, {'header': 'Cuerpo Ocurrencias Parciales (Keywords%)'}]
    Celdas = 'A12:AR' + str(num+12)
    worksheet.add_table(Celdas, {'data': aTable, 'columns': columns})


    #Tabla Medias Local
    mediaL, DataL = tablaMedias(num,aTable)

    #Añade datos BD
    tabPos=[2,4,6,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43]
    for x in range(0,numBus):
        for y in range(0,39):
            y2=tabPos[y]
            x2=id*numBus+x
            try:
                DataTable[x2][y]=DataL[x][y2]
            except Exception:
                try:
                    DataTY[x2]=0
                    DataTY10[x2]=0
                except Exception:
                    print(len(DataTY))
                    print(id)
                    print(numBus)
                    print(x2)



    #Información adicional Excel
    bold2 = workbook.add_format({'bold': True, 'border': True,'align': 'center'})
    bold = workbook.add_format({'bold': True, 'align': 'center'})

    bTable[id]=mediaL
    mediaL=[mediaL]
    if(idioma):
        keyStr= "Limit Ranking: " + str(limiteInf)
    else:
        keyStr= "Limite Ranking: " + str(limiteInf)
    numC=num+13
    if(idioma): mediaCols = [{'header': 'Summary'}, {'header': 'Url'}, {'header': 'Url Length'}, {'header': 'Title'}, {'header': 'Title Length'}, {'header': 'Description'}, {'header': 'Description Length'}, {'header': 'Creation Date'}, {'header': 'Days Alive'}, {'header': 'Use Https'}, {'header': 'Friendly Url'}, {'header': 'Robots.txt'}, {'header': 'Sitemap.xml'}, {'header': 'Mobile Compatibility'}, {'header': 'Mobile Overall Performance'}, {'header': 'Mobile Load Speed'}, {'header': 'Desktop Oveall Performance'}, {'header': 'Desktop Load Speed'}, {'header': 'Headers'}, {'header': 'Structed Data'}, {'header': 'Schema Markup'}, {'header': 'Total Images'}, {'header': 'Total Videos'}, {'header': 'Total Links'}, {'header': 'Internal Links'}, {'header': 'External Links'}, {'header': 'Internal Follow'}, {'header': 'Internal NoFollow'}, {'header': 'External Follow'}, {'header': 'External NoFollow'}, {'header': 'Errors 404'}, {'header': 'Conection Errors'}, {'header': 'Title Total Occurrences (Keywords%)'}, {'header': 'Title Partial Occurrences (Keywords%)'}, {'header': 'Description Total Occurrences (Keywords%)'}, {'header': 'Description Partial Occurrences (Keywords%)'}, {'header': 'H1 Total Occurrences (Keywords%)'}, {'header': 'H1 Partial Occurrences (Keywords%)'}, {'header': 'Alt Total Occurrences (Keywords%)'}, {'header': 'Alt Partial Occurrences (Keywords%)'}, {'header': 'Src Total Occurrences (Keywords%)'}, {'header': 'Src Partial Occurrences (Keywords%)'}, {'header': 'Body Total Occurrences (Keywords%)'}, {'header': 'Body Partial Occurrences (Keywords%)'}]
    else: mediaCols = [{'header': 'Resumen'}, {'header': 'Url'}, {'header': 'Url Longitud'}, {'header': 'Título'}, {'header': 'Título Longitud'}, {'header': 'Descripción'}, {'header': 'Descripción Longitud'}, {'header': 'Fecha Creación'}, {'header': 'Dias Vivo'}, {'header': 'Usa Https'}, {'header': 'Url Amigable'}, {'header': 'Robots.txt'}, {'header': 'Sitemap.xml'}, {'header': 'Compatibilidad Móvil'}, {'header': 'Rendimiento Móvil'}, {'header': 'Velocidad Carga Móvil'}, {'header': 'Rendimiento Sobremesa'}, {'header': 'Velocidad Carga Sobremesa'}, {'header': 'Cabeceras'}, {'header': 'Datos Estructurados'}, {'header': 'Schema Markup'}, {'header': 'Imágenes Totales'}, {'header': 'Videos Totales'}, {'header': 'Links Totales'}, {'header': 'Links Internos'}, {'header': 'Links Externos'}, {'header': 'Follow Internos'}, {'header': 'NoFollow Internos'}, {'header': 'Follow Externos'}, {'header': 'NoFollow Externos'}, {'header': 'Errores 404'}, {'header': 'Errores Conexión'}, {'header': 'Título Ocurrencias Totales (Keywords%)'}, {'header': 'Título Ocurrencias Parciales (Keywords%)'}, {'header': 'Descripción Ocurrencias Totales (Keywords%)'}, {'header': 'Descripción Ocurrencias Parciales (Keywords%)'}, {'header': 'H1 Ocurrencias Totales (Keywords%)'}, {'header': 'H1 Ocurrencias Parciales (Keywords%)'}, {'header': 'Alt Ocurrencias Totales (Keywords%)'}, {'header': 'Alt Ocurrencias Parciales (Keywords%)'}, {'header': 'Src Ocurrencias Totales (Keywords%)'}, {'header': 'Src Ocurrencias Parciales (Keywords%)'}, {'header': 'Cuerpo Ocurrencias Totales (Keywords%)'}, {'header': 'Cuerpo Ocurrencias Parciales (Keywords%)'}]
    worksheet.add_table('A4:AR5', {'data': mediaL, 'columns':mediaCols})
    worksheet.write(0, 0, "Keywords: ", bold)
    if(idioma):
        worksheet.write(2, 0, "Mean Table", bold2)
        worksheet.write(6, 0, "Superior Mean Table", bold2)
        worksheet.write(10, 0, "Results Table", bold2)
        worksheet.write(numC, 0, "Keywords Table", bold2)
    else:
        worksheet.write(2, 0, "Tabla Medias", bold2)
        worksheet.write(6, 0, "Tabla Medias Superior", bold2)
        worksheet.write(10, 0, "Tabla Resultados", bold2)
        worksheet.write(numC, 0, "Tabla Keywords", bold2)
    worksheet.write(0, 1, str(keywords), bold)
    worksheet.write(6, 1, keyStr, bold)



    #Creamos Tabla de Keywords para esta Busqueda
    keyCols = [{'header': 'Ranking'}, {'header': 'Url'}]
    keyStr=""
    if(idioma): 
        ocurS = [" Total Occurrences", " Partial Occurrences"]
        ocurS2 = ["Title ","Description ","Body ","H1 ","Alt ","Src "]
    else: 
        ocurS = [" Ocurrencias Totales", " Ocurencias Parciales"]
        ocurS2 = ["Título ","Descripción ","Cuerpo ","H1 ","Alt ","Src "]
    for tipo in ocurS2:
        for key in keywords:
            for ocur in ocurS:
                keyStr= tipo + key + ocur
                keyStr=[{'header': keyStr}]
                keyCols = keyCols+keyStr
    letter = ''
    while clength > 25 + 1:
        letter += chr(65 + int((clength-1)/26) - 1)
        clength = clength - (int((clength-1)/26))*26
    letter += chr(65 - 1 + (int(clength)))
    Celdas = 'A' + str(num+15) +  ':'+ letter + str(num+num+15)

    worksheet.add_table(Celdas, {'data': cTable,'columns':keyCols})


    #Tabla Media superior Local + datos Global
    mediaL, none= tablaMedias(limiteInf,aTable)
    dTable[id]=mediaL
    mediaL=[mediaL]
    worksheet.add_table('A8:AR9', {'data': mediaL, 'columns':mediaCols})

    if(idioma):
        print("Finished query (%d)" %(id+1))
    else:
        print("Busqueda terminada (%d)" %(id+1))

    return










###################################################################################################################################################

#############################################
###                                       ###
###   Parte 2.2 Crador Tablas de Medias   ###
###                                       ###
#############################################



def tablaMedias(num,aTable):

    """
    Num ejecuciones: (Num Queries) * 2 + 1

    Función para crear una Tabla de Medias tanto para una query total como para aquella limitada a los Rankings superiores.

    param:  "num": Numero de resultados de la Query sobre la que crear la Tabla.
            "aTable": Tabla Variables SEO.

    return: "mediaL": Tabla de Medias.
            "dataL": Datos estilizados para el clasificador

    """

    #Creamos Tabla de Medias para esta Busqueda
    resTabla=[0]*44
    resTabla[0]="Medias"
    DataTableAux = list(map(lambda _: [0]*44, range(num)))

    tabPos=[2,4,6,8,14,16,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43]
    for i in range(len(tabPos)):
        tabP=tabPos[i]
        cont=num
        for k in range(num):
            ins = aTable[k][tabP]
            if "Unknown" in str(ins):
                cont=cont-1
                DataTableAux[k][tabP]=0
            else:
                if tabP>=32:
                    resTabla[tabP]=resTabla[tabP]+ ins
                    DataTableAux[k][tabP]=ins
                else:
                    resTabla[tabP]=resTabla[tabP]+ int(ins)
                    DataTableAux[k][tabP]=float(ins)
        if(cont==0):
            resTabla[tabP]=0
        else:
            resTabla[tabP]=resTabla[tabP]/cont

    tabPos=[9,10,11,12,13,19,20]
    for i in range(len(tabPos)):
        tabP=tabPos[i]
        cont=num
        for k in range(num):
            ins = str(aTable[k][tabP])
            if "Unknown" in ins:
                cont=cont-1
                DataTableAux[k][tabP]=0
            elif (not ("No" in ins)) or "https" in ins:
                resTabla[tabP]=resTabla[tabP]+ 1
                DataTableAux[k][tabP]=1
            else:
                DataTableAux[k][tabP]=0
        if(cont==0):
            resTabla[tabP]=0
        else:
            resTabla[tabP]=resTabla[tabP]/cont

    tabPos=[15,17]
    for i in range(len(tabPos)):
        tabP=tabPos[i]
        cont=num
        for k in range(num):
            ins = str(aTable[k][tabP])
            if "Unknown" in ins:
                cont=cont-1
                DataTableAux[k][tabP]=0
            else:
                ins=ins.replace(",",".")
                ins=ins.replace("s","")
                ins=ins.replace(" ","")
                resTabla[tabP]=resTabla[tabP]+ float(ins)
                DataTableAux[k][tabP]=float(ins)
        if(cont==0):
            resTabla[tabP]=0
        else:
            resTabla[tabP]=resTabla[tabP]/cont

    cont=num
    for k in range(num):
        ins = str(aTable[k][18])
        if "Unknown" in ins:
            cont=cont-1
            DataTableAux[k][18]=0
        else:
            resTabla[18]=resTabla[18]+ ins.count('h')
            DataTableAux[k][18]=ins.count('h')
    if(cont==0):
        resTabla[18]=0
    else:
        resTabla[18]=resTabla[18]/cont

    tabPos=[1,3,5,7]
    for i in tabPos:
        resTabla[i]="-"

    return resTabla, DataTableAux





def tablaMediasTotales(bdTable,numQuery,desN,desviaciónT):

    """
    Num ejecuciones: 2 

    Función para crear una Tabla de Medias Total a partir de las Tablas de Medias de cada query  
    (tanto toatales como aquellas imitada a los Rankings superiores).

    param:  "bdTable": Tabla de Medias o Tabla de Medias limitado a limiteInf
            "numQuery": Numero de queries buscadas.
            "desN": Booleano que decide si calcular la desviación típica.
            "desviaciónT": Array de desviación típica.

    return: "mediaT/mediaS": Tabla de Medias Total.

    """

    mediasLista = [0]*44
    mediasLista[0]="Medias"
    tabPos=[1,3,5,7]
    for i in tabPos:
        mediasLista[i]="-"

    tabPos=[2,4,6,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43]
    if(desN):
        for i in tabPos:
            auxVari = []
            for j in range(numQuery):
                auxVari.append(bdTable[j][i])
                mediasLista[i]=mediasLista[i]+bdTable[j][i]
            desviaciónT[i]=np.std(auxVari)
            mediasLista[i]=mediasLista[i]/numQuery
    else:
        for i in tabPos:
            for j in range(numQuery):
                mediasLista[i]=mediasLista[i]+bdTable[j][i]
            mediasLista[i]=mediasLista[i]/numQuery
    return mediasLista



###################################################################################################################################################



































############################################################################
###                                                                      ###
###             PARTE 3: Inicialización de Busqueda Realizada            ###
###                                                                      ###
############################################################################





###################################################################################################################################################

#############################################
###                                       ###
### Parte 3.1 Iniciador Busquedas & Excel ###
###                                       ###
#############################################



def inicializar(idioma,numQuery,numBus,limiteInf,keywordL,queryL,keyU,q2,keywords2):

    """
    Num ejecuciones: 1

    Función para lanzar los threads para crear las tablas y xmlx de cada query especifica, y para crear los xmlx de la media de medias
    y del infrome de una URL/Query específica.

    param:  "idioma": Idioma seleccionado (Ingles/Español).
            "numQuery": Numero de queries a buscar.
            "nuBus": Numero de resultados por query a buscar.
            "limiteInf": Ranking inferior limite para la Table Media de mejores resultados.
            "keywords": Listado de Keywords para cada Query.
            "queryL": Listado de queries a buscar.
            "KeyU": Booleano que indica si se quiere realizar un informe de una URL/Query. 
            "q2": URL/Query de la que realizar un informe/analísis.
            "keywords2": Keywords de la Query/URL q2 
            
    """


    t = time.time()
    print("\n")

    #Datos basicos excel
    hoy=datetime.datetime.now().date()

    path = os.getcwd()
    if(keyU):
        qstr= q2.replace(":","&")
        qstr= qstr.replace(".","&")
        qstr= qstr.replace("/","")
        if(idioma): path=path + '/Report/'
        else: path=path + '/Informe/'
        isExist = os.path.exists(path)
    else:
        path=path+'/Query/'
        isExist = os.path.exists(path)
    if not isExist:
        os.makedirs(path)

    global fileName
    if(idioma):
        if(keyU):
            fileName = "Report/" + str(hoy) + "_q=" + qstr + "_[" + datetime.datetime.now().strftime("%H;%M;%S") + "]" + ".xlsx"
        else:
            fileName = "Query/EN_" + str(hoy) + "_q=" + str(numQuery) + "_num=" + str(numBus) + "_[" + datetime.datetime.now().strftime("%H;%M;%S") + "]" + ".xlsx"
    else:
        if(keyU):
            fileName = "Informe/" + str(hoy) + "_q=" + qstr + "_[" + datetime.datetime.now().strftime("%H;%M;%S") + "]" + ".xlsx"
        else:
            fileName = "Query/ES_" + str(hoy) + "_q=" + str(numQuery) + "_num=" + str(numBus) + "_[" + datetime.datetime.now().strftime("%H;%M;%S") + "]" + ".xlsx"

    workbook = xlsxwriter.Workbook(fileName)
    if(idioma):
        worksheetI = workbook.add_worksheet('Report')
    else:
        worksheetI = workbook.add_worksheet('Informe')
    worksheetP = workbook.add_worksheet('Data')

    bTable = [0]*numQuery
    dTable = [0]*numQuery
    desviaciónT = [0]*44 



    #Data Clasifier Etiquetas
    DataTable = list(map(lambda _: [0]*39, range(numQuery*numBus))) 
    DataTableY = [0]*(numBus)
    DataTableY10 = [0]*(numBus)
    numMin = [0, 0, 0, 0, 0]
    numMin10 = [0, 0]
    numClases=5
    numClases10=2

    clfList=["Unknown"]*5
    clfList[0]=limiteInf
    for bd in range(limiteInf):
        DataTableY[bd]=1
        DataTableY10[bd]=1
        numMin[0]=numMin[0]+1
        numMin10[0]=numMin10[0]+1
    if(numBus>10):
        clfList[1]=10
        for bd in range(limiteInf,10):
            DataTableY[bd]=2
            DataTableY10[bd]=1
            numMin[1]=numMin[1]+1
            numMin10[0]=numMin10[0]+1

        if(numBus>=25):
            DataAux=int((numBus-10)/3)
            for bd2 in range(0,3):
                DataAux2=10+DataAux*bd2
                clfList[2+bd2]=DataAux2+DataAux
                for bd in range(DataAux2,DataAux2+DataAux):
                    DataTableY[bd]=bd2+3
                    DataTableY10[bd]=2
                    numMin[bd2+2]=numMin[bd2+2]+1
                    numMin10[1]=numMin10[1]+1
            DataAux3 = (numBus-10)%3
            for bd in range(DataAux2+DataAux,DataAux2+DataAux+DataAux3):
                DataTableY[bd]=5
                DataTableY10[bd]=2
                numMin[4]=numMin[4]+1
                numMin10[1]=numMin10[1]+1

        else:
            if(numBus>=20):
                numClases=4
                DataAux=int((numBus-10)/2)
                for bd2 in range(0,2):
                    DataAux2=10+DataAux*bd2
                    clfList[2+bd2]=DataAux2+DataAux
                    for bd in range(DataAux2,DataAux2+DataAux):
                        DataTableY[bd]=bd2+3
                        DataTableY10[bd]=2
                        numMin[bd2+2]=numMin[bd2+2]+1
                        numMin10[1]=numMin10[1]+1
                DataAux3 = (numBus-10)%2
                for bd in range(DataAux2+DataAux,DataAux2+DataAux+DataAux3):
                    DataTableY[bd]=5
                    DataTableY10[bd]=2
                    numMin[4]=numMin[4]+1
                    numMin10[1]=numMin10[1]+1
            else:
                numClases=3
                clfList[2]=numBus
                for bd in range(10,numBus):
                    DataTableY[bd]=3
                    DataTableY10[bd]=2
                    numMin[2]=numMin[2]+1
                    numMin10[1]=numMin10[1]+1

    else:
        numClases=2
        clfList[1]=numBus
        for bd in range(limiteInf,numBus):
            DataTableY[bd]=2
            DataTableY10[bd]=1
            numMin[1]=numMin[1]+1
            numMin10[0]=numMin10[0]+1
    
    if(limiteInf==10 or limiteInf==numBus):
        numClases=numClases-1
        numClases10=1
    else:
        if(numBus<=10):
            numClases10=1
        if(limiteInf==0):
            numClases=numClases-1

    DataTY=[]
    DataTY10=[]
    for _ in range(numQuery):
        DataTY = np.concatenate((DataTY, DataTableY))
        DataTY10 = np.concatenate((DataTY10, DataTableY10))


    

    #Lanzamos queries en threads
    try:
        if numQuery == 1:
            thread= threading.Thread(target=principal, daemon=True, args=(DataTable,DataTY,DataTY10,queryL[0],numBus,keywordL[0],workbook,50,0,bTable,dTable,limiteInf,idioma))
            thread.start()
            if(numBus > 0):
                while thread.is_alive():
                    time.sleep(60)
            thread.join()
        else:
            lnumgLen=numQuery
            lnum1=0
            lnum2=5
            while(lnumgLen>5):
                threads = [threading.Thread(target=principal, daemon=True, args=(DataTable,DataTY,DataTY10,queryL[i],numBus,keywordL[i],workbook,10,i,bTable,dTable,limiteInf,idioma,)) for i in range(lnum1,lnum2)]
                for thread in threads:
                    thread.start()
                for thread in threads:
                    while thread.is_alive():
                        time.sleep(60)
                    thread.join()
                lnumgLen = lnumgLen - 5
                lnum1 = lnum2
                lnum2=lnum2+5
                # print("num")
                # print(threading.active_count())
                #gc.collect()
            if(lnumgLen>=1):
                limitT=50/lnumgLen
                lnum2=lnum1+lnumgLen
                threads = [threading.Thread(target=principal, daemon=True, args=(DataTable,DataTY,DataTY10,queryL[i],numBus,keywordL[i],workbook,int(limitT),i,bTable,dTable,limiteInf,idioma,)) for i in range(lnum1,lnum2)]
                for thread in threads:
                    thread.start()
                for thread in threads:
                    while thread.is_alive():
                        time.sleep(60)
                    thread.join()
                #gc.collect()
    except KeyboardInterrupt:
        sys.exit("User Quited")
    # print("num")
    # print(threading.active_count())


    #Datos worksheet excel globales
    worksheetP.set_column('A:A', 19)
    worksheetP.set_column('C:AR', 15)
    worksheetP.set_column('B:B', 50)
    worksheetP.set_column('D:D', 50)
    worksheetP.set_column('G:G', 18)
    worksheetP.set_column('F:F', 50)
    worksheetP.set_column('H:H', 14)
    worksheetP.set_column('N:N', 20)
    worksheetP.set_column('O:T', 26)
    worksheetP.set_column('U:U', 16)
    worksheetP.set_column('V:Z', 14)
    worksheetP.set_column('AA:AF', 17)
    worksheetP.set_column('AG:AR', 38)

    bold2 = workbook.add_format({'bold': True, 'border': True,'align': 'center'})
    bold = workbook.add_format({'bold': True, 'align': 'center'})


    #Tabla medias Totales global
    mediaT= tablaMediasTotales(bTable,numQuery,False,desviaciónT)
    mediasLista=[mediaT]
    if(idioma): mediaCols = [{'header': 'General Summary'}, {'header': 'Url'}, {'header': 'Url Length'}, {'header': 'Title'}, {'header': 'Title Length'}, {'header': 'Description'}, {'header': 'Description Length'}, {'header': 'Creation Date'}, {'header': 'Days Alive'}, {'header': 'Use Https'}, {'header': 'Friendly Url'}, {'header': 'Robots.txt'}, {'header': 'Sitemap.xml'}, {'header': 'Mobile Compatibility'}, {'header': 'Mobile Overall Performance'}, {'header': 'Mobile Load Speed'}, {'header': 'Desktop Oveall Performance'}, {'header': 'Desktop Load Speed'}, {'header': 'Headers'}, {'header': 'Structed Data'}, {'header': 'Schema Markup'}, {'header': 'Total Images'}, {'header': 'Total Videos'}, {'header': 'Total Links'}, {'header': 'Internal Links'}, {'header': 'External Links'}, {'header': 'Internal Follow'}, {'header': 'Internal NoFollow'}, {'header': 'External Follow'}, {'header': 'External NoFollow'}, {'header': 'Errors 404'}, {'header': 'Conection Errors'}, {'header': 'Title Total Occurrences (Keywords%)'}, {'header': 'Title Partial Occurrences (Keywords%)'}, {'header': 'Description Total Occurrences (Keywords%)'}, {'header': 'Description Partial Occurrences (Keywords%)'}, {'header': 'H1 Total Occurrences (Keywords%)'}, {'header': 'H1 Partial Occurrences (Keywords%)'}, {'header': 'Alt Total Occurrences (Keywords%)'}, {'header': 'Alt Partial Occurrences (Keywords%)'}, {'header': 'Src Total Occurrences (Keywords%)'}, {'header': 'Src Partial Occurrences (Keywords%)'}, {'header': 'Body Total Occurrences (Keywords%)'}, {'header': 'Body Partial Occurrences (Keywords%)'}]
    else: mediaCols = [{'header': 'Resumen General'}, {'header': 'Url'}, {'header': 'Url Longitud'}, {'header': 'Título'}, {'header': 'Título Longitud'}, {'header': 'Descripción'}, {'header': 'Descripción Longitud'}, {'header': 'Fecha Creación'}, {'header': 'Dias Vivo'}, {'header': 'Usa Https'}, {'header': 'Url Amigable'}, {'header': 'Robots.txt'}, {'header': 'Sitemap.xml'}, {'header': 'Compatibilidad Móvil'}, {'header': 'Rendimiento Móvil'}, {'header': 'Velocidad Carga Móvil'}, {'header': 'Rendimiento Sobremesa'}, {'header': 'Velocidad Carga Sobremesa'}, {'header': 'Cabeceras'}, {'header': 'Datos Estructurados'}, {'header': 'Schema Markup'}, {'header': 'Imágenes Totales'}, {'header': 'Videos Totales'}, {'header': 'Links Totales'}, {'header': 'Links Internos'}, {'header': 'Links Externos'}, {'header': 'Follow Internos'}, {'header': 'NoFollow Internos'}, {'header': 'Follow Externos'}, {'header': 'NoFollow Externos'}, {'header': 'Errores 404'}, {'header': 'Errores Conexión'}, {'header': 'Título Ocurrencias Totales (Keywords%)'}, {'header': 'Título Ocurrencias Parciales (Keywords%)'}, {'header': 'Descripción Ocurrencias Totales (Keywords%)'}, {'header': 'Descripción Ocurrencias Parciales (Keywords%)'}, {'header': 'H1 Ocurrencias Totales (Keywords%)'}, {'header': 'H1 Ocurrencias Parciales (Keywords%)'}, {'header': 'Alt Ocurrencias Totales (Keywords%)'}, {'header': 'Alt Ocurrencias Parciales (Keywords%)'}, {'header': 'Src Ocurrencias Totales (Keywords%)'}, {'header': 'Src Ocurrencias Parciales (Keywords%)'}, {'header': 'Cuerpo Ocurrencias Totales (Keywords%)'}, {'header': 'Cuerpo Ocurrencias Parciales (Keywords%)'}]
    worksheetP.add_table('A4:AR5', {'data': mediasLista, 'columns':mediaCols})



    #Tabla Media superior global
    mediaS= tablaMediasTotales(dTable,numQuery,True,desviaciónT)
    mediasLista=[mediaS]
    worksheetP.add_table('A8:AR9', {'data': mediasLista, 'columns':mediaCols})

    #Información adicional Excel
    wStr="Querys: " + str(numQuery)
    worksheetP.write(0, 0, wStr,bold)
    if(idioma):
        wStr="Results/Querys: " + str(numBus)
        worksheetP.write(2, 0, "Total Mean Table", bold2)
        worksheetP.write(6, 0, "Superior Mean Table", bold2)
    else:
        wStr="Resultados/Querys: " + str(numBus)
        worksheetP.write(2, 0, "Tabla Medias Totales", bold2)
        worksheetP.write(6, 0, "Tabla Medias Superior", bold2)
    worksheetP.write(0, 1, wStr, bold)
    if(idioma):
        wStr= "Limit Ranking: " + str(limiteInf)
    else:
        wStr= "Limite Ranking: " + str(limiteInf)
    worksheetP.write(6, 1, wStr, bold)


    print("")
    if(idioma):
        print("All querys completed \n")
    else:
        print("Todas las busquedas completadas \n")

    print("")
    print("")



    


    #Creación de clasificadores
    try:
        clf, cvStr = clasificador(numMin,DataTable,DataTY)
        clf10, cvStr10 = clasificador(numMin10,DataTable,DataTY10)
        print("")
    except Exception:
        if(idioma): fileName="Error 429 - Try later"
        else: fileName="Error 429 - Prueba más tarde"
        return



    #Analisis Pagina seleccionada
    if(keyU):
        if(idioma):
            print("Collecting URL/Query to be compared")
        else:
            print("Recopilando URL/Busqueda a comparar")
        keywords=keywords2
        q = q2.replace(' ', '+')
        q = q.replace(':', '%3A')
        q = q.replace('/', '%2F')
        URL = "https://google.com/search?q="+q
        USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.14; rv:65.0) Gecko/20100101 Firefox/65.0"
        headers = {"user-agent" : USER_AGENT, "Cache-Control": "no-cache", "Pragma": "no-cache"}

        aTable = list(map(lambda _: ["Unknown"]*44, range(1)))
        clength=2 + len(keywords)*12
        cTable = list(map(lambda _: ["Unknown"]*clength, range(1)))

        ListGdes=[]
        ListGdes1=[]

        #Sacamos el HTML de google search
        try:
            resp = requests.get(URL, headers=headers, timeout=60)
            if resp.status_code == 200:
                soupg = BeautifulSoup(resp.content, "html.parser")


                #Primer resultado
                for gdes in soupg.find_all("div", {"class":["tF2Cxc","jtfYYd"]}):
                    gdes1=gdes.find_all("a")
                    break
                
                print(URL)
                print(q2)
                gauxiliar=gdes
                urlBool = False
                try:
                    resp=requests.get(q2, headers=headers, timeout=60)
                    if resp.status_code == 200:
                        urlBool = True
                except Exception:
                    urlBool = False
                print(urlBool)
                nplus=0
                
                #Buqueda de URL
                if(urlBool):
                    q3 = q2 + "/"
                    gbool=True
                    try:
                        while(gbool):

                            for gdes in soupg.find_all("div", {"class":["tF2Cxc","jtfYYd"]}):
                                gdes1=gdes.find_all("a")
                                if q2 == gdes1[0].get("href") or q3 == gdes1[0].get("href"):
                                    gbool=False
                                    break

                                if gbool:
                                    nplus=nplus+1
                                    URL2=URL + "&start=" + str(nplus)
                                    resp = requests.get(URL2, headers=headers, timeout=60)
                                    if resp.status_code == 200:
                                        soupg = BeautifulSoup(resp.content, "html.parser")
                                    elif nplus>30:
                                        gdes=gauxiliar
                                        gbool=False
                                    else:
                                        gdes=gauxiliar
                                        gbool=False
                                    break

                        gdes1=q2

                    except Exception:
                        gdes=gauxiliar
                        gdes1=q2
                            
                else:
                    gdes1=gdes1[0].get("href")


                ListGdes.append(gdes)
                ListGdes1.append(gdes1)
                print(gdes1)
                print(nplus) 

                if(idioma):
                    print("Analising URL/Query")
                else:
                    print("Analizando URL/Busqueda")

            
                #blockPrint()
                crearTabla(0,aTable,cTable,ListGdes,ListGdes1,headers,keywords,0)
                #enablePrint()
                aTable[0][0]="Resultados"
                cTable[0][0]="Resultados"
                a2Table, nada =tablaMedias(1,aTable)
                tabPos=[1,3,5,7]
                for i in tabPos:
                    a2Table[i]=aTable[0][i]

                if(idioma): mediaCols = [{'header': 'Web Analisis'}, {'header': 'Url'}, {'header': 'Url Length'}, {'header': 'Title'}, {'header': 'Title Length'}, {'header': 'Description'}, {'header': 'Description Length'}, {'header': 'Creation Date'}, {'header': 'Days Alive'}, {'header': 'Protocol'}, {'header': 'Friendly Url'}, {'header': 'Robots.txt'}, {'header': 'Sitemap.xml'}, {'header': 'Mobile Compatibility'}, {'header': 'Mobile Overall Performance'}, {'header': 'Mobile Load Speed'}, {'header': 'Desktop Oveall Performance'}, {'header': 'Desktop Load Speed'}, {'header': 'Headers'}, {'header': 'Structed Data'}, {'header': 'Schema Markup'}, {'header': 'Total Images'}, {'header': 'Total Videos'}, {'header': 'Total Links'}, {'header': 'Internal Links'}, {'header': 'External Links'}, {'header': 'Internal Follow'}, {'header': 'Internal NoFollow'}, {'header': 'External Follow'}, {'header': 'External NoFollow'}, {'header': 'Errors 404'}, {'header': 'Conection Errors'}, {'header': 'Title Total Occurrences (Keywords%)'}, {'header': 'Title Partial Occurrences (Keywords%)'}, {'header': 'Description Total Occurrences (Keywords%)'}, {'header': 'Description Partial Occurrences (Keywords%)'}, {'header': 'H1 Total Occurrences (Keywords%)'}, {'header': 'H1 Partial Occurrences (Keywords%)'}, {'header': 'Alt Total Occurrences (Keywords%)'}, {'header': 'Alt Partial Occurrences (Keywords%)'}, {'header': 'Src Total Occurrences (Keywords%)'}, {'header': 'Src Partial Occurrences (Keywords%)'}, {'header': 'Body Total Occurrences (Keywords%)'}, {'header': 'Body Partial Occurrences (Keywords%)'}]
                else: mediaCols = [{'header': 'Web Análisis'}, {'header': 'Url'}, {'header': 'Url Longitud'}, {'header': 'Título'}, {'header': 'Título Longitud'}, {'header': 'Descripción'}, {'header': 'Descripción Longitud'}, {'header': 'Fecha Creación'}, {'header': 'Dias Vivo'}, {'header': 'Protocolo'}, {'header': 'Url Amigable'}, {'header': 'Robots.txt'}, {'header': 'Sitemap.xml'}, {'header': 'Compatibilidad Móvil'}, {'header': 'Rendimiento Móvil'}, {'header': 'Velocidad Carga Móvil'}, {'header': 'Rendimiento Sobremesa'}, {'header': 'Velocidad Carga Sobremesa'}, {'header': 'Cabeceras'}, {'header': 'Datos Estructurados'}, {'header': 'Schema Markup'}, {'header': 'Imágenes Totales'}, {'header': 'Videos Totales'}, {'header': 'Links Totales'}, {'header': 'Links Internos'}, {'header': 'Links Externos'}, {'header': 'Follow Internos'}, {'header': 'NoFollow Internos'}, {'header': 'Follow Externos'}, {'header': 'NoFollow Externos'}, {'header': 'Errores 404'}, {'header': 'Errores Conexión'}, {'header': 'Título Ocurrencias Totales (Keywords%)'}, {'header': 'Título Ocurrencias Parciales (Keywords%)'}, {'header': 'Descripción Ocurrencias Totales (Keywords%)'}, {'header': 'Descripción Ocurrencias Parciales (Keywords%)'}, {'header': 'H1 Ocurrencias Totales (Keywords%)'}, {'header': 'H1 Ocurrencias Parciales (Keywords%)'}, {'header': 'Alt Ocurrencias Totales (Keywords%)'}, {'header': 'Alt Ocurrencias Parciales (Keywords%)'}, {'header': 'Src Ocurrencias Totales (Keywords%)'}, {'header': 'Src Ocurrencias Parciales (Keywords%)'}, {'header': 'Cuerpo Ocurrencias Totales (Keywords%)'}, {'header': 'Cuerpo Ocurrencias Parciales (Keywords%)'}]
                worksheetP.add_table('A13:AR14', {'data': [a2Table], 'columns':mediaCols})

                #Creamos Tabla de Keywords para esta Busqueda
                keyCols = [{'header': 'Web Analisis'}, {'header': 'Url'}]
                keyStr=""
                if(idioma): 
                    ocurS = [" Total Occurrences", " Partial Occurrences"]
                    ocurS2 = ["Title ","Description ","Body ","H1 ","Alt ","Src "]
                else: 
                    ocurS = [" Ocurrencias Totales", " Ocurencias Parciales"]
                    ocurS2 = ["Título ","Descripción ","Cuerpo ","H1 ","Alt ","Src "]
                for tipo in ocurS2:
                    for key in keywords:
                        for ocur in ocurS:
                            keyStr= tipo + key + ocur
                            keyStr=[{'header': keyStr}]
                            keyCols = keyCols+keyStr
                letter = ''
                while clength > 25 + 1:
                    letter += chr(65 + int((clength-1)/26) - 1)
                    clength = clength - (int((clength-1)/26))*26
                letter += chr(65 - 1 + (int(clength)))
                Celdas = 'A17:'+ letter + '18'
                worksheetP.add_table(Celdas, {'data': cTable,'columns':keyCols})

                #Información adicional Excel
                if(idioma):
                    worksheetP.write(11, 1, 'URL to analise'), bold
                    worksheetP.write(11, 0, 'Base Table', bold2)
                    worksheetP.write(15, 0, 'Keywords Table', bold2)
                else:
                    worksheetP.write(11, 1, 'URL a analizar', bold)
                    worksheetP.write(11, 0, 'Tabla Base', bold2)
                    worksheetP.write(15, 0, 'Tabla Keywords', bold2)
                worksheetP.write(15, 2, "Keywords: ", bold)
                worksheetP.write(15, 3, str(keywords), bold)

                informe(cvStr,clf,cvStr10,clf10,numClases,numClases10,clfList,workbook,worksheetI,mediaS,mediaT,aTable,a2Table,desviaciónT,idioma)


            else:
                #sys.exit("Report URL Google Conection Error")
                if(idioma):
                    worksheetI.write(0, 0, 'Error at extracting URL')
                else:
                    worksheetI.write(11, 1, 'Error al extraer la URL')
        
        except Exception:
            print("Report URL Google Conection Error")
            if(idioma):
                worksheetI.write(0, 0, 'Error at extracting URL')
            else:
                worksheetI.write(11, 1, 'Error al extraer la URL')


    else: 
        aTable = list(map(lambda _: ["Unknown"]*44, range(1)))
        a2Table, none = tablaMedias(1,aTable)
        informe(cvStr,clf,cvStr10,clf10,numClases,numClases10,clfList,workbook,worksheetI,mediaS,mediaT,aTable,a2Table,desviaciónT,idioma)

    workbook.close()


    if(idioma):
        print("Completed: %s generated \n\n" %(fileName))
        print("Execution time: %f s" %(round(time.time() - t, 4)))
    else:
        print("Completado: %s generado \n\n" %(fileName))
        print("Tiempo de ejecución: %f s" %(round(time.time() - t, 4)))



    return










###################################################################################################################################################

#############################################
###                                       ###
###   Parte 3.2 Crador de clasificadores  ###
###                                       ###
#############################################



def clasificador(numM,DT,DY):

    """
    Num ejecuciones: 2

    Función para crear el clasificador gradientBoosting que estimará los pesados de la variables y la URL del informe.

    param:  "numM": Array que musetra el numero de etiquetas de cada clase. Clase i de DT y DY = a clase i-1.
            "DT": Tabla de datos.
            "DY": Tabla de etiquetas.

             return: "clf": Clasificador obtenido.
                     "cvStr2": Precisión del clasificador (Str).
            
    """

    #Clasificador
    auxCero=[0]*39
    auxCero=[auxCero]
    numMinF=numM[0]
    if(numMinF==0):numMinF=numM[1]
    for i in numM:
        if(i>0 and i<numMinF): numMinF=i
    #clf = GradientBoostingClassifier(n_estimators=100, learning_rate=1.0, max_depth=1, random_state=0).fit(DataTable,DataTY)
    if(numMinF>=5):
        cv = RepeatedStratifiedKFold(n_splits=5, n_repeats=3, random_state=1)
        for _ in range(5):
            DT = np.concatenate((DT, auxCero))
            DY = np.concatenate((DY, [0]))
    else:
        cv = RepeatedStratifiedKFold(n_splits=numMinF, n_repeats=3, random_state=1)
        for _ in range(numMinF):
            DT = np.concatenate((DT, auxCero))
            DY = np.concatenate((DY, [0]))
    
    
    learnR=0
    n_est=0

    if(numMinF==1): 
        if(idioma): cvStr2 = "Classifier Accuracy: Unavilable"
        else: cvStr2 = "Precisión del Clasificado: No Disponible"
        clf2 = GradientBoostingClassifier()
        
    else:
        crossVS = np.array([[0], [0]])
        crossi=[0.1,0.25,0.5,0.75,1]
        crossj=[100,250,500,750,1000]

        for i in crossi:

            for j in crossj:

                clf3 = GradientBoostingClassifier(learning_rate=i,n_estimators=j)
                crossVS2 = cross_val_score(clf3, DT, DY, scoring='accuracy',cv=cv)
                if(np.mean(crossVS2)>np.mean(crossVS)):
                    crossVS=crossVS2
                    clf2=clf3
                    learnR=i
                    n_est=j



        if(idioma):
            cvStr2= "Classifier Accuracy: " + str(round(np.mean(crossVS),3)) + " (Deviation: " + str(round(np.std(crossVS),3)) + ")"
        else:
            cvStr2= "Precisión del clasificador: " + str(round(np.mean(crossVS),3)) + " (Desviación: " + str(round(np.std(crossVS),3)) + ")"

    print(cvStr2)
    print(learnR)
    print(n_est)

    return clf2.fit(DT,DY), cvStr2












###################################################################################################################################################

#############################################
###                                       ###
###      Parte 3.3 Crador Informe URL     ###
###                                       ###
#############################################



def informe(cvStr,clf,cvStr10,clf10,numClases,numClases10,clfList,workbook,worksheetI,mediaS,mediaT,aTable,a2Table,desviaciónT,idioma):

    """
    Num ejecuciones: 1

    Función para crear el apartado informe del xmlx para la URL/Query principal a analizar.

    param:  "workbook": Woorkbook xlsx en el que mostrar los datos.
            "worksheetI": Woorksheet para el Informe dentro del Workbook.
            "mediaS": Medias Total de todas las queries limitado por el Ranking inferior límite (Mejores resultados).
            "mediaT": Media Total de todas las queries.
            "aTable": Variables SEO de la URL/Query principal (acondicionado para ilustrarlo en xmlx).
            "a2Table": Variables SEO de la URL/Query principal (acondicionado para calculos).
            "idioma": Idioma seleccionado (Ingles/Español). 
            "desviaciónT": Array de desviación típica.
            
    """

    if(idioma): print("Generating Report")
    else: print("Generando Informe")

    red = workbook.add_format({'bg_color': 'red', 'border': True}) #Top Priority 5
    orange = workbook.add_format({'bg_color': 'orange', 'border': True}) #Highly Recommended 4
    yellow = workbook.add_format({'bg_color': 'yellow', 'border': True}) #Recommended 3
    lime = workbook.add_format({'bg_color': '#d0ff00', 'border': True}) #Low priority 2
    green = workbook.add_format({'bg_color': 'green', 'border': True}) #Adequate 1
    grey = workbook.add_format({'bg_color': '#cccccc', 'border': True}) #Indifferent 0
    align = workbook.add_format({'align': 'left'})
    blue = workbook.add_format({'bold': True, 'bg_color': '#0099f7', 'border': True})
    bold = workbook.add_format({'bold': True, 'align': 'center'})

    red2 = workbook.add_format({'bold': True,'bg_color': 'red', 'border': True,'align': 'center'}) #Top Priority 5
    orange2 = workbook.add_format({'bold': True,'bg_color': 'orange', 'border': True,'align': 'center'}) #Highly Recommended 4
    yellow2 = workbook.add_format({'bold': True,'bg_color': 'yellow', 'border': True,'align': 'center'}) #Recommended 3
    lime2 = workbook.add_format({'bold': True,'bg_color': '#d0ff00', 'border': True,'align': 'center'}) #Low priority 2
    green2 = workbook.add_format({'bold': True,'bg_color': 'green', 'border': True,'align': 'center'}) #Adequate 1
    grey2 = workbook.add_format({'bold': True,'bg_color': '#cccccc', 'border': True,'align': 'center'}) #Indifferent 0
    blue2 = workbook.add_format({'bold': True, 'bg_color': '#0099f7', 'border': True,'align': 'center'})
    bold2 = workbook.add_format({'bold': True, 'border': True,'align': 'center'})
    cen = workbook.add_format({'align': 'center'})

    worksheetI.write_blank (1, 0, '', red)
    worksheetI.write_blank (2, 0, '', orange)
    worksheetI.write_blank (3, 0, '', yellow)
    worksheetI.write_blank (4, 0, '', lime)
    worksheetI.write_blank (5, 0, '', green)
    worksheetI.write_blank (6, 0, '', grey)

    
    if(idioma):
        
        worksheetI.write(0,0, "Legend",bold2)
        worksheetI.write(0,3, "Error Count",bold2)
        strC="General (" + str(numClases) + " classes)"
        worksheetI.write(0,5, strC,bold2)
        strC="Top 10 (" + str(numClases10) + " classes)"
        worksheetI.write(5,5, strC,bold2)
        worksheetI.write(1,5, "Score",blue2)
        worksheetI.write(1,6, "Estimated Ranking",blue2)
        worksheetI.write(6,5, "Score",blue2)
        worksheetI.write(6,6, "In Top Ten",blue2)
        worksheetI.write(1,1, "Max")
        worksheetI.write(2,1, "High")
        worksheetI.write(3,1, "Moderate")
        worksheetI.write(4,1, "Low")
        worksheetI.write(5,1, "Adequate")
        worksheetI.write(6,1, "Indifferent")
        worksheetI.write(9, 0, 'Report',bold2)
        worksheetI.write(15, 0, 'Error level',bold2)
        worksheetI.write(15, 1, 'Atribute',bold2)
        worksheetI.write(15, 2, 'Your Values',bold2)
        worksheetI.write(15, 3, 'Recommended Values',bold2)
        worksheetI.write(15, 4, 'Max Error',bold2)
        worksheetI.write(15, 5, 'Weight (General)',bold2)
        worksheetI.write(15, 6, 'Weight (Top 10)',bold2)
    else:
        worksheetI.write(0,0,"Leyenda",bold2)
        worksheetI.write(0,3, "Conteo Errores",bold2)
        strC="General (" + str(numClases) + " clases)"
        worksheetI.write(0,5, strC,bold2)
        strC="Top 10 (" + str(numClases10) + " clases)"
        worksheetI.write(5,5, strC,bold2)
        worksheetI.write(1,5, "Puntuación",blue2)
        worksheetI.write(1,6, "Ranking Estimado",blue2)
        worksheetI.write(6,5, "Puntuación",blue2)
        worksheetI.write(6,6, "En Top 10",blue2)
        worksheetI.write(1,1, "Máximo")
        worksheetI.write(2,1, "Alto")
        worksheetI.write(3,1, "Moderado")
        worksheetI.write(4,1, "Bajo")
        worksheetI.write(5,1, "Adecuado")
        worksheetI.write(6,1, "Indiferente")
        worksheetI.write(9, 0, 'Informe',bold2)
        worksheetI.write(15, 0, 'Nivel de Error',bold2)
        worksheetI.write(15, 1, 'Atributo',bold2)
        worksheetI.write(15, 2, 'Valores Actuales',bold2)
        worksheetI.write(15, 3, 'Valores Recomendados',bold2)
        worksheetI.write(15, 4, 'Error Máximo',bold2)
        worksheetI.write(15, 5, 'Peso (General)',bold2)
        worksheetI.write(15, 6, 'Peso (Top 10)',bold2)
    


    prioridad = [0]*44
    prioridadMax = [0]*44
    rVal = [0]*44

    #0-1
    ceroUno=[9,10,11,12,13,19,20]
    for cu in ceroUno:
        rVal[cu]="Yes"
        if(mediaS[cu]>0.9 or mediaS[cu]-mediaT[cu]>0.3):
            prioridad[cu]=5
        elif(mediaS[cu]>0.75 or mediaS[cu]-mediaT[cu]>0.2):
            prioridad[cu]=4
        elif(mediaS[cu]>0.6 or mediaS[cu]-mediaT[cu]>0.1):
            prioridad[cu]=3
        elif(mediaS[cu]>0.5):
            prioridad[cu]=2
        else:
            prioridad[cu]=0
            if(idioma):
                rVal[cu]="Indifferent"
            else:
                rVal[cu]="Indiferente"

        prioridadMax[cu]=prioridad[cu]
        if a2Table[cu] == 1 and (not (prioridad[cu] == 0)):
            prioridad[cu]=1

    
    #int
    numer=[2,4,6,8,14,15,16,17,18,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43]
    for cu in numer:

        roundaux=0
        if cu in [15,17]:
            roundaux=1
        elif cu in [32,33,34,35,36,37,38,39,40,41,42,43]:
            roundaux=2
        difMed=round(mediaS[cu],roundaux)
        dif=round(abs(a2Table[cu]-mediaS[cu]),roundaux)
        

        if(dif<= difMed*0.25):
            prioridad[cu]=1
        elif(dif<= difMed*0.45):
            prioridad[cu]=2
        elif(dif<= difMed*0.65):
            prioridad[cu]=3
        elif(dif<= difMed*0.85):
            prioridad[cu]=4
        else:
            prioridad[cu]=5

        #if cu not in [14,15,16,17]:
        dt=round(desviaciónT[cu],roundaux)
        if(dt<= difMed*0.45):
            dtp=5
        elif(dt<= difMed*0.65):
            dtp=4
        elif(dt<= difMed*0.85):
            dtp=3
        elif(dt<= difMed*0.95):
            dtp=2
        else:
            if cu in [26,27,30,31,32,33,34,35,36,37,38,39,40,41,42,43]:
                dtp=2
            else:
                dtp=0

        prioridadMax[cu]=dtp

        if(dtp<prioridad[cu]): prioridad[cu] = dtp

        


        raux = mediaS[cu]-difMed*0.25
        raux2 = mediaS[cu]+difMed*0.25


        if difMed == 1 or difMed == 0:
            if dif <= 1 and not (prioridad[cu]==0):
                prioridad[cu]=1
            raux =0
            raux2=1

        if cu in [14,16]:
            if a2Table[cu] > mediaS[cu]-difMed*0.05:
                prioridad[cu]=1
            elif(a2Table[cu] > mediaS[cu]-difMed*0.10):
                prioridad[cu]=2
            elif(a2Table[cu] > mediaS[cu]-difMed*0.20):
                prioridad[cu]=3
            elif(a2Table[cu] > mediaS[cu]-difMed*0.30):
                prioridad[cu]=4
            else:
                prioridad[cu]=5
            raux=round(mediaS[cu]-difMed*0.05,0)
            raux=int(raux)
            if(raux <= a2Table[cu] <= 100): prioridad[cu]=1
            rVal[cu]= "[" + str(raux) + "-100]"
        elif cu in [15,17]:
            if a2Table[cu] < mediaS[cu]:
                prioridad[cu]=1
            raux2=round(raux2,1)
            if(0 <= a2Table[cu] <= raux2): prioridad[cu]=1
            rVal[cu]= "[0-" + str(raux2) + "]"
        elif cu in [30,31]:
            if a2Table[cu] < mediaS[cu]:
                prioridad[cu]=1
            raux=round(raux,0)
            raux2=round(raux2,0)
            raux=int(raux)
            raux2=int(raux2)
            if(0 <= a2Table[cu] <= raux2): prioridad[cu]=1
            rVal[cu]= "[0-" + str(raux2) + "]"
        elif cu in [32,33,34,35,36,37,38,39,40,41,42,43]:
            raux=round(raux,2)
            raux2=round(raux2,2)
            if(raux <= a2Table[cu] <= raux2 and not (prioridad[cu]==0)): prioridad[cu]=1
            rVal[cu]= "[" + str(raux) + "%-" + str(raux2) + "%]"
        else:
            raux=round(raux,0)
            raux2=round(raux2,0)
            raux=int(raux)
            raux2=int(raux2)
            if(raux <= a2Table[cu] <= raux2 and not (prioridad[cu]==0)): prioridad[cu]=1
            rVal[cu]= "[" + str(raux) + "-" + str(raux2) + "]"

    for cu in numer:
        if "Unknown"==str(aTable[0][cu]) :
            prioridad[cu]=prioridadMax[cu]

    
    #if(a2Table[8]==0):
    #    prioridad[8]=0

    #ignorar variables compuestas
    # for z in [23,24,25]:
    #     prioridad[z]=0


    # No Ajustables
    if(idioma):
        columns = ['Web Analisis', 'Url', 'Url Length', 'Title', 'Title Length', 'Description', 'Description Length', 'Creation Date', 'Days Alive', 'Use Https', 'Friendly Url', 'Robots.txt', 'Sitemap.xml', 'Mobile Compatibility', 'Mobile Overall Performance', 'Mobile Load Speed', 'Desktop Oveall Performance', 'Desktop Load Speed', 'Headers', 'Structed Data', 'Schema Markup', 'Total Images', 'Total Videos', 'Total Links', 'Internal Links', 'External Links', 'Internal Follow', 'External Follow', 'Internal NoFollow', 'External NoFollow', 'Errors 404', 'Conection Errors', 'Title Total Occurrences (Keywords%)', 'Title Partial Occurrences (Keywords%)', 'Description Total Occurrences (Keywords%)', 'Description Partial Occurrences (Keywords%)', 'H1 Total Occurrences (Keywords%)', 'H1 Partial Occurrences (Keywords%)', 'Alt Total Occurrences (Keywords%)', 'Alt Partial Occurrences (Keywords%)', 'Src Total Occurrences (Keywords%)', 'Src Partial Occurrences (Keywords%)', 'Body Total Occurrences (Keywords%)', 'Body Partial Occurrences (Keywords%)']
        columns2 = ['Web Analisis', 'Url', 'Url Length', 'Title', 'Title Length', 'Description', 'Description Length', 'Creation Date', 'Days Alive', 'Protocol', 'Friendly Url', 'Robots.txt', 'Sitemap.xml', 'Mobile Compatibility', 'Mobile Overall Performance', 'Mobile Load Speed', 'Desktop Oveall Performance', 'Desktop Load Speed', 'Headers', 'Structed Data', 'Schema Markup', 'Total Images', 'Total Videos', 'Total Links', 'Internal Links', 'External Links', 'Internal Follow', 'External Follow', 'Internal NoFollow', 'External NoFollow', 'Errors 404', 'Conection Errors', 'Title Total Occurrences (Keywords%)', 'Title Partial Occurrences (Keywords%)', 'Description Total Occurrences (Keywords%)', 'Description Partial Occurrences (Keywords%)', 'H1 Total Occurrences (Keywords%)', 'H1 Partial Occurrences (Keywords%)', 'Alt Total Occurrences (Keywords%)', 'Alt Partial Occurrences (Keywords%)', 'Src Total Occurrences (Keywords%)', 'Src Partial Occurrences (Keywords%)', 'Body Total Occurrences (Keywords%)', 'Body Partial Occurrences (Keywords%)']
    else:
        columns = ['Web Analisis', 'Url', 'Url Longitud', 'Título', 'Título Longitud', 'Descripción', 'Descripción Longitud', 'Fecha Creación', 'Dias Vivo', 'Uso Https', 'Url Amigable', 'Robots.txt', 'Sitemap.xml', 'Compatibilidad Móvil', 'Rendimiento Móvil', 'Velocidad Carga Móvil', 'Rendimiento Sobremesa', 'Velocidad Carga Sobremesa', 'Cabeceras', 'Datos Estructurados', 'Schema Markup', 'Imágenes Totales', 'Videos Totales', 'Links Totales', 'Links Internos', 'Links Externos', 'Follow Internos', 'Follow Externos', 'NoFollow Internos', 'NoFollow Externos', 'Errores 404', 'Errores Conexión', 'Título Ocurrencias Totales (Keywords%)', 'Título Ocurrencias Parciales (Keywords%)', 'Descripción Ocurrencias Totales (Keywords%)', 'Descripción Ocurrencias Parciales (Keywords%)', 'H1 Ocurrencias Totales (Keywords%)', 'H1 Ocurrencias Parciales (Keywords%)', 'Alt Ocurrencias Totales (Keywords%)', 'Alt Ocurrencias Parciales (Keywords%)', 'Src Ocurrencias Totales (Keywords%)', 'Src Ocurrencias Parciales (Keywords%)', 'Cuerpo Ocurrencias Totales (Keywords%)', 'Cuerpo Ocurrencias Parciales (Keywords%)']
        columns2 = ['Web Analisis', 'Url', 'Url Longitud', 'Título', 'Título Longitud', 'Descripción', 'Descripción Longitud', 'Fecha Creación', 'Dias Vivo', 'Protocolo', 'Url Amigable', 'Robots.txt', 'Sitemap.xml', 'Compatibilidad Móvil', 'Rendimiento Móvil', 'Velocidad Carga Móvil', 'Rendimiento Sobremesa', 'Velocidad Carga Sobremesa', 'Cabeceras', 'Datos Estructurados', 'Schema Markup', 'Imágenes Totales', 'Videos Totales', 'Links Totales', 'Links Internos', 'Links Externos', 'Follow Internos', 'Follow Externos', 'NoFollow Internos', 'NoFollow Externos', 'Errores 404', 'Errores Conexión', 'Título Ocurrencias Totales (Keywords%)', 'Título Ocurrencias Parciales (Keywords%)', 'Descripción Ocurrencias Totales (Keywords%)', 'Descripción Ocurrencias Parciales (Keywords%)', 'H1 Ocurrencias Totales (Keywords%)', 'H1 Ocurrencias Parciales (Keywords%)', 'Alt Ocurrencias Totales (Keywords%)', 'Alt Ocurrencias Parciales (Keywords%)', 'Src Ocurrencias Totales (Keywords%)', 'Src Ocurrencias Parciales (Keywords%)', 'Cuerpo Ocurrencias Totales (Keywords%)', 'Cuerpo Ocurrencias Parciales (Keywords%)']
    
    NoAjustables=[1,3,5,7]

    #Ajustables y Noajustables
    cuAux=1
    tAux=1
    i=0
    numPrior=[0,0,0,0,0,0]
    importancias=[0,0,0,0,0,0]
    importancias10=[0,0,0,0,0,0]
    ajustables=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43]
    clfImportances=clf.feature_importances_
    clfImportances10=clf10.feature_importances_
    errorZ=0
    errorZ2=0
    for cu in ajustables:
        cuPos=16+cuAux

        if cu not in NoAjustables:
            cuAux=cuAux+3
            aVal=""

            if(prioridad[cu]==5):
                worksheetI.write_blank (cuPos, 0, '', red)
                worksheetI.write(11, tAux, aTable[0][cu],red)
                numPrior[5]=numPrior[5]+1
                importancias[5]=importancias[5]+clfImportances[i]
                importancias10[5]=importancias10[5]+clfImportances10[i]
            elif(prioridad[cu]==4):
                worksheetI.write_blank (cuPos, 0, '', orange)
                worksheetI.write(11, tAux, aTable[0][cu],orange)
                numPrior[4]=numPrior[4]+1
                importancias[4]=importancias[4]+clfImportances[i]
                importancias10[4]=importancias10[4]+clfImportances10[i]
            elif(prioridad[cu]==3):
                worksheetI.write_blank (cuPos, 0, '', yellow)
                worksheetI.write(11, tAux, aTable[0][cu],yellow)
                numPrior[3]=numPrior[3]+1
                importancias[3]=importancias[3]+clfImportances[i]
                importancias10[3]=importancias10[3]+clfImportances10[i]
            elif(prioridad[cu]==2):
                worksheetI.write_blank (cuPos, 0, '', lime)
                worksheetI.write(11, tAux, aTable[0][cu],lime)
                numPrior[2]=numPrior[2]+1
                importancias[2]=importancias[2]+clfImportances[i]
                importancias10[2]=importancias10[2]+clfImportances10[i]
            elif(prioridad[cu]==1):
                worksheetI.write_blank (cuPos, 0, '', green)
                worksheetI.write(11, tAux, aTable[0][cu],green)
                numPrior[1]=numPrior[1]+1
                importancias[1]=importancias[1]+clfImportances[i]
                importancias10[1]=importancias10[1]+clfImportances10[i]
            else:
                worksheetI.write_blank (cuPos, 0, '', grey)
                worksheetI.write(11, tAux, aTable[0][cu],grey)
                numPrior[0]=numPrior[0]+1


            if "Unknown"==str(aTable[0][cu]) :
                aVal = "Unknown"
            
            else:
                if cu in ceroUno:
                    if a2Table[cu] == 0:
                        aVal="No"
                    else:
                        aVal="Yes"
                elif cu in [32,33,34,35,36,37,38,39,40,41,42,43]:
                    aVal = str(round(a2Table[cu],3)) + "%"
                else:
                    aVal = str(a2Table[cu])

            

            worksheetI.write(cuPos, 1, columns[cu])
            worksheetI.write(cuPos, 2, aVal,align)
            worksheetI.write(cuPos, 3, rVal[cu],align)
            worksheetI.write(cuPos, 4, prioridadMax[cu],cen)
            worksheetI.write(cuPos, 5, clfImportances[i],bold)
            worksheetI.write(cuPos, 6, clfImportances10[i],bold)
            errorZ= clfImportances[i]+errorZ
            errorZ2= clfImportances10[i]+errorZ2
            i=i+1

        else:
            worksheetI.write(11, tAux, aTable[0][cu],grey)

        worksheetI.write(10, tAux, columns2[cu], blue)
        tAux=tAux+1

    print(errorZ)
    print(errorZ2)
    worksheetI.write(10, 0, columns[0], blue)
    worksheetI.write(11, 0, aTable[0][0], grey)

    
    # numPriorCount=numPrior[1]+numPrior[2]+numPrior[3]+numPrior[4]+numPrior[5]
    # if(numPriorCount==0):numPriorCount=1
    # numScore=100-(numPrior[5]*100+numPrior[4]*75+numPrior[3]*50+numPrior[2]*25)/numPriorCount
    numPriorCount=importancias[1]+importancias[2]+importancias[3]+importancias[4]+importancias[5]
    if(numPriorCount==0):numPriorCount=1
    numScore=100*((importancias[1]+importancias[2]*0.75+importancias[3]*0.5+importancias[4]*0.25)/numPriorCount)

    numScore=int(numScore)
    strNumScor= str(numScore) + "/" "100"

    if(numScore>=90):
        worksheetI.write(2, 5, strNumScor, green2)
    elif(numScore>=80):
        worksheetI.write(2, 5, strNumScor, lime2)
    elif(numScore>=65):
        worksheetI.write(2, 5, strNumScor, yellow2)
    elif(numScore>=50):
        worksheetI.write(2, 5, strNumScor, orange2)
    else:
        worksheetI.write(2, 5, strNumScor, red2)

    

    numPriorCount=importancias10[1]+importancias10[2]+importancias10[3]+importancias10[4]+importancias10[5]
    if(numPriorCount==0):numPriorCount=1
    numScore=100*((importancias10[1]+importancias10[2]*0.75+importancias10[3]*0.5+importancias10[4]*0.25)/numPriorCount)

    numScore=int(numScore)
    strNumScor= str(numScore) + "/" "100"

    if(numScore>=90):
        worksheetI.write(7, 5, strNumScor, green2)
    elif(numScore>=80):
        worksheetI.write(7, 5, strNumScor, lime2)
    elif(numScore>=65):
        worksheetI.write(7, 5, strNumScor, yellow2)
    elif(numScore>=50):
        worksheetI.write(7, 5, strNumScor, orange2)
    else:
        worksheetI.write(7, 5, strNumScor, red2)


    worksheetI.write(1, 3, str(numPrior[5]), red2)
    worksheetI.write(2, 3, str(numPrior[4]), orange2)
    worksheetI.write(3, 3, str(numPrior[3]), yellow2)
    worksheetI.write(4, 3, str(numPrior[2]), lime2)
    worksheetI.write(5, 3, str(numPrior[1]), green2)
    worksheetI.write(6, 3, str(numPrior[0]), grey2)






    #Predicción clasificador
    a3Table = [0]*39
    tabPos=[2,4,6,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43]
    for y in range(0,39):
        x = tabPos[y]
        a3Table[y]=a2Table[x]

    clfPre=clf.predict([a3Table])[0]
    if clfPre == 1:
        clfPreS="Top " + str(clfList[0])
        worksheetI.write(2,6, clfPreS,green2)
    elif clfPre == 2:
        clfPreS="Top " + str(clfList[1])
        worksheetI.write(2,6, clfPreS,lime2)
    elif clfPre == 3:
        clfPreS="Pos: " + str(clfList[1]) + "-" + str(clfList[2])
        worksheetI.write(2,6, clfPreS,yellow2)
    elif clfPre == 4:
        clfPreS="Pos: " + str(clfList[2]) + "-" + str(clfList[3])
        worksheetI.write(2,6, clfPreS,orange2)
    elif clfPre == 5:
        clfPreS="Pos " + str(clfList[3]) + "+"
        worksheetI.write(2,6, clfPreS,red2)
    else:
        worksheetI.write(2,6, "Unknown",grey2)

    worksheetI.write(3,6, cvStr)


    clfPre=clf10.predict([a3Table])[0]
    if clfPre == 1:
        if(idioma): worksheetI.write(7,6, "Yes",green2)
        else: worksheetI.write(7,6, "Si",green2)
    elif clfPre == 2:
        worksheetI.write(7,6, "No",red2)
    else:
        worksheetI.write(7,6, "Unknown",grey2)

    worksheetI.write(8,6, cvStr10)

    
    


    #Datos worksheet excel globales
    worksheetI.set_column('A:AR', 18)
    worksheetI.set_column('E:AR', 13)
    worksheetI.set_column('B:B', 45)
    worksheetI.set_column('C:D', 21)
    worksheetI.set_column('F:G', 41)
    worksheetI.set_column('H:H', 14)
    worksheetI.set_column('N:N', 20)
    worksheetI.set_column('O:T', 26)
    worksheetI.set_column('U:U', 16)
    worksheetI.set_column('V:Z', 14)
    worksheetI.set_column('AA:AF', 17)
    worksheetI.set_column('AG:AR', 38)

    return



###################################################################################################################################################



































############################################################################
###                                                                      ###
###                             PARTE 4: GUI                             ###
###                                                                      ###
############################################################################





###################################################################################################################################################

#############################################
###                                       ###
###     Parte 4.1 Ventanas Emergentes     ###
###                                       ###
#############################################



def launch():

    """

    Ventana en la que se ha de especificar los detalles de la busqueda a realizar:
        -Numero de Queries
        -Resultados por Query
        -Keywords=Query (yes/no)
        -Ranking Inferior Limite para calcular la media de los mejores resultados
        -Comparar resultados con una URL/Query

        ###Se muestra ventana de error en caso de input inadecuado

    """

    #Terminación del programa en caso de cerrar ventana
    def on_closing():
        second.destroy()
        root.quit()

    #Crea la ventana
    second = tk.Toplevel(root)
    second.iconbitmap("4693059.ico")
    second.title("SEO Variables Analyzer")
    second.geometry("300x530")
    second.resizable(False, False)
    root.withdraw()
    second.protocol("WM_DELETE_WINDOW", on_closing)
    

    #Widgets de la ventana 
    labelQ = ttk.Label(second,text="QUERY", justify='center',font='Helvetica 12 bold')
    labelQ.pack(pady=(10,0))

    if(idioma): labelQuery = ttk.Label(second,text="Insert number of Queries: \n (Recommended 2+)")
    else: labelQuery = ttk.Label(second,text="Ingrese el numero de Queries: \n (Recomendado 2+)")
    labelQuery.pack(fill=tk.X, padx=5, pady=5)
    entryQuery = tk.Entry(second,width=45)
    entryQuery.pack()

    if(idioma): labelRes = ttk.Label(second,text="Insert number of results per Query: \n (Recommended 20-200)")
    else: labelRes = ttk.Label(second,text="Ingrese el numero de resultados por Query: \n (Recomendado 20-200)")
    labelRes.pack(fill=tk.X, padx=5, pady=(20,5))
    entryRes = tk.Entry(second,width=45)
    entryRes.pack()

    if(idioma): labelkey = ttk.Label(second,text="Querys = Keywords [Y/n]:")
    else: labelkey = ttk.Label(second,text="Querys = Keywords [S/n]:")
    labelkey.pack(fill=tk.X, padx=5, pady=(20,5))
    selected_key = tk.StringVar(master=second)
    key_cb = ttk.Combobox(second, textvariable=selected_key)
    key_cb['state'] = 'readonly'
    key_cb.pack(fill=tk.X, padx=12, pady=5)
    if(idioma):
        key_cb['values'] = ["Yes", "No"]
        selected_key.set("Yes")
    else:
        key_cb['values'] = ["Si", "No"]
        selected_key.set("Si")

    if(idioma): labelRank = ttk.Label(second,text="Inferior Ranking limit -> Mean best results: \n (Recommended 1-3)")
    else: labelRank = ttk.Label(second,text="Ranking inferior limite -> Media mejores resultados: \n (Recomendado 1-3)")
    labelRank.pack(fill=tk.X, padx=5, pady=(20,5))
    entryRank = tk.Entry(second,width=45)
    entryRank.pack()

    if(idioma): labelI = ttk.Label(second,text="REPORT", justify='center',font='Helvetica 12 bold')
    else: labelI = ttk.Label(second,text="INFORME", justify='center',font='Helvetica 12 bold')
    labelI.pack(pady=(40,0))


    if(idioma): labelInfo = ttk.Label(second,text="Compare results with a URL/Query [Y/n]:")
    else: labelInfo = ttk.Label(second,text="Comparar resultados con una URL/Busqueda [S/n]:")
    labelInfo.pack(fill=tk.X, padx=5, pady=5)
    selected_info = tk.StringVar(master=second)
    info_cb = ttk.Combobox(second, textvariable=selected_info, justify='left')
    info_cb['state'] = 'readonly'
    info_cb.pack(fill=tk.X, padx=12, pady=5)
    if(idioma):
        info_cb['values'] = ["Yes", "No"]
        selected_info.set("Yes")
    else:
        info_cb['values'] = ["Si", "No"]
        selected_info.set("Si")


    #Abre la siguiente ventana (de query o de informe) si los datos son correctos (ventana de error en caso contrario)
    def handle_click(event):

        if(buttonB['state'] == 'normal'):
            global numQuery,numBus,limiteInf,keyU,keyQ
            keyU = (info_cb.get() == 'Yes') or (info_cb.get() == 'Si') or (info_cb.get() == '')
            keyQ = (key_cb.get() == 'Yes') or (key_cb.get() == 'Si') or (key_cb.get() == '')
            numQuery= entryQuery.get()
            numBus= entryRes.get()
            limiteInf= entryRank.get()

            numBool = False
            try:
                numQuery = int(numQuery)
                numBus = int(numBus)
                limiteInf = int(limiteInf)
            except ValueError:
                numQuery = ""
                numBus = ""
                limiteInf = ""
            numBool = isinstance(numQuery, int)
            numBool = numBool and isinstance(numBus, int)
            numBool = numBool and isinstance(limiteInf, int)
            if(numBool and ((numQuery<1 or numBus<1 or limiteInf<0) or numBus<limiteInf or limiteInf>10)):
                numBool=False

            if(numBool):
                buttonB['state'] = 'disabled'
                if(keyU):
                    launchInf(second,buttonB,keyQ,numQuery)
                else:
                    launchQue(second,buttonB,keyQ,numQuery,1,[],[])
            else:
                if(idioma): messagebox.showerror("Input Error", "\n Inputs must be Numerical: \n\n (Num. Queries > 0) \n (Num. Results > 0) \n (Limit Ranking <= 10) \n (Limit Ranking <= Num. Results)")
                else: messagebox.showerror("Input Error", "\n Las entradas deben ser numéricas: \n\n (Num. Queries > 0) \n (Num. Resultados > 0) \n (Limite Ranking <= 10) \n (Limite Ranking <= Num. Resultados)")


    #Boton de aceptar
    if(idioma): buttonB = tk.Button(second, text="Accept",height=1, width = 15)
    else: buttonB = tk.Button(second, text="Aceptar",height=1, width = 15)
    buttonB.pack(padx=5, pady=(35,0))

    #Bindings
    buttonB.bind("<Button-1>", handle_click)
    second.bind('<Return>', handle_click)





def launchInf(second,buttonB, aux,aux2):

    """

    Ventana en la que se ha de especificar los detalles de la URL/Query analizar (para crear un informe):
        -Query(Primer reultado)/URL
        -Keywords a utilizar

    param:  "second": Ventana de especificación de detalles.
            "buttonB": Botton de aceptar de la ventana de especificación de detalles (para activarlo/desactivarlo)
            "aux": Booleano que indica si Query=Keywords (Solo necesario para pasarlo a la siguiente ventana)
            "aux2": Número de Queries a buscar (Solo necesario para pasarlo a la siguiente ventana)

    """

    #En caso de cierre volver a venta de especificación de detalles (reiniciar otra busqeuda)
    def on_closing():
        buttonB['state'] = 'normal'
        third.destroy()

    #Crear la ventana
    third = tk.Toplevel(second)
    third.iconbitmap("4693059.ico")
    if(idioma): third.title("SVA - Report Data")
    else: third.title("SVA - Datos Informe")
    third.geometry("300x190")
    third.resizable(False, False)
    third.protocol("WM_DELETE_WINDOW", on_closing)

    #Widgets de la ventana 
    if(idioma): labelQueryInf = ttk.Label(third,text="Insert your URL or Query (Report):")
    else: labelQueryInf = ttk.Label(third,text="Ingrese su URL o Busqueda (Report):")
    labelQueryInf.pack(fill=tk.X, padx=5, pady=5)
    entryQueryInf = tk.Entry(third,width=45)
    entryQueryInf.pack()

    if(idioma): labelQueryInfKey = ttk.Label(third,text="Insert your Keywords:")
    else: labelQueryInfKey = ttk.Label(third,text="Ingrese sus Keywords:")
    labelQueryInfKey.pack(fill=tk.X, padx=5, pady=(20,5))
    entryQueryInfKey = tk.Entry(third,width=45)
    entryQueryInfKey.pack()

    #Boton de Aceptar
    if(idioma): buttonC = tk.Button(third, text="Accept",height=1, width = 15)
    else: buttonC = tk.Button(third, text="Aceptar",height=1, width = 15)
    buttonC.pack(padx=5, pady=(35,0))

    #Abre la siguiente ventana (Queries)
    def handle_click(event):
        global q2,keywords2
        q2 = entryQueryInf.get()
        keywords2 = entryQueryInfKey.get()
        keywords2 = tokenizer.sub(' ', keywords2.lower()).split()
        third.destroy()
        launchQue(second,buttonB,aux,aux2,1,[],[])

    #bindings
    buttonC.bind("<Button-1>", handle_click)
    third.bind('<Return>', handle_click)





def launchQue(second,buttonB,keys,nQuery,nPos,qList,kList):

    """

    Ventana en la que se ha de especificar los detalles de la URL/Query analizar (para crear un informe):
        -Query(Primer reultado)/URL
        -Keywords a utilizar (En caso de que keys=False)

    param:  "second": Ventana de especificación de detalles.
            "buttonB": Botton de aceptar de la ventana de especificación de detalles (para activarlo/desactivarlo)
            "keys": Booleano que indica si Query=Keywords
            "nQuery": Numero de Queries a buscar
            "nPos": Posición de la Query
            "qList": Listado de Queries
            "kList": Listado de Keywords

    """

    #En caso de cierre volver a venta de especificación de detalles (reiniciar otra busqeuda)
    def on_closing():
        buttonB['state'] = 'normal'
        fourth.destroy()

    fourth = tk.Toplevel(second)
    fourth.iconbitmap("4693059.ico")
    titStr= "SVA - Query " + str(nPos) + "/" + str(nQuery)
    fourth.title(titStr)
    if(not keys): fourth.geometry("300x190")
    else: fourth.geometry("300x120")
    fourth.resizable(False, False)
    fourth.protocol("WM_DELETE_WINDOW", on_closing)

    if(idioma): labelQueryQ = ttk.Label(fourth,text="Insert Query:")
    else: labelQueryQ = ttk.Label(fourth,text="Ingrese su Busqueda:")
    labelQueryQ.pack(fill=tk.X, padx=5, pady=5)
    entryQueryQ = tk.Entry(fourth,width=45)
    entryQueryQ.pack()

    if(not keys):
        if(idioma): labelQueryQKey = ttk.Label(fourth,text="Insert your Keywords:")
        else: labelQueryQKey = ttk.Label(fourth,text="Ingrese sus Keywords:")
        labelQueryQKey.pack(fill=tk.X, padx=5, pady=(20,5))
        entryQueryQKey = tk.Entry(fourth,width=45)
        entryQueryQKey.pack()

    if(nQuery==nPos):
        if(idioma): buttonD = tk.Button(fourth, text="Accept",height=1, width = 15)
        else: buttonD = tk.Button(fourth, text="Aceptar",height=1, width = 15)
    else:
        if(idioma): buttonD = tk.Button(fourth, text="Next",height=1, width = 15)
        else: buttonD = tk.Button(fourth, text="Siguiente",height=1, width = 15)
    buttonD.pack(padx=5, pady=(35,0))

    def handle_click(event):

        global queryL,keywordL
        eqq=entryQueryQ.get()
        qList.append(eqq)
        if(not keys):
            eqqk=entryQueryQKey.get()
            keywordQuery = tokenizer.sub(' ', eqqk.lower()).split()
        else:
            if(idioma): stop_words = swEn
            else: stop_words = swEs
            keywordQuery = [w for w in tokenizer.sub(' ', eqq.lower()).split() if not w in stop_words]
        kList.append(keywordQuery)


        if(nQuery==nPos):
            global queryL,keywordL,auxGui
            queryL=qList
            keywordL=kList
            second.destroy()
            fourth.destroy()
            launchWait()
        else:
            fourth.destroy()
            launchQue(second,buttonB,keyQ,numQuery,nPos+1,qList,kList)



    buttonD.bind("<Button-1>", handle_click)
    fourth.bind('<Return>', handle_click)





def launchWait():

    """

    Se lanza el thread que comienza el analisis/busqueda, se muestra Ventana de Espera.
    Cuando el thread finalice se sale de la venta para comenzar el cierre del programa (ventana de cierre).

    """

    #Ilustrar resultados terminal (desarrollador)
    print(idioma)

    print(numQuery)
    print(numBus)
    print(keyQ)
    print(limiteInf)

    print(queryL)
    print(keywordL)

    print(keyU)
    print(q2)
    print(keywords2)

    print("Fin GUI")

    #Lanza el thread que comieza la busqueda/analisis
    threadI= threading.Thread(target=inicializar, daemon=True, args=(idioma,numQuery,numBus,limiteInf,keywordL,queryL,keyU,q2,keywords2))
    threadI.start()

    #Crea página de espera
    wait = tk.Toplevel(root)
    wait.iconbitmap("4693059.ico")
    wait.geometry('300x100')
    wait.resizable(False, False)
    wait.title('SVA wait')

    #Widgets de la página de espera
    if(idioma):
        labelf = ttk.Label(wait,text="Waiting...", justify='center',font='Helvetica 12 bold')
        labelf.pack(pady=(5,10))
        labelf2 = ttk.Label(wait,text="This might take a while. \nLarge queries might take extra time.")
        labelf2.pack(pady=(5,5))
    else:
        labelf = ttk.Label(wait,text="Esperando...", justify='center',font='Helvetica 12 bold')
        labelf.pack(pady=(5,10))
        labelf2 = ttk.Label(wait,text="Esto puede llevar unos minutos. \nQueries largas pueden tomar más tiempo.")
        labelf2.pack(pady=(5,5))

    #Preservar Página de espera mientras el thread esté vivo
   
    global exitAux
    exitAux = True
    while(exitAux and threadI.is_alive()):
        wait.update()
        wait.update_idletasks()

        #Terminación del programa en caso de cerrar ventana
        def on_closing():
            global exitAux
            exitAux = False
        wait.protocol("WM_DELETE_WINDOW", on_closing)
        time.sleep(1)
    
        
    #Comenzar terminación
    if(exitAux): launchTerminacion(wait)
    else: root.quit()



def launchTerminacion(wait):

    """

    Ventana que indica que la ejecución del programa a concluido y se a creado el archivo xmlx.
    
    """

    #Eliminar ventana, Todo ha sido completado
    def on_closing():
        endr.destroy()
        root.quit()

    #Crea Ventana
    endr = tk.Toplevel(root)
    endr.iconbitmap("4693059.ico")
    endr.geometry('500x100')
    endr.resizable(False, False)
    endr.title('SVA End')
    endr.protocol("WM_DELETE_WINDOW", on_closing)
    wait.destroy()

    #Widgets
    if(idioma):
        labelf = ttk.Label(endr,text="Execution Completed", justify='center',font='Helvetica 12 bold')
        labelf.pack(pady=(5,10))
        tstr= "File created: \n" + fileName
        labelf2 = ttk.Label(endr,text=tstr)
        labelf2.pack(pady=(5,5))
    else:
        labelf = ttk.Label(endr,text="Proceso Completado", justify='center',font='Helvetica 12 bold')
        labelf.pack(pady=(5,10))
        tstr= "Fichero generado: \n" + fileName
        labelf2 = ttk.Label(endr,text=tstr)
        labelf2.pack(pady=(5,5))







###################################################################################################################################################

#############################################
###                                       ###
###   Parte 4.2 Main, Cierre y Root(Gui)  ###
###                                       ###
#############################################



if __name__=='__main__':

    """
    
    Creamos el root de la GUI y esperamos a que se complete la creación del archivo xmlx o a que se cierre manuealmente
    el programa, para así poder finalizar el programa.

    """

    print("Inicio GUI")

    #Crea variables globales (Valores Default)
    global q2,keywords2,keyU,keyQ,idioma,numQuery,numBus,limiteInf,queryL,keywordL
    global exitAux, fileName
    exitAux = False
    fileName = ""
    limiteInf = ""
    numBus = ""
    numQuery = ""
    q2 = ""
    keywords2 = ""
    keyU = ""
    keyQ = ""
    idioma = False
    queryL=[]
    keywordL=[]


    #Ventana de seleccion de idioma y Root
    #Creación ventana
    root = tk.Tk()
    root.iconbitmap("4693059.ico")

    root.geometry('300x120')
    root.resizable(False, False)
    root.title('SEO Variables Analyzer')

    #Widgets de la ventana
    label = ttk.Label(text="Select language/Selecciona el idioma:")
    label.pack(fill=tk.X, padx=5, pady=5)

    selected_lan = tk.StringVar()
    lan_cb = ttk.Combobox(root, textvariable=selected_lan)
    lan_cb['values'] = ["English", "Español"]

    lan_cb['state'] = 'readonly'

    lan_cb.pack(fill=tk.X, padx=5, pady=0)

    lan_cb.set("Español")

    #Lanzar ventana de especificación de detalles
    def handle_click(event):
        global idioma
        idioma = selected_lan.get() == "English"
        launch()

    #Boton de Aceptar
    buttonA = tk.Button(root, text="Accept",height=1, width = 15)
    buttonA.pack(padx=5, pady=(35,0))

    #Bindings
    buttonA.bind("<Button-1>", handle_click)
    root.bind('<Return>', handle_click)

    #Bloqueo de la gui hasta quit() o boton X
    root.mainloop()



    #Fin
    print("")
    if(idioma):
        sys.exit("Bye Bye")
    else:
        sys.exit("Adiós")



###################################################################################################################################################



# Fin :)