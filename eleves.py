# -*- coding: utf-8 -*-


"""
Documentation

OFFICIELLE :
https://help.libreoffice.org/latest/sq/text/sbasic/python/python_programming.html

Un tuto en anglais, incomplet mais donnant les grandes lignes.
https://tutolibro.tech/category/python/

Une explication pour avoir des box interactives :
https://help.libreoffice.org/latest/sq/text/sbasic/python/python_screen.html?&DbPAR=WRITER&System=UNIX

christopher5106.github.io/office/2015/12/06/openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html
"""



import uno, unohelper
    
from random import sample, randint

#********** Box de messages *************

""" A utiliser pour les tests : on afiche le message qu'on veut dans
une BOX.

_zinfosheet utilise cette box pour donner des informations sur les
sheets.


"""

from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
#trouvé sur https://wiki.documentfoundation.org/Macros/Python_Guide/Useful_functions


CTX = uno.getComponentContext()
SM = CTX.getServiceManager()


def create_instance(name, with_context=False):
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance

def msgbox(message, title='LibreOffice', buttons=MSG_BUTTONS.BUTTONS_OK, type_msg='infobox'):
    """ Create message box
        type_msg: infobox, warningbox, errorbox, querybox, messbox
        https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMessageBoxFactory.html
    """
    toolkit = create_instance('com.sun.star.awt.Toolkit')
    parent = toolkit.getDesktopWindow()
    mb = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    return mb.execute()


#********** Fin box de msg *************

""" les types de cellules. Pratique : j'aurais dû les utiliser 
avant d'écrire mes fonctions de recherche de types.
"""
from com.sun.star.table.CellContentType import TEXT, EMPTY, VALUE

# valeurs de justifications à l'intérieur des cellules
from com.sun.star.table.CellHoriJustify  import STANDARD, RIGHT, LEFT

# valeurs d'orientation dans les cellules
from com.sun.star.table.CellOrientation  import STANDARD, TOPBOTTOM ,BOTTOMTOP ,STACKED 


#from apso_utils import console, msgbox, mri, xray

# ---------------------- OUTILS ----------------------------

#************ NUMEROTATION DES COLONNES ********
"""
passage du nom de colonne à son numéro et réciproquement.

Les numéros de lignes et de colonnes affichés sur le tableur ne 
sont pas ceux qu'il faut utiliser pour getCellByPosition(col,line).

En effet, la numérotation libreoffice commence à 1 mais en Python, on
compte à partir de zéro : il y a un décalage de 1.

"""

def _name2nb(name):
    """
    passe du nom de colonne au numéro 
    ex : colonne ACP -> 26**2+3*26**1+16 soit 770
    """
    n , res = len(name), 0 
    for i in range(n): 
        res = (ord(name[i])-64)+26*res 
    return res

def _nb2name(n):
    """
    passe du num de colonne au nom de colonne
    ex _nb2name(770) donne ACP
    Implantation : galère car ce n'est pas juste un passage de la
    base 10 à la base 26. Il y a des décalages d'indices résolus
    ici sans grande élégance.
    """
    res=[]
    while n!=0:
        r = n%26
        res.append(r)
        n = n//26
    res.reverse()
    def elim_zero():
        for i in range(1,len(res)):
            if res[i]==0:
                res[i-1]-=1
                res[i]=26
    while 0 in res[1:]:
        elim_zero()
    while res[0]==0:
        res = res[1:]
    return "".join([chr(e+64) for e in res])

# ********* MAX MIN dans une colonne ***************

def _max_col(nameCol,fin=1500, sheet=0):
    doc = XSCRIPTCONTEXT.getDocument()
    rg = nameCol+"2:"+nameCol+str(fin)
    cells = doc.Sheets[sheet].getCellRangeByName(rg)
    maxi = 0
    for e in cells.getDataArray():
        try:
            v = float(e[0])
            if v>maxi:
                maxi=v
        except ValueError as ve:
            pass
    return maxi

#************ variables globales (à configurer) ********



def _dichot(sheet,colUAI,colkly,uai,size=800):
    g,d=2, size
    while g < d:
        m=(g+d)//2
        cell=sheet[colUAI+str(m)]
        if cell.getString()==uai:
            return sheet[colkly+str(m)].getString()
        elif cell.getString()<uai:
            g=m+1
        else :
            d=m
    return -1

def uai2kly(doc,code):
    """Passer d'un numéro détablissement à un Kly"""
    doc = XSCRIPTCONTEXT.getDocument()
    NUM_MAX_COLUAI1 = 776#par observation du fichier
    NUM_MAX_COLUAI2 = 612#par observation du fichier
    if code == "":
        return "UAI non renseigné"
    sheet=doc.Sheets["Kly"]
    kly = _dichot(sheet,"A","E",code,NUM_MAX_COLUAI1)
    if kly!=-1:
        return kly
    else:#il y a 2 colonnes d'UAI et donc deux où chercher le kly!!
        kly =  _dichot(sheet,"I","J",code,NUM_MAX_COLUAI2)
        if kly!=-1:
            return kly
    return "Pas de KLY trouvé"

def _positionnement(doc,colcandidat,i):
    #document, col à lire, ligne à lire
    moycandidat = doc.Sheets[0][colcandidat+str(i)].getValue()
    #par observation : la moy basse apparaît 2 colonnes après moyenne candidat
    moyclasse = doc.Sheets[0][_nb2name(_name2nb(colcandidat)+1)+str(i)].getValue()
    #par observation : la moy best apparaît 3 colonnes après moyenne candidat
    moybest = doc.Sheets[0][_nb2name(_name2nb(colcandidat)+3)+str(i)].getValue()
    try :
        pos = (moycandidat-moyclasse)/(moybest - moyclasse)
        #msgbox("moycandidat={},moybasse{},moybest = {}, pos={}".format(moycandidat, moybasse,moybest,pos))
        return pos
    except ZeroDivisionError as ze:
        return ""


    
def findcols_that_contains_string(doc,l,d):
    """ajoute au dictionnaire d les tuples colonnes/sigle dont le titre
    contient les mots de la liste l
    """
    debut = _name2nb("A")#1ere colonne
    fin = _name2nb("ACO")#par observation
    for i in range(debut,fin+1):
        cpt = 2
        col = _nb2name(i)
        cellIN=doc.Sheets[0][col+str(1)]
        test = True
        for s in l:
           test = test and s in cellIN.getString() 
        if test:
           if  cellIN.getString() in d:
                #les notes de 1ere après celles de terminale
               d["{} ({})".format(cellIN.getString(),str(cpt))]=col
               cpt+=1
           else:
               d[cellIN.getString()]=col
    #msgbox(d)



def col2explore(doc):    

    def col_bac():
            """Renvoie un dictionnaire avec les cases à explorer pour chercher les
            résultats des élèves au bac. pour le moment on ne
            s'intéresse qu'au bac de Francais.
            """
            d={}
            #suivant les colonnes oral est écrit Oral ou oral
            findcols_that_contains_string(doc,["épreuve","ral","Français"],d)
            #suivant les colonnes écrit est écrit Ecrit ou écrit
            findcols_that_contains_string(doc,["épreuve","crit","Français"],d)    
            #msgbox(d)
            return d

    
    def col_for_note():
            """Renvoie un dictionnaire avec les cases à explorer pour
            chercher les moyennes des élèves à leurs spécialités"""
            les_spés = ["Mathématiques Spécialité","Physique-Chimie Spécialité",\
                        "Numérique et Sciences Informatiques", "ingénieur","Français","Expertes"]
            d={}
            for e in les_spés:
                findcols_that_contains_string(doc,["candidat",e],d)
            # traitement spécial LVA  : Parfois 'Langue', parfois 'langue'....
            findcols_that_contains_string(doc,["candidat","angue","vivante","A"],d)
            #msgbox(d)
            return d

    def col_sans_note():
            """renvoie un dico où on trouve la colonne du nom, prénom, établissement etc."""
            infos = ["Nom","Prénom","Niveau étude actuel","UAI établissement",\
                     "EDS BAC Terminale","EDS BAC Terminale","EDS BAC Abandonné"]
            d={}
            for e in infos:
                findcols_that_contains_string(doc,[e],d)
            #traitement spécial LVA au bac et dans la scolarité
            findcols_that_contains_string(doc,["LV","A"],d)
            findcols_that_contains_string(doc,["ection","linguistique"],d)
            # un truc idiot : le nombre de parts peut apparaître dans les zinfos :
            asup =""
            for k in d:
                if "parts" in k:
                    asup=k
                    break
            d.pop(k)#on supprime toute mention du nombre de parts
            #msgbox(d)
            return d        

    return col_for_note(), col_sans_note(), col_bac() 

    
def infos_eleve(doc,d1,d2,d3,i):
    """
    if d1={}:
        d1=col_for_note(doc)
    if d2={}
        d2=col_sans_note(doc)
    if d3={}
        d3=col_bac(doc)#provisoirement, seulement le français
    """
    """i est un numéro de ligne"""
    d={}
    #informations générales non quantifiées (nom, UAI, kly, EDS etc)
    for k in d2:
        cellOUT=doc.Sheets[0][d2[k]+str(i)]
        d[k]=cellOUT.getString()
    uai=doc.Sheets[0]["R"+str(i)].getString()#par observation R est la colonne des UAI
    d["Kly"]=uai2kly(doc,uai)
    #notes au bac (seulement le français pour le moment)
    for k in d3:
        cellNOTE = doc.Sheets[0][d3[k]+str(i)]
        try:
           d[k]=int(cellNOTE.getString())
        except ValueError as ve:
           #champ non renseigné
           pass
    #liste des spés de l'élève
    les_spes = []
    les_spes.append(doc.Sheets[0][d2['EDS BAC Terminale']+str(i)].getString())
    les_spes.append(doc.Sheets[0][d2['EDS BAC Terminale (2)']+str(i)].getString())
    les_spes.append(doc.Sheets[0][d2['EDS BAC Abandonné']+str(i)].getString())
    les_spes.extend(["Expertes","Français","vivante"])
    #msgbox(les_spes)
    for spe in les_spes:
        for k in d1:
            if spe in k:#une colonne concernant la spé
                 #msgbox("spe={},k={},d1[k]={}".format(spe,k,d1[k]))
                 try:
                     pos=round(_positionnement(doc,d1[k],i),3)
                     cellNOTE = doc.Sheets[0][d1[k]+str(i)]
                     d[k]=cellNOTE.getString()#la note de l'élève
                     d[k+" : positionnement"]=pos#son positionnement
                 except TypeError as te:
                     pos='non renseigné'
    #msgbox(d)
    return d



    
import time   
"""
FF
"""
def eleves():
    """Donne des zinfos pour chaque élève"""
    doc = XSCRIPTCONTEXT.getDocument()
    NOMBRE_ELEVES=1350#par observation du fichier
    exists_eleves = 'Eleves' in doc.Sheets
    if not exists_eleves:#insertion de la feuille de résultat en position 4
        doc.Sheets.insertNewByName('Eleves', 4)
    # les dictionnaires donnant un numéro de colonne pour chaque nom
    d1,d2,d3=col2explore(doc)
    # d1 colonnes avec des notes
    # d2 colonnes sans note (nom, prénom, kly, LV A suivie...)
    # d3 notes du bac, provisoirement, seulement le français
    #les noms des colonnes
    colonnes = list(d2.keys())+['Kly']+list(d3.keys())
    #ajouter les positionnements aux infos sur les spés
    for k in d1:
        colonnes.append(k)
        colonnes.append(k+" : positionnement")
    #msgbox(d1)
    #Mettre les titre aux colonnes :
    titre2sigle = {}
    for i,name in enumerate(colonnes):
        nom_col =  _nb2name(i+1)
        titre2sigle[name]=nom_col
        cell = doc.Sheets["Eleves"][nom_col+str(1)]
        cell.setString(name)
        cell.Orientation = BOTTOMTOP
    #msgbox(titre2sigle)
    #remplir les résultats par élève :
    def _traitement_eleve(i):
        d_eleve=infos_eleve(doc,d1,d2,d3,i)
        #msgbox({k:v for k,v in titre2sigle.items() if k in d_eleve})
        for k in d_eleve:
            cell = doc.Sheets["Eleves"][titre2sigle[k]+str(i)]
            cell.setString(d_eleve[k])
        #msgbox(d_eleve)
    start=time.time()
    #for i in range(2,NOMBRE_ELEVES):
    for i in range(2,500):
        _traitement_eleve(i)
    msgbox("Exécuté en  : {} s".format(round(time.time()-start),2))
    #_traitement_eleve(39)
    

"""
DEPRECATED

Ci-dessous des fonctions utilisées pour des tests et que je ne veux pas jeter
"""
def section_linguistique(doc=XSCRIPTCONTEXT.getDocument()):
    l=[]
    for i in range(2,1500):
        cell = doc.Sheets[0]["AN"+str(i)]
        l.append(cell.getString())
    msgbox(set(l))

def _zinfos():
    """
    Soulève une exception volotairement.
    Dans la box d'erreur qui s'affiche, on met le message qu'on veut
    """
    doc = XSCRIPTCONTEXT.getDocument()
    cell = doc.Sheets[SHEETOUT]["A1"]
    #assert 1==0, "cel.getElementType()={}".format(cell.getElementType())
    #assert 1==0, "dir(cell)={}".format(dir(cell))
    l = [cell.RotateAngle, cell.Orientation, cell.CellBackColor, cell.CellStyle, cell.HoriJustify]
    cell.setString("GGGGG dans zinfos")
    assert 1==0, "{}".format(l)


def _zinfossheets():
    #https://wiki.documentfoundation.org/Macros/Python_Guide/Calc/Calc_sheets
    doc = XSCRIPTCONTEXT.getDocument()
    count = doc.Sheets.Count
    msg = "nb de sheets : {}\n".format(count)
    #msgbox(count)
    for sheet in doc.Sheets:
        #msgbox(sheet.Name)
        msg=msg+sheet.Name+"\n"
    exists_eleves = 'Eleves' in doc.Sheets
    msg+="existe une sheet Eleves : {}\n".format(exists_eleves)
    msgbox(msg)


def col_lva(doc=XSCRIPTCONTEXT.getDocument()):
    """
    zinfos sur la lva
    """
    d={}
    findcols_that_contains_string(doc,["LV","A"],d)
    msgbox(d)
    
