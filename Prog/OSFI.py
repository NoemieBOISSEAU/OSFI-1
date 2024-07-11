# -*- coding: utf-8 -*-
"""
Created on Wed Jan 24 20:18:26 2024

@author: sacha
"""
from XLS import Excel
from Regression import get_regression, draw_acceptable_html
from _Progression import progression
import os
import json
Base1=os.path.dirname(os.path.dirname(os.getcwd()))
Base=os.path.dirname(Base1)
print(Base)
if Base==Base1 :
    if len(Base1.split("\\")[-1])==0 :
        Base=Base[:-len(Base1.split("\\")[-2])-1]
    else :
        Base=Base[:-len(Base1.split("\\")[-1])-1]
    if Base.endswith("\\") :
        Base=Base+"Data"
    else :
        Base=Base+"\\Data"
else :
    Base = os.path.join(Base,"Data")
print(Base)
Base = os.path.join(os.path.join(Base,"Cerema"),"OSFI")
xls_name="Consommations annuelles des équipements.xlsx"
Lx=["Surface totale du bâtiment",["Degrés jours unifiés","Dégres-jours (DJ) de chauffage"]]
Lx_poste=["Surface totale du bâtiment",["Degrés jours unifiés","Dégres-jours (DJ) de chauffage"]]
Lx_pers=["Surface totale du bâtiment",["Degrés jours unifiés","Dégres-jours (DJ) de chauffage"]]
Lmeta=["Code bâtiment RT","Nom du bâtiment","Code Site","Nom du site","Typologie","Année","Etat du bâtiment"]
Ly=[["Consommation d'électricité (kWh)","Électricité - Consommation"],["Consommation de gaz (kWh PCS)","Gaz - Consommation"],
    ["Consommation du réseau de chaud (kWh)","Réseau de chaleur - Consommation"],["Consommation du réseau de froid (kWh)","Réseau de froid - Consommation"],
    ["Consommation de fioul (kWh PCS)","Fioul - Consommation"],["Consommation de granulés de bois (kWh)","Consommation de granulés de bois (kWh)"]]
def get_admin_regression(progress_path=None):
    if not(progress_path==None) :
        pr=progression("Traitement calcul des régressions à l'échelle nationalle",path=progress_path)
    else :
        pr=progression("Traitement calcul des régressions à l'échelle nationalle")
    admin_path=os.path.join(Base,"__admin__")
    print(admin_path)
    if os.path.exists(os.path.join(admin_path,xls_name)) and os.path.exists(os.path.join(admin_path,"new.txt")) :
        print(1)
        #Lecture du fichier excel :
        XL = Excel(os.path.join(admin_path,xls_name))
        #Récupération des paramètres de la régression
        X=XL.get_typed_lists(Lmeta+Lx,"float")
        for i in range(len(X)) :
            for j in range(len(Lmeta)):
                del X[i][0]
        print(len(X))
        pr.actualize(15)
        #Récupération des métadonnées
        META=XL.get_lists(Lmeta)
        print(len(META))
        pr.actualize(30)
        # for _meta_ in META :
        #     print(_meta_)
        #Récupération de l'image de la régression
        Y=XL.get_typed_lists(Lmeta+Ly,"float")
        print(len(Y))
        for i in range(len(Y)) :
            for j in range(len(Lmeta)):
                del Y[i][0]
        pr.actualize(45)
        XL.close()
        pr.actualize(50)
        #Création des listes de choix des typologies
        Typologies=["Toutes"]
        for _meta_ in META :
            if not((_meta_[4]) in Typologies) :
                Typologies.append(_meta_[4])
        LTP=[]
        pr.actualize(60)
        for typo in Typologies :
            if typo=="Toutes" :
                LTP.append([(True) for _meta_ in META])
            else :
                LTP.append([(_meta_[4]==typo) for _meta_ in META])
        #Création des listes de validation (surface strictement positive, DJU strictements positifs)
        sp,djup=[(_x_[0]>0) for _x_ in X],[(_x_[1]>=0) for _x_ in X]
        #Création des listes de validation (consommation electrique strictement positive, valeur de la consommation totale, toutes consommations positives)
        ye,yt,yp=[_y_[0]>0 for _y_ in Y],[sum(_y_) for _y_ in Y],[(min(_y_)>=0) for _y_ in Y]
        #Création des listes de validation (Code batiment connu, code site connu, batiment ouvert)
        cbp,csp,bo=[not(_meta_[0]=="") for _meta_ in META],[not(_meta_[2]=="") for _meta_ in META],[(_meta_[6]=="Ouvert") for _meta_ in META]
        regress_data={}
        pr.actualize(70)
        for i in range(len(Typologies)) :
            pr.actualize(70+30*((i+1)/len(Typologies)))
            print(Typologies[i])
            _X_,_Y_=[],[]
            for j in range(len(META)) :
                if sp[j] and djup[j] and ye[j] and yp[j] and cbp[j] and csp[j] and bo[j] and LTP[i][j] :
                    print(True)
                    _X_.append(X[j][0])
                    _Y_.append(yt[j])
                else :
                    #print(sp[j], djup[j], ye[j], yp[j], cbp[j], csp[j], bo[j], LTP[i][j])
                    None
            if len(_X_)>1000 :
                coeffs=get_regression(_Y_,_X_,[],[])
            else :
                coeffs=regress_data["Toutes"]
            print(coeffs)
            regress_data[Typologies[i]]=coeffs
        with open(os.path.join(Base,"regress_coefficients.json"),"w",encoding="utf-8") as file :
            json.dump(regress_data, file)
        get_table_from_regression("admin",progress_path=progress_path)
def get_admin_regression_poste(progress_path=None):
    if not(progress_path==None) :
        pr=progression("Traitement calcul des régressions à l'échelle nationalle",path=progress_path)
    else :
        pr=progression("Traitement calcul des régressions à l'échelle nationalle")
    admin_path=os.path.join(Base,"__admin__")
    print(admin_path)
    if os.path.exists(os.path.join(admin_path,xls_name)) and os.path.exists(os.path.join(admin_path,"new.txt")) :
        print(1)
        #Lecture du fichier excel :
        XL = Excel(os.path.join(admin_path,xls_name))
        #Récupération des paramètres de la régression
        X=XL.get_typed_lists(Lmeta+Lx_poste,"float")
        for i in range(len(X)) :
            for j in range(len(Lmeta)):
                del X[i][0]
        print(len(X))
        pr.actualize(15)
        #Récupération des métadonnées
        META=XL.get_lists(Lmeta)
        print(len(META))
        pr.actualize(30)
        # for _meta_ in META :
        #     print(_meta_)
        #Récupération de l'image de la régression
        Y=XL.get_typed_lists(Lmeta+Ly,"float")
        print(len(Y))
        for i in range(len(Y)) :
            for j in range(len(Lmeta)):
                del Y[i][0]
        pr.actualize(45)
        XL.close()
        pr.actualize(50)
        #Création des listes de choix des typologies
        Typologies=["Toutes"]
        for _meta_ in META :
            if not((_meta_[4]) in Typologies) :
                Typologies.append(_meta_[4])
        LTP=[]
        pr.actualize(60)
        for typo in Typologies :
            if typo=="Toutes" :
                LTP.append([(True) for _meta_ in META])
            else :
                LTP.append([(_meta_[4]==typo) for _meta_ in META])
        #Création des listes de validation (surface strictement positive, DJU strictements positifs)
        sp,djup=[(_x_[0]>0) for _x_ in X],[(_x_[1]>=0) for _x_ in X]
        #Création des listes de validation (consommation electrique strictement positive, valeur de la consommation totale, toutes consommations positives)
        ye,yt,yp=[_y_[0]>0 for _y_ in Y],[sum(_y_) for _y_ in Y],[(min(_y_)>=0) for _y_ in Y]
        #Création des listes de validation (Code batiment connu, code site connu, batiment ouvert)
        cbp,csp,bo=[not(_meta_[0]=="") for _meta_ in META],[not(_meta_[2]=="") for _meta_ in META],[(_meta_[6]=="Ouvert") for _meta_ in META]
        regress_data={}
        pr.actualize(70)
        for i in range(len(Typologies)) :
            pr.actualize(70+30*((i+1)/len(Typologies)))
            print(Typologies[i])
            _X_,_Y_=[],[]
            for j in range(len(META)) :
                if sp[j] and djup[j] and ye[j] and yp[j] and cbp[j] and csp[j] and bo[j] and LTP[i][j] :
                    print(True)
                    _X_.append(X[j][0])
                    _Y_.append(yt[j])
                else :
                    #print(sp[j], djup[j], ye[j], yp[j], cbp[j], csp[j], bo[j], LTP[i][j])
                    None
            if len(_X_)>1000 :
                coeffs=get_regression(_Y_,_X_,[],[])
            else :
                coeffs=regress_data["Toutes"]
            print(coeffs)
            regress_data[Typologies[i]]=coeffs
        with open(os.path.join(Base,"regress_coefficients.json"),"w",encoding="utf-8") as file :
            json.dump(regress_data, file)
        get_table_from_regression("admin",progress_path=progress_path)
def get_admin_regression_pers(progress_path=None):
    if not(progress_path==None) :
        pr=progression("Traitement calcul des régressions à l'échelle nationalle",path=progress_path)
    else :
        pr=progression("Traitement calcul des régressions à l'échelle nationalle")
    admin_path=os.path.join(Base,"__admin__")
    print(admin_path)
    if os.path.exists(os.path.join(admin_path,xls_name)) and os.path.exists(os.path.join(admin_path,"new.txt")) :
        print(1)
        #Lecture du fichier excel :
        XL = Excel(os.path.join(admin_path,xls_name))
        #Récupération des paramètres de la régression
        X=XL.get_typed_lists(Lmeta+Lx_pers,"float")
        for i in range(len(X)) :
            for j in range(len(Lmeta)):
                del X[i][0]
        print(len(X))
        pr.actualize(15)
        #Récupération des métadonnées
        META=XL.get_lists(Lmeta)
        print(len(META))
        pr.actualize(30)
        # for _meta_ in META :
        #     print(_meta_)
        #Récupération de l'image de la régression
        Y=XL.get_typed_lists(Lmeta+Ly,"float")
        print(len(Y))
        for i in range(len(Y)) :
            for j in range(len(Lmeta)):
                del Y[i][0]
        pr.actualize(45)
        XL.close()
        pr.actualize(50)
        #Création des listes de choix des typologies
        Typologies=["Toutes"]
        for _meta_ in META :
            if not((_meta_[4]) in Typologies) :
                Typologies.append(_meta_[4])
        LTP=[]
        pr.actualize(60)
        for typo in Typologies :
            if typo=="Toutes" :
                LTP.append([(True) for _meta_ in META])
            else :
                LTP.append([(_meta_[4]==typo) for _meta_ in META])
        #Création des listes de validation (surface strictement positive, DJU strictements positifs)
        sp,djup=[(_x_[0]>0) for _x_ in X],[(_x_[1]>=0) for _x_ in X]
        #Création des listes de validation (consommation electrique strictement positive, valeur de la consommation totale, toutes consommations positives)
        ye,yt,yp=[_y_[0]>0 for _y_ in Y],[sum(_y_) for _y_ in Y],[(min(_y_)>=0) for _y_ in Y]
        #Création des listes de validation (Code batiment connu, code site connu, batiment ouvert)
        cbp,csp,bo=[not(_meta_[0]=="") for _meta_ in META],[not(_meta_[2]=="") for _meta_ in META],[(_meta_[6]=="Ouvert") for _meta_ in META]
        regress_data={}
        pr.actualize(70)
        for i in range(len(Typologies)) :
            pr.actualize(70+30*((i+1)/len(Typologies)))
            print(Typologies[i])
            _X_,_Y_=[],[]
            for j in range(len(META)) :
                if sp[j] and djup[j] and ye[j] and yp[j] and cbp[j] and csp[j] and bo[j] and LTP[i][j] :
                    print(True)
                    _X_.append(X[j][0])
                    _Y_.append(yt[j])
                else :
                    #print(sp[j], djup[j], ye[j], yp[j], cbp[j], csp[j], bo[j], LTP[i][j])
                    None
            if len(_X_)>1000 :
                coeffs=get_regression(_Y_,_X_,[],[])
            else :
                coeffs=regress_data["Toutes"]
            print(coeffs)
            regress_data[Typologies[i]]=coeffs
        with open(os.path.join(Base,"regress_coefficients.json"),"w",encoding="utf-8") as file :
            json.dump(regress_data, file)
        get_table_from_regression("admin",progress_path=progress_path)
def get_table_from_regression(user,subfolder="",progress_path=None):
    if not(progress_path==None) :
        pr=progression("Traitement de données de l'utilisateur "+user,path=progress_path)
    else :
        pr=progression("Traitement de données de l'utilisateur "+user)
    if subfolder=="" :
        xls_path=os.path.join(os.path.join(Base,"__"+user+"__"),xls_name)
        if not(os.path.exists(xls_path)) :
            L=os.listdir(os.path.dirname(xls_path))
            for x in L :
                if x.endswith(".xlsx") :
                    xls_path=os.path.join(os.path.dirname(xls_path),x)
                    break
    else :
        xls_path=os.path.join(os.path.join(os.path.join(Base,"__"+user+"__"),subfolder),xls_name)
        if not(os.path.exists(xls_path)) :
            L=os.listdir(os.path.dirname(xls_path))
            for x in L :
                if x.endswith(".xlsx") :
                    xls_path=os.path.join(os.path.dirname(xls_path),x)
                    break
    if os.path.exists(xls_path) and os.path.exists(os.path.join(os.path.dirname(xls_path),"new.txt")) :
        if os.path.exists(os.path.join(Base,"regress_coefficients.json")) :
            #Lecture des résultats des régressions :
            with open(os.path.join(Base,"regress_coefficients.json"),"r",encoding="utf-8") as file :
                regress_data=json.load(file)
            #Lecture du fichier excel :
            XL = Excel(xls_path)
            #Récupération des paramètres de la régression
            X=XL.get_typed_lists(Lmeta+Lx,"float")
            for i in range(len(X)) :
                for j in range(len(Lmeta)):
                    del X[i][0]
            pr.actualize(15)
            #Récupération des métadonnées
            META=XL.get_lists(Lmeta)
            pr.actualize(30)
            #Récupération des consommations
            Y=XL.get_typed_lists(Lmeta+Ly,"float")
            print(len(Y))
            for i in range(len(Y)) :
                for j in range(len(Lmeta)):
                    del Y[i][0]
            pr.actualize(45)
            XL.close()
            pr.actualize(50)
            #Création des listes de choix des typologies
            Typologies=["Toutes"]
            for _meta_ in META :
                if not((_meta_[4]) in Typologies) :
                    Typologies.append(_meta_[4])
            pr.actualize(60)
            LTP=[]
            for typo in Typologies :
                if typo=="all" :
                    LTP.append([(True) for _meta_ in META])
                else :
                    LTP.append([(_meta_[4]==typo) for _meta_ in META])
            #Création des listes de validation (surface strictement positive, DJU strictements positifs)
            sp,djup=[(_x_[0]>=0) for _x_ in X],[(_x_[1]>=0) for _x_ in X]
            #Création des listes de validation (consommation electrique strictement positive, valeur de la consommation totale, toutes consommations positives)
            ye,yt,yp=[_y_[0]>0 for _y_ in Y],[sum(_y_) for _y_ in Y],[(min(_y_)>=0) for _y_ in Y]
            #Création des listes de validation (Code batiment connu, code site connu, batiment ouvert)
            cbp,csp,bo=[not(_meta_[0]=="") for _meta_ in META],[not(_meta_[2]=="") for _meta_ in META],[(_meta_[6]=="Ouvert") for _meta_ in META]
            pr.actualize(70)
            supp_Lmeta=["Etat de la donnée","Etat détaillé de la donnée"]
            for i in range(len(Typologies)) :
                pr.actualize(70+30*((i+1)/len(Typologies)))
                coeffs=regress_data[Typologies[i]]
                for j in range(len(META)) :
                    if LTP[i][j] :
                        if not bo[j] :
                            META[j].append("Bâtiment fermé")
                            META[j].append("")
                        elif not(sp[j] and djup[j] and ye[j] and yp[j] and cbp[j] and csp[j]) :
                            META[j].append("Erreur")
                            err="<ul>"
                            if not(sp[j]) :
                                err+="<li>Surface négative</li>"
                            if not(djup[j]) :
                                err+="<li>DJU négatif</li>"
                            if not(ye[j]):
                                err+="<li>Consommation électrique négative</li>"
                            if not(yp[j]):
                                err+="<li>Consommations strictement négatives</li>"
                            if not(cbp[j]) :
                                err+="<li>Code bâtiment absent</li>"
                            if not(csp[j]) :
                                err+="<li>Code site absent</li>"
                            err+="</ul>"
                            META[j].append(err)
                        else :
                            A0inf,A0sup,A1inf,A1sup=float(coeffs["A0inf"]),float(coeffs["A0sup"]),float(coeffs["A1inf"]),float(coeffs["A1sup"])
                            #Regarder si on est au dessus, en dessous ou dans la normale
                            META[j].append("A valider")
                            if yt[j]>(A0sup+A1sup*X[j][0]) :
                                META[j].append("Consommations anormalement hautes")
                            elif yt[j]<(A0inf+A1inf*X[j][0]) :
                                META[j].append("Consommations anormalement basses")
                            else :
                                META[j].append("Condommations normales")
            tbl_head=[]
            tbl_body=[]
            tbl="<table>\n\t<thead><tr>"
            for v in Lmeta :
                if isinstance(v,list) :
                    v=v[0]
                tbl_head.append(v)
                tbl+="<th>"+v+"</th>"
            for v in supp_Lmeta :
                if isinstance(v,list) :
                    v=v[0]
                tbl_head.append(v)
                tbl+="<th>"+v+"</th>"
            for v in Lx :
                if isinstance(v,list) :
                    v=v[0]
                tbl_head.append(v)
                tbl+="<th>"+v+"</th>"
            for v in Ly :
                if isinstance(v,list) :
                    v=v[0]
                tbl_head.append(v)
                tbl+="<th>"+v+"</th>"
            tbl+="</tr></thead>\n\t<tbody>\n"
            # tbl_global="<table>\n\t<thead><tr><th>Erreur</th><th>Erreures détaillées</th><th>Nombre</th></tr></thead>"
            
            # tbl_global+="</table>"
            for i in range(len(META)) :
                tbl_body.append([])
                tbl+="\t\t<tr>"
                for j in range(len(META[i])-2):
                    v=META[i][j]
                    tbl+="<td>"+v+"</td>"
                    tbl_body[-1].append(v)
                #style="display:block"
                tbl+='<td>'+META[i][-2]+'</td><td>'+META[i][-1]+'</td>'
                tbl_body[-1].append(META[i][-2])
                tbl_body[-1].append(META[i][-1])
                for v in X[i] :
                    tbl+="<td>"+str(v)+"</td>"
                    tbl_body[-1].append(str(v))
                for v in Y[i] :
                    tbl+="<td>"+str(v)+"</td>"
                    tbl_body[-1].append(str(v))
                tbl+="</tr>\n"
            tbl+="\t</tbody>\n</table>"
            with open(os.path.join(Base,"user.html"),"r",encoding="utf-8") as file :
                txt=file.read()
            if subfolder=="" :
                with open(os.path.join(os.path.join(Base,"__"+user+"__"),user+".html"),"w",encoding="utf-8") as file :
                    file.write(txt.replace("{_{table}_}",tbl).replace("{_{thead}_}",str(tbl_head)).replace("{_{tbody}_}",str(tbl_body)))
            else :
                with open(os.path.join(os.path.join(os.path.join(Base,"__"+user+"__"),subfolder),user+".html"),"w",encoding="utf-8") as file :
                    file.write(txt.replace("{_{table}_}",tbl).replace("{_{thead}_}",str(tbl_head)).replace("{_{tbody}_}",str(tbl_body)))
            os.remove(os.path.join(os.path.dirname(xls_path),"new.txt"))
if __name__=="__main__" :
    import time
    progress_path = os.path.join(Base,"_progression_")
    while True :
        L=os.listdir(Base)
        for x in L :
            print(x)
            if x=="__admin__" :
                get_admin_regression(progress_path=progress_path)
            elif x.startswith('__') and x.endswith('__') and (len(x)>4):
                user=x[2:-2]
                L2=os.listdir(os.path.join(Base,x))
                for y in L2 :
                    get_table_from_regression(user,subfolder=y,progress_path=progress_path)
        pr=progression("En attente de modifications",path=progress_path)
        pr.actualize(0)
        for i in range(100) :
            time.sleep(1)
            pr.actualize(i+1)
        