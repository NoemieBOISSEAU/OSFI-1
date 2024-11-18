# -*- coding: utf-8 -*-
"""
Created on Tue Sep 10 14:57:52 2024

@author: sacha.mailler
"""

import pandas
from Avancement import progression
class Excel:
    def __init__(self,path=None) :
        self.loaded,self.aux_loaded=False,""
        self.Data_frame,self.Aux_Data_frame=None,None
        self.path=path
        if not(path==None) :
            self.read()
    def __is(self,a,b):
        if isinstance(a,list) and not(isinstance(b,list)):
            for x in a :
                if self.__is(x,b):
                    return True
            return False
        if isinstance(b,list) and not(isinstance(a,list)):
            for x in b :
                if self.__is(a,x):
                    return True
            return False
        if a==None :
            return b in [None,"",'""']
        if isinstance(a,float):
            _a_=str(a)
        elif isinstance(a,bool) :
            _a_=str(a)
        elif isinstance(a,int):
            _a_=str(a)
        elif isinstance(a,str):
            _a_=a
            while _a_.startswith('"') and _a_.endswith('"'):
                _a_=_a_[1:-1]
        else :
            _a_=str(a)
        if isinstance(b,float):
            _b_=str(b)
        elif isinstance(b,bool) :
            _b_=str(b)
        elif isinstance(b,int):
            _b_=str(b)
        elif isinstance(b,str):
            _b_=b
            while _b_.startswith('"') and _b_.endswith('"'):
                _b_=_b_[1:-1]
        elif b==None :
            return a in [None,"",'""']
        else :
            _b_=str(b)
        return _a_==_b_
    def __to_num(self,a):
        if a==None :
            return 0
        if isinstance(a,float) or isinstance(a,int) :
            return a
        if isinstance(a,str):
            val=-1
            _a_=a
            while _a_.startswith('"') and _a_.endswith('"') :
                _a_=_a_[1:-1]
            if _a_=="" :
                return 0
            try :
                val=float(_a_)
            except :
                try :
                    val = float(_a_.replace(',','.'))
                except :
                    print("Erreur de convertion : "+a+" n'est pas un nombre")
            return val
        print("Erreur de convertion : "+a+" n'est pas un nombre")
        return -1
    def __create_column(self,name):
        if not(name) in self.Data_frame.columns :
            n= self.Data_frame.shape[0]
            L=["" for i in range(n)]
            self.Data_frame[name]=L
    def read(self,path=None) :
        self.loaded=True
        if path==None :
            PR = progression("Lecture du fichier :"+self.path.split("\\")[-1].split('/')[-1])
            if self.path==None :
                self.loaded = False
            else :
                self.Data_frame = pandas.read_excel(self.path)
        else :
            PR = progression("Lecture du fichier :"+path.split("\\")[-1].split('/')[-1])
            self.Data_frame = pandas.read_excel(path)
        PR.actualize(100)
    def add_values(self,known_element,element_to_add):
        for a in element_to_add :
            self.__create_column(a)
        nrow,ncol = self.Data_frame.shape
        for i in range(nrow):
            flag=False
            for x in known_element :
                flag = not(self.__is(self.Data_frame.at[i,x],known_element[x]))
                if flag :
                    break
            if not(flag) :
                for x in element_to_add :
                    self.Data_frame.loc[i,x] = element_to_add[x]
    def is_in(self,col_names,col_values,final_name):
        LF=[]
        nrow,ncol = self.Data_frame.shape
        for i in range(nrow) :
            for j in range(len(col_values)):
                flag=True
                for k in range(len(col_values[j])):
                    flag = flag and self.__is(self.Data_frame.at[i,col_names[k]],col_values[j][k])
                if flag :
                    break
            LF.append(flag)
        self.Data_frame[final_name]=LF
    def virtual_group_by_sum(self,known_element,to_sum_name,result_prefix="group_",add_count=False) :
        for a in to_sum_name :
            self.__create_column(to_sum_name[a])
        if add_count :
            self.__create_column(result_prefix+"count")
        # PR = progression("Ajout des colonne "+str(to_sum_name)+" de la somme groupée sur "+str(known_element))
        concerned_lines=[]
        # PR.actualize(0)
        n= self.Data_frame.shape[0]
        concerned_lines=[]
        for i in range(n):
            flag=False
            for x in known_element :
                flag = not(self.__is(self.Data_frame.at[i,x],known_element[x]))
                if flag :
                    break
            if not(flag) :
                concerned_lines.append(i)
        count = len(concerned_lines)
        for i in concerned_lines :
            if add_count :
                self.Data_frame.loc[i,result_prefix+"count"]=count
        for name in to_sum_name :
            value=0
            for i in concerned_lines :
                value+=self.__to_num(self.Data_frame.at[i,name])
            for i in concerned_lines :
                self.Data_frame.loc[i,result_prefix+name]=value
    def import_columns_from(self,path,links_main_to_imported,cols_to_import,where={},collapsed="summ",count_imported=False) :
        def to_name():
            name_unique="count_import_"+path.split("\\")[-1].split('/')[-1]
            for x in cols_to_import :
                name_unique+=x
            return name_unique
        uname=to_name()
        for a in cols_to_import :
            self.__create_column(cols_to_import[a])
        if (count_imported) :
            self.__create_column(uname)
        PR = progression("Lecture du fichier :"+path.split("\\")[-1].split('/')[-1])
        PR.actualize(0)
        aux_Data_frame = pandas.read_excel(path)
        PR.actualize(100)
        PR = progression("Filtrage")
        PR.actualize(0)
        ncount,count=len(where),1
        for x in where :
            PR.actualize(count/ncount)
            n= aux_Data_frame.shape[0]
            for j in range(n) :
                if not(self.__is(aux_Data_frame.at[n-1-j,x],where[x])) :
                        aux_Data_frame.drop([n-1-j],axis="index")
        PR.actualize(100)
        n= aux_Data_frame.shape[0]
        n2= self.Data_frame.shape[0]
        PR = progression("Ajout des éléments de la table chargée (partie longue de l'algorithme)")
        for j in range(n2) :
            PR.actualize(99*(j+1)/n2)
            for i in range(n) :
                flag=True
                for x in links_main_to_imported:
                    flag = self.__is(self.Data_frame.at[j,x],aux_Data_frame.at[i,links_main_to_imported[x]])
                    if not(flag) :
                        break
                if flag :
                    if(count_imported) :
                        self.Data_frame.loc[j,uname] = self.__to_num(self.Data_frame.at[j,uname])+1
                    for x in cols_to_import :
                        if collapsed=="summ" :
                            self.Data_frame.loc[j,cols_to_import[x]] = self.__to_num(self.Data_frame.at[j,cols_to_import[x]])+self.__to_num(aux_Data_frame.at[i,x])
                        elif collapsed=="concat" :
                            if str(self.Data_frame.at[j,cols_to_import[x]]) in ["None","",'""',"nan"] :
                                self.Data_frame.loc[j,cols_to_import[x]]=str(aux_Data_frame.at[i,x])
                            else :
                                self.Data_frame.loc[j,cols_to_import[x]]=str(self.Data_frame.at[j,cols_to_import[x]])+";"+str(aux_Data_frame.at[i,x])
                        else :
                            self.Data_frame.loc[j,cols_to_import[x]]=aux_Data_frame.at[i,x]
        PR.actualize(100)
    def get_list_from_col(self,name,known_element={}):
        PR = progression("Récupération des valeurs possibles de la colonne : "+name)
        PR.actualize(0)
        L=[]
        n= self.Data_frame.shape[0]
        for j in range(n) :
            PR.actualize((j+1)/n*100)
            flag=True
            for x in known_element :
                if not(self.__is(self.Data_frame.at[j,x],known_element[x])) :
                    flag=False
                    break
            if flag and not(self.Data_frame.at[j,x] in L) :
                L.append(self.Data_frame.at[j,x])
        PR.actualize(100)
        return L
    def get_list_from_cols(self,names,known_element={}):
        PR = progression("Récupération des valeurs possibles des colonnes : "+str(names))
        PR.actualize(0)
        L=[]
        n= self.Data_frame.shape[0]
        for j in range(n) :
            PR.actualize((j+1)/n*100)
            flag=True
            for x in known_element :
                if not(self.__is(self.Data_frame.at[j,x],known_element[x])) :
                    flag=False
                    break
            if flag :
                l=[]
                for x in names :
                    l.append(self.Data_frame.at[j,x])
                L.append(l)
        PR.actualize(100)
        return L
    def remove(self,known_element):
        PR = progression("Suppressino des éléments : "+str(known_element))
        n= self.Data_frame.shape[0]
        for j in range(n) :
            PR.actualize((j+1)/n*100)
            flag=True
            for x in known_element :
                if not(self.__is(self.Data_frame.at[n-1-j,x],known_element[x])) :
                    flag=False
                    break
            if flag :
                print(n-1-j)
                self.Data_frame.drop([n-1-j],axis="index")
    def __extract_ending_num(self,value):
        if isinstance(value,int):
            return str(value)
        if isinstance(value,float):
            return str(value)
        if isinstance(value,str):
            while value.endswith(" ") :
                value=value[:-1]
            val=""
            n=len(value)
            cp=0
            for i in range(n):
                if value[n-1-i] in [".",","] :
                    if cp==0 :
                        cp=1
                        val=value[n-1-i]+val
                    else :
                        break
                elif value[n-1-i] in "0123456789" :
                    val=value[n-1-i]+val
                else :
                    break
            return val
        return ""
    def extract_ending_num(self,columns):
        for a in columns :
            self.__create_column(columns[a])
        n= self.Data_frame.shape[0]
        for i in range(n) :
            for x in columns :
                self.Data_frame.loc[i,columns[x]] = self.__extract_ending_num(self.Data_frame.loc[i,x])
    def get_stat_by_element(self,col,value,known_element={}):
        L = XL.get_list_from_col(col,known_element)
        i=0
        for x in L :
            print(i)
            i+=1
            known_element[col] = x
            m = self.__get_mean(value,known_element)
            sig = self.__get_sig(value,known_element,m)
            print(m,sig)
            self.add_values(known_element,{"moyenne : "+value:m,"ecart type : "+value:sig})
        
    def __get_sig(self,value,known_element={},m=None):
        variance=0
        nv=0
        if m==None :
            m=self.__get_mean()
        for i in range(len(self.Values)) :
            flag = True
            for x in known_element :
                flag = self.__is(self.Data_frame.at[i,x],known_element[x])
                if not(flag) :
                    break
            if flag :
                if not(self.__to_num(self.Data_frame.at[i,value])==0) :
                    nv=nv+1
                    variance+=(m-self.__to_num(self.Data_frame.at[i,value]))*(m-self.__to_num(self.Data_frame.at[i,value]))
        variance/=nv
        return variance**0.5
    def __get_mean(self,value,known_element={}):
       main_index = self.index(value)
       Indexes=[]
       Values=[]
       nm=0
       for x in known_element :
           Indexes.append(self.index(x))
           Values.append(known_element[x])
       m=0
       for i in range(len(self.Values)) :
           flag = True
           for j in range(len(Indexes)) :
               flag = self.__is(self.Values[i][Indexes[j]],Values[j])
               if not(flag) :
                   break
           if flag :
               if not(self.__to_num(self.Values[i][main_index])==0) :
                   nm=nm+1
                   m+=(self.__to_num(self.Values[i][main_index]))
       m/=nm
       return m
    def save(self,path):
        self.Data_frame.to_excel(path)
    
if __name__=="__main__" :
    import os
    import time
    def get_last_file_of(path):
        maxt=0
        L = os.listdir(path)
        for x in L :
            if x.endswith(".xlsx") :
                maxt = max(os.path.getmtime(os.path.join(path,x)),maxt)
        if not(maxt==0) :
            for x in L :
                if x.endswith(".xlsx") and os.path.getmtime(os.path.join(path,x))==maxt :
                    return os.path.join(path,x)
        return None
    
    path0 = get_last_file_of(os.path.join(os.getcwd(),"OSFI_lien_entre_compteur_et_batiment"))
    path1 = get_last_file_of(os.path.join(os.getcwd(),"OSFI_donnees_du_RT"))
    path2023 = get_last_file_of(os.path.join(os.getcwd(),"OSFI_consommation_mensuelle_2023"))
    path2022 = get_last_file_of(os.path.join(os.getcwd(),"OSFI_consommation_mensuelle_2022"))
    path3 = os.path.join(os.getcwd(),"_result.xlsx")
    oad_path = get_last_file_of(os.path.join(os.getcwd(),"OAD_liste_des_batiments_et_de_leurs_surface"))
    t=time.time()
    XL = Excel(path1)
    XL.remove(known_element={"Typologie du bâtiment":["OUVRAGES D'ART RÉSEAUX VOIRIES","MONUMENT ET MÉMORIAL","ESPACE AMÉNAGÉ","ESPACE NATUREL","RÉSEAUX ET VOIRIES"]})
    XL.extract_ending_num({"Code Site":"Code Site RT"})
    XL.import_columns_from(path=path0,links_main_to_imported={"Identifiant du bâtiment":"Identifiant du bâtiment"},cols_to_import={"Fluide":"Fluides relevés pour le bâtiment"},collapsed="concat")
    
    XL.save(path3[:-5]+"__step1.xlsx")
    # XL = Excel(path3[:-5]+"_step1.xlsx")
    
    XL.import_columns_from(path=path2022,links_main_to_imported={"Identifiant du bâtiment":"Identifiant du bâtiment"},cols_to_import={"Gaz - Consommation":"2022 Gaz - Consommation","Électricité - Consommation":"2022 Électricité - Consommation","Réseau de chaleur - Consommation":"2022 Réseau de chaleur - Consommation","Réseau de froid - Consommation":"2022 Réseau de froid - Consommation","Fioul - Consommation":"2022 Fioul - Consommation","Consommation de granulés de bois":"2022 Consommation de granulés de bois"},where={},collapsed="summ",count_imported=True)
    XL.import_columns_from(path=path2022,links_main_to_imported={"Identifiant du bâtiment":"Identifiant du bâtiment"},cols_to_import={"Gaz - Consommation":"Mai 2022 Gaz - Consommation","Électricité - Consommation":"Mai 2022 Électricité - Consommation","Réseau de chaleur - Consommation":"Mai 2022 Réseau de chaleur - Consommation","Réseau de froid - Consommation":"Mai 2022 Réseau de froid - Consommation","Fioul - Consommation":"Mai 2022 Fioul - Consommation","Consommation de granulés de bois":"Mai 2022 Consommation de granulés de bois"},where={"Date":"2022-05-01"},collapsed="summ",count_imported=False)
    XL.import_columns_from(path=path2022,links_main_to_imported={"Identifiant du bâtiment":"Identifiant du bâtiment"},cols_to_import={"Gaz - Consommation":"Hiver 2022 Gaz - Consommation","Électricité - Consommation":"Hiver 2022 Électricité - Consommation","Réseau de chaleur - Consommation":"Hiver 2022 Réseau de chaleur - Consommation","Réseau de froid - Consommation":"Hiver 2022 Réseau de froid - Consommation","Fioul - Consommation":"Hiver 2022 Fioul - Consommation","Consommation de granulés de bois":"Hiver 2022 Consommation de granulés de bois"},where={"Date":["2022-01-01","2022-02-01","2022-01-01","2022-03-01","2022-04-01","2022-05-01","2022-10-01","2022-11-01","2022-12-01"]},collapsed="summ",count_imported=False)
    XL.import_columns_from(path=path2022,links_main_to_imported={"Identifiant du bâtiment":"Identifiant du bâtiment"},cols_to_import={"Gaz - Consommation":"Ete 2022 Gaz - Consommation","Électricité - Consommation":"Ete 2022 Électricité - Consommation","Réseau de chaleur - Consommation":"Ete 2022 Réseau de chaleur - Consommation","Réseau de froid - Consommation":"Ete 2022 Réseau de froid - Consommation","Fioul - Consommation":"Ete 2022 Fioul - Consommation","Consommation de granulés de bois":"Ete 2022 Consommation de granulés de bois"},where={"Date":["2022-06-01","2022-07-01","2022-08-01","2022-09-01"]},collapsed="summ",count_imported=False)
    
    XL.save(path3[:-5]+"__step2.xlsx")
    # XL = Excel(path3[:-5]+"_step2.xlsx")
    
    XL.import_columns_from(path=path2023,links_main_to_imported={"Identifiant du bâtiment":"Identifiant du bâtiment"},cols_to_import={"Gaz - Consommation":"2023 Gaz - Consommation","Électricité - Consommation":"2023 Électricité - Consommation","Réseau de chaleur - Consommation":"2023 Réseau de chaleur - Consommation","Réseau de froid - Consommation":"2023 Réseau de froid - Consommation","Fioul - Consommation":"2023 Fioul - Consommation","Consommation de granulés de bois":"2023 Consommation de granulés de bois"},where={},collapsed="summ",count_imported=True)
    XL.import_columns_from(path=path2023,links_main_to_imported={"Identifiant du bâtiment":"Identifiant du bâtiment"},cols_to_import={"Gaz - Consommation":"Mai 2023 Gaz - Consommation","Électricité - Consommation":"Mai 2023 Électricité - Consommation","Réseau de chaleur - Consommation":"Mai 2023 Réseau de chaleur - Consommation","Réseau de froid - Consommation":"Mai 2023 Réseau de froid - Consommation","Fioul - Consommation":"Mai 2023 Fioul - Consommation","Consommation de granulés de bois":"Mai 2023 Consommation de granulés de bois"},where={"Date":"2023-05-01"},collapsed="summ",count_imported=False)
    XL.import_columns_from(path=path2023,links_main_to_imported={"Identifiant du bâtiment":"Identifiant du bâtiment"},cols_to_import={"Gaz - Consommation":"Hiver 2023 Gaz - Consommation","Électricité - Consommation":"Hiver 2023 Électricité - Consommation","Réseau de chaleur - Consommation":"Hiver 2023 Réseau de chaleur - Consommation","Réseau de froid - Consommation":"Hiver 2023 Réseau de froid - Consommation","Fioul - Consommation":"Hiver 2023 Fioul - Consommation","Consommation de granulés de bois":"Hiver 2023 Consommation de granulés de bois"},where={"Date":["2023-01-01","2023-02-01","2023-01-01","2023-03-01","2023-04-01","2023-05-01","2023-10-01","2023-11-01","2023-12-01"]},collapsed="summ",count_imported=False)
    XL.import_columns_from(path=path2023,links_main_to_imported={"Identifiant du bâtiment":"Identifiant du bâtiment"},cols_to_import={"Gaz - Consommation":"Ete 2023 Gaz - Consommation","Électricité - Consommation":"Ete 2023 Électricité - Consommation","Réseau de chaleur - Consommation":"Ete 2023 Réseau de chaleur - Consommation","Réseau de froid - Consommation":"Ete 2023 Réseau de froid - Consommation","Fioul - Consommation":"Ete 2023 Fioul - Consommation","Consommation de granulés de bois":"Ete 2023 Consommation de granulés de bois"},where={"Date":["2023-06-01","2023-07-01","2023-08-01","2023-09-01"]},collapsed="summ",count_imported=False)
    
    XL.save(path3[:-5]+"__step3.xlsx")
    # XL = Excel(path3[:-5]+"_step3.xlsx")
    
    XL.import_columns_from(path=oad_path,links_main_to_imported={"Code bâtiment RT":"Code bât/ter","Code Site RT":"Code Site"},cols_to_import={"Surface de plancher":"SDP RT","SUB":"SUB RT"},collapsed="last")
    
    XL.save(path3[:-5]+"__step4.xlsx")
    # XL = Excel(path3[:-5]+"_step4.xlsx")
    
    L=XL.get_list_from_col("Code Site",known_element={"Nombre d éléments importés":0})
    for x in L :
        XL.virtual_group_by_sum(known_element={"Code Site":x}, to_sum_name=["Surface au sol","Gaz - Consommation","Électricité - Consommation","Réseau de chaleur - Consommation","Réseau de froid - Consommation","Fioul - Consommation","Consommation de granulés de bois"],result_prefix="Regroupé par site ",add_count=True)
    #XL = Excel(path3[:-5]+"_step2.xlsx")
    XL.import_columns_from(path=oad_path,links_main_to_imported={"Code bâtiment RT":"Code bât/ter","Code Site RT":"Code Site"},cols_to_import={"Surface de plancher":"SDP RT","SUB":"SUB RT"},collapsed="last")
    L = XL.get_list_from_cols(["Code Site","Code bâtiment RT"],{})
    for [x,y] in L :
        XL.virtual_group_by_sum(known_element={"Code Site":x,"Code bâtiment RT":y}, to_sum_name=["Surface au sol"],result_prefix="Regroupé par bâtiment")
    XL.is_in(["Typologie du bâtiment","Typologie détaillée"],[["BÂT. ENSEIGNEMENT OU SPORT"],
		["BÂTIMENT CULTUREL"],
		["BATIMENT SANITAIRE OU SOCIAL"],
		["BUREAU"],
		["BATIMENT TECHNIQUE","BÂTIMENT TECHNIQUE"],
		["BATIMENT TECHNIQUE","BÂTIMENT SCIENTIFIQUE"],
		["BATIMENT TECHNIQUE","CENTRE DE RECHERCHE OU D'ESSAI"],
		["BATIMENT TECHNIQUE","DÉPÔT D'ARCHIVES"],
		["BATIMENT TECHNIQUE","CENTRE INFORMATIQUE"],
		["BATIMENT TECHNIQUE","POSTE DE COMMANDEMENT"]],"dans le périmètre choisi")
    #XL.get_stat_by_element(col = "Typologie du bâtiment", value = "Surface au sol")
    XL.save(path3)
    XL.close()
    print(str(int(time.time()-t))+" secondes")