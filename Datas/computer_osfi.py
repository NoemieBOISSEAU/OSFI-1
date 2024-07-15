# -*- coding: utf-8 -*-
"""
Created on Thu Jul 11 11:18:30 2024

@author: sacha.mailler
"""
import sys
import os
import json
import openpyxl
import statsmodels.api as sm
import time
def return_line(txt,n=52):
    T="\n"
    len_line=0
    for i in range(len(txt)) :
        if i==0 :
            T="\n"
        elif len_line==52 :
            T+="\n"
            len_line=0
        if not(len_line==0 and txt[i]=="\n"):
            if txt[i]=="\n" :
                len_line=0
                T+=txt[i]
            elif txt[i]=="\t":
                T+=" "*4
                len_line+=4
            else :
                T+=txt[i]
                len_line+=1
    return T
#Les objets progression servent à afficher des barres de progressions plutôt  #
#Interactives                                                                 #
###############################################################################
class progression:
    #Initialise l'affichage de la barre de progression :                      #
    #   name : Le nom à afficher au dessus de la barre de progression         #
    ###########################################################################
    def __init__(self,name):
        print(return_line(name,n=52))
        sys.stdout.write(f"\r{' '*52}\r")
        sys.stdout.flush()
        self.progression=0
        self.position=0
        self.max,self.min=100,0
        self.time=0
    #Actuallise le chargement de la barre de progression                      #
    #   progression : pourcentage de progression                              #
    ###########################################################################
    def actualize(self,progression):
        #correction self.min et self.max
        self.min,self.max=max(min(min(self.min,self.max),100),0),max(min(max(self.min,self.max),100),0)
        if progression < 0 :
            progression=0
        if progression > 100 :
            progression=100
        if self.min==self.max :
            progression=self.min
        else :
            progression=self.min+(self.max-self.min)*progression/100
        n=int(progression//2)
        if n==self.progression :
            self.position=(self.position+1)%4
        self.progression=n
        if progression < 100 :
            if self.position==0 :
                car='-'
            if self.position==1:
                car='\\'
            if self.position==2:
                car='|'
            if self.position==3 :
                car='/'
        else :
            car=""
        sys.stdout.flush()
        sys.stdout.write(f"\r{' '*52}\r")
        sys.stdout.write("["+("#"*n)+car+(" "*(49-n))+"]")
        sys.stdout.flush()
    def alert(self,Message):
        print("\n"+Message)
        


class OSFI_regress:
    def __init__(self,path):
        self.path=path
        self.__params={
            "consommation chaud": None,
            "consommation froid" : None,
            "consommation non thermique" : None,
            "Surface" : None,
            "DJU" : None,
            "DJF" : None,
            "typologie" : None,
            "filtres simples":None,
            "verif dqc":None,
            "dqc to osfi":None
            }
        self.__params_num={}
        self.to_remove_dqc=[]
        self.dqc_done=False
        self.excel=None
        self.loaded=False
    def read_json_param(self,name):
        # displayer = progression("Lecture des paramètres dans "+name)
        # displayer.actualize(0)
        with open(os.path.join(self.path,name), 'r', encoding='utf-8') as file :
            temp_dict = json.load(file)
        # displayer.actualize(50)
        for x in self.__params :
            self.__params[x]=temp_dict[x]
        # displayer.actualize(100)
    def add_dqc(self,xls_path):
        self.dqc_done=True
        self.to_remove_dqc=[]
        dqcxlsx = openpyxl.load_workbook(xls_path)
        def get_index(value,sheet):
            flag_found,flag_end=False,False
            i=0
            while not(flag_found or flag_end) :
                i+=1
                if (dqcxlsx[sheet].cell(row=1,column = i).value)==value :
                    flag_found=True
                if (dqcxlsx[sheet].cell(row=1,column = i).value) in [None,"",'""'] :
                    flag_end=True
            if flag_found :
                return i
            return -1
        def get_list(sheet,L_column) :
            count=0
            line=1
            result=[]
            while count<10 :
                line+=1
                flag_void,L=True,[]
                for col in L_column :
                    L.append(dqcxlsx[sheet].cell(row=line,column = col).value)
                    if not(L[-1] in [None,"",'""']) :
                        flag_void = False
                if flag_void:
                    count+=1
                else :
                    count=0
                    result.append(L)
            return result
        for sheet in self.__params["verif dqc"] :
            L_col,L_names=[],[]
            for col in self.__params["verif dqc"][sheet] :
                L_col.append(get_index(col,sheet))
                L_names.append(col)
            rem = get_list(sheet,L_col)
            for x in rem :
                rem2 = {}
                for i in range(len(x)):
                    if L_names[i] in self.__params["dqc to osfi"] :
                        rem2[self.__params["dqc to osfi"][L_names[i]]]=x[i]
                    else :
                        rem2[L_names[i]]=x[i]
                if not(rem2 in self.to_remove_dqc) :
                    self.to_remove_dqc.append(rem2)
        dqcxlsx.close()
    def load_excel(self,xls_name):
        loading=True
        if os.path.exists(os.path.join(self.path,xls_name[:-5]+".json")) :
            self.read_json_param(xls_name[:-5]+".json")
        elif os.path.exists(os.path.join(self.path,"all.json")) :
            self.read_json_param("all.json")
        else :
            loading=False
        if os.path.exists(os.path.join(self.path,"dqc_"+xls_name)) :
            self.add_dqc(os.path.join(self.path,"dqc_"+xls_name))
        if loading :
            # displayer = progression("Lecture du fichier Excel "+xls_name)
            # displayer.actualize(0)
            self.loaded=True
            try :
                self.workbook = openpyxl.load_workbook(os.path.join(self.path,xls_name))
            except :
                self.loaded=False
            # displayer.actualize(10)
            counter=0
            for x in self.__params :
                counter+=1
                # displayer.actualize(10+80*counter/len(self.__params))
                if not(x=="filtres simples") :
                    if self.__params[x] in [None,""]:
                        self.__params_num[x]=[-1]
                    elif isinstance(self.__params[x],list) :
                        self.__params_num[x]=[]
                        for y in self.__params[x] :
                            self.__params_num[x].append(self.get_num_from_col_title(y))
                    elif isinstance(self.__params[x],str) :
                        self.__params_num[x]=[self.get_num_from_col_title(self.__params[x])]
                    elif isinstance(self.__params[x],int) :
                        self.__params_num[x]=[self.get_num_from_col_title(str(self.__params[x]))]
                    elif isinstance(self.__params[x],float) :
                        self.__params_num[x]=[self.get_num_from_col_title(str(self.__params[x]))]
            # displayer.actualize(100)
    def get_num_from_col_title(self,title):
        if self.loaded :
            count,index,flag=0,0,True
            while count<2 and flag :
                value = self.workbook.active.cell(row=1,column=index+1).value
                if value in ['"'+title+'"',title] :
                    return index+1
                elif value in [None,"",'""']:
                    count+=1
                    index+=1
                else :
                    count=0
                    index+=1
            return -1
        else :
            return -1
    def get_list_of_typologies(self) :
        # displayer = progression("Récupération des différentes typologies de bâtiment")
        # displayer.actualize(0)
        Typologies=[]
        if self.loaded :
            if not(self.__params_num["typologie"][0]==-1) :
                count,index=0,0
                # displayer.actualize((1-1/(index/1000+1))*100)
                while count<100 :
                    value = self.workbook.active.cell(row=index+2,column=self.__params_num["typologie"][0]).value
                    if value in [None,"",'""']:
                        count+=1
                        index+=1
                    elif value in Typologies :
                        count=0
                        index+=1
                    else :
                        Typologies.append(value)
                        count=0
                        index+=1
        # displayer.actualize(100)
        return Typologies
    def __to_val(self,row,col):
        if self.loaded :
            value = self.workbook.active.cell(row=row,column=col).value
            if value in [None,"",'""'] :
                return 0
            if isinstance(value,str) :
                try :
                    v=float(value.replace("'",'').replace('"','').replace(",","."))
                except :
                    return 0
            else :
                if isinstance(value,int) or isinstance(value,float) :
                    v=value
            return v
        return 0
    def get_values_for_typologie(self,typologie):
        # displayer = progression("Récupération des grandeurs pour la typologie "+typologie)
        # displayer.actualize(0)
        L_values=[]
        count,index=0,0
        while count<100 :
            # displayer.actualize((1-1/(index/1000+1))*100)
            value = self.workbook.active.cell(row=index+2,column=self.__params_num["typologie"][0]).value
            if value in [None,"",'""']:
                count+=1
                index+=1
            elif value == typologie :
                if self.__params["filtres simples"]==None :
                    for x in self.to_remove_dqc :
                        flag = False
                        for attr in x :
                            attr_index = self.get_num_from_col_title(attr)
                            flag = flag or  (attr_index==-1) or not(self.workbook.active.cell(row=index+2,column=attr_index).value==x[attr])
                            if flag :
                                break
                        if not(flag):
                            break
                    if flag :
                        L_values.append({})
                        for elem in self.__params_num :
                            if not(elem=="typologie") :
                                L_values[-1][elem]=0
                                for new_index in self.__params_num[elem] :
                                    if (new_index>=0) :
                                        L_values[-1][elem]+=self.__to_val(row=index+2,col=new_index)
                        L_values[-1]["typologie"]=typologie
                        L_values[-1]["dqc error"]=False
                    else :
                        L_values.append({})
                        for elem in self.__params_num :
                            if not(elem=="typologie") :
                                L_values[-1][elem]=0
                                for new_index in self.__params_num[elem] :
                                    if (new_index>=0) :
                                        L_values[-1][elem]+=self.__to_val(row=index+2,col=new_index)
                        L_values[-1]["typologie"]=typologie
                        L_values[-1]["dqc error"]=True
                else :
                    flag=True
                    for Filter in self.__params["filtres simples"] :
                        filter_index = self.get_num_from_col_title(Filter)
                        flag = flag and ((filter_index==-1) or (self.workbook.active.cell(row=index+2,column=filter_index).value==self.__params["filtres simples"][Filter]))
                        if not(flag):
                            break
                    if flag :
                        for x in self.to_remove_dqc :
                            flag = False
                            for attr in x :
                                attr_index = self.get_num_from_col_title(attr)
                                flag = flag or  (attr_index==-1) or not(self.workbook.active.cell(row=index+2,column=attr_index).value==x[attr])
                                if flag :
                                    break
                            if not(flag):
                                break
                        if flag :
                            L_values.append({})
                            for elem in self.__params_num :
                                if not(elem=="typologie") :
                                    L_values[-1][elem]=0
                                    for new_index in self.__params_num[elem] :
                                        if (new_index>=0) :
                                            L_values[-1][elem]+=self.__to_val(row=index+2,col=new_index)
                            L_values[-1]["typologie"]=typologie
                            L_values[-1]["dqc error"]=False
                        else :
                            L_values.append({})
                            for elem in self.__params_num :
                                if not(elem=="typologie") :
                                    L_values[-1][elem]=0
                                    for new_index in self.__params_num[elem] :
                                        if (new_index>=0) :
                                            L_values[-1][elem]+=self.__to_val(row=index+2,col=new_index)
                            L_values[-1]["typologie"]=typologie
                            L_values[-1]["dqc error"]=True
                count=0
                index+=1
            else :
                count=0
                index+=1
        # displayer.actualize(100)
        return L_values
        #format de la régression : Conso = A0 + A1*DJU*(S**0.5) + A2*DJF*(S**0.5) +A3*S
    def exclude_neg_surf(self,Lval):
        # displayer = progression("Suppression des bâtiments ayant une surface négative")
        # displayer.actualize(0)
        count=0
        n=len(Lval)
        for i in range(n) :
            # displayer.actualize((i+1)/n*100)
            if Lval[n-1-i]["Surface"]<=0  :
                count+=1
                del Lval[n-1-i]
        # displayer.actualize(100)
        return count
    def exclude_neg_elec(self,Lval):
        # displayer = progression("Suppression des bâtiments ayant une consomation électrique négative")
        # displayer.actualize(0)
        count=0
        n=len(Lval)
        for i in range(n) :
            # displayer.actualize((i+1)/n*100)
            if Lval[n-1-i]["consommation non thermique"]<=0  :
                count+=1
                del Lval[n-1-i]
        # displayer.actualize(100)
        return count
    def exclude_strict_neg_consos(self,Lval):
        # displayer = progression("Suppression des bâtiments ayant une consomation strictement négative")
        # displayer.actualize(0)
        count=0
        n=len(Lval)
        for i in range(n) :
            # displayer.actualize((i+1)/n*100)
            if Lval[n-1-i]["consommation froid"]<0 or Lval[n-1-i]["consommation chaud"]<0  :
                count+=1
                del Lval[n-1-i]
        # displayer.actualize(100)
        return count
    def exclude_dqc(self,Lval) :
        count=0
        n=len(Lval)
        for i in range(n) :
            if Lval[n-1-i]["dqc error"] :
                count+=1
                del Lval[n-1-i]
        return count
    def create_meta_excel(self,excel_name,alpha=0.05):
        LL=[["Typologie","Nombre de Bâtiments initiaux",
             "Bâtiments supprimés","","","","Nombre de bâtiments utilisables",
             "Régression par rapport à la surface","","","",
             "Régression par rapport à la surface et aux DJU","","","",
             ],
            ["Typologie","Nombre de Bâtiments initiaux",
             "Nombre de bâtiments exclus pour surface négative",
             "Nombre de bâtiments exclus pour consommation électrique négative",
             "Nombre de bâtiments exclus pour une consommation strictement négative",
             "Nombre de bâtiments exclus par l'annalyse DQC",
             "Nombre de bâtiments utilisables",
             "r²",
             "Valeur de la consommation",
             "Valeur minimale de la consommation",
             "Valeur maximale de la consommation",
             "r²",
             "Valeur de la consommation",
             "Valeur minimale de la consommation",
             "Valeur maximale de la consommation"]]
        self.load_excel(excel_name)
        List_of_typologies = self.get_list_of_typologies()
        for x in List_of_typologies :
            l = [x]
            L=self.get_values_for_typologie(x)
            l.append(len(L))
            l.append(self.exclude_neg_surf(L))
            l.append(self.exclude_neg_elec(L))
            l.append(self.exclude_strict_neg_consos(L))
            l.append(self.exclude_dqc(L))
            l.append(len(L))
            X1,X2,Y=[],[],[]
            for y in L :
                Y.append(y['consommation chaud']+y['consommation froid']+y['consommation non thermique'])
                X1.append([]),X2.append([])
                X1[-1].append(1)
                X2[-1].append(1)
                X1[-1].append(y["Surface"])
                X2[-1].append(y["Surface"])
                X2[-1].append((y["Surface"]**0.5)*(y["DJU"]))
            if (len(Y)>2) :
                regression1=sm.OLS(Y,X1).fit()
                l.append(regression1.rsquared)
                params=regression1.params
                Iconf=regression1.conf_int(alpha=alpha)
                l.append(str(params[0])+'+'+str(params[1])+'*Surface')
                l.append(str(Iconf[0][1])+'+'+str(Iconf[1][1])+'*Surface')
                l.append(str(Iconf[0][0])+'+'+str(Iconf[1][0])+'*Surface')
                regression2=sm.OLS(Y,X2).fit()
                l.append(regression2.rsquared)
                params=regression2.params
                Iconf=regression2.conf_int(alpha=alpha)
                l.append(str(params[0])+'+'+str(params[1])+'*Surface+'+str(params[2])+'*DJU*(Surface**0.5)')
                l.append(str(Iconf[0][1])+'+'+str(Iconf[1][1])+'*Surface+'+str(Iconf[2][1])+'*DJU*(Surface**0.5)')
                l.append(str(Iconf[0][0])+'+'+str(Iconf[1][0])+'*Surface+'+str(Iconf[2][0])+'*DJU*(Surface**0.5)')
            else :
                for i in range(8) :
                    l.append("Contient moins de 2 batiments")
            LL.append(l)
        self.loaded=False
        try :
            self.workbook.close()
        except :
            self.loaded=True
        workbook = openpyxl.Workbook()
        for line in range(len(LL)):
            for col in range(len(LL[line])) :
                workbook.active.cell(row=line+1,column=col+1).value = str(LL[line][col])
        if (self.dqc_done) :
            excel_name = "dqc_"+excel_name
        workbook.save(os.path.join(self.path,"meta_"+excel_name))
        workbook.close()
if __name__=="__main__" :
    path = os.getcwd()
    XL = OSFI_regress(path)
    L_files = os.listdir(path)
    for file in L_files :
        if not(file.startswith("meta_") or file.startswith("dqc_")) and file.endswith(".xlsx") :
            XL.create_meta_excel(file)
            
            
        
        
