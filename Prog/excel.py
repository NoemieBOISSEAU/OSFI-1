# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 14:28:40 2025

@author: sacha.mailler
"""

from reader import basic_reader as openpyxl
import random
from Avancement import progression
class str_date:
    def names(self):
        return ["Année","Mois","Jour"]
    def calculate(self,value):
        return value[0].split('-')
class Excel :
    def __init__(self):
        self.prev_path=None
        self.path=None
        self.workbok = None
        self.L_names = []
        self.L_values = []
        self.max_row = None
        self.max_line = None
    def read(self,path=None,read_only=True):
        avancement = progression("Reading : "+str(path))
        avancement.actualize(0)
        self.prev_path = self.path
        if not(self.path==None) :
            if not(self.path==path) :
                avancement = progression("Reading : "+str(path))
                avancement.actualize(0)
                self.workbook.close()
                avancement.actualize(50)
                self.workbook = openpyxl().load_workbook(path=self.path,read_only=read_only)
                avancement.actualize(100)
        if not(path==None) :
            avancement = progression("Reading : "+str(path))
            avancement.actualize(0)
            self.path = path
            self.workbook = openpyxl().load_workbook(path=self.path,read_only=read_only)
            avancement.actualize(100)
    def close(self):
        avancement = progression("Closing : "+str(path))
        avancement.actualize(0)
        self.prev_path = self.path
        if not(self.path==None) :
            self.workbook.close()
            del self.workbook
            self.path=None
        avancement.actualize(100)
    def __get_max_col(self,max_void=5):
        col,count_void=0,0
        while count_void<max_void :
            col+=1
            if self.workbook.active.cell(row=1,column=col).value in [None,""] :
                count_void+=1
            else :
                count_void = 0
        return col-count_void
    def __get_max_line(self,max_col=None,max_void = 5):
        if max_col==None :
            max_col = self.__get_max_col(max_void)
        def is_void_line(line):
            flag = True
            for i in range(max_col) :
                flag = self.workbook.active.cell(row=line,column=i+1).value in [None,""]
                if not(flag) :
                    break
            return flag
        line,count_void=1,0
        while count_void<max_void :
            line+=1
            if is_void_line(line) :
                count_void+=1
            else :
                count_void = 0
        return line-count_void
    def __index(self,name, simplified=False,max_col=None):
        if max_col==None :
            if self.path==None :
                max_col = len(self.L_names)
            else :
                max_col  = self.__get_max_col()
        to_replace = "äâàABCDêéèëEFGHIîïJKLMNOôöPQRSTUûüùVWXYZ"
        replacer   = "aaaabcdeeeeefghiiijklmnooopqrstuuuuvwxyz"
        def simplify(word):
            if simplified :
                result=""
                for i in range(len(word)) :
                    if word[i] in replacer :
                        result+=word[i]
                    elif word[i] in to_replace :
                        result+=replacer[to_replace.index(word[i])]
                return result
            else :
                return word
        for i in range(max_col) :
            if self.path==None :
                if simplify(self.L_names[i])==simplify(name) :
                    return i
            else :
                if simplify(self.workbook.active.cell(row=1,column=i+1).value)==simplify(name) :
                    return i
        return -1
    def load(self,path,dict_of_names,rad="",term="",read_only=True):
        flag = True
        for x in dict_of_names :
            if rad+x+term in self.L_names :
                flag = False
                raise Exception("Error : La grandeur "+rad+x+term+ "à charger à déjà été chargée ! ")
        if flag :
            self.read(path,read_only=read_only)
            avancement = progression("Loadig : "+str(dict_of_names))
            avancement.actualize(0)
            line_num = self.__get_max_line()
            n,_i_=len(dict_of_names),-1
            for name in dict_of_names :
                _i_+=1
                avancement.actualize(_i_/n)
                if isinstance(dict_of_names[name],list) :
                    for s_name in dict_of_names[name] :
                        ind = self.__index(s_name,simplified=False)
                        if not(ind==-1) :
                            break
                    if ind == -1 :
                        for s_name in dict_of_names[name] :
                            ind = self.__index(s_name,simplified=True)
                            if not(ind==-1) :
                                break
                else :
                    ind = self.__index(s_name,simplified=False)
                    if ind == -1 :
                        ind = self.__index(s_name,simplified=True)
                if ind==-1 :
                    print("Error : Aucune des grandeurs "+str(dict_of_names[name])+" n'a pas été trouvée dans le fichier"+self.path)
                else :
                    self.L_names.append(rad+name+term)
                    for j in range(line_num-1):
                        if len(self.L_values)<=j :
                            self.L_values.append([])
                            for i in range(len(self.L_names)-1) :
                                self.L_values[-1].append(None)
                        self.L_values[j].append(self.workbook.active.cell(row=j+2,column=ind+1).value)
            avancement.actualize(100)
    def calculate(self,calculation_object,L_inputs) :
        avancement = progression("Calculating element from : "+str(L_inputs))
        avancement.actualize(0)
        L = calculation_object.names()
        L_indexes = []
        for x in L :
            if x in self.L_names :
                raise Exception("Error : Un des éléments ("+x+") à calculer est déjà dans le tableau")
            else :
                self.L_names.append(x)
        for x in L_inputs :
            index = self.__index(x,simplified=False)
            if index == -1 :
                index = self.__index(x,simplified=True)
                if index == -1 :
                    raise Exception("Error : Entrée de calcul ("+x+")non trouvée")
            L_indexes.append(index)
        n  = len(self.L_values)
        for i in range(n) :
            avancement.actualize(i/n)
            inp=[]
            for j in L_indexes :
                inp.append(self.L_values[i][j])
                result = calculation_object.calculate(inp)
                for k in range(len(L)):
                    self.L_values[i].append(result[k])
        avancement.actualize(100)
    def remove_col(self,List_of_cols):
        L_indexes = []
        for x in List_of_cols :
            index = self.__index(x,simplified=False)
            if index==-1 :
                index = self.__index(x,simplified=True)
                if index == -1 :
                    raise Exception("Error : la colonne à supprimer n'existe pas")
            L_indexes.append(index)
        L_indexes.sort(reverse=True)
        for index in L_indexes :
            del self.L_names[index]
        for i in range(len(self.L_values)) :
            for index in L_indexes :
                del self.L_values[i][index]
    def sort(self,elements_to_sort):
        def num_compear(A,B) :
            a,b=A,B
            if a==None :
                return (b==None)*1
            if b==None :
                return 2
            if isinstance(a,str):
                while a.startswith("0") :
                    a = a[1:]
                if a=="" :
                    a="0"
                if isinstance(b,str):
                    while b.startswith("0") :
                        b = b[1:]
                    if b=="" :
                        b="0"
                    return (float(a.replace(",","."))>=float(b.replace(",",".")))*1+(float(a.replace(",","."))>float(b.replace(",",".")))*1
                else :
                    return (float(a.replace(",","."))>=b)*1+(float(a.replace(",","."))>b)*1
            if isinstance(b,str):
                while b.startswith("0") :
                    b = b[1:]
                if b=="" :
                    b="0"
                return (a>=float(b.replace(",",".")))*1+(a>float(b.replace(",",".")))*1
            else :
                return (a>=b)*1+(a>b)*1
        def isnum(elem) :
            if elem==None :
                return True
            if isinstance(elem,int):
                return True
            if isinstance(elem,float):
                return True
            if isinstance(elem,str):
                for x in elem :
                    if not(x in "0123456789,."):
                        return False
                return (elem.count('.')+elem.count(','))<=1
            return False
        def is_num(column) :
            for i in range(len(self.L_values)):
                if not(isnum(self.L_values[i][column])) :
                    return False
            return True
        L_indexes = []
        for x in elements_to_sort :
            i = self.__index(x,simplified=False)
            if i==-1 :
                i = self.__index(x,simplified=True)
                if i==-1 :
                    raise Exception("Error : l'élément sur lequel trier "+x+" n'existe pas")
            if i>=0 :
                L_indexes.append(i)
        for index in L_indexes : 
            if not(is_num(index)) :
                raise Exception("Error : l'algorithme de tri n'est codé que pour les colonnes numériques ce qui n'est pas le cas de la colonne "+self.L_names[index])
        flag = True
        while flag :
            flag = False
            for i in range(len(self.L_values)-1):
                to_change = False
                for index in L_indexes :
                    check = num_compear(self.L_values[i][index],self.L_values[i+1][index])
                    if check == 0 :
                        break
                    elif check == 2 :
                        to_change = True
                if to_change :
                    flag = True
                    self.L_values[i],self.L_values[i+1] = self.L_values[i+1],self.L_values[i]
    def linearize(self,id_name,name_to_linearize,sorted_liearrized=True) :
        if (sorted_liearrized) :
            self.sort(name_to_linearize)#self.sort(id_name+name_to_linearize)
        fixed_indexes = []
        prev_indexes = []
        for x in id_name :
            a = self.__index(x,simplified=False)
            if a==-1 :
                a = self.__index(x,simplified=True)
                if a==-1 :
                    raise Exception("Error : Attribut fixe de la linéarisation "+x+" non trouvé")
            if a>=0 :
                fixed_indexes.append(a)
        for x in name_to_linearize :
            a = self.__index(x,simplified=False)
            if a==-1 :
                a = self.__index(x,simplified=True)
                if a==-1 :
                    raise Exception("Error : Attribut fixe de la linéarisation "+x+" non trouvé")
            if a>=0 :
                prev_indexes.append(a)
        i=0
        #vérifications des valeurs à linéariser
        L_fixed,L_var=[],[]
        for k in range(len(self.L_names)):
            if not((k in fixed_indexes) or (k in prev_indexes)) :
                L_fixed.append(k)
        to_avoid=[]
        while i<len(self.L_values) :
            to_avoid.append(i)
            j=i+1
            while j<len(self.L_values) :
                flag = True
                for index in fixed_indexes :
                    if not(self.L_values[i][index]==self.L_values[j][index]) :
                        flag = False
                        break
                if flag :
                    to_avoid.append(j)
                    for k in range(len(L_fixed)):
                        if not(self.L_values[i][L_fixed[len(L_fixed)-1-k]]==self.L_values[j][L_fixed[len(L_fixed)-1-k]]) :
                            L_var.append(L_fixed[len(L_fixed)-1-k])
                            del L_fixed[len(L_fixed)-1-k]
                j+=1
            while i in to_avoid :
                i+=1
        L_var.sort()
        i=0
        #vérifications des valeurs à linéariser
        while i<len(self.L_values) :
            j=i+1
            L_lines=[i]
            while j<len(self.L_values) :
                flag = True
                for index in fixed_indexes :
                    if not(self.L_values[i][index]==self.L_values[j][index]) :
                        flag = False
                        break
                if flag :
                    L_lines.append(j)
                j+=1
            delta=0
            L_lines.sort()
            for j in L_lines :
                prefix=""
                for index in prev_indexes :
                    if self.L_values[j-delta][index] ==None :
                        prefix+="_"
                    else :
                        prefix+=(str(self.L_values[j-delta][index])+"_")
                for k in L_var :
                    name = prefix+self.L_names[k]
                    index =self.__index(name)
                    if index==-1 :
                        index = len(self.L_names)
                        self.L_names.append(name)
                        for l in range(len(self.L_values)):
                            self.L_values[l].append(None)
                    if not(self.L_values[i][index]==None):
                        self.__display([i,j])
                        raise Exception("Error : L'une des variables à affecter à déjà été affectée")
                    else :
                        self.L_values[i][index] = self.L_values[j-delta][k]
                        self.L_values[j-delta][k] = None
                if not(j==i) :
                    delta+=1
                    del self.L_values[j-delta]
            i+=1
    def __display(self,L):
        print("")
        L_size=[]
        for j in range(len(self.L_names)) :
            L_size.append(len(self.L_names[j]))
        for i in range(len(L)) :
            for j in range(len(self.L_names)) :
                try :
                    L_size[j]=max(L_size[j],len(str(self.L_values[L[i]][j])))
                except Exception :
                    print(i,j)
                    raise Exception
        txt=""
        for j in range(len(self.L_names)) :
            txt+=self.L_names[j]+" "*(L_size[j]-len(self.L_names[j]))+"|"
        print(txt)
        for i in range(len(L)) :
            txt=""
            for j in range(len(self.L_names)) :
                txt+=str(self.L_values[L[i]][j])+" "*(L_size[j]-len(str(self.L_values[L[i]][j])))+"|"
            print(txt)
    def display(self):
        L=[]
        for i in range(100) :
            L.append(random.randint(0,len(self.L_values)))
        L.sort()
        self.__display(L)
    def save(self,path):
        wb = openpyxl().Workbook()
        for i in range(len(self.L_names)) :
            wb.active.cell(row=1,column=i+1).value = self.L_names[i]
            for j in range(len(self.L_values)):
                wb.active.cell(row=j+2,column=i+1).value = self.L_values[j][i]
        wb.save(path)
        wb.close()
    def save_separately(self,L_columns):
        if not(os.path.exists(os.path.join(os.path.dirname(self.prev_path),"__work__"))) :
            os.mkdir(os.path.join(os.path.dirname(self.prev_path),"__work__"))
        def egual(L1,L2):
            n = len(L1)
            if n==len(L2) :
                for i in range(n):
                    if not(L1[i]==L2[i]) :
                        return False
                return True
            return False
        def get(L,line):
            val = []
            for x in L :
                if isinstance(x,int):
                    index = x
                elif isinstance(x,str):
                    index = self.__index(x,simplified=False)
                    if index == -1 :
                        index = self.__index(x,simplified=True)
                if index == -1 :
                    val.append(None)
                else :
                    val.append(self.L_values[line][index])
            return val
        def inside(L,line):
            for x in L :
                if egual(x,line) :
                    return True
            return False
        Values = []
        for line in range(len(self.L_values)) :
            value = get(L_columns,line)
            if not(inside(Values,value)):
                Values.append(value)
        for value in Values :
            prev=""
            for x in value :
                prev += (str(x)+"_")
            wb = openpyxl().Workbook()
            path = os.path.join(os.path.join(os.path.dirname(self.prev_path),"__work__"),prev+os.path.basename(self.prev_path))
            for i in range(len(self.L_names)) :
                wb.active.cell(row=1,column=i+1).value = self.L_names[i]
            line_num = 1
            for i in range(len(self.L_values)) :
                if egual(get(L_columns,i),value) :
                    line_num += 1
                    for j in range(len(self.L_names)) :
                        wb.active.cell(row=line_num,column=j+1).value = self.L_values[i][j]
            wb.save(path)
            wb.close()
    def __exit__(self,exc_type,exc_value,traceback):
        if not(self.path==None) :
            self.workbook.close()
        print("Type : "+str(exc_type))
        print("Value : "+str(exc_value))
        print("Traceback : "+str(traceback))
    def get(self):
        return self.L_names,self.L_values
        
if __name__=="__main__" :
    import os
    import time
    #path = os.path.join(os.path.join(os.path.dirname(os.getcwd()),"Datas"),"Consommations_mensualisees_des_equipements.xlsx")
    path = "C:\\Users\\sacha.mailler\\Desktop\\GIT\\OSFI\\Datas\\Consommations_mensualisees_des_equipements.xlsx"
    IDS = {"ID":["Identifiant du bâtiment"],"DATE":["Date"]}
    Consos = {"DJU":["Degrés-jours (DJ) de chauffage"],
             "DJF":["Degrés-jours (DJ) de refroidissement"],
             "fioul":["Fioul - Consommation"],
             "gaz":["Gaz - Consommation"],
             "rcu":["Réseau de chaleur - Consommation"],
             "bois":["Consommation de granulés de bois"],
             "froid":["Réseau de froid - Consommation"],
             "elec":["Électricité - Consommation"]}
    Datas = {"code bat RT":["Code bâtiment RT"],
             "code site RT":["Code Site"],
             "surface":["Surface au sol"],
             "typologie 1":["Typologie du bâtiment"],
             "typologie 2": ["Typologie détaillée"],}
    E = Excel()
    t0 = time.time()
    E.load(path,IDS,read_only=False)
    E.load(path,Datas,read_only=False)
    E.load(path,Consos,read_only=False)
    E.close()
    t1=time.time()
    E.calculate(str_date(),["DATE"])
    t2=time.time()
    E.remove_col(["DATE","Jour"])
    E.linearize(id_name=["ID","Année"],name_to_linearize=["Mois"])
    t3 = time.time()
    E.save_separately(["ID"])
    del E
    print("reading : "+str(int(t1-t0))+"s")
    print("calculating : "+str(int(t2-t1))+"s")
    print("linearising : "+str(int(t3-t2))+"s")
    