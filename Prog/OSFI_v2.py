# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 15:21:31 2024

@author: sacha.mailler
"""
import openpyxl
class Excel:
    def __init__(self,path=None) :
        self.Names=[]
        self.Valuse=[]
        self.flag=False
        self.path=path
        if not(path==None) :
            self.read()
    def __is(self,a,b):
        if a==None :
            return b in [None,"",'""']
        if isinstance(a,float):
            _a_=str(a)
        elif isinstance(a,int):
            _a_=str(a)
        elif isinstance(a,str):
            _a_=a
            while _a_.startswith('"') and _a_.endswith('"'):
                _a_=_a_[1:-1]
        if isinstance(b,float):
            _b_=str(b)
        elif isinstance(b,int):
            _b_=str(b)
        elif isinstance(b,str):
            _b_=b
            while _b_.startswith('"') and _b_.endswith('"'):
                _b_=_b_[1:-1]
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
    def read(self,path=None) :
        self.Names=[]
        self.Values=[]
        self.flag=False
        if path==None :
            if self.path==None :
                self.flag = False
            else :
                workbook = openpyxl.load_workbook(self.path)
        else :
            workbook = openpyxl.load_workbook(path)
        for col in range(workbook.active.max_column):
            self.Names.append(workbook.active.cell(row=1,column=col+1).value)
        for rw in range(workbook.active.max_row-1):
            self.Values.append([])
            for col in range(workbook.active.max_column):
                self.Values[-1].append(workbook.active.cell(row=rw+2,column=col+1).value)
        workbook.close()
    def index(self,name):
        if isinstance(name,str):
            flag=False
            for i in range(len(self.Names)):
                if self.__is(self.Names[i],name):
                    flag=True
                    break
            if flag :
                return i
            return -1
        print("Erreur de repérage : le nom d'une colonne doit être une chaîne de caractères")
    def add_values(self,known_element,element_to_add):
        for x in element_to_add :
            if self.index(x)==-1 :
                self.Names.append(x)
                for i in range(len(self.Values)):
                    self.Values[i].append(None)
        I=[]
        I2=[]
        for x in known_element :
            i=self.index(x)==-1 
            if i==-1 :
                print("Erreur d'attribution d'une valeur : l'un des éléments à reconnaitre n'est pas dans le fichier de base")
            else :
                I.append(i)
                I2.append(x)
        for i in range(len(self.Values)):
            flag=False
            for j in range(I):
                flag = flag or not(self.__is(self.Values[i][I[j]],known_element[I2[i]]))
                if flag :
                    break
            if not(flag) :
                for x in element_to_add :
                    self.Values[i][self.index(x)] = element_to_add[x]
    def sum_values(self,known_element,element_to_add):
        for x in element_to_add :
            if self.index(x)==-1 :
                self.Names.append(x)
                for i in range(len(self.Values)):
                    self.Values[i].append(None)
        I=[]
        I2=[]
        for x in known_element :
            i=self.index(x)==-1 
            if i==-1 :
                print("Erreur d'attribution d'une valeur : l'un des éléments à reconnaitre n'est pas dans le fichier de base")
            else :
                I.append(i)
                I2.append(x)
        for i in range(len(self.Values)):
            flag=False
            for j in range(I):
                flag = flag or not(self.__is(self.Values[i][I[j]],known_element[I2[i]]))
                if flag :
                    break
            if not(flag) :
                for x in element_to_add :
                    self.Values[i][self.index(x)] = self.__to_num(self.Values[i][self.index(x)])+self.__to_num(element_to_add[x])
    def virtual_group_by_sum(self,known_element,to_sum_name,result_prefix="group_",add_count=False) :
        concerned_lines=[]
        I=[]
        I2=[]
        for x in known_element :
            i=self.index(x)==-1 
            if i==-1 :
                print("Erreur d'attribution d'une valeur : l'un des éléments à reconnaitre n'est pas dans le fichier de base")
            else :
                I.append(i)
                I2.append(x)
        for i in range(len(self.Values)):
            flag=False
            for j in range(I):
                flag = flag or not(self.__is(self.Values[i][I[j]],known_element[I2[i]]))
                if flag :
                    break
            if not(flag) :
                concerned_lines.append(i)
        if add_count :
            add_count_index = self.index(result_prefix+"count")
            if add_count_index==-1 :
                add_count_index=len(self.Names)
                self.Names.append(result_prefix+"count")
                for i in range(len(self.Values)) :
                    self.Values.append(None)
        count = len(concerned_lines)
        init_indexes=[]
        final_indexes=[]
        for j in range(len(to_sum_name)) :
            i = self.index(to_sum_name[j])
            if not(i==-1) :
                init_indexes.append(i)
                i2 = self.index(result_prefix+to_sum_name[j])
                if i2==-1 :
                    final_indexes.append(len(self.Names))
                    self.Names.append(result_prefix+to_sum_name[j])
                    for i3 in range(len(self.Values)) :
                        self.Values[i3].append(None)
                else :
                    final_indexes.append(i2)
        for i in range(count) :
            if add_count :
                self.Values[concerned_lines[i]][add_count_index]=count
        for j in range(len(init_indexes)) :
            value = 0
            for i in range(count) :
                value+=self.__to_num(self.Values[concerned_lines[i]][init_indexes[j]])
            for i in range(count) :
                self.Values[concerned_lines[i]][final_indexes[j]]=value
    def import_columns_from(self,path,links_main_to_imported,cols_to_import,where={},collapsed="summ",count_imported=False) :
        if count_imported :
            count_index = len(self.Names)
            name = "Nombre d éléments importés"
            if self.index(name)==-1 :
                self.Names.append(name)
            else :
                endname=1
                while not(self.index(name+" "+str(endname))==-1) :
                    endname+=1
                self.Names.append(name+" "+str(endname))
            for i in range(len(self.Values)) :
                self.Values[i].append(0)
        L_names=[]
        L_values=[]
        def index(name):
            for i in range(len(L_names)) :
                if self.__is(name,L_names[i]) :
                    return i
            return -1
        workbook = openpyxl.load_workbook(path)
        for col in range(workbook.active.max_column):
            L_names.append(workbook.active.cell(row=1,column=col+1).value)
        for rw in range(workbook.active.max_row-1):
            L_values.append([])
            for col in range(workbook.active.max_column):
                L_values[-1].append(workbook.active.cell(row=rw+2,column=col+1).value)
        workbook.close()
        for x in where :
            i = index(x)
            if not(i==-1) :
                n=len(L_values)
                for j in range(n) :
                    if not(self.__is(L_values[n-1-j][i],where[x])) :
                        del L_values[n-1-j]
        L_import_index = []
        L_imported_indexes=[]
        if isinstance(cols_to_import,list):
            for x in cols_to_import :
                i = index(x)
                if not(i==-1) :
                    L_import_index.append(i)
                    i = self.index(x)
                    if i==-1 :
                        L_imported_indexes.append(len(self.Names))
                        self.Names.append(x)
                        for i in range(len(self.Values)) :
                            self.Values[i].append(None)
                    else :
                        L_imported_indexes.append(i)
                        print("Attention dans l'import des colonnes vous impoprtez des colonnes existant déjà")
        elif isinstance(cols_to_import,dict) :
            for x in cols_to_import :
                i = index(x)
                if not(i==-1) :
                    L_import_index.append(i)
                    i = self.index(cols_to_import[x])
                    if i==-1 :
                        L_imported_indexes.append(len(self.Names))
                        self.Names.append(cols_to_import[x])
                        for i in range(len(self.Values)) :
                            self.Values[i].append(None)
                    else :
                        L_imported_indexes.append(i)
                        print("Attention dans l'import des colonnes vous impoprtez des colonnes existant déjà")
        links_base=[]
        links_imported=[]
        for x in  links_main_to_imported :
            i,j=self.index(x),index(links_main_to_imported[x])
            if not(i==-1) and not(j==-1) :
                links_base.append(i)
                links_imported.append(j)
        n = len(self.Values)
        for j in range(len(self.Values)) :
            print(str(int(100*(j+1)/n))+"%")
            for i in range(len(L_values)) :
                flag=True
                for k in range(len(links_base)):
                    flag = self.__is(self.Values[j][links_base[k]],L_values[i][links_imported[k]])
                    if not(flag) :
                        break
                if flag :
                    if(count_imported) :
                        self.Values[j][count_index] = self.__to_num(self.Values[j][count_index])+1
                    for k in range(len(L_import_index)) :
                        if collapsed=="summ" :
                            self.Values[j][L_imported_indexes[k]] = self.__to_num(self.Values[j][L_imported_indexes[k]])+self.__to_num(L_values[i][L_import_index[k]])
                        else :
                            self.Values[j][L_imported_indexes[k]]=L_values[i][L_import_index[k]]
    def save(self,path):
        if path.endswith(".xlsx") :
            workbook = openpyxl.Workbook()
            for i in range(len(self.Names)) :
                workbook.active.cell(row=1,column = i+1).value = self.Names[i]
                for j in range(len(self.Values)) :
                    workbook.active.cell(row=j+2,column = i+1).value = self.Values[j][i]
            workbook.save(path)
            workbook.close()
        elif path.endswith(".csv") :
            with open(path,"w",encoding="utf-8") as file :
                for i in range(len(self.Names)) :
                    file.write(self.Names[i].replace(";",",")+";")
                file.write("\n")
                for i in range(len(self.Values)) :
                    for j in range(len(self.Values[i])) :
                        file.write(str(self.Values[i][j]).replace(";",",")+";")
                    file.write("\n")
        else :
            print("Le type de fichier pour l'enregistrement est inconnu")
if __name__=="__main__" :
    import os
    path1 = os.path.join(os.getcwd(),"Compteurs.xlsx")
    path2 = os.path.join(os.getcwd(),"Lien_Compteur_-_Batiment.xlsx")
    path3 = os.path.join(os.getcwd(),"result.xlsx")
    XL = Excel(path1)
    XL.import_columns_from(path=path2,links_main_to_imported={"Code du Point de livraison":"Compteur"},cols_to_import=["Ratio"],where={},collapsed="summ",count_imported=False)
    XL.save(path3)
    
        
        
            
    
        