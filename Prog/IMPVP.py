# -*- coding: utf-8 -*-
"""
Created on Mon Feb  3 08:16:03 2025

@author: sacha.mailler
"""
from excel import Excel
import numpy as np
from sklearn.linear_model import LinearRegression
import plotly.graph_objects as go
class IPMVP :
    def __init__(self,path):
        self.L_names = []
        self.L_values = []
        self.to_load = {"ID":["ID"],"Année":["Année"],
                   "code bat RT":["code bat RT"],
                   "code site RT":["code site RT"],
                   "surface":["surface"],
                   "typologie 1":["typologie 1"],
                   "typologie 2": ["typologie 2"],}
        self.Consos = ["DJU","DJF","fioul","gaz","rcu","bois","froid","elec"]
        self.IPMVP_ends = ["_use","_fact_dju","_fact_djf"]
        self.IMPVP_vals = []
        self.read(path)
    def read(self,path):
        to_load = self.to_load
        for x in self.Consos :
            for i in range(12) :
                to_load[self.__month(i)+"_"+x]=[self.__month(i)+"_"+x]
        E = Excel()
        E.load(path,to_load,read_only=False)
        self.L_names,self.L_values = E.get()
        
    def load(self,IPMVP_path):
        col = self.__index("typologie 1")
        if col>=0 :
            typo = self.L_values[0][col]
        to_load = {"Mois":["Mois"]}
        for x in self.IPMVP_ends :
            to_load[typo+x]=[typo+x]
        E = Excel()
        E.load(IPMVP_path,to_load,read_only=False)
        t,self.IMPVP_vals = E.get()
        del t
        del E
    def __index(self,name) :
        to_replace = "äâàABCDêéèëEFGHIîïJKLMNOôöPQRSTUûüùVWXYZ"
        replacer   = "aaaabcdeeeeefghiiijklmnooopqrstuuuuvwxyz"
        def simplify(word,simplified=False):
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
        for i in range(len(self.L_names)) :
            if simplify(self.L_names[i],False)==simplify(name,False) :
                return i
        for i in range(len(self.L_names)) :
            if simplify(self.L_names[i],True)==simplify(name,True) :
                return i
        return -1
    def get_regression(self):
        def to_float(st):
            if st==None :
                return 0
            if st=="" :
                return 0
            return float(st)
        year_index = self.__index("Année")
        years=[]
        for i in range(len(self.L_values)):
            years.append(int(self.L_values[i][year_index]))
        years.sort()
        y,X=[],[]
        for i in range(len(self.L_values)):
            if not(int(self.L_values[i][year_index])==years[-1]) :
                for j in range(12) :
                    use,DJU,DJF = self.IMPVP_vals[j][1],0,0
                    if self.__index(self.__month(j)+"_DJU")>=0 :
                        DJU = to_float(self.IMPVP_vals[j][2])*to_float(self.L_values[i][self.__index(self.__month(j)+"_DJU")])
                    if self.__index(self.__month(j)+"_DJF")>=0 :
                        DJF = to_float(self.IMPVP_vals[j][2])*to_float(self.L_values[i][self.__index(self.__month(j)+"_DJF")])
                    conso = 0
                    for x in self.Consos :
                        if self.__index(self.__month(j)+"_"+x)>=0 :
                            conso+=to_float(self.L_values[i][self.__index(self.__month(j)+"_"+x)])
                    y.append(conso)
                    X.append([use,DJU,DJF])
        reg = LinearRegression().fit(X, y)
        tal,usa,cha,fro=[],[],[],[]
        y,yth=[],[]
        for i in range(len(self.L_values)):
            if int(self.L_values[i][year_index])==years[-1] :
                break
        for j in range(12) :
            use,DJU,DJF = self.IMPVP_vals[j][1],0,0
            if self.__index(self.__month(j)+"_DJU")>=0 :
                DJU = to_float(self.IMPVP_vals[j][2])*to_float(self.L_values[i][self.__index(self.__month(j)+"_DJU")])
            if self.__index(self.__month(j)+"_DJF")>=0 :
                DJF = to_float(self.IMPVP_vals[j][2])*to_float(self.L_values[i][self.__index(self.__month(j)+"_DJF")])
            conso = 0
            for x in self.Consos :
                if self.__index(self.__month(j)+"_"+x)>=0 :
                    conso+=to_float(self.L_values[i][self.__index(self.__month(j)+"_"+x)])
            tal.append(reg.predict(np.array([[0,0,0]]))[0])
            usa.append(reg.predict(np.array([[use,0,DJF]]))[0]-tal[-1])
            cha.append(reg.predict(np.array([[0,DJU,0]]))[0]-tal[-1])
            fro.append(reg.predict(np.array([[0,0,DJF]]))[0]-tal[-1])
            y.append(conso)
            yth.append(reg.predict(np.array([[use,DJU,DJF]]))[0])
        tot=0
        tot_th=0
        for i in range(len(y)):
            tot+=y[i]
            tot_th+=yth[i]
        print(str((tot-tot_th)/tot_th*100)+" %")
        x=["01","02","03","04","05","06","07","08","09","10","11","12"]
        fig = go.Figure(go.Bar(x=x, y=tal, name='Talon'))
        fig.add_trace(go.Bar(x=x, y=usa, name='Usage'))
        fig.add_trace(go.Bar(x=x, y=cha, name='Chaud'))
        fig.add_trace(go.Bar(x=x, y=fro, name='Froid'))
        fig.update_layout(barmode='stack', xaxis={'categoryorder':'category ascending'})
        #fig.add_trace(go.Scatter(x=x,y=y,name="Consommation réelle"))
        fig.show()
    def __month(self,month_num) :
        if month_num>=12 :
            return "00"
        else :
            if len(str(int(month_num)))==2 :
                return str(int(month_num))
            return "0"+str(int(month_num))
if __name__=="__main__":
    import os
    ipmvppath="C:\\Users\\sacha.mailler\\Desktop\\GIT\\OSFI\\Datas\\IPMVP_fact_typ.xlsx"
    path="C:\\Users\\sacha.mailler\\Desktop\\GIT\\OSFI\\Datas\\__work__"
    path = os.path.join(path,"178869_1000007090_Consommations_mensualisees_des_equipements.xlsx")
    a=IPMVP(path)
    a.load(IPMVP_path=ipmvppath)
    a.get_regression()