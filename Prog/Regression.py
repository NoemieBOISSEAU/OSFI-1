# -*- coding: utf-8 -*-
"""
Created on Wed Nov 29 14:29:58 2023

@author: sacha.mailler
"""
X=[[0,0,1],[1,0,1],[2,1,1],[3,1,1]]
Y=[0+0+1+0.0001,1+0+1-0.001,2+2+1,3+2+1]
# import statsmodels.api as sm
import plotly.graph_objects as go
import os
def get_regression(Consommation,Surface,DJU,DJF,alpha=0.05) :
    if DJF==[] :
        if DJU==[]:
            import scipy
            print(Surface)
            result=scipy.stats.linregress(Surface,Consommation)
            # regression=sm.OLS(Consommation,[[1,Surface[i]] for i in range(len(Surface))]).fit()
            # params=regression.params
            # Iconf=regression.conf_int(alpha=alpha)
            return {"A0" : result.intercept,"A1":result.slope,"A2":0,"A3":0,
                "A0inf":result.intercept-result.intercept_stderr,"A1inf":result.slope-result.stderr,"A2inf":0,"A3inf":0,
                "A0sup":result.intercept+result.intercept_stderr,"A1sup":result.slope+result.stderr,"A2sup":0,"A3sup":0,"r":result.rvalue}
    #     else :
    #         import statsmodels.api as sm
    #         regression=sm.OLS(Consommation,[[1,Surface[i],DJU[i]] for i in range(len(Surface))]).fit()
    #         params=regression.params
    #         Iconf=regression.conf_int(alpha=alpha)
    #         return {"A0" : params[0],"A1":params[1],"A2":params[2],"A3":0,
    #             "A0inf":Iconf[0][0],"A1inf":Iconf[1][0],"A2inf":Iconf[2][0],"A3inf":0,
    #             "A0sup":Iconf[0][1],"A1sup":Iconf[1][1],"A2sup":Iconf[2][1],"A3sup":0}
    # else :
    #     import statsmodels.api as sm
    #     regression=sm.OLS(Consommation,[[1,Surface[i],DJU[i],DJF[i]] for i in range(len(Surface))]).fit()
    #     params=regression.params
    #     Iconf=regression.conf_int(alpha=alpha)
    #     return {"A0" : params[0],"A1":params[1],"A2":params[2],"A3":params[3],
    #         "A0inf":Iconf[0][0],"A1inf":Iconf[1][0],"A2inf":Iconf[2][0],"A3inf":Iconf[3][0],
    #         "A0sup":Iconf[0][1],"A1sup":Iconf[1][1],"A2sup":Iconf[2][1],"A3sup":Iconf[3][1]}
#En supposant que A3 est très proche de 0, la représentation tient compte que de A0, A1, et A2
def draw_acceptable_html(Dict,LP,LDatas=[],surf_min=1,surf_max=1000,DJU_min=800,DJU_max=1800,n_subdiv=10):
    hoover_template="<br>Consommation : %{z} <br>Surface : %{x} <br>DJU : %{y} <br>"
    if len(LDatas)==len(LP) :
        T=LDatas[0]
        Datas=[]
        for i in range(len(LDatas)) :
            Datas.append([])
            j=0
            for key in T :
                if i==0 :
                    hoover_template+=(key+" : "+'%{customdata['+str(j)+']} <br>')
                    j+=1
                Datas[-1].append(LDatas[i][key])
    elif len(LDatas)==(len(LP)+1) :#la première liste c'est les noms
        T=LDatas[0]
        Datas=[]
        j=0
        for key in T :
            hoover_template+=(key+" : "+'%{customdata['+str(j)+']} <br>')
        for i in range(len(LDatas)-1):
            Datas.append([])
            for val in LDatas[i+1] :
                Datas[-1].append(val)
    dsurf=(surf_max-surf_min)/n_subdiv
    dDJU=(DJU_max-DJU_min)/n_subdiv
    surf=[surf_min]
    DJU=[DJU_min]
    for i in range(n_subdiv) :
        surf.append(surf_min+dsurf*(i+1))
        DJU.append(DJU_min+dDJU*(i+1))
    conso_inf,conso_sup=[],[]
    for j in range(len(DJU)):
        conso_inf.append([])
        conso_sup.append([])
        for i in range(len(surf)) :
            conso_inf[-1].append(Dict["A0inf"]+Dict["A1inf"]*surf[i]+Dict["A2inf"]*DJU[j])
            conso_sup[-1].append(Dict["A0sup"]+Dict["A1sup"]*surf[i]+Dict["A2sup"]*DJU[j])
    colorscale = [[0, 'red'], [1, 'red']]
    fig=go.Figure(data=[go.Surface(z=conso_inf, x=surf, y=DJU, colorscale=colorscale,showscale=False),go.Surface(z=conso_sup, x=surf, y=DJU,colorscale=colorscale,showscale=False)])
    LGP=[]
    LMP=[]
    LBP=[]
    for i in range(len(LP)) :
        if LP[i][2]> Dict["A0sup"]+Dict["A1sup"]*LP[i][0]+Dict["A2sup"]*LP[i][1] :
            LBP.append(i)
        elif LP[i][2]< Dict["A0inf"]+Dict["A1inf"]*LP[i][0]+Dict["A2inf"]*LP[i][1]:
            LGP.append(i)
        else :
            LMP.append(i)
    if len(Datas)==len(LP) :
        fig.add_scatter3d(x=[LP[i][0] for i in LGP],y=[LP[i][1] for i in LGP],z=[LP[i][2] for i in LGP],customdata=[Datas[i] for i in LGP],mode='markers',marker=dict(color="red",opacity=0.5),name="Consomation trop faible",hovertemplate=hoover_template)
        fig.add_scatter3d(x=[LP[i][0] for i in LMP],y=[LP[i][1] for i in LMP],z=[LP[i][2] for i in LMP],customdata=[Datas[i] for i in LMP],mode='markers',marker=dict(color="green",opacity=0.5),name="Consomation cohérente",hovertemplate=hoover_template)
        fig.add_scatter3d(x=[LP[i][0] for i in LBP],y=[LP[i][1] for i in LBP],z=[LP[i][2] for i in LBP],customdata=[Datas[i] for i in LBP],mode='markers',marker=dict(color="red",opacity=0.5),name="Consomation trop élevée",hovertemplate=hoover_template)
    else :
        fig.add_scatter3d(x=[LP[i][0] for i in LGP],y=[LP[i][1] for i in LGP],z=[LP[i][2] for i in LGP],mode='markers',marker=dict(color="red",opacity=0.5),name="Consomation trop faible",hovertemplate=hoover_template)
        fig.add_scatter3d(x=[LP[i][0] for i in LMP],y=[LP[i][1] for i in LMP],z=[LP[i][2] for i in LMP],mode='markers',marker=dict(color="green",opacity=0.5),name="Consomation cohérente",hovertemplate=hoover_template)
        fig.add_scatter3d(x=[LP[i][0] for i in LBP],y=[LP[i][1] for i in LBP],z=[LP[i][2] for i in LBP],mode='markers',marker=dict(color="red",opacity=0.5),name="Consomation trop élevée",hovertemplate=hoover_template)
    fig.write_html(os.path.join(os.path.dirname(__file__),"file.html"))

if __name__=="__main__" :
    # Surface,Dju,Djf,Conso,Annee,Type="Surface totale","Degrés jours unifiés","Degrés jours unifiés",[["Consommation de gaz (kWh PCS)","Consommation réseau de chaleur (kWh)","Consommation fioul (kWh PCS)","Consommation de granulés (kWh)"],"Consommation froid urbain (kWh)","Consommation d'électricité (kWh)"],"Année","Typologie"
    # ANNEE,SURFACE,DJU,DJF,CCHAUD,CFROID,CELEC,TYPE=[],[],[],[],[],[],[],[]
    # erreur_surface,erreur_dj,erreur_annee,erreur_conso=0,0,0,0
    import os
    import random
    import json
    import pandas
    # file = os.path.join(os.getcwd(),"Consommations annuelles des équipements.csv")
    # data=pandas.read_csv(file,sep=";",encoding="ansi")
    # for index in data.index :
    #     surface=data[index,Surface]
    #     if isinstance(surface,float) and surface>0 :
    #         annee=data[index,Annee]
    #         if isinstance(annee,float) and annee>=1995 :
    #             dju=data[index,Dju]
    #             if isinstance(dju,float) and dju>0 :
    #                 djf=data[index,Djf]
    #                 if isinstance(djf,float) and djf>0 :
    #                     flag=True
    #                     cchaud=0
    #                     for i in range(len(Conso[0])):
    #                         temp=data[index,Conso[0][i]]
    #                         if isinstance(temp,float) :
    #                             if temp<0 :
    #                                 flag=False
    #                                 break
    #                             else :
    #                                 cchaud+=temp
    #                     if flag :
    #                         temp=data[index,Conso[1]]
    #                         cfroid=0
    #                         if isinstance(temp,float) :
    #                             if temp<0 :
    #                                 flag=False
    #                             else :
    #                                 cfroid=temp
    #                         if flag :
    #                             temp=data[index,Conso[2]]
    #                             celec=0
    #                             if isinstance(temp,float) :
    #                                 if temp<0 :
    #                                     flag=False
    #                                 else :
    #                                     cfroid=temp
    #                             if flag :
    #                                 SURFACE.append(surface)
    #                                 DJU.append(DJU)
    #                                 DJF.append(DJF)
    #                                 CCHAUD.append(cchaud)
    #                                 CFROID.append(cfroid)
    #                                 CELEC.append(celec)
    #                                 TYPE.append(data[index,Type])
    #                                 #Tout semble correct
    #                             else :
    #                                 erreur_conso+=1
    #                         else :
    #                             #Consommation de froid négative
    #                             erreur_conso+=1
    #                     else :
    #                         #Consommation de chaud négative
    #                         erreur_conso+=1
    #                 else :
    #                     #djf négatifs
    #                     erreur_dj+=1
    #             else :
    #                 #dju négatif
    #                 erreur_dj+=1
    #         else :
    #             #annee incorrecte
    #             erreur_annee+=1
                
    #     else :
    #         #surface incorrecte
    #         erreur_surface+=1
    # #Extraction des erreures :
    #     #Erreures d'année :
    #     #Erreures de DJU-DJF :
    #     #Erreures de Surfaces :
    #     #Erreures de consommation :
    # # Séparer chauffage connu de chauffage inconnus :
    # CSurface,CDJU,CDJF,CConso_chauffage,CConso_elec,CConso_froid=[],[],[],[],[],[]
    # #Séparer par catégorie
    # ASurface,ADJU,ADJF,AConso_chauffage,AConso_elec,AConso_froid=[],[],[],[],[],[]
    
    path=os.path.join(os.path.dirname(__file__),)
    data={"premier argument" : "1","deuxieme argument" : "2"}
    S=[random.randint(1,10000) for i in range(1000)]
    Data=[data for i in range(len(S))]
    DJU=[1000+random.randint(-500,500) for i in range(len(S))]
    Consommation=[200+random.randint(0,100000000)+S[i]*1500+DJU[i]*15 for i in range(len(S))]
    with open("sample.json", "w") as outfile: 
        json.dump(get_regression(Consommation,S,DJU,[]), outfile)
    # dataframe=pandas.DataFrame(data={"Consommation":Consommation,"Surface":S,"DJU":DJU})
    # dataframe.to_csv('sample.csv',sep=';', encoding='ansi')
    draw_acceptable_html((get_regression(Consommation,S,DJU,[])),[[S[i],DJU[i],Consommation[i]] for i in range(len(S))],LDatas=Data,surf_min=min(S),surf_max=max(S),DJU_min=min(DJU),DJU_max=max(DJU))
    