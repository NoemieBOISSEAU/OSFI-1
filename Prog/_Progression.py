# -*- coding: utf-8 -*-
"""
Created on Thu Nov  2 11:02:55 2023

@author: sacha.mailler
"""
import sys
import os
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
    def __init__(self,name,path):
        print(return_line(name,n=52))
        sys.stdout.write(f"\r{' '*52}\r")
        sys.stdout.flush()
        self.progression=0
        self.position=0
        self.max,self.min=100,0
        self.path=path
        with open(os.path.join(self.path,"progress_comment.txt"),"w",encoding="utf-8") as file :
            file.write(name)
        with open(os.path.join(self.path,"progress_value.txt"),"w",encoding="utf-8") as file :
            file.write("0")
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
        else :
            with open(os.path.join(self.path,"progress_value.txt"),"w",encoding="utf-8") as file :
                file.write(str(int(progression)))
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
