import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from win32com.client import Dispatch
s = Dispatch("SAPI.SpVoice")
print("write EXACT file name with format as saved in your computer and make sure it is present in the same directory in which programme file is present")
s.Speak("write EXACT file name with format as saved in your computer and make sure it is present in the same directory in which programe file is present")
file = input()#input file name 
mydata =  pd.read_csv(file)
print("enter the X label")
s.Speak("enter the X label")
plt.xlabel(input())
print("enter the Y label")
s.Speak("enter the y label")
plt.ylabel(input())
#plt.plot(mydata.Year, mydata.Australia, )
#plt.plot(mydata.Year, mydata.Canada, )
print("enter the graph title")
s.Speak("enter the graph title")
plt.title(input())
print("enter the EXACT COLOUMN NAME you want to choose at Y AXIS")
s.Speak("enter the EXACT COLOUMN NAME you want to choose at Y AXIS")
t= input()
#so that i don't have to input all the colums of the data it will iterate its coloums
def linegraph():
   for i in mydata:
        if i != t:
            print(i)
            plt.plot(mydata[t],mydata[i],label = i,marker='.')
            plt.legend(loc = (1.1,0))#to put legend outside the graph
def scatter():        
    for i in mydata:
        if i != t:
           # print(i)
            plt.scatter(mydata[t], mydata[i],label = i )
            plt.legend(loc = (1.1,0))
def histogram():
    for i in mydata:
        if i != t:
            # print(i)
            plt.hist(mydata[i],label = i )
            plt.legend(loc = (1.1,0))
def piegraph():
    val = []
    print("enter the coloumn you want to remove from pie graph note, it will only be removed for pie graph")
    s.Speak("enter the coloumn you want to remove from pie graph note, it will only be removed for pie graph")
    f=input()
    for i in mydata.columns:
        if i != f:
            z = mydata[i].tolist()
            m = np.mean(z)
            val.append(m)
    print(val)
    t= mydata.columns.tolist()
    t.remove(f)
    print(t)
    plt.pie(val,labels = t,autopct = '%0.3f%%',shadow = True,radius = 1.3)
    plt.legend(loc = (1.3,0))
def boxplot():
    val = []
    print("enter the coloumn you want to remove or just write no then enter")
    s.Speak("enter the coloumn you want to remove or just write no then enter")
    f=input()
    for i in mydata.columns:
        if i != f:
            z = mydata[i].tolist()
            m = np.mean(z)
            val.append(m)
    print(val)
    print(np.mean(val))
    t= mydata.columns.tolist()
    if f!="no":
        t.remove(f)
    print(t)
    plt.figure()
    plt.boxplot(val)
    box = []
    label = []
    for j in mydata.columns:
        if j != f:
            z = mydata[j].tolist()#convert all the row entries into a list of that coloumn 
            box.append(z)
            label.append(j)
            plt.figure()
            plt.boxplot(z,labels = [j])
    plt.figure()
    print("enter the title of main graph")
    s.Speak("enter the title of main graph")
    plt.title(input())
    plt.boxplot(box,labels = label)
try:
    plt.figure(1)
    scatter()
except:
    print("some error occured during processing of scatter plot follow the commands given by the bot ")
    s.Speak("some error occured during processing of your scatter plot follow the commands given by the bot ")
try:
    plt.figure(2)
    linegraph()
except:
    print("some error occured during processing of line graph follow the commands given by the bot ")
    s.Speak("some error occured during processing of line graph follow the commands given by the bot ")
try:
    plt.figure(3)
    histogram()
except:
    print("some error occured during processing of histogram graph follow the commands given by the bot ")
    s.Speak("some error occured during processing of histogram graph follow the commands given by the bot ")
try:
    plt.figure(4)
    piegraph()
except:
    print("some error occured during processing of pie graph follow the commands given by the bot ")
    s.Speak("some error occured during processing of pie graph follow the commands given by the bot ")
try:
    plt.figure(5)
    boxplot()
except:
    print("some error occured during processing of box plot follow the commands given by the bot ")
    s.Speak("some error occured during processing of box plot follow the commands given by the bot ")


