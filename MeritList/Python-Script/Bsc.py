import pandas as pd
import xlrd
import matplotlib.pyplot as pt
import matplotlib.pyplot as pt2
import numpy as np
import sys
import time

graphNames = ""
flag=1
def extractMarks(records,valGP,valHSC):
    i=0
    while i!=len(records):
        valGP.append(records[i]["12th GP Marks"])
        valHSC.append(records[i]["% Marks @ hsc"])
        i+=1
    return;

def plotGraph(high,low,title):
    global graphNames
    global flag
    N=6
    ind = np.arange(N)
    width = 0.27
    ax = pt.subplot(111)
    ax.bar(ind, high,width,color='g')
    ax.bar(ind+width, low,width,color='r')
    ax.set_xticks(ind+.14)
    ax.set_xticklabels(('OPEN','OBC','SBC','SB','ST','NT'))
    ax.legend(('Highest','Lowest'))
    ax.set_ylabel('Marks')
    ax.set_title(title)
    pos=-0.25
    for i in high:
        ax.text(pos, i + 1, str(round(i,2)), color='blue', fontsize=6)
        pos+=1
    pos=-0.05+width
    for i in low:
        ax.text(pos, i + 1, str(round(i,2)), color='indigo', fontsize=6)
        pos+=1
    name = title+".png"
    pt.show(block=False)
    time.sleep(1)
    pt.savefig("../Resource/Graphs/"+title)
    pt.close()
    if flag==1:
        graphNames = name
        flag=2
    else:
        graphNames = graphNames+","+name
    return;



def checkLength(Marks):
    l=len(Marks)
    if len(Marks)==0:
            Marks.append(0)
    return;

if __name__ == '__main__':

    fn=sys.argv[1]
    fn2 = "../Resource/"+fn
    book = xlrd.open_workbook(fn2)
    sheets=book.sheet_names()#sheetnames
    data={}
    finalData={}

    for x in sheets:
        if x!="MASTER":
            data[x]=pd.read_excel(fn2,sheet_name=x)

    for y in data:
        df=[]
        i=0
        for x in data[y]["Remarks"]:
            if x=="SELECTED":
                df.append(data[y].iloc[i])
            i+=1
        finalData[y]=df


    openCaste={}
    obcCaste={}
    scCaste={}
    stCaste={}
    ntCaste={}
    sbcCaste={}

    for x in finalData:
        i=0
        openCaste[x]=[]
        obcCaste[x]=[]
        scCaste[x]=[]
        stCaste[x]=[]
        ntCaste[x]=[]
        sbcCaste[x]=[]
        while i!=len(finalData[x]):
            if finalData[x][i]["Category"]=="OPEN":
                openCaste[x].append(finalData[x][i])
            elif finalData[x][i]["Category"]=="OBC":
                obcCaste[x].append(finalData[x][i])
            elif finalData[x][i]["Category"]=="SC":
                scCaste[x].append(finalData[x][i])
            elif finalData[x][i]["Category"]=="ST":
                stCaste[x].append(finalData[x][i])
            elif finalData[x][i]["Category"]=="NT":
                ntCaste[x].append(finalData[x][i])
            elif finalData[x][i]["Category"]=="SBC":
                sbcCaste[x].append(finalData[x][i])
            i+=1


    for x in finalData:
        openCasteValGP=[]
        obcCasteValGP=[]
        scCasteValGP=[]
        stCasteValGP=[]
        ntCasteValGP=[]
        sbcCasteValGP=[]

        openCasteValHSC=[]
        obcCasteValHSC=[]
        scCasteValHSC=[]
        stCasteValHSC=[]
        ntCasteValHSC=[]
        sbcCasteValHSC=[]

        extractMarks(openCaste[x],openCasteValGP,openCasteValHSC)
        extractMarks(obcCaste[x],obcCasteValGP,obcCasteValHSC)
        extractMarks(sbcCaste[x],sbcCasteValGP,sbcCasteValHSC)
        extractMarks(scCaste[x],scCasteValGP,scCasteValHSC)
        extractMarks(stCaste[x],stCasteValGP,stCasteValHSC)
        extractMarks(ntCaste[x],ntCasteValGP,ntCasteValHSC)

        checkLength(openCasteValGP)
        checkLength(openCasteValHSC)
        checkLength(obcCasteValGP)
        checkLength(obcCasteValHSC)
        checkLength(sbcCasteValGP)
        checkLength(sbcCasteValHSC)
        checkLength(scCasteValGP)
        checkLength(scCasteValHSC)
        checkLength(stCasteValGP)
        checkLength(stCasteValHSC)
        checkLength(ntCasteValGP)
        checkLength(ntCasteValHSC)

        highMeritGP=[max(openCasteValGP),max(obcCasteValGP),max(sbcCasteValGP),max(scCasteValGP),max(stCasteValGP),max(ntCasteValGP)]
        highMeritHSC=[max(openCasteValHSC),max(obcCasteValHSC),max(sbcCasteValHSC),max(scCasteValHSC),max(stCasteValHSC),max(ntCasteValHSC)]

        lowMeritGP=[min(openCasteValGP),min(obcCasteValGP),min(sbcCasteValGP),min(scCasteValGP),min(stCasteValGP),min(ntCasteValGP)]
        lowMeritHSC=[min(openCasteValHSC),min(obcCasteValHSC),min(sbcCasteValHSC),min(scCasteValHSC),min(stCasteValHSC),min(ntCasteValHSC)]

        plotGraph(highMeritGP,lowMeritGP,x+" MERIT LIST BSC (Group Percentage List)")
        plotGraph(highMeritHSC,lowMeritHSC,x+" MERIT LIST BSC (HSC Percentage List)")
    print(graphNames)
