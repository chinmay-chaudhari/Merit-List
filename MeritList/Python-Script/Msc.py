import pandas as pd
import xlrd
import matplotlib.pyplot as pt
import numpy as np
import time
import sys

graphNames = ""
flag=1

def extractMarks(records,val12th,valPI,valEntrance,valTotal):
    i=0
    while i!=len(records):
        val12th.append(records[i]["50% Vatage @12 Marks"])
        valPI.append(records[i]["P.I. Marks"])
        valEntrance.append(records[i]["Entrance"])
        valTotal.append(records[i]["Total"])
        i+=1
    return;

def plotGraph(high,low,title):
    global graphNames
    global flag
    N=8
    ind = np.arange(N)
    width = 0.27
    ax = pt.subplot(111)
    ax.bar(ind, high,width,color='g',)
    ax.bar(ind+width, low,width,color='r')
    ax.set_xticks(ind+width)
    ax.set_xticklabels(('OPEN','OBC','SBC','SB','ST','NT','NT2','VJ'))
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
    vjCaste={}
    nt2Caste={}

    for x in finalData:
        i=0
        openCaste[x]=[]
        obcCaste[x]=[]
        scCaste[x]=[]
        stCaste[x]=[]
        ntCaste[x]=[]
        sbcCaste[x]=[]
        vjCaste[x]=[]
        nt2Caste[x]=[]
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
            elif finalData[x][i]["Category"]=="VJ":
                vjCaste[x].append(finalData[x][i])
            elif finalData[x][i]["Category"]=="NT2":
                nt2Caste[x].append(finalData[x][i])
            i+=1


    for x in finalData:
        openCasteVal12th=[]
        obcCasteVal12th=[]
        scCasteVal12th=[]
        stCasteVal12th=[]
        ntCasteVal12th=[]
        sbcCasteVal12th=[]
        nt2CasteVal12th=[]
        vjCasteVal12th=[]

        openCasteValPI=[]
        obcCasteValPI=[]
        scCasteValPI=[]
        stCasteValPI=[]
        ntCasteValPI=[]
        sbcCasteValPI=[]
        vjCasteValPI=[]
        nt2CasteValPI=[]

        openCasteValEntrance=[]
        obcCasteValEntrance=[]
        scCasteValEntrance=[]
        stCasteValEntrance=[]
        ntCasteValEntrance=[]
        sbcCasteValEntrance=[]
        vjCasteValEntrance=[]
        nt2CasteValEntrance=[]

        openCasteValTotal=[]
        obcCasteValTotal=[]
        scCasteValTotal=[]
        stCasteValTotal=[]
        ntCasteValTotal=[]
        sbcCasteValTotal=[]
        vjCasteValTotal=[]
        nt2CasteValTotal=[]

        extractMarks(openCaste[x],openCasteVal12th,openCasteValPI,openCasteValEntrance,openCasteValTotal)
        extractMarks(obcCaste[x],obcCasteVal12th,obcCasteValPI,obcCasteValEntrance,obcCasteValTotal)
        extractMarks(sbcCaste[x],sbcCasteVal12th,sbcCasteValPI,sbcCasteValEntrance,sbcCasteValTotal)
        extractMarks(scCaste[x],scCasteVal12th,scCasteValPI,scCasteValEntrance,scCasteValTotal)
        extractMarks(stCaste[x],stCasteVal12th,stCasteValPI,stCasteValEntrance,stCasteValTotal)
        extractMarks(ntCaste[x],ntCasteVal12th,ntCasteValPI,ntCasteValEntrance,ntCasteValTotal)
        extractMarks(nt2Caste[x],nt2CasteVal12th,nt2CasteValPI,nt2CasteValEntrance,nt2CasteValTotal)
        extractMarks(vjCaste[x],vjCasteVal12th,vjCasteValPI,vjCasteValEntrance,vjCasteValTotal)


        checkLength(openCasteVal12th)
        checkLength(openCasteValPI)
        checkLength(openCasteValEntrance)
        checkLength(openCasteValTotal)
        checkLength(obcCasteVal12th)
        checkLength(obcCasteValPI)
        checkLength(obcCasteValEntrance)
        checkLength(obcCasteValTotal)
        checkLength(sbcCasteVal12th)
        checkLength(sbcCasteValPI)
        checkLength(sbcCasteValEntrance)
        checkLength(sbcCasteValTotal)
        checkLength(scCasteVal12th)
        checkLength(scCasteValPI)
        checkLength(scCasteValEntrance)
        checkLength(scCasteValTotal)
        checkLength(stCasteVal12th)
        checkLength(stCasteValPI)
        checkLength(stCasteValEntrance)
        checkLength(stCasteValTotal)
        checkLength(ntCasteVal12th)
        checkLength(ntCasteValPI)
        checkLength(ntCasteValEntrance)
        checkLength(ntCasteValTotal)
        checkLength(nt2CasteVal12th)
        checkLength(nt2CasteValPI)
        checkLength(nt2CasteValEntrance)
        checkLength(nt2CasteValTotal)
        checkLength(vjCasteVal12th)
        checkLength(vjCasteValPI)
        checkLength(vjCasteValEntrance)
        checkLength(vjCasteValTotal)

        highMerit12th=[max(openCasteVal12th),max(obcCasteVal12th),max(sbcCasteVal12th),max(scCasteVal12th),max(stCasteVal12th),max(ntCasteVal12th),max(nt2CasteVal12th),max(vjCasteVal12th)]
        highMeritPI=[max(openCasteValPI),max(obcCasteValPI),max(sbcCasteValPI),max(scCasteValPI),max(stCasteValPI),max(ntCasteValPI),max(nt2CasteValPI),max(vjCasteValPI)]
        highMeritEntrance=[max(openCasteValEntrance),max(obcCasteValEntrance),max(sbcCasteValEntrance),max(scCasteValEntrance),max(stCasteValEntrance),max(ntCasteValEntrance),max(nt2CasteValEntrance),max(vjCasteValEntrance)]
        highMeritTotal=[max(openCasteValTotal),max(obcCasteValTotal),max(sbcCasteValTotal),max(scCasteValTotal),max(stCasteValTotal),max(ntCasteValTotal),max(nt2CasteValTotal),max(vjCasteValTotal)]

        lowMerit12th=[min(openCasteVal12th),min(obcCasteVal12th),min(sbcCasteVal12th),min(scCasteVal12th),min(stCasteVal12th),min(ntCasteVal12th),min(nt2CasteVal12th),min(vjCasteVal12th)]
        lowMeritPI=[min(openCasteValPI),min(obcCasteValPI),min(sbcCasteValPI),min(scCasteValPI),min(stCasteValPI),min(ntCasteValPI),min(nt2CasteValPI),min(vjCasteValPI)]
        lowMeritEntrance=[min(openCasteValEntrance),min(obcCasteValEntrance),min(sbcCasteValEntrance),min(scCasteValEntrance),min(stCasteValEntrance),min(ntCasteValEntrance),min(nt2CasteValEntrance),min(vjCasteValEntrance)]
        lowMeritTotal=[min(openCasteValTotal),min(obcCasteValTotal),min(sbcCasteValTotal),min(scCasteValTotal),min(stCasteValTotal),min(ntCasteValTotal),min(nt2CasteValTotal),min(vjCasteValTotal)]

        plotGraph(highMerit12th,lowMerit12th,x+" MERIT LIST MSC (12th 50 Percentage)")
        plotGraph(highMeritPI,lowMeritPI,x+" MERIT LIST MSC (Personal Interview)")
        plotGraph(highMeritEntrance,lowMeritEntrance,x+" MERIT LIST MSC (Entrance)")
        plotGraph(highMeritTotal,lowMeritTotal,x+" MERIT LIST MSC (Total)")
    print(graphNames)
