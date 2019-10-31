#PANDAGANTT python PRINCE2 plan generator
#Charles Fox, University of Lincoln, 2019

import pandas as pd
import numpy as np
import subprocess, pdb
import datetime
from dateutil.relativedelta import *



import GanttChart

pd.set_option('display.max_columns', 20)

############LATEX FUNTIONS##########

def latexHeader():
    docTitle = "ARWAC second stage plan"
    dateStr = ""+datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    s='\\documentclass[english]{article}\n'
    s+='\\usepackage[T1]{fontenc}\n'
    s+='\\usepackage[latin9]{inputenc}\n'
    s+='\\usepackage{babel}\n'
    s+='\\usepackage{graphicx}\n'
    s+='\\begin{document}\n'
    s+="\\title{"+docTitle+"}\\author{ARWAC Consortium}\\date{"+dateStr+"}\\maketitle\n"
    s+="\\begin{center}\n\n"
    s+= "\\includegraphics[width=6cm]{arwac.png}\n\n"
    s+= "\\includegraphics[width=6cm]{lincoln.jpeg}\n\n"
    s+= "\\includegraphics[width=6cm]{innovateuk.png}\n\n"
    s+="\\end{center}\n\n"
    s+= "\\newpage\n\n"
    s+= getIntroText()
    return s

def latexFooter():
    s= '\end{document}'
    return s

def makeLatexTable(M, colNames=[], rowNames=[]):
    (nrows, ncols) = M.shape

    if rowNames==[]:
        colDescriptor = "|r"*ncols + "|"	
    else:
        colDescriptor = "|r"*(1+ncols) + "|"	

    s = "\\begin{tabular}{%s}"%colDescriptor
    if not colNames==[]:
        s+= "\\hline\n "
        srow=""
	
        if not rowNames==[]:	
            srow += str("X") + " & "

        for col in range(0,ncols):
            srow += str(colNames[col]) + " & "
        srow=srow[0:-2]  #remove trailing &
        srow+=" \\tabularnewline"
        s+=srow+"\n"
        s+="\hline\n"
    for row in range(0, nrows):
        s+= "\\hline\n "
        srow = ""
        if not rowNames==[]:
            srow += str(rowNames[row]) + " & "
        for col in range(0,ncols):
            srow +=   '{:,.2f}'.format(M[row,col]) + " & "
        srow=srow[0:-2]  #remove trailing &
        srow+="\\tabularnewline\n"
        s+=srow
    s+= "\\hline "
    s+= "\\end{tabular}"
    return s

##############PRINCE2 FUNCTIONS############

def getIntroText():
    s="\\section{Project overview}\n\n"
    s+=df_Project['Description'][0]   
    s+= "\\newpage\n\n"

    s+="\\section{Work package overview}\n\n"
    #pdb.set_trace()
    for i in range(0, df_WorkPackage.shape[0]):
        s+="\\subsection*{WP"+str(df_WorkPackage.iloc[i]['id'])+": "+df_WorkPackage.iloc[i]['Name']+ "}\n\n"
        s+="Objectives: " + df_WorkPackage.iloc[i]['Objectives']+"\n\n"
        s+="Description: " +  df_WorkPackage.iloc[i]['Description']+"\n\n"
    s+= "\\newpage\n\n"
    
    s+="\\section{Gantt chart}\n\n"
    s+= "\\includegraphics[width=16cm]{gantt.png}\n\n"
    return s

def getDeliverableText(D_id):
    s=""
    #look up the id in the deliverables tables
    df = df_deliverable[df_deliverable['Deliverable']==D_id]
    s = "\\subsection*{Deliverable D" +str(df['Deliverable'].values[0]) +": "+ str(df['Name'].values[0]) + "}\n\n"

    quarter_start = df_deliverable[df_deliverable['Deliverable']==D_id]['Quarter'].iloc[0]
    dur = df_deliverable[df_deliverable['Deliverable']==1.12]['Quarter'].iloc[0]
    quarter_end = quarter_start+dur
    owner = df_deliverable[df_deliverable['Deliverable']==D_id]['Owner'].iloc[0]
    status = df_deliverable[df_deliverable['Deliverable']==D_id]['Status'].iloc[0]


    date_quarter_start = date_project_start+ relativedelta(months=+((quarter_start-1)*3))
    date_quarter_end = date_project_start+ relativedelta(months=+((quarter_start-1)*3)) + relativedelta(months=+(dur*3))

    s+=  "Start quarter: Q%i (%s) \n \n End Quarter: Q%i (%s) \n\n Leader: %s\n\n  Status: %s \n\n "%(quarter_start, date_quarter_start.strftime("%Y-%m-%d"), quarter_end, date_quarter_start.strftime("%Y-%m-%d"),  owner, status)

    #print each row of description text
    s += "\subsubsection*{Description of work}\n\n"
    df = df_desc[df_desc['Deliverable']==D_id]
    for row in df.iterrows():
        s += str(row[1][1]) + "\n\n"

    #get list of associated risks
    s += "\\subsubsection*{Associated risks:}\n\n"
    df = rd[rd['Deliverable']==D_id]
    b_found_anything = False
    for i in range(0, df.shape[0]):
        s+= "R%s: %s\n\n"%( df.iloc[i]['RiskID'] , df.iloc[i]['RiskName'])
        b_found_anything=True
    if not b_found_anything:
        s+="None\n\n"

    #get dependencies
    s += "\subsubsection*{Depends on deliverables:}\n\n"
    df = df_Dependency.merge(df_deliverable)
    df = df[df['Deliverable']==D_id]
    #	df = df_Dependency[df_Dependency['Deliverable']==D_id]
    b_found_anything = False
    for i in range(0, df.shape[0]):
        s+= "D"+str(df.iloc[i]['DependsOn'])+": "+df.iloc[i]['Name']+"\n\n"
        b_found_anything=True
    if not b_found_anything:
        s+="None\n\n"
    s+="\n\n"
    s += "\subsubsection*{Prerequisite for deliverables:}\n\n"
    df = df_Dependency.merge(df_deliverable)
    df = df[df['DependsOn']==D_id]
    b_found_anything = False
    for i in range(0, df.shape[0]):
        s+= "D"+str(df.iloc[i]['Deliverable'])+": "+df.iloc[i]['Name']+"\n\n"
        b_found_anything=True
    if not b_found_anything:
        s+="None\n\n"
    s+="\n\n"

    #costings
    s+= "\subsubsection*{Resources}\n\n"

    #per spend category
    df_cost_del = df_cost[df_cost['Deliverable']==D_id]
    s+= "\\begin{tabular}{ | l | l | r | }\n"    #TODO I made a latextabl;e function above, use it!
    s+= "\\hline\n "
    s+= "Category & Partner & Cost \\\\ \n "
    s+= "\\hline\n "
#    if D_id==1.2:
#        pdb.set_trace()
    #for category in  df_cost_del.groupby(['Category','Partner']).indices:
    #    cost = df_cost_del.groupby(['Category','Partner']).get_group(category)['Cost'].iloc[0]

    sums = df_cost_del.groupby(['Category','Partner']).sum()
    for i in range(0, sums.shape[0]):
        row = sums.iloc[i]
        s+= "%s & %s &  %s \\\\ \n"%(row.name[0], row.name[1]  , '{:,.2f}'.format(row.Cost))
    s+= "\\hline\n "
    s+= "\\end{tabular}\n\n"


    #total deliverable cost
    cost = df_cost[df_cost['Deliverable']==D_id].sum()['Cost']
    s+= "Total deliverable cost:  %s"%'{:,.2f}'.format(cost)

    s+="\n\n"


    s+= "\subsubsection*{Per-item costs}\n\n"
    #per item (TODO only big ones for MO version;  show all for internal version?)
    df_cost_del = df_cost[df_cost['Deliverable']==D_id]
    s+= "\\begin{tabular}{ | l | c | c | r | c | }\n"    #TODO I made a latextable function above, use it!
    s+= "\\hline\n "
    s+= "Item & Partner & Category & Cost \\\\ \n "
    s+= "\\hline\n "
    for i in range(0, df_cost_del.shape[0]):
        s+= "%s & %s & %s & %s %s\\\\ \n"%( df_cost_del.iloc[i]['Item'], df_cost_del.iloc[i]['Partner'], df_cost_del.iloc[i]['Category'],  '{:,.2f}'.format(df_cost_del.iloc[i]['Cost']), df_cost_del.iloc[i]['SpendStatus'] )
    s+= "\\hline\n "
    s+= "\\end{tabular}\n\n"

    s+="\\newpage"
    return s

def getPartnerSpendTable():
    s="\\newpage\n\n"
    s="\section{Spend profiles}\n\n"

    partners = ["LINCOLN", "ARWAC"]
    categories = ["MATERIALS", "LABOUR", "CAPEX", "TRAVEL", "OTHER", "OVERHEADS", "SUBCON"]
    quarters = [1,2,3,4,5,6,7,8]

    for partner in partners:
        M = np.zeros((len(categories),len(quarters)))
        s_partner="\subsection{Spend profile for partner %s}\n\n"%partner
        icategory=-1
        for category in categories:
            icategory+=1
            iquarter=-1
            for quarter in quarters:
                iquarter+=1
                cost = 0
                if (partner, quarter, category) in gb.indices:
                    cost = gb.get_group((partner, quarter, category)).sum()['Cost'] 

                #s_partner += "%s %s %s %f"%(partner, quarter, category, cost)
                M[icategory, iquarter] = cost
            #s_partner+="\n\n"
        #pdb.set_trace()
        s_partner += makeLatexTable(M, quarters, categories)
        s+=s_partner
    return s

def getRiskMatrix():
    s="\\newpage\n\n"
    s+="\\section{Risk register}\n\n"
    nrisks = df_risk.shape[0]
    for i in range(0, nrisks):
        row = df_risk.iloc[i]
        s+= "\\subsection*{Risk R%s}\n\n"%(str(row['RiskID'])+": "+row['RiskName'])
        s+= "%s\n\n"%row['RiskDesc']
        s+= "Category: %s\n\n"%row['Category']
        s+= "Treatment: %s\n\n"%row['Treatment']
        s+= " %s\n\n"%row['MitigationDesc']

        s+= "Impact Pre-Mitigation: %s\n\n"%row['ImpactPre']
        s+= "Probability Pre-Mitigation: %s\n\n"%row['ProbabilityPre']
        s+= "Score Pre-Mitigation: %s\n\n"%(row['ImpactPre']*row['ProbabilityPre'])

        s+= "Impact Post-Mitigation: %s\n\n"%row['ImpactPost']
        s+= "Probability Post-Mitigation: %s\n\n"%row['ProbabilityPost']
        s+= "Score Post-Mitigation: %s\n\n"%(row['ImpactPost']*row['ProbabilityPost'])

        s+= "Owner: %s\n\n"%row['RiskOwner']



        #get list of associated deliverables
        s += "Associated deliverables: "
        df = rd[rd['RiskID']==row['RiskID']]
        for i in range(0, df.shape[0]):
            s+= "D%s "%( df.iloc[i]['Deliverable'])
        s+="\n\n"

    return s

def makeStageTwoPlan(fn_out, df_deliverable):
    f_out = open(fn_out, 'w')
    s = latexHeader()
    s+="\\newpage\n\n"
    s+="\\section{Deliverables}\n\n"
    #get list of deliverables and text for each one
    for index, row in df_deliverable.iterrows():
        D_id = row['Deliverable']
        s += getDeliverableText(D_id)

    s+= getPartnerSpendTable()

    s+= getRiskMatrix()

    s += latexFooter()
    f_out.write(s)
    cmd = "./runpdflatex.sh"
    subprocess.call(cmd, shell="True")



def makeGanttChart(fn_gantt_png, date_project_start):
    g = GanttChart.GanttChart(date_project_start)

    for i in range(0, df_deliverable.shape[0]):
        t_start = df_deliverable.iloc[i]['Quarter']
        t_end   = df_deliverable.iloc[i]['Quarter'] + df_deliverable.iloc[1]['Duration']
        name    = "D"+str(df_deliverable.iloc[i]['Deliverable'])+": "+str(df_deliverable.iloc[i]['Name'])
        g.drawTask( i, t_start, t_end, name)

    g.draw(fn_gantt_png)


dir_csv = "csv/ARWAC_costs_sheets - "
dir_out = "out/"
df_cost = pd.read_csv(dir_csv+'Cost.csv', skiprows=0, header=0) 
df_deliverable = pd.read_csv(dir_csv+'Deliverable.csv', skiprows=0, header=0) 
df_desc = pd.read_csv(dir_csv+'DeliverableText.csv', skiprows=0, header=0) 
df_risk = pd.read_csv(dir_csv+'Risk.csv', skiprows=0, header=0) 
df_RiskDeliverable = pd.read_csv(dir_csv+'RiskDeliverable.csv', skiprows=0, header=0) 
df_Dependency = pd.read_csv(dir_csv+'Dependency.csv', skiprows=0, header=0) 
df_Project = pd.read_csv(dir_csv+'Project.csv', skiprows=0, header=0) 
df_WorkPackage = pd.read_csv(dir_csv+'WorkPackage.csv', skiprows=0, header=0) 

gb = df_cost.merge(df_deliverable)[['Partner', 'Quarter', 'Category', 'Cost']].groupby(['Partner','Quarter', 'Category'])


rd = df_RiskDeliverable.merge(df_risk)
#rd = df_RiskDeliverable.merge(df_risk).merge(df_deliverable)  #testing risk-deliverable join


date_project_start = datetime.datetime.strptime('2019-04-01_00:09:00.00', "%Y-%m-%d_%H:%M:%S.%f" )


makeGanttChart(dir_out+"gantt.png", date_project_start)
fn_tex = dir_out+"ARWAC_stage2plan.tex"
makeStageTwoPlan(fn_tex, df_deliverable)
