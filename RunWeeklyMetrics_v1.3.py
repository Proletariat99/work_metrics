__author__ = 'Dave Dyer'
__version__ = '1.0'

import os
import xlrd
from matplotlib import pyplot as plt
import operator
import numpy as np
import datetime
from operator import itemgetter


#################################################
#           Paths, files and dates              #
#################################################
# Paths ==============
CTpath = "C:\\pathname\\"
oCTpath = CTpath
CTfn = os.listdir(CTpath)
oCTfn = os.listdir(oCTpath)

# Files ==============
# Some psuedo code to user here (to write in my "extra" time)
# Per line in file (old and current)
    # if date(file[42:]) is greater than date.today - 30:
        # filenamelist.append(file)
    # Then take all files and get the numbers using xlrd.open_workbook
# Instead, we'll just do this:
                           #"SEMA_Report_and_Alert_Metrics_week_Ending_08-29-2013"
fn_txt = CTpath + "filename"
fntxt = "output_filename"
# Dates ==============
curDate = "08-29-2013"
ofnDates = ["08-01-2013", "08-08-2013", "08-15-2013","08-22-2013"]
lmDates = ["06-27-2013", "07-04-2013", "07-11-2013","07-18-2013", "07-25-2013"]
curfn = fn_txt + curDate + ".xlsx"

################################################################
#           Getting Current Data in                            #
################################################################

bk = xlrd.open_workbook(curfn)
    # print str(xl)
n = 0
sheets = bk.sheet_by_name("Totals for Plot")

# making a giant list of all field values
# comes from the last 4 weeks of data (4 xl files)
values = []
for m in sheets._cell_values:
    values.append(m)
    # for tabs in range(file.nsheets):
    #     spreadsheets[n].append(file.sheet_by_index(tab))
    #     print 'tab = ' + str(tab)
    #     print 'n = ' + str(n)
    # n = n+1


# This Week's Dates ================
#print "first day of the  current week's spreadsheet is " + str(xlrd.xldate_as_tuple(values[0][1],1))
week0dates = values[0][1:]

################################################################
#            Getting Old Data in                               #
################################################################
# Grabbing data using xlrd and making pretty lists
ofns = []
for n in ofnDates:
    ofns.append(oCTpath + fn_txt + n + ".xlsx")

obk = []
for xl in ofns:
    obk.append(xlrd.open_workbook(xl))
    # print str(xl)

n = 0
osheets = []
for xls in obk:
    osheets.append(xls.sheet_by_name("Totals for Plot"))

# making a giant list of all field values
# comes from the last 4 weeks of data (4 xl files)
ovalues = []
for n in osheets:
    # print n
    for m in n._cell_values:
        ovalues.append(m)
    # for tabs in range(file.nsheets):
    #     spreadsheets[n].append(file.sheet_by_index(tab))
    #     print 'tab = ' + str(tab)
    #     print 'n = ' + str(n)
    # n = n+1
# Last 4 weeks of dates ==================
week1dates = ovalues[0][1:8]
week2dates = ovalues[19][1:8]
week3dates = ovalues[38][1:8]
week4dates = ovalues[57][1:8]

# More Date Munging (as datetime.date())
#########################
# Getting Last Month In #
#########################
lmfns = []
for n in lmDates:
    lmfns.append(oCTpath + fn_txt + n + ".xlsx")

lmbk = []
for xl in lmfns:
    lmbk.append(xlrd.open_workbook(xl))
    # print str(xl)

n = 0
lmsheets = []
for xls in lmbk:
    lmsheets.append(xls.sheet_by_name("Totals for Plot"))

# making a giant list of all field values
# comes from the last 4 weeks of data (4 xl files)
lmValues = []
for n in lmsheets:
    # print n
    for m in n._cell_values:
        lmValues.append(m)
week5dates = lmValues[0][1:8]
week6dates = lmValues[19][1:8]
week7dates = lmValues[38][1:8]
week8dates = lmValues[57][1:8]
week9dates = lmValues[76][1:8]

lastMonthDates = week5dates + week6dates + week7dates + week8dates + week9dates

#####################################################
# Mushing all Dates together into tuples lists, etc #
# May need to order as set() later, just in case. ###
#####################################################
allDates = week1dates + week2dates + week3dates + week4dates + week0dates
aD = allDates           # apparently laziness is fueling my variable naming convention today.
# Can Use these below tuples later, but they're only used for labels currently
dateListOfTuples = []
mDay = []
dList = []
mList = []
for entry in allDates:
    dateListOfTuples.append(xlrd.xldate_as_tuple(entry, 1))
    mDay.append(xlrd.xldate_as_tuple(entry, 1)[1:3])
    mList.append(list(xlrd.xldate_as_tuple(entry, 1))[1])
    dList.append(list(xlrd.xldate_as_tuple(entry, 1))[2])

dateTimes = []
dateNames = []
for dateTuple in dateListOfTuples:
    dateTimes.append(datetime.date(dateTuple[0], dateTuple[1], dateTuple[2]))
    dateNames.append(datetime.date(dateTuple[0], dateTuple[1], dateTuple[2]).isoformat())
################################################
# Grabbing Daily Totals                        #
################################################
# Some notes on totals:
# DTE = "Daily Total Events.csv" from my email / Envision.  .csv is manually updated by me from email.
# totesExternal = (Total ISS + Total SecureWorks + Total Bruteforce Events)
# totesInternal = (Total Windows Events + Total ISS Events)
# totesBadActor = (Total Watchlist Firewall + Total Watchlist Bluecoat)
#
# >>> mV[0]
# ['', 40012.0, 40013.0, 40014.0, 40015.0, 40016.0, 40017.0, 40018.0]
# >>> mV[1] - 07-26-2013
# [u'Total ISS Events', 38073916.0, 38235015.0, 39683115.0, 38371166.0, 38858729.0, 39031435.0, 38136072.0]

# >>> mV[20] - 07-18-2013
# [u'Total ISS Events', 37114112.0, 37840593.0, 38561997.0, 39298130.0, 37695179.0, 37893162.0, 38933676.0]

# >>> mV[39] - 07-11-2013
# [u'Total ISS Events', 34542314.0, 35858019.0, 37567913.0, 37321014.0, 36295982.0, 37581699.0, 37741778.0]

# >>> mV[58] - 07-04-2013
# [u'Total ISS Events', 32748087.0, 33205475.0, 34144845.0, 34294079.0, 35377910.0, 35318954.0, 34359102.0]

# >>> mV[77] - 08-01-2013
# [u'Total ISS Events', 37370335.0, 37834224.0, 24871817.0, 21795117.0, 20702037.0, 17775681.0, 16837840.0]
#TotalEnvisionEvents = [x for x in jN[::7]]
# keyNames = []
# for line in mergedValues:
#         for stuff in line:
#             if type(stuff) == str:
#                 keyNames.append(stuff)
#             else:
#                 pass

mergedValues = ovalues + values


mV = mergedValues # that name sucked.
justNums = []
jN = justNums          #  This one too
for lineItem in mV:
    jN.append(lineItem[1:])

lmjN = []
for lineItem in lmValues:
    lmjN.append(lineItem[1:])

n19 = 19                    # Because there are 19 lines until mV reset (loops back around to the next day.)

nISS = 1
nEnv = 2
nSW = 3
nWin = 4
nUx = 5
nWLfw = 6
nWLbc = 7

#  these are mushing together the totals from the top 7 lines (totals)
tISS = jN[nISS+0*n19] + jN[nISS+1*n19] + jN[nISS+2*n19] + jN[nISS+3*n19] + jN[nISS+4*n19]
tEnv = jN[nEnv+0*n19] + jN[nEnv+1*n19] + jN[nEnv+2*n19] + jN[nEnv+3*n19] + jN[nEnv+4*n19]
tSW= jN[nSW+0*n19] + jN[nSW+1*n19] + jN[nSW+2*n19] + jN[nSW+3*n19] + jN[nSW+4*n19]
tWin = jN[nWin+0*n19] + jN[nWin+1*n19] + jN[nWin+2*n19] + jN[nWin+3*n19] + jN[nWin+4*n19]
tUx = jN[nUx+0*n19] + jN[nUx+1*n19] + jN[nUx+2*n19] + jN[nUx+3*n19] + jN[nUx+4*n19]
tWLfw = jN[nWLfw+0*n19] + jN[nWLfw+1*n19] + jN[nWLfw+2*n19] + jN[nWLfw+3*n19] + jN[nWLfw+4*n19]
tWLbc = jN[nWLbc+0*n19] + jN[nWLbc+1*n19] + jN[nWLbc+2*n19] + jN[nWLbc+3*n19] + jN[nWLbc+4*n19]

tInternal = map(operator.add, tWin, tISS)
tExtTemp = map(operator.add, tISS, tSW)
tExternal = map(operator.add, tExtTemp, tUx)
tBA = map(operator.add, tWLfw, tWLbc)

################################################
# Now Events Analyzed and Escalated Counts     #
################################################
# Some Notes on Internal tickets
# IA = aI + aP  Internal analyzed (total) = analyzed internal + analyzed policy
# IE = eI + eP same with escalated
# AE = aE
# "last week" has been deprecated.  Monthly data gives context.
naS = 8
neS = 9
naI = 10
neI = 11
naE = 12
neE = 13
naP = 14
neP = 15
naB = 16
neB = 17

# These are mushing together the ticket counts from the next 10 lines
taS = jN[naS+0*n19] + jN[naS+1*n19] + jN[naS+2*n19] + jN[naS+3*n19] + jN[naS+4*n19]
teS = jN[neS+0*n19] + jN[neS+1*n19] + jN[neS+2*n19] + jN[neS+3*n19] + jN[neS+4*n19]
taI = jN[naI+0*n19] + jN[naI+1*n19] + jN[naI+2*n19] + jN[naI+3*n19] + jN[naI+4*n19]
teI = jN[neI+0*n19] + jN[neI+1*n19] + jN[neI+2*n19] + jN[neI+3*n19] + jN[neI+4*n19]
taE = jN[naE+0*n19] + jN[naE+1*n19] + jN[naE+2*n19] + jN[naE+3*n19] + jN[naE+4*n19]
teE = jN[neE+0*n19] + jN[neE+1*n19] + jN[neE+2*n19] + jN[neE+3*n19] + jN[neE+4*n19]
taP = jN[naP+0*n19] + jN[naP+1*n19] + jN[naP+2*n19] + jN[naP+3*n19] + jN[naP+4*n19]
teP = jN[neP+0*n19] + jN[neP+1*n19] + jN[neP+2*n19] + jN[neP+3*n19] + jN[neP+4*n19]
taB = jN[naB+0*n19] + jN[naB+1*n19] + jN[naB+2*n19] + jN[naB+3*n19] + jN[naB+4*n19]
teB = jN[neB+0*n19] + jN[neB+1*n19] + jN[neB+2*n19] + jN[neB+3*n19] + jN[neB+4*n19]
#taB and teB are good right here.

# <------------- Aggregates Ticket Values ---------------->
tAI = np.array(taI) + np.array(taP)
tEI = np.array(teI)
tAE = np.array(taE)
tEE = np.array(teE)
tAB = np.array(taB)
tEB = np.array(teB)
# tAB and tEB good right here, too.
lmnaS = 8
lmneS = 9
lmnaI = 10
lmneI = 11
lmnaE = 12
lmneE = 13
lmnaP = 14
lmneP = 15
lmnaB = 16
lmneB = 17

# These are mushing together events from the previous 5 weeks (last month = lm)
lmtaS = lmjN[lmnaS+0*n19] + lmjN[lmnaS+1*n19] + lmjN[lmnaS+2*n19] + lmjN[lmnaS+3*n19] + lmjN[lmnaS+4*n19]
lmteS = lmjN[lmneS+0*n19] + lmjN[lmneS+1*n19] + lmjN[lmneS+2*n19] + lmjN[lmneS+3*n19] + lmjN[lmneS+4*n19]
lmtaI = lmjN[lmnaI+0*n19] + lmjN[lmnaI+1*n19] + lmjN[lmnaI+2*n19] + lmjN[lmnaI+3*n19] + lmjN[lmnaI+4*n19]
lmteI = lmjN[lmneI+0*n19] + lmjN[lmneI+1*n19] + lmjN[lmneI+2*n19] + lmjN[lmneI+3*n19] + lmjN[lmneI+4*n19]
lmtaE = lmjN[lmnaE+0*n19] + lmjN[lmnaE+1*n19] + lmjN[lmnaE+2*n19] + lmjN[lmnaE+3*n19] + lmjN[lmnaE+4*n19]
lmteE = lmjN[lmneE+0*n19] + lmjN[lmneE+1*n19] + lmjN[lmneE+2*n19] + lmjN[lmneE+3*n19] + lmjN[lmneE+4*n19]
lmtaP = lmjN[lmnaP+0*n19] + lmjN[lmnaP+1*n19] + lmjN[lmnaP+2*n19] + lmjN[lmnaP+3*n19] + lmjN[lmnaP+4*n19]
lmteP = lmjN[lmneP+0*n19] + lmjN[lmneP+1*n19] + lmjN[lmneP+2*n19] + lmjN[lmneP+3*n19] + lmjN[lmneP+4*n19]
lmtaB = lmjN[lmnaB+0*n19] + lmjN[lmnaB+1*n19] + lmjN[lmnaB+2*n19] + lmjN[lmnaB+3*n19] + lmjN[lmnaB+4*n19]
lmteB = lmjN[lmneB+0*n19] + lmjN[lmneB+1*n19] + lmjN[lmneB+2*n19] + lmjN[lmneB+3*n19] + lmjN[lmneB+4*n19]

lmtAI = np.array(lmtaI) + np.array(lmtaP)
lmtEI = np.array(lmteI)
lmtAE = np.array(lmtaE)
lmtEE = np.array(lmteE)
lmtAB = np.array(lmtaB)
lmtEB = np.array(lmteB)


#uc_group = sorted(uc_group.iteritems(), key=operator.itemgetter(1))
################################################################
#            List comprehension example / Test                 #
################################################################
# l = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 'steve', 'gordo', 'stringy', 5.0]
# [type(y) == str for type(x) for x in l]
# [type(x) for x in l if type(x) == str]  - includes an if statement

#
# keyNames = []
# for element in mergedValues:
#     keyNames.append(x for x in element if type(x) == unicode)

################################################################
#            Plotting and Labeling                             #
################################################################
TTFS = 20      # Title font size
AXFS = 18      # Axis font size
TFS = 14       # text fart size
LFS = 16       # legernd flaunt sides
pCol1 = "#339900"
pCol2 = "#ff9900"
pCol3 = "#000000"


#############################################
#               Internal                    #
#############################################
# <------------ Totals ---------------->
fig0 = plt.figure()
axI1 = fig0.add_subplot(111)

# indexes
ix = np.arange(len(mDay)) # Bar xticks
tx = np.arange(len(mDay)) + 0.4# points xticks
tx2 = [tx, tx]

#some last internal calculations for limits)
meanTI = np.mean(tInternal)
sdTI = np.std(tInternal)
topTI = meanTI + sdTI * 4
if meanTI - sdTI * 4 < 0:
    botTI = 0
elif meanTI - sdTI * 4 >= 0:
    botTI = meanTI - sdTI * 4
yIticks = np.arange(botTI, topTI+sdTI, sdTI)/1e6

axI1.bar(ix, tInternal, width=0.75,  color='#99CCFF', label='Event Totals')

axI1.set_xticklabels(dateNames, rotation=270)
axI1.set_xticks(ix)
axI1.set_yticklabels(yIticks.astype(int))
axI1.set_ylim(0, 2e9)
axI1.set_yticklabels(yIticks.astype(int))
axI1.set_title("Internal Threat Metric", fontsize=TTFS)
axI1.set_xlabel('Date', fontsize=AXFS)
axI1.set_ylabel('Event Totals (Millions)', fontsize=AXFS)

# <------------ Tickets ---------------->

axI2 = plt.twinx(axI1)

lI1 = axI2.plot(tx, tAI, ms=13, marker="o", color=pCol1, linewidth=3)
lI2 = axI2.plot(tx, tEI, ms=13, marker="s", color=pCol2, linewidth=3)

lmai_horizDash = [tAI, lmtAI]
lmei_horizDash = [tEI, lmtEI]
lmai_diff = [(tAI-lmtAI)]
lmei_diff = [(tEI-lmtEI)]

lmI1 = axI2.plot(tx2, lmai_horizDash, ls="dashed", color=pCol3)
lmI3 = axI2.plot(tx, lmtAI, "b_", ms=13, color=pCol3, ls="None")
lmI2 = axI2.plot(tx2, lmei_horizDash, ls="dashed", color=pCol3)
lmI4 = axI2.plot(tx, lmtEI, "b_", ms=13, color=pCol3, ls="None")

axI2.set_ylabel("Number of Events Analyzed or Escalated", fontsize=AXFS, rotation=270)
xI1, xI2, yI1, yI2 = axI2.axis()         # get the current axis
axI2.axis = (xI1, xI2, 0, yI2)

# Legend Stuff (Internal)
handles1, labels1 = axI1.get_legend_handles_labels()
handles2, labels2 = axI2.get_legend_handles_labels()
labels1 = ["Event Totals (Millions)"]
labels2 = ["Events Analyzed", "Events Escalated", "Last Week's Data"]
lines2 = [lI1, lI2]
lines1 = [handles1]
leg1 = axI1.legend(handles1, labels1, loc=2, fancybox=True, fontsize=LFS)
leg2 = axI2.legend(labels2, loc=1, markerscale=0.25, fontsize=LFS)
leg1.get_frame().set_alpha(0.5)
leg2.get_frame().set_alpha(0.5)


#############################################
#               external                    #
#############################################
# <------------ Totals ---------------->
fig2 = plt.figure()
axE1 = fig2.add_subplot(111)

# Some graph calculations
meanTE = np.mean(tExternal)
sdTE = np.std(tExternal)
topTE = meanTE + sdTE * 4
if meanTE - sdTE * 4 < 0:
    botTE = 0
elif meanTE - sdTE * 4 >= 0:
    botTE = meanTE - sdTE * 4

axE1.bar(ix, tExternal, width=0.75,  color='#99CCFF', label='Event Totals')

axE1.set_xlabel("Date", fontsize=AXFS)
axE1.set_ylabel("Event Totals (Millions)", fontsize=AXFS, rotation=90)
axE1.set_xticklabels(dateNames, rotation=270)
axE1.set_xticks(ix+0.5)

# <------------ Tickets ---------------->
axE2 = plt.twinx(axE1)

lmae_horizDash = [tAE, lmtAE]
lmee_horizDash = [tEE, lmtEE]
lmae_diff = [(tAE-lmtAE)]
lmee_diff = [(tEE-lmtEE)]

lE1 = axE2.plot(tx, tAE, ms=13, marker="o", color=pCol1, linewidth=2)
lE2 = axE2.plot(tx, tEE, ms=13, marker="s", color=pCol2, linewidth=2)
lmE1 = axE2.plot(tx2, lmae_horizDash, ls="dashed", color=pCol3)
lmE2 = axE2.plot(tx2, lmee_horizDash, ls="dashed", color=pCol3)
lmE3 = axE2.plot(tx, lmtAE, "b_", ms=13, color=pCol3, ls="None")
lmE4 = axE2.plot(tx, lmtEE, "b_", ms=13, color=pCol3)

axE1.set_title("External Threat", fontsize=TTFS)
axE1.set_ylabel('Event Totals (Millions)', fontsize=AXFS)
axE2.set_ylabel('Number of Events Analyzed or Escalated', fontsize=AXFS, rotation=270)
yEticks = np.arange(botTE, topTE+sdTE, sdTE)/1e6
extYmaxP = max(tAE)
EYPM = extYmaxP+(0.40*extYmaxP)
axE2.set_ylim(top=EYPM)
axE1.set_title("External Threat Metric", fontsize=TTFS)
axE1.set_yticklabels(yEticks.astype(int))
axE1.set_ylim(0, topTE)
axE1.set_yticklabels(yEticks.astype(int))

# Legend Stuff (Internal)
handles3, labels3 = axE1.get_legend_handles_labels()
handles4, labels4 = axE2.get_legend_handles_labels()
labels3 = ["Event Totals (Millions)"]
labels4 = ["Events Analyzed", "Events Escalated", "Last Week's Data"]
lines4 = [lE1, lE2]
lines3 = [handles3]
leg3 = axE1.legend(handles1, labels1, loc=2, fancybox=True, fontsize=LFS)
leg4 = axE2.legend(labels2, loc=1, markerscale=0.25, fontsize=LFS)
leg3.get_frame().set_alpha(0.5)
leg4.get_frame().set_alpha(0.5)





#############################################
#               Watchlist/BadActor          #
#############################################
# <------------ Totals ---------------->
fig3 = plt.figure()
axB1 = fig3.add_subplot(111)

# Some graph calculations
meanTB = np.mean(tBA)
sdTB = np.std(tBA)
topTB = meanTB + sdTB * 4
if meanTB - sdTB * 4 < 0:
    botTB = 0
elif meanTB - sdTB * 4 >= 0:
    botTB = meanTB - sdTB * 4

axB1.bar(ix, tBA, width=0.75,  color='#99CCFF', label='Event Totals')

axB1.set_xlabel("Date", fontsize=AXFS)
axB1.set_ylabel("Event Totals (Millions)", fontsize=AXFS, rotation=90)
axB1.set_xticklabels(dateNames, rotation=270)
axB1.set_xticks(ix+0.5)

# <------------ Tickets ---------------->
axB2 = plt.twinx(axB1)

lmab_horizDash = [tAB, lmtAB]
lmeb_horizDash = [tEB, lmtEB]
lmab_diff = [(tAB-lmtAB)]
lmeb_diff = [(tEB-lmtEB)]

lB1 = axB2.plot(tx, tEB, ms=13, marker="s", color=pCol1, linewidth=2)
lB2 = axB2.plot(tx, tAB, ms=13, marker="o", color=pCol2, linewidth=2)
lmB1 = axB2.plot(tx2, lmab_horizDash, ls="dashed", color=pCol3)
lmB2 = axB2.plot(tx2, lmeb_horizDash, ls="dashed", color=pCol3)
lmB3 = axB2.plot(tx, lmtAB, "b_", ms=13, color=pCol3, ls="None")
lmB4 = axB2.plot(tx, lmtEB, "b_", ms=13, color=pCol3, ls="None")

yBticks = np.arange(botTB, topTB+sdTB, sdTB)/1e6
extYmaxP = max(tAB)
BYPM = extYmaxP+(0.40*extYmaxP)
axB2.set_ylim(bottom=0, top=7)

axB1.set_title("Watchlist Threat Metric", fontsize=TTFS)
axB1.set_yticklabels(yBticks.astype(int))
axB1.set_ylim(0, topTB)
axB1.set_yticklabels(yBticks.astype(int))
axB2.set_ylabel('Number of Events Analyzed or Escalated', fontsize=AXFS, rotation=270)

# Legend Stuff (Internal)
handles3, labels3 = axB1.get_legend_handles_labels()
handles4, labels4 = axB2.get_legend_handles_labels()
labels3 = ["Event Totals (Millions)"]
labels4 = ["Events Analyzed", "Events Escalated", "Last Beek's Data"]
lines4 = [lB1, lB2]
lines3 = [handles3]
leg3 = axB1.legend(handles1, labels1, loc=2, fancybox=True, fontsize=LFS)
leg4 = axB2.legend(labels2, loc=1, markerscale=0.25, fontsize=LFS)
leg3.get_frame().set_alpha(0.5)
leg4.get_frame().set_alpha(0.5)

weekends = np.arange(7*5)


######################################################
# I fire my magic missle at the darkenss             #
######################################################

plt.show()
today = datetime.datetime.now()
today = today.isoformat()[0:10]

win_savepath = "C:\\savepath\\"
mac_savepath = "/Users/savepath"


fig0.savefig(win_savepath + today + ' InternalThreat' + '.png')
fig2.savefig(win_savepath + today + ' ExternalThreat' + '.png')
fig3.savefig(win_savepath + today + ' WatchlistThreat' + '.png')
