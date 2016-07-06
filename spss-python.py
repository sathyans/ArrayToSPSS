# Python SPSS Integration
import spss, spssaux
def spssloadfromarray(arr):
    cmd="DATA LIST /a 1-3. \nBEGIN DATA \n"
    for case in arr:
        cmd = cmd + str(case) + "\n"
    cmd=cmd + "END DATA."
    print cmd
    spss.Submit(cmd)
    spss.Submit('FREQ VAR a.')
#cmd =r"""FILE HANDLE data /name='M:\Institutional Assessment and Research\private\COMMON FOLDER\Retention\CSRDE\CO-CSRDE-FTF-fall2013.sav'.
#FILE HANDLE report /name='M:\Institutional Assessment and Research\private\COMMON FOLDER\Peer Institutions\Peers_Output_2013.xls'.
#GET FILE=data.
#DATASET NAME CSRDE.
#DATASET ACTIVATE CSRDE.
#RECODE CAMPUS_NAME ('Bakersfield'=1) ('Channel Islands'=1) ('Dominguez Hills'=1) 
#('East Bay'=1) ('Humboldt'=1) ('Monterey Bay'=1) ('San Marcos'=1) ('Sonoma'=1) 
#('Stanislaus'=1) (ELSE=0) INTO PEERS.
#VARIABLE LABELS  PEERS 'Peer Institution'.
#VALUE LABELS PEERS 1 'Yes' 0 'No'.
#EXECUTE.
#COMPUTE UseThis = 0.
#EXECUTE.
#VALUE LABELS UseThis 1 'Yes' 0 'No'.
#IF ( SEX EQ '' & ETHNICITY EQ '' & PEERS EQ 1) UseThis =1.
#EXECUTE.
#FILTER BY UseThis.
#EXECUTE.
#COMPUTE Ret1 = 100 *  CONT_1_YR.
#COMPUTE Ret2 = 100 *  CONT_2_YR.
#EXECUTE.
#VARIABLE LABELS Ret1 'After \n 1 Year'.
#VARIABLE LABELS Ret2 'After \n 2 Years'.
#COMPUTE Ret4 = 100 *  CONT_4_YR.
#COMPUTE Ret6 = 100 *  CONT_6_YR.
#COMPUTE Grad4 = 100 *  GRAD_4_YR. 
#COMPUTE Grad6 = 100 * GRAD_6_YR.
#COMPUTE RG4 = Ret4 + Grad4.
#COMPUTE RG6 = Ret6 + Grad6.
#EXECUTE.
#VARIABLE LABELS Ret4 'Within 4 Years \n CONT'.
#VARIABLE LABELS Ret6 'Within 6 Years \n CONT'.
#VARIABLE LABELS Grad4 'Within 4 Years \n GRAD'.
#VARIABLE LABELS Grad6 'Within 6 Years \n GRAD'.
#VARIABLE LABELS RG4 'Campus Overall (4-yr)'.
#VARIABLE LABELS RG6 'Campus Overall (6-yr)'.
#SORT CASES  BY FALL.
#SPLIT FILE SEPARATE BY FALL.
#OMS /SELECT TABLES /IF COMMANDS =['CTABLES'] 
#/EXCEPTIF SUBTYPES =['Notes','Case Summary']
#/DESTINATION OUTFILE = report FORMAT =XLS
#/NOWARN.
#ECHO 'Systemwide numbers are calculated by the CO'.
#ECHO 'See http://www.asd.calstate.edu/csrde/index.shtml#ftf'.
#ECHO 'Excel workbook calculates unweighted peer average'.
#ECHO 'Table for CSU Peers Benchmark Report'.
#CTABLES 
  #/VLABELS VARIABLES=CAMPUS_NAME FALL Ret1 Ret2 Grad4 Ret4 RG4 Grad6 Ret6 RG6 
    #DISPLAY=LABEL 
  #/TABLE CAMPUS_NAME BY FALL > (Ret1 [MEAN PCT40.0] + Ret2 [MEAN PCT40.0] + Grad4 [MEAN PCT40.0] + Ret4 [MEAN PCT40.0] + RG4 [MEAN PCT40.0] + Grad6 [MEAN PCT40.0] + Ret6 [MEAN PCT40.0] + RG6 [MEAN PCT40.0]) 
  #/CATEGORIES VARIABLES=CAMPUS_NAME FALL ORDER=D KEY=MEAN(Ret1) EMPTY=EXCLUDE
  #/TITLES 
  #TITLE = 'Degree-Seeking First Time, Full Time Freshman Retention Rates, by Campus'
  #CAPTION ='Source: Data Compiled from Chancellors Office Website: http://www.asd.calstate.edu/csrde/index.shtml'															
  #CORNER = 'CAMPUS'.		
#OMSEND.
#DATASET CLOSE CSRDE."""
#spss.Submit(cmd)

# fetch from web
import urllib, os
baseurl = "http://www.asd.calstate.edu/csrde/ftf/2011xls/"
csu = ["Sys","Bak","CI","Chi","DH","EB","Fre","Ful","Hum","LA","LB","MA","MB","Nor","Pom","Sac","SB","SD","SF","SJ","SLO","SM","Son","Sta"]
ext = ".xls"
d = "c:/temp/"    
a = [["Sys",92,85,64],["Bak",85,74,56],["CI",56,48,37]]
for element in a:
    print element[0][0:3]
#for campus in csu:
    #urltarget = baseurl + campus + ext
    #storeit = d + campus + ext
    #urllib.urlretrieve(urltarget,storeit)
    #print urltarget + " downloaded and saved to " + storeit +"."


#grab xls ranges
#combine in one sheet with campus name

# deletes the downloaded excel files
#for campus in csu:
    #storeit = d + campus + ext
    #os.remove(storeit) 
#print "Files removed!"

#load to spss
#run script
#import win32com.client
arr = [123,456,789,123]
spssloadfromarray(arr)