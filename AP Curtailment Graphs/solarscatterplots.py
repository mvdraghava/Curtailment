import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import numpy as np
from matplotlib import rcParams
from matplotlib.ticker import (MultipleLocator, FormatStrFormatter,
                               AutoMinorLocator)
from matplotlib.transforms import Bbox
from matplotlib.backends.backend_pdf import PdfPages

def freqinrange(row):
    if row['Frequency']>=49.9 and row['Frequency']<=50.05:
        return True
    else:
        return False

plt.style.use('seaborn')
df = pd.read_excel('solarcurtailment.xlsx')

with PdfPages('AP_Solar_Curtailment.pdf') as pdf:

#############  Plot with Frequency in range points and out range points #####################################################
    rcParams['figure.figsize'] = 18, 12
    freqinrangedf = df[df.apply(freqinrange,axis=1)]
    freqlessband = df[df['Frequency'] < 49.9]
    frqgreatband = df[df['Frequency'] > 50.05]
    fig, ax = plt.subplots()
    plt.scatter(freqinrangedf['Curtailment'],freqinrangedf['Deviation'],label='Frequency in range 49.95 to 50.05')
    plt.scatter(freqlessband['Curtailment'],freqlessband['Deviation'],c = sns.xkcd_rgb["pale red"],label='Grid Frequency Less than 49.9')
    plt.scatter(frqgreatband['Curtailment'],frqgreatband['Deviation'],c = sns.xkcd_rgb["pale orange"],label='Grid Frequency Greater than 50.05')
    ax.xaxis.set_major_locator(MultipleLocator(50))
    ax.xaxis.set_minor_locator(MultipleLocator(25))
    ax.xaxis.grid(True, which='minor')
    ax.yaxis.set_major_locator(MultipleLocator(100))
    ax.yaxis.set_minor_locator(MultipleLocator(50))
    ax.yaxis.grid(True, which='minor')
    plt.xlabel("Curtailment in MW")
    plt.ylabel("Deviation in MW")
    plt.legend(loc='upper left')
    plt.title('AP Wind Curtailment Vs Deviation',fontdict = {'fontsize': 24,
     'fontweight' : rcParams['axes.titleweight'],
     'verticalalignment': 'baseline',
     'horizontalalignment': 'center'})
    
    pdf.savefig(bbox_inches='tight')
    plt.close()
######################################################################################################################################
##
##
######################################## Plot with Frequency less than 50.0 and greater than 50.0 #######################################
##

    fig, ax = plt.subplots()
    frqgreatband = df[df['Frequency'] >= 50.0]
    frqlessband = df[df['Frequency'] < 50.0]
    plt.scatter(frqgreatband['Curtailment'],frqgreatband['Deviation'],label='Grid Frequency Greater than 50.0 Hz')
    plt.scatter(frqlessband['Curtailment'],frqlessband['Deviation'],c = sns.xkcd_rgb["pale red"],label='Grid Frequency Less than 50.0 Hz')


    ax.xaxis.set_major_locator(MultipleLocator(50))
    ax.xaxis.set_minor_locator(MultipleLocator(25))
    ax.xaxis.grid(True, which='minor')
    ax.yaxis.set_major_locator(MultipleLocator(100))
    ax.yaxis.set_minor_locator(MultipleLocator(50))
    ax.yaxis.grid(True, which='minor')
    plt.xlabel("Curtailment in MW")
    plt.ylabel("Deviation in MW")
    plt.legend(loc='upper left')
    plt.title('APWind Curtailment Vs AP ISTS Deviation',fontdict = {'fontsize': 24,
     'fontweight' : rcParams['axes.titleweight'],
     'verticalalignment': 'baseline',
     'horizontalalignment': 'center'})

    pdf.savefig(bbox_inches='tight')
    plt.close()

#######################################################################################################################################


###################################### Plot with Frequency and Deviation in x and y axis and curtailment as size of the marker #########################

    fig, ax = plt.subplots()

    plt.scatter(df['Frequency'],df['Deviation'],marker = '.', c = sns.xkcd_rgb["pale red"],s = df['Curtailment'],label = 'Size represents Wind Curtailment in MW')

    plt.xlabel("Frequency in Hz")
    plt.ylabel("Deviation in MW")
    plt.legend(loc='upper left')
    ax.xaxis.set_major_locator(MultipleLocator(.05))
    ax.xaxis.set_minor_locator(MultipleLocator(.025))
    ax.xaxis.grid(True, which='minor')
    ax.yaxis.set_major_locator(MultipleLocator(100))
    ax.yaxis.set_minor_locator(MultipleLocator(50))
    ax.yaxis.grid(True, which='minor')
    plt.title('Grid Frequency Vs AP ISTS Deviation at times of Wind Curtailment',fontdict = {'fontsize': 20,
     'fontweight' : rcParams['axes.titleweight'],
     'verticalalignment': 'baseline',
     'horizontalalignment': 'center'})
    pdf.savefig(bbox_inches='tight')
    plt.close()

########################################################################################################################################################

################################  Plot with Frequnecy vs Deviation at times of Wind Curtailment ###########################################

##plt.figure(5)
##plt.scatter(df['Frequency'],df['Deviation'])
##plt.plot([49.80,50.20],[0,0], c=sns.xkcd_rgb["black"])
##plt.plot([49.9 , 49.9],[-600 , 300], c=sns.xkcd_rgb["pale red"])
##plt.plot([50.05, 50.05],[-600 , 300], c=sns.xkcd_rgb["pale red"])
##plt.xlabel("Frequency in Hz")
##plt.ylabel("Deviation in MW")
##plt.title('Frequency Vs AP ISTS Deviation at times of AP Wind Curtailemnt',fontdict = {'fontsize': 24,
## 'fontweight' : rcParams['axes.titleweight'],
## 'verticalalignment': 'baseline',
## 'horizontalalignment': 'center'})
##plt.rcParams['axes.xmargin'] = 0
##plt.show()

###########################################################################################################################################


############################### Plot with Frequency vs Curtailment vs Deviation ###########################################################

    fig, ax = plt.subplots()
    underdrawalpoints = df[df['Deviation']<0]
    overdrawalpoints = df[df['Deviation']>0]
    plt.scatter(underdrawalpoints['Frequency'],underdrawalpoints['Curtailment'],marker = '*',label = 'Underdrawal Points')
    plt.scatter(overdrawalpoints['Frequency'],overdrawalpoints['Curtailment'],marker = ',',label = 'Overdrawal Points')
    ax.xaxis.set_major_locator(MultipleLocator(.05))
    ax.xaxis.set_minor_locator(MultipleLocator(.025))
    ax.xaxis.grid(True, which='minor')
    ax.yaxis.set_major_locator(MultipleLocator(50))
    ax.yaxis.set_minor_locator(MultipleLocator(25))
    ax.yaxis.grid(True, which='minor')
    plt.xlabel("Frequency in Hz")
    plt.ylabel("Curtailment in MW")
    plt.legend(loc='upper left')
    plt.title('Grid Frequency Vs AP Wind Curtailment',fontdict = {'fontsize': 24,
     'fontweight' : rcParams['axes.titleweight'],
     'verticalalignment': 'baseline',
     'horizontalalignment': 'center'})
    pdf.savefig(bbox_inches='tight')
    plt.close()

###########################################################################################################################################

################################# Plot freq vs curt size and marker -- deviation ###########################################################

    fig, ax = plt.subplots()
    underdrawalpoints = df[df['Deviation']<0]
    overdrawalpoints = df[df['Deviation']>0]
    plt.scatter(underdrawalpoints['Frequency'],underdrawalpoints['Curtailment'],marker = '*',s = underdrawalpoints['Deviation'].apply(abs),label = 'Underdrawal Points with size representing deviation magnitude')
    plt.scatter(overdrawalpoints['Frequency'],overdrawalpoints['Curtailment'],marker = ',',s = overdrawalpoints['Deviation'].apply(abs),label = 'Overdrawal Points with size representing deviation magnitude')
    ax.xaxis.set_major_locator(MultipleLocator(.05))
    ax.xaxis.set_minor_locator(MultipleLocator(.025))
    ax.xaxis.grid(True, which='minor')
    ax.yaxis.set_major_locator(MultipleLocator(50))
    ax.yaxis.set_minor_locator(MultipleLocator(25))
    ax.yaxis.grid(True, which='minor')
    plt.xlabel("Grid Frequency in Hz")
    plt.ylabel("Curtailment in MW")
    plt.legend(loc='upper left')
    plt.title('Grid Frequency Vs AP Wind Curtailemnt',fontdict = {'fontsize': 24,
     'fontweight' : rcParams['axes.titleweight'],
     'verticalalignment': 'baseline',
     'horizontalalignment': 'center'})
    pdf.savefig(bbox_inches='tight')
    plt.close()

##############################################################################################################################################
