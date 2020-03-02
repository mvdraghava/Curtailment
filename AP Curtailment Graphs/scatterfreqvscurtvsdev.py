import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import numpy as np
from matplotlib import rcParams

def freqinrange(row):
    if row['Frequency']>=49.95 and row['Frequency']<=50.05:
        return True
    else:
        return False

plt.style.use('seaborn')
df = pd.read_excel('freqvscurtailvsdev.xlsx')
freqinrangedf = df[df.apply(freqinrange,axis=1)]
freqlessband = df[df['Frequency'] < 50.0]
frqgreatband = df[df['Frequency'] > 50.05]
plt.scatter(freqinrangedf['Curtailment'],freqinrangedf['Deviation'],label='Frequency in range 49.95 to 50.05')
plt.scatter(freqlessband['Curtailment'],freqlessband['Deviation'],c = sns.xkcd_rgb["pale red"],label='Frequency Less than 50.05')
plt.scatter(frqgreatband['Curtailment'],frqgreatband['Deviation'],c = sns.xkcd_rgb["pale orange"],label='Frequency Greater than 50.05')
plt.xlabel("Curtailment in MW")
plt.ylabel("Deviation in MW")
plt.legend(loc='upper left')
plt.title('Frequency Vs Curtailment Vs Deviation',fontdict = {'fontsize': 24,
 'fontweight' : rcParams['axes.titleweight'],
 'verticalalignment': 'baseline',
 'horizontalalignment': 'center'})
plt.show()
import pdb
pdb.set_trace()
