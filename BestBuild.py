# This script finds the best character-vehicle combination for a Mario Kart World race in 150cc with normal items.
# It achieves its objective by maximising its expected speed during a race while fulfilling a minimum acceleration requirement.
# The evaluation of the expected speed takes the effects of coins during races in terms of the probabilities of the numbers of coins the player is expected to possess and the combination's responsiveness to the coins.

import numpy as np
import openpyxl as pyxl
import matplotlib.pyplot as plt
import Functions as f
import colorsys

## Initiate and preallocate

file=pyxl.load_workbook("Mario Kart World stats.xlsx",data_only=True)
# loads the excel file
characters=file["Characters"]
vehicles=file["Vehicles"]
weightandcoins=file["Weight & Coins"]
# separates the spreadsheets in the excel files

nc=20
# 20 unique character classes
nv=24
# 24 unique vehicle stat configurations
nB=nc*nv
# total number of possible unique combinations
iB=0
# preallocates the considered combination counter

AccelLvlLTL=12
# the minimum acceleration level requirement for a combination to be considered

comboname=[[0]*1 for i2 in range(0,nB,1)]
speedlvl_solid=np.zeros(nB)
speedlvl_grainy=np.zeros(nB)
speedlvl_water=np.zeros(nB)
speedincrease_solid=np.zeros((nB,21))
speedincrease_grainy=np.zeros((nB,21))
speedincrease_water=np.zeros((nB,21))
speedincrease_railairorwall=np.zeros((nB,21))
speedincrease_overall=np.zeros((nB,21))
speedincrease_expected=np.zeros(nB)
accelerationlevel=np.zeros(nB)
weight=np.zeros(nB)
handling_solid=np.zeros(nB)
handling_grainy=np.zeros(nB)
handling_water=np.zeros(nB)
# preallocates arrays to store the considered combinations

## Analyse the stats

# Chooses whether to use Queueing Theory probabilities (True) or from raw data (False)
if True:
    P=f.CoinCountProbabilities(True)
    # loads the state probabilities of the number of coins the player posseses at any point during a race
    # The boolean input indicates whether the focus is on 3-lap races (True) or 6-lap knockout tours (False)
else:
    file2=pyxl.load_workbook("Mario Kart World real time coin possession data.xlsx",data_only=True)
    spreadsheet=file2["shortcat"]
    # loads the raw data spreadsheet
    P=np.zeros(21)
    # preallocates the probability matrix
    for i in range(0,21,1):
        P[i]=spreadsheet.cell(row=2+i,column=7).value
        # imports the probabilities

for ic in range(0,nc,1):
    for iv in range(0,nv,1):
        # for every character-vehicle combination
        accelerationlevel[iB]=characters.cell(row=3+ic,column=5).value+vehicles.cell(row=3+iv,column=5).value
        comboname[iB]=f"{characters.cell(row=3+ic,column=1).value} with the {vehicles.cell(row=3+iv,column=1).value}"
        speedlvl_solid[iB]=characters.cell(row=3+ic,column=2).value+vehicles.cell(row=3+iv,column=2).value
        speedlvl_grainy[iB]=characters.cell(row=3+ic,column=3).value+vehicles.cell(row=3+iv,column=3).value
        speedlvl_water[iB]=characters.cell(row=3+ic,column=4).value+vehicles.cell(row=3+iv,column=4).value
        weight[iB]=characters.cell(row=3+ic,column=6).value+vehicles.cell(row=3+iv,column=6).value
        handling_solid[iB]=characters.cell(row=3+ic,column=7).value+vehicles.cell(row=3+iv,column=7).value
        handling_grainy[iB]=characters.cell(row=3+ic,column=8).value+vehicles.cell(row=3+iv,column=8).value
        handling_water[iB]=characters.cell(row=3+ic,column=9).value+vehicles.cell(row=3+iv,column=9).value
        # records all relevant stat levels by the summation of character and vehicle stat levels
        for i in range(0,21,1):
            speedincrease_solid[iB,i]=(1+weightandcoins.cell(row=2+int(speedlvl_solid[iB]),column=2).value)*(1+weightandcoins.cell(row=2+int(weight[iB]),column=5+i).value)-1
            speedincrease_grainy[iB,i]=(1+weightandcoins.cell(row=2+int(speedlvl_grainy[iB]),column=2).value)*(1+weightandcoins.cell(row=2+int(weight[iB]),column=5+i).value)-1
            speedincrease_water[iB,i]=(1+weightandcoins.cell(row=2+int(speedlvl_water[iB]),column=2).value)*(1+weightandcoins.cell(row=2+int(weight[iB]),column=5+i).value)-1
            speedincrease_railairorwall[iB,i]=(1+weightandcoins.cell(row=15,column=2).value)*(1+weightandcoins.cell(row=2+int(weight[iB]),column=5+i).value)-1
            # recalls the speed increases for each terrain at every coin count according to the speed level and weight
        speedincrease_overall[iB,:]=(0.5188*speedincrease_solid[iB,:]+0.2131*speedincrease_grainy[iB,:]+0.0747*speedincrease_water[iB,:]+0.1934*speedincrease_railairorwall[iB,:])
        speedincrease_expected[iB]=np.dot(speedincrease_overall[iB,:],P).item()
        # evaluates the overall speed increase by the proportions of solid, grainy, and water terrains
        # treats the number of coins possessed and the terrain on which the player is driving as independent events
        iB+=1
        # counts and saves the considered combination

## Find the best combination and viable alternatives

jB0=np.where(accelerationlevel>=AccelLvlLTL)
speedincrease_expected_filtered=speedincrease_expected[[jB0]]
# Filters for combinations that meet or exceed the acceleration level requirement

jB1=jB0[0][speedincrease_expected_filtered.argmax()]
bestcombo=comboname[jB1]
# finds the best combination
print(f"Best combination: {bestcombo}")
print(f"Expected speed: +{np.round(100*speedincrease_expected[jB1],3)}% from baseline")
print(f"Speed level: {int(speedlvl_solid[jB1])}-{int(speedlvl_grainy[jB1])}-{int(speedlvl_water[jB1])}")
print(f"Acceleration level: {int(accelerationlevel[jB1])}")
print(f"Weight: {int(weight[jB1])}")
print(f"Handling: {int(handling_solid[jB1])}-{int(handling_grainy[jB1])}-{int(handling_water[jB1])}")
# prints the details of the best combination

plt.rcParams.update({'font.size': 10})
plt.rcParams['figure.constrained_layout.use']=True

jB=np.where(speedincrease_expected>=0.98*np.max(speedincrease_expected_filtered))
jB=np.intersect1d(jB0,jB)
# finds the indices of all viable alternatives (within 2% of the maximum relative expected speed increase)
jB=np.delete(jB,np.where(jB==jB1))
# prvents double-counting the best combination

plt.figure(1)
plt.plot(np.arange(0,21,1),100*speedincrease_overall[jB1,:],color=colorsys.hsv_to_rgb(0,1,1),marker=".",markersize=6,linestyle='',label=bestcombo)
# plots the speed-coincount curve for the best combination

for i in range(0,len(jB),1):
    i2=i+1
    print("")
    print(f"Viable alternative {i+1}: {comboname[jB[i]]}")
    print(f"Expected speed: +{np.round(100*speedincrease_expected[jB[i]],3)}% from baseline")
    print(f"Speed level: {int(speedlvl_solid[jB[i]])}-{int(speedlvl_grainy[jB[i]])}-{int(speedlvl_water[jB[i]])}")
    print(f"Acceleration level: {int(accelerationlevel[jB[i]])}")
    print(f"Weight: {int(weight[jB[i]])}")
    print(f"Handling: {int(handling_solid[jB[i]])}-{int(handling_grainy[jB[i]])}-{int(handling_water[jB[i]])}")
    # prints the details of each viable alternative
    plt.plot(np.arange(0,21,1),100*speedincrease_overall[jB[i],:],color=colorsys.hsv_to_rgb(28/36*(i+1)/(len(jB)),1,1),marker=".",markersize=(i2/len(jB)*3+(len(jB)-i2)/len(jB)*6),linestyle='',label=comboname[jB[i]])
    # plots the speed-coincount curve for each viable alternative

## find how a particular chosen combination compares (optinal; toggle by boolean)

if True:
    IC=10 # Mario
    IV=3 # Baby Blooper
    IB=IC*nv+IV
    s1=np.sort(speedincrease_expected_filtered[0][0])
    s1=s1[-1::-1]
    rank=np.array(np.where(s1==speedincrease_expected[IB])).flatten()[0]+1
    print("")
    print(f"User-selected combination: {comboname[IB]}")
    print(f"Expected speed: +{np.round(100*speedincrease_expected[IB],3)}% from baseline")
    print(f"Speed level: {int(speedlvl_solid[IB])}-{int(speedlvl_grainy[IB])}-{int(speedlvl_water[IB])}")
    print(f"Acceleration level: {int(accelerationlevel[IB])}")
    print(f"Weight: {int(weight[IB])}")
    print(f"Handling: {int(handling_solid[IB])}-{int(handling_grainy[IB])}-{int(handling_water[IB])}")
    if rank==1:
        print(f"This combination is the best")
    if rank==2:
        print(f"This combination is the 2nd best")
    if rank==3:
        print(f"This combination is the 3rd best")
    elif rank%10>=4 or rank%10==0 or (rank%100>=10 and rank%100<=20):
        print(f"This combination is the {rank}th best")
    elif rank%10==3:
        print(f"This combination is the {rank}rd best")
    elif rank%10==2:
        print(f"This combination is the {rank}nd best")
    elif rank%10==1:
        print(f"This combination is the {rank}st best")
    # prints the details of the selected combination
    plt.plot(np.arange(0,21,1),100*speedincrease_overall[IB,:],'kx',label=comboname[IB])
    # plots the speed-coincount curve for the selected combination

## Dress up the graph

plt.xlabel("Number of coins in possession")
plt.ylabel("Expected speed increase relative to baseline (%)")
plt.xlim(0,20)
plt.ylim(2,11)
plt.xticks(np.arange(0,21,1))
plt.yticks(np.arange(2,12,1))
plt.legend(loc="upper right")
plt.grid()

plt.figure(2)
plt.bar(np.arange(0,21,1),100*P[:,0])
plt.xlabel("Number of coins in possession")
plt.ylabel("Coin possession probability (%)")
plt.xlim(0,20)
plt.ylim(0,10)
plt.xticks(np.arange(0,21,1))
plt.yticks(np.arange(0,11,1))

plt.grid()

if True:
    plt.show()
