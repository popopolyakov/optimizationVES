import matplotlib.pyplot as plt
import numpy as np
import pyexcel
import pandas as pd

#-----------------------------------------
#------КОНСТАНТЫ ВЛИЯЮЩИЕ НА РАБОТУ-------
#-----------------------------------------
numVES=2
alf = 0.2 #коэф сглаживания
coup_limit = 730  # Вопрос 250 чего?? Это сделано чтобы срезать пики
kSafeFluctuation = 0.1  # коэф безопасной скорости зарядки

lumb = 0
#-----------------------------------------

# for numVES in range(1):
book = pyexcel.get_book(file_name="VES_Dataset.xlsx")

data = np.array(book['Мощности'].column[numVES + 7][2:1010])
# data2 = np.array(book['Мощности'].column[1 + 7][2:1010])
# data3 = np.array(book['Мощности'].column[2 + 7][2:1010])
# data=data1+data2+data3
times = np.array(book['Мощности'].column[6][2:1010])

for i in range(data.size):
    if data[i] < 0:
        data[i] = 0

numberOfIn = 7
# T = times[-1] - times[0]
T = 10
pRated = np.amax(data)
weights = np.zeros(2 * numberOfIn + 1)
weights[-1] = data[0]
omg = 1 / T


power=0
outData = np.zeros(data.size) # массив под изменение батареей
outData2 = np.zeros(data.size) # массив исходного сглаженного графика
outData3 = np.zeros(data.size) 
x_ar = np.zeros(2 * numberOfIn + 1)
testData = np.zeros(data.size)
BESS_values = np.zeros(data.size)
pfmc = []
finalLimitsOfPower = []
finalLimitsOfPower2=[]
pH = np.zeros(data.size)
pL = np.zeros(data.size)
PL30real=np.zeros(data.size)
PH30real=np.zeros(data.size)




coup = coup_limit*0.6
coup_ar = np.zeros(data.size)




wp1 = np.zeros(3)


BESS_status = True  # True зарядка False разрядка
BESS_value = coup/coup_limit # Уровень зарядки
BESS_status_log = np.zeros(data.size)
coefBatteryLog=np.zeros(data.size)

prev_BESS_status = BESS_status
powerBattery = np.zeros(data.size)
plus_coup_ar=np.zeros(data.size)
teor_capacity=np.zeros(data.size)
for x in range(data.size):
    t = times[x]
    for k in range(numberOfIn):
        x_ar[2 * k] = np.cos(t * k * omg)
        x_ar[2 * k + 1] = np.sin(t * k * omg)

    x_ar[2 * numberOfIn] = 1
    f_k = np.dot(weights, x_ar)
    
    weights = weights + alf * (data[x] - f_k) * x_ar / (np.dot(x_ar, x_ar) + lumb)
    outData[x] = np.dot(weights, x_ar)
    
    

    if len(pfmc) < 3:
        pfmc.append(outData[x])
    else:
        pfmc.pop(0)
        pfmc.append(outData2[x-1])

    pH30 = min(pfmc) + pRated * kSafeFluctuation
    pL30 = max(pfmc) - pRated * kSafeFluctuation
    if pL30 < 0:
        pL30=0

    pH[x] = pH30
    pL[x] = pL30
    limitPower=False
    if outData[x] <= pL30:
        outData[x] = pL30
        limitPower=True
    if outData[x] >= pH30:
        outData[x] = pH30
        limitPower=True
    outData2[x] = outData[x]
    # if x % 6 == 0:
    #     coup = 0
    #     outData[x] = data[x]
    #
    # else:
    #     coup += (f_k - data[x]) * T / 60
    oldCoup = coup
    oldOutData = outData[x]
    coefOnBattery = 1
    
    if data[x] >= outData[x]:
        BESS_status = True
    else:
        BESS_status = False
        

    if x > 0:
        teor_capacity[x] = (outData2[x] - data[x]) * T / 60 + teor_capacity[x - 1]
    else:
        teor_capacity[x] = (outData2[x] - data[x]) * T / 60


    if BESS_status == True:
        if BESS_value > 0 and BESS_value <= 0.6:
            coefOnBattery=1
        elif BESS_value > 0.6 and BESS_value <= 1:
            coefOnBattery=1 - (1/0.4)*(BESS_value-0.6)
        else:
            pass
            # raise IOError('BESS value > 1 | < 0 при зарядке')
    
    

    else:
        if BESS_value <= 0.4:
            coefOnBattery = 1 - (1/0.4) * (0.4-BESS_value)
            # print(BESS_value, coefOnBattery)
        elif BESS_value > 0.4 and BESS_value <= 1:
            coefOnBattery=1
        else:
            pass
            # raise IOError('BESS value > 1 | < 0 при разрядке')
    
    
        

    
    plus_coup = coefOnBattery * (data[x] - outData[x]) * T / 60
    
    coefBatteryLog[x]=coefOnBattery*100
    newOutData = data[x] - plus_coup * 60 / T
    

    # if BESS_value > 1:
    #     BESS_value = 1
    #     BESS_status = False

    # print(coup, coup_ar[x - 1])
    
    
    

    # if (x > 2) and not limitPower and  outData[x]>0:
    #     deltaFluctuation = max(abs(coup - coup_ar[x - 3]), abs(coup - coup_ar[x - 2]), abs(coup - coup_ar[x - 1]))
    #     deltaCurFluctuation = coup - coup_ar[x - 1]
    #     if abs(deltaFluctuation) > kSafeFluctuation * coup_limit:
    #         curDeltaCoup = deltaCurFluctuation + (kSafeFluctuation * coup_limit/3)
    #         newOutData = data[x] - curDeltaCoup * 60 / T
    #         if not (newOutData > pH30 or newOutData < pL30):
    #             print(times[x], ' ПРЕДЕЛ ПРИ ОГРАНИЧЕНИИ КОЛЕБАНИЙ ЕМКОСТИ')
    #             plus_coup = (data[x] - outData[x]) * T / 60
    #         else:
    #             print(times[x], ' ПРЕДЕЛ БЕЗ НИХ')
    #             plus_coup = curDeltaCoup
    
    # if times[x] > 11.65 and times[x] < 11.73:
    #     print(newOutData, pH[x], newOutData >= pH[x], pL[x])
        
    if newOutData >= pH[x]:
        # print('pH30 limit')
        outData[x] = pH[x]
    elif newOutData <= pL[x]:
        # print('pL30 limit')
        outData[x] = pL[x]
    else:
        outData[x] = newOutData
    plus_coup = (data[x] - outData[x]) * T / 60    
    coup += plus_coup
    
    # if times[x] > 6.115 and times[x] < 6.120:
    #     print(oldCoup, plus_coup)
    # powerBattery[x]=plus_coup*60/T
    
    if coup < 0 or coup > coup_limit:
        # print('coup<0')
        coup = oldCoup
        BESS_value = coup / coup_limit
        powerBattery[x]=0
        outData[x] = data[x]
    
    

    # if (BESS_value > 1 or BESS_value < 0):
    #     BESS_value = coup / coup_limit
    #     print(coup, 'coup', coup_limit, 'coup limit', BESS_value, 'BESS value', prev_BESS_status, 'prev BESS STATUS')
    #     #raise IOError('BESS жопа')
    
    # if (outData[x] != outData2[x]):
        # print(BESS_status, BESS_value, coup, outData[x], outData2[x], outData[x]-outData2[x])
   
    # power += (outData[x]-data[x])*T/60
    plus_coup_ar[x] = plus_coup

    if len(finalLimitsOfPower) < 3:
        finalLimitsOfPower.append(outData[x])
    else:
        finalLimitsOfPower.pop(0)
        finalLimitsOfPower.append(outData[x - 1])
    testLimits = (max(finalLimitsOfPower) - min(finalLimitsOfPower)) / pRated * 100
    PL30real[x] = max(finalLimitsOfPower) - pRated * kSafeFluctuation
    PH30real[x] = min(finalLimitsOfPower) + pRated * kSafeFluctuation
    
    if outData[x] > PH30real[x]:
        print('pH30real limit')
        if testLimits < 15:
            deltaHigh=data[x] - PH30real[x]
            outData[x] = PH30real[x]
            coup = oldCoup + deltaHigh * T / 60
            if coup < 0 or coup > coup_limit:
                print('coup<0')
                coup = oldCoup
                powerBattery[x]=0
                outData[x] = data[x]
                # raise IOError('Будет неоптимально. Выбери другую емкость батареи или просто поколдуй как то еще')
            
            BESS_value = coup / coup_limit
        else:
            raise IOError('Будет неоптимально. Выбери другую емкость батареи или просто поколдуй как то еще')
            # pass
        
    elif outData[x] < PL30real[x]:
        print('pL30real limit')
        
        if testLimits < 15:
            deltaLow=data[x] - PL30real[x]
            outData[x] = PL30real[x]
            coup = oldCoup + deltaLow * T / 60
            if coup < 0 or coup > coup_limit:
                print('coup<0')
                coup = oldCoup
                powerBattery[x]=0
                outData[x] = data[x]
                raise IOError('Будет неоптимально. Выбери другую емкость батареи или просто поколдуй как то еще')
            else:
                pass
            BESS_value = coup / coup_limit
        else:
            raise IOError('Будет неоптимально. Выбери другую емкость батареи или просто поколдуй как то еще')
            # pass  
            
    else:
        pass

    if len(finalLimitsOfPower2) < 3:
        finalLimitsOfPower2.append(outData[x])
    else:
        finalLimitsOfPower2.pop(0)
        finalLimitsOfPower2.append(outData[x - 1])
    testLimits = (max(finalLimitsOfPower2) - min(finalLimitsOfPower2)) / pRated * 100
    if testLimits > 10:
        print(pH[x], pL[x], outData[x], plus_coup, times[x], coup)
    coup_ar[x] = coup
    outData3[x] = testLimits
    BESS_value = coup / coup_limit
    BESS_values[x] = BESS_value * 100
    BESS_status_log[x] = BESS_status * 100
    

fig, ax = plt.subplots()

# plt.plot(times, outData2, label='Power without BESS')
# plt.plot(times, data - outData)
plt.plot(times, data, label='Wind Power', color=(0/255,91/255,149/255, 0.5), linewidth = 1)
plt.plot(times, outData, label='Power with BESS')
plt.plot(times, outData3, label='30min changes')
# plt.plot(times, BESS_status_log, label='BESS Status', color='pink', linewidth=1)
# plt.plot(times, coefBatteryLog, label='BESS coef log', color='grey', linewidth = 1)
#plt.plot(times, pL, label='low', color='green', linewidth = 1)
#plt.plot(times, pH, label="high", color='red', linewidth=1)
#plt.plot(times, PL30real, label='low real', color='green', linewidth = 1)
#plt.plot(times, PH30real, label="high real",color='red', linewidth = 1)


# plt.plot(times, testData)
# ax.plot(times, outData, label='с батареей')
# ax.plot(times, data, label='исходные данные')

# ax.plot(times, coup_ar, label='coup')
# ax.plot(times, teor_capacity, label='теоретически требуемая емкость')
ax.plot(times, BESS_values, label='BESS')


ax.legend()

plt.show()



resultForAnalitycs= {
        'numberOfLimitCoup': len(list(filter(lambda x: (80 < x and x < 100) or (0 < x and x < 20), BESS_values))),
        'qualityPower': len(list(filter(lambda x: x > 10.000000000000005, outData3))),
        'alfa': alf,
        'limitCoup': coup_limit,
        'VES': 'все сразу',
        'capacityVeight': coup_limit
    }
print(resultForAnalitycs)
# finalLimitsOfPower=[]
                        

# minCountOfLimit = resultForAnalitycs[0]['numberOfLimitCoup']
# minObjectLimit={}
# minCountOfQuality = resultForAnalitycs[0]['qualityPower']
# minObjectQuality={}
# for item in resultForAnalitycs:
#     if item['numberOfLimitCoup'] < minCountOfLimit:
#         minCountOfLimit = item['numberOfLimitCoup']
#         minObjectLimit = item
#     if item['qualityPower'] < minCountOfQuality:
#         minCountOfQuality = item['qualityPower']
#         minObjectQuality=item
    
    
# print('Количество 20 - 80% превышений', str(minCountOfLimit), str(minObjectLimit), 'для ВЭС № всех сразу')
# print('Количество выходов из стандартов', str(minCountOfQuality), str(minObjectQuality), 'для ВЭС №')
# resultForAnalitycs=[]

# sheet = pyexcel.Sheet([[times[i], outData[i], data[i]] for i in range(outData.size)])
# sheet.save_as("VEU" + str(numVES + 1) + "__alfa_" + str(alf) + "__capacity_" + str(coup_limit) + ".xsl")

df1 = pd.DataFrame({'with BESS': outData,'without BESS': data, 'quality Power': outData3, 'BESS status': BESS_status_log, 'BESS value': BESS_values, 'BESS capacitive': coup_ar})


# Create a Pandas Excel writer using XlsxWriter as the engine.
fileNameVESstr='VES'+str(numVES+1)+'.xlsx'
writer = pd.ExcelWriter(fileNameVESstr, # pylint: disable=abstract-class-instantiated
                  date_format='YYYY-MM-DD',
                  datetime_format='YYYY-MM-DD HH:MM:SS') 
# Write each dataframe to a different worksheet.
df1.to_excel(writer)
writer.save()



