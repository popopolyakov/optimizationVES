import matplotlib.pyplot as plt
import numpy as np
import pyexcel

book = pyexcel.get_book(file_name="VES_Dataset.xlsx")

data = np.array(book['Мощности'].column[8][2:1010])
times = np.array(book['Мощности'].column[6][2:1010])

for i in range(data.size):
    if data[i] < 0:
        data[i] = 0

numberOfIn = 5
# T = times[-1] - times[0]
T = 10
pRated = 1200
weights = np.zeros(2 * numberOfIn + 1)
weights[-1] = data[0]
omg = 1 / T
alf = 0.1
lumb = 0
power=0
outData = np.zeros(data.size)
outData2 = np.zeros(data.size)
x_ar = np.zeros(2 * numberOfIn + 1)
testData = np.zeros(data.size)
BESS_values = np.zeros(data.size)
pfmc = []
pH = np.zeros(data.size)
pL = np.zeros(data.size)

coup = 200
coup_limit = 500  # Вопрос 250 чего?? Это сделано чтобы срезать пики
coup_ar = np.zeros(data.size)


kSafeFluctuation = 0.1  # коэф безопасной скорости зарядки

wp1 = np.zeros(3)


BESS_status = True  # True зарядка False разрядка
BESS_value = coup/coup_limit # Уровень зарядки

prev_BESS_status = BESS_status
for x in range(data.size):
    t = times[x]
    for k in range(numberOfIn):
        x_ar[2 * k] = np.cos(t * k * omg)
        x_ar[2 * k + 1] = np.sin(t * k * omg)

    x_ar[2 * numberOfIn] = 1
    f_k = np.dot(weights, x_ar)

    weights = weights + alf * (data[x] - f_k) * x_ar / (np.dot(x_ar, x_ar) + lumb)

    outData[x] = np.dot(weights, x_ar)

    if len(pfmc) < 4:
        pfmc.append(outData[x])
    else:
        pfmc.pop(0)
        pfmc.append(outData[x])

    pH30 = min(pfmc) + pRated * kSafeFluctuation
    pL30 = max(pfmc) - pRated * kSafeFluctuation

    pH[x] = pH30
    pL[x] = pL30
    limitPower=False
    if outData[x] <= pL30:
        outData[x] = pL30
        limitPower=True
    if outData[x] >= pH30:
        outData[x] = pH30
        limitPower=True
    
    # if x % 6 == 0:
    #     coup = 0
    #     outData[x] = data[x]
    #
    # else:
    #     coup += (f_k - data[x]) * T / 60
    oldCoup = coup
    oldOutData = outData[x]
    if BESS_status == True:
        
        if BESS_value <= 0.2:
            coup += (outData[x] - data[x]) * T / 60

        elif BESS_value > 0.2 and BESS_value <= 0.6:
            coup += (outData[x] - data[x]) * T / 60

        elif BESS_value > 0.6 and BESS_value <= 1:
            plus_coup = (1 - (1/0.6)*BESS_value) * ((outData[x] - data[x]) * T / 60)
            newOutData = outData[x] - plus_coup * 60 / T
            if not (newOutData > pH30 or newOutData < pL30) and newOutData>0:
                coup += plus_coup
                outData[x] = newOutData
        else:
            pass
            #raise IOError('BESS value > 1 | < 0 при зарядке')





    else:
        if BESS_value <= 0.2:
            plus_coup = (1 - 5 * BESS_value) * (outData[x] - data[x]) * T / 60
            
            newOutData = outData[x] - plus_coup * 60 / T
            if not (newOutData > pH30 or newOutData < pL30) and newOutData>0:
                coup += plus_coup
                outData[x] = newOutData

        elif BESS_value > 0.2 and BESS_value <= 0.6:
            coup += (outData[x] - data[x]) * T / 60

        elif BESS_value > 0.6 and BESS_value <= 1:
            coup += (outData[x] - data[x]) * T / 60

        else:
            pass
            #raise IOError('BESS value > 1 | < 0 при разрядке')
        
    
    # if BESS_value > 1:
    #     BESS_value = 1
    #     BESS_status = False
    saveBatteryLimit = False
    print(coup, coup_ar[x - 1])
    prev_BESS_status = BESS_status
    if coup <= coup_ar[x-1] and BESS_status == True:
        BESS_status = False
    if coup >= coup_ar[x-1] and BESS_status == False:
        BESS_status = True





    # if (x > 2) and not limitPower and not saveBatteryLimit and outData[x]>0:
    #     deltaFluctuation = max(abs(coup - coup_ar[x - 3]), abs(coup - coup_ar[x - 2]), abs(coup - coup_ar[x - 1]))
    #     deltaCurFluctuation = abs(coup - coup_ar[x - 1])
    #     if deltaFluctuation > kSafeFluctuation * coup_limit:
    #         if (prev_BESS_status == True):
    #             curDeltaCoup = deltaCurFluctuation + (kSafeFluctuation * coup_limit/3)
    #             newOutData = outData[x] + curDeltaCoup * 60 / T
    #             if not (newOutData > pH30 or newOutData < pL30) and newOutData>0:
    #                 coup += curDeltaCoup
    #                 outData[x] = newOutData
    #         else:
    #             curDeltaCoup = deltaCurFluctuation + (kSafeFluctuation * coup_limit/3)
    #             newOutData = outData[x] - curDeltaCoup * 60 / T
    #             if not (newOutData > pH30 or newOutData < pL30) and newOutData>0:
    #                 coup += curDeltaCoup
    #                 outData[x] = newOutData
    if coup < 0 or coup > coup_limit:
        print('coup<0')
        coup = oldCoup
        outData[x] = oldOutData
    BESS_value = coup / coup_limit
    
    if (BESS_value > 1 or BESS_value < 0):
        print(coup, 'coup', coup_limit, 'coup limit', BESS_value, 'BESS value', prev_BESS_status, 'prev BESS STATUS')
        #raise IOError('BESS жопа')
    coup_ar[x] = coup
    

    # if (outData[x] != outData2[x]):
        # print(BESS_status, BESS_value, coup, outData[x], outData2[x], outData[x]-outData2[x])
    BESS_values[x] = BESS_value * 100
    power += (outData[x]-data[x])*T/60
    outData2[x] = power

# plt.plot(times, outData, label='coup')
# plt.plot(times, outData2, label='coup')
# plt.plot(times, data - outData)
# plt.plot(times, pL)
# plt.plot(times, pH)

# plt.plot(times, testData)
fig, ax = plt.subplots()
ax.plot(times, outData, label='с батареей')
ax.plot(times, outData2, label='суммарная мощность')
#ax.plot(times, data, label='исходные данные')
ax.legend()
plt.show()
fig, ax = plt.subplots()
ax.plot(times, coup_ar, label='coup')
ax.plot(times, BESS_values, label='BESS')

ax.legend()

plt.show()

# sheet = pyexcel.Sheet([[times[i], outData[i], data[i]] for i in range(outData.size)])
# sheet.save_as("output.csv")
