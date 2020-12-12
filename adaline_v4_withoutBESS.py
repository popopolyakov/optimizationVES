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
T = 10
pRated = 1500
weights = np.zeros(2 * numberOfIn + 1)
weights[-1] = data[0]
omg = 1 / T
alf = 0.3
lumb = 0
outData = np.zeros(data.size)
outData2 = np.zeros(data.size)
x_ar = np.zeros(2 * numberOfIn + 1)
testData = np.zeros(data.size)
pfmc = []
pH = np.zeros(data.size)
pL = np.zeros(data.size)

coup_ar = np.zeros(data.size)

coup_size = 300  # емкость BESS

kSafeFluctuation10 = 0.1  # коэф безопасной скорости зарядки

SOC = 0.5  # Уровень зарядки
SOCmin = 0.2
SOCmax = 0.8

for x in range(data.size):
    t = times[x]
    for k in range(numberOfIn):
        x_ar[2 * k] = np.cos(t * k * omg)
        x_ar[2 * k + 1] = np.sin(t * k * omg)

    x_ar[2 * numberOfIn] = 1
    f_k = np.dot(weights, x_ar)

    weights = weights + alf * (data[x] - f_k) * x_ar / (np.dot(x_ar, x_ar) + lumb)

    Pco = np.dot(weights, x_ar)

    # SCFC 1 and 2
    if len(pfmc) < 4:
        pfmc.append(Pco)
    else:
        pfmc.pop(0)
        pfmc.append(Pco)

    pH30 = min(pfmc) + pRated * kSafeFluctuation10
    pL30 = max(pfmc) - pRated * kSafeFluctuation10

    if Pco <= pL30:
        Pco = pL30
    if Pco >= pH30:
        Pco = pH30

    outData2[x] = Pco

    Peref = data[x] - Pco
    SOCpred = (coup_size * SOC + Peref * T / 60) / coup_size

    if SOCmin < SOCpred < SOCmax:
        SOC = SOCpred
        outData[x] = Pco
    elif SOC <= SOCmax and SOCpred > SOCmax:
        SOC = SOCmax
        outData[x] = data[x] + (SOC - SOCmax) * coup_size / T * 60
    elif SOC >= SOCmin and SOCpred < SOCmin:
        SOC = SOCmin
        outData[x] = data[x] + (SOC - SOCmin) * coup_size / T * 60
    else:
        outData[x] = data[x]

    
    coup_ar[x] = SOC
    pH[x] = pH30
    pL[x] = pL30


plt.plot(times, outData)
#plt.plot(times, data)
#plt.plot(times, outData2)
# plt.plot(times, data - outData)
#plt.plot(times, coup_ar * 1000)
plt.plot(times, pL)
plt.plot(times, pH)

# plt.plot(times, testData)
plt.show()
# fig, ax = plt.subplots()
# ax.plot(times, coup_ar, label='coup')
# ax.plot(times, BESS_values, label='BESS')
#
# ax.legend()
plt.plot(times, coup_ar)
plt.show()

# sheet = pyexcel.Sheet([[times[i], outData[i], data[i]] for i in range(outData.size)])
# sheet.save_as("output.csv")
