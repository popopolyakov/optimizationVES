import matplotlib.pyplot as plt
import numpy as np
import pyexcel


def dis_func(arr, times, t):
    for i in range(len(times)):
        if times[i] == t:
            return arr[i]
        if times[i] > t:
            a = (arr[i - 1] - arr[i]) / (times[i - 1] - times[i])
            b = (arr[i - 1] * times[i] - arr[i] * times[i - 1]) / (times[i - 1] - 1)
            return a * t + b


book = pyexcel.get_book(file_name="VES_Dataset.xlsx")

data = np.array(book['Мощности'].column[9][2:1010])
times = np.array(book['Мощности'].column[6][2:1010])

for i in range(data.size):
    if data[i] < 0:
        data[i] = 0

numberOfIn = 5
# T = times[-1] - times[0]
T = 5
weights = np.zeros(2 * numberOfIn + 1)
weights[-1] = data[0]
omg = 1 / T
alf = 0.3
lumb = 0
outData = np.zeros(data.size)
x_ar = np.zeros(2 * numberOfIn + 1)
testData = np.zeros(data.size)
BESS_values = np.zeros(data.size)

coup = 30
coup_ar = np.zeros(data.size)

coup_limit = 200  #Вопрос 50 чего?? Это сделано чтобы срезать пики
kSafeFluctuation=0.1 # коэф безопасной скорости зарядки

wp1 = np.zeros(3)

BESS_limit=1500000 # МВт*ч изначальная версия предела зарядки
BESS_status=True # True зарядка False разрядка
BESS_value=0.3 #Уровень зарядки

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
    # if x % 6 == 0:
    #     coup = 0
    #     outData[x] = data[x]
    #
    # else:
    #     coup += (f_k - data[x]) * T / 60


    if BESS_status == True:
        if BESS_value < 0.2:
            coup += (outData[x] - data[x]) * T / 60

        elif BESS_value > 0.2 and BESS_value < 0.6:
            coup += (outData[x] - data[x]) * T / 60

        elif BESS_value > 0.6 and BESS_value <= 1:
            plus_coup=(1-BESS_value)*((outData[x] - data[x]) * T / 60)
            coup += plus_coup
            outData[x] = outData[x] - plus_coup*60/T
        else:
            raise IOError('BESS value > 1 | < 0 при зарядке')
        
        BESS_value=abs(coup)/coup_limit
        if coup <= coup_ar[x-1]:
            prev_BESS_status=BESS_status
            BESS_status = False
    else:
        if BESS_value < 0.2:
            plus_coup=(1-5*BESS_value) * (outData[x] - data[x]) * T / 60
            coup += plus_coup
            outData[x] = outData[x] - plus_coup * 60 / T
            
        elif BESS_value > 0.2 and BESS_value < 0.6:
            coup += (outData[x] - data[x]) * T / 60

        elif BESS_value > 0.6 and BESS_value <= 1:
            coup += (outData[x] - data[x]) * T / 60

        else:
            raise IOError('BESS value > 1 | < 0 при разрядке')
        
        
        if coup >= coup_ar[x-1]:
            prev_BESS_status=BESS_status
            BESS_status = True
    print(BESS_value)
    # if BESS_value > 1:
    #     BESS_value = 1
    #     BESS_status = False
    
    if (x>2):
        deltaFluctuation = max(abs(coup - coup_ar[x - 3]), abs(coup - coup_ar[x - 2]), abs(coup - coup_ar[x - 1]))
        deltaCurFluctuation = abs(coup - coup_ar[x - 1])
        if deltaFluctuation > kSafeFluctuation * coup_limit:
            if (prev_BESS_status == True):
                coup = coup - deltaCurFluctuation + (kSafeFluctuation * coup_limit/3) # во время зарядки вычесть разницу и прибавить допустимое значение
                outData[x] = outData[x] + (deltaCurFluctuation - (kSafeFluctuation * coup_limit/3)) * 60 / T # эта строчка под вопросом я не умею считать но вроде по уму
            else:
                coup = coup + deltaCurFluctuation - (kSafeFluctuation * coup_limit/3) # наоборот во время разрядки
                outData[x] = outData[x] - (deltaCurFluctuation - (kSafeFluctuation * coup_limit/3)) * 60 / T # эта строчка под вопросом я не умею считать но вроде по уму
    
    if coup < 0:
        delta_coup=coup
        coup = 0
        outData[x] = outData[x] + delta_coup * 60 / T



    coup_ar[x] = coup
    if (coup_ar[x] < 0):
        raise IOError('coup < 0')
    BESS_value = coup / coup_limit
    
    
    if (BESS_value > 1 or BESS_value < 0):
        print(coup, 'coup', coup_limit, 'coup limit', BESS_value, 'BESS value', prev_BESS_status, 'prev BESS STATUS')
        raise IOError('BESS жопа')
    BESS_values[x] = BESS_value*100



plt.plot(times, outData)
plt.plot(times, data)
plt.plot(times, data - outData)

#plt.plot(times, testData)
plt.show()
fig, ax = plt.subplots()
ax.plot(times, coup_ar, label='coup')
ax.plot(times, BESS_values, label='BESS')

ax.legend()

plt.show()

# sheet = pyexcel.Sheet([[times[i], outData[i], data[i]] for i in range(outData.size)])
# sheet.save_as("output.csv")
