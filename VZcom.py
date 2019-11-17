'''
Программа для передачи в COM-порт  автоинкриментных записей в формате WITS
Передача ведётся в порт с наименьшим номером, найденный в системе.
ОБЯЗАТЕЛЬНО переведите клавиатуру в режим английского языка!!!
краткая справка в периуд исполнения программы - нажмите h
код написан NykSu (c) нояюрь 2019.  v 0.1.0
GitHub NykSu
'''

import os
import time
import serial
import msvcrt

class KBHit:

    def __init__(self):
        if os.name == 'nt':
            pass

    def set_normal_term(self):
        if os.name == 'nt':
            pass

    def getch(self):
        s = ''
        if os.name == 'nt':
            return msvcrt.getch().decode('utf-8')

    def kbhit(self):
        if os.name == 'nt':
            return msvcrt.kbhit()


def get_WITS_date_time(): # формат WITS строки времени и даты
    dat_str = ''
    tim_str = ''
    tim = time.localtime(time.time())
    # создаём строку времени
    if len(str(tim[3])) == 1:
        tim_str = '0' + str(tim[3])
    else:
        tim_str = str(tim[3])
    if len(str(tim[4])) == 1:
        tim_str += '0' + str(tim[4])
    else:
        tim_str += str(tim[4])
    if len(str(tim[5])) == 1:
        tim_str += '0' + str(tim[5])
    else:
        tim_str += str(tim[5])
    # создаём строку даты
    if len(str(tim[0])) == 4:
        dat_str = str(tim[0])[2::]
    if len(str(tim[1])) == 1:
        dat_str += '0' + str(tim[1])
    else:
        dat_str += str(tim[1])
    if len(str(tim[2])) == 1:
        dat_str += '0' + str(tim[2])
    else:
        dat_str += str(tim[2])
    return ('0105' + dat_str, '0106' + tim_str)


def make_WITS_msg(record, sequence, deep, deep_d): # формирование пакета данных в формате WITS
    result = ['&&','0101Oil Hole 1','01020']
    result.extend(['0103' + str(record),'0104' + str(sequence)])
    result.extend(get_WITS_date_time())
    result.append('01070')
    result.append('0108' + str(round(deep_d, 2))) # глубина долота
    result.append('01090')
    result.append('0110' + str(round(deep, 2))) # глубина скважины
    result.extend(['01110','0112','0113','0114','0117','01410','0142','!!'])
    return tuple(result)


def push_to_com_port(num_port, data, end_str = '\r\n'): # отправка данных в COM-порт
    result = False 
    if num_port == 0:
        return result
    try :
        port = "COM%s" % num_port
        ser = serial.Serial(port)
        for st in data:
            ser.write((st + end_str).encode())
        ser.close()
        result = True
    except serial.serialutil.SerialException :
        pass
    return result


def tornado(ds, dd, d, de, p, record, sequence, num_port, chars = [], tm = 0): # главный цикл, ожидание коанд, отправка пакетов, расчёт инкриментов
    dp = 0
    kb = KBHit()
    pause = p
    print('Нажмите ESC для выхода из программы, для справки нажмите ---> h')

    start = time.time()
    if tm == 0:
        result = ['', ds, dd, start, sequence]
    else:
        result = ['', ds, dd, tm, sequence]
    
    while True:
        time.sleep(p/50)
        if de != 0 and ds >= de:
            result[1] = ds
            break
        end = time.time()
        if end - start >= pause:
            dd += d
            if dd > ds:
                ds = dd
            print('Глубина скважины и долота (м):', round(ds, 2), round(dd, 2), ' интервал (c): ', p, ' приращение глубины (м): ', d, ' Время: ', round(end - result[3], 2), '--> в порт COM%s' % num_port)
            dp = p - (end - start)
            pause += dp 
            start = time.time()
            if not push_to_com_port(num_port, make_WITS_msg(record, sequence, ds, dd)): # Передать значение в COM-порт (строка возможной корректировки)
                print('Ошибка записи в COM-порт!!!')
                break
            # print('Пакет успешно отправлен в порт COM%s' % num_port)
            sequence += 1
        if kb.kbhit():
            c = kb.getch()
            if ord(c) != 0:
                if ord(c) == 27 or c in chars:  # ESC ord(c) == 27
                    result[0] = c
                    result[1] = ds
                    result[2] = dd
                    result[4] = sequence
                    break
    kb.set_normal_term()
    return result


def print_help(): # вывод на экран справки
    print('Краткая справка. Команды клавиш: *, /, d, +, -, h')
    print('"*" - увеличение временного интервала в 2 раза')
    print('"/" - уменьшение временного интервала в 2 раза')
    print('"d" - измнение (корректировка) глубины')
    print('"+" - увеличение интервала на 2 сек')
    print('"-" - уменьшение интервала на 2 сек')
    print('"s" - остановка для подъёма долота. Ввод глубины долота')
    print('"h" - вызов этой справки')


if __name__ == "__main__":

    found = False 
    num_port = 0
    for i in range(256) :
        try :
            port = "COM%s" % (i + 1)
            ser = serial.Serial(port)
            ser.close()
            print("Найден последовательный порт: ", port)
            found = True
            if num_port == 0:
                num_port = i + 1
        except serial.serialutil.SerialException :
            pass
    if not found :
        print("Последовательных портов не обнаружено")
    if num_port > 0:
        print('Передача пакетов идёт на порт COM%s' % num_port)

    time_str = 0
    sequence = 1
    deep = float(input('Введите начальную глубину скважины (м): '))
    deep_d = float(input('Введите начальную глубину долота (м): '))
    delta = float(input('Введите шаг приращения глубины (м): '))
    deep_end = float(input('Введите конечную глубину (0 - нет ограничений)(м): '))
    pause = float(input('Введите интервал времени в секудах: '))
    record = int(input('Введите номер записи: '))
    print_help()
    while True:
        res = tornado(deep, deep_d, delta, deep_end, pause, record, sequence, num_port, ['*', '/', 'd', '+', '-', 's', 'h'], time_str)
        if type(res) != type([]): 
            print('Ошибка!!..')
            break
        if res[0] == '': # автоматическое завершение программы по достижению конечной глубины
            print('Достигнута глубина (м)', deep_end)
            break
        elif ord(res[0]) == 27: # команда выхода из прораммы
            print('Работа программы прервана пользователем.')
            break
        elif res[0] == '*': # команда двукратного увеличения интервала времени
            pause += pause
            deep = res[1]
            deep_d = res[2]
            time_str = res[3]
            sequence = res[4]
            print('Интервал увеличен до: ', pause, ' сек')
        elif res[0] == '/':  # команда двукратного уменьшения интервала времени
            pause = pause / 2
            deep = res[1]
            deep_d = res[2]
            time_str = res[3]
            sequence = res[4]
            print('Интервал уменьшен до: ', pause, ' сек')
        elif res[0] == '+': # команда увеличения интервала времени на 2
            pause += 2
            deep = res[1]
            deep_d = res[2]
            time_str = res[3]
            sequence = res[4]
            print('Интервал увеличен до: ', pause, ' сек')
        elif res[0] == '-': # команда уменьшения интервала времени на 2
            if pause > 2:
                pause -= 2
            deep = res[1]
            deep_d = res[2]
            time_str = res[3]
            sequence = res[4]
            print('Интервал уменьшен до: ', pause, ' сек')
        elif res[0] == 'h': # команда вывода справки
            deep = res[1]
            deep_d = res[2]
            time_str = res[3]
            sequence = res[4]
            print_help()
        elif res[0] == 'd': # команда приостановки и корректирования глубины скважины
            time_str = 0
            deep = float(input('Введите корректировочную глубину скважины(м): '))
            if deep < res[2]:
                res[2] = deep
            deep_d = res[2]
            sequence = res[4]
            print('Исправлена глубина. Новое значение глубины скважины (м): ', deep, ' глубина долота (м): ', deep_d)
            res[2] = float(input('Показания высоты подвеса долота верны? Введите 0 - верны или число для корректировки: '))
            if res[2] > 0:
                deep_d = res[2]
        elif res[0] == 's': # команда приостановки и корректирования глубины долота
            time_str = 0
            deep = res[1]
            sequence = res[4]
            deep_d = float(input('Пауза в программе. Остановка или подъём долота. Введите корректировочную глубину долота(м): '))
            print('Исправлена глубина долота. Новое значение глубины (м): ', deep_d)
            if deep_d > deep:
                print('Долото ниже последнего уровня глубины скважины. Глубина скважины скорректируется на уровень (м): ', deep_d)

'''
The Program is for transferring auto-recording records in WITS format to COM-port
Transfer is carried out to the port with the lowest number found in the system.
ALWAYS put the keyboard in English mode !!!
quick reference during program execution period - press h
This code was written by NykSu (c) November 2019. v 0.1.0
GitHub NykSu
'''