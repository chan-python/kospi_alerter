# 1분 봉으로 체크
import win32com.client
import datetime
import time
import schedule
import telegram

percent = 1  # default 1%
rapid_percent = 2 # default 2%
alert_once_percent = 1 # default 1%
counter_min = [0] * 2 # 연속 카운트 방지 변수
latest_before_value = [0] * 2 # 코인 중 마지막으로 발생한 알람 시점의 퍼센트를 저장하는 변수
# 0 : KOSPI
# 1 : KOSDAQ
stock = {'KOSPI': 0, 'KOSDAQ': 1}
stock_no = ('KOSPI', 'KOSDAQ')
alert_moment_default = [[60, 0, 0], [30, 0, 0], [15, 0, 0]] # 경과시간 (60, 30, 15), 당시금액 혹은 지수, 당시~현재 변동률
alert_moment = [alert_moment_default] * len(stock)
telgm_default_list = [] 

# 초기 시작시간 기록
now = datetime.datetime.now()
nowtime_start = int(now.strftime('%Y%m%d%H%M'))

# 텔레그램의 chat id를 추가하는 telegram_chat_id_add 함수 선언
def telegram_chat_id_add(updates, telgm_default_list, telgm_extra):
    telgm_list = []
    for telgm_i in updates:
        try:
            telgm_list.append(telgm_i['message']['chat']['id'])
        except Exception as e:
            telgm_list.append(telgm_i['channel_post']['chat']['id'])
    telgm_list.extend(telgm_default_list)
    telgm_list = list(set(telgm_list))
    return telgm_list

telgm_token = '1272415618:AAHJ2b-Fz701zzazQCwDPMS8IGCKWoSuECI'
telgm_bot = telegram.Bot(token=telgm_token)
telgm_updates = telgm_bot.get_updates()
telgm_list = telegram_chat_id_add(telgm_updates, telgm_default_list, [])

def telgm_message(alert_message):
    global telgm_list
    telgm_i = 0
    while telgm_i < len(telgm_list):
        try:
            telgm_bot.sendMessage(chat_id=telgm_list[telgm_i], text=alert_message)
            telgm_i += 1
        except:
            telgm_i += 1

# 현재가격을 체크한 후 알람을 발생하는 함수
def current_percent(term, stock_type, name, updown, value, alert_on, price_received):
    price = '%.4f' % price_received
    price = str(price)

    if alert_on == True:
        updown = ""
    if updown == "+":
        alert_message = str(term) + "분 기준 " + name + " 상승, 퍼센트는 " + str('%.4f' % value)\
                        + " %, " + str(price) + " 입니다."
        telgm_message(alert_message)
        return 30
    if updown == "-":
        alert_message = str(term) + "분 기준 " + name + " 하락, 퍼센트는 " + str('%.4f' % value)\
                        + " %, " + str(price) + " 입니다."
        telgm_message(alert_message)
        return 30
    if updown == "r+":
        alert_message = str(term) + "분 기준 " + name + " 빠른 상승, 퍼센트는 " + str('%.4f' % value)\
                        + " %, " + str(price) + " 입니다."
        telgm_message(alert_message)
        return 30
    if updown == "r-":
        alert_message = str(term) + "분 기준 " + name + " 빠른 하락, 퍼센트는 " + str('%.4f' % value)\
                        + " %, " + str(price) + " 입니다."
        telgm_message(alert_message)
        return 30

    if name == stock_no[0]:
        return counter_min[0]
    elif name == stock_no[1]:
        return counter_min[1]

    return counter_min

# 알람이 발생한 이후 일정 퍼센트 이상 추가 변동이 있었는지 체크하는 함수
def check_alert_once(stock_type, latest_value):
    try:
        if abs(latest_before_value[stock_type] - latest_value) < alert_once_percent:
            return True
        else:
            return False
    except:
        return False

# 알림이 발생할 시 값을 리턴하는 함수
def check_latest_percent(latest_value, stock_type, alert_on):
    if alert_on == False:
        return latest_value
    else:
        return latest_before_value[stock_type]


# 연결 여부 체크
objCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()


def current_stock():
    current = []
    try:
        now = datetime.datetime.now()
        nowtime = now.strftime('%H%M')
        objStockChart = win32com.client.Dispatch('CpSysDib.StockChart')
        objStockChart.SetInputValue(0, 'U001')  # 종목코드
        objStockChart.SetInputValue(5, [0, nowtime, 2, 3, 4, 5, 8])  # 요청항목 - 날짜, 시간,시가,고가,저가,종가,거래량
        objStockChart.SetInputValue(7, 1)  # 분틱차트 주기  #10은 10분봉
        objStockChart.BlockRequest()
        #print("KOSPI : ", objStockChart.GetHeaderValue(7), sep='')
        current.append(float(objStockChart.GetHeaderValue(7)))

        objStockChart.SetInputValue(0, 'U201')  # 종목코드
        objStockChart.SetInputValue(5, [0, nowtime, 2, 3, 4, 5, 8])  # 요청항목 - 날짜, 시간,시가,고가,저가,종가,거래량
        objStockChart.SetInputValue(7, 1)  # 분틱차트 주기  #10은 10분봉
        objStockChart.BlockRequest()
        #print("KOSDAQ : ", objStockChart.GetHeaderValue(7), sep='')
        current.append(float(objStockChart.GetHeaderValue(7)))
    except:
        pass
    for c in current:
        if c == 0:
            return False
    return current

current = []
current = current_stock()

# 코인 가격 60분 - [0] ~ [59] 까지 현재 시세로 초기화
history_kospi_default = []
history_kospi_default.append(current[stock['KOSPI']])
history_kospi_default *= 60
history_kosdaq_default = []
history_kosdaq_default.append(current[stock['KOSDAQ']])
history_kosdaq_default *= 60
history = []
history.append(history_kospi_default)
history.append(history_kosdaq_default)

def nowtime_check():
    now = datetime.datetime.now()
    nowtime = int(now.strftime('%H%M'))
    nowtime_start_check = int(now.strftime('%Y%m%d%H%M'))
    return nowtime, nowtime_start_check

# 스케쥴러 함수
def job():
    global latest_before_value
    global counter_min
    global alert_on
    global stock, stock_no, history
    global nowtime_start
    global current

    alert_on = [False, False]

    nowtime, nowtime_start_check = nowtime_check()
    #print(nowtime_start_check, nowtime_start, (nowtime_start_check - nowtime_start), sep=' / ')
    
    # 장 종료 타임에는 Sleep 호출
    if nowtime > 1530 and nowtime <= 2400: # 15시 30분에 KOSPI, KOSDAQ 제공 종료
        print('15:30~24:00 - 슬립진입', nowtime, nowtime_start_check,sep=', ')
        time.sleep(3600)
    elif nowtime > 0 and nowtime <= 800:
        print('00:00~08:00 - 슬립진입', nowtime, nowtime_start_check, sep=', ')
        time.sleep(3600)
    elif nowtime_start_check - nowtime_start > 1\
            and history[0][0] == history[0][30] and history[0][0] == history[0][59]:
        print('동일한 값이 60분 동안 지속 - 슬립진입', nowtime, nowtime_start_check, sep=', ')
        #print(history[0][0], history[0][30], history[0][59])
        #print(history)
        nowtime_start = nowtime_start_check
        time.sleep(3600)
    else:
        # 현재 값을 호출해오는 함수
        current = current_stock()

    # 60분, 30분, 15분 별로 값 입력, 변동률 계산
    if current == False:
        pass
    else:
        try:
            for scode in stock_no:
                alert_moment[stock[scode]][0][1] = history[stock[scode]][60 - alert_moment[stock[scode]][0][0]]
                alert_moment[stock[scode]][1][1] = history[stock[scode]][60 - alert_moment[stock[scode]][1][0]]
                alert_moment[stock[scode]][2][1] = history[stock[scode]][60 - alert_moment[stock[scode]][2][0]]

                alert_moment[stock[scode]][0][2] = \
                    float(
                        (current[stock[scode]] - float(alert_moment[stock[scode]][0][1])) / current[stock[scode]] * 100)
                alert_moment[stock[scode]][1][2] = \
                    float(
                        (current[stock[scode]] - float(alert_moment[stock[scode]][1][1])) / current[stock[scode]] * 100)
                alert_moment[stock[scode]][2][2] = \
                    float(
                        (current[stock[scode]] - float(alert_moment[stock[scode]][2][1])) / current[stock[scode]] * 100)
        except:
            pass

    # KOSPI 변화율 체크후 알람 발생
    for scode in stock_no:
        if alert_moment[stock[scode]][2][2] > rapid_percent and alert_on[stock[scode]] == False:
            alert_on[stock[scode]] = check_alert_once(stock[scode], alert_moment[stock[scode]][2][2])
            counter_min[stock[scode]] = current_percent(alert_moment[stock[scode]][2][0], stock[scode], scode, "r+",
                                                        alert_moment[stock[scode]][2][2], alert_on[stock[scode]],
                                                        current[stock[scode]])
            latest_before_value[stock[scode]] = check_latest_percent(alert_moment[stock[scode]][2][2], stock[scode],
                                                                     alert_on[stock[scode]])
            alert_on[stock[scode]] = True
        if alert_moment[stock[scode]][2][2] < (-rapid_percent) and alert_on[stock[scode]] == False:
            alert_on[stock[scode]] = check_alert_once(0, alert_moment[stock[scode]][2][2])
            counter_min[stock[scode]] = current_percent(alert_moment[stock[scode]][2][0], stock[scode], scode, "r-",
                                                        alert_moment[stock[scode]][2][2], alert_on[stock[scode]],
                                                        current[stock[scode]])
            latest_before_value[stock[scode]] = check_latest_percent(alert_moment[stock[scode]][2][2], stock[scode],
                                                                     alert_on[stock[scode]])
            alert_on[stock[scode]] = True
        if alert_moment[stock[scode]][1][2] > percent and alert_on[stock[scode]] == False:
            alert_on[stock[scode]] = check_alert_once(0, alert_moment[stock[scode]][1][2])
            counter_min[stock[scode]] = current_percent(alert_moment[stock[scode]][1][0], stock[scode], scode, "+",
                                                        alert_moment[stock[scode]][1][2], alert_on[stock[scode]],
                                                        current[stock[scode]])
            latest_before_value[stock[scode]] = check_latest_percent(alert_moment[stock[scode]][1][2], stock[scode],
                                                                     alert_on[stock[scode]])
            alert_on[stock[scode]] = True
        if alert_moment[stock[scode]][1][2] < (-percent) and alert_on[stock[scode]] == False:
            alert_on[stock[scode]] = check_alert_once(0, alert_moment[stock[scode]][1][2])
            counter_min[stock[scode]] = current_percent(alert_moment[stock[scode]][1][0], stock[scode], scode, "-",
                                                        alert_moment[stock[scode]][1][2], alert_on[stock[scode]],
                                                        current[stock[scode]])
            latest_before_value[stock[scode]] = check_latest_percent(alert_moment[stock[scode]][1][2], stock[scode],
                                                                     alert_on[stock[scode]])
            alert_on[stock[scode]] = True
        if alert_moment[stock[scode]][0][2] > percent and alert_on[stock[scode]] == False:
            alert_on[stock[scode]] = check_alert_once(0, alert_moment[stock[scode]][0][2])
            counter_min[stock[scode]] = current_percent(alert_moment[stock[scode]][0][0], stock[scode], scode, "+",
                                                        alert_moment[stock[scode]][0][2], alert_on[stock[scode]],
                                                        current[stock[scode]])
            latest_before_value[stock[scode]] = check_latest_percent(alert_moment[stock[scode]][0][2], stock[scode],
                                                                     alert_on[stock[scode]])
            alert_on[stock[scode]] = True
        if alert_moment[stock[scode]][0][2] < (-percent) and alert_on[stock[scode]] == False:
            alert_on[stock[scode]] = check_alert_once(0, alert_moment[stock[scode]][0][2])
            counter_min[stock[scode]] = current_percent(alert_moment[stock[scode]][0][0], stock[scode], scode, "-",
                                                        alert_moment[stock[scode]][0][2], alert_on[stock[scode]],
                                                        current[stock[scode]])
            latest_before_value[stock[scode]] = check_latest_percent(alert_moment[stock[scode]][0][2], stock[scode],
                                                                     alert_on[stock[scode]])
            alert_on[stock[scode]] = True

    # 1분이 지날때마다 카운터를 1분씩 줄이며, 알림 울린 뒤 30분이 지났을 시 마지막 가격을 초기화
    for coin_no in range(0, len(counter_min)):
        if counter_min[coin_no] != 0:
            counter_min[coin_no] -= 1
        elif counter_min[coin_no] == 0:
            latest_before_value[coin_no] = 0

    now = datetime.datetime.now()
    now_view = ('%04d/%02d/%02d %02d:%02d' % (now.year, now.month, now.day, now.hour, now.minute))
    myview_message = 'KOSPI : ' + str(current[stock['KOSPI']])\
                     + ' ' + '%.4f' %float(alert_moment[stock['KOSPI']][0][2])\
                     + ' ' + '%.4f' %float(alert_moment[stock['KOSPI']][1][2])\
                     + ' ' + '%.4f' %float(alert_moment[stock['KOSPI']][2][2])\
                     + ' ' + '%.4f' %float(latest_before_value[stock['KOSPI']])\
                     + ' ' + str(counter_min[stock['KOSPI']]) + ' '\
                     + 'KOSDAQ : ' + str(current[stock['KOSDAQ']])\
                     + ' ' + '%.4f' %float(alert_moment[stock['KOSDAQ']][0][2])\
                     + ' ' + '%.4f' %float(alert_moment[stock['KOSDAQ']][1][2])\
                     + ' ' + '%.4f' %float(alert_moment[stock['KOSDAQ']][2][2])\
                     + ' ' + '%.4f' %float(latest_before_value[stock['KOSDAQ']])\
                     + ' ' + str(counter_min[stock['KOSDAQ']])
    print(now_view, myview_message, sep=', ')
schedule.every(1).minutes.do(job)

while True:
    schedule.run_pending()
    time.sleep(1)
