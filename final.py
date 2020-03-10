#coding=utf8
import itchat
from itchat.content import *
import os
import time, datetime
from openpyxl import  Workbook 
from openpyxl  import load_workbook
from openpyxl.styles import numbers


folder = './wechatmsg'
path_prefix = folder+'/'
chatroom_ids = []


#xlsx文件存储位置
def mkdir(path):
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        return True
    else:
        return False

#按日期登记，xlsx文件初始化
def mkxls(date):
    filename = date+'.xlsx'
    path = path_prefix + filename
    
    isExists = os.path.exists(path)
    if not isExists:
        # 实例化
        wb = Workbook()
        # 激活 worksheet
        ws = wb.active
        
        ws.append(['微信昵称','客户名称', '产品名称', '金额','期数','提交时间'])
        col_d = ws.column_dimensions['D']
        col_d.number_format = numbers.FORMAT_NUMBER_00
        col_e = ws.column_dimensions['E']
        col_e.number_format = numbers.FORMAT_NUMBER_00
        col_f = ws.column_dimensions['F']
        col_f.number_format = numbers.FORMAT_DATE_DATETIME
        wb.save(path)
        return True
    else:
        return False


#是否为有效数字
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
    return False
 

#是否为有效消息：含有四个信息（'客户名称', '产品名称', '金额','期数'）
#金额，期数字段为有效数字
def verify_group_msg(contentlist):
    if len(contentlist) !=4 :
        return False
    else:
        if not is_number(contentlist[2]) :
            return False
        else:
            if not is_number(contentlist[3]) :
                return False
            else:
                return True
#提示信息回复
def reply(content):
    itchat.send_msg(content, chatroom_ids[0])
        
#保存群消息
def save_group_msg(msg):
    msg_id = msg['MsgId']
    msg_from_user = msg['ActualNickName']
    msg_content = msg['Content']
    msg_create_time = msg['CreateTime']
    msg_type = msg['Type']
    
    timeArray = time.localtime(msg_create_time)
    otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S", timeArray)

    contentlist = msg_content.split()

    if not verify_group_msg(contentlist):
        reply('登记失败')
        return
    else:
        date = time.strftime("%Y%m%d", timeArray)
        mkxls(date)
        filename = date+'.xlsx'
        path = path_prefix + filename
        wb = load_workbook(path)
        ws = wb.active
        ws.append([msg_from_user,contentlist[0],contentlist[1],contentlist[2],contentlist[3],otherStyleTime])
        #ws.append(['微信昵称','客户名称', '产品名称', '金额','几期','提交时间'])
        wb.save(path)
        reply('登记成功')
    return 


#注册对文本消息进行监听，对群聊进行监听
@itchat.msg_register(itchat.content.TEXT, isGroupChat=True) 
def print_content(msg):
    
    msg_from_user = msg['ActualNickName']
    msg_content = msg['Content']
    msg_chatroom_id = msg['User']['UserName']
    print("群聊信息: ",msg_from_user, msg_content)
    if not msg_chatroom_id in chatroom_ids:
        return
    save_group_msg(msg)

   

# 轮询任务
def start_schedule():
    sched.add_job(logout, 'interval', minutes=1)
    sched.start()


def lc():
    print('login')
    print (time.strftime('%H:%M:%S',time.localtime(time.time())))
def ec():
    print('exit')
    print (time.strftime('%H:%M:%S',time.localtime(time.time())))


if __name__=="__main__":
    itchat.auto_login(hotReload=True,enableCmdQR=False,loginCallback=lc, exitCallback=ec)
    mkdir(folder)
    chat_rooms = itchat.search_chatrooms(name='test')
    if len(chat_rooms) > 0:
        chatroom_ids.append(chat_rooms[0]['UserName'])
    itchat.run()


