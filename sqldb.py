#数据库操作
import sqlite3
import re


def newdb(year, month):
    # 打开或创建数据库文件
    conn = sqlite3.connect('sichangyun.db')

    # #创建表
    # #获取游标
    c = conn.cursor()
    # #执行sql
    # #primary 主键
    # 判断表是否存在
    table_check = "SELECT name FROM sqlite_master WHERE type='table';"
    c.execute(table_check)
    # 返回所有表名[[('cq_2021_1',), ('cq_2021_2',), ('cq_2021_3',), ('cq_2021_4',), ('cq_2021_5',), ('cq_2021_6',), ('cq_2021_7',), ('cq_2021_8',)]]
    tables = [c.fetchall()]
    # 返回所有的表名["'cq_2021_1'", "'cq_2021_2'", "'cq_2021_3'", "'cq_2021_4'", "'cq_2021_5'", "'cq_2021_6'", "'cq_2021_7'", "'cq_2021_8'"]
    table_list = re.findall('(\'.*?\')', str(tables))
    # 去掉双引号
    table_list = [re.sub("'", '', each) for each in table_list]
    for m in range(month):
        #新建重庆表
        if ('cq_' + str(year) + '_' + str(m + 1)) not in table_list:
           #如果不存在则创建表
            sql_newcq = '''
                create table cq_{}
                    (id int primary key not null,
                    daysal int not null,
                    daymoney float not null,
                    totalmoney float not null,
                    yuebao int not null,
                    duoribao int not null,
                    shengjibao int not null,
                    ribao int not null,
                    yuyinbao int not null,
                    
                    faqidinggou int not null,
                    dinggou int not null,
                    yuebao_money float not null,
                    duoribao_money float not null,
                    shengjibao_money float not null,
                    ribao_money float not null,
                    yuyinbao_money float not null
                    );
            '''.format(str(year) + '_' + str(m + 1))
            c.execute(sql_newcq)
            conn.commit()
        #新建河北表
        if ('hb_' + str(year) + '_' + str(m + 1)) not in table_list:
            #如果不存在则创建表
            sql_newhb = '''
                create table hb_{} (
                    id int primary key not null,
                    daysal int not null,
                    daymoney float not null,
                    totalmoney float not null,
                    tehuiyuebao int not null,
                    duoribao int not null,
                    yuyinbao int not null,
                    quanyibao int not null,
                    shengjibao int not null,
                    jiasubao int not null,
                    yuebao int not null,
                    xianshibao int not null,
                    
                    faqidinggou int not null,
                    erciqueren int not null,
                    dinggou int not null,
                    
                    yuebaomoney float not null,
                    tehuimoney float not null,
                    ribaomoney float not null,
                    xianshibaomoney float not null,
                    quanyibaomoney float not null,
                    jiasubaomoney float not null,
                    shengjibaomoney float not null,
                    yuyinbaomoney float not null
                );
            '''.format(str(year) + '_' + str(m + 1))
            c.execute(sql_newhb)
            conn.commit()

        # 新建山西表
        if ('sx_' + str(year) + '_' + str(m + 1)) not in table_list:
            # 如果不存在则创建表
            sql_newsx = '''
                    create table sx_{} (
                        id int primary key not null,
                        daysal int not null,
                        daymoney float not null,
                        totalmoney float not null,
                        duoribao int not null,
                        yuebao int not null,
                        yuyinbao int not null,
                        shengjibao int not null,

                        faqidinggou int not null,
                        erciqueren int not null,
                        dinggou int not null,
                        
                        duoribaomoney float not null,
                        yuebaomoney float not null,
                        yuyinbaomoney float not null,
                        shengjibaomoney float not null,
                        shengdaibaomoney float not null
                    );
                '''.format(str(year) + '_' + str(m + 1))
            c.execute(sql_newsx)
            conn.commit()

        # 新建黑龙江表
        if ('hlj_' + str(year) + '_' + str(m + 1)) not in table_list:
            # 如果不存在则创建表
            sql_newhlj = '''
                    create table hlj_{} (
                        id int primary key not null,
                        daysal int not null,
                        daymoney float not null,
                        totalmoney float not null,
                        duoribao int not null,
                        shengdaibao int not null,
                        chaohui int not null,
                        yuyinbao int not null,
                        yuebao int not null,
                        jiasubao int not null,
                        zunxiangbao int not null,

                        faqidinggou int not null,
                        erciqueren int not null,
                        dinggou int not null,
                        duoribaomoney float not null,
                        bingjilingmoney float not null,
                        chaohuimoney float not null,
                        yuyinbaomoney float not null,
                        yuebaomoney float not null,
                        jiasubaomoney float not null,
                        zunxiangbaomoney float not null
                    );
                '''.format(str(year) + '_' + str(m + 1))
            c.execute(sql_newhlj)
            conn.commit()
    c.close()
    # #提交数据库操作，关闭数据库链接
    conn.close()


#插入重庆数据
def insertcq(year, month, day_sale, day_money, total_money, yuebao_sale, duoribao_sale, shengjibao_sale, ribao_sale, yuyinbao_sale, day_faqi, day_dinggou,yuebao_money, duoribao_money,shengjibao_money,ribao_money,yuyinbao_money):
    connection = sqlite3.connect('sichangyun.db')
    cur = connection.cursor()
    newdb(year, month)
    j = 0
    #插入数据
    for m in range(month):
        # print('重庆数据：'+ str(m))
        for i in range(31):
            insert_cq_sql = '''
                insert into {} (id,daysal,daymoney,totalmoney,yuebao,duoribao,shengjibao,ribao,yuyinbao,faqidinggou,dinggou,yuebao_money, duoribao_money,shengjibao_money,ribao_money,yuyinbao_money)
                values ({},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{});
            '''.format('cq_'+str(year)+'_'+str(m+1), i+1, day_sale[j*31+i], day_money[j*31+i], total_money[j*31+i], yuebao_sale[j*31+i], duoribao_sale[j*31+i],
                       shengjibao_sale[j*31+i], ribao_sale[j*31+i], yuyinbao_sale[j*31+i], day_faqi[j*31+i], day_dinggou[j*31+i],yuebao_money[j*31+i],
                       duoribao_money[j*31+i],shengjibao_money[j*31+i],ribao_money[j*31+i],yuyinbao_money[j*31+i])
            cur.execute(insert_cq_sql)
            connection.commit()
        j += 1
    cur.close()
    connection.close()


#插入河北数据
def inserthb(year, month, day_sale, day_money, total_money, tehuiyuebao_sale, duoribao_sale, yuyinbao_sale, quanyibao_sale, shengjibao_sale, jiasubao_sale,
             yuebao_sale, xianshibao_sale, day_faqi,day_erciqueren, day_dinggou,yuebao_money,tehui_money,ribao_money,xianshibao_money,quanyibao_money,
             jiasubao_money,shengjibao_money,yuyinbao_money):
    connection = sqlite3.connect('sichangyun.db')
    cur = connection.cursor()
    newdb(year, month)
    j = 0
    # 插入数据
    for m in range(month):
        # print('河北数据：' + str(m))
        for i in range(31):
            insert_hb_sql = '''
                    insert into {} (id,daysal,daymoney,totalmoney,tehuiyuebao,duoribao,yuyinbao,quanyibao,shengjibao,jiasubao,yuebao,xianshibao,faqidinggou,erciqueren,dinggou,
                    yuebaomoney,tehuimoney,ribaomoney,xianshibaomoney,quanyibaomoney,jiasubaomoney,shengjibaomoney,yuyinbaomoney)
                    values ({},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{});
                '''.format('hb_' + str(year) + '_' + str(m + 1), i + 1, day_sale[j * 31 + i], day_money[j * 31 + i],
                           total_money[j * 31 + i], tehuiyuebao_sale[j * 31 + i], duoribao_sale[j * 31 + i],yuyinbao_sale[j * 31 + i],
                           quanyibao_sale[j * 31 + i], shengjibao_sale[j * 31 + i], jiasubao_sale[j * 31 + i], yuebao_sale[j * 31 + i],
                           xianshibao_sale[j * 31 + i], day_faqi[j * 31 + i],day_erciqueren[j * 31 + i], day_dinggou[j * 31 + i],yuebao_money[j * 31 + i],
                           tehui_money[j * 31 + i],ribao_money[j * 31 + i],xianshibao_money[j * 31 + i],quanyibao_money[j * 31 + i],jiasubao_money[j * 31 + i],
                           shengjibao_money[j * 31 + i],yuyinbao_money[j * 31 + i])
            cur.execute(insert_hb_sql)
            connection.commit()
        j += 1
    cur.close()
    connection.close()


#插入山西数据
def insertsx(year, month, day_sale, day_money, total_money, duoribao_sale, yuebao_sale, yuyinbao_sale,shengjibao_sale,day_faqi,day_erciqueren, day_dinggou,
             duoribao_money, yuebao_money, yuyinbao_money, shengjibao_money, shengdaibao_money):
    connection = sqlite3.connect('sichangyun.db')
    cur = connection.cursor()
    newdb(year, month)
    j = 0
    # 插入数据
    for m in range(month):
        # print('河北数据：' + str(m))
        for i in range(31):
            insert_sx_sql = '''
                    insert into {} (id,daysal,daymoney,totalmoney,duoribao,yuebao,yuyinbao,shengjibao,faqidinggou,erciqueren,dinggou,duoribaomoney, yuebaomoney, yuyinbaomoney, shengjibaomoney, shengdaibaomoney)
                    values ({},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{});
                '''.format('sx_' + str(year) + '_' + str(m + 1), i + 1, day_sale[j * 31 + i], day_money[j * 31 + i],
                           total_money[j * 31 + i], duoribao_sale[j * 31 + i],yuebao_sale[j * 31 + i], yuyinbao_sale[j * 31 + i],
                           shengjibao_sale[j * 31 + i], day_faqi[j * 31 + i],day_erciqueren[j * 31 + i], day_dinggou[j * 31 + i],
                           duoribao_money[j * 31 + i], yuebao_money[j * 31 + i], yuyinbao_money[j * 31 + i], shengjibao_money[j * 31 + i], shengdaibao_money[j * 31 + i])
            cur.execute(insert_sx_sql)
            connection.commit()
        j += 1
    cur.close()
    connection.close()


#插入黑龙江数据
def inserthlj(year, month, day_sale, day_money, total_money, duoribao_sale, shengdaibao_sale, chaohui_sale, yuyinbao_sale,yuebao_sale,
                    jiasubao_sale, zunxiangbao_sale,day_faqi,day_erciqueren, day_dinggou,duoribao_money,bingjiling_money,chaohui_money,yuyinbao_money,
                    yuebao_money,jiasubao_money,zunxiangbao_money):
    connection = sqlite3.connect('sichangyun.db')
    cur = connection.cursor()
    newdb(year, month)
    j = 0
    # 插入数据
    for m in range(month):
        # print('河北数据：' + str(m))
        for i in range(31):
            insert_hlj_sql = '''
                    insert into {} (id,daysal,daymoney,totalmoney,duoribao,shengdaibao,chaohui,yuyinbao,yuebao,jiasubao,zunxiangbao,faqidinggou,erciqueren,dinggou,duoribaomoney,bingjilingmoney,chaohuimoney,yuyinbaomoney,yuebaomoney,jiasubaomoney,zunxiangbaomoney)
                    values ({},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{});
                '''.format('hlj_' + str(year) + '_' + str(m + 1), i + 1, day_sale[j * 31 + i], day_money[j * 31 + i],
                           total_money[j * 31 + i], duoribao_sale[j * 31 + i],shengdaibao_sale[j * 31 + i],chaohui_sale[j * 31 + i],yuyinbao_sale[j * 31 + i],
                           yuebao_sale[j * 31 + i],jiasubao_sale[j * 31 + i], zunxiangbao_sale[j * 31 + i],day_faqi[j * 31 + i],day_erciqueren[j * 31 + i], day_dinggou[j * 31 + i],
                           duoribao_money[j * 31 + i], bingjiling_money[j * 31 + i], chaohui_money[j * 31 + i], yuyinbao_money[j * 31 + i],yuebao_money[j * 31 + i], jiasubao_money[j * 31 + i], zunxiangbao_money[j * 31 + i])
            cur.execute(insert_hlj_sql)
            connection.commit()
        j += 1
    cur.close()
    connection.close()


#删除表
def deldb(year, month):
    connection = sqlite3.connect('sichangyun.db')
    cur = connection.cursor()
    for m in range(month):
        if month == 1:
            year -= 1
        # 判断表是否存在
        table_check = "SELECT name FROM sqlite_master WHERE type='table';"
        cur.execute(table_check)
        # 返回所有表名[[('cq_2021_1',), ('cq_2021_2',), ('cq_2021_3',), ('cq_2021_4',), ('cq_2021_5',), ('cq_2021_6',), ('cq_2021_7',), ('cq_2021_8',)]]
        tables = [cur.fetchall()]
        # 返回所有的表名["'cq_2021_1'", "'cq_2021_2'", "'cq_2021_3'", "'cq_2021_4'", "'cq_2021_5'", "'cq_2021_6'", "'cq_2021_7'", "'cq_2021_8'"]
        table_list = re.findall('(\'.*?\')', str(tables))
        # 去掉双引号
        table_list = [re.sub("'", '', each) for each in table_list]
        sql = ''
        if ('cq_' + str(year) + '_' + str(m + 1)) in table_list:
            sql_delcq = 'drop table cq_{}_{};'.format(year, m + 1)
            cur.execute(sql_delcq)
            connection.commit()
        if ('hb_' + str(year) + '_' + str(m + 1)) in table_list:
            sql_delhb = 'drop table hb_{}_{};'.format(year, m + 1)
            cur.execute(sql_delhb)
            connection.commit()
        if ('hlj_' + str(year) + '_' + str(m + 1)) in table_list:
            sql_delhlj = 'drop table hlj_{}_{};'.format(year, m + 1)
            cur.execute(sql_delhlj)
            connection.commit()
        if ('sx_' + str(year) + '_' + str(m + 1)) in table_list:
            sql_delsx = 'drop table sx_{}_{};'.format(year, m + 1)
            cur.execute(sql_delsx)
            connection.commit()
    cur.close()
    connection.close()