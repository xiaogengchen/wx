# -*- coding: UTF-8 -*- 
  
import os
import sys
import codecs
import sqlite3
  
#set default encoding as UTF-8 
reload(sys)
sys.setdefaultencoding('utf-8')
#param list for select,insert,delete,update
paramDict = {} 
#state for every action's state.show in statebar
state = ''

#connect_database 
#if db exists,connect it 
#else create db file 
def connect_db(db_name):
    conn = sqlite3.connect(db_name)
    return conn
  
#close connect  
def close_db(conn):
    conn.close()
  
#insert values into archives table 
def insert_values(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("insert into %s values (%s,'%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
                   % (paramDict.get('table_name'),
                      paramDict.get('archiveID'),
                      paramDict.get('menlei'),
                      paramDict.get('guidang'),
                      paramDict.get('qixian'),
                      paramDict.get('anjuantiming'),
                      paramDict.get('danwei'),
                      paramDict.get('lijuanriqi'),
                      paramDict.get('weizhi'),
                      paramDict.get('miji'),
                      paramDict.get('zerenren'),
                      paramDict.get('quhao'),
                      paramDict.get('guihao'),
                      paramDict.get('hehao'),
                      paramDict.get('juanhao'),
                      paramDict.get('hujianhao'),
                      paramDict.get('kemu'),
                      paramDict.get('beizhu'),
                      paramDict.get('inputter'),
                      u'否',
                      paramDict.get('lijuanriqi')[:4]
                      )
                   )
    except sqlite3.Error,e:
        print 'insert value failed:',e.args[0]
        return
    conn.commit()
    #将存入的字符串拆分区第0个元素(例如:J_技术与管理档案,取J)
    return u'新增档案成功!'

#insert values into files table 
def insert_values_into_files(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("insert into %s values (%s,%s,'%s','%s','%s','%s','%s','%s','%s','%s','%s')"
                   % (paramDict.get('table_name'),
                      paramDict.get('fileID'),
                      paramDict.get('archiveID'),
                      paramDict.get('wenjiantimu'),
                      paramDict.get('wenjianbianhao'),
                      paramDict.get('fawendanwei'),
                      paramDict.get('xingchengriqi'),
                      paramDict.get('yeshu'),
                      paramDict.get('miji'),
                      paramDict.get('beizhu'),
                      paramDict.get('inputter'),
                      paramDict.get('filepath')
                      )
                   )
    except sqlite3.Error,e:
        print 'insert value into files failed:',e.args[0]
        return
    conn.commit()
    return u"新增文件成功!"

#insert user into users table 
def add_user(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("insert into users values ('%s','%s')"
                   % (paramDict.get('username'),
                      paramDict.get('password')
                      )
                   )
    except sqlite3.Error,e:
        print 'add user failed:',e.args[0]
        return
    conn.commit()


#insert values into destoryedArchives table 
def insert_values_into_destroyedArchives(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("insert into %s select *,'%s','%s' from %s where qixian not like '%s' and year like '%s' " %
                   (paramDict.get('target_table'),paramDict.get('people'),paramDict.get('date'),paramDict.get('from_table'),paramDict.get('qixian'),paramDict.get('year'))
                  )
       
    except sqlite3.Error,e:
        print 'insert value into destroyedArchives failed:',e.args[0]
        return
    conn.commit()
    
#insert values into destoryedArchives table 
def insert_values_into_destroyedFiles(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("insert into %s select *,'%s','%s' from %s where archiveID in (select archiveID from archives where qixian not like '%s' and year like '%s') " %
                   (paramDict.get('target_table'),paramDict.get('people'),paramDict.get('date'),paramDict.get('from_table'),paramDict.get('qixian'),paramDict.get('year'))
                  )
       
    except sqlite3.Error,e:
        print 'insert value into destroyedFiles failed:',e.args[0]
        return
    conn.commit()
    
#query archiveID from archives for destory
def query_archiveID_for_destory(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select archiveID from archives where qixian not like '%s' and year like '%s'  " % 
                   (paramDict.get("qixian"),paramDict.get("year")))

    except sqlite3.Error,e:
        print 'query archiveID data for destory faied:',e.args[0]
        return
    return cu.fetchall()

#query password from users
def query_password_for_modifyPassword(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select password from users where username like '%s' " % 
                   paramDict.get("username"))

    except sqlite3.Error,e:
        print 'query password  faied:',e.args[0]
        return
    return cu.fetchall()

#query user from users
def query_user_for_resetUser(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select username from users where username not like 'admin' ")
    except sqlite3.Error,e:
        print 'query username faied:',e.args[0]
        return
    return cu.fetchall()

#query fro guihao print
def query_for_guihao_print(conn,paramDict={}):
    cu = conn.cursor()
    try:
        temp_sql = "select \
	                  substr(archives.menlei,1,1),\
	                  substr(archives.guidang,1,2),\
	                  substr(archives.danwei,1,2),\
	                  archives.juanhao,\
	                  archives.quhao,\
	                  archives.guihao,\
	                  archives.hehao,\
	                  archives.juanhao,\
	                  archives.anjuantiming,\
	                  (select total(abs(files.yeshu)) from files where files.archiveID=archives.archiveID),\
	                  substr(archives.qixian,3,2),\
	                  substr(archives.menlei,3,length(menlei)),\
	                  archives.beizhu \
                    from archives where quhao like '%s' and guihao like '%s' " % \
        (paramDict.get('quhao'),paramDict.get('guihao'))
        #执行temp_sql
        cu.execute(temp_sql)
    except sqlite3.Error,e:
        print 'query guihao for print faied:',e.args[0]
        return
    return cu.fetchall()

#query fro guihao print
def query_for_hehao_print(conn,paramDict={}):
    cu = conn.cursor()
    try:
        temp_sql = "select \
	                  substr(archives.menlei,1,1),\
	                  substr(archives.guidang,1,2),\
	                  substr(archives.danwei,1,2),\
	                  archives.juanhao,\
	                  archives.quhao,\
	                  archives.guihao,\
	                  archives.hehao,\
	                  archives.juanhao,\
	                  archives.anjuantiming,\
	                  (select total(abs(files.yeshu)) from files where files.archiveID=archives.archiveID),\
	                  substr(archives.qixian,3,2),\
	                  substr(archives.menlei,3,length(menlei)),\
	                  archives.beizhu \
                    from archives where quhao like '%s' and guihao like '%s' and hehao like '%s' " % \
        (paramDict.get('quhao'),paramDict.get('guihao'),paramDict.get('hehao'))
        #执行temp_sql
        cu.execute(temp_sql)
    except sqlite3.Error,e:
        print 'query hehao for print faied:',e.args[0]
        return
    return cu.fetchall()


#query files for juanhao print
def query_files_for_juan_print(conn,paramDict={}):
    cu = conn.cursor()
    try:
        temp_sql = "select quhao,guihao,hehao,juanhao,\
                           wenjianbianhao,fawendanwei,wenjiantimu,xingchengriqi,yeshu,hujianhao,hujianhao,files.beizhu\
                    from archives,files where archives.archiveID=files.archiveID and \
                    quhao like '%s' and guihao like '%s' and hehao like '%s' and juanhao like '%s'" % \
        (paramDict.get('quhao'),paramDict.get('guihao'),paramDict.get('hehao'),paramDict.get('juanhao'))
        #执行temp_sql
        cu.execute(temp_sql)
    except sqlite3.Error,e:
        print 'query files for juan print faied:',e.args[0]
        return
    return cu.fetchall()

#query beikaobiao for print
def query_beikaobiao_for_print(conn,paramDict={}):
    cu = conn.cursor()
    try:
        temp_sql = "select quhao,guihao,hehao,juanhao,hujianhao,\
                           (select count(fileID) from files where files.archiveID = archives.archiveID),\
                           (select total(abs(yeshu)) from files where files.archiveID = archives.archiveID),\
                           inputter,lijuanriqi \
                    from archives where \
                    quhao like '%s' and guihao like '%s' and hehao like '%s' and juanhao like '%s'" % \
        (paramDict.get('quhao'),paramDict.get('guihao'),paramDict.get('hehao'),paramDict.get('juanhao'))
        #执行temp_sql
        cu.execute(temp_sql)
    except sqlite3.Error,e:
        print 'query beikaobiao for print faied:',e.args[0]
        return
    return cu.fetchall()


def query_anjuanbiaoti_print(conn,paramDict={}):
    cu = conn.cursor()    
    try:
        temp_sql = "select substr(kemu,1,2),substr(danwei,1,2),substr(weizhi,1,2),substr(menlei,1,1),substr(guidang,1,2),substr(danwei,1,2),\
                           juanhao,quhao,guihao,hehao,juanhao,substr(danwei,4,length(danwei)),\
                           (select min(substr(xingchengriqi,1,4)||substr(xingchengriqi,6,2)||substr(xingchengriqi,9,2)) \
                                   ||'-'|| max(substr(xingchengriqi,1,4)||substr(xingchengriqi,6,2)||substr(xingchengriqi,9,2)) \
                            from files where files.archiveID = archives.archiveID\
                            ),\
                           substr(qixian,3,2),substr(miji,4,2),anjuantiming\
                    from archives where quhao like '%s' and guihao like '%s' and hehao like '%s' " % \
                    (paramDict.get('quhao'),paramDict.get('guihao'),paramDict.get('hehao'))
        #执行temp_sql
        cu.execute(temp_sql)
    except sqlite3.Error,e:
        print 'query anjuanbiaoti print faied:',e.args[0]
        return
    return cu.fetchall()        

#query all values from table 
def query_all_values(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select  archiveID,    \
                            menlei,       \
                            guidang,      \
                            danwei,       \
                            anjuantiming, \
                            lijuanriqi,   \
                            weizhi,       \
                            miji,         \
                            zerenren,     \
                            quhao,        \
                            guihao,       \
                            hehao,        \
                            juanhao,      \
                            hujianhao,    \
                            kemu,         \
                            qixian,       \
                            beizhu,       \
                            inputter,     \
                            isGuidang     \
                    from %s" % paramDict.get('table_name'))
    except sqlite3.Error,e:
        print 'query data failed:',e.args[0]
        return
    return cu.fetchall()

#query some values from archives table 
def query_some_values_from_archivetable(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select  archiveID,    \
                                        menlei,       \
                                        guidang,      \
                                        danwei,      \
                                        anjuantiming, \
                                        lijuanriqi,   \
                                        weizhi,       \
                                        miji,         \
                                        zerenren,     \
                                        quhao,        \
                                        guihao,       \
                                        hehao,        \
                                        juanhao,      \
                                        hujianhao,    \
                                        kemu,         \
                                        qixian,       \
                                        beizhu,       \
                                        inputter,     \
                                        isGuidang     \
                    from %s where quhao like '%%%s%%' and guihao like '%%%s%%' and hehao like '%%%s%%' and juanhao like '%%%s%%' and anjuantiming like '%%%s%%' \
                    and menlei like '%%%s%%' and guidang like '%%%s%%' and qixian like '%%%s%%' and danwei like '%%%s%%' and lijuanriqi like '%%%s%%' and weizhi like '%%%s%%' \
                    and kemu like '%%%s%%' and miji like '%%%s%%' and inputter like '%%%s%%' and hujianhao like '%%%s%%' " % 
                    (paramDict.get('table_name'),paramDict.get('quhao'),paramDict.get('guihao'),paramDict.get('hehao'),paramDict.get('juanhao'),paramDict.get('anjuantiming'),
                     paramDict.get('menlei'),paramDict.get('guidang'),paramDict.get('qixian'),paramDict.get('danwei'),paramDict.get('lijuanriqi'),paramDict.get('weizhi'),
                     paramDict.get('kemu'),paramDict.get('miji'),paramDict.get('inputter'),paramDict.get('hujianhao')
                     )
                   )
    except sqlite3.Error,e:
        print 'query data failed:',e.args[0]
        return
    return cu.fetchall()

#query some values from files table 
def query_some_values_from_filestable(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select fileID,         \
                           archiveID,      \
                           wenjiantimu,    \
                           wenjianbianhao, \
                           fawendanwei,    \
                           xingchengriqi,  \
                           yeshu,          \
                           miji,           \
                           beizhu,         \
                           inputter        \
                    from %s where archiveID=%s and wenjianbianhao like '%%%s%%' and wenjiantimu like '%%%s%%' and fawendanwei like '%%%s%%' \
                    and xingchengriqi like '%%%s%%' and miji like '%%%s%%' and beizhu like '%%%s%%' and %s in (select archiveID from archives where weizhi like '%%%s%%') " % 
                    (paramDict.get('table_name'),
                     paramDict.get('archiveID'),
                     paramDict.get('wenjianbianhao'),
                     paramDict.get('wenjiantimu'),
                     paramDict.get('fawendanwei'),
                     paramDict.get('xingchengriqi'),
                     paramDict.get('miji'),
                     paramDict.get('beizhu'),  
                     paramDict.get('archiveID'),
                     paramDict.get('weizhi')
                     )
                   )
    except sqlite3.Error,e:
        print 'query some data from files failed:',e.args[0]
        return
    return cu.fetchall()

#query some values from files table no archivID
def query_some_values_from_filestable_noArchiveID(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select fileID,         \
                           archiveID,      \
                           wenjiantimu,    \
                           wenjianbianhao, \
                           fawendanwei,    \
                           xingchengriqi,  \
                           yeshu,          \
                           miji,           \
                           beizhu,         \
                           inputter        \
                    from %s where wenjianbianhao like '%%%s%%' and wenjiantimu like '%%%s%%' and fawendanwei like '%%%s%%' \
                    and xingchengriqi like '%%%s%%' and miji like '%%%s%%' and beizhu like '%%%s%%' and \
                    archiveID in (select distinct archiveID from archives where weizhi like '%%%s%%' ) " % 
                    (paramDict.get('table_name'),
                     paramDict.get('wenjianbianhao'),
                     paramDict.get('wenjiantimu'),
                     paramDict.get('fawendanwei'),
                     paramDict.get('xingchengriqi'),
                     paramDict.get('miji'),
                     paramDict.get('beizhu'),
                     paramDict.get('weizhi')
                     )
                   )
    except sqlite3.Error,e:
        print 'query some data from files no archiveID failed:',e.args[0]
        return
    return cu.fetchall()

#update some values from archives table 
def update_some_values_from_archivetable(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("update %s set menlei='%s',      \
                                  guidang='%s',     \
                                  qixian='%s',      \
                                  anjuantiming='%s',\
                                  danwei='%s',      \
                                  lijuanriqi='%s',  \
                                  weizhi='%s',      \
                                  miji='%s',        \
                                  zerenren='%s',    \
                                  quhao='%s',       \
                                  guihao='%s',      \
                                  hehao='%s',       \
                                  juanhao='%s',     \
                                  hujianhao='%s',   \
                                  kemu='%s',        \
                                  beizhu='%s',      \
                                  inputter='%s'     \
                                  where archiveID=%s " % 
                    (paramDict.get('table_name'),
                     paramDict.get('menlei'),
                     paramDict.get('guidang'),
                     paramDict.get('qixian'),
                     paramDict.get('anjuantiming'),
                     paramDict.get('danwei'),
                     paramDict.get('lijuanriqi'),
                     paramDict.get('weizhi'),
                     paramDict.get('miji'),
                     paramDict.get('zerenren'),
                     paramDict.get('quhao'),
                     paramDict.get('guihao'),
                     paramDict.get('hehao'),
                     paramDict.get('juanhao'),
                     paramDict.get('hujianhao'),
                     paramDict.get('kemu'),
                     paramDict.get('beizhu'),
                     paramDict.get('inputter'),
                     paramDict.get('archiveID')
                     )
                   )
    except sqlite3.Error,e:
        print 'update data failed:',e.args[0]
        return u'fail'
    conn.commit()
    return u'ok'

#update some values from files table
def update_some_values_from_filetable(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("update %s set archiveID=%s,         \
                                  wenjiantimu='%s',     \
                                  wenjianbianhao='%s',  \
                                  fawendanwei='%s',     \
                                  xingchengriqi='%s',   \
                                  yeshu='%s',           \
                                  miji='%s',            \
                                  beizhu='%s',          \
                                  inputter='%s',        \
                                  filepath='%s'         \
                                  where fileID=%s " % 
                        (
                          paramDict.get('table_name'),
                          paramDict.get('archiveID'),
                          paramDict.get('wenjiantimu'),
                          paramDict.get('wenjianbianhao'),
                          paramDict.get('fawendanwei'),
                          paramDict.get('xingchengriqi'),
                          paramDict.get('yeshu'),
                          paramDict.get('miji'),
                          paramDict.get('beizhu'),
                          paramDict.get('inputter'),
                          paramDict.get('filepath'),
                          paramDict.get('fileID')
                        )                        
                   )
    except sqlite3.Error,e:
        print 'update files data failed:',e.args[0]
        return u'fail'
    conn.commit()
    return u'ok'

#update password
def update_password(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("update users set password='%s' where username like '%s' " % 
                    (paramDict.get('password'),paramDict.get('username'))                        
                  )
    except sqlite3.Error,e:
        print 'update password data failed:',e.args[0]
        return u'fail'
    conn.commit()
    return u'ok'


#query juanhao from table
def query_juanhao_values(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select juanhao from %s where quhao like '%s' and guihao like '%s' and hehao like '%s' " % 
                   (paramDict.get("table_name"),paramDict.get("quhao"),paramDict.get("guihao"),paramDict.get("hehao")))

    except sqlite3.Error,e:
        print 'query data faied:',e.args[0]
        return
    return cu.fetchall()

def query_isGuidang_values_for_addfile(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select isGuidang from %s where quhao like '%s' and guihao like '%s' and hehao like '%s' and juanhao like '%s' " % 
                   (paramDict.get("table_name"),paramDict.get("quhao"),paramDict.get("guihao"),paramDict.get("hehao"),paramDict.get("juanhao"))
                  )
    except sqlite3.Error,e:
        print 'query isGuidang data faied:',e.args[0]
        return
    return cu.fetchall()


#query archiveID from  files table
def query_archiveID_values(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select archiveID from %s where quhao like '%s' and guihao like '%s' and hehao like '%s' and juanhao like '%s' " % 
                   (paramDict.get("table_name"),paramDict.get("quhao"),paramDict.get("guihao"),paramDict.get("hehao"),paramDict.get("juanhao")))
    except sqlite3.Error,e:
        print 'query archiveID for wenjianbianhao data faied:',e.args[0]
        return
    return cu.fetchall()

#query wenjianbianhao from  files table
def query_wenjianbianhao_values(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select wenjianbianhao from %s where archiveID like %s " % 
                   (paramDict.get("table_name"),paramDict.get("archiveID")))
    except sqlite3.Error,e:
        print 'query  wenjianbianhao data faied:',e.args[0]
        return
    return cu.fetchall()

#query isGuidang from  archives table
def query_isGuidang_values(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select isGuidang from %s where archiveID like %s " % 
                   (paramDict.get("table_name"),paramDict.get("archiveID")))
    except sqlite3.Error,e:
        print 'query  isGuidang data faied:',e.args[0]
        return
    return cu.fetchall()


#query juanhao from table
def query_juanhaoArchiveID_values_for_addfile(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select juanhao,archiveID from %s where quhao like '%s' and guihao like '%s' and hehao like '%s' " % 
                   (paramDict.get("table_name"),paramDict.get("quhao"),paramDict.get("guihao"),paramDict.get("hehao")))
    except sqlite3.Error,e:
        print 'query data faied:',e.args[0]
        return
    return cu.fetchall()

# query archiveID from archive table
def query_archiveID_for_selectfile(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select archiveID from %s" % paramDict.get("table_name"))
    except sqlite3.Error,e:
        print 'query archiveID data for selectfile faied:',e.args[0]
        return
    return cu.fetchall()

# query menlei guidang danwei juanhao from archives table
def query_mgdj_from_archives(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select quhao,guihao,hehao,juanhao from %s where archiveID=%s " % (paramDict.get("table_name"),paramDict.get("archiveID")))
    except sqlite3.Error,e:
        print 'query quhao guihao hehao juanhao data  faied:',e.args[0]
        return
    return cu.fetchall()

# query neirong from files table
def query_neirong_from_files(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select filepath from %s where fileID=%s "  %  (paramDict.get("table_name"),paramDict.get("fileID")))
    except sqlite3.Error,e:
        print 'query neirong(filepath) data  faied:',e.args[0]
        return
    return cu.fetchall()

# query neirong from files table
def query_fileID_filepath_from_files(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("select fileID,filepath from %s where archiveID=%s "  %  (paramDict.get("table_name"),paramDict.get("archiveID")))
    except sqlite3.Error,e:
        print 'query fileID,filepath data  faied:',e.args[0]
        return
    return cu.fetchall()

def query_filepath_from_files(conn,paramDict={}):
    cu = conn.cursor()
    try:   
        cu.execute("select fileID from %s where filepath='%s' "  %  (paramDict.get("table_name"),paramDict.get("filepath")))
        
    except sqlite3.Error,e:
        print 'query filepath data  faied:',e.args[0]
        return
    return cu.fetchall()

def query_user_count(conn,paramDict={}):
    cu = conn.cursor()
    try:   
        cu.execute("select username from users")
    except sqlite3.Error,e:
        print 'query user count data  faied:',e.args[0]
        return
    return cu.fetchall()

#create archives table 
def create_archives_table(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute('create table if not exists %s (archiveID integer primary key autoincrement,\
                                                    menlei text not null,\
                                                    guidang text not null,\
                                                    qixian text not null,\
                                                    anjuantiming text not null,\
                                                    danwei text not null,\
                                                    lijuanriqi text not null,\
                                                    weizhi text not null,\
                                                    miji text not null,\
                                                    zerenren text not null,\
                                                    quhao text not null,\
                                                    guihao text not null,\
                                                    hehao text not null,\
                                                    juanhao text not null,\
                                                    hujianhao text not null,\
                                                    kemu text not null,\
                                                    beizhu text not null,\
                                                    inputter text not null,\
                                                    isGuidang text not null,\
                                                    year text not null\
                                                   )' % paramDict.get('table_name'))
    except sqlite3.Error,e:
        print 'create archives or destroyArchives table failed:',e.args[0]
        return
    conn.commit()

#create archives table (people,date---销毁人、销毁日期)
def create_destroy_archives_table(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute('create table if not exists %s (archiveID integer not null,\
                                                    menlei text not null,\
                                                    guidang text not null,\
                                                    qixian text not null,\
                                                    anjuantiming text not null,\
                                                    danwei text not null,\
                                                    lijuanriqi text not null,\
                                                    weizhi text not null,\
                                                    miji text not null,\
                                                    zerenren text not null,\
                                                    quhao text not null,\
                                                    guihao text not null,\
                                                    hehao text not null,\
                                                    juanhao text not null,\
                                                    hujianhao text not null,\
                                                    kemu text not null,\
                                                    beizhu text not null,\
                                                    inputter text not null,\
                                                    isGuidang text not null,\
                                                    year text not null,\
                                                    people text not null,\
                                                    date text not null\
                                                   )' % paramDict.get('table_name'))
    except sqlite3.Error,e:
        print 'create destory archives  table failed:',e.args[0]
        return
    conn.commit()


#create files table
def create_files_table(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute('create table if not exists %s (fileID integer primary key autoincrement,\
                                                    archiveID integer not null,\
                                                    wenjiantimu text not null,\
                                                    wenjianbianhao text not null,\
                                                    fawendanwei text not null,\
                                                    xingchengriqi text not null,\
                                                    yeshu text not null,\
                                                    miji text not null,\
                                                    beizhu text not null,\
                                                    inputter text not null,\
                                                    filepath text not null\
                                                   )' % paramDict.get('table_name')
                                                      )
    except sqlite3.Error,e:
        print 'create files table failed:',e.args[0]
        return
    conn.commit()

#create files table
def create_destory_files_table(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute('create table if not exists %s (fileID integer not null,\
                                                    archiveID integer not null,\
                                                    wenjiantimu text not null,\
                                                    wenjianbianhao text not null,\
                                                    fawendanwei text not null,\
                                                    xingchengriqi text not null,\
                                                    yeshu text not null,\
                                                    miji text not null,\
                                                    beizhu text not null,\
                                                    inputter text not null,\
                                                    filepath text not null,\
                                                    people text not null ,\
                                                    date text not null \
                                                   )' % paramDict.get('table_name')
                                                      )
    except sqlite3.Error,e:
        print 'create files table failed:',e.args[0]
        return
    conn.commit()

#create users table
def create_users_table(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute('create table if not exists users (username text primary key,password text not null)')
    except sqlite3.Error,e:
        print 'create users table failed:',e.args[0]
        return
    conn.commit()
    
#drop table if exist 
def drop_table(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute('drop table if exists %s' % paramDict.get('table_name'))
    except suqlit3.Error,e:
        print 'drop table failed:',e.args[0]
        return
    conn.commit()

#delete archive values from table
def delete_row_from_archivetable(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("delete from %s where archiveID=%s" % (paramDict.get('table_name'),
                                                           paramDict.get('archiveID')
                                                           )
                   )
    except suqlit3.Error,e:
        print 'delete archive table failed:',e.args[0]
        return
    conn.commit()
    
#delete file values from table
def delete_row_from_filetable(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("delete from %s where fileID=%s" % (paramDict.get('table_name'),
                                                           paramDict.get('fileID')
                                                           )
                   )
    except suqlit3.Error,e:
        print 'delete file table failed:',e.args[0]
        return
    conn.commit()

#delete file values from table by archiveID
def delete_row_from_filetable_by_ArchiveID(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("delete from %s where archiveID=%s" % (paramDict.get('table_name'),
                                                           paramDict.get('archiveID')
                                                           )
                   )
    except suqlit3.Error,e:
        print 'delete file table by archiveID failed:',e.args[0]
        return
    conn.commit() 

#update guidang from archvies table
def update_guidang(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("update 'archives' set isGuidang='%s' where year like '%s' " % 
                        (
                          u'是',
                          paramDict.get('year')
                        )                        
                   )
    except sqlite3.Error,e:
        print 'update guidang data failed:',e.args[0]
        return u'fail'
    conn.commit()
    return u'ok'    

def update_quxiao_guidang(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("update 'archives' set isGuidang='%s' where year like '%s' " % 
                        (
                          u'否',
                          paramDict.get('year')
                        )                        
                   )
    except sqlite3.Error,e:
        print 'update quxiao guidang data failed:',e.args[0]
        return u'fail'
    conn.commit()
    return u'ok' 

def query_already_year(conn,paramDict={}):
    cu = conn.cursor()
    try:   
        cu.execute("select distinct year from 'archives' where isGuidang like '%s' "  %  u'是')
        
    except sqlite3.Error,e:
        print 'query already year data  faied:',e.args[0]
        return
    return cu.fetchall()

def query_noGuidang_year(conn,paramDict={}):
    cu = conn.cursor()
    try:   
        cu.execute("select distinct year from 'archives' where isGuidang like '%s' "  %  u'否')
        
    except sqlite3.Error,e:
        print 'query no guidang year data  faied:',e.args[0]
        return
    return cu.fetchall()

#destroy guidang from archives
def destroy_guidang_by_year(conn,paramDict={}):
    cu = conn.cursor()
    try:
        cu.execute("insert into 'destroyArchives' values (select * from 'archives' where isGuidang like u'是' and year like '%s')"
                   %(paramDict.get('year')))
        #cu.execute("delete from 'archives' where isGuidang like u'是' and year like '%s'" % (paramDict.get('year')))        
    except sqlite3.Error,e:
        print 'insert value into destroy failed:',e.args[0]
        return
    conn.commit()

#main function 
if __name__ == '__main__':
    db_name = './Per.db'
    conn = connect_db(db_name)
    table_name = 'archives'
    close_db(conn)
 