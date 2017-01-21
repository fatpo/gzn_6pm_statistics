# #coding:utf-8  
from PyQt4 import QtGui
from PyQt4 import QtCore
from PyQt4.QtWebKit import*
from PyQt4 import QtCore, QtGui
from PyQt4 import uic
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4.QtWebKit import *
from PyQt4.QtNetwork import *

import os,sqlite3,re
import sys
import csv
import xlrd
import os
import datetime

reload(sys)
sys.setdefaultencoding('utf-8')
 
class MyConsole(QWidget):
    def __init__(self,parent):
        QWidget.__init__(self)
        self.parent = parent
         
        self.initUI()
        self.initConfig()
        
    #初始化UI    
    def initUI(self):
        self.gridlayout = QtGui.QGridLayout()
        for i in range(10):
            self.gridlayout.setColumnStretch(i,1)
        for i in range(7):
            self.gridlayout.setRowStretch(i,1)
        
        self.loadFileBtn = QPushButton(u"选择文件")
        self.connect(self.loadFileBtn, QtCore.SIGNAL('clicked()'), self.onLoadFileBtn)
        self.fileEntry = QLineEdit()
        
        self.gridlayout.addWidget(self.loadFileBtn, 0, 0)
        self.gridlayout.addWidget(self.fileEntry, 0, 1,1,4)
        
        lb1 = QLabel(u'车站集合:')
        lb2 = QLabel(u'一次筛选担当局:')
        lb3 = QLabel(u'多次筛选担当局:')
        lb4 = QLabel(u'贵广南广车站:')
        lb5 = QLabel(u'高铁小计:')
        lb6 = QLabel(u'合股小计:')
        
        self.stationsEntry = QTextEdit()
        self.oneRespEntry = QTextEdit()
        self.multRespEntry = QTextEdit()
        self.ggngEntry = QTextEdit()
        self.highwayEntry = QTextEdit()
        self.heguEntry = QTextEdit()
        
        self.gridlayout.addWidget(lb1, 1, 0)
        self.gridlayout.addWidget(lb2, 2, 0)
        self.gridlayout.addWidget(lb3, 3, 0)
        self.gridlayout.addWidget(lb4, 4, 0)
        self.gridlayout.addWidget(lb5, 5, 0)
        self.gridlayout.addWidget(lb6, 6, 0)
        
        self.gridlayout.addWidget(self.stationsEntry, 1, 1,1,8)
        self.gridlayout.addWidget(self.oneRespEntry,2, 1,1,8)
        self.gridlayout.addWidget(self.multRespEntry,3, 1,1,8)
        self.gridlayout.addWidget(self.ggngEntry,4, 1,1,8)
        self.gridlayout.addWidget(self.highwayEntry,5, 1,1,8)
        self.gridlayout.addWidget(self.heguEntry,6, 1,1,8)
        
      
        
        self.goBtn = QtGui.QPushButton(u"统计")
        self.connect(self.goBtn, QtCore.SIGNAL('clicked()'), self.onGoButton)
        self.gridlayout.addWidget(self.goBtn, 0, 5,1,1)
     
        self.saveConfigBtn = QtGui.QPushButton(u"保存")
        self.connect(self.saveConfigBtn, QtCore.SIGNAL('clicked()'), self.onSaveConfigButton)
        self.gridlayout.addWidget(self.saveConfigBtn, 0, 8,1,1)
        
        self.setLayout(self.gridlayout)        
    
    #处理数据
    def onGoButton(self):
        filename = self.fileEntry.text()
        if len(filename) == 0:
            QtGui.QMessageBox.about(self, u'提醒', u"未选中数据文件，格式为：xx.xls")
            print u"未选中数据文件，格式为：xx.xls"
        else:
            if 0 == self.xls2csv(filename):
                stations = unicode(self.stationsEntry.toPlainText(),"utf-8").encode('gbk').split(',')
                oneResp = unicode(self.oneRespEntry.toPlainText(),"utf-8").encode('gbk').split(',')
                multResp = unicode(self.multRespEntry.toPlainText(),"utf-8").encode('gbk').split(',')
                ggng = unicode(self.ggngEntry.toPlainText(),"utf-8").encode('gbk').split(',')
                highway = unicode(self.highwayEntry.toPlainText(),"utf-8").encode('gbk').split(',')
                hegu = unicode(self.heguEntry.toPlainText(),"utf-8").encode('gbk').split(',')
                
                resps2_dic = {u'广铁集团'.encode('gbk'):(0,0),u'国铁妹纸'.encode('gbk'):(0,0)}
                
                with open('out.txt','wb') as out: 
                    for (d,x) in self.station_dic.items():
                        with open(x, 'rb') as f:
                            #初始化
                            resp_dic = {}
                            for i in oneResp:
                                resp_dic[i] = (0,0)
                            for i in resps2_dic:
                                resps2_dic[i] = (0,0)
                                
                            reader = csv.reader(f)
                            for row in reader:
                                trainNO = row[3]
                                start = row[4]
                                arrive = row[5]
                                resp = row[6]
                                price = row[7]
                                num = row[8]
                                
                                #一次筛选的
                                if resp in resp_dic:
                                    v1 = resp_dic[resp][0]
                                    v2 = resp_dic[resp][1]
                                    v1 += float(num)
                                    v2 += float(price)
                                    resp_dic[resp] = (v1,v2)
                                # D开头且到发有深圳北都是厦深广东
                                elif trainNO[0]=='D' and u'深圳北'.encode('gbk') in [start,arrive] and (start not in ggng) and (arrive not in ggng):
                                    k = u'厦深广东'.encode('gbk')
                                    v1 = resp_dic[k][0]
                                    v2 = resp_dic[k][1]
                                    v1 += float(num)
                                    v2 += float(price)
                                    resp_dic[k] = (v1,v2)
                                # G开头且到发有深圳北都是广深港
                                elif trainNO[0]=='G' and u'深圳北'.encode('gbk') in [start,arrive]:
                                    k = u'广深港公司'.encode('gbk')
                                    v1 = resp_dic[k][0]
                                    v2 = resp_dic[k][1]
                                    v1 += float(num)
                                    v2 += float(price)
                                    resp_dic[k] = (v1,v2)
                                # 福田出发或者达到也都是广深港
                                elif u'福田'.encode('gbk') in [start,arrive]:
                                    k = u'广深港公司'.encode('gbk')
                                    v1 = resp_dic[k][0]
                                    v2 = resp_dic[k][1]
                                    v1 += float(num)
                                    v2 += float(price)
                                    resp_dic[k] = (v1,v2)
                                # 含有广铁集团的
                                elif resp  == u'广铁集团'.encode('gbk'):
                                    k = u'广铁集团'.encode('gbk')
                                    v1 = resps2_dic[k][0]
                                    v2 = resps2_dic[k][1]
                                    v1 += float(num)
                                    v2 += float(price)
                                    resps2_dic[k] = (v1,v2)
                                # 其它都是国铁
                                else:
                                    k = u'国铁妹纸'.encode('gbk')
                                    v1 = resps2_dic[k][0]
                                    v2 = resps2_dic[k][1]
                                    v1 += float(num)
                                    v2 += float(price)
                                    resps2_dic[k] = (v1,v2)
                            
                            
                            out.write('##############'+d+'##############\r\n')
                            ans_v1 = 0
                            ans_v2 = 0
                            #高铁小计
                            ans_highway_v1 = 0
                            ans_highway_v2 = 0
                            #合股小计
                            ans_he_gu_v1 = 0
                            ans_he_gu_v2 = 0
                            
                            
                            for i in resps2_dic:
                                ans_v1 += resps2_dic[i][0]
                                ans_v2 += resps2_dic[i][1]
                                #四舍五入
                                v1 = int((resps2_dic[i])[0])
                                v2 = int((resps2_dic[i])[1]) 
                                resps2_dic[i] = (v1,v2)
                                out.write(i+'\t'+str(resps2_dic[i][0])+'\t'+str(resps2_dic[i][1])+'\r\n')
                            
                            #排个序
                            tmplist = sorted(resp_dic.iteritems(), key=lambda d:d[1][1], reverse = True)
                            for i in tmplist:
                                #统计
                                if i[0] in highway:
                                    ans_highway_v1 += i[1][0]
                                    ans_highway_v2 += i[1][1]
                                if i[0] in hegu:
                                    ans_he_gu_v1 += i[1][0]
                                    ans_he_gu_v2 += i[1][1]
                                    
                                ans_v1 += i[1][0]
                                ans_v2 += i[1][1]
                                out.write(i[0]+'\t'+str(int(i[1][0]))+'\t'+str(int(i[1][1]))+'\r\n')
                                
                            out.write(u'高铁小计:'.encode('gbk')+str(int(ans_highway_v1))+'\t'+str(int(ans_highway_v2))+' \r\n')
                            out.write(u'合股小计:'.encode('gbk')+str(int(ans_he_gu_v1))+'\t'+str(int(ans_he_gu_v2))+' \r\n')
                            out.write(u'总张数:'.encode('gbk')+str(int(ans_v1))+u'\t总票额:'.encode('gbk')+str(int(ans_v2))+' \r\n')
                            out.write('################################\r\n\r\n')

                QtGui.QMessageBox.about(self, u'成功', u"请查看:out.txt")
                print u"成功：请查看:out.txt"
                os.system('notepad out.txt')
             
    #保存
    def onSaveConfigButton(self):
        stations = self.stationsEntry.toPlainText()
        oneResp = self.oneRespEntry.toPlainText()
        multResp = self.multRespEntry.toPlainText()
        ggng = self.ggngEntry.toPlainText()
        highway = self.highwayEntry.toPlainText()
        hegu = self.heguEntry.toPlainText()
        
        if  os.path.exists('./config/my.db'):
            #检查输入是否正确
            inputs = [stations,oneResp,multResp,ggng,highway,hegu]
            ugly = [u'，','.',':',u'。']
            a = True
            for i in range(len(inputs)):
                for j in inputs[i]:
                    if j in ugly:
                        a = False
                        QtGui.QMessageBox.about(self, u'出了点问题', u"检查第"+str(i+1)+"个框框，不能有以下符号："+j)
                        break
            if a:
                cqssc_db = sqlite3.connect("./config/my.db")
                sql = "update config set stations = '"+stations+ "',oneResp = '"+oneResp + "',multResp = '" + multResp + "', ggng = '"+ggng+"',highway='"+highway+"',hegu = '"+hegu+"';"
               
                cqssc_db.execute(str(sql))
                cqssc_db.commit()
                cqssc_db.close()
                QtGui.QMessageBox.about(self, u'成功', u"数据库保存成功")
        else:
            print './config/my.db No Exist!'
            QtGui.QMessageBox.about(self, u'出了点问题', u"数据库文件怎么不见了..重启下程序吧")
    
    #选择文件
    def onLoadFileBtn(self):
        fname = QFileDialog.getOpenFileName(self, u'打开', 
         './',u"符合格式的文件 (*.xls)")
        self.fileEntry.setText(fname)
        
    #初始化配置
    def initConfig(self):
        if os.path.exists('./config/my.db'):
            cqssc_db = sqlite3.connect("./config/my.db")
            cursor = cqssc_db.execute("select * from config;")
            
            for row in cursor:
                self.stationsEntry.setText(row[0])
                self.oneRespEntry.setText(row[1])
                self.multRespEntry.setText(row[2])
                self.ggngEntry.setText(row[3])
                self.highwayEntry.setText(row[4])
                self.heguEntry.setText(row[5])
                
            cursor.close()            
            cqssc_db.commit()
            cqssc_db.close()
        else:
            cqssc_db = sqlite3.connect("./config/my.db")
            sql = "create table config(stations text NULL,oneResp text NULL,multResp text NULL,ggng text NULL,highway text NULL,hegu text NULL);"
            cqssc_db.execute(sql)
            cqssc_db.commit()
            
            sql = "insert into config(stations,oneResp,multResp,ggng,highway,hegu)"
            sql+=" values('','','','','','');"
            cqssc_db.execute(sql)
            cqssc_db.commit()
            cqssc_db.close()
    
    #先把xls转成csv
    def xls2csv(self,filename):
        #先把xls转成csv
        a = xlrd.open_workbook(filename,formatting_info=True)
        i = a.sheets()[3] #车次表

        #按照第三列的“售票处” 找到 合并列 
        b = i.merged_cells
        c = [k for k in b if k[2]==2 and k[3] == 3]


        stations = unicode(self.stationsEntry.toPlainText(),"utf-8").encode('gbk').split(',')
        d = [k for k in c if (i.cell_value(k[0],k[2])).encode('gbk') in stations]
        
        self.station_dic = {}
        for k in stations:
             self.station_dic[k] = k+'.csv'
             
        for k in d:
            output = open(self.station_dic[(i.cell_value(k[0],k[2])).encode('gbk')],'w')
            for r in range(k[0],k[1]):
                linevalue = []
                for l in range(i.ncols):
                    if i.cell(r,l).ctype == 0:
                        linevalue.append(('%*s'%(4,' ')))
                    else:
                        b = i.cell(r,l).value
                        if type(b) == type(4.0):
                            b = str(b)
                        linevalue.append(b)
                line = ','.join(linevalue).encode('gbk')
                output.write(line+"\n")
            output.close()

        return 0
            

class MainWindow(QtGui.QMainWindow):
    def __init__(self):
        QtGui.QMainWindow.__init__(self)

        tabs = QtGui.QTabWidget(self)
     
        tab1 = QtGui.QWidget()   
        self.console = MyConsole(self)
        
        vBoxlayout = QtGui.QVBoxLayout()
        vBoxlayout.addWidget(self.console)
        tab1.setLayout(vBoxlayout)
        tabs.addTab(tab1,u"控制台")
        
        tabs.resize(700, 500)
        self.resize(700, 500)
        
        self.setWindowTitle(u'后台18点数据统计助手')
        
        #禁止最大化
        self.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)  
        
        self.show()
       
    
    @pyqtSlot(str) 
    def mySetWindowTitle(self,title):
        self.setWindowTitle(title)
        
        
    def closeEvent(self, event):
        reply = QtGui.QMessageBox.question(self, u'退出',u"您确定离开吗？", QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
            
def main():
    app = QtGui.QApplication(sys.argv)
    win = MainWindow()
    sys.exit(app.exec_())
 
if __name__ == '__main__':
    main()
