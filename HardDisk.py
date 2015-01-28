# -*- coding: UTF-8 -*- 
import os, sys
import time
import wmi,zlib
import hashlib
import base64

cpu_str = ''
hd_str = ''
def get_cpu_info() :
    tmpdict = {}
    tmpdict["CpuCores"] = 0
    c = wmi.WMI()
#print c.Win32_Processor().['ProcessorId']
#print c.Win32_DiskDrive()
    for cpu in c.Win32_Processor():    
    #print cpu
        print "cpu id:", cpu.ProcessorId.strip()
        tmpdict["CpuType"] = cpu.Name
        try:
            tmpdict["CpuCores"] = cpu.NumberOfCores
        except:
            tmpdict["CpuCores"] += 1
            tmpdict["CpuClock"] = cpu.MaxClockSpeed
            return tmpdict
 
def _read_cpu_usage():
    c = wmi.WMI ()
    for cpu in c.Win32_Processor():
        return cpu.LoadPercentage
 
def get_cpu_usage():
    cpustr1 =_read_cpu_usage()
    if not cpustr1:
        return 0
    time.sleep(2)
    cpustr2 = _read_cpu_usage()
    if not cpustr2:
        return 0
    cpuper = int(cpustr1)+int(cpustr2)/2
    return cpuper

def get_disk_info():
    tmplist = []
    encrypt_str = ""    
    c = wmi.WMI ()
    for cpu in c.Win32_Processor():
        #cpu 序列号
        encrypt_str = encrypt_str+cpu.ProcessorId.strip()
        #print "cpu id:", cpu.ProcessorId.strip()    
    for physical_disk in c.Win32_DiskDrive():
        encrypt_str = encrypt_str+physical_disk.SerialNumber.strip()
        #硬盘序列号
        #print 'disk id:', physical_disk.SerialNumber.strip()
    return encrypt_str

#获取机器码
def getMachineCode():
    return get_disk_info()

#获取加密码
def getSecretCode():
    mc = getMachineCode()
    sc1 = hashlib.md5(mc.encode('utf-8')).hexdigest()
    sc2 = hashlib.md5(sc1.encode('utf-8')).hexdigest()
    return sc2
def encodeBase64(s=''):
    return base64.b64encode(s)
def decodeBase64(s=''):
    return base64.b64decode(s).decode('utf-8')

def createSN(mcdw=''):
    mc_dw = mcdw.split('@')
    temp = ''
    temp = hashlib.md5(unicode(mc_dw[0])).hexdigest() + u'@' + hashlib.md5(unicode(encodeBase64(mc_dw[1]))).hexdigest()
    return temp

#if __name__ == "__main__":
    #print decodeBase64(u'5aSp5L+d5Yqe5YWs5a6k')
    #print createSN('BFEBFBFF000306C3S4Y0N22W@5aSp5L+d5Yqe5YWs5a6k')
    ##a = get_cpu_info()
    #encrypt_str = get_disk_info()
    ##加密
    #print "org :",encrypt_str
    #md5a = hashlib.md5(encrypt_str.encode('utf-8')).hexdigest()
    #print "md5a:",md5a
    #md5b = hashlib.md5(md5a.encode('utf-8')).hexdigest()
    #print "md5b:",md5b    
    #print "machine code : " , getMachineCode()
    #print "secret code : " , getSecretCode()
    
    #machine code :  BFEBFBFF000306C3S4Y0N22W
    #secret code :  618bac5dcbf70c214438ec0994ceac4c
