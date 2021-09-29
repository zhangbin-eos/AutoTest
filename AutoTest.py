from openpyxl import load_workbook as LoadExcel
import urllib.parse as urlparse

import logging
import os
import argparse
import sys

class CmdObj:
    def __init__(self,config):
        return ;



class ExcelToDict:
    def __init__(self,file,title_row=0):
        self.title_row = int(title_row)
        self.ExcelBook = LoadExcel(filename=file);
        self.ExcelDate = {}

    def SheetToList(self,SheetName):
        DataObjList=[]
        titlelist=[]
        if not SheetName in self.ExcelBook.sheetnames:
            raise Exception('SheetName {0} not Existence'.format(SheetName));
        SheetObj = self.ExcelBook[SheetName];
        self.ExcelDate[SheetName]={}

        rows = tuple(SheetObj.rows)
        for row in rows:
            rowsDataObj={}
            if rows.index(row)==self.title_row:
                #这行是title
                for i in row:
                    titlelist.append(i.value)
                continue;

            # 读取行里的所有元素
            for i in row:
                rowsDataObj[titlelist[row.index(i)]]=i.value;
            DataObjList.append(rowsDataObj);
            self.ExcelDate[SheetName]=DataObjList;
        return self.ExcelDate[SheetName];


def SheetDatatoPortlist(PortSheetData):
    Portlist={}
    if type(PortSheetData)!=list:
        raise Exception('config type is {0} , Exception list'.format(type(PortSheetData)));
    for item in PortSheetData:
        Portlist[item['vport']]={}
        parsed = urlparse.urlparse(item['config']);
        Portlist[item['vport']]['type']=parsed.scheme;
        Portlist[item['vport']]['config']={};
        if parsed.scheme=='uart':
            print(urlparse.parse_qs(parsed.query))
            Portlist[item['vport']]['config'].update(urlparse.parse_qs(parsed.query))
        elif parsed.scheme=='tcp':
            Portlist[item['vport']]['config']['ip']=parsed.netloc
        elif parsed.scheme=='udp':
            Portlist[item['vport']]['config']['ip']=parsed.netloc
        elif parsed.scheme=='http':
            Portlist[item['vport']]['config']['url']=item['config'];
    return Portlist;


import  serial
import  time

def TestOneCmd(cmdobj,rcvdfunc,sendfunc):
    sendfunc(cmdobj["send"])
    #logger.debug('send: "%s" OK', cmdobj["send"].replace('\r','<CR>').replace('\n','<LF>'))

    relrcvd = rcvdfunc(cmdobj["rcvd"],cmdobj["timeout"])
    #logger.debug('rcvd "%s" ',relrcvd)
    cmpRes=True;
    CmdTimeoutError=False;

    if len(relrcvd)==0:
        #接收超时了
        cmpRes=False;
        CmdTimeoutError=True;
    else:
        relrcvd = relrcvd.replace('\n','')
        for i in range(0,len(cmdobj["rcvd"])):
            if relrcvd[i] != cmdobj["rcvd"][i]:
                cmpRes=False;
                break;
    
    relrcvd = relrcvd.replace('\r','<CR>').replace('\n','<LF>');

    if cmpRes == True:
        #logger.debug('rcvd "%s" OK',relrcvd)
        time.sleep(cmdobj["wait"])
    else:
        if( cmdobj.has_key("ignor_error") == True) and (cmdobj["ignor_error"] == True ):
            #logger.warning('rcvd "%s",but except "%s",ignor error',relrcvd,cmdobj["rcvd"].replace('\r','<CR>').replace('\n','<LF>'))
            return True;
        if 	CmdTimeoutError==False:
            #logger.error('rcvd: "%s",but except "%s"',relrcvd,cmdobj["rcvd"].replace('\r','<CR>').replace('\n','<LF>'))
            CmdTimeoutError=False
        else:
            CmdTimeoutError=False
            #logger.error('rcvd: "%s",timeout %d',relrcvd,cmdobj["timeout"])
        return False;
    return True;


def openlogfile(logfilepath):
    global LogF
    global logger
    logger=logging.getLogger();
    logger.setLevel(logging.DEBUG)
    LogF = logging.FileHandler(logfilepath, mode='w+')
    LogF.setLevel(logging.DEBUG)
    LogF.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
    logger.addHandler(LogF)




if __name__ == '__main__':

    parser = argparse.ArgumentParser()
    parser.add_argument("-f","--file", type=str,help="input file(.xlsx)")
    parser.add_argument(
        "-o",
        "--log",
        nargs='?',
        default="log_"+time.strftime("%Y%m%d-%H%M%S",time.localtime(time.time()))+".log",
        type=str,
        help="output logfile(.txt)")
    #parser.add_argument("-v", "--verbosity", help="increase output verbosity")

    args = parser.parse_args()
    openlogfile(args.log)

    if args.file==None:
        logger.info("input file is None");
        logger.info("{0}".format(parser.print_help()));

    Excelfile=args.file;
    logger.info('open '+Excelfile);
    ExcelObj=ExcelToDict(Excelfile);

    PortList=SheetDatatoPortlist(ExcelObj.SheetToList("Port"));

    Cmdlist=ExcelObj.SheetToList("CMD")

    while True:

        for iter in Cmdlist:
            if not iter["port"] in PortList:
                raise Exception('maybe excel port info err? no found cmd sheet {0} in port sheet'.format(iter["port"]));
            config=PortList[iter["port"]]["config"];
            
            if PortList[iter["port"]]["type"]=='uart':
                if not "PortHandle" in PortList[iter["port"]]:
                    PortHandle=serial.Serial(str(config['port'][0]));
                    PortHandle.baudrate=115200;
                    PortHandle.bytesize=int(config['databits'][0])
                    PortHandle.stopbits=int(config['stopbits'][0]);
                    PortHandle.parity=serial.PARITY_NONE;
                    PortHandle.xonxoff=False;
                    PortHandle.rtscts=False;
                    PortHandle.dsrdtr=False;

                    PortList[iter["port"]]["PortHandle"]=PortHandle;
                    print("open {0} baud {1}".format(str(config['port'][0]),int(config['baud'][0])));
                    del PortHandle

                PortHandle=PortList[iter["port"]]["PortHandle"];
                PortHandle.timeout=float(iter['timeout']); 
                PortHandle.inter_byte_timeout=0.5;

                # 如果字符串中有<HEX>标记,表示后面的数据发送使用HEX
                SendByteData=iter['send'].replace('<CR>','\r').replace('<LF>','\n').encode();
                ExceptRcvdByteData=iter['rcvd'].encode()
                PortHandle.flush()
                PortHandle.reset_input_buffer()
                PortHandle.reset_output_buffer()
                PortHandle.write(SendByteData);
                logmsg="send to {0}:{1}".format(iter["port"],SendByteData);
                logger.info(logmsg);
                print(logmsg);


                RcvdByteData=PortHandle.read_until(b'\n');
                if(len(RcvdByteData)==0):
                    logmsg="rcvd from {0}:timeout({1})s".format(iter["port"],float(iter['timeout']));
                    logger.error(logmsg);
                else:
                    if RcvdByteData.find(ExceptRcvdByteData)==-1:
                        logmsg='rcvd from {0}: "{1}",but except "{2}"'.format(iter["port"],RcvdByteData,ExceptRcvdByteData);
                        logger.error(logmsg);
                    else:
                        logmsg="rcvd from {0}:{1}".format(iter["port"],RcvdByteData);
                        logger.info(logmsg);
                print(logmsg);
                del PortHandle
            elif PortList[iter["port"]]["type"]=='tcp':
                print("send to",PortList[iter["port"]]["type"],"msg:",iter["send"]);
            elif PortList[iter["port"]]["type"]=='udp':
                print("send to",PortList[iter["port"]]["type"],"msg:",iter["send"]);
            elif PortList[iter["port"]]["type"]=='http':
                print("send to",PortList[iter["port"]]["type"],"msg:",iter["send"]);
            if str(iter['wait']) != "0":
                logmsg="wait {0}".format(iter['wait']);
                print(logmsg);
                time.sleep(float(iter['wait']))
        #for end
    #while end


    for iter in PortList:
        if "PortHandle" in iter:
            iter["PortHandle"].close();
            print("close {0} ".format(str(iter['port'][0])));

    LogF.close()


