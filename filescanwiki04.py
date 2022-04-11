# -*- coding: shift_jis -*-

from __future__ import print_function

import sys
import os
import time
import sqlite3
import xlrd
import win32com.client

import psycopg2

SCAN_FILEEXT = [".txt", ".xls", ".xlsx", ".doc", ".docx"]

#SCANNER_DB_PATH = 'C:\\trac\\mydata02.db'
#SCANNER_DB_PATH = 'C:\\py_virenv\\test\\testenv1\\trac\\mydata02.db'
SCANNER_DB_PATH = 'C:\\py_virenv\\trac1.2env\\trac\\mydata02.db'

TMP_TXT_FOR_WORD = "c:\\tmp\\tmpword.txt"

LOG_STD = "filescanwiki04.log"
LOG_ERR = "filescanwiki04_err.log"


def pathNormalize(fpath):
    """�������A�啶���̕Њ�A�p�X�̕W����
    """
    return os.path.normpath(os.path.normcase(fpath))


# def fild_all_files(directory):
#     for root, dirs, files in os.walk(directory):
#         yield root
#         for f in files:
#             yield os.path.join(root, f)

def getNextSibling(curDir, curFile):
    """���̌Z��G���g����Ԃ�
    Args:
        curDir: �f�B���N�g��
        curFile: ���݂̃t�@�C��
    Returns:
        curDir�z���̃f�B���N�g�����ŁAcurFile��
        ��(�t�@�C�����\�[�g����)�̃t�@�C����Ԃ��B
        �t�@�C�����Ȃ����None��Ԃ�
    """

    lstDirEnt = map(pathNormalize, sorted(os.listdir(curDir)))

    pos = lstDirEnt.index(curFile)

    if len(lstDirEnt) == pos + 1:
        return None
    else:
        return pathNormalize(lstDirEnt[pos + 1])


def getFirstChild(curDir):
    """�f�B���N�g�����̐擪�̃t�@�C����Ԃ�
    Args:
        curDir: �f�B���N�g��
    Returns:
        curDir�z���̃f�B���N�g�����ŁA
        (�t�@�C�����\�[�g����)�擪�̃t�@�C����Ԃ��B
        �t�@�C�����Ȃ����None��Ԃ�
    """
    if os.path.isfile(curDir):
        print("getFirstChild Error. %s is File." % curDir)
        return None

    lstDirEnt = sorted(os.listdir(curDir))
    if len(lstDirEnt) == 0:
        return None
    
    return pathNormalize(lstDirEnt[0])


def getParentNextSibling(baseDir, curDir):
    """���݂̃f�B���N�g���̎��̃G���g����Ԃ�
    ���̃G���g�����Ȃ���΍ċA�I�ɐe�ɑk���Ă���
    baseDir�܂ők������None��Ԃ��Ĕ�����
    Args:
        baseDir: ��f�B���N�g��
        curDir: ���݂̃f�B���N�g��
    Returns:
        (nextDir, nextFile)��Ԃ�
        nextDir: ���̃f�B���N�g��
        nextFile: ���̃t�@�C��
        baseDir�z���̃t�@�C�������ׂĒT�������ꍇ�́A
        (None, None)��Ԃ�
    """
    #print("    getParentNextSibling() called. baseDir=[%s], curDir=[%s]" %
    #      (baseDir, curDir))
    
    (curDir, curFile) = os.path.split(curDir)
    nf = getNextSibling(curDir, curFile)
    
    if nf is None:
        logstd("    No Entry left in the directory, %s" % curDir)
        logstd("        baseDir=[%s], curDir=[%s]" % (baseDir, curDir))
        if baseDir == curDir:
            print("    baseDir[%s] scanning done." % baseDir)
            return (None, None)
        return getParentNextSibling(baseDir, curDir)
    else:
        return (curDir, nf)


def getNextEntry(baseDir, curDir, curFile, skipfiles):
    """���̃G���g����Ԃ�
    Args:
        baseDir: ��f�B���N�g��
        curDir: ���݂̃f�B���N�g��
        curFile: ���݂̃t�@�C��
        skipfiles: �X�L�b�v����t�@�C���i�f�B���N�g���܂ށj
    Returns:
        (nextDir, nextFile)��Ԃ�
        nextDir: ���̃f�B���N�g��
        nextFile: ���̃t�@�C��
        baseDir�z���̃t�@�C�������ׂĒT�������ꍇ�́A
        (None, None)��Ԃ�
    """
    filepath = pathNormalize(os.path.join(curDir, curFile))

    #print("    getNextEntry called. filepath=%s" % filepath)
    nf2 = None

    if os.path.isfile(filepath):
        nf = getNextSibling(curDir, curFile)
        if nf is None:
            if baseDir == curDir:
                print("    baseDir[%s] scanning done." % baseDir)
                return (None, None)

            nf2 = getParentNextSibling(baseDir, curDir)
        else:
            nf2 = (curDir, nf)
    else:
        print("    %s is Directory." % filepath)
        nf = getFirstChild(filepath)

        if (filepath in skipfiles) or (nf is None):
            logstd("%s: No Children or skip-directory" % filepath)

            nf = getNextSibling(curDir, curFile)
            if nf is None:
                if baseDir == curDir:
                    print("    baseDir[%s] scanning done." % baseDir)
                    return (None, None)
                
                nf2 = getParentNextSibling(baseDir, curDir)
            else:
                nf2 = (curDir, nf) 
        else:
            nf2 = (filepath, nf)

	if not (nf2[0] is None):
	    filepath2 = os.path.join(nf2[0], nf2[1])
    	    #print("filepath2=" + filepath2)
	    if filepath2 in skipfiles:
                logstd(filepath2 + " is excluded by t_skip_file.")
        	return getNextEntry(baseDir, nf2[0], nf2[1], skipfiles)

    return nf2


def getFileExtension(filepath):
    """�t�@�C���̊g���q���擾����
    Args: 
        filepath: �t�@�C����
    """
    periodpos = filepath.rfind('.')
    if periodpos > 0:
        return filepath[periodpos:]
    else:
        return ""


def get_file_contents(filepath):
    """�t�@�C���̒��g��ǂݎ��A���e��Ԃ�
    Args:
        filepath: �t�@�C���̃t���p�X
    Returns:
        �t�@�C���̒��g
    """

    f = open(filepath)
    s = f.read()
    f.close()

    flg_success = False

    for encoding in ['cp932', 'utf-8']:
        try:
            #print encoding
            us = s.decode(encoding)
            flg_success = True
            break
        except UnicodeDecodeError:
            pass
            
    if not flg_success:
        #print("all of trying 's.decode()' is failed  .", file=sys.stderr)
        logerr("filepath: %s" % filepath)
        logerr("all of trying 's.decode()' is failed  .")
        #raise Exception("all of trying 's.decode()' is failed  .")
        us = ""
        
    return us


def get_excel_contents(filepath):
    """Excel�t�@�C���̒��g��ǂݎ��A���e��Ԃ�
    Args:
        filepath: �t�@�C���̃t���p�X
    Returns:
        �t�@�C���̒��g
    """
    book = None
    try:
        book = xlrd.open_workbook(filepath)
    except xlrd.biffh.XLRDError as e:
        logstd("%s cause XLRDError. %s" % (filepath, str(e)))
        logerr("%s cause XLRDError. %s" % (filepath, str(e)))
        return ""
    except UnicodeDecodeError as e:
        logstd("%s cause UnicodeDecodeError. %s" % (filepath, str(e)))
        logerr("%s cause UnicodeDecodeError. %s" % (filepath, str(e)))
        return ""
    except ValueError as e:
        logstd("%s cause ValueError. %s" % (filepath, str(e)))
        logerr("%s cause ValueError. %s" % (filepath, str(e)))
        return ""
        
    content = ""

    for shtName in book.sheet_names():
        content += "==========================\n"
        content += shtName + "\n"
        content += "--------------------------\n"
        sht = book.sheet_by_name(shtName)

        for row in range(sht.nrows):
            for col in range(sht.ncols):

                try:
                    cellvalue = sht.cell(row,col).value
                except IndexError:
                    cellvalue = ""
                
                cs = ""
                if isinstance(cellvalue, str):
                    cs = cellvalue.decode('cp932')
                elif isinstance(cellvalue, unicode):
                    cs = cellvalue
                else:
                    cs = str(cellvalue)

                content += cs + ","

            content += "\n"

    book.release_resources()

    return content


def getNewWikiPageName(connScanner):
    '''�V�����y�[�W�����蓖�Ă�ꍇ��wiki�y�[�W�̃L�[���擾
    Args: 
        connScanner: scannerDB�̃R�l�N�V����
    Returns:
        �V����wiki�y�[�W�̃L�[
    '''
    sql = "select wikiPageName from t_seq"

    cur = connScanner.cursor()

    mappage_seq = cur.execute(sql).fetchone()[0]

    sql = "update t_seq set wikiPageName = ?"

    connScanner.execute(sql, (mappage_seq + 1,))

    return "__MapPage%08d" % (mappage_seq)


def registerScannedFile(connScanner, scannedFilePath):
    '''�X�L���������t�@�C���̏���o�^����
    Args: 
        connScanner: scannerDB�̃R�l�N�V����
    Returns:
        �Ώۃt�@�C����wikiPageName
    '''

    sql = "select wikiPageName from t_scan_file where fpath=?"

    cur = connScanner.cursor()
    row = cur.execute(sql, (scannedFilePath, )).fetchone()

    pageName = ""
    
    if row is None:
        #�V�����G���g����������scan_dir�Ƀ��R�[�h��ǉ�
        pageName = getNewWikiPageName(connScanner)
    
        sql = "insert into t_scan_file(fpath, last_checked, wikiPageName) "
        sql += " values(?, ?, ?) "
        connScanner.execute(sql, (scannedFilePath, time.time() * 1000000, pageName))
    else:
        #�����G���g����������scan_dir��last_checked���������ōX�V
        pageName = row[0]

        sql = "update t_scan_file set last_checked = ? "
        sql += "where fpath = ? "
        connScanner.execute(sql, (time.time() * 1000000, scannedFilePath))

    return pageName


def registerFile(filepath, tracDb, connScanner, msWordRdr):
    '''1) �t�@�C���̓��e���X�L��������
       2) �X�L���������t�@�C���̏���o�^����
       3) �t�@�C���̓��e��Trac DB�ɓo�^����
    Args: 
        filepath: �X�L��������t�@�C��
        connTrac: TracDB�̃R�l�N�V����
        connScanner: scannerDB�̃R�l�N�V����
        msWordRdr: MsWordReader�N���X�̃C���X�^���X
    Returns:
        �i�Ȃ��j
    '''

    try:
        fileext = getFileExtension(filepath)
        pageData = ""
        if fileext in [".xls", ".xlsx"]:
            pageData = get_excel_contents(filepath)
        elif fileext in [".doc", "docx"]:
            pageData = msWordRdr.readTextData(filepath)
        else:
            pageData = get_file_contents(filepath)
        
        #print "-- start contents --"
        #print pageData
        #print "-- end contents --"
        
        # ���g���Ȃ���Β��f
        if 0 == len(pageData):
            return
        
        pageName = registerScannedFile(connScanner, filepath)
        logstd("%s regisgered to wiki as %s." % (filepath, pageName))
        tracDb.register2TracDb(pageName, pageData, filepath)

        connScanner.commit()
        tracDb.conn.commit()

    except:
        print("filepath: " + filepath)
        print("Unexpected error:", sys.exc_info()[0])
        connScanner.rollback()
        tracDb.conn.rollback()
        raise
        

#
# t_scan_status�ŏ������̃t�@�C���̏����Ǘ�
#

def initScanStatus(connScanner, baseDir, curDir, curFile):
    sql = "insert into t_scan_status(base_dir, cur_dir, cur_file) "
    sql += " values(?, ?, ?) "
    connScanner.execute(sql, (baseDir, curDir, curFile))
    connScanner.commit()

def updateScanStatus(connScanner, curDir, curFile):
    sql = "update t_scan_status set cur_dir = ?, cur_file = ? "
    connScanner.execute(sql, (curDir, curFile))
    connScanner.commit();

def clearScanStatus(connScanner):
    sql = "delete from t_scan_status "
    connScanner.execute(sql);
    connScanner.commit()

def getScanStatus(connScanner):
    sql = "select base_dir, cur_dir, cur_file "
    sql += " from t_scan_status"

    row = connScanner.execute(sql).fetchone()

    if row is None:
        return (None, None, None)
    else:
        return (row[0], row[1], row[2])

class TracDb:
    '''TracDb�N���X

    '''

    def __init__(self):
        self.conn = psycopg2.connect(
            "dbname=trac2 host=localhost user=tracuser password=tracuser")

    def register2TracDb(self, pageName, pageData, filepath):
        '''�t�@�C���̓��e��Trac DB�ɓo�^����
        Args: 
            connTrac: TracDB�̃R�l�N�V����
            pageName: wiki�y�[�W�L�[
            pageData: �t�@�C���̓��e
        Returns:
            �i�Ȃ��j
        '''
        sql = "select count(*) from wiki where name=%s;"
        cur = self.conn.cursor()
        cur.execute(sql, (pageName,))
        row = cur.fetchone()
        recCnt = row[0]
        cur.close()
        #print sql
        #print recCnt

        wikiContent = u"filepath: {{{ " + filepath + u" }}}\n"
        wikiContent += u"{{{\n" + pageData + u"\n}}}\n "
    
        if recCnt == 0:
            sql = "insert into WIKI \n"
            sql += "  (NAME,VERSION,TIME,AUTHOR,IPNR,TEXT,COMMENT,READONLY) \n"
            sql += "values(%s, 1, %s, 'filescanner', '127.0.0.1', %s, '', 0) "

            cur2 = self.conn.cursor()
            cur2.execute(sql, (pageName, time.time() * 1000000, wikiContent))
            #self.conn.commit()
            cur2.close()
        else:
            sql = "update wiki set text=%s, time=%s where name=%s"

            #print sql
            #print time.time()

            #�����͂Ȃ�1M�{����Ƃ��܂��s�����͕s��
            cur2 = self.conn.cursor()
            cur2.execute(sql, (wikiContent, time.time() * 1000000, pageName))
            #self.conn.commit()
            cur2.close()


class ScannerDb:
    '''ScannerDb�N���X

    �ڐA�̓r���A�����Â��̃N���X�ɂ܂Ƃ߂�
    SCANNER_DB_PATH�Ƀt�@�C���p�X�̐ݒ������
    '''
    def __init__(self):
        self.conn = sqlite3.connect(SCANNER_DB_PATH)
        self.loadSkipfile()
        
    def needScan(self, fpath):
        '''�t�@�C���̒��g���X�L��������K�v�����邩�`�F�b�N

        Args:
            fpath: �Ώۂ̃t�@�C���i�t���p�X�j
        Returns:
            �t�@�C���̃^�C���X�^���v��t_scan_file.last_checked���
            �傫��������A�X�L�������K�v�Ɣ��肵�ATrue��Ԃ��B
            ����ȊO��False��Ԃ�
        '''

        sql = "select last_checked from t_scan_file where fpath=?"

        cur = self.conn.cursor()
        row = cur.execute(sql, (fpath, )).fetchone()

        if row is None:
            return True
        else:
            fpath_tstmp = os.stat(fpath).st_mtime
            last_checked = row[0] / 1000000 #1M�{���ēo�^���Ă���̂Ŗ߂�
            #print("fpath_tstmp=%f, last_checked=%f" % (fpath_tstmp, last_checked))
            return (fpath_tstmp > last_checked)

    def loadSkipfile(self):
        '''skip�Ώہi�t�H���_�^�t�@�C���̂ǂ���ł�OK)��ǂݍ���
        
        '''
        sql = "select fpath from t_skip_file"
        cur = self.conn.cursor()
        cur.execute(sql)
        self.skipfiles = []
        for row in cur:
            fpath = pathNormalize(row[0])
            print("skipfiles=%s" % fpath)
            self.skipfiles.append(fpath)

def logstd(txt):
    #glbLogStdF.write(txt + u'\n')
    glbLogStdF.write(txt.encode('cp932'))
    glbLogStdF.write('\n')
    glbLogStdF.flush()

def logerr(txt):
    glbLogErrF.write(txt.encode('cp932'))
    glbLogErrF.write('\n')
    glbLogErrF.flush()
    

class MsWordReader:
    '''MS WORD�����N���X
    
    MS WORD�t�@�C�����e�L�X�g�Ƃ��ēǂݎ��N���X

    TMP_TXT_FOR_WORD���ꎞ�t�@�C���Ƃ��Ďg�p����B�v�ݒ�
    '''

    def __init__(self):
        pass

    def startWord(self):
        self.wordApp = win32com.client.gencache.EnsureDispatch("Word.Application")
    def quitWord(self):
        self.wordApp.Quit()

    def saveToTxt(self, filepath):
        self.wordApp.Documents.Open(filepath)
        self.wordApp.ActiveDocument.SaveAs(
            TMP_TXT_FOR_WORD,
            FileFormat=win32com.client.constants.wdFormatText)
        self.wordApp.ActiveDocument.Close()
        
    def readTextData(self, filepath):
        '''Reads text data from MSWord file

        Args:
            filepath: ms-word file(*.doc, *.docx)
        Returns:
            text data from ms-word file.
        '''

        self.saveToTxt(filepath)
        return get_file_contents(TMP_TXT_FOR_WORD)


#############################
# main routine
#############################
if 2 > len(sys.argv):
    sys.exit("Usage: python filescanwiki04.py (directory path)")

#�������A�啶���̕Њ�A�p�X�̕W����
#baseDir = os.path.normcase(sys.argv[1].decode('cp932'))
#baseDir = os.path.normpath(baseDir)
baseDir = pathNormalize(sys.argv[1].decode('cp932'))

#baseDir = u'c:\\tmp\\work03'

#���O�t�@�C���̃n���h���I�[�v��
glbLogStdF = open(LOG_STD, 'w')
glbLogErrF = open(LOG_ERR, 'w')

print(u'�i�x�[�X�f�B���N�g���jbase_dir=%s' % baseDir)
logstd(u'�i�x�[�X�f�B���N�g���jbase_dir=%s' % baseDir)

# #Test Routine.


# for r in scannerDb.skipfiles:
#     print(r)

# print(scannerDb.isSkippingOver('c:\\tmp\\work03\\text01.txt'))
# print(scannerDb.isSkippingOver('c:\\tmp\\work03\\text05.txt'))
    
#sys.exit("test end.")

# #Test Routine End.

tracDb = TracDb()
connTrac = tracDb.conn
scannerDb = ScannerDb()
connScanner = scannerDb.conn

#connTrac = sqlite3.connect('C:\\trac\\wachipj\\db\\trac.db')

(baseDir2, curDir, curFile) = getScanStatus(connScanner)

print("baseDir2=%s, curDir=%s, curFile=%s" % (baseDir2, curDir, curFile))

if baseDir2 is None:
    curDir = baseDir
    curFile = getFirstChild(curDir)
    initScanStatus(connScanner, baseDir, curDir, curFile)
else:
    baseDir = baseDir2

msWordRdr = MsWordReader()
msWordRdr.startWord()

waitCnt = 0
for i in range(1, 10000):

    #print isinstance(f, unicode)  #print True
    #print filepath

    updateScanStatus(connScanner, curDir, curFile)
    
    filepath = os.path.join(curDir, curFile)

    if os.path.isfile(filepath):
        #�g���q���Y�����邩�`�F�b�N���A�Y�������
        # �t�@�C���̒��g��o�^���Ascannerdb, tracdb�ɓo�^
        fileext = getFileExtension(filepath)
        if (0 < len(fileext)) and (fileext in SCAN_FILEEXT):
            if scannerDb.needScan(filepath):
                #registerFile(filepath, connTrac, connScanner, msWordRdr)
                registerFile(filepath, tracDb, connScanner, msWordRdr)
            else:
                logstd(filepath + " is skipped(already scanned).")
        else:
            #print(filepath + " is excluded.")
            logstd(filepath + " is excluded.")


    #logstd("curDir=%s, curFile=%s" % (curDir, curFile))
    (nextDir, nextFile) = getNextEntry(baseDir, curDir, curFile, scannerDb.skipfiles)
    #logstd("nextDir=%s, nextFile=%s" % (nextDir, nextFile))

    if nextDir is None:
        break

    #
    # wait 0.5sec each 10 entries processed.
    #
    if waitCnt >= 10:
        time.sleep(0.5)
        waitCnt = 0
        print("wait 0.5sec...")
    waitCnt += 1
    
    curDir = nextDir
    curFile = nextFile

    if i == 9999:
        print("loop exausted(9999).")
        break

msWordRdr.quitWord()

clearScanStatus(connScanner)

connScanner.close()
connTrac.close()

glbLogStdF.close()
glbLogErrF.close()
