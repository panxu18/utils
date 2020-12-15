import requests
import re

def getIdAndName(cardId,dwId,headers):
    p=re.compile(r'.+XM=(.+)GRSXH=(.+)GMSFHM',re.I)
    url="http://221.232.64.242:7101/dwws/action/AjaxAction?ActionType=ajax_common_proc"
    data={
        "confId": "_UserProc",
        "ProcName": "pkg_ajax.dwws_common_grjbzl",
        "FunName":"MyFunGRSXH",
        "Params":"GRBS="+cardId+"%03XTJGDM=1008%03DWSXH="+dwId
    }
    r=requests.post(url=url,data=data,headers=headers)
    string=r.text
    me=p.match(str(string.strip()))
    if me is None:
        print("getIdAndName error", string)
        return ("","")
    else:
        xm=me.group(1).replace("\x03","")
        gr=me.group(2).replace("\x03","")
        return (xm,gr)

def getFile(tid,headers):
    baseurl="http://221.232.64.242:7101/dwws/action/MainAction"
    data={
        "ActionType":"printtest",
        "isUsePOIExport":True,
        "BAE001":"10",
        "whereCls":"where g.GRSXH="+tid+" and SHNY between'201909' and '202008' and DWJFBZ in ('1','2') and XZLX in ('10','20','30','40','50','53','54') order by JFNY desc",
        "uploadType":"dwryjfxxcx_export"
    }
    res=requests.post(url=baseurl,data=data,headers=headers)
    if res.status_code==200:
        return res.content
    else:
        return b''

def grjf_dowmload(dwId, headers, inputFile, outputPath):
    f=open(inputFile)
    cardList=[]
    for line in f:
        cardList.append(line.replace("\n",""))
    f.close()
    for cardId in cardList:
        t=getIdAndName(cardId,dwId,headers)
        if(t[0]==""):
            print("error: ", cardId)
            continue
        else:
            print(cardId, t)
            content=getFile(t[1],headers)
            with open(outputPath + "/"+t[0]+".xls","wb") as f1:
                f1.write(content)


if __name__=="__main__":
    # 公司ID
    dwId = ""
    # 公司账号cookie
    cookie = ""
    # 身份证号文件
    input = ""
    # 下载目录
    downloadPath = ""
    headers={
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36",
        "Cookie":cookie ,
        "Content-Type":"application/x-www-form-urlencoded"
    }
    grjf_dowmload(dwId, headers, input , downloadPath )
