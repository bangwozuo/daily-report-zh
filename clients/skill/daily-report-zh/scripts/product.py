"""Standard-library product SDK. Fixed local deployment, no Agent/model loop."""
from urllib.request import Request,build_opener,HTTPCookieProcessor,ProxyHandler,HTTPRedirectHandler
from http.cookiejar import CookieJar
from urllib.parse import quote
import argparse,json,re,sys
from pathlib import Path

class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self,*args,**kwargs):return None

class Client:
    def __init__(self,connection):
        self.c=connection
        if set(connection)!={'origin','project_id','agent_id','installation_id','manifest_sha256'}:raise ValueError('连接字段不完整')
        if connection['origin']!='http://127.0.0.1:8940':raise ValueError('未授权的服务目的地')
        for key in ('project_id','agent_id','installation_id'):
            if not re.fullmatch('[A-Za-z0-9_-]{1,100}',connection[key]):raise ValueError('连接身份格式错误')
        if not re.fullmatch('[a-f0-9]{64}',connection['manifest_sha256']):raise ValueError('安装指纹缺失')
        if connection['project_id']!='prj_e9adecb93e4845b3b157' or connection['agent_id']!='daily-report-zh':raise ValueError('Wrong installed employee identity')
        self.http=build_opener(ProxyHandler({}),HTTPCookieProcessor(CookieJar()),NoRedirect())
        self.prefix='/daily-role/v1'
    def _request(self,path,body=None):
        request=Request(self.c['origin']+path,data=json.dumps(body,ensure_ascii=False).encode() if body is not None else None,
            headers={'Content-Type':'application/json','Origin':self.c['origin']})
        with self.http.open(request,timeout=180)as response:return json.load(response)
    def contract(self):
        value=self._request(self.prefix+'/contract')
        if value['installation_id']!=self.c['installation_id'] or value['manifest_sha256']!=self.c['manifest_sha256']:
            raise ValueError('服务安装已变化；不切换版本，不发送业务')
        return value
    def submit(self,operation_id,fields,text):
        if not re.fullmatch('[A-Za-z0-9][A-Za-z0-9_.:-]{11,159}',operation_id):raise ValueError('必须保留一个准确原操作编号')
        contract=self.contract()
        if 'submit'not in contract['allowed_actions']:raise ValueError('当前产品未获准办理')
        return self._request(self.prefix+'/operations',{'business_operation_id':operation_id,
            'expected_installation_id':self.c['installation_id'],'fields':fields,
            'chat':{'text':text,'fields':fields,'attachments':[],'instructions':[]}})
    def operation(self,operation_id):
        if not re.fullmatch('[A-Za-z0-9][A-Za-z0-9_.:-]{0,159}',operation_id):raise ValueError('原操作编号格式错误')
        value=self._request(self.prefix+'/operations/'+quote(operation_id,safe=''))
        if value['installation_id']!=self.c['installation_id']:raise ValueError('原操作属于其他安装')
        return value
    def review(self,operation_id,decision,expected_revision,notes):
        if decision not in {'revise','confirm','reject'}:raise ValueError('核对动作不支持')
        self.operation(operation_id)
        return self._request(self.prefix+'/operations/'+quote(operation_id,safe='')+'/review',
            {'action':decision,'expected_revision':expected_revision,'notes':notes})

def main():
    p=argparse.ArgumentParser();p.add_argument('action',choices=('contract','submit','lookup','review'))
    p.add_argument('--connection',type=Path,required=True);p.add_argument('--operation');p.add_argument('--input',type=Path)
    p.add_argument('--decision',choices=('revise','confirm','reject'));p.add_argument('--expected-revision',type=int);p.add_argument('--notes')
    a=p.parse_args();client=Client(json.loads(a.connection.read_bytes()))
    if a.action=='contract':value=client.contract()
    elif a.action=='lookup':value=client.operation(a.operation)
    elif a.action=='submit':
        value=json.loads(a.input.read_bytes())
        if set(value)!={'fields','text'}:raise ValueError('输入只允许业务字段与任务原文')
        value=client.submit(a.operation,value['fields'],value['text'])
    else:value=client.review(a.operation,a.decision,a.expected_revision,a.notes)
    print(json.dumps(value,ensure_ascii=False))
if __name__=='__main__':main()
