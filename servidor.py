#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dashboard local 100% biblioteca padrão Python. Monitora XLSX do OneDrive.
Não modifica o Excel e não envia nenhum dado à internet.
"""
import argparse
import datetime as dt
import json
import mimetypes
import os
import posixpath
import re
import threading
import time
import unicodedata
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

DEFAULT_SOURCE = r'C:\Users\BR0101687168\OneDrive - Enel Spa\Documents\Programações_2026\Programações_2026.xlsx'
BASE_DIR = Path(__file__).resolve().parent
NS_URI = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P_URI = 'http://schemas.openxmlformats.org/package/2006/relationships'
M = '{' + NS_URI + '}'
R = '{' + R_URI + '}'
STANDARD_FIELDS = {
    'projeto': ('projeto','projetos','numero projeto','codigo projeto'),
    'data': ('data programacao','data de programacao','data programada','data'),
    'responsavel': ('responsavel','tecnico responsavel','tecnico'),
    'status_programacao': ('status programacao','status da programacao'),
    'tipo_intervencao': ('tipo intervencao','tipo de intervencao','intervencao'),
    'status_gilmar': ('status gilmar',),
}

def norm(s):
    s = ''.join(ch for ch in unicodedata.normalize('NFKD', str(s or '')) if not unicodedata.combining(ch))
    return re.sub(r'[^a-z0-9]+',' ',s.lower()).strip()

def column_number(ref):
    result = 0
    for ch in re.match(r'[A-Z]+',ref.upper()).group(0):
        result = result * 26 + ord(ch) - 64
    return result - 1

def parse_xlsx(path):
    """Lê arquivo .xlsx somente para leitura usando ZIP/XML, sem dependências extras."""
    output=[]; warnings=[]; names=[]; matched=0
    with zipfile.ZipFile(path,'r') as z:
        # Alguns arquivos gravados na nuvem podem estar temporariamente incompletos.
        valid = set(z.namelist())
        shared = []
        if 'xl/sharedStrings.xml' in valid:
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            shared = [''.join(t.text or '' for t in si.iter(M+'t')) for si in root.findall(M+'si')]
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rel_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rel_map = {x.get('Id'):x.get('Target') for x in rel_root.findall(P_URI and '{'+P_URI+'}Relationship')}
        date1904 = (wb.find(M+'workbookPr') is not None and wb.find(M+'workbookPr').get('date1904') in ('1','true'))
        date_styles = set()
        if 'xl/styles.xml' in valid:
            styles=ET.fromstring(z.read('xl/styles.xml'))
            fmt_el = styles.find(M+'numFmts')
            custom={int(x.get('numFmtId')):x.get('formatCode','') for x in (list(fmt_el) if fmt_el is not None else [])}
            xfs=styles.find(M+'cellXfs')
            for idx,x in enumerate(list(xfs) if xfs is not None else []):
                fmt_id=int(x.get('numFmtId','0'))
                if fmt_id in set(range(14,23))|{45,46,47} or (fmt_id in custom and re.search(r'[dymhs]',re.sub(r'"[^"]*"','',custom[fmt_id]),re.I)):
                    date_styles.add(idx)
        for sh in wb.findall('./'+M+'sheets/'+M+'sheet'):
            name=sh.get('name','Planilha')
            rid=sh.get(R+'id')
            target=rel_map.get(rid)
            if not target: continue
            if target.startswith('/'):
                sheet_path=target.lstrip('/')
            else:
                sheet_path=posixpath.normpath(posixpath.join('xl',target))
            if sheet_path not in valid:
                warnings.append(f'Aba {name}: não foi possível localizar os dados.')
                continue
            names.append(name)
            root=ET.fromstring(z.read(sheet_path))
            rows=[]
            for rw in root.findall('./'+M+'sheetData/'+M+'row'):
                values={}
                for cell in rw.findall(M+'c'):
                    ref=cell.get('r','')
                    if not ref:continue
                    idx=column_number(ref)
                    kind=cell.get('t','n')
                    v_el=cell.find(M+'v')
                    raw=v_el.text if v_el is not None else None
                    if kind=='s':
                        try: val=shared[int(raw)] if raw is not None else ''
                        except (ValueError,IndexError):val=''
                    elif kind=='inlineStr':
                        inline=cell.find(M+'is')
                        val=''.join(t.text or '' for t in inline.iter(M+'t')) if inline is not None else ''
                    elif kind=='b': val='Sim' if raw=='1' else 'Não'
                    elif kind=='e': val=''
                    elif raw is None:val=''
                    elif kind=='d':val=raw
                    else:
                        try:
                            numeric=float(raw)
                            if int(cell.get('s','0')) in date_styles:
                                base=dt.datetime(1904,1,1) if date1904 else dt.datetime(1899,12,30)
                                val=(base + dt.timedelta(days=numeric)).date().isoformat()
                            else:val= int(numeric) if numeric.is_integer() else numeric
                        except (ValueError,OverflowError):val=raw
                    values[idx]=val
                if values:rows.append((int(rw.get('r','0')),values))
            header_index=None; field_index={}
            for rowno, vals in rows[:35]:
                mapping={}
                for col, value in vals.items():
                    key=norm(value)
                    for field, options in STANDARD_FIELDS.items():
                        if key in options and field not in mapping:mapping[field]=col
                if all(x in mapping for x in ('projeto','data')) and len(mapping)>=4:
                    header_index=rowno;field_index=mapping;break
            if header_index is None:
                warnings.append(f'Aba {name}: ignorada porque não apresenta os cabeçalhos de programação.')
                continue
            matched+=1
            for rowno,vals in rows:
                if rowno<=header_index:continue
                record={}
                for field in STANDARD_FIELDS:
                    v=vals.get(field_index.get(field,-1),'')
                    record[field]=str(v).strip() if v is not None else ''
                # Não contar linhas completamente vazias ou apenas formatadas.
                if not any(record.values()):continue
                if not record['projeto'] and not record['data'] and not record['responsavel']:
                    continue
                original_date=record['data']
                if original_date:
                    try:
                        number=float(original_date)
                        base=dt.datetime(1904,1,1) if date1904 else dt.datetime(1899,12,30)
                        record['data']=(base+dt.timedelta(days=number)).date().isoformat()
                    except (ValueError,OverflowError):
                        date_text=original_date.strip()
                        parsed=False
                        for format_code in ('%d/%m/%Y','%Y-%m-%d','%d-%m-%Y','%Y-%m-%dT%H:%M:%S','%d/%m/%y'):
                            try:
                                record['data']=dt.datetime.strptime(date_text,format_code).date().isoformat();parsed=True;break
                            except ValueError:pass
                        if not parsed:
                            warnings.append(f'{name}, linha {rowno}: data não reconhecida: {date_text[:35]}')
                project=norm(record['projeto']).upper().replace(' ','')
                # Regional derivada do código do projeto, não é uma coluna original.
                match=re.search(r'(?:^|[/.])(LES|ABC|NOR|NTE)(?:[/.]|$)',record['projeto'].upper())
                record['regional']=match.group(1) if match else 'Não identificada'
                record['aba']=name
                record['linha']=rowno
                output.append(record)
    if not matched:
        raise ValueError('Nenhuma aba com os cabeçalhos Projeto, Data Programacao e campos de programação foi encontrada.')
    return output,names,warnings

class ExcelCache:
    def __init__(self, path):
        self.path=Path(path)
        self.lock=threading.RLock()
        self.signature=None
        self.last_good=None
        self.last_error=None
        self.error_time=None
        self.first=True
    def get(self):
        with self.lock:
            try:
                stat=self.path.stat()
                signature=f'{stat.st_mtime_ns}-{stat.st_size}'
            except FileNotFoundError:
                self.last_error=f'Arquivo não encontrado no caminho configurado: {self.path}'
                signature=None
            except OSError as ex:
                self.last_error=f'Erro ao acessar o Excel: {ex}'
                signature=None
            if signature is not None and (signature!=self.signature or self.last_good is None):
                try:
                    records,tabs,warnings=parse_xlsx(self.path)
                    self.last_good={
                        'revision':signature,
                        'records':records,
                        'sheets':tabs,
                        'warnings':warnings,
                        'read_at':dt.datetime.now().astimezone().isoformat(timespec='seconds'),
                        'file_modified_at':dt.datetime.fromtimestamp(stat.st_mtime).astimezone().isoformat(timespec='seconds'),
                        'source':str(self.path),
                        'count':len(records),
                    }
                    self.signature=signature
                    self.last_error=None
                    print(f'[Excel] Atualizado: {len(records)} programações | {len(tabs)} aba(s) | {self.last_good["read_at"]}',flush=True)
                except (OSError,KeyError,ValueError,zipfile.BadZipFile,ET.ParseError,IndexError) as ex:
                    self.last_error=f'Excel temporariamente indisponível ou inválido: {ex}'
                    self.error_time=dt.datetime.now().astimezone().isoformat(timespec='seconds')
                    if self.first or self.last_good is None: print('[Excel] '+self.last_error,flush=True)
            self.first=False
            if self.last_good is not None:
                return {'ok':True, **self.last_good, 'stale':bool(self.last_error), 'error':self.last_error}
            return {'ok':False,'records':[],'sheets':[],'warnings':[],'error':self.last_error or 'Aguardando arquivo Excel.', 'source':str(self.path),'revision':'none', 'stale':True}

class LocalHandler(BaseHTTPRequestHandler):
    cache=None
    def log_message(self, format, *args):
        if not self.path.startswith('/api'):super().log_message(format,*args)
    def send_json(self, payload,code=200):
        data=json.dumps(payload,ensure_ascii=False,separators=(',',':')).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type','application/json; charset=utf-8')
        self.send_header('Cache-Control','no-store, max-age=0')
        self.send_header('X-Content-Type-Options','nosniff')
        self.send_header('Content-Length',str(len(data)))
        self.end_headers();self.wfile.write(data)
    def do_GET(self):
        pathname=urlparse(self.path).path
        if pathname in ('/api/data','/api/status'):
            payload=self.cache.get()
            if pathname=='/api/status':
                payload={k:v for k,v in payload.items() if k!='records'}
            return self.send_json(payload)
        routes={'/':'index.html','/index.html':'index.html','/estilos.css':'estilos.css','/dashboard.js':'dashboard.js','/dados_iniciais.js':'dados_iniciais.js'}
        if pathname not in routes:
            self.send_error(404,'Página não encontrada');return
        file=BASE_DIR/routes[pathname]
        try:
            data=file.read_bytes()
        except OSError:
            self.send_error(404,'Arquivo não encontrado');return
        media_type=mimetypes.guess_type(file.name)[0] or 'application/octet-stream'
        self.send_response(200)
        self.send_header('Content-Type',media_type+'; charset=utf-8')
        self.send_header('Cache-Control','no-store, max-age=0')
        self.send_header('X-Content-Type-Options','nosniff')
        self.send_header('Content-Length',str(len(data)))
        self.end_headers();self.wfile.write(data)

def main():
    parser=argparse.ArgumentParser(description='Dashboard Programações 2026 - atualização automática Excel')
    parser.add_argument('--excel',default=os.getenv('EXCEL_PROGRAMACOES',DEFAULT_SOURCE),help='Caminho completo do .xlsx')
    parser.add_argument('--porta',type=int,default=8765,help='Porta localhost, padrão 8765')
    parser.add_argument('--sem-navegador',action='store_true')
    args=parser.parse_args()
    LocalHandler.cache=ExcelCache(args.excel)
    LocalHandler.cache.get()
    try:
        httpd=ThreadingHTTPServer(('127.0.0.1',args.porta),LocalHandler)
    except OSError as exc:
        print(f'Não foi possível iniciar a porta {args.porta}: {exc}',flush=True)
        raise SystemExit(1)
    url=f'http://127.0.0.1:{args.porta}/'
    print('\n'+'='*69)
    print(' DASHBOARD PROGRAMACOES 2026 - ATUALIZACAO AUTOMATICA')
    print(' Abrir: '+url)
    print(' Planilha: '+str(args.excel))
    print(' Atualização: navegador verifica o Excel a cada 5 segundos.')
    print(' Para encerrar: CTRL + C nesta janela.')
    print('='*69+'\n',flush=True)
    if not args.sem_navegador:
        threading.Timer(0.9,lambda:webbrowser.open(url)).start()
    try:httpd.serve_forever(poll_interval=0.3)
    except KeyboardInterrupt:pass
    finally:httpd.server_close()

if __name__=='__main__':main()
