from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

def _table_rows(ws, table_name):
    if table_name not in ws.tables:
        raise RuntimeError(f"ไม่พบ Excel Table {table_name}")
    min_col,min_row,max_col,max_row = range_boundaries(ws.tables[table_name].ref)
    headers=[ws.cell(min_row,c).value for c in range(min_col,max_col+1)]
    for row_idx in range(min_row+1,max_row+1):
        vals=[ws.cell(row_idx,c).value for c in range(min_col,max_col+1)]
        if not any(v is not None and str(v).strip() for v in vals):
            continue
        d={headers[i]:vals[i] for i in range(len(headers))}
        d["_row"]=row_idx
        yield d

def load_jobs(path: Path, job_sheet="JOB_INPUT", service_sheet="SERVICE_INPUT"):
    wb=load_workbook(path)
    jws=wb[job_sheet]; sws=wb[service_sheet]
    services=list(_table_rows(sws,"ServiceInputV5Table"))
    jobs=[]
    for job in _table_rows(jws,"JobInputV5Table"):
        ref=str(job.get("job_ref") or "").strip()
        status=str(job.get("bot_status") or "").strip().upper()
        if not ref or status!="WAIT":
            continue
        matched=[s for s in services
                 if str(s.get("job_ref") or "").strip()==ref
                 and str(s.get("service_status") or "").strip().upper() in {"WAIT",""}]
        matched.sort(key=lambda x:int(x.get("service_seq") or 999999))
        job["_services"]=matched
        jobs.append(job)
    return jobs

def write_job_result(path: Path,row,status,result="",tvc_job_no=""):
    wb=load_workbook(path); ws=wb["JOB_INPUT"]
    h={c.value:i+1 for i,c in enumerate(ws[1])}
    ws.cell(row=row,column=h["bot_status"]).value=status
    ws.cell(row=row,column=h["bot_result"]).value=result
    if tvc_job_no and "tvc_job_no" in h:
        ws.cell(row=row,column=h["tvc_job_no"]).value=tvc_job_no
    wb.save(path)

def write_service_result(path: Path,row,status,result=""):
    wb=load_workbook(path); ws=wb["SERVICE_INPUT"]
    h={c.value:i+1 for i,c in enumerate(ws[1])}
    ws.cell(row=row,column=h["service_status"]).value=status
    ws.cell(row=row,column=h["service_result"]).value=result
    wb.save(path)
