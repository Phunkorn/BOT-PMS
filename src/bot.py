from pathlib import Path
import configparser, json, logging, sys, time
from excel_io import load_jobs, write_job_result, write_service_result
from tvc_driver import TVCDriver
from utils import stamp, ensure_dirs

BASE=Path(__file__).resolve().parents[1]
ensure_dirs(BASE)

def main():
    cfg=configparser.ConfigParser()
    cfg.read(BASE/"config.ini",encoding="utf-8")
    fmap=json.loads((BASE/"field_map.json").read_text(encoding="utf-8"))
    excel_path=BASE/cfg.get("excel","file")
    jobs=load_jobs(excel_path,cfg.get("excel","job_sheet"),cfg.get("excel","service_sheet"))
    if not jobs:
        print("ไม่พบ JOB ที่ bot_status = WAIT")
        return

    log_file=BASE/"logs"/f"bot_v051_{stamp()}.log"
    logging.basicConfig(level=logging.INFO,format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_file,encoding="utf-8"),logging.StreamHandler(sys.stdout)])

    tvc=TVCDriver(cfg.get("tvc","window_title_regex"),cfg.get("tvc","backend"),
                  cfg.getfloat("bot","step_delay"),cfg.getfloat("bot","typing_interval"))
    win=tvc.connect()
    print("พบหน้าต่าง:",win.window_text())

    for job in jobs:
        ref=str(job["job_ref"]).strip()
        row=int(job["_row"])
        try:
            write_job_result(excel_path,row,"RUNNING","Bot กำลังทำงาน")

            for f in ["plate_no","visit_no","chassis_no","warranty_no","brand_model","color",
                      "claim_no","reference_no","customer_name_2"]:
                v=job.get(f)
                if v is None or str(v).strip()=="":
                    print("SKIP",f); continue
                tvc.set_text(fmap[f],v); print("OK",f,"=",v)

            for f in ["job_date","return_date"]:
                v=job.get(f)
                if v is None or str(v).strip()=="":
                    print("SKIP",f); continue
                tvc.set_date(fmap[f],v); print("OK",f,"=",v)

            services=job.get("_services",[])
            if not services:
                raise RuntimeError("JOB นี้ไม่มีรายการใน SERVICE_INPUT")

            for svc in services:
                code=str(svc.get("service_code") or "").strip()
                before=tvc.list_count(fmap["service_list"])
                tvc.set_text(fmap["service_code"],code,press_enter=True)
                time.sleep(cfg.getfloat("bot","service_load_wait",fallback=0.8))
                qty=svc.get("qty")
                if qty is not None and str(qty).strip()!="":
                    tvc.set_text(fmap["service_qty"],qty)
                tvc.click(fmap["service_add_button"])
                time.sleep(cfg.getfloat("bot","service_load_wait",fallback=0.8))
                after=tvc.list_count(fmap["service_list"])
                if before is not None and after is not None and after<=before:
                    raise RuntimeError(f"กดเพิ่มแล้วจำนวนรายการไม่เพิ่ม สำหรับ {code}")
                write_service_result(excel_path,int(svc["_row"]),"ADDED",f"เพิ่ม {code} สำเร็จ")
                print("ADDED",code)

            tvc.screenshot(BASE/"screenshots"/f"BEFORE_SAVE_{ref}_{stamp()}.png")

            if cfg.getboolean("bot","save_enabled",fallback=True):
                print("SAVE FLOW v0.5.1: F10 -> Y(YES) -> N(NO)")
                result=tvc.save_flow_yes_then_no(
                    fmap["save_job_button"],
                    cfg.getfloat("bot","dialog_wait",fallback=1.0)
                )
                write_job_result(excel_path,row,"DONE",
                                 f"บันทึกสำเร็จ รอบแรก={result['first_clicked']} รอบสอง={result['second_clicked']}")
                print("DONE",ref,"| YES -> NO")
            else:
                write_job_result(excel_path,row,"DONE","กรอกครบแต่ยังไม่บันทึก")

        except Exception as e:
            logging.exception("JOB %s ERROR",ref)
            write_job_result(excel_path,row,"ERROR",str(e))
            try: tvc.screenshot(BASE/"screenshots"/f"ERROR_{ref}_{stamp()}.png")
            except Exception: pass
            print("ERROR:",e)
            break

    print("Log:",log_file)

if __name__=="__main__":
    main()
