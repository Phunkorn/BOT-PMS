from pathlib import Path
import argparse, configparser, json, logging, sys, time
from excel_io import (
    COMMITTING_TVC,
    TVC_SAVED_PENDING_EXCEL,
    UNCERTAIN_TVC_SAVE,
    load_jobs,
    write_job_result,
    write_service_result,
)
from tvc_driver import TVCDriver
from utils import stamp
from version import APP_NAME, APP_VERSION
from runtime_paths import resolve_runtime_paths

RUNTIME_PATHS=resolve_runtime_paths()
BASE=RUNTIME_PATHS.app_dir
RUNTIME_PATHS.logs_dir.mkdir(parents=True,exist_ok=True)
RUNTIME_PATHS.screenshots_dir.mkdir(parents=True,exist_ok=True)

EVENT_PREFIX="TVCBOT_EVENT "


def parse_args(argv=None):
    parser=argparse.ArgumentParser(description=f"{APP_NAME} v{APP_VERSION}")
    parser.add_argument("--excel",help="Path to the .xlsx input file")
    parser.add_argument("--stop-file",help="Internal cooperative stop signal file")
    return parser.parse_args(argv)


def emit_event(event,**values):
    payload={"event":event,**values}
    print(EVENT_PREFIX+json.dumps(payload,ensure_ascii=False),flush=True)


def stop_requested(stop_file):
    return stop_file is not None and stop_file.exists()


def main(argv=None):
    args=parse_args(argv)
    cfg=configparser.ConfigParser()
    cfg.read(RUNTIME_PATHS.config_file,encoding="utf-8")
    fmap=json.loads(RUNTIME_PATHS.field_map_file.read_text(encoding="utf-8"))
    if args.excel:
        excel_path=Path(args.excel).expanduser().resolve()
    else:
        excel_path=(BASE/cfg.get("excel","file")).resolve()
    if not excel_path.is_file():
        raise FileNotFoundError(f"ไม่พบไฟล์ Excel: {excel_path}")
    if excel_path.suffix.lower()!=".xlsx":
        raise ValueError("รองรับเฉพาะไฟล์ Excel .xlsx")
    stop_file=Path(args.stop_file).expanduser().resolve() if args.stop_file else None

    jobs=load_jobs(excel_path,cfg.get("excel","job_sheet"),cfg.get("excel","service_sheet"))
    if not jobs:
        print("ไม่พบ JOB ที่ bot_status = WAIT")
        emit_event("batch_complete",total=0,success=True)
        return 0

    version_tag=APP_VERSION.replace(".","")
    log_file=RUNTIME_PATHS.logs_dir/f"bot_v{version_tag}_{stamp()}.log"
    logging.basicConfig(level=logging.INFO,format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_file,encoding="utf-8"),logging.StreamHandler(sys.stdout)])

    total_jobs=len(jobs)
    emit_event("batch_start",total=total_jobs,excel=str(excel_path),log=str(log_file))

    # Intentional Stop contract: checks happen only at JOB boundaries.
    # Never interrupt fields, services, duplicate flow, F10, or YES/NO mid-JOB.
    if stop_requested(stop_file):
        print("STOP: ผู้ใช้สั่งหยุดก่อนเริ่ม JOB")
        emit_event("stopped",ref="",phase="before_job")
        return 2

    tvc=TVCDriver(cfg.get("tvc","window_title_regex"),cfg.get("tvc","backend"),
                  cfg.getfloat("bot","step_delay"),cfg.getfloat("bot","typing_interval"))
    win=tvc.connect()
    print("พบหน้าต่าง:",win.window_text())

    if stop_requested(stop_file):
        print("STOP: ผู้ใช้สั่งหยุดก่อนเริ่ม JOB")
        emit_event("stopped",ref="",phase="before_job")
        return 2

    had_error=False
    for job_index,job in enumerate(jobs,start=1):
        ref=str(job["job_ref"]).strip()
        row=int(job["_row"])

        if stop_requested(stop_file):
            print("STOP: หยุดก่อนเริ่ม JOB ถัดไป")
            emit_event("stopped",ref="",phase="before_job",index=job_index,total=total_jobs)
            return 2

        commit_started=False
        try:
            write_job_result(excel_path,row,"RUNNING","Bot กำลังทำงาน")
            emit_event("job_start",ref=ref,index=job_index,total=total_jobs)

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

            tvc.screenshot(RUNTIME_PATHS.screenshots_dir/f"BEFORE_SAVE_{ref}_{stamp()}.png")

            if cfg.getboolean("bot","save_enabled",fallback=True):
                write_job_result(excel_path,row,"RUNNING",COMMITTING_TVC)
                commit_started=True
                emit_event("commit_state",ref=ref,state=COMMITTING_TVC)
                print("SAVE FLOW v0.5.1: F10 -> Y(YES) -> N(NO)")
                result=tvc.save_flow_yes_then_no(
                    fmap["save_job_button"],
                    cfg.getfloat("bot","dialog_wait",fallback=1.0)
                )
                write_job_result(excel_path,row,"RUNNING",TVC_SAVED_PENDING_EXCEL)
                emit_event("commit_state",ref=ref,state=TVC_SAVED_PENDING_EXCEL)
                write_job_result(excel_path,row,"DONE",
                                 f"บันทึกสำเร็จ รอบแรก={result['first_clicked']} รอบสอง={result['second_clicked']}")
                print("DONE",ref,"| YES -> NO")
            else:
                write_job_result(excel_path,row,"DONE","กรอกครบแต่ยังไม่บันทึก")
            emit_event("job_done",ref=ref,index=job_index,total=total_jobs)

            if stop_requested(stop_file):
                print("STOP: JOB ปัจจุบันเสร็จแล้ว หยุดก่อนเริ่ม JOB ถัดไป")
                emit_event(
                    "stopped",
                    ref=ref,
                    phase="after_job",
                    index=job_index,
                    total=total_jobs,
                )
                return 2

        except Exception as e:
            logging.exception("JOB %s ERROR",ref)
            error_result=UNCERTAIN_TVC_SAVE if commit_started else str(e)
            try:
                write_job_result(excel_path,row,"ERROR",error_result)
            except Exception:
                logging.exception("เขียนสถานะ ERROR กลับ Excel ไม่สำเร็จสำหรับ JOB %s",ref)
            try: tvc.screenshot(RUNTIME_PATHS.screenshots_dir/f"ERROR_{ref}_{stamp()}.png")
            except Exception: pass
            print("ERROR:",e)
            emit_event(
                "job_error",
                ref=ref,
                index=job_index,
                total=total_jobs,
                message=error_result,
                commit_uncertain=commit_started,
            )
            had_error=True
            break

    print("Log:",log_file)
    emit_event("batch_complete",total=total_jobs,success=not had_error,log=str(log_file))
    return 1 if had_error else 0


if __name__=="__main__":
    sys.exit(main())
