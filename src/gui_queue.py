from __future__ import annotations

from dataclasses import dataclass, field
import os
from pathlib import Path


EMPTY_STATS={
    "WAIT":0,
    "DONE":0,
    "ERROR":0,
    "RUNNING":0,
    "OTHER":0,
    "TOTAL":0,
    "COMPLETED":0,
}


def normalize_excel_path(path):
    return Path(path).expanduser().resolve()


def normalized_path_key(path):
    return os.path.normcase(str(normalize_excel_path(path)))


@dataclass
class ExcelQueueItem:
    path: Path
    status: str="PENDING"
    stats: dict=field(default_factory=lambda:dict(EMPTY_STATS))
    message: str=""
    error_jobs: list=field(default_factory=list)

    def __post_init__(self):
        self.path=normalize_excel_path(self.path)

    @property
    def key(self):
        return normalized_path_key(self.path)


class ExcelQueue:
    def __init__(self):
        self.items=[]
        self.locked=False

    def add_paths(self,paths):
        if self.locked:
            raise RuntimeError("Queue ถูกล็อกระหว่างทำงาน")
        existing={item.key for item in self.items}
        added=[]
        duplicates=[]
        for raw_path in paths:
            path=normalize_excel_path(raw_path)
            key=normalized_path_key(path)
            if key in existing:
                duplicates.append(path)
                continue
            item=ExcelQueueItem(path)
            self.items.append(item)
            existing.add(key)
            added.append(item)
        return added,duplicates

    def remove_indices(self,indices):
        if self.locked:
            raise RuntimeError("Queue ถูกล็อกระหว่างทำงาน")
        removed=[]
        for index in sorted(set(indices),reverse=True):
            if 0<=index<len(self.items):
                removed.append(self.items.pop(index))
        removed.reverse()
        return removed

    def clear(self):
        if self.locked:
            raise RuntimeError("Queue ถูกล็อกระหว่างทำงาน")
        removed=list(self.items)
        self.items.clear()
        return removed

    def move(self,index,offset):
        if self.locked:
            raise RuntimeError("Queue ถูกล็อกระหว่างทำงาน")
        target=index+offset
        if not (0<=index<len(self.items) and 0<=target<len(self.items)):
            return index
        self.items[index],self.items[target]=self.items[target],self.items[index]
        return target

    def totals(self):
        totals=dict(EMPTY_STATS)
        for item in self.items:
            for key in EMPTY_STATS:
                totals[key]+=int(item.stats.get(key,0) or 0)
        return totals

    def all_ready(self):
        return bool(self.items) and all(item.status=="READY" for item in self.items)


class QueueRunController:
    """Small deterministic state machine used by the GUI subprocess orchestrator."""

    def __init__(self,total_files):
        self.total_files=max(int(total_files),0)
        self.current_index=-1
        self.active=False
        self.outcome="IDLE"

    def start(self):
        if self.total_files<=0:
            self.outcome="EMPTY"
            return None
        self.active=True
        self.current_index=0
        self.outcome="RUNNING"
        return self.current_index

    def complete_current(self,return_code,stop_requested=False):
        """Return the next index only for a clean exit without a stop request."""
        if not self.active:
            return None
        if stop_requested or return_code==2:
            self.active=False
            self.outcome="STOPPED"
            return None
        if return_code!=0:
            self.active=False
            self.outcome="ERROR"
            return None
        next_index=self.current_index+1
        if next_index>=self.total_files:
            self.active=False
            self.outcome="COMPLETE"
            return None
        self.current_index=next_index
        return next_index


def calculate_queue_progress(items,current_index=-1):
    totals=dict(EMPTY_STATS)
    for item in items:
        for key in EMPTY_STATS:
            totals[key]+=int(item.stats.get(key,0) or 0)
    if 0<=current_index<len(items):
        current=dict(EMPTY_STATS)
        for key in EMPTY_STATS:
            current[key]=int(items[current_index].stats.get(key,0) or 0)
    else:
        current=dict(EMPTY_STATS)
    return {
        "current":current,
        "overall":totals,
        "current_processed":current["DONE"]+current["ERROR"],
        "overall_processed":totals["DONE"]+totals["ERROR"],
    }


def build_queue_summary(items):
    totals=dict(EMPTY_STATS)
    error_files=[]
    error_file_keys=set()
    problems=[]
    completed_files=0
    for item in items:
        for key in EMPTY_STATS:
            totals[key]+=int(item.stats.get(key,0) or 0)
        if item.status=="DONE":
            completed_files+=1
        if item.status in {"ERROR","INVALID"} or int(item.stats.get("ERROR",0) or 0)>0:
            error_files.append(str(item.path))
            error_file_keys.add(item.key)
        if item.message and item.status in {"ERROR","INVALID","STOPPED"}:
            problems.append(f"{item.path.name} | {item.message}")
        for error in item.error_jobs:
            ref=str(error.get("job_ref") or "-")
            result=str(error.get("bot_result") or "ERROR")
            problems.append(f"{item.path.name} | {ref} | {result}")
    return {
        "files_total":len(items),
        "files_completed":completed_files,
        "files_error":len(error_file_keys),
        "error_files":error_files,
        "problems":problems,
        **totals,
    }
