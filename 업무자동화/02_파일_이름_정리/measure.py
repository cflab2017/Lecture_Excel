# -*- coding: utf-8 -*-
"""업무 자동화 2편 — 파일 이름 한꺼번에 정리, 실측 스크립트 (코딩나우 · 2026-11-25)

    py measure.py            → measure.json (같은 폴더)

필요: Windows · 엑셀(숨긴 COM 으로만 씀, 창 안 띄움) · Windows PowerShell 5.1 · pywin32 · openpyxl
재는 것
  1) 파일 수정 시각으로 날짜를 정하면 몇 개가 맞나(zip 안 시각 = 메일에서 받은 날 흉내)
  2) 파일 이름만으로 날짜를 꺼내면 — 연 · 월 · 일이 다 있음 / 월 · 일만 / 없음
  3) 같은 날 판이 여럿인 묶음 — 이름 꼬리(최종 · 수정 · (2) …) · 문서 안 「■ 고친 것」 줄 수 · 수정 시각 순서
  4) 손으로 바꾸기: PowerShell 로 목록.csv 만들기 → 엑셀에서 열어 NewName 열 채우기 → 「CSV UTF-8」로 저장
     → PowerShell Import-Csv + Rename-Item. 함정 비교로 그냥 「CSV」로 저장한 것도 같은 명령으로 돌린다.
  5) 검산: 파일 수 · 내용 SHA-256 · 바뀐 이름이 정답(정답_날짜_제목.csv + 클로드 대조표)과 같은가
"""
import csv
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
ZIP = HERE / "예제_회의록_30개.zip"
TRUTH = HERE / "정답_날짜_제목.csv"
TABLE = HERE / "이름_바꾼_목록.xlsx"   # 클로드가 만든 대조표(원래 이름 → 바뀐 이름)
WORK = Path(tempfile.gettempdir()) / "cn_auto2_measure"
# 가이드에 그대로 적는 명령(바꾸기 전 · 뒤 내용 해시 비교, 되돌리기)
HASH_CMD = "Get-ChildItem files -File | Get-FileHash | ForEach-Object Hash | Sort-Object > {out}"
UNDO_CMD = ("Import-Csv {csv} | Where-Object {{ $_.NewName }} | ForEach-Object {{ "
            "Rename-Item -LiteralPath (Join-Path files $_.NewName) -NewName $_.Name }}")


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def unzip(dst):
    if dst.exists():
        shutil.rmtree(dst)
    dst.mkdir(parents=True)
    with zipfile.ZipFile(ZIP) as z:
        for i in z.infolist():
            if i.is_dir():
                continue
            p = dst / i.filename
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_bytes(z.read(i))
            t = time.mktime(i.date_time + (0, 0, -1))
            os.utime(p, (t, t))
    return dst / "files"


def ps(cmd, cwd):
    """Windows PowerShell 5.1 — 출력은 UTF-8 로 받는다."""
    full = "[Console]::OutputEncoding=[Text.Encoding]::UTF8; $ErrorActionPreference='Continue'; " + cmd
    r = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", full],
                       cwd=str(cwd), capture_output=True)
    return r.returncode, r.stdout.decode("utf-8", "replace"), r.stderr.decode("utf-8", "replace")


# ---- 2) 이름에서 날짜 꺼내기 ----
PATS = [
    ("연.월.일", re.compile(r"(20\d{2})\.(\d{1,2})\.(\d{1,2})")),
    ("연-월-일", re.compile(r"(20\d{2})-(\d{1,2})-(\d{1,2})")),
    ("연월일 8자리", re.compile(r"(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)")),
    ("○월○일", re.compile(r"(?<!\d)(\d{1,2})월\s*(\d{1,2})일")),
    ("월-일", re.compile(r"(?<!\d)(\d{1,2})-(\d{1,2})(?!\d)")),
    ("월.일", re.compile(r"(?<!\d)(\d{1,2})\.(\d{1,2})(?!\d)")),
    ("월일 4자리", re.compile(r"(?<!\d)(\d{2})(\d{2})(?!\d)")),
]


def name_date(stem):
    for kind, rx in PATS:
        m = rx.search(stem)
        if m:
            g = [int(x) for x in m.groups()]
            if len(g) == 3:
                return kind, date(*g), True
            return kind, (g[0], g[1]), False
    return "없음", None, False


def docx_lines(p):
    x = zipfile.ZipFile(p).read("word/document.xml").decode("utf-8")
    paras = [re.sub(r"<[^>]+>", "", s) for s in re.findall(r"<w:p[ >].*?</w:p>", x, flags=re.S)]
    return [s for s in paras if s.strip()]


def edits(p):
    """문서 안 「■ 고친 것」「■ 덧붙임」 아래 · 줄 수 (docx 만)"""
    if p.suffix.lower() != ".docx":
        return None
    ls, n, on = docx_lines(p), 0, False
    for s in ls:
        if s.startswith("■"):
            on = ("고친 것" in s) or ("덧붙임" in s)
            continue
        if on and s.lstrip().startswith("·"):
            n += 1
    return n


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    out = {"measured": datetime.now().strftime("%Y-%m-%d %H:%M"), "zip_sha256": sha(ZIP)}
    truth = {r["원래 이름"]: r for r in csv.DictReader(io.open(TRUTH, encoding="utf-8-sig"))}
    from openpyxl import load_workbook
    ws = load_workbook(TABLE, data_only=True).worksheets[0]
    claude = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if isinstance(row[0], int) and row[1]:
            claude[row[1]] = None if str(row[2]).startswith("(") else row[2]
    out["claude_table_rows"] = len(claude)
    # 클로드 대조표의 바뀐 이름을 정답 표(날짜 · 제목)와 견줌 — 번호(_2, _3) · 확장자 소문자 · 빈 파일 그대로
    rx = re.compile(r"^(\d{4}-\d{2}-\d{2})_(.+?)(?:_(\d+))?(\.[A-Za-z0-9]+)$")
    bad = []
    for orig, new in claude.items():
        t = truth.get(orig)
        if t is None:
            bad.append([orig, new, "정답 표에 없음"])
        elif not t["제목"]:
            if new is not None:
                bad.append([orig, new, "빈 파일을 바꿈"])
        else:
            m = rx.match(new or "")
            if not m or m.group(1) != t["정답 날짜"] or m.group(2) != t["제목"] or m.group(4) != Path(orig).suffix.lower():
                bad.append([orig, new, f"정답 {t['정답 날짜']}_{t['제목']}"])
    out["claude_vs_truth_bad"] = bad

    files = unzip(WORK / "orig")
    names = sorted(p.name for p in files.iterdir())
    out["files"] = len(names)

    # 1) 수정 시각
    mt = {}
    for n in names:
        d = datetime.fromtimestamp((files / n).stat().st_mtime)
        mt[n] = d
    good = [n for n in names if truth[n]["정답 날짜"][:4] == "2026" and mt[n].date().isoformat() == truth[n]["정답 날짜"]]
    out["mtime"] = {
        "range": [min(mt.values()).strftime("%Y-%m-%d %H:%M"), max(mt.values()).strftime("%Y-%m-%d %H:%M")],
        "same_as_truth": len(good), "same_names": good,
        "dated_files": sum(1 for n in names if truth[n]["정답 날짜"][:4] == "2026"),
    }

    # 2) 이름에서 날짜
    kinds = defaultdict(list)
    wrong = []
    for n in names:
        kind, d, full = name_date(Path(n).stem)
        kinds[kind].append(n)
        t = truth[n]["정답 날짜"]
        if d is None:
            continue
        got = d.isoformat() if full else f"2026-{d[0]:02d}-{d[1]:02d}"
        if got != t:
            wrong.append([n, got, t])
    out["name_dates"] = {k: v for k, v in kinds.items()}
    out["name_date_counts"] = {k: len(v) for k, v in kinds.items()}
    out["name_date_wrong_if_2026"] = wrong
    out["sha_twins"] = [[a, b] for a in names for b in names if a < b and sha(files / a) == sha(files / b)]
    out["empty_files"] = [n for n in names if (files / n).stat().st_size == 0]
    out["upper_ext"] = [n for n in names if Path(n).suffix != Path(n).suffix.lower()]

    # 3) 같은 날 판
    grp = defaultdict(list)
    for n in names:
        t = truth[n]
        if t["제목"]:
            grp[(t["정답 날짜"], t["제목"])].append(n)
    same = []
    for (d, title), ns in sorted(grp.items()):
        if len(ns) < 2:
            continue
        rows = []
        for n in ns:
            rows.append({"name": n, "edits": edits(files / n), "mtime": mt[n].strftime("%m-%d %H:%M"),
                         "claude": claude.get(n)})
        same.append({"date": d, "title": title, "files": rows})
    out["same_day"] = same

    # 4) 손으로 바꾸기 — PowerShell → 엑셀 → PowerShell
    import win32com.client as wc
    runs = {}
    base = WORK / "hand"
    if base.exists():
        shutil.rmtree(base)
    base.mkdir(parents=True)
    unzip(base)  # base/files
    (base / "orig").exists() and shutil.rmtree(base / "orig")
    code, so, se = ps("Get-ChildItem files -File | Select-Object Name | Export-Csv 목록.csv -NoTypeInformation -Encoding UTF8", base)
    out["powershell"] = ps("$PSVersionTable.PSVersion.ToString()", base)[1].strip()
    raw = (base / "목록.csv").read_bytes()
    out["export_csv"] = {"exit": code, "bom": raw[:3] == b"\xef\xbb\xbf", "lines": raw.count(b"\n"), "stderr": se.strip()[:300]}

    xl = wc.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        out["excel"] = f"{xl.Version} build {xl.Build}"
        wb = xl.Workbooks.Open(str(base / "목록.csv"))
        sh = wb.Worksheets(1)
        n_rows = sh.UsedRange.Rows.Count
        col_a = [sh.Cells(r, 1).Value for r in range(1, n_rows + 1)]
        out["excel_open"] = {"rows": n_rows, "header": col_a[0],
                             "names_match": sorted(col_a[1:]) == names,
                             "mismatch": [v for v in col_a[1:] if v not in names][:5]}
        sh.Cells(1, 2).Value = "NewName"
        for r in range(2, n_rows + 1):
            new = claude.get(sh.Cells(r, 1).Value)
            if new:
                sh.Cells(r, 2).Value = new
        wb.SaveAs(str(base / "목록_utf8.csv"), 62)       # xlCSVUTF8 — 「CSV UTF-8(쉼표로 분리)」
        wb.SaveAs(str(base / "목록_ansi.csv"), 6)        # xlCSV — 「CSV(쉼표로 분리)」
        wb.Close(False)
    finally:
        xl.Quit()
        del xl
    for k in ("목록_utf8.csv", "목록_ansi.csv"):
        b = (base / k).read_bytes()
        out.setdefault("saved_csv", {})[k] = {"bom": b[:3] == b"\xef\xbb\xbf", "bytes": len(b),
                                               "utf8_ok": _utf8(b)}

    cmd_tpl = ("Import-Csv {csv} | Where-Object {{ $_.NewName }} | ForEach-Object {{ "
               "Rename-Item -LiteralPath (Join-Path files $_.Name) -NewName $_.NewName{whatif} }}")
    for label, csvname in (("utf8", "목록_utf8.csv"), ("ansi", "목록_ansi.csv")):
        d = WORK / f"run_{label}"
        if d.exists():
            shutil.rmtree(d)
        shutil.copytree(base / "files", d / "files")
        shutil.copy2(base / csvname, d / csvname)
        before = sorted(sha(p) for p in (d / "files").iterdir())
        ps(HASH_CMD.format(out="전.txt"), d)
        c1, w_out, w_err = ps(cmd_tpl.format(csv=csvname, whatif=" -WhatIf"), d)
        c2, r_out, r_err = ps(cmd_tpl.format(csv=csvname, whatif=""), d)
        ps(HASH_CMD.format(out="후.txt"), d)
        _, cmp_out, _ = ps("Compare-Object (Get-Content 전.txt) (Get-Content 후.txt)", d)
        after_names = sorted(p.name for p in (d / "files").iterdir())
        after = sorted(sha(p) for p in (d / "files").iterdir())
        want = sorted((claude.get(n) or n) for n in names)
        undo = {}
        if label == "utf8":
            _, _, u_err = ps(UNDO_CMD.format(csv=csvname), d)
            undo = {"errors": len(re.findall(r"Rename-Item\s*:", u_err)),
                    "names_back": sorted(p.name for p in (d / "files").iterdir()) == names,
                    "content_same": sorted(sha(p) for p in (d / "files").iterdir()) == before}
        runs[label] = {
            "hash_compare_lines": len([l for l in cmp_out.splitlines() if l.strip()]),
            "hash_file_lines": len((d / "전.txt").read_text(encoding="utf-16").split()),
            "undo": undo,
            "whatif_lines": len([l for l in w_out.splitlines() if l.strip()]),
            "whatif_first": [l for l in w_out.splitlines() if l.strip()][:2],
            "errors": len(re.findall(r"Rename-Item\s*:", r_err)),
            "error_first": r_err.strip().splitlines()[:3],
            "renamed": sum(1 for n in after_names if n not in names),
            "files_after": len(after_names),
            "content_same": before == after,
            "names_equal_claude_table": after_names == want,
        }
    out["hand_runs"] = runs
    (HERE / "measure.json").write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    print(json.dumps(out, ensure_ascii=False, indent=1, default=str))


def _utf8(b):
    try:
        b.decode("utf-8")
        return True
    except UnicodeDecodeError:
        return False


if __name__ == "__main__":
    main()
