#!/usr/bin/env python3
"""
SERP順位チェッカー（SerpApi使用）

keywords.csv のキーワードを Google（日本）で検索し、domains.csv に登録したドメインの
・オーガニック順位
・リスティング広告の出現率／掲載順位
を記録して、Excelレポート（reports/rank_report_latest.xlsx）を出力する。
結果は data/history.csv に蓄積されるので、定期実行すると推移・出現率が見られる。
"""
import os
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 設定
BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
REPORT_DIR = BASE / "reports"
HIST_PATH = DATA_DIR / "history.csv"
RUNS_PATH = DATA_DIR / "runs.csv"
REPORT_PATH = REPORT_DIR / "rank_report_latest.xlsx"
JST = timezone(timedelta(hours=9))

API_KEY = os.environ.get("SERPAPI_KEY", "").strip()
DEVICES = [d.strip().lower() for d in os.environ.get("DEVICES", "mobile").split(",") if d.strip()]
PAGES = max(1, int(os.environ.get("PAGES", "2") or 2))          # 1ページ = 上位10件 = 1クレジット
LOCATION = os.environ.get("LOCATION", "").strip()               # 例: Tokyo,Tokyo,Japan（空なら日本全体）
WINDOW_DAYS = int(os.environ.get("WINDOW_DAYS", "30") or 30)     # 出現率・推移の集計期間（日）
MAX_TREND_COLS = 60

HIST_COLS = ["run_at", "keyword", "device", "type", "block", "rank", "domain", "url", "title"]
RUN_COLS = ["run_at", "keyword", "device", "status"]
DEV_JA = {"mobile": "スマホ", "desktop": "PC", "tablet": "タブレット"}
OUT_OF_RANGE = "圏外"


# ---------------------------------------------------------------- 入力
def read_csv_jp(path):
    """UTF-8 / Shift-JIS どちらで保存されたCSVでも読めるようにする"""
    for enc in ("utf-8-sig", "cp932"):
        try:
            return pd.read_csv(path, dtype=str, encoding=enc).fillna("")
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f"{path} の文字コードを判定できません（UTF-8で保存してください）")


def clean_domain(s):
    s = (s or "").strip().lower()
    if "://" in s:
        s = urlparse(s).netloc
    parts = s.split()
    s = parts[0] if parts else ""
    s = s.split("/")[0].split(":")[0]
    return s[4:] if s.startswith("www.") else s


def host_of(url):
    if not url:
        return ""
    if "://" not in url:
        url = "https://" + url
    return clean_domain(urlparse(url).netloc)


def load_keywords():
    df = read_csv_jp(BASE / "keywords.csv")
    kws = [k.strip() for k in df["keyword"] if k.strip()]
    return list(dict.fromkeys(kws))


def load_targets():
    df = read_csv_jp(BASE / "domains.csv")
    targets = []
    for _, r in df.iterrows():
        d = clean_domain(r.get("domain", ""))
        if d:
            targets.append([d, (r.get("label", "") or "").strip() or d])
    labels = [l for _, l in targets]
    for t in targets:                      # ラベル重複時はドメインを付けて区別
        if labels.count(t[1]) > 1:
            t[1] = f"{t[1]}（{t[0]}）"
    # サブドメイン指定を優先してマッチさせるため、長いドメインから判定
    return sorted([tuple(t) for t in targets], key=lambda t: -len(t[0]))


def match_label(host, targets):
    for d, label in targets:
        if host == d or host.endswith("." + d):
            return label
    return ""


# ---------------------------------------------------------------- 取得
def serp(q, device, start):
    params = {
        "engine": "google", "q": q, "google_domain": "google.co.jp",
        "gl": "jp", "hl": "ja", "device": device, "start": start,
        "no_cache": "true", "api_key": API_KEY,
    }
    if LOCATION:
        params["location"] = LOCATION
    last = None
    for attempt in range(3):
        try:
            r = requests.get("https://serpapi.com/search.json", params=params, timeout=90)
            data = r.json()
        except Exception as e:  # 通信エラー等はリトライ
            last = e
            time.sleep(5 * (attempt + 1))
            continue
        err = data.get("error")
        if err:
            if "hasn't returned any results" in err:
                return {}
            if r.status_code in (429, 500, 502, 503):
                last = RuntimeError(err)
                time.sleep(10 * (attempt + 1))
                continue
            raise RuntimeError(err)
        return data
    raise RuntimeError(f"取得失敗: {last}")


def fetch(keywords, run_at):
    rows, runs = [], []
    for kw in keywords:
        for device in DEVICES:
            status, rank = "ok", 0
            try:
                for p in range(PAGES):
                    data = serp(kw, device, p * 10)
                    if p == 0:
                        for ad in data.get("ads", []) or []:
                            url = ad.get("link") or ad.get("displayed_link") or ""
                            rows.append({
                                "run_at": run_at, "keyword": kw, "device": device, "type": "ad",
                                "block": ad.get("block_position") or "", "rank": ad.get("position") or "",
                                "domain": host_of(url), "url": url, "title": ad.get("title") or "",
                            })
                    organic = data.get("organic_results", []) or []
                    if not organic:
                        break
                    for res in organic:
                        rank += 1
                        url = res.get("link") or ""
                        rows.append({
                            "run_at": run_at, "keyword": kw, "device": device, "type": "organic",
                            "block": "", "rank": rank, "domain": host_of(url), "url": url,
                            "title": res.get("title") or "",
                        })
                    time.sleep(1)
            except Exception as e:
                status = f"error: {e}"[:300]
                print(f"[ERROR] {kw} / {device}: {e}", file=sys.stderr)
            runs.append({"run_at": run_at, "keyword": kw, "device": device, "status": status})
            print(f"{kw} / {DEV_JA.get(device, device)}: {status}（オーガニック {rank}件）")
    return rows, runs


def append_csv(path, rows, cols):
    new = pd.DataFrame(rows, columns=cols)
    if path.exists():
        new = pd.concat([read_csv_jp(path), new], ignore_index=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    new.to_csv(path, index=False, encoding="utf-8-sig")


# ---------------------------------------------------------------- レポート
def to_int(v):
    return int(v) if pd.notna(v) else None


def top_domains(df, n=3):
    return " / ".join(df["domain"].value_counts().head(n).index) if len(df) else ""


def build_report(run_at, keywords, targets):
    labels = [l for _, l in sorted(targets, key=lambda t: load_order[t[0]])]
    hist = read_csv_jp(HIST_PATH)
    runs = read_csv_jp(RUNS_PATH)
    hist["rank"] = pd.to_numeric(hist["rank"], errors="coerce")
    hist["target"] = hist["domain"].map(lambda h: match_label(h, targets))

    cutoff = (datetime.now(JST) - timedelta(days=WINDOW_DAYS)).strftime("%Y-%m-%d %H:%M")
    hist_w = hist[hist["run_at"] >= cutoff]
    runs_ok = runs[(runs["status"] == "ok") & (runs["run_at"] >= cutoff)]
    status_latest = {(r.keyword, r.device): r.status for r in runs[runs["run_at"] == run_at].itertuples()}

    # --- オーガニック（最新）
    latest = hist[(hist["run_at"] == run_at) & (hist["type"] == "organic")]
    best = latest[latest["target"] != ""].groupby(["keyword", "device", "target"])["rank"].min().to_dict()
    org_rows = []
    for kw in keywords:
        for dev in DEVICES:
            row = {"キーワード": kw, "端末": DEV_JA.get(dev, dev)}
            ok = status_latest.get((kw, dev)) == "ok"
            for l in labels:
                v = best.get((kw, dev, l))
                row[l] = (to_int(v) if v is not None else OUT_OF_RANGE) if ok else "取得失敗"
            row["参考：上位3ドメイン"] = top_domains(
                latest[(latest["keyword"] == kw) & (latest["device"] == dev) & (latest["rank"] <= 3)]
                .sort_values("rank"), 3)
            org_rows.append(row)
    org_df = pd.DataFrame(org_rows)

    # --- 広告（集計期間内の出現率・上部平均順位）
    devs_seen = [d for d in ("mobile", "desktop", "tablet") if d in set(runs_ok["device"])] or DEVICES
    n_runs = runs_ok.groupby(["keyword", "device"])["run_at"].nunique().to_dict()
    ads = hist_w[hist_w["type"] == "ad"]
    ads_t = ads[ads["target"] != ""]
    appear = ads_t.groupby(["keyword", "device", "target"])["run_at"].nunique().to_dict()
    top_ads = ads_t[ads_t["block"] == "top"]
    top_avg = {}
    if len(top_ads):
        per_run = top_ads.groupby(["keyword", "device", "target", "run_at"])["rank"].min()
        top_avg = per_run.groupby(level=[0, 1, 2]).mean().to_dict()

    rate_rows, pos_rows = [], []
    for kw in keywords:
        for dev in devs_seen:
            n = n_runs.get((kw, dev), 0)
            base = {"キーワード": kw, "端末": DEV_JA.get(dev, dev), "計測回数": n}
            r1, r2 = dict(base), dict(base)
            for l in labels:
                r1[l] = (appear.get((kw, dev, l), 0) / n) if n else "-"
                v = top_avg.get((kw, dev, l))
                r2[l] = round(v, 1) if v is not None else "-"
            sub = ads[(ads["keyword"] == kw) & (ads["device"] == dev)]
            r1["参考：広告出稿の多いドメイン"] = top_domains(sub.drop_duplicates(["run_at", "domain"]), 3)
            rate_rows.append(r1)
            pos_rows.append(r2)
    rate_df, pos_df = pd.DataFrame(rate_rows), pd.DataFrame(pos_rows)

    # --- オーガニック推移
    run_list = sorted(runs_ok["run_at"].unique())[-MAX_TREND_COLS:]
    org_w = hist_w[(hist_w["type"] == "organic") & (hist_w["target"] != "")]
    best_t = org_w.groupby(["keyword", "device", "target", "run_at"])["rank"].min().to_dict()
    ok_set = set(zip(runs_ok["keyword"], runs_ok["device"], runs_ok["run_at"]))
    trend_rows = []
    for kw in keywords:
        for dev in devs_seen:
            for l in labels:
                row = {"キーワード": kw, "端末": DEV_JA.get(dev, dev), "ドメイン": l}
                for r in run_list:
                    if (kw, dev, r) not in ok_set:
                        row[r] = ""
                    else:
                        v = best_t.get((kw, dev, l, r))
                        row[r] = to_int(v) if v is not None else OUT_OF_RANGE
                trend_rows.append(row)
    trend_df = pd.DataFrame(trend_rows)

    # --- 生データ・実行ログ
    raw = hist[hist["run_at"] == run_at].copy()
    raw["device"] = raw["device"].map(lambda d: DEV_JA.get(d, d))
    raw["type"] = raw["type"].map({"organic": "オーガニック", "ad": "広告"})
    raw["block"] = raw["block"].map({"top": "上部", "bottom": "下部"}).fillna("")
    raw = raw[["keyword", "device", "type", "block", "rank", "domain", "target", "title", "url"]]
    raw.columns = ["キーワード", "端末", "種別", "広告枠", "順位", "ドメイン", "追跡対象", "タイトル", "URL"]
    log = runs.iloc[::-1].head(1000).rename(columns={
        "run_at": "実行日時", "keyword": "キーワード", "device": "端末", "status": "結果"})

    info = pd.DataFrame([
        ["最新の実行日時（JST）", run_at],
        ["検索条件", f"google.co.jp / 日本語 / 地域: {LOCATION or '指定なし（日本）'}"],
        ["端末", "、".join(DEV_JA.get(d, d) for d in DEVICES)],
        ["取得範囲", f"上位{PAGES * 10}位まで（それより下は「{OUT_OF_RANGE}」）"],
        ["出現率・推移の集計期間", f"直近{WINDOW_DAYS}日"],
        ["", ""],
        ["オーガニック_最新", "最新の実行時点での各ドメインの最高順位"],
        ["オーガニック_推移", "実行ごとの順位の推移（空欄は未計測）"],
        ["広告_出現率", "集計期間の計測回数のうち、そのドメインの広告が表示された割合"],
        ["広告_上部平均順位", "検索結果上部に広告が出たときの平均掲載順位（1=最上位）"],
        ["生データ_最新", "最新実行で取得した全結果（追跡対象外のサイトも含む）"],
        ["", ""],
        ["注意", "広告の掲載はオークションで都度変わるため、1回の結果ではなく出現率で見ること。"
                 "検索結果はパーソナライズ等により実際の表示と差が出る場合がある。"],
    ], columns=["項目", "内容"])

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as w:
        info.to_excel(w, sheet_name="説明", index=False)
        org_df.to_excel(w, sheet_name="オーガニック_最新", index=False)
        trend_df.to_excel(w, sheet_name="オーガニック_推移", index=False)
        rate_df.to_excel(w, sheet_name="広告_出現率", index=False)
        pos_df.to_excel(w, sheet_name="広告_上部平均順位", index=False)
        raw.to_excel(w, sheet_name="生データ_最新", index=False)
        log.to_excel(w, sheet_name="実行ログ", index=False)
    style_workbook(len(labels))
    print(f"レポート出力: {REPORT_PATH}")


def style_workbook(n_labels):
    wb = load_workbook(REPORT_PATH)
    head_fill = PatternFill("solid", fgColor="1F3864")
    g1 = PatternFill("solid", fgColor="A9D08E")
    g2 = PatternFill("solid", fgColor="E2EFDA")
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
            c.alignment = Alignment(vertical="center", wrap_text=True)
        for i, col in enumerate(ws.iter_cols(min_row=1, max_row=min(ws.max_row, 200)), start=1):
            width = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            ws.column_dimensions[get_column_letter(i)].width = min(max(8, width * 1.6), 60)
    wb["説明"].column_dimensions["B"].width = 90

    def rank_rules(ws, first_col, last_col):
        if ws.max_row < 2 or last_col < first_col:
            return
        rng = f"{get_column_letter(first_col)}2:{get_column_letter(last_col)}{ws.max_row}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["1", "3"], fill=g1))
        ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["4", "10"], fill=g2))

    ws = wb["オーガニック_最新"]; ws.freeze_panes = "C2"; rank_rules(ws, 3, 2 + n_labels)
    ws = wb["オーガニック_推移"]; ws.freeze_panes = "D2"; rank_rules(ws, 4, ws.max_column)
    ws = wb["広告_上部平均順位"]; ws.freeze_panes = "D2"; rank_rules(ws, 4, 3 + n_labels)
    ws = wb["広告_出現率"]; ws.freeze_panes = "D2"
    if ws.max_row >= 2 and n_labels:
        for row in ws.iter_rows(min_row=2, min_col=4, max_col=3 + n_labels):
            for c in row:
                if isinstance(c.value, (int, float)):
                    c.number_format = "0%"
        rng = f"D2:{get_column_letter(3 + n_labels)}{ws.max_row}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0.5"], fill=g1))
        ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["0.0001", "0.4999"], fill=g2))
    for name in ("生データ_最新", "実行ログ"):
        wb[name].freeze_panes = "A2"
    wb.save(REPORT_PATH)


# ---------------------------------------------------------------- メイン
load_order = {}


def run(run_at=None):
    global load_order
    if not API_KEY:
        sys.exit("SERPAPI_KEY が設定されていません（GitHub の Secrets に登録してください）")
    bad = [d for d in DEVICES if d not in DEV_JA]
    if bad:
        sys.exit(f"DEVICES の指定が不正です: {bad}（mobile / desktop / tablet）")
    keywords, targets = load_keywords(), load_targets()
    if not keywords:
        sys.exit("keywords.csv にキーワードがありません")
    df = read_csv_jp(BASE / "domains.csv")
    load_order = {clean_domain(d): i for i, d in enumerate(df["domain"])}
    run_at = run_at or datetime.now(JST).strftime("%Y-%m-%d %H:%M")
    print(f"実行: {run_at} / キーワード {len(keywords)}件 × 端末 {len(DEVICES)} × {PAGES}ページ"
          f" = 最大 {len(keywords) * len(DEVICES) * PAGES} クレジット")

    rows, runs = fetch(keywords, run_at)
    if all(r["status"] != "ok" for r in runs):
        sys.exit(f"全件取得に失敗しました。最初のエラー: {runs[0]['status']}")
    append_csv(HIST_PATH, rows, HIST_COLS)
    append_csv(RUNS_PATH, runs, RUN_COLS)
    build_report(run_at, keywords, targets)


if __name__ == "__main__":
    run()
