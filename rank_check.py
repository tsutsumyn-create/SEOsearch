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
import re
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


def col(df, *names):
    for n in names:
        if n in df.columns:
            return df[n]
    return pd.Series([""] * len(df), index=df.index)


def split_cats(c):
    return {x.strip() for x in re.split(r"[／/・|、]", c or "") if x.strip()}


def load_keywords():
    """keywords.csv → {キーワード: (分野, 意図)}（分野・意図の列は任意）"""
    df = read_csv_jp(BASE / "keywords.csv")
    meta = {}
    for k, c, i in zip(col(df, "keyword", "キーワード"), col(df, "分野", "category"), col(df, "意図", "intent")):
        k = k.strip()
        if k and k not in meta:
            meta[k] = (c.strip(), i.strip())
    return meta


def load_targets():
    """domains.csv → (マッチ用[(domain,label)], 表示順のlabel一覧, {label: 分野set})
    同じ label の行は1社として扱う（複数ドメインのうち最高順位を採用）。
    分野が空欄 or「総合」なら全キーワードが対象。"""
    df = read_csv_jp(BASE / "domains.csv")
    targets, labels, label_cats = [], [], {}
    for d, l, c in zip(col(df, "domain", "ドメイン"), col(df, "label", "表示名"), col(df, "分野", "category")):
        d = clean_domain(d)
        if not d:
            continue
        l = l.strip() or d
        targets.append((d, l))
        label_cats.setdefault(l, set()).update(split_cats(c) or {"総合"})
        if l not in labels:
            labels.append(l)
    # サブドメイン指定を優先してマッチさせるため、長いドメインから判定
    targets.sort(key=lambda t: -len(t[0]))
    return targets, labels, label_cats


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
NA = "－"   # 分野対象外


def to_int(v):
    return int(v) if v is not None and pd.notna(v) else None


def top_domains(df, n=3):
    return " / ".join(df["domain"].value_counts().head(n).index) if len(df) else ""


def build_report(run_at, meta, targets, labels, label_cats):
    keywords = list(meta)

    def applicable(kw, label):
        kcat = meta[kw][0]
        cats = label_cats.get(label, {"総合"})
        return (not kcat) or kcat == "総合" or "総合" in cats or kcat in cats

    def head(kw, dev):
        return {"キーワード": kw, "分野": meta[kw][0], "意図": meta[kw][1], "端末": DEV_JA.get(dev, dev)}

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
            row = head(kw, dev)
            ok = status_latest.get((kw, dev)) == "ok"
            for l in labels:
                if not applicable(kw, l):
                    row[l] = NA
                elif not ok:
                    row[l] = "取得失敗"
                else:
                    v = to_int(best.get((kw, dev, l)))
                    row[l] = v if v is not None else OUT_OF_RANGE
            row["参考：上位3ドメイン"] = top_domains(
                latest[(latest["keyword"] == kw) & (latest["device"] == dev) & (latest["rank"] <= 3)]
                .sort_values("rank"), 3)
            org_rows.append(row)
    org_df = pd.DataFrame(org_rows)

    # --- ドメイン別サマリ（最新・分野対象セルのみ）
    intents = list(dict.fromkeys(i for _, i in meta.values() if i))
    lim = PAGES * 10
    sum_rows = []
    for dev in DEVICES:
        tot_n = tot_in = 0
        tot_int = {i: [0, 0] for i in intents}
        for l in labels:
            tgt = [kw for kw in keywords if applicable(kw, l) and status_latest.get((kw, dev)) == "ok"]
            rk = {kw: to_int(best.get((kw, dev, l))) for kw in tgt}
            ranks = [r for r in rk.values() if r is not None]
            n = len(tgt)
            row = {"端末": DEV_JA.get(dev, dev), "ドメイン": l, "対象KW数": n, "圏内KW数": len(ranks),
                   "圏内率": len(ranks) / n if n else "-",
                   "1〜3位": sum(r <= 3 for r in ranks), "4〜10位": sum(3 < r <= 10 for r in ranks),
                   f"11〜{lim}位": sum(r > 10 for r in ranks),
                   "平均順位（圏内のみ）": round(sum(ranks) / len(ranks), 1) if ranks else "-"}
            for i in intents:
                ks = [kw for kw in tgt if meta[kw][1] == i]
                hit = sum(rk[kw] is not None for kw in ks)
                row[f"圏内率_{i}"] = hit / len(ks) if ks else "-"
                tot_int[i][0] += len(ks); tot_int[i][1] += hit
            tot_n += n; tot_in += len(ranks)
            sum_rows.append(row)
        total = {"端末": DEV_JA.get(dev, dev), "ドメイン": "【全体】", "対象KW数": tot_n, "圏内KW数": tot_in,
                 "圏内率": tot_in / tot_n if tot_n else "-",
                 "1〜3位": "", "4〜10位": "", f"11〜{lim}位": "", "平均順位（圏内のみ）": ""}
        for i in intents:
            total[f"圏内率_{i}"] = tot_int[i][1] / tot_int[i][0] if tot_int[i][0] else "-"
        sum_rows.append(total)
    sum_df = pd.DataFrame(sum_rows)

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
            r1 = {**head(kw, dev), "計測回数": n}
            r2 = dict(r1)
            for l in labels:
                if not applicable(kw, l):
                    r1[l] = r2[l] = NA
                    continue
                r1[l] = (appear.get((kw, dev, l), 0) / n) if n else "-"
                v = top_avg.get((kw, dev, l))
                r2[l] = round(v, 1) if v is not None else "-"
            sub = ads[(ads["keyword"] == kw) & (ads["device"] == dev)]
            r1["参考：広告出稿の多いドメイン"] = top_domains(sub.drop_duplicates(["run_at", "domain"]), 3)
            rate_rows.append(r1)
            pos_rows.append(r2)
    rate_df, pos_df = pd.DataFrame(rate_rows), pd.DataFrame(pos_rows)

    # --- オーガニック推移（分野対象の組み合わせのみ）
    run_list = sorted(runs_ok["run_at"].unique())[-MAX_TREND_COLS:]
    org_w = hist_w[(hist_w["type"] == "organic") & (hist_w["target"] != "")]
    best_t = org_w.groupby(["keyword", "device", "target", "run_at"])["rank"].min().to_dict()
    ok_set = set(zip(runs_ok["keyword"], runs_ok["device"], runs_ok["run_at"]))
    trend_rows = []
    for kw in keywords:
        for dev in devs_seen:
            for l in labels:
                if not applicable(kw, l):
                    continue
                row = {**head(kw, dev), "ドメイン": l}
                for r in run_list:
                    if (kw, dev, r) not in ok_set:
                        row[r] = ""
                    else:
                        v = to_int(best_t.get((kw, dev, l, r)))
                        row[r] = v if v is not None else OUT_OF_RANGE
                trend_rows.append(row)
    trend_df = pd.DataFrame(trend_rows)

    # --- 生データ・実行ログ
    raw = hist[hist["run_at"] == run_at].copy()
    raw["device"] = raw["device"].map(lambda d: DEV_JA.get(d, d))
    raw["type"] = raw["type"].map({"organic": "オーガニック", "ad": "広告"})
    raw["block"] = raw["block"].map({"top": "上部", "bottom": "下部"}).fillna("")
    raw["cat"] = raw["keyword"].map(lambda k: meta.get(k, ("", ""))[0])
    raw["intent"] = raw["keyword"].map(lambda k: meta.get(k, ("", ""))[1])
    raw = raw[["keyword", "cat", "intent", "device", "type", "block", "rank", "domain", "target", "title", "url"]]
    raw.columns = ["キーワード", "分野", "意図", "端末", "種別", "広告枠", "順位", "ドメイン", "追跡対象", "タイトル", "URL"]
    log = runs.iloc[::-1].head(1000).rename(columns={
        "run_at": "実行日時", "keyword": "キーワード", "device": "端末", "status": "結果"})

    info = pd.DataFrame([
        ["最新の実行日時（JST）", run_at],
        ["検索条件", f"google.co.jp / 日本語 / 地域: {LOCATION or '指定なし（日本）'}"],
        ["端末", "、".join(DEV_JA.get(d, d) for d in DEVICES)],
        ["取得範囲", f"上位{lim}位まで（それより下は「{OUT_OF_RANGE}」）"],
        ["分野対象外", f"キーワードの分野がサイトの分野と合わない組み合わせは「{NA}」（集計からも除外）"],
        ["出現率・推移の集計期間", f"直近{WINDOW_DAYS}日"],
        ["", ""],
        ["ドメイン別サマリ", "最新実行での各サイトの圏内率・順位分布（検索意図別の圏内率つき）。【全体】は表のマスのうち圏内の割合"],
        ["オーガニック_最新", "最新の実行時点での各サイトの最高順位"],
        ["オーガニック_推移", "実行ごとの順位の推移（空欄は未計測）"],
        ["広告_出現率", "集計期間の計測回数のうち、そのサイトの広告が表示された割合"],
        ["広告_上部平均順位", "検索結果上部に広告が出たときの平均掲載順位（1=最上位）"],
        ["生データ_最新", "最新実行で取得した全結果（追跡対象外のサイトも含む）"],
        ["", ""],
        ["注意", "広告の掲載はオークションで都度変わるため、1回の結果ではなく出現率で見ること。"
                 "検索結果はパーソナライズ等により実際の表示と差が出る場合がある。"],
    ], columns=["項目", "内容"])

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as w:
        info.to_excel(w, sheet_name="説明", index=False)
        sum_df.to_excel(w, sheet_name="ドメイン別サマリ", index=False)
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
    grey = Font(color="BFBFBF")
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
            c.alignment = Alignment(vertical="center", wrap_text=True)
        for i, colc in enumerate(ws.iter_cols(min_row=1, max_row=min(ws.max_row, 200)), start=1):
            width = max((len(str(c.value)) if c.value is not None else 0) for c in colc)
            ws.column_dimensions[get_column_letter(i)].width = min(max(8, width * 1.6), 60)
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if c.value == NA:
                    c.font = grey
                    c.alignment = Alignment(horizontal="center")
    wb["説明"].column_dimensions["B"].width = 100

    def rank_rules(ws, first_col, last_col):
        if ws.max_row < 2 or last_col < first_col:
            return
        rng = f"{get_column_letter(first_col)}2:{get_column_letter(last_col)}{ws.max_row}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["1", "3"], fill=g1))
        ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["4", "10"], fill=g2))

    ws = wb["オーガニック_最新"]; ws.freeze_panes = "E2"; rank_rules(ws, 5, 4 + n_labels)
    ws = wb["オーガニック_推移"]; ws.freeze_panes = "F2"; rank_rules(ws, 6, ws.max_column)
    ws = wb["広告_上部平均順位"]; ws.freeze_panes = "F2"; rank_rules(ws, 6, 5 + n_labels)
    ws = wb["広告_出現率"]; ws.freeze_panes = "F2"
    if ws.max_row >= 2 and n_labels:
        for row in ws.iter_rows(min_row=2, min_col=6, max_col=5 + n_labels):
            for c in row:
                if isinstance(c.value, (int, float)):
                    c.number_format = "0%"
        rng = f"F2:{get_column_letter(5 + n_labels)}{ws.max_row}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0.5"], fill=g1))
        ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["0.0001", "0.4999"], fill=g2))
    ws = wb["ドメイン別サマリ"]; ws.freeze_panes = "C2"
    for j, h in enumerate(ws[1], start=1):
        if h.value and "圏内率" in str(h.value):
            for row in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                for c in row:
                    if isinstance(c.value, (int, float)):
                        c.number_format = "0%"
    for row in ws.iter_rows(min_row=2):
        if row[1].value == "【全体】":
            for c in row:
                c.font = Font(bold=True)
    for name in ("生データ_最新", "実行ログ"):
        wb[name].freeze_panes = "A2"
    wb.save(REPORT_PATH)


# ---------------------------------------------------------------- メイン
def run(run_at=None):
    if not API_KEY:
        sys.exit("SERPAPI_KEY が設定されていません（GitHub の Secrets に登録してください）")
    bad = [d for d in DEVICES if d not in DEV_JA]
    if bad:
        sys.exit(f"DEVICES の指定が不正です: {bad}（mobile / desktop / tablet）")
    meta = load_keywords()
    targets, labels, label_cats = load_targets()
    if not meta:
        sys.exit("keywords.csv にキーワードがありません")
    run_at = run_at or datetime.now(JST).strftime("%Y-%m-%d %H:%M")
    print(f"実行: {run_at} / キーワード {len(meta)}件 × 端末 {len(DEVICES)} × {PAGES}ページ"
          f" = 最大 {len(meta) * len(DEVICES) * PAGES} クレジット")

    rows, runs = fetch(list(meta), run_at)
    if all(r["status"] != "ok" for r in runs):
        sys.exit(f"全件取得に失敗しました。最初のエラー: {runs[0]['status']}")
    append_csv(HIST_PATH, rows, HIST_COLS)
    append_csv(RUNS_PATH, runs, RUN_COLS)
    build_report(run_at, meta, targets, labels, label_cats)


if __name__ == "__main__":
    run()
