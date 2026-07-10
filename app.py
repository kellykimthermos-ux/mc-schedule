# -*- coding: utf-8 -*-
"""
마컴 선택적 근무시간제 스케줄 관리 앱 (최종)

확정 스펙
- 로그인: 이름 선택만 (PIN 없음)
- 팀원 주간 등록: 제출 → 팀장 승인 후 확정 (주 단위 일괄 승인 카드 1장)
- 주간 반려: 해당 주 등록 취소(삭제) → 팀원 수정 후 재제출
- 지난주 복사: 같은 요일 기준 유형/시간/비고 복사 (이미 등록된 날짜 제외)
- 변경 요청: 확정 스케줄의 일단위 수정, 건별 승인/반려
- 팀장(김지윤) 본인 등록: 즉시 확정 (자기 승인 생략)
- 이메일: 요청 발생 → 팀장 / 처리 결과 → 팀원 회신 (양방향)
- 데이터: Google Sheets — schedule, requests 탭 자동 생성
"""

import smtplib
import uuid
from datetime import datetime, date, time, timedelta
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

KST = ZoneInfo("Asia/Seoul")
WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]
WORK_TYPES = ["근무", "연차", "오전반차", "오후반차", "기타"]

SCHEDULE_HEADERS = ["이름", "날짜", "유형", "출근", "퇴근", "비고", "상태"]
REQUEST_HEADERS = [
    "요청ID", "구분", "이름", "날짜",
    "기존유형", "기존출근", "기존퇴근",
    "신규유형", "신규출근", "신규퇴근",
    "사유", "상태", "요청일시", "처리일시", "처리메모",
]

st.set_page_config(page_title="마컴 근무 스케줄", page_icon="🗓️", layout="wide")


# ─────────────────────────────────────────────
# Google Sheets 연결
# ─────────────────────────────────────────────
@st.cache_resource
def get_spreadsheet():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    gc = gspread.authorize(creds)
    return gc.open_by_key(st.secrets["app"]["sheet_key"])


def get_ws(name: str, headers: list):
    ss = get_spreadsheet()
    try:
        ws = ss.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=name, rows=1000, cols=len(headers))
        ws.append_row(headers)
        return ws
    if ws.row_values(1) != headers:
        ws.update("A1", [headers])
    return ws


@st.cache_data(ttl=20)
def load_df(name: str) -> pd.DataFrame:
    headers = SCHEDULE_HEADERS if name == "schedule" else REQUEST_HEADERS
    ws = get_ws(name, headers)
    df = pd.DataFrame(ws.get_all_records())
    if df.empty:
        df = pd.DataFrame(columns=headers)
    return df.astype(str)


def clear_cache():
    load_df.clear()


def now_str() -> str:
    return datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")


# ─────────────────────────────────────────────
# 이메일 (Gmail SMTP)
# ─────────────────────────────────────────────
def send_mail(subject: str, body: str, to: str = None) -> bool:
    """to 미지정 시 팀장에게 발송"""
    try:
        cfg = st.secrets["smtp"]
        msg = MIMEText(body, _charset="utf-8")
        msg["Subject"] = subject
        msg["From"] = cfg["sender"]
        msg["To"] = to or st.secrets["app"]["leader_email"]
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15) as s:
            s.login(cfg["sender"], cfg["app_password"])
            s.send_message(msg)
        return True
    except Exception as e:
        st.warning(f"이메일 발송 실패 (처리는 정상 반영됨): {e}")
        return False


def member_email(name: str) -> str:
    return str(st.secrets["users"].get(name, {}).get("email", "")).strip()


def notify_member(req_row, approved: bool, memo: str):
    """승인/반려 결과를 요청한 팀원에게 회신. email 미설정 팀원은 건너뜀."""
    addr = member_email(req_row["이름"])
    if not addr:
        return
    result = "승인" if approved else "반려"
    if req_row["구분"] == "주간신규":
        detail = req_row["사유"]  # 주간 요약이 저장되어 있음
        outcome = "요청한 주간 스케줄 전체가 확정되었습니다." if approved \
            else "주간 등록이 취소되었습니다. 스케줄을 수정하여 다시 제출해주세요."
        subject = f"[근무스케줄] {req_row['날짜']} 주간 등록 {result} 안내"
        body = (
            f"{req_row['이름']} 님, {req_row['날짜']} 주간 스케줄 등록이 [{result}]되었습니다.\n\n"
            f"제출 내용: {detail}\n결과: {outcome}\n처리 시각: {now_str()}\n"
        )
    else:
        old_label = label_of(req_row["기존유형"], req_row["기존출근"], req_row["기존퇴근"])
        new_label = label_of(req_row["신규유형"], req_row["신규출근"], req_row["신규퇴근"])
        final = new_label if approved else f"{old_label} (기존 스케줄 유지)"
        subject = f"[근무스케줄] {req_row['날짜']} 변경 요청 {result} 안내"
        body = (
            f"{req_row['이름']} 님, {req_row['날짜']} 스케줄 변경 요청이 [{result}]되었습니다.\n\n"
            f"요청 내용: {old_label} → {new_label}\n최종 확정: {final}\n처리 시각: {now_str()}\n"
        )
    if memo:
        body += f"팀장 메모: {memo}\n"
    body += "\n자세한 내역은 앱의 '내 요청 현황' 탭에서 확인하세요."
    send_mail(subject, body, to=addr)


# ─────────────────────────────────────────────
# 헬퍼
# ─────────────────────────────────────────────
def find_schedule_row(df: pd.DataFrame, name: str, d: str):
    hit = df[(df["이름"] == name) & (df["날짜"] == d)]
    return None if hit.empty else hit.index[0]


def update_schedule_cells(df_idx: int, values: dict):
    ws = get_ws("schedule", SCHEDULE_HEADERS)
    sheet_row = df_idx + 2
    for col_name, val in values.items():
        ws.update_cell(sheet_row, SCHEDULE_HEADERS.index(col_name) + 1, val)


def delete_schedule_rows(df_indices: list):
    """행 삭제 시 인덱스 밀림 방지를 위해 아래쪽부터 삭제"""
    ws = get_ws("schedule", SCHEDULE_HEADERS)
    for df_idx in sorted(df_indices, reverse=True):
        ws.delete_rows(df_idx + 2)


def update_request_cells(df: pd.DataFrame, req_id: str, values: dict):
    hit = df[df["요청ID"] == req_id]
    if hit.empty:
        return
    ws = get_ws("requests", REQUEST_HEADERS)
    sheet_row = hit.index[0] + 2
    for col_name, val in values.items():
        ws.update_cell(sheet_row, REQUEST_HEADERS.index(col_name) + 1, val)


def fmt_time(t) -> str:
    return t.strftime("%H:%M") if isinstance(t, time) else str(t)


def parse_hhmm(s: str, fallback: time) -> time:
    try:
        h, m = str(s).split(":")
        return time(int(h), int(m))
    except Exception:
        return fallback


def label_of(wtype: str, t_in: str, t_out: str) -> str:
    if wtype == "연차":
        return "연차"
    base = f"{t_in}–{t_out}"
    if wtype and wtype != "근무":
        base += f" ({wtype})"
    return base


def schedule_label(row) -> str:
    return label_of(row["유형"], row["출근"], row["퇴근"])


def next_monday(base: date) -> date:
    return base + timedelta(days=(7 - base.weekday()) % 7 or 7)


def week_dates(monday: date) -> list:
    return [monday + timedelta(days=i) for i in range(5)]


# ─────────────────────────────────────────────
# 로그인 (이름 선택만)
# ─────────────────────────────────────────────
users_cfg = st.secrets["users"]
LEADER = next(n for n, u in users_cfg.items() if u["role"] == "leader")

if "user" not in st.session_state:
    st.session_state.user = None

with st.sidebar:
    st.title("🗓️ 마컴 근무 스케줄")
    if st.session_state.user is None:
        sel = st.selectbox("이름을 선택하세요", list(users_cfg.keys()))
        if st.button("로그인", use_container_width=True, type="primary"):
            st.session_state.user = sel
            st.rerun()
        st.stop()
    else:
        user = st.session_state.user
        role_label = "팀장" if users_cfg[user]["role"] == "leader" else "팀원"
        st.success(f"{user} ({role_label})")
        if st.button("로그아웃", use_container_width=True):
            st.session_state.user = None
            st.rerun()

user = st.session_state.user
is_leader = users_cfg[user]["role"] == "leader"

req_df = load_df("requests")
pending = req_df[req_df["상태"] == "대기"]
pending_chg = pending[pending["구분"] == "변경"]
if is_leader and not pending.empty:
    st.sidebar.error(f"🔔 승인 대기 {len(pending)}건")


# ─────────────────────────────────────────────
# 전체 스케줄 표
# ─────────────────────────────────────────────
def render_week_table(monday: date):
    sch = load_df("schedule")
    days = week_dates(monday)
    names = list(users_cfg.keys())
    data = {}
    for d in days:
        ds = d.isoformat()
        col = f"{d.month}/{d.day}({WEEKDAY_KR[d.weekday()]})"
        vals = []
        for n in names:
            row = sch[(sch["이름"] == n) & (sch["날짜"] == ds)]
            if row.empty:
                vals.append("—")
                continue
            r = row.iloc[0]
            label = schedule_label(r)
            if r["상태"] == "승인대기":
                pr = pending_chg[(pending_chg["이름"] == n) & (pending_chg["날짜"] == ds)]
                if not pr.empty:
                    p = pr.iloc[0]
                    label = f"{label} → {label_of(p['신규유형'], p['신규출근'], p['신규퇴근'])} ⏳승인대기"
                else:
                    label += " ⏳승인대기"
            vals.append(label)
        data[col] = vals
    st.dataframe(pd.DataFrame(data, index=names), use_container_width=True)
    st.caption("⏳승인대기: 팀장 승인 전 미확정 상태 (변경 건은 승인 전까지 기존 스케줄이 유효)")


def tab_schedule(container):
    with container:
        today = datetime.now(KST).date()
        this_mon = today - timedelta(days=today.weekday())
        offset = st.radio("주 선택", ["이번 주", "다음 주", "직접 선택"], horizontal=True)
        if offset == "이번 주":
            mon = this_mon
        elif offset == "다음 주":
            mon = this_mon + timedelta(days=7)
        else:
            picked = st.date_input("주 시작일(월요일)", value=this_mon)
            mon = picked - timedelta(days=picked.weekday())
        st.caption(f"{mon.isoformat()} ~ {(mon + timedelta(days=4)).isoformat()}")
        render_week_table(mon)


# ─────────────────────────────────────────────
# 주간 스케줄 등록 (팀원: 승인 필요 / 팀장: 즉시 확정)
# ─────────────────────────────────────────────
def tab_register(container):
    with container:
        st.subheader("주간 스케줄 등록")
        if is_leader:
            st.caption("팀장 등록은 즉시 확정됩니다.")
        else:
            st.caption("제출하면 팀장 승인 후 확정됩니다. 이미 등록/제출된 날짜는 잠기며, 확정된 날짜는 '변경 요청'으로 수정하세요.")

        today = datetime.now(KST).date()
        picked = st.date_input("등록할 주의 월요일", value=next_monday(today), key="reg_week")
        mon = picked - timedelta(days=picked.weekday())
        days = week_dates(mon)

        sch = load_df("schedule")
        my_rows = sch[sch["이름"] == user]
        existing = {r["날짜"]: r for _, r in my_rows.iterrows()}

        # 지난주 복사
        if st.button("↩️ 지난주 스케줄 복사"):
            copied = 0
            for d in days:
                ds = d.isoformat()
                if ds in existing:
                    continue
                prev = my_rows[my_rows["날짜"] == (d - timedelta(days=7)).isoformat()]
                if prev.empty:
                    continue
                p = prev.iloc[0]
                st.session_state[f"wt_{ds}"] = p["유형"] if p["유형"] in WORK_TYPES else "근무"
                st.session_state[f"in_{ds}"] = parse_hhmm(p["출근"], time(9, 0))
                st.session_state[f"out_{ds}"] = parse_hhmm(p["퇴근"], time(18, 0))
                st.session_state[f"memo_{ds}"] = p["비고"]
                copied += 1
            if copied:
                st.info(f"지난주 {copied}일 복사됨 — 연차·반차도 그대로 복사되니 제출 전 확인하세요.")
            else:
                st.warning("복사할 지난주 스케줄이 없거나, 이 주는 이미 모두 등록되어 있습니다.")

        entries = {}
        for d in days:
            ds = d.isoformat()
            day_label = f"{d.month}/{d.day}({WEEKDAY_KR[d.weekday()]})"
            c1, c2, c3, c4 = st.columns([1.2, 1.5, 1.5, 2])
            if ds in existing:
                r = existing[ds]
                c1.markdown(f"**{day_label}**")
                if r["상태"] == "승인대기":
                    c2.warning(f"{schedule_label(r)} — 승인 대기 중")
                else:
                    c2.info(f"{schedule_label(r)} 확정 → 수정은 변경 요청")
                continue
            with c1:
                st.markdown(f"**{day_label}**")
                wt = st.selectbox("유형", WORK_TYPES, key=f"wt_{ds}", label_visibility="collapsed")
            t_in = c2.time_input("출근", value=time(9, 0), key=f"in_{ds}", step=1800)
            t_out = c3.time_input("퇴근", value=time(18, 0), key=f"out_{ds}", step=1800)
            memo = c4.text_input("비고", key=f"memo_{ds}", label_visibility="collapsed", placeholder="비고")
            entries[ds] = (wt, t_in, t_out, memo)

        if not entries:
            return

        btn_label = "이 주 스케줄 등록 (즉시 확정)" if is_leader else "주간 스케줄 제출 (팀장 승인 후 확정)"
        if st.button(btn_label, type="primary"):
            status = "확정" if is_leader else "승인대기"
            rows, summary = [], []
            for ds in sorted(entries.keys()):
                wt, t_in, t_out, memo = entries[ds]
                d = date.fromisoformat(ds)
                if wt == "연차":
                    rows.append([user, ds, wt, "", "", memo, status])
                    summary.append(f"{WEEKDAY_KR[d.weekday()]} 연차")
                else:
                    rows.append([user, ds, wt, fmt_time(t_in), fmt_time(t_out), memo, status])
                    tag = f"({wt})" if wt != "근무" else ""
                    summary.append(f"{WEEKDAY_KR[d.weekday()]} {fmt_time(t_in)}–{fmt_time(t_out)}{tag}")
            get_ws("schedule", SCHEDULE_HEADERS).append_rows(rows)

            if is_leader:
                clear_cache()
                st.success(f"{len(rows)}건 등록 완료 (확정)")
                st.rerun()

            date_range = f"{min(entries)}~{max(entries)}"
            summary_txt = " / ".join(summary)
            get_ws("requests", REQUEST_HEADERS).append_row([
                uuid.uuid4().hex[:8], "주간신규", user, date_range,
                "", "", "", "", "", "",
                summary_txt, "대기", now_str(), "", "",
            ])
            clear_cache()
            send_mail(
                subject=f"[근무스케줄] {user} 주간 등록 요청 — {date_range}",
                body=(
                    f"{user} 님이 주간 스케줄 등록을 요청했습니다.\n\n"
                    f"기간: {date_range}\n내용: {summary_txt}\n요청 시각: {now_str()}\n\n"
                    f"앱에 접속해 승인/반려를 처리해주세요."
                ),
            )
            st.success(f"{len(rows)}건 제출 완료 — 팀장 승인 후 확정됩니다.")
            st.rerun()


# ─────────────────────────────────────────────
# 변경 요청 (확정 스케줄 대상)
# ─────────────────────────────────────────────
def tab_change_request(container):
    with container:
        st.subheader("스케줄 변경 요청")
        st.caption("확정된 스케줄만 변경 요청할 수 있습니다. 승인 전까지 기존 스케줄이 유효해요.")
        sch = load_df("schedule")
        my_pending_chg_dates = set(pending_chg[pending_chg["이름"] == user]["날짜"])
        mine = sch[(sch["이름"] == user) &
                   ((sch["상태"] == "확정") | (sch["날짜"].isin(my_pending_chg_dates)))].sort_values("날짜")
        if mine.empty:
            st.info("변경 가능한 확정 스케줄이 없습니다. 주간 스케줄을 먼저 등록하고 승인을 받아주세요.")
            return

        target_date = st.selectbox(
            "변경할 날짜",
            mine["날짜"].tolist(),
            format_func=lambda ds: f"{ds} ({WEEKDAY_KR[date.fromisoformat(ds).weekday()]}) — 현재: "
                                   f"{schedule_label(mine[mine['날짜'] == ds].iloc[0])}",
        )
        cur = mine[mine["날짜"] == target_date].iloc[0]

        if target_date in my_pending_chg_dates:
            st.warning("이 날짜는 승인 대기 중인 변경 요청이 있습니다. 새로 제출하면 기존 대기 건은 취소됩니다.")

        c1, c2, c3 = st.columns(3)
        cur_type_idx = WORK_TYPES.index(cur["유형"]) if cur["유형"] in WORK_TYPES else 0
        new_type = c1.selectbox("변경 유형", WORK_TYPES, index=cur_type_idx)
        new_in = c2.time_input("새 출근", value=parse_hhmm(cur["출근"], time(9, 0)), step=1800)
        new_out = c3.time_input("새 퇴근", value=parse_hhmm(cur["퇴근"], time(18, 0)), step=1800)
        reason = st.text_input("변경 사유 (필수)", placeholder="예: 오후 병원 방문으로 조기 출근")

        if st.button("변경 요청 제출", type="primary"):
            if not reason.strip():
                st.error("변경 사유를 입력해주세요.")
                return
            req_all = load_df("requests")
            dup = req_all[(req_all["이름"] == user) & (req_all["날짜"] == target_date) &
                          (req_all["상태"] == "대기") & (req_all["구분"] == "변경")]
            for _, dp in dup.iterrows():
                update_request_cells(req_all, dp["요청ID"], {
                    "상태": "취소", "처리일시": now_str(), "처리메모": "신규 요청으로 대체",
                })

            n_in = "" if new_type == "연차" else fmt_time(new_in)
            n_out = "" if new_type == "연차" else fmt_time(new_out)
            get_ws("requests", REQUEST_HEADERS).append_row([
                uuid.uuid4().hex[:8], "변경", user, target_date,
                cur["유형"], cur["출근"], cur["퇴근"],
                new_type, n_in, n_out,
                reason, "대기", now_str(), "", "",
            ])
            idx = find_schedule_row(sch, user, target_date)
            update_schedule_cells(idx, {"상태": "승인대기"})
            clear_cache()

            send_mail(
                subject=f"[근무스케줄] {user} 변경 요청 — {target_date}",
                body=(
                    f"{user} 님이 스케줄 변경을 요청했습니다.\n\n"
                    f"날짜: {target_date} ({WEEKDAY_KR[date.fromisoformat(target_date).weekday()]})\n"
                    f"변경 전: {schedule_label(cur)}\n"
                    f"변경 후: {label_of(new_type, n_in, n_out)}\n"
                    f"사유: {reason}\n요청 시각: {now_str()}\n\n"
                    f"앱에 접속해 승인/반려를 처리해주세요."
                ),
            )
            st.success("변경 요청이 제출되었습니다. 팀장 승인 후 확정됩니다.")
            st.rerun()


# ─────────────────────────────────────────────
# 내 요청 현황 (팀원)
# ─────────────────────────────────────────────
def tab_my_requests(container):
    with container:
        st.subheader("내 요청 현황")
        req_all = load_df("requests")
        mine = req_all[req_all["이름"] == user].sort_values("요청일시", ascending=False)
        if mine.empty:
            st.info("요청 내역이 없습니다.")
            return
        show = mine[["구분", "날짜", "기존출근", "기존퇴근", "신규유형", "신규출근", "신규퇴근",
                     "사유", "상태", "요청일시", "처리메모"]]
        st.dataframe(show, use_container_width=True, hide_index=True)
        st.caption("주간신규 반려 시 해당 주 등록은 취소되며, 수정 후 다시 제출하면 됩니다.")


# ─────────────────────────────────────────────
# 승인 대기 (팀장) — 주간신규는 일괄, 변경은 건별
# ─────────────────────────────────────────────
def tab_approvals(container):
    with container:
        st.subheader("승인 대기 요청")
        if pending.empty:
            st.success("대기 중인 요청이 없습니다. ✨")
            return
        sch = load_df("schedule")
        for _, r in pending.sort_values("요청일시").iterrows():
            with st.container(border=True):
                if r["구분"] == "주간신규":
                    start, _, end = r["날짜"].partition("~")
                    st.markdown(
                        f"🗂️ **주간 신규** — **{r['이름']}** · {r['날짜']}  \n"
                        f"{r['사유']}  \n"
                        f"요청: {r['요청일시']}"
                    )
                else:
                    st.markdown(
                        f"🔁 **변경** — **{r['이름']}** · {r['날짜']} "
                        f"({WEEKDAY_KR[date.fromisoformat(r['날짜']).weekday()]})  \n"
                        f"변경 전 `{label_of(r['기존유형'], r['기존출근'], r['기존퇴근'])}` → "
                        f"변경 후 `{label_of(r['신규유형'], r['신규출근'], r['신규퇴근'])}`  \n"
                        f"사유: {r['사유']} · 요청: {r['요청일시']}"
                    )
                memo = st.text_input("처리 메모(반려 시 사유)", key=f"memo_{r['요청ID']}")
                c1, c2 = st.columns(2)
                approve = c1.button("✅ 승인", key=f"ok_{r['요청ID']}", use_container_width=True)
                reject = c2.button("❌ 반려", key=f"no_{r['요청ID']}", use_container_width=True)
                if not (approve or reject):
                    continue

                if r["구분"] == "주간신규":
                    start, _, end = r["날짜"].partition("~")
                    week_rows = sch[(sch["이름"] == r["이름"]) &
                                    (sch["날짜"] >= start) & (sch["날짜"] <= end) &
                                    (sch["상태"] == "승인대기")]
                    if approve:
                        for idx in week_rows.index:
                            update_schedule_cells(idx, {"상태": "확정"})
                    else:
                        delete_schedule_rows(list(week_rows.index))
                else:
                    idx = find_schedule_row(sch, r["이름"], r["날짜"])
                    if idx is not None:
                        if approve:
                            update_schedule_cells(idx, {
                                "유형": r["신규유형"], "출근": r["신규출근"],
                                "퇴근": r["신규퇴근"], "상태": "확정",
                            })
                        else:
                            update_schedule_cells(idx, {"상태": "확정"})  # 기존 값 유지

                update_request_cells(req_df, r["요청ID"], {
                    "상태": "승인" if approve else "반려",
                    "처리일시": now_str(),
                    "처리메모": memo if approve else (memo or "사유 미기재"),
                })
                clear_cache()
                notify_member(r, approved=approve, memo=memo)
                st.rerun()


# ─────────────────────────────────────────────
# 요청 이력 (팀장)
# ─────────────────────────────────────────────
def tab_history(container):
    with container:
        st.subheader("전체 요청 이력")
        req_all = load_df("requests")
        if req_all.empty:
            st.info("이력이 없습니다.")
            return
        st.dataframe(req_all.sort_values("요청일시", ascending=False),
                     use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────
# 탭 구성 및 렌더링
# ─────────────────────────────────────────────
if is_leader:
    tabs = st.tabs([f"✅ 승인 대기 ({len(pending)})", "📅 전체 스케줄", "📝 내 스케줄 등록", "📚 요청 이력"])
    tab_approvals(tabs[0])
    tab_schedule(tabs[1])
    tab_register(tabs[2])
    tab_history(tabs[3])
else:
    tabs = st.tabs(["📅 전체 스케줄", "📝 주간 스케줄 등록", "🔁 변경 요청", "📋 내 요청 현황"])
    tab_schedule(tabs[0])
    tab_register(tabs[1])
    tab_change_request(tabs[2])
    tab_my_requests(tabs[3])
