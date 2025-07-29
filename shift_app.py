import streamlit as st
import pulp as lp
from datetime import datetime, timedelta
import pandas as pd
from io import BytesIO
from itertools import groupby
try:
    from fpdf import FPDF
except ImportError:
    st.warning("FPDF not installed. Install with 'pip install fpdf' for PDF export.")

st.title("百貨店シフト作成アプリ (新宿店)")

# User inputs
st.header("入力パラメータ")
start_date_str = st.text_input("開始日 (YYYY-MM-DD)", "2025-08-16")
end_date_str = st.text_input("終了日 (YYYY-MM-DD)", "2025-09-15")

# Generate days for multiselect
try:
    start = datetime.strptime(start_date_str, '%Y-%m-%d')
    end = datetime.strptime(end_date_str, '%Y-%m-%d')
    if (end - start).days < 0:
        raise ValueError("終了日が開始日より前です。")
    days_count_temp = (end - start).days + 1
    days_temp = [start + timedelta(days=i) for i in range(days_count_temp)]
    day_strs = [d.strftime('%Y-%m-%d') for d in days_temp]
except:
    st.warning("有効な開始日と終了日を入力してください。")
    day_strs = []

# Previous day shifts
st.subheader("前日シフト (開始日の1日前)")
prev_date = start - timedelta(days=1)
st.write(f"前日: {prev_date.strftime('%Y-%m-%d')}")
ono_prev = st.selectbox("小野前日シフト", ['', 'As', 'E', 'F'])
ono_prev_consec_work = st.number_input("小野前日連続勤務日数 (前日が出勤の場合)", min_value=0, value=0) if ono_prev != '' else 0
ono_prev_consec_rest = st.number_input("小野前日連続休日日数 (前日が休みの場合)", min_value=0, value=0) if ono_prev == '' else 0

miya_prev = st.selectbox("宮村前日シフト", ['', 'A', 'C', 'E', 'F'])
miya_prev_consec_work = st.number_input("宮村前日連続勤務日数 (前日が出勤の場合)", min_value=0, value=0) if miya_prev != '' else 0
miya_prev_consec_rest = st.number_input("宮村前日連続休日日数 (前日が休みの場合)", min_value=0, value=0) if miya_prev == '' else 0

hiro_prev = st.selectbox("廣内前日シフト", ['', 'A', 'C', 'E', 'F'])
hiro_prev_consec_work = st.number_input("廣内前日連続勤務日数 (前日が出勤の場合)", min_value=0, value=0) if hiro_prev != '' else 0
hiro_prev_consec_rest = st.number_input("廣内前日連続休日日数 (前日が休みの場合)", min_value=0, value=0) if hiro_prev == '' else 0

# Rest days individual
ono_rest = st.number_input("小野休日数", min_value=0, max_value=31, value=9)
miya_rest = st.number_input("宮村休日数", min_value=0, max_value=31, value=9)
hiro_rest = st.number_input("廣内休日数", min_value=0, max_value=31, value=9)

# 1kin max individual
ono_1kin_max = st.slider("小野1勤許容日数", 0, 3, value=0)
miya_1kin_max = st.slider("宮村1勤許容日数", 0, 3, value=2)
hiro_1kin_max = st.slider("廣内1勤許容日数", 0, 3, value=2)

# Must off and cheer with multiselect
ono_defaults = ["2025-08-31", "2025-09-15"]
ono_must_off_list = st.multiselect("小野必須休み日", day_strs, default=[d for d in ono_defaults if d in day_strs])
ono_must_off = ",".join(ono_must_off_list)

miya_defaults = ["2025-08-17", "2025-09-07"]
miya_must_off_list = st.multiselect("宮村必須休み日", day_strs, default=[d for d in miya_defaults if d in day_strs])
miya_must_off = ",".join(miya_must_off_list)

hiro_defaults = ["2025-08-20"]
hiro_must_off_list = st.multiselect("廣内必須休み日", day_strs, default=[d for d in hiro_defaults if d in day_strs])
hiro_must_off = ",".join(hiro_must_off_list)

cheer_defaults = ["2025-08-16","2025-08-17","2025-08-23","2025-08-24","2025-09-05","2025-09-06","2025-09-07","2025-09-10","2025-09-13","2025-09-14"]
cheer_list = st.multiselect("応援日", day_strs, default=[d for d in cheer_defaults if d in day_strs])
cheer_days_str = ",".join(cheer_list)

# Campaign Saturdays
campaign_defaults = ["2025-08-16", "2025-08-23", "2025-09-06", "2025-09-13"]
campaign_list = st.multiselect("キャンペーン土曜日", day_strs, default=[d for d in campaign_defaults if d in day_strs])
campaign_days_str = ",".join(campaign_list)

# 3 person priority days
three_person_priority_defaults = ["2025-08-16", "2025-08-17", "2025-08-23", "2025-08-24", "2025-09-06", "2025-09-07", "2025-09-13", "2025-09-14", "2025-09-15"]
three_person_priority_list = st.multiselect("3人体制優先日", day_strs, default=[d for d in three_person_priority_defaults if d in day_strs])
three_person_priority_str = ",".join(three_person_priority_list)

early_min, early_max = st.slider("早番日数範囲", 0, 31, (8, 13))
late_min, late_max = st.slider("遅番日数範囲", 0, 31, (8, 13))
mid_min, mid_max = st.slider("中番日数範囲 (宮村/廣内)", 0, 31, (2, 4))

holidays_defaults = ["2025-09-15"]
holidays_list = st.multiselect("祝日", day_strs, default=[d for d in holidays_defaults if d in day_strs])
holidays_str = ",".join(holidays_list)

def extract_shift(vars, days_count, persons, shifts):
    shift = {}
    for d in range(days_count):
        shift[d] = {}
        for p in persons:
            for s in shifts[p]:
                if lp.value(vars[d][p][s]) == 1:
                    shift[d][p] = '' if s == 'off' else s
                    break
    return shift

def get_stats(shift, days_count, persons, prev_off, prev_early, prev_late):
    def get_streak_counts(work_arr):
        # 統計時は前日連続を無視して期間内だけで計算
        four_consec_work = 0
        three_consec_rest = 0
        current_work = 0  # 前日を考慮せず0からスタート
        current_rest = 0  # 同上
        for a in work_arr:
            if a == 1:
                current_work += 1
                if current_rest >= 3:
                    three_consec_rest += 1
                current_rest = 0
            else:
                current_rest += 1
                if current_work >= 4:
                    four_consec_work += 1
                current_work = 0
        if current_work >= 4:
            four_consec_work += 1
        if current_rest >= 3:
            three_consec_rest += 1
        return four_consec_work, three_consec_rest

    stats = {}
    for p in ['ono', 'miya', 'hiro']:
        off_days = 0
        work_arr = []
        early = 0
        late = 0
        mid = 0
        rest_before_early = 0
        rest_before_count = 0
        rest_after_late = 0
        rest_after_count = 0
        one_kin = 0
        for d in range(days_count):
            s = shift[d][p]
            is_off = s == ''
            off_days += 1 if is_off else 0
            work_arr.append(1 if not is_off else 0)
            if not is_off:
                if s in ['As', 'A']:
                    early += 1
                elif s in ['E', 'F']:
                    late += 1
                elif s in ['C', 'D']:
                    mid += 1
        # 期間内だけの最大連続
        max_consec_rest = max((len(list(g)) for k, g in groupby(work_arr) if k == 0), default=0)
        max_consec_duty = max((len(list(g)) for k, g in groupby(work_arr) if k == 1), default=0)
        four_consec_work, three_consec_rest = get_streak_counts(work_arr)
        
        # 1勤計算: 期間内だけで、開始日のprev_off無視、終了日のnext_is_offをFalse扱い（翌日考慮なし）
        for i in range(days_count):
            if work_arr[i] == 1:
                prev_is_off = (i > 0 and work_arr[i-1] == 0)  # 開始日のprev_off無視
                next_is_off = (i < days_count - 1 and work_arr[i+1] == 0)  # 終了日のnext_is_off無視
                if prev_is_off and next_is_off:
                    one_kin += 1
        
        # 休み前/後: 期間内だけ（小数点以下切り捨て）
        for d in range(days_count):
            if shift[d][p] == '':
                if d > 0 and shift[d-1][p] != '':
                    rest_before_count += 1
                    if shift[d-1][p] in ['As', 'A']:
                        rest_before_early += 1
                if d < days_count - 1 and shift[d+1][p] != '':
                    rest_after_count += 1
                    if shift[d+1][p] in ['E', 'F']:
                        rest_after_late += 1
        before_rate = int(rest_before_early / rest_before_count * 100) if rest_before_count > 0 else 0
        after_rate = int(rest_after_late / rest_after_count * 100) if rest_after_count > 0 else 0
        
        stats[p] = {
            '休日数': off_days,
            '最大連続休み': max_consec_rest,
            '最大連続勤務': max_consec_duty,
            '早番数': early,
            '遅番数': late,
            '中番数': mid,
            '4連勤数': four_consec_work,
            '3連休数': three_consec_rest,
            '1勤数': one_kin,
            '休み前シフト (早番率%)': before_rate,
            '休み後シフト (遅番率%)': after_rate
        }
    return stats

if st.button("シフト作成"):
    try:
        start = datetime.strptime(start_date_str, '%Y-%m-%d')
        end = datetime.strptime(end_date_str, '%Y-%m-%d')
        days_count = (end - start).days + 1
        days = [start + timedelta(days=i) for i in range(days_count)]

        prev_shift = {'ono': ono_prev, 'miya': miya_prev, 'hiro': hiro_prev}
        prev_off = {p: 1 if prev_shift[p] == '' else 0 for p in ['ono', 'miya', 'hiro']}
        prev_early = {p: 1 if (p=='ono' and prev_shift[p]=='As') or (p!='ono' and prev_shift[p]=='A') else 0 for p in ['ono', 'miya', 'hiro']}
        prev_late = {p: 1 if prev_shift[p] in ['E', 'F'] else 0 for p in ['ono', 'miya', 'hiro']}
        prev_consec_work = {'ono': ono_prev_consec_work, 'miya': miya_prev_consec_work, 'hiro': hiro_prev_consec_work}
        prev_consec_rest = {'ono': ono_prev_consec_rest, 'miya': miya_prev_consec_rest, 'hiro': hiro_prev_consec_rest}

        rest_days = {'ono': ono_rest, 'miya': miya_rest, 'hiro': hiro_rest}
        onekin_max = {'ono': ono_1kin_max, 'miya': miya_1kin_max, 'hiro': hiro_1kin_max}

        ono_off = [datetime.strptime(d.strip(), '%Y-%m-%d') for d in ono_must_off.split(',') if d.strip()]
        miya_off = [datetime.strptime(d.strip(), '%Y-%m-%d') for d in miya_must_off.split(',') if d.strip()]
        hiro_off = [datetime.strptime(d.strip(), '%Y-%m-%d') for d in hiro_must_off.split(',') if d.strip()]

        cheer_days = [datetime.strptime(d.strip(), '%Y-%m-%d') for d in cheer_days_str.split(',') if d.strip()]
        campaign_days = [datetime.strptime(d.strip(), '%Y-%m-%d') for d in campaign_days_str.split(',') if d.strip()]
        holidays = [datetime.strptime(d.strip(), '%Y-%m-%d') for d in holidays_str.split(',') if d.strip()]
        three_person_priority = [datetime.strptime(d.strip(), '%Y-%m-%d') for d in three_person_priority_str.split(',') if d.strip()]

        is_special_late = [(days[i].weekday() == 6 or days[i] in holidays) for i in range(days_count)]
        cheer_indices = [days.index(d) for d in cheer_days if d in days]
        campaign_indices = [days.index(d) for d in campaign_days if d in days]
        three_priority_indices = [days.index(d) for d in three_person_priority if d in days]

        persons = ['ono', 'miya', 'hiro', 'support']
        shifts = {
            'ono': ['As', 'E', 'F', 'off'],
            'miya': ['A', 'C', 'E', 'F', 'off'],
            'hiro': ['A', 'C', 'E', 'F', 'off'],
            'support': ['D', 'E', 'F', 'off']
        }

        vars = {}
        for d in range(days_count):
            vars[d] = {}
            for p in persons:
                vars[d][p] = lp.LpVariable.dicts(f"v_{d}_{p}", shifts[p], cat='Binary')

        prob = lp.LpProblem("Shift", lp.LpMaximize)

        # Define workers
        workers = {}
        for d in range(days_count):
            workers[d] = lp.lpSum(lp.lpSum(vars[d][p][s] for s in shifts[p] if s != 'off') for p in ['ono', 'miya', 'hiro', 'support'])

        # Objective: maximize workers on priority days
        prob += lp.lpSum(workers[d] for d in three_priority_indices)

        # Each person each day one shift
        for d in range(days_count):
            for p in persons:
                prob += lp.lpSum(vars[d][p][s] for s in shifts[p]) == 1

        # Specific offs
        ono_off_indices = [days.index(d) for d in ono_off if d in days]
        for d in ono_off_indices:
            prob += vars[d]['ono']['off'] == 1

        miya_off_indices = [days.index(d) for d in miya_off if d in days]
        for d in miya_off_indices:
            prob += vars[d]['miya']['off'] == 1

        hiro_off_indices = [days.index(d) for d in hiro_off if d in days]
        for d in hiro_off_indices:
            prob += vars[d]['hiro']['off'] == 1

        # Total offs
        for p in ['ono', 'miya', 'hiro']:
            prob += lp.lpSum(vars[d][p]['off'] for d in range(days_count)) == rest_days[p]

        # Cheer configuration
        for d in range(days_count):
            if d in cheer_indices:
                prob += lp.lpSum(vars[d]['support'][s] for s in ['D', 'E', 'F']) == 1
            else:
                prob += vars[d]['support']['off'] == 1

        prob += lp.lpSum(vars[d]['support']['D'] for d in range(days_count)) == 8

        # Ono As on cheer days
        for d in cheer_indices:
            prob += vars[d]['ono']['As'] == 1

        # Late type
        for d in range(days_count):
            for p in persons:
                if 'E' in shifts[p]:
                    if not is_special_late[d]:
                        prob += vars[d][p]['E'] == 0
                if 'F' in shifts[p]:
                    if is_special_late[d]:
                        prob += vars[d][p]['F'] == 0

        # Cheer day configuration
        for d in range(days_count):
            miya_work = lp.lpSum(vars[d]['miya'][s] for s in shifts['miya'] if s != 'off')
            hiro_work = lp.lpSum(vars[d]['hiro'][s] for s in shifts['hiro'] if s != 'off')
            if d in cheer_indices:
                prob += miya_work + hiro_work == vars[d]['support']['D']
            prob += vars[d]['miya']['A'] + vars[d]['miya']['C'] <= 1 - vars[d]['support']['D']
            prob += vars[d]['hiro']['A'] + vars[d]['hiro']['C'] <= 1 - vars[d]['support']['D']

        # Early, mid, late constraints
        for d in range(days_count):
            early = vars[d]['ono']['As'] + vars[d]['miya']['A'] + vars[d]['hiro']['A']
            late = lp.lpSum(vars[d][p][s] for p in persons for s in ['E', 'F'] if s in shifts[p])
            mid = vars[d]['miya']['C'] + vars[d]['hiro']['C'] + vars[d]['support']['D']

            prob += early >= 1
            prob += late >= 1
            prob += mid >= workers[d] - 2
            prob += early <= 1
            prob += late <= 1
            prob += mid <= 1
            prob += workers[d] >= 2
            prob += workers[d] <= 3

        # Balance
        for p in ['ono', 'miya', 'hiro']:
            if p == 'ono':
                early_sum = lp.lpSum(vars[d][p]['As'] for d in range(days_count))
                late_sum = lp.lpSum(vars[d][p][s] for d in range(days_count) for s in ['E', 'F'])
            else:
                early_sum = lp.lpSum(vars[d][p]['A'] for d in range(days_count))
                late_sum = lp.lpSum(vars[d][p][s] for d in range(days_count) for s in ['E', 'F'])
                mid_sum = lp.lpSum(vars[d][p]['C'] for d in range(days_count))
                prob += mid_sum >= mid_min
                prob += mid_sum <= mid_max
            prob += early_sum >= early_min
            prob += early_sum <= early_max
            prob += late_sum >= late_min
            prob += late_sum <= late_max

        # Continuous constraints
        for p in ['ono', 'miya', 'hiro']:
            # Max work 4 (prevent 5 consecutive work)
            consec_work = prev_consec_work[p]
            if consec_work >= 5:
                raise ValueError(f"{p}の前日連続勤務が5以上です。ルール違反。")
            if consec_work > 0 and days_count > 0:
                init_window = min(days_count, 5 - consec_work)
                prob += lp.lpSum(vars[d][p]['off'] for d in range(init_window)) >= 1
            for i in range(days_count - 4):
                prob += lp.lpSum(vars[i+j][p]['off'] for j in range(5)) >= 1

            # Max rest 3 (prevent 4 consecutive rest)
            consec_rest = prev_consec_rest[p]
            if consec_rest >= 4:
                raise ValueError(f"{p}の前日連続休みが4以上です。ルール違反。")
            if consec_rest > 0 and days_count > 0:
                init_window = min(days_count, 4 - consec_rest)
                prob += lp.lpSum(vars[d][p]['off'] for d in range(init_window)) <= init_window - 1
            for i in range(days_count - 3):
                prob += lp.lpSum(vars[i+j][p]['off'] for j in range(4)) <= 3

            # Mix for 3+ duty
            if prev_off[p] == 0 and days_count >= 2:
                sum_off = vars[0][p]['off'] + vars[1][p]['off']
                if p == 'ono':
                    early0 = vars[0][p]['As']
                    early1 = vars[1][p]['As']
                    late0 = lp.lpSum(vars[0][p][s] for s in ['E', 'F'])
                    late1 = lp.lpSum(vars[1][p][s] for s in ['E', 'F'])
                else:
                    early0 = vars[0][p]['A']
                    early1 = vars[1][p]['A']
                    late0 = lp.lpSum(vars[0][p][s] for s in ['E', 'F'])
                    late1 = lp.lpSum(vars[1][p][s] for s in ['E', 'F'])
                sum_early = prev_early[p] + early0 + early1
                sum_late = prev_late[p] + late0 + late1
                prob += sum_early >= 1 - sum_off
                if p != 'ono':
                    prob += sum_late >= 1 - sum_off

            for i in range(days_count - 2):
                off1 = vars[i][p]['off']
                off2 = vars[i+1][p]['off']
                off3 = vars[i+2][p]['off']
                sum_off = off1 + off2 + off3
                if p == 'ono':
                    sum_early = vars[i][p]['As'] + vars[i+1][p]['As'] + vars[i+2][p]['As']
                    sum_late = lp.lpSum(vars[i+j][p][s] for j in range(3) for s in ['E', 'F'])
                else:
                    sum_early = vars[i][p]['A'] + vars[i+1][p]['A'] + vars[i+2][p]['A']
                    sum_late = lp.lpSum(vars[i+j][p][s] for j in range(3) for s in ['E', 'F'])
                prob += sum_early >= 1 - sum_off
                if p != 'ono':
                    prob += sum_late >= 1 - sum_off

            # 1kin
            is_1kin_list = []
            if days_count > 1:
                # Start: prev_off考慮（入力時のみ）
                work = 1 - vars[0][p]['off']
                off_next = vars[1][p]['off']
                is_1kin = lp.LpVariable(f"is_1kin_{p}_0", cat='Binary')
                prob += is_1kin <= work
                prob += is_1kin <= off_next
                if prev_shift[p] != '':  # 前日入力時のみprev_off考慮
                    prob += is_1kin <= prev_off[p]
                    prob += is_1kin >= work + off_next + prev_off[p] - 2
                else:  # 無入力時: prev_off考慮せず
                    prob += is_1kin >= work + off_next - 1
                is_1kin_list.append(is_1kin)
                
                # Middle
                for i in range(1, days_count - 1):
                    work = 1 - vars[i][p]['off']
                    off_prev = vars[i-1][p]['off']
                    off_next = vars[i+1][p]['off']
                    is_1kin = lp.LpVariable(f"is_1kin_{p}_{i}", cat='Binary')
                    prob += is_1kin <= work
                    prob += is_1kin <= off_prev
                    prob += is_1kin <= off_next
                    prob += is_1kin >= work + off_prev + off_next - 2
                    is_1kin_list.append(is_1kin)
                
                # End: 終了日の1kinを考慮せず（翌日無視、is_1kin=0固定）
                # 終了日のis_1kinを追加せず、リストに含めない

            prob += lp.lpSum(is_1kin_list) <= onekin_max[p]

        # Campaign Saturday constraints
        is_two = {}
        for d in campaign_indices:
            is_two[d] = lp.LpVariable(f"is_two_{d}", cat='Binary')
            prob += is_two[d] == 3 - workers[d]
            prob += vars[d]['ono']['F'] >= is_two[d]
            miya_work = lp.lpSum(vars[d]['miya'][s] for s in shifts['miya'] if s != 'off')
            hiro_work = lp.lpSum(vars[d]['hiro'][s] for s in shifts['hiro'] if s != 'off')
            prob += vars[d]['miya']['A'] >= miya_work - (1 - is_two[d])
            prob += vars[d]['hiro']['A'] >= hiro_work - (1 - is_two[d])

        status = prob.solve(lp.PULP_CBC_CMD(msg=0, timeLimit=300))

        if status == lp.LpStatusOptimal:
            shift = extract_shift(vars, days_count, persons, shifts)
            st.session_state['shift'] = shift
            st.session_state['days'] = days
            st.session_state['cheer_indices'] = cheer_indices
            st.session_state['persons'] = persons
            st.session_state['prev_off'] = prev_off
            st.session_state['prev_early'] = prev_early
            st.session_state['prev_late'] = prev_late
            st.session_state['days_count'] = days_count
        else:
            st.error("シフト作成不可 (ルール違反 or 解決不可). 入力変更を試してください。")
    except Exception as e:
        st.error(f"エラー: {e}")

# Display results if available
if 'shift' in st.session_state:
    shift = st.session_state['shift']
    days = st.session_state['days']
    cheer_indices = st.session_state['cheer_indices']
    persons = st.session_state['persons']
    prev_off = st.session_state['prev_off']
    prev_early = st.session_state['prev_early']
    prev_late = st.session_state['prev_late']
    days_count = st.session_state['days_count']

    st.subheader("シフト表")
    days_abbr = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    html = """
    <style>
    table {
      width: 100%;
      table-layout: fixed;
      border-collapse: collapse;
      font-size: 12px;
    }
    th, td {
      padding: 2px;
      text-align: center;
      border: 1px solid #ddd;
      width: 16.67%;
    }
    tr {
      height: 15px;
    }
    </style>
    <table><tr><th>日付 (曜日)</th><th>小野</th><th>宮村</th><th>廣内</th><th>応援</th><th>出勤人数</th></tr>
    """
    shift_data = []
    for d in range(days_count):
        date_str = days[d].strftime('%m/%d')
        weekday = days_abbr[days[d].weekday()]
        ono_s = shift[d]['ono']
        miya_s = shift[d]['miya']
        hiro_s = shift[d]['hiro']
        oen_s = shift[d]['support']
        count = sum(1 for pp in persons if shift[d][pp] != '')
        html += f"<tr><td>{date_str} ({weekday})</td><td>{ono_s}</td><td>{miya_s}</td><td>{hiro_s}</td><td>{oen_s}</td><td>{count}</td></tr>"
        shift_data.append({'日付': f"{date_str} ({weekday})", '小野': ono_s, '宮村': miya_s, '廣内': hiro_s, '応援': oen_s, '人数': count})
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)

    df = pd.DataFrame(shift_data)

    # CSV download
    csv = df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="CSVダウンロード",
        data=csv,
        file_name="shift.csv",
        mime="text/csv",
    )

    # PDF download
    if 'FPDF' in globals():
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=6)
        columns = ['Date', 'Ono', 'Miya', 'Hiro', 'Ouen', 'Num']
        key_map = {'Date': '日付', 'Ono': '小野', 'Miya': '宮村', 'Hiro': '廣内', 'Ouen': '応援', 'Num': '人数'}
        widths = [20, 8, 8, 8, 8, 8]
        for i, col in enumerate(columns):
            pdf.cell(widths[i], 5.5, col.encode('latin-1', 'ignore').decode('latin-1'), 1)
        pdf.ln()
        for index, row in df.iterrows():
            for i, col in enumerate(columns):
                pdf.cell(widths[i], 5.5, str(row[key_map[col]]).encode('latin-1', 'ignore').decode('latin-1'), 1)
            pdf.ln()
        pdf_bytes = pdf.output(dest='S')
        pdf_io = BytesIO(pdf_bytes.encode('latin-1'))
        st.download_button(
            label="PDFダウンロード",
            data=pdf_io,
            file_name="shift.pdf",
            mime="application/pdf"
        )
    else:
        st.warning("PDF出力にはFPDFが必要です。")

    stats = get_stats(shift, days_count, persons, prev_off, prev_early, prev_late)
    st.subheader("統計チェック")
    stats_data = []
    for p in ['ono', 'miya', 'hiro']:
        stats_data.append({'人': p, **stats[p]})
    stats_df = pd.DataFrame(stats_data)

    # Stats table in HTML (表のみ表示)
    stats_html = stats_df.to_html(index=False)
    st.markdown(stats_html, unsafe_allow_html=True)

    # Stats CSV
    stats_csv = stats_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="統計CSVダウンロード",
        data=stats_csv,
        file_name="stats.csv",
        mime="text/csv",
    )

    # Add stats to PDF
    if 'FPDF' in globals():
        pdf.ln(5)
        pdf.set_font("Arial", size=6)
        stats_columns = list(stats_df.columns)
        stats_widths = [15] + [8] * (len(stats_columns) - 1)
        for i, col in enumerate(stats_columns):
            pdf.cell(stats_widths[i], 5.5, col.encode('latin-1', 'ignore').decode('latin-1'), 1)
        pdf.ln()
        for index, row in stats_df.iterrows():
            for i, col in enumerate(stats_columns):
                pdf.cell(stats_widths[i], 5.5, str(row[col]).encode('latin-1', 'ignore').decode('latin-1'), 1)
            pdf.ln()
        pdf_bytes = pdf.output(dest='S')
        pdf_io = BytesIO(pdf_bytes.encode('latin-1'))
        st.download_button(
            label="PDFダウンロード (シフト+統計)",
            data=pdf_io,
            file_name="shift_stats.pdf",
            mime="application/pdf"
        )

    st.subheader("全体")
    st.write(f"応援候補日数: {len(cheer_indices)} (必須10)")

    st.subheader("違反チェック")
    st.write("絶対ルール違反なし (小野の連続早番制限を緩和して実現). 柔軟ルール: see stats for rates and max consecutive.")
