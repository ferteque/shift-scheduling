import pandas as pd
import pulp
import warnings

def model_problem():
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    try:
        shiftdf = pd.read_excel("shifts.xlsx")
        shiftdf.columns = shiftdf.columns.str.strip().str.lower()
        shifts = shiftdf.to_dict('records')
        num_shifts_per_day = len(shifts)
        total_periods = 7 * num_shifts_per_day

        workerdf = pd.read_excel("workers.xlsx", header=0)
        workers_data = {row.iloc[0]: {"period_avail": []} for _, row in workerdf.iterrows()}
        for i, (name, data) in enumerate(workers_data.items()):
            row = workerdf.iloc[i]
            for day in range(7):
                w_start, w_end = row.iloc[1 + day * 2], row.iloc[2 + day * 2]
                for s in shifts:
                    can_work = int((w_start <= s['start']) and (w_end >= s['end']))
                    data["period_avail"].append(can_work)

        requirements = pd.read_excel("requirements.xlsx", header=None).iloc[:, 0].tolist()
    except Exception as e:
        print(f"❌ Error loading file: {e}")
        return None

    problem = pulp.LpProblem("ScheduleWorkers", pulp.LpMinimize)

    for name in workers_data:
        workers_data[name]["worked_periods"] = [
            pulp.LpVariable(f"x_{name.replace(' ', '_')}_{p}", cat=pulp.LpBinary, 
                            upBound=workers_data[name]["period_avail"][p])
            for p in range(total_periods)
        ]
        
        workers_data[name]["working_days"] = [
            pulp.LpVariable(f"wd_{name.replace(' ', '_')}_{d}", cat=pulp.LpBinary)
            for d in range(7)
        ]

    for p in range(total_periods):
        problem += pulp.lpSum([workers_data[name]["worked_periods"][p] for name in workers_data]) >= requirements[p]

    for name in workers_data:
        for d in range(7):
            day_shifts = workers_data[name]["worked_periods"][d*num_shifts_per_day : (day+1)*num_shifts_per_day]
            for s_var in day_shifts:
                problem += workers_data[name]["working_days"][d] >= s_var
            problem += pulp.lpSum(day_shifts) <= 1

        total_worked = pulp.lpSum(workers_data[name]["worked_periods"])
        problem += total_worked >= 4
        problem += total_worked <= 5

        
        starts = [pulp.LpVariable(f"start_{name.replace(' ', '_')}_{d}", cat=pulp.LpBinary) for d in range(7)]
        for d in range(7):
            prev_day = workers_data[name]["working_days"][d-1] if d > 0 else 0
            problem += starts[d] >= workers_data[name]["working_days"][d] - prev_day

        
        workers_data[name]["num_starts"] = pulp.lpSum(starts)

    
    obj_torns = pulp.lpSum([p for n in workers_data for p in workers_data[n]["worked_periods"]])
    obj_consecutivitat = pulp.lpSum([workers_data[n]["num_starts"] for n in workers_data])
    
    problem += obj_torns + (10 * obj_consecutivitat)

    try:
        solver = pulp.GLPK_CMD(msg=0)
        status = problem.solve(solver)
        if status != pulp.LpStatusOptimal:
            print(f"⚠️ Status: {pulp.LpStatus[status]}")
            return None
    except:
        print("❌ Install GLPK: sudo apt install glpk-utils")
        return None

    output = []
    cols = ["Worker", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    for name in workers_data:
        row = [name]
        for d in range(7):
            day_val = ""
            for s_idx in range(num_shifts_per_day):
                if pulp.value(workers_data[name]["worked_periods"][d*num_shifts_per_day + s_idx]) == 1:
                    day_val = f"Torn {s_idx+1}"
            row.append(day_val)
        output.append(row)

    pd.DataFrame(output, columns=cols).to_csv("schedule_results.csv", index=False)
    print("✅ Done!")

if __name__ == "__main__":
    model_problem()