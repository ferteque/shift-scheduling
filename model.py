import pandas as pd
import pulp
import warnings

def model_problem():
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    # 1. CÀRREGA DE DADES (Shifts, Workers, Requirements)
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
        print(f"❌ Error carregant fitxers: {e}")
        return None

    # 2. DEFINICIÓ DEL PROBLEMA
    problem = pulp.LpProblem("ScheduleWorkers", pulp.LpMinimize)

    # Variables: Treballa el torn P?
    for name in workers_data:
        workers_data[name]["worked_periods"] = [
            pulp.LpVariable(f"x_{name.replace(' ', '_')}_{p}", cat=pulp.LpBinary, 
                            upBound=workers_data[name]["period_avail"][p])
            for p in range(total_periods)
        ]
        
        # Variables auxiliars: Treballa el dia D? (1 si fa qualsevol torn, 0 si lliure)
        workers_data[name]["working_days"] = [
            pulp.LpVariable(f"wd_{name.replace(' ', '_')}_{d}", cat=pulp.LpBinary)
            for d in range(7)
        ]

    # 3. RESTRICCIONS DE COBERTURA I BÀSIQUES
    for p in range(total_periods):
        problem += pulp.lpSum([workers_data[name]["worked_periods"][p] for name in workers_data]) >= requirements[p]

    for name in workers_data:
        # Enllaçar torns amb dies: si treballa un torn, el dia compta com a treballat
        for d in range(7):
            day_shifts = workers_data[name]["worked_periods"][d*num_shifts_per_day : (day+1)*num_shifts_per_day]
            for s_var in day_shifts:
                problem += workers_data[name]["working_days"][d] >= s_var
            # Màxim un torn per dia
            problem += pulp.lpSum(day_shifts) <= 1

        # Mínim i màxim de torns setmanals (Equitat)
        total_worked = pulp.lpSum(workers_data[name]["worked_periods"])
        problem += total_worked >= 4
        problem += total_worked <= 5

        # --- NOVA LÒGICA: DIES LLIURES CONSECUTIUS ---
        # Variable que detecta l'inici d'un bloc de treball (0 -> 1)
        # Si un treballador comença a treballar, descansa i torna, tindrà 2 "inicis".
        # Si els dies lliures estan junts, només tindrà 1 "inici" a la setmana.
        starts = [pulp.LpVariable(f"start_{name.replace(' ', '_')}_{d}", cat=pulp.LpBinary) for d in range(7)]
        for d in range(7):
            prev_day = workers_data[name]["working_days"][d-1] if d > 0 else 0
            problem += starts[d] >= workers_data[name]["working_days"][d] - prev_day

        # Minimitzant la suma de 'starts' obliguem al solver a crear un sol bloc de treball
        # i, per tant, un sol bloc de dies lliures.
        workers_data[name]["num_starts"] = pulp.lpSum(starts)

    # 4. FUNCIÓ OBJECTIU
    # Minimitzem el total de torns PERÒ també les interrupcions (els 'starts')
    # Posem un pes gran a 'num_starts' per prioritzar els dies junts
    obj_torns = pulp.lpSum([p for n in workers_data for p in workers_data[n]["worked_periods"]])
    obj_consecutivitat = pulp.lpSum([workers_data[n]["num_starts"] for n in workers_data])
    
    problem += obj_torns + (10 * obj_consecutivitat)

    # 5. RESOLUCIÓ (Amb GLPK per evitar errors de ruta)
    try:
        solver = pulp.GLPK_CMD(msg=0)
        status = problem.solve(solver)
        if status != pulp.LpStatusOptimal:
            print(f"⚠️ Estat: {pulp.LpStatus[status]}")
            return None
    except:
        print("❌ Instal·la GLPK: sudo apt install glpk-utils")
        return None

    # 6. RESULTATS
    output = []
    cols = ["Treballador", "Dilluns", "Dimarts", "Dimecres", "Dijous", "Divendres", "Dissabte", "Diumenge"]
    for name in workers_data:
        row = [name]
        for d in range(7):
            day_val = ""
            for s_idx in range(num_shifts_per_day):
                if pulp.value(workers_data[name]["worked_periods"][d*num_shifts_per_day + s_idx]) == 1:
                    day_val = f"Torn {s_idx+1}"
            row.append(day_val)
        output.append(row)

    pd.DataFrame(output, columns=cols).to_csv("schedule_resultat.csv", index=False)
    print("✅ Quadrant optimitzat amb dies consecutius generat!")

if __name__ == "__main__":
    model_problem()