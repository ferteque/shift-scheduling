import pandas as pd
import pulp
import warnings

def model_problem():
    # Silenciar avisos de format d'Excel
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    # 1. CARREGAR DEFINICIÓ DE TORNS
    try:
        shiftdf = pd.read_excel("shifts.xlsx")
        shiftdf.columns = shiftdf.columns.str.strip().str.lower()
        shifts = shiftdf.to_dict('records')
        num_shifts_per_day = len(shifts)
        total_periods = 7 * num_shifts_per_day
    except Exception as e:
        print(f"❌ Error llegint shifts.xlsx: {e}")
        return None

    # 2. CARREGAR TREBALLADORS
    try:
        workerdf = pd.read_excel("workers.xlsx", header=0)
        workerdf.columns = workerdf.columns.str.strip()
        workers_data = {}
        
        for _, row in workerdf.iterrows():
            name = row.iloc[0]
            workers_data[name] = {"period_avail": []}
            
            for day in range(7):
                # Accés per posició per evitar errors de noms de columnes
                w_start = row.iloc[1 + day * 2]
                w_end = row.iloc[2 + day * 2]
                
                for s in shifts:
                    # Comprovem si el treballador cobreix el torn sencer
                    can_work = int((w_start <= s['start']) and (w_end >= s['end']))
                    workers_data[name]["period_avail"].append(can_work)
    except Exception as e:
        print(f"❌ Error llegint workers.xlsx: {e}")
        return None

    # 3. CARREGAR DEMANDA
    try:
        # Llegim la primera columna del fitxer requirements
        requirements = pd.read_excel("requirements.xlsx", header=None).iloc[:, 0].tolist()
        if len(requirements) < total_periods:
            print(f"⚠️ Alerta: requirements.xlsx només té {len(requirements)} valors, se'n necessiten {total_periods}.")
            return None
    except Exception as e:
        print(f"❌ Error llegint requirements.xlsx: {e}")
        return None

    # 4. DEFINICIÓ DEL PROBLEMA
    problem = pulp.LpProblem("ScheduleWorkers", pulp.LpMinimize)

    # Crear variables de decisió
    for name in workers_data:
        workers_data[name]["worked_periods"] = [
            pulp.LpVariable(f"x_{name.replace(' ', '_')}_{p}", cat=pulp.LpBinary, 
                            upBound=workers_data[name]["period_avail"][p])
            for p in range(total_periods)
        ]

    # 5. RESTRICCIONS

    # A) Cobrir la demanda mínima de cada torn
    for p in range(total_periods):
        problem += pulp.lpSum([workers_data[name]["worked_periods"][p] for name in workers_data]) >= requirements[p]

    # B) Restriccions individuals
    for name in workers_data:
        total_worked = pulp.lpSum(workers_data[name]["worked_periods"])
        
        # --- RESTRICCIÓ D'EQUITAT (Mínims i Màxims) ---
        # Ajusta aquests valors segons el teu contracte (ex: 4 a 5 torns per setmana)
        problem += total_worked >= 4, f"MinTorns_{name}"
        problem += total_worked <= 5, f"MaxTorns_{name}"

        # C) Màxim un torn per dia (Evita sobreposicions)
        for day in range(7):
            day_start = day * num_shifts_per_day
            day_end = (day + 1) * num_shifts_per_day
            problem += pulp.lpSum(workers_data[name]["worked_periods"][day_start:day_end]) <= 1

    # 6. RESOLUCIÓ
    try:
        # Importem el solver de forma nativa per evitar problemes d'executable
        from pulp import HiGHS_CMD
        
        # Intentem forçar el solver a través de la interfície nativa de highspy
        # Si 'highspy' està instal·lat, PuLP l'hauria de detectar així:
        solver = pulp.getSolver('HiGHS', msg=0)
        
        status = problem.solve(solver)
        
        if status != pulp.LpStatusOptimal:
            print(f"⚠️ Estat del solver: {pulp.LpStatus[status]}")
            if status == pulp.LpStatusInfeasible:
                print("❌ Impossible: Revisa si demanes més gent de la que tens o si les disponibilitats són massa curtes.")
            return None
            
    except Exception as e:
        print(f"⚠️ Error amb HiGHS: {e}")
        print("Intentant Pla de Reserva: Solver CBC del sistema...")
        try:
            # Si HiGHS falla, provem amb el CBC que hem instal·lat abans via 'apt'
            # No posem ruta manual, deixem que PuLP el busqui al sistema
            solver_cbc = pulp.PULP_CBC_CMD(msg=0)
            status = problem.solve(solver_cbc)
            if status != pulp.LpStatusOptimal:
                return None
        except Exception as e2:
            print(f"❌ Cap solver funciona: {e2}")
            return None

    # 7. GENERAR RESULTATS
    output = []
    for name in workers_data:
        row = [name]
        # Creem una llista buida per cada dia per omplir-la ordenadament
        schedule_by_day = ["" for _ in range(7)]
        for p in range(total_periods):
            var = workers_data[name]["worked_periods"][p]
            if var.varValue is not None and var.varValue > 0.9:
                day = p // num_shifts_per_day
                shift_idx = p % num_shifts_per_day
                schedule_by_day[day] = f"Torn {shift_idx + 1}"
        
        row.extend(schedule_by_day)
        output.append(row)

    return output

if __name__ == "__main__":
    result = model_problem()
    
    if result:
        # Columnes per al CSV
        cols = ["Treballador", "Dilluns", "Dimarts", "Dimecres", "Dijous", "Divendres", "Dissabte", "Diumenge"]
        df_out = pd.DataFrame(result, columns=cols)
        df_out.to_csv("schedule_resultat.csv", index=False)
        print("✅ Quadrant generat amb èxit a 'schedule_resultat.csv'")
    else:
        print("❌ No s'ha pogut generar el quadrant.")