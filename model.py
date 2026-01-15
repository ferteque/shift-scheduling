import pandas as pd
import pulp

def model_problem():
    # 1. CARREGAR DEFINICIÓ DE TORNS
    # Esperem un excel amb: shift_id, start, end
    shiftdf = pd.read_excel("shifts.xlsx")
    shifts = shiftdf.to_dict('records')
    num_shifts_per_day = len(shifts)
    total_periods = 7 * num_shifts_per_day

    # 2. CARREGAR TREBALLADORS
    # Estructura: Nom, Dilluns_Inici, Dilluns_Final, Dimarts_Inici...
    workerdf = pd.read_excel("workers.xlsx", header=0)
    workers_data = {}
    
    for _, row in workerdf.iterrows():
        name = row[0]
        workers_data[name] = {"period_avail": []}
        
        for day in range(7):
            w_start = row[1 + day * 2]
            w_end = row[2 + day * 2]
            
            for s in shifts:
                # El treballador pot fer el torn si la seva disponibilitat cobreix tot el torn
                can_work = int((w_start <= s['start']) and (w_end >= s['end']))
                workers_data[name]["period_avail"].append(can_work)

    # 3. CARREGAR DEMANDA (quants treballadors per cada torn)
    # Un excel amb una sola fila o columna amb els 21 valors (7 dies * 3 torns)
    requirements = pd.read_excel("requirements.xlsx", header=None).iloc[:, 0].tolist()

    # 4. DEFINICIÓ DEL PROBLEMA
    problem = pulp.LpProblem("ScheduleWorkers", pulp.LpMinimize)

    # Crear variables de decisió (X_treballador_torn)
    for name in workers_data:
        workers_data[name]["worked_periods"] = [
            pulp.LpVariable(f"x_{name}_{p}", cat=pulp.LpBinary, upBound=workers_data[name]["period_avail"][p])
            for p in range(total_periods)
        ]

    # FUNCIÓ OBJECTIU: Minimitzar el total de torns assignats (o optimitzar segons convingui)
    total_shifts = []
    for name in workers_data:
        total_shifts.extend(workers_data[name]["worked_periods"])
    problem += pulp.lpSum(total_shifts)

    # 5. RESTRICCIONS

    # A) Cobrir la demanda de cada torn
    for p in range(total_periods):
        problem += pulp.lpSum([workers_data[name]["worked_periods"][p] for name in workers_data]) >= requirements[p]

    # B) Màxim un torn per dia (Evita que si els torns se sobreposen, un treballador faci dos alhora)
    for name in workers_data:
        for day in range(7):
            day_start = day * num_shifts_per_day
            day_end = (day + 1) * num_shifts_per_day
            problem += pulp.lpSum(workers_data[name]["worked_periods"][day_start:day_end]) <= 1

    # C) Descans setmanal (Exemple: Mínim 2 dies lliures a la setmana)
    # Creem una variable auxiliar per saber si un treballador treballa un dia concret
    for name in workers_data:
        days_worked = []
        for day in range(7):
            is_working_day = pulp.LpVariable(f"working_{name}_day_{day}", cat=pulp.LpBinary)
            day_shifts = workers_data[name]["worked_periods"][day*num_shifts_per_day : (day+1)*num_shifts_per_day]
            
            # Si treballa qualsevol torn, is_working_day serà 1
            for s_var in day_shifts:
                problem += is_working_day >= s_var
            
            days_worked.append(is_working_day)
        
        # Obliguem a que treballi com a màxim 5 dies de 7
        problem += pulp.lpSum(days_worked) <= 5

    # 6. RESOLUCIÓ
    try:
        problem.solve(pulp.PULP_CBC_CMD(msg=0))
    except Exception as e:
        print(f"Error resolent el problema: {e}")

    # 7. GENERAR RESULTATS
    output = []
    for name in workers_data:
        row = [name]
        for p in range(total_periods):
            if pulp.value(workers_data[name]["worked_periods"][p]) == 1:
                day = p // num_shifts_per_day
                shift_idx = p % num_shifts_per_day
                row.append(f"Dia {day+1}-Torn {shift_idx+1}")
        output.append(row)

    return output

if __name__ == "__main__":
    schedule = model_problem()
    
    # Guardar a CSV
    df_out = pd.DataFrame(schedule)
    df_out.to_csv("schedule_resultat.csv", index=False, header=False)
    print("Quadrant generat amb èxit a 'schedule_resultat.csv'")