import gurobipy as gp
from gurobipy import GRB
import csv
import pandas as pd
# list of indices of variables b[i]
materiales = [0,1,2]
periodos = [0,1,2,3, 4, 5]
periodos_2 = [1,2,3, 4, 5]

#parametros:
alph = 2
beta = 5
gamma= 300
delta = 500
s=4
epsilon = 6
phi=30

model = gp.Model('Minimizacion costos Plan 1')

# Variables b[i]

#listas de datos variables
cb = list()
cb.append(12) #coste bloque celulosa
cb.append(10)
cb.append(20)
q = list()
q.append(1) #tiempo para producir celulosa
q.append(2)
q.append(11)
PM= list()
PM.append(.25) #participacion en mercado de materiales
PM.append(.5)
PM.append(.25)
D = [2000, 3200, 3400, 3800, 2200, 2200] #demanda en 6 periodos


#popular parametros
Inv = model.addVars(6, vtype='I', name='I') # inventario
inv = model.addVars(3,6, vtype='I', name='i') #inv de cada producto
c = model.addVars(3,6, lb=0,vtype='I', name="c") #costos totales
p = model.addVars(3,6, lb=0,vtype='I', name="p") #produccion producto j en tiempo t
h = model.addVars(3,6, lb=0,vtype='I', name="h") # horas hombre trabajadas
e = model.addVars(3,6, lb=0,vtype='I', name="e") # horas hombre trabajadas EXTRAS
S = model.addVars(3,6, lb=0,vtype='I', name="S") # horas hombre trabajadas subcontratadas
contrato = model.addVars(6, lb=0,vtype='I',name="contrato")
despido = model.addVars(6, lb=0,vtype='I', name="despido")
o = model.addVars(6, vtype='B',name="o") #variable que decide el inv + -
o1 = model.addVars(6, vtype='B',name="o1")
o2 = model.addVars(6, vtype='B',name="o2")
o3 = model.addVars(6, vtype='B',name="o3")
H = model.addVars(6, vtype='I', lb=0, name='H') #maximo horas trabajables en horario normal
k = model.addVars(3,6, vtype='I', lb=0, name='k') #costo del tiempo total para producir producto en tiempo t
Tb = model.addVars(6, vtype='I', lb=0 , name='Tb') #staff en tiempo t personal de trabajo

# Decreasing value of coefficients
sa0 = model.addConstrs((h[j,t] + e[j,t] + S[j,t] == q[j]*p[j,t] for j in materiales for t in periodos), name='sa0')
sa1= model.addConstrs((k[j,t] == h[j,t]*s + e[j,t]*epsilon + S[j,t]*phi for j in materiales for t in periodos), name='sa1')
sa2 = model.addConstrs(  (c[j,t] == cb[j]*p[j,t] + k[j,t] for j in materiales for t in periodos), name='sa2')
sa3 = model.addConstrs((gp.quicksum(h[j,t] for j in materiales) <= H[t] for t in periodos), name='sa3')
sa4 = model.addConstrs((H[t] == Tb[t]*160 for t in periodos), name='sa4')
sa4_5 = model.addConstr((Tb[0] == 80 - despido[0] + contrato[0]), name='sa4.5')
sa5 = model.addConstrs((Tb[t] == Tb[t-1] - despido[t] + contrato[t] for t in periodos_2), name='sa5')
sa6_5 = model.addConstr((inv[0,0] == PM[0]*1000 ), name='sa6.3')
sa6_6 = model.addConstr((inv[1,0] == PM[1]*1000), name='sa6.4')
sa6_7 = model.addConstr((inv[2,0] == PM[2]*1000), name='sa6.5')
sa6 = model.addConstrs((inv[j,t] == inv[j,t-1] + p[j,t-1]- D[t-1]*PM[j] for j in materiales for t in periodos_2), name='sa6')
sa7 = model.addConstrs((Inv[t] == gp.quicksum(inv[j,t] for j in materiales) for t in periodos), name='sa7')
#sa8 = model.addConstrs((-o[t]*100000000 <= Inv[t]  for t in periodos), name='sa8')
sa8_1 = model.addConstrs((-o1[t]*100000000 <= inv[0,t]  for t in periodos), name='sa8.1')
sa8_2 = model.addConstrs((-o2[t]*100000000 <= inv[1,t]  for t in periodos), name='sa8.2')
sa8_3 = model.addConstrs((-o3[t]*100000000 <= inv[2,t]  for t in periodos), name='sa8.3')
sa9 = model.addConstr((Inv[5] >= 500), name='sa9')
sa10 = model.addConstrs((gp.quicksum(e[j,t] for j in materiales) <= 10*Tb[t] for t in periodos), name='sa10')
# Objective function
model.setObjective(gp.quicksum( c[j,t] + (1-o1[t])*alph*inv[0,t] + o1[t]*beta*inv[0,t] +(1-o1[t])*alph*inv[1,t] + o1[t]*beta*inv[1,t] + (1-o3[t])*alph*inv[2,t] + o3[t]*beta*inv[2,t] + contrato[t]*gamma + despido[t]*phi
                               for j in materiales
                               for t in periodos),
                   GRB.MINIMIZE)



# Verify model formulation

model.write('OptimizeConstraint.lp')

# Run optimization engine

modelo_final = model.optimize()

var_names = list()
var_values= list()

for var in model.getVars():
    var_names.append(str(var.varName))
    var_values.append(var.X)

# Write to csv
print(var_names)
print(var_values)
df = pd.DataFrame({'Var names': var_names, 'var values': var_values})

df.to_excel('case.xlsx',sheet_name='sheet1', index=False)