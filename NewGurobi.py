from gurobipy import *
import os
import openpyxl



def ReadData(floderName, txtName):
    '''read the common data'''
    fileName = "D:\\Instance\\" + floderName + "\\SamePara.txt"

    with open(fileName) as f:
        readLine = f.readlines()
        for i in range(0, len(readLine)):
            if i == 1:
                jobNum, machineNum = map(lambda x: int(x), readLine[i].strip("\n").split())
            if i == 3:
                capacity = int(readLine[i].strip("\n"))
            if i == 5:
                jobSize = list(map(lambda x: int(x), readLine[i].strip("\n").split()))
            if i == 7:
                jobReadyTime = list(map(lambda x: int(x), readLine[i].strip("\n").split()))

    #read data from different scenarios
    scenarioName = "D:\\Instance\\" + floderName + "\\" + str(txtName) + ".txt"
    jobProcessTime = []
    with open(scenarioName) as s:
        scenarioLine = s.readlines()
        for j in range(0, len(scenarioLine)):
            if j >= 1:
                jobProcessTime.append(list(map(lambda x: int(x), scenarioLine[j].strip("\n").split())))

    batchNum = jobNum
    jobInfo = {}
    for i in range(jobNum):
        jobInfo[i] = [jobSize[i], jobReadyTime[i], jobProcessTime[i]]

    return jobNum, machineNum, batchNum, capacity, jobInfo


# read data
path = "D:\\2"
dirs = os.listdir(path)
for folderName in dirs:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['scenario index', 'GurobiResult', 'Gap(%)'])

    for txtName in range(20):
        jobNum, machineNum, batchNum, capacity, jobInfo = ReadData(folderName, txtName)
        jobIndex, jobSize, jobReadyTime, jobProcessTime = multidict(jobInfo)
        print(jobReadyTime)

        # set up model
        m = Model("MILP model")
        x = m.addVars(range(0, jobNum), range(0, batchNum), range(0, machineNum), vtype=GRB.BINARY, name="x")
        S = m.addVars(range(0, batchNum), range(0, machineNum), name = "S")
        P = m.addVars(range(0, batchNum), range(0, machineNum), name = "P")
        C_max = m.addVar(name = "C_max")

        
        m.update()
        # add objective
        m.setObjective(C_max, sense=GRB.MINIMIZE)
        # add constraints
        m.addConstrs(quicksum(x[j, b, m] for b in range(0, batchNum) for m in range(0, machineNum)) == 1
                     for j in range(0, jobNum))

        m.addConstrs(quicksum(x[j, b, m] * jobSize[j] for j in range(0, jobNum)) <= capacity
                     for b in range(0, batchNum) for m in range(0, machineNum))
        m.addConstrs(jobReadyTime[j] * x[j, b, m] <= S[b, m] for j in range(0, jobNum)
                     for b in range(0, batchNum) for m in range(0, machineNum))
        m.addConstrs(jobProcessTime[j][m] * x[j, b, m] <= P[b, m] for j in range(0, jobNum)
                     for b in range(0, batchNum) for m in range(0, machineNum))

        m.addConstrs(S[b, m] + P[b, m] <= S[b + 1, m] for b in range(0, batchNum - 1) for m in range(0, machineNum))

        m.addConstrs(P[b, m] + S[b, m] <= C_max for b in range(0, batchNum)
                     for m in range(0, machineNum))

        m.write("1.lp")
        m.Params.TimeLimit = 3600
        m.optimize()

        ResultDir = path + "\\" + folderName
        if m.status == GRB.OPTIMAL:
            ResultFile = 'Gurobi_' + str(txtName) + ".txt"
            f = open(ResultDir + "\\" + ResultFile, 'w+')
            f.write("RunTime is " + str(m.Runtime) + "\n")
            f.write("BestObj is " + str(m.objval) + "\n")
            f.write('gap=:' + str(m.mipgap) + "\n")
            f.write("-----------------optimal solution--------------------\n")
            for v in m.getVars():
                if v.x > 0:
                    f.write(str(v.varName) + "\t" + str(v.x) + "\n")
            ws.append([txtName, round(m.objval, 2), round(m.mipgap, 2), "Optimal"])

        else:
            ResultFile = 'Gurobi_' + str(txtName) + ".txt"
            f = open(ResultDir + "\\" + ResultFile, 'w+')
            f.write("RunTime is " + str(m.Runtime) + "\n")
            f.write("BestObj is " + str(m.objval) + "\n")
            f.write('gap=:' + str(m.mipgap) + "\n")
            f.write("-----------------feasible solution--------------------\n")
            for v in m.getVars():
                if v.x > 0:
                    f.write(str(v.varName) + "\t" + str(v.x) + "\n")
            ws.append([txtName, round(m.objval, 2), round(m.mipgap, 2), "Feasible"])
        f.close()

    wb.save(path + "\\" + folderName + "\\" + "GurobiResult.xlsx")






