import sys
import numpy as np
import pandas as pd
if len(sys.argv)!=5:
    print("incorrect number of arguments")
    sys.exit()
input_file=sys.argv[1]
weights=sys.argv[2]
impacts=sys.argv[3]
output_file=sys.argv[4]
try:
    if input_file.endswith(".csv"):
        data = pd.read_csv(input_file)

    elif input_file.endswith(".xlsx"):
        data = pd.read_excel(input_file, engine="openpyxl")

        # auto convert Excel to CSV
        csv_name = input_file.replace(".xlsx", ".csv")
        data.to_csv(csv_name, index=False)

    else:
        print("Unsupported file format. Use CSV or XLSX.")
        sys.exit()

except FileNotFoundError:
    print("no file")
    sys.exit()
if data.shape[1]<3:
    print("insufficient columns")    
    sys.exit()
weights=list(map(float,weights.split(",")  ) )
impacts=impacts.split(",")
if len(weights)!=data.shape[1]-1 or len(impacts)!=data.shape[1]-1:
    print("length of weights or impacts is not equal to number of criteria")
    sys.exit()
paramter=data.iloc[:,1:]
for i in range(paramter.shape[1]):
    if impacts[i] not in ['+','-']:
        print("impacts should be either + or -")
        sys.exit()
if not paramter.applymap(lambda x: isinstance(x, (int, float))).all().all():
    print("non numeric values found")
    sys.exit()



criteria = data.iloc[:, 1:].astype(float).values


norm = np.sqrt((criteria ** 2).sum(axis=0))
normalized = criteria / norm

weights = np.array(weights)
weighted = normalized * weights


ideal_best = []
ideal_worst = []

for i in range(len(impacts)):
    if impacts[i] == "+":
        ideal_best.append(weighted[:, i].max())
        ideal_worst.append(weighted[:, i].min())
    else:
        ideal_best.append(weighted[:, i].min())
        ideal_worst.append(weighted[:, i].max())

ideal_best = np.array(ideal_best)
ideal_worst = np.array(ideal_worst)


dist_best = np.sqrt(((weighted - ideal_best) ** 2).sum(axis=1))
dist_worst = np.sqrt(((weighted - ideal_worst) ** 2).sum(axis=1))


topsis_score = dist_worst / (dist_best + dist_worst)


data["Topsis Score"] = topsis_score
data["Rank"] = data["Topsis Score"].rank(ascending=False, method="dense")


data.to_csv(output_file, index=False)

print("TOPSIS result saved to", output_file)
      
