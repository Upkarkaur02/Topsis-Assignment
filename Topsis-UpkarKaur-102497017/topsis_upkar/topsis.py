import sys
import pandas as pd
import numpy as np
import os

def topsis(input_file, weights, impacts, output_file):

    # -------- File Handling --------
    if not os.path.exists(input_file):
        print("Error: File not found")
        sys.exit()

    # Read Excel or CSV
    try:
        if input_file.endswith('.xlsx'):
            df = pd.read_excel(input_file)
        elif input_file.endswith('.csv'):
            df = pd.read_csv(input_file)
        else:
            print("Error: Only .csv or .xlsx files are supported")
            sys.exit()
    except Exception as e:
        print("Error reading file:", e)
        sys.exit()

    # -------- Column Check --------
    if df.shape[1] < 3:
        print("Error: Input file must contain at least 3 columns")
        sys.exit()

    # Extract numeric data (2nd to last column)
    data = df.iloc[:, 1:]

    # -------- Numeric Check --------
    try:
        data = data.apply(pd.to_numeric)
    except:
        print("Error: All columns except first must contain numeric values")
        sys.exit()

    # -------- Weights & Impacts --------
    weights = weights.split(',')
    impacts = impacts.split(',')

    if len(weights) != data.shape[1] or len(impacts) != data.shape[1]:
        print("Error: Number of weights and impacts must match number of criteria columns")
        sys.exit()

    try:
        weights = np.array(weights, dtype=float)
    except:
        print("Error: Weights must be numeric")
        sys.exit()

    for impact in impacts:
        if impact not in ['+', '-']:
            print("Error: Impacts must be either '+' or '-'")
            sys.exit()

    # -------- Step 1: Normalization --------
    normalized = data / np.sqrt((data**2).sum(axis=0))

    # -------- Step 2: Weighted Matrix --------
    weighted = normalized * weights

    # -------- Step 3: Ideal Best & Worst --------
    ideal_best = []
    ideal_worst = []

    for i in range(len(weights)):
        if impacts[i] == '+':
            ideal_best.append(weighted.iloc[:, i].max())
            ideal_worst.append(weighted.iloc[:, i].min())
        else:
            ideal_best.append(weighted.iloc[:, i].min())
            ideal_worst.append(weighted.iloc[:, i].max())

    ideal_best = np.array(ideal_best)
    ideal_worst = np.array(ideal_worst)

    # -------- Step 4: Distance --------
    dist_best = np.sqrt(((weighted - ideal_best)**2).sum(axis=1))
    dist_worst = np.sqrt(((weighted - ideal_worst)**2).sum(axis=1))

    # -------- Step 5: Score --------
    score = dist_worst / (dist_best + dist_worst)

    # -------- Step 6: Ranking --------
    df['Topsis Score'] = score
    df['Rank'] = score.rank(ascending=False)

    # -------- Save Output --------
    if output_file.endswith('.xlsx'):
        df.to_excel(output_file, index=False)
    else:
        df.to_csv(output_file, index=False)

    print("TOPSIS completed successfully.")
    print("Output saved to:", output_file)


if __name__ == "__main__":

    if len(sys.argv) != 5:
        print("Usage: python topsis.py <InputFile> <Weights> <Impacts> <OutputFile>")
        sys.exit()

    input_file = sys.argv[1]
    weights = sys.argv[2]
    impacts = sys.argv[3]
    output_file = sys.argv[4]

    topsis(input_file, weights, impacts, output_file)
