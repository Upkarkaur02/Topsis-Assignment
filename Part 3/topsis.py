import pandas as pd
import numpy as np
import sys

def calculate_topsis(input_file, weights, impacts, output_file):
    # 1. Handle File Reading
    try:
        # Explicitly use openpyxl to prevent the ImportError seen in your logs
        if input_file.endswith('.xlsx'):
            df = pd.read_excel(input_file, engine='openpyxl')
        else:
            df = pd.read_csv(input_file)
    except Exception as e:
        raise Exception(f"Could not read file: {str(e)}")

    # 2. Validation: Ensure we have enough columns
    if df.shape[1] < 3:
        raise Exception("Input file must have at least 3 columns (ID/Name + 2 or more numeric columns)")

    # Extract numeric data (skipping the first column)
    data = df.iloc[:, 1:].values
    
    # 3. Parse and Clean Weights/Impacts
    try:
        # .strip() is crucial to fix the "Impacts must be + or -" error from your screenshot
        weights_list = [float(w.strip()) for w in weights.split(',')]
        impact_list = [i.strip() for i in impacts.split(',')]
    except ValueError:
        raise Exception("Weights must be numeric and comma-separated.")

    # 4. Validation: Check if counts match columns
    num_cols = data.shape[1]
    if len(weights_list) != num_cols or len(impact_list) != num_cols:
        raise Exception(f"Number of weights/impacts ({len(weights_list)}) must match numeric columns ({num_cols})")

    if not all(i in ['+', '-'] for i in impact_list):
        # This addresses the error: "Error: Impacts must be either '+' or '-'"
        raise Exception("Impacts must be either '+' or '-'")

    # 5. TOPSIS Algorithm
    # Vector Normalization
    norm_data = data / np.sqrt((data**2).sum(axis=0))
    
    # Weighted Normalized Matrix
    weighted_data = norm_data * weights_list

    # Ideal Best and Worst
    ideal_best = []
    ideal_worst = []

    for i in range(len(weights_list)):
        col_values = weighted_data[:, i]
        if impact_list[i] == '+':
            ideal_best.append(np.max(col_values))
            ideal_worst.append(np.min(col_values))
        else:
            ideal_best.append(np.min(col_values))
            ideal_worst.append(np.max(col_values))

    # Calculate Euclidean Distances
    s_best = np.sqrt(np.sum((weighted_data - ideal_best)**2, axis=1))
    s_worst = np.sqrt(np.sum((weighted_data - ideal_worst)**2, axis=1))

    # Performance Score (P-Score)
    score = s_worst / (s_best + s_worst)

    # 6. Append Results and Save
    df['Topsis Score'] = score
    df['Rank'] = score.rank(ascending=False).astype(int)

    # Save using openpyxl engine to ensure compatibility
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    return output_file

if __name__ == "__main__":
    # Optional: Allow command line usage as seen in your terminal
    if len(sys.argv) == 5:
        calculate_topsis(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
        print("TOPSIS completed successfully.")