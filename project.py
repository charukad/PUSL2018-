import pandas as pd
import numpy as np
import random
import matplotlib.pyplot as plt

throws = []

def throw_dart():
    x = random.uniform(-50, 50)
    y = random.uniform(-50, 50)
    dec = incircle((x, y))
    throws.append((x, y, dec))

def incircle(throw):
    x, y = throw
    distance = (x ** 2 + y ** 2) ** 0.5
    if distance <= 50:
        return 1
    else:
        return 0

def plot_dart_throws(df):
    plt.figure(figsize=(10, 6))
    inside_circle = df[df['InsideCircle'] == 1]
    outside_circle = df[df['InsideCircle'] == 0]

    plt.scatter(outside_circle['x'], outside_circle['y'], color='red', label='Outside Circle')
    plt.scatter(inside_circle['x'], inside_circle['y'], color='blue', label='Inside Circle')

    circle = plt.Circle((0, 0), 50, color='green', fill=False, linestyle='--', label='Circle')
    plt.gca().add_patch(circle)

    plt.xlabel('X')
    plt.ylabel('Y')
    plt.title('Dart Throws: Inside and Outside Circle')
    plt.legend()
    plt.show()


count = 100
for i in range(count): 
    throw_dart()

df = pd.DataFrame(throws, columns=['x', 'y', 'InsideCircle'])
df['pi'] = (4 * np.cumsum(df['InsideCircle'])) / (df.index + 1)
df.to_csv('dart_throws.csv', index=False)

excel_file = "dart_simulation_results.xlsx"
generate_excel_with_mean(df, excel_file)

plt.figure(figsize=(10, 6))
plt.plot(df['pi'], label='Estimate of pi')
plt.axhline(y=3.141592653589793, color='r', linestyle='--', label='Actual value of pi')
plt.xlabel('Number of Throws')
plt.ylabel('Estimated Value of pi')
plt.title('Monte Carlo Simulation: Dart Throwing')
plt.legend()
plt.show()

plot_dart_throws(df)
def generate_excel_with_mean(df, excel_file):
    try:
        existing_data = pd.read_excel(excel_file, sheet_name='Summary')
        execution_count = existing_data['Execution Count'].max() + 1
    except FileNotFoundError:
        execution_count = 1

    mean_pi = df['pi'].mean()
    dart_count = df.index[-1] + 1  
    summary_df = pd.DataFrame({
        'Execution Count': [execution_count],
        'Dart Count': [dart_count],
        'Mean Pi': [mean_pi]
    })

    if execution_count == 1:
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
    else:
        updated_data = pd.concat([existing_data, summary_df], ignore_index=True)
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            updated_data.to_excel(writer, sheet_name='Summary', index=False)
