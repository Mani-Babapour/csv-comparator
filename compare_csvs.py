import os
import pandas as pd

# Set your main folder path here
main_folder = 'path_to_your_folder'

# List to collect the results
results = []

# Go through folders
for student_folder in os.listdir(main_folder):
    student_path = os.path.join(main_folder, student_folder)
    if os.path.isdir(student_path):
        # Get all CSV files
        csv_files = [f for f in os.listdir(student_path) if f.endswith('.csv')]
        
        if len(csv_files) == 2:
            # Read both CSV files
            file1_path = os.path.join(student_path, csv_files[0])
            file2_path = os.path.join(student_path, csv_files[1])
            
            try:
                df1 = pd.read_csv(file1_path, encoding='utf-8')
                df2 = pd.read_csv(file2_path, encoding='utf-8')
                
                # Compare shapes first
                if df1.shape != df2.shape:
                    explanation = f"Difference found: different shape {df1.shape} vs {df2.shape}"
                else:
                    # Compare the actual content
                    if df1.equals(df2):
                        explanation = "No difference"
                    else:
                        # Find where the difference is
                        diffs = (df1 != df2)
                        diff_cells = diffs.sum().sum()
                        explanation = f"Difference found: {diff_cells} different cells"
            
            except Exception as e:
                explanation = f"Error reading files: {str(e)}"
            
            results.append({
                'Student Name': student_folder,
                'Explanation': explanation
            })
        else:
            results.append({
                'Student Name': student_folder,
                'Explanation': "Error: Expected 2 CSV files, found " + str(len(csv_files))
            })

# Save results to Excel
output_df = pd.DataFrame(results)
output_df.to_excel('comparison_result.xlsx', index=False)

print("Done! Result saved as 'comparison_result.xlsx'")
