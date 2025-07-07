import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from matplotlib.ticker import MaxNLocator

# Setup
if not os.path.exists('images'):
    os.makedirs('images')

def load_data():
    try:
        xls = pd.ExcelFile('student_performance_analysis.xlsx')
        sheets = {
            'basic_stats': xls.parse('Basic_Stats_By_Subject_Term'),
            'gender_stats': xls.parse('Subject_Perf_By_Gender'),
            'grade_dist': xls.parse('Grade_Distribution'),
            'term_trends': xls.parse('Subject_Trend'),
            'class_perf': xls.parse('Class_Subject_Perf'),
            'best_worst_subjects': xls.parse('Best_Worst_Subjects'),  # ✅ Fixed
            'subject_popularity': xls.parse('Subject_Popularity'),
            'top_bottom_performers': xls.parse('Top_Bottom_Performers'),
            'improved_declined': xls.parse('Improved_Declined'),
            'gender_performance': xls.parse('Gender_Performance')
        }
        return sheets
    except Exception as e:
        print("Error loading data:", e)
        return None

def create_db(data):
    conn = sqlite3.connect('student_performance.db')
    for name, df in data.items():
        df.to_sql(name.lower(), conn, if_exists='replace', index=False)
    conn.close()
    print("✅ Database created successfully")

def plot_best_worst_subjects(conn):
    df = pd.read_sql("SELECT * FROM best_worst_subjects", conn)

    plt.figure(figsize=(10, 6))
    colors = {'Top 3 Subjects': 'green', 'Bottom 3 Subjects': 'red'}
    bars = plt.barh(df['Subject'], df['Avg_Total_Score'], 
                   color=[colors[x] for x in df['Category']])
    plt.title('Best and Worst Performing Subjects', pad=20)
    plt.xlabel('Average Total Score')
    plt.ylabel('Subject')

    for bar in bars:
        width = bar.get_width()
        plt.text(width + 1, bar.get_y() + bar.get_height()/2, 
                 f'{width:.1f}', va='center', ha='left')

    plt.grid(axis='x', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.savefig('images/best_worst_subjects.png', dpi=300)
    plt.close()

# Add more plotting functions as needed, for example:
def plot_subject_popularity(conn):
    df = pd.read_sql("SELECT * FROM subject_popularity", conn)
    plt.figure(figsize=(10, 5))
    plt.bar(df['Subject'], df['Frequency'], color='purple')
    plt.xticks(rotation=45)
    plt.title('Subject Popularity')
    plt.xlabel('Subject')
    plt.ylabel('Enrollment')
    plt.tight_layout()
    plt.savefig('images/subject_popularity.png', dpi=300)
    plt.close()

def generate_all_visuals():
    conn = sqlite3.connect('student_performance.db')

    print("\n📊 Generating Visualizations...")
    plot_best_worst_subjects(conn)
    plot_subject_popularity(conn)

    # Add more function calls like:
    # plot_gender_performance(conn)
    # plot_top_bottom_performers(conn)

    conn.close()
    print("✅ Charts saved to /images")

def main():
    print("📘 STUDENT PERFORMANCE ANALYSIS")

    print("\n📥 Loading data...")
    data = load_data()
    if not data:
        return

    print("🛠️ Creating database...")
    create_db(data)

    generate_all_visuals()

    print("\n✅ All done! Check the /images folder for charts.")

if __name__ == "__main__":
    main()