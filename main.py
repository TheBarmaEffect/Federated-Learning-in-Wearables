import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from mpl_toolkits.mplot3d import Axes3D
from fpdf import FPDF
import os

# Step 1: Load the Dataset
excel_file_path = "C:/Users/shane/OneDrive/Documents/synthetic_wearable_weekly_data_with_timestamps (1).xlsx"
dataset = pd.read_excel(excel_file_path)

# Data Preprocessing
dataset['ExercisingThisWeek'] = dataset['ExercisingThisWeek'].astype(int)
X = dataset.drop(['ExercisingThisWeek', 'Timestamp'], axis=1)
X = pd.get_dummies(X, columns=['ActivityType'], drop_first=True)
y = dataset['ExercisingThisWeek']

# Create individual histograms with deeper explanations and save them as images
def create_histogram_with_deeper_explanation(data, title, x_label, explanation, image_filename):
    plt.figure(figsize=(8, 6))
    sns.histplot(data, bins=20, kde=True)
    plt.title(title)
    plt.xlabel(x_label)
    plt.ylabel("Frequency")

    # Add a text box with the deeper explanation
    plt.gcf().text(0.12, 0.5, explanation, fontsize=10, bbox=dict(facecolor='white', alpha=0.7))

    plt.savefig(image_filename)
    plt.close()

# Create individual 3D plots and save them as images
def create_3d_plot(data, title, x_label, y_label, z_label, image_filename):
    fig = plt.figure(figsize=(10, 8))
    ax = fig.add_subplot(111, projection='3d')

    x = data[x_label]
    y = data[y_label]
    z = data[z_label]

    ax.scatter(x, y, z, c='r', marker='o')

    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    ax.set_zlabel(z_label)

    ax.set_title(title)

    plt.savefig(image_filename)
    plt.close()

# Create a folder to store histogram and 3D plot images
output_folder = "images"
os.makedirs(output_folder, exist_ok=True)

# Create 2D histograms and 3D plots for individual columns
create_histogram_with_deeper_explanation(
    dataset['TotalKmWalked'],
    "Total Kilometers Walked",
    "Total Kilometers Walked",
    "This histogram shows the distribution of total kilometers walked, which indicates your physical activity level.",
    os.path.join(output_folder, "TotalKmWalked_Histogram.png")
)

create_3d_plot(
    dataset[['TotalKmWalked', 'AvgRestingHeartRate', 'AvgRestfulSleep']],
    "3D Health Plot 1",
    'TotalKmWalked', 'AvgRestingHeartRate', 'AvgRestfulSleep',
    os.path.join(output_folder, "3D_Health_Plot_1.png")
)

# Define explanations and suggestions for each column
column_info = {
    'TotalKmWalked': {
        'title': "Total Kilometers Walked Analysis:",
        'x_label': "Total Kilometers Walked",
        'explanation': "The histogram below shows the distribution of total kilometers walked in the past week. "
                       "This metric indicates your physical activity level.",
        'suggestions': "You have walked a total of X kilometers in the past week, which suggests a moderately active lifestyle. "
                       "Consider increasing your daily steps to improve your overall fitness. Regular exercise can lead to a healthier lifestyle."
    },
    'AvgRestingHeartRate': {
        'title': "Average Resting Heart Rate Analysis:",
        'x_label': "Average Resting Heart Rate (bpm)",
        'explanation': "The histogram displays your resting heart rate distribution, measured in beats per minute (bpm).",
        'suggestions': "Your average resting heart rate is X bpm, which is within the healthy range. Maintain an active lifestyle to keep your heart healthy."
    },
    # Add more columns here...
}

# Define 3D histograms for the specified columns
additional_3d_histograms = {
    'CaloriesBurned': {
        'title': "Calories Burned Analysis:",
        'x_label': "Calories Burned",
        'y_label': "AvgRestingHeartRate",
        'z_label': "AvgRestfulSleep",
        'explanation': "Explanation for Calories Burned.",
        'suggestions': "Suggestions for Calories Burned."
    },
    'TotalActiveMinutes': {
        'title': "Total Active Minutes Analysis:",
        'x_label': "TotalActiveMinutes",
        'y_label': "AvgRestingHeartRate",
        'z_label': "AvgRestfulSleep",
        'explanation': "Explanation for Total Active Minutes.",
        'suggestions': "Suggestions for Total Active Minutes."
    },
    'AvgHrsWith250PlusSteps': {
        'title': "Average Hrs With 250+ Steps Analysis:",
        'x_label': "AvgHrsWith250PlusSteps",
        'y_label': "AvgRestingHeartRate",
        'z_label': "AvgRestfulSleep",
        'explanation': "Explanation for Average Hrs With 250+ Steps.",
        'suggestions': "Suggestions for Average Hrs With 250+ Steps."
    },
    'ActivityHeartRate': {
        'title': "Activity Heart Rate Analysis:",
        'x_label': "ActivityHeartRate",
        'y_label': "AvgRestingHeartRate",
        'z_label': "AvgRestfulSleep",
        'explanation': "Explanation for Activity Heart Rate.",
        'suggestions': "Suggestions for Activity Heart Rate."
    },
    'BodyWeight': {
        'title': "Body Weight Analysis:",
        'x_label': "BodyWeight",
        'y_label': "AvgRestingHeartRate",
        'z_label': "AvgRestfulSleep",
        'explanation': "Explanation for Body Weight.",
        'suggestions': "Suggestions for Body Weight."
    }
}

# Create subplots for additional 3D histograms and save them as images
fig = plt.figure(figsize=(15, 10))
for i, (column_name, column_data) in enumerate(additional_3d_histograms.items()):
    ax = fig.add_subplot(2, 3, i + 1, projection='3d')  # 2 rows, 3 columns of subplots
    create_3d_plot(
        dataset[[column_name, 'AvgRestingHeartRate', 'AvgRestfulSleep']],
        column_data['title'],
        column_name, 'AvgRestingHeartRate', 'AvgRestfulSleep',
        os.path.join(output_folder, f"{column_name}_3D_Plot.png")
    )
    ax.set_title(column_data['title'])

# Adjust the layout
plt.tight_layout()

# Save the figure
plt.savefig(os.path.join(output_folder, "additional_3d_histograms.png"))

class DetailedPDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Your Comprehensive Health Analysis Report', 0, 1, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(10)

    def chapter_body(self, body):
        self.set_font('Arial', '', 12)
        self.multi_cell(0, 10, body)
        self.ln()

    def add_image(self, image_path):
        self.image(image_path, x=10, w=190)

    def explanation_and_suggestions(self, title, explanation, suggestions, image_path):
        self.chapter_title(title)
        self.add_image(image_path)
        self.chapter_body(explanation)
        self.chapter_body(suggestions)

# Create a PDF instance for the detailed report
detailed_pdf = DetailedPDF()
detailed_pdf.add_page()

# Title
detailed_pdf.set_title("Your Comprehensive Health Analysis Report")
detailed_pdf.set_author("Your Health Analysis Team")

# Add content to the detailed PDF report

# Introduction
detailed_pdf.chapter_title("Introduction:")
detailed_pdf.chapter_body(
    "Welcome to your comprehensive health analysis report. This report provides an in-depth analysis of your health and fitness data, "
    "along with suggestions to improve your overall well-being. Below, you'll find detailed explanations and insights "
    "about your health metrics and what they mean for your lifestyle."
)

# Add explanations and suggestions for the 10 columns
for column_name, column_data in column_info.items():
    image_path = os.path.join(output_folder, f"{column_name}_Histogram.png")
    detailed_pdf.explanation_and_suggestions(
        column_data['title'],
        column_data['explanation'],
        column_data['suggestions'],
        image_path
    )

# Add explanations and suggestions for 6 additional 3D histograms
for column_name, column_data in additional_3d_histograms.items():
    image_path = os.path.join(output_folder, f"{column_name}_3D_Plot.png")
    detailed_pdf.explanation_and_suggestions(
        column_data['title'],
        column_data['explanation'],
        column_data['suggestions'],
        image_path
    )

# Save the detailed PDF report
detailed_pdf_output_filename = "detailed_health_analysis_report.pdf"
detailed_pdf.output(detailed_pdf_output_filename)

# Display the detailed PDF report
import subprocess

subprocess.Popen([detailed_pdf_output_filename], shell=True)
