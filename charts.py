import tkinter as tk
from tkinter import messagebox
import threading
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import numpy as np
import os
import tempfile

# Define the path for the Excel file
excel_file = 'output_data.xlsx'

# Load the data from the Excel file
df = pd.read_excel(excel_file, engine="openpyxl", sheet_name=None)

# Replace 'Not found' with NaN and extract the DataFrames
data_frames = {
    'Valuation Data': df['Valuation Data'].replace('Not found', np.nan),
    'Margin Data': df['Margin Data'].replace('Not found', np.nan),
    'Management Data': df['Management Data'].replace('Not found', np.nan),
    'Growth Data': df['Growth Data'].replace('Not found', np.nan),
    'Financial Strength Data': df['Financial Strength Data'].replace('Not found', np.nan)
}

# Define cell positions for charts
cell_positions = {
    'Valuation Data': ['A10', 'K10', 'U10', 'A31', 'K31', 'U31'],
    'Margin Data': ['A10', 'K10', 'U10', 'A31', 'K31', 'U31'],
    'Management Data': ['A10', 'K10', 'U10', 'A31', 'K31', 'U31'],
    'Growth Data': ['A10', 'K10', 'U10', 'A31', 'K31', 'U31'],
    'Financial Strength Data': ['A10', 'K10', 'U10', 'A31', 'K31', 'U31']
}

# Define colours for charts
colors = ['#6a1b9a', '#7b1fa2', '#8e24aa', '#9c27b0', '#ab47bc', '#ba68c8']

# Apply a figure style
plt.style.use('fivethirtyeight')

# Function to create, sort, and save aesthetically enhanced charts
def create_and_save_charts(excel_file, data_frames, cell_positions, colors, image_scale=0.6):
    for sheet_name, data_frame in data_frames.items():
        book = load_workbook(excel_file)
        ws = book[sheet_name]

        with tempfile.TemporaryDirectory() as temp_dir:
            for i, column in enumerate(data_frame.columns[1:]):  # Skip the symbol column
                print(f"Plotting {column}")
                # Sorting the values
                sorted_df = data_frame.sort_values(by=column, ascending=False)
                values = pd.to_numeric(sorted_df[column], errors='coerce').dropna()
                valid_symbols = sorted_df.iloc[:, 0][values.index]
                fig, ax = plt.subplots(figsize=(14, 7))
                bars = ax.bar(valid_symbols, values, color=[colors[i % len(colors)] for i in range(len(valid_symbols))])

                # Adding value labels on top of each bar
                for bar in bars:
                    height = bar.get_height()
                    label_x_pos = bar.get_x() + bar.get_width() / 2
                    ax.text(label_x_pos, height, s=f'{height:.2f}', ha='center', va='bottom', fontsize=12)

                # Set the chart title and labels
                ax.set_title(f'{column} - {sheet_name}', fontsize=20, weight='bold', color='#333333')
                ax.set_xlabel('Symbol', fontsize=14, weight='bold')
                ax.set_ylabel(column, fontsize=14, weight='bold')

                # Customize the grid
                ax.grid(True, color='#CCCCCC', linestyle='--', linewidth=0.5)
                ax.set_axisbelow(True)

                # Set the background color
                ax.set_facecolor('#FAFAFA')

                # Rotate x-ticks
                ax.tick_params(axis='x', rotation=45)

                # Tight layout
                plt.tight_layout()

                # Save the figure
                chart_image_path = os.path.join(temp_dir, f'chart_{sheet_name}_{i}.png')
                fig.savefig(chart_image_path, transparent=False)
                plt.close(fig)

                # Insert the image into Excel
                img = OpenpyxlImage(chart_image_path)
                img.width, img.height = img.width * image_scale, img.height * image_scale
                if i < len(cell_positions[sheet_name]):
                    ws.add_image(img, cell_positions[sheet_name][i])

            # Save the workbook
            book.save(excel_file)
            print(f"Charts saved in {sheet_name} sheet.")

class ChartApp(tk.Tk):
    def __init__(self, excel_file, data_frames, cell_positions, colors):
        super().__init__()

        self.title("Charts Generation")
        self.excel_file = excel_file
        self.data_frames = data_frames
        self.cell_positions = cell_positions
        self.colors = colors
        self.label = tk.Label(self, text="Generating charts. Please wait...")
        self.label.pack(pady=10)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.thread = threading.Thread(
            target=self.threaded_chart_creation,
            args=(0.5,),
            daemon=True
        )
        self.thread.start()

    def threaded_chart_creation(self, image_scale):
        try:
            create_and_save_charts(self.excel_file, self.data_frames, self.cell_positions, self.colors, image_scale)
            messagebox.showinfo("Success", "Charts have been generated successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.on_finished()

    def on_finished(self):
        self.label['text'] = "Chart generation completed!"
        self.update_idletasks()
        self.after(2000, self.destroy)  # Close the window after a delay

    def on_closing(self):
        # This can be used to stop the chart generation thread if necessary
        # self.thread.stop()
        if messagebox.askokcancel("Quit", "Do you want to quit the application?"):
            self.destroy()

#Generate GUI
def generate_charts_gui(excel_file, data_frames, cell_positions, colors):
    app = ChartApp(excel_file, data_frames, cell_positions, colors)
    app.mainloop()

generate_charts_gui(excel_file, data_frames, cell_positions, colors)
