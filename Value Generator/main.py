import random
import xlsxwriter
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import messagebox

# Set appearance mode and color theme
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def get_vibration_range(mode):
    """Return min and max vibration values based on selected mode"""
    ranges = {
        "Ideal": (0.5, 2.5),    # Very low vibration range
        "Normal": (1.5, 8.0),   # Typical operating range
        "Harsh": (8.0, 25.0),   # Critical/Warning range
        "Random": (0.5, 25.0)   # Full range
    }
    return ranges.get(mode)

def configure_and_generate():
    try:
        time_interval = int(entry_interval.get())
        data_points = int(entry_points.get())
        num_channels = int(entry_channels.get())
        mode = mode_selector.get()
        
        # Input validation
        if time_interval <= 0 or data_points <= 0 or num_channels <= 0:
            raise ValueError("Please enter positive integers!")
        if num_channels > 50:
            raise ValueError("Number of channels cannot exceed 50!")
            
    except ValueError as e:
        error_label.configure(text=str(e))
        return

    generate_excel_file(time_interval, data_points, num_channels, mode)
    error_label.configure(text="")

def generate_excel_file(time_interval, data_points, num_channels, mode):
    output_file = "Generated_Data.xlsx"
    min_value, max_value = get_vibration_range(mode)
    
    # Create column headers with channel numbers
    columns = ["Date & Time"]
    for i in range(1, num_channels + 1):
        columns.append(f"Channel {i}")

    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    # Set column widths
    worksheet.set_column('A:A', 18)  # Date & Time column
    worksheet.set_column(1, num_channels, 12)  # Data columns

    # Create formats
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1
    })

    units_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1
    })

    data_format = workbook.add_format({
        'align': 'center',
        'border': 1,
        'num_format': '0.0'
    })

    date_format = workbook.add_format({
        'align': 'center',
        'border': 1,
        'num_format': 'dd-mmm-yyyy hh:mm:ss'
    })

    # Write headers
    for col, header in enumerate(columns):
        worksheet.write(1, col, header, header_format)
        if col > 0:  # Skip Date & Time column
            worksheet.write(0, col, "mm/s RMS", units_format)

    # Generate and write data starting from current system time
    start_time = datetime.now()
    for row in range(data_points):
        current_time = start_time + timedelta(seconds=row * time_interval)
        # Write timestamp
        worksheet.write_datetime(row + 2, 0, current_time, date_format)
        
        # Write RMS values
        for col in range(1, len(columns)):
            if mode == "Ideal":
                # For ideal mode, generate more stable values with less variation
                base_value = random.uniform(min_value, (min_value + max_value) / 2)
                rms_value = round(base_value + random.uniform(-0.2, 0.2), 1)
            else:
                rms_value = round(random.uniform(min_value, max_value), 1)
            worksheet.write(row + 2, col, rms_value, data_format)

    workbook.close()
    messagebox.showinfo("Success", f"Generated Excel file ({mode} mode): {output_file}")

# --- CustomTkinter GUI ---
root = ctk.CTk()
root.title("Vibration Data Generator")

# Configure grid layout
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# Create frame for inputs
input_frame = ctk.CTkFrame(root)
input_frame.grid(row=0, column=0, columnspan=2, padx=20, pady=20, sticky="nsew")

# Mode Selector
label_mode = ctk.CTkLabel(input_frame, text="Generation Mode:")
label_mode.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

mode_selector = ctk.CTkOptionMenu(
    input_frame,
    values=["Ideal", "Normal", "Harsh", "Random"],
    width=200
)
mode_selector.grid(row=0, column=1, padx=20, pady=(20, 10), sticky="ew")
mode_selector.set("Normal")  # Default value

# Time Interval Input
label_interval = ctk.CTkLabel(input_frame, text="Time Interval (seconds):")
label_interval.grid(row=1, column=0, padx=20, pady=10, sticky="w")

entry_interval = ctk.CTkEntry(input_frame)
entry_interval.grid(row=1, column=1, padx=20, pady=10, sticky="ew")
entry_interval.insert(0, "10")  # Default value

# Number of Channels Input
label_channels = ctk.CTkLabel(input_frame, text="Number of Channels:")
label_channels.grid(row=2, column=0, padx=20, pady=10, sticky="w")

entry_channels = ctk.CTkEntry(input_frame)
entry_channels.grid(row=2, column=1, padx=20, pady=10, sticky="ew")
entry_channels.insert(0, "16")  # Default value

# Number of Data Points Input
label_points = ctk.CTkLabel(input_frame, text="Number of Data Points:")
label_points.grid(row=3, column=0, padx=20, pady=10, sticky="w")

entry_points = ctk.CTkEntry(input_frame)
entry_points.grid(row=3, column=1, padx=20, pady=10, sticky="ew")
entry_points.insert(0, "100")  # Default value

# Generate Button
generate_button = ctk.CTkButton(
    root, 
    text="Generate Excel File", 
    command=configure_and_generate,
    height=40
)
generate_button.grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 10), sticky="ew")

# Error Label
error_label = ctk.CTkLabel(root, text="", text_color="red")
error_label.grid(row=2, column=0, columnspan=2, padx=20, pady=(0, 20))

# Configure frame grid
input_frame.grid_columnconfigure(1, weight=1)

# Run the main event loop
root.mainloop()