import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import mplcursors

# Initialize CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Main application window
app = ctk.CTk()
app.title("Interactive Graph Application")
app.geometry("1200x800")

# Global variables
data = None
cursor = None  # To track the cursor instance

# Validate and load the data
def validate_and_load(file_path):
    global data
    required_columns = ["Date & Time"]
    try:
        temp_data = pd.read_excel(file_path, sheet_name='Sheet1')
        temp_data.columns = temp_data.iloc[0]
        temp_data = temp_data[1:]
        if not all(col in temp_data.columns for col in required_columns):
            raise ValueError("Missing required columns: 'Date & Time'")
        temp_data.rename(columns={temp_data.columns[0]: "Date & Time"}, inplace=True)
        temp_data["Date & Time"] = pd.to_datetime(temp_data["Date & Time"])
        temp_data.set_index("Date & Time", inplace=True)
        temp_data.dropna(axis=1, how="all", inplace=True)
        temp_data = temp_data.apply(pd.to_numeric, errors='coerce')
        data = temp_data
        dropdown.configure(values=["All Channels"] + list(data.columns))
        update_graph()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load data: {e}")

# Select a file
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")], title="Select an Excel File")
    if file_path:
        validate_and_load(file_path)

# Update the graph
def update_graph():
    global cursor
    ax.clear()
    if data is not None:
        # Clear any existing cursor instance
        if cursor is not None:
            cursor.remove()
            cursor = None
        
        choice = selected_channel.get()
        if choice == "All Channels":
            for column in data.columns:
                ax.plot(data.index, data[column], label=column)
        else:
            ax.plot(data.index, data[choice], label=choice)
        
        ax.set_title("Interactive Graph", fontsize=14)
        ax.set_xlabel("Date & Time", fontsize=12)
        ax.set_ylabel("Measurements", fontsize=12)
        ax.legend(title="Channels", loc="upper left")
        ax.grid(True)
        
        # Add a new cursor instance
        cursor = mplcursors.cursor(ax, hover=True)
        canvas.draw()

# Header
header = ctk.CTkLabel(app, text="Interactive Graph Application", font=("Arial", 20, "bold"))
header.pack(pady=10)

# Frame for dropdown, graph, and file selection
main_frame = ctk.CTkFrame(app, corner_radius=10)
main_frame.pack(fill="both", expand=True, padx=20, pady=20)

# File selection button
file_button = ctk.CTkButton(main_frame, text="Select Excel File", command=select_file, corner_radius=8)
file_button.pack(pady=10)

# Dropdown for channels
selected_channel = ctk.StringVar(value="All Channels")
dropdown = ctk.CTkComboBox(main_frame, values=["All Channels"], variable=selected_channel, width=300, command=lambda _: update_graph())
dropdown.pack(pady=10)

# Frame for the graph
graph_frame = ctk.CTkFrame(main_frame, corner_radius=10)
graph_frame.pack(fill="both", expand=True, padx=10, pady=10)

# Matplotlib figure
fig = Figure(figsize=(10, 5), dpi=100)
ax = fig.add_subplot(111)
canvas = FigureCanvasTkAgg(fig, master=graph_frame)
canvas.draw()
canvas.get_tk_widget().pack(fill="both", expand=True)

# Footer with quit button
footer_frame = ctk.CTkFrame(app, corner_radius=10)
footer_frame.pack(fill="x", padx=20, pady=10)

quit_button = ctk.CTkButton(footer_frame, text="Quit", command=app.quit, corner_radius=8)
quit_button.pack(side="right", padx=10, pady=10)

# Run the application
app.mainloop()
