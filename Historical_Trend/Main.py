import pandas as pd
import customtkinter as ctk
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.widgets import Cursor

# Function to load and clean the data
def load_data():
    global data
    file_path = 'Historical_Trend\Demo_Data.xlsx'   #your Excel file path
    data = pd.read_excel(file_path, sheet_name='Sheet1')
    
    # Clean the data
    data.columns = data.iloc[0]  # Set the first row as column headers
    data = data[1:]  # Skip the header row used for column names
    data.rename(columns={data.columns[0]: "Date & Time"}, inplace=True)  # Rename the first column
    data["Date & Time"] = pd.to_datetime(data["Date & Time"])  # Convert to datetime
    data.set_index("Date & Time", inplace=True)
    data.dropna(axis=1, how="all", inplace=True)  # Drop columns with all NaN values
    data = data.apply(pd.to_numeric, errors='coerce')  # Convert data to numeric

# Initialize CustomTkinter
ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue"

# Main application window
app = ctk.CTk()
app.title("Historical Trend")
app.geometry("1200x800")

# Load the initial data
load_data()

# Header
header = ctk.CTkLabel(app, text="Historical Trend", font=("Arial", 20, "bold"))
header.pack(pady=10)

# Frame for dropdown, graph, and refresh button
main_frame = ctk.CTkFrame(app, corner_radius=10)
main_frame.pack(fill="both", expand=True, padx=20, pady=20)

# Dropdown menu for selecting channels
selected_channel = ctk.StringVar(value="All Channels")
dropdown = ctk.CTkComboBox(
    main_frame,
    values=["All Channels"] + list(data.columns),
    variable=selected_channel,
    width=300,
    command=lambda _: update_graph(),
)
dropdown.pack(pady=10)

# Refresh button
def refresh_data():
    load_data()  # Reload the data from the Excel file
    dropdown.configure(values=["All Channels"] + list(data.columns))  # Update dropdown values
    update_graph()  # Redraw the graph

refresh_button = ctk.CTkButton(main_frame, text="Refresh Data", command=refresh_data, corner_radius=8)
refresh_button.pack(pady=10)

# Frame for the graph
graph_frame = ctk.CTkFrame(main_frame, corner_radius=10)
graph_frame.pack(fill="both", expand=True, padx=10, pady=10)

# Create a Matplotlib figure
fig = Figure(figsize=(10, 5), dpi=100)
ax = fig.add_subplot(111)

# Function to update the graph based on the selected channel
def update_graph():
    ax.clear()
    choice = selected_channel.get()

    if choice == "All Channels":
        for column in data.columns:
            ax.plot(data.index, data[column], label=column)
    else:
        ax.plot(data.index, data[choice], label=choice)

    ax.set_title("Historical Trend Graph", fontsize=14)
    ax.set_xlabel("Date & Time", fontsize=12)
    ax.set_ylabel("Measurements (mm/s RMS)", fontsize=12)
    ax.legend(title="Channels", loc="upper left")
    ax.grid(True)
    canvas.draw()

# Embed the Matplotlib figure into CustomTkinter
canvas = FigureCanvasTkAgg(fig, master=graph_frame)
canvas.draw()
canvas.get_tk_widget().pack(fill="both", expand=True)

# Cursor for hovering over points
cursor = Cursor(ax, useblit=True, color='red', linewidth=1)

# Initial graph update
update_graph()

# Footer with quit button
footer_frame = ctk.CTkFrame(app, corner_radius=10)
footer_frame.pack(fill="x", padx=20, pady=10)

quit_button = ctk.CTkButton(
    footer_frame, text="Quit", command=app.quit, corner_radius=8
)
quit_button.pack(side="right", padx=10, pady=10)

# Run the application
app.mainloop()