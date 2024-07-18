import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pymol import cmd
import openpyxl

# Function to split and export docked poses
def split_and_export_docked_poses(input_file, output_prefix, output_dir):
    # Load the multi-state object from the input file
    cmd.load(input_file, "docked")

    # Count the number of states in the "docked" object
    num_states = cmd.count_states("docked")

    # Loop through each state, create a new object, and save it
    for i in range(1, num_states + 1):
        # Create a name for the new object
        docked_name = f"{output_prefix}_{i:04d}"

        # Create a new object from the specific state
        cmd.create(docked_name, "docked", i, 1)

        # Save the new object to a file in the specified directory
        output_file = os.path.join(output_dir, f"{docked_name}.mol2")
        cmd.save(output_file, docked_name)

        # Print the name of the saved file for reference
        print(f"Saved {output_file}")

    # Clean up the original multi-state object
    cmd.delete("docked")

# Function to calculate RMSD values and collect them
def calculate_rmsd(undocked_file, docked_dir, num_docked, output_prefix):
    rmsd_results = []

    # Load the undocked structure
    cmd.load(undocked_file, "undocked")

    # Iterate through each docked file
    for i in range(1, num_docked + 1):
        # Generate the file name for the docked structure
        docked_file = os.path.join(docked_dir, f"{output_prefix}_{i:04d}.mol2")
        docked_name = f"{output_prefix}_{i:04d}"

        # Load the docked structure
        cmd.load(docked_file, docked_name)

        # Align the structures without fitting (transform=0)
        rmsd = cmd.align("undocked", docked_name, cycles=0, transform=0)[0]

        # Append RMSD result to list
        rmsd_results.append((i, rmsd))

        # Clean up the created object to avoid interference
        cmd.delete(docked_name)

    # Clean up the undocked structure
    cmd.delete("undocked")

    return rmsd_results

# Function to export RMSD results to Excel
def export_to_excel(rmsd_results, output_file):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = 'RMSD Results'

    # Write headers
    sheet['A1'] = 'Pose Number'
    sheet['B1'] = 'RMSD'

    # Write RMSD results
    for idx, (pose_num, rmsd_value) in enumerate(rmsd_results, start=2):
        sheet[f'A{idx}'] = pose_num
        sheet[f'B{idx}'] = rmsd_value

    # Save workbook
    wb.save(output_file)

# Function to run splitting and RMSD calculation
def run_split_and_rmsd():
    input_file = entry_input_file.get()
    output_prefix = entry_output_prefix.get()
    undocked_file = entry_undocked_file.get()
    docked_dir = entry_docked_dir.get()
    num_docked = int(entry_num_docked.get())

    if not input_file or not output_prefix or not undocked_file or not docked_dir or not num_docked:
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    # Split and export docked poses
    split_and_export_docked_poses(input_file, output_prefix, docked_dir)

    # Calculate RMSD values
    rmsd_results = calculate_rmsd(undocked_file, docked_dir, num_docked, output_prefix)

    # Export RMSD results to Excel
    excel_output_file = os.path.join(docked_dir, f"RMSD_results_{output_prefix}.xlsx")
    export_to_excel(rmsd_results, excel_output_file)

    messagebox.showinfo("Success", f"Docked poses split and RMSD values saved to:\n{excel_output_file}")

# Function to browse for a file
def browse_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("MOL2 files", "*.mol2"), ("PDB files", "*.pdb"), ("PDBQT files", "*.pdbqt")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)

# Function to browse for a directory
def browse_directory(entry):
    directory = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, directory)

# Create the main window
root = tk.Tk()
root.title("Docked Pose Splitter and RMSD Calculator")

# Create and place labels and entry widgets
tk.Label(root, text="Input file with multiple docked poses:").grid(row=0, column=0, padx=10, pady=5)
entry_input_file = tk.Entry(root, width=50)
entry_input_file.grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=lambda: browse_file(entry_input_file)).grid(row=0, column=2, padx=10, pady=5)

tk.Label(root, text="Output prefix for docked pose files:").grid(row=1, column=0, padx=10, pady=5)
entry_output_prefix = tk.Entry(root, width=50)
entry_output_prefix.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Undocked structure file:").grid(row=2, column=0, padx=10, pady=5)
entry_undocked_file = tk.Entry(root, width=50)
entry_undocked_file.grid(row=2, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=lambda: browse_file(entry_undocked_file)).grid(row=2, column=2, padx=10, pady=5)

tk.Label(root, text="Directory to save docked pose files:").grid(row=3, column=0, padx=10, pady=5)
entry_docked_dir = tk.Entry(root, width=50)
entry_docked_dir.grid(row=3, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=lambda: browse_directory(entry_docked_dir)).grid(row=3, column=2, padx=10, pady=5)

tk.Label(root, text="Number of docked poses:").grid(row=4, column=0, padx=10, pady=5)
entry_num_docked = tk.Entry(root, width=50)
entry_num_docked.grid(row=4, column=1, padx=10, pady=5)

# Create and place the run button
tk.Button(root, text="Run", command=run_split_and_rmsd).grid(row=5, column=0, columnspan=3, pady=10)

# Run the GUI event loop
root.mainloop()
