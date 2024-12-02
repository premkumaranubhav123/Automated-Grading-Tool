import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import os

class MultiHandleSliderApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Grading Tool")
        self.root.configure(bg='#FFFFFF')  # background

        self.data = None  # Initialize data
        self.criteria = {}  
        self.filePath = None
        self.average_value = None

        self.update_interval = 0  # milliseconds (approx. 60 FPS)
        self.last_update_time = 0
        self.is_dragging = False
        self.update_after_id = None

        self.heading_label = tk.Label(root, text="Grading Tool", font=("open sans", 28, "bold"), bg='#FFFFFF')
        self.heading_label.pack(pady=10)

        # Add button to load Excel file
        self.load_button = tk.Button(self.root, text="Load Excel File", command=self.load_data, font=("open sans", 12, "bold" ))
        self.load_button.pack(pady=5)

        # File selection frame
        file_frame = tk.Frame(root,bg='#FFFFFF')
        file_frame.pack(pady=10)

        tk.Label(file_frame, text="Output Excel File:",font=("open sans", 12, "bold" )).grid(row=1, column=0, padx=10, pady=10)
        self.output_file_entry = tk.Entry(file_frame, width=50)
        self.output_file_entry.grid(row=1, column=1, padx=10, pady=10)
        tk.Button(file_frame, text="Save As", command=self.save_file,font=("open sans", 12, "bold" )).grid(row=1, column=2, padx=10, pady=10)

        tk.Button(file_frame, text="Run Grading", command=self.run_grading,font=("open sans", 12, "bold" )).grid(row=2, column=1, padx=10, pady=20)

        # Create a canvas for scrolling
        self.scroll_canvas = tk.Canvas(self.root,bg='#FFFFFF')
        self.scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.scroll_canvas.yview)
        self.scrollbar.pack(side="right", fill="y")

        # Create a frame inside the canvas
        self.scrollable_frame = tk.Frame(self.scroll_canvas,bg='#FFFFFF')

        # Add the frame to the canvas
        self.scroll_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scroll_canvas.pack(side="left", fill="both", expand=True)
        self.scroll_canvas.config(yscrollcommand=self.scrollbar.set)
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)

        # Create a frame for the table and plot
        self.main_frame = tk.Frame(self.scrollable_frame,bg='#FFFFFF')
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Frame for the table
        self.table_frame = tk.Frame(self.main_frame,bg='#FFFFFF',padx=5)
        self.table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Frame for the plot and sliders
        self.plot_frame = tk.Frame(self.main_frame,bg='#FFFFFF')
        self.plot_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create the matplotlib figure and axis for the plot
        self.figure, self.ax = plt.subplots(figsize=(12, 6), dpi=100)
        self.plot_canvas = FigureCanvasTkAgg(self.figure, master=self.plot_frame)
        self.plot_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Create the matplotlib figure and axis for the table
        self.table_figure, self.table_ax = plt.subplots(figsize=(10, 2), dpi=100)
        self.table_ax.axis('tight')
        self.table_ax.axis('off')

        # Create a frame for the slider
        self.slider_frame = tk.Frame(self.plot_frame)
        self.slider_frame.pack(fill=tk.X, padx=50)

        # Slider properties
        self.slider_width = 1100  # Adjusted width for 9 sliders
        self.slider_frame_width = 1100
        self.slider_height = 20
        self.min_value = 0
        self.max_value = 100
        self.num_handles = 9  # Number of sliders

        # Create a Canvas for the custom slider
        self.slider_canvas = tk.Canvas(self.slider_frame,width=self.slider_frame_width, height=50, bg='white')
        self.slider_canvas.pack(fill=tk.X)

        # Draw slider track
        self.track = self.slider_canvas.create_line(50, 25, 1050, 25, width=3, fill="grey")

        # Create slider handles
        self.handles = []
        self.value_labels = []
        self.value_entries = []  # Entry widgets for manual input
        handle_spacing = (self.slider_width - 100) / (self.num_handles - 1)  # Adjust spacing between handles

        default_values = [0, 15, 25, 40, 50, 65, 75, 85, 95]

        # Add these new attributes
        self.update_after_id = None
        self.is_dragging = False

        for i in range(self.num_handles):
            x_position = 50 + i * handle_spacing
            handle = self.slider_canvas.create_oval(x_position - 5, 20, x_position + 5, 30, fill='navy blue', outline='black')
            label = self.slider_canvas.create_text(x_position, 40, text="0", fill="black")
            self.handles.append(handle)
            self.value_labels.append(label)

            # Entry widgets for direct input
            entry = tk.Entry(self.slider_frame, width=5 , bg='#FFFFFF')
            entry.pack(side=tk.LEFT, padx=22)
            self.value_entries.append(entry)
            entry.insert(0, str(default_values[i]))
            # Bind the Entry widgets to update slider on Return key
            entry.bind("<Return>", lambda _, index=i: self.update_handle_position(index))

            # Bind mouse events
            self.slider_canvas.tag_bind(handle, "<Button-1>", self.on_click)
            self.slider_canvas.tag_bind(handle, "<B1-Motion>", self.on_drag)
            self.slider_canvas.tag_bind(handle, "<ButtonRelease-1>", self.on_release)

        # Store handle positions
        self.handle_positions = [50 + i * handle_spacing for i in range(self.num_handles)]
        self.hand_name = ['F', 'E', 'D', 'C-', 'C', 'B-', 'B', 'A-', 'A']

         # Set default positions for sliders
        self.set_default_handle_positions(default_values)

         # Add a label for displaying the mark difference
        self.diff_label = tk.Label(self.slider_frame, text="", font=("open sans", 16,"bold"))
        self.diff_label.pack(pady=5)

        # Store sorted marks
        self.sorted_marks = []

        # Initial plot
        self.update_histogram()

    def set_default_handle_positions(self, default_values):
        # Update handle positions based on default values
        handle_spacing = (self.slider_width - 100) / (self.num_handles - 1)
        for i, value in enumerate(default_values):
            new_x = 50 + (value / self.max_value) * (self.slider_width - 100)
            self.handle_positions[i] = new_x
            self.slider_canvas.coords(self.handles[i], new_x - 5, 20, new_x + 5, 30)
            self.slider_canvas.itemconfig(self.value_labels[i], text=f"{value}")
            # Clear entry widget and set default value
            self.value_entries[i].delete(0, tk.END)
            self.value_entries[i].insert(0, str(value))

    def on_click(self, event):
        closest_handle = event.widget.find_closest(event.x, event.y)
        if closest_handle:
            self.drag_data = {'handle': closest_handle[0], 'start_x': event.x}
            self.is_dragging = True

    def on_frame_configure(self, event):
        # Update the scroll region of the scroll canvas
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def load_data(self):
        self.filePath = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.filePath:
            self.data = self.read_excel_file(self.filePath)
            self.average_value = self.data.mean()
            if self.data is not None:
                self.sorted_marks = sorted(self.data, reverse=True)    
                self.update_histogram()

    def read_excel_file(self, file_path):
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the Excel file: {e}")
            return
        if 'Marks' not in df.columns:
            messagebox.showerror("Error", "The Excel file must contain a 'Marks' column.")
            return
        return df['Marks']

    def save_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.output_file_entry.delete(0, tk.END)
            self.output_file_entry.insert(0, file_path)

    def run_grading(self):
        input_file = self.filePath
        output_file = self.output_file_entry.get()

        if not os.path.exists(input_file):
            messagebox.showerror("Error", "The selected input file does not exist.")
            return

        if not output_file:
            messagebox.showerror("Error", "Please select an output file.")
            return

        self.update_grades(input_file, output_file)

    def update_grades(self, input_file, output_file):
        try:
            df = pd.read_excel(input_file)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the Excel file: {e}")
            return

        if 'Marks' not in df.columns:
            messagebox.showerror("Error", "The Excel file must contain a 'Marks' column.")
            return

        df['Grade'] = df['Marks'].apply(self.assign_grade)

        try:
            df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Grades updated and saved to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save the updated Excel file: {e}")

    def assign_grade(self, marks):
        for grade, range_str in self.criteria.items():
            lower_bound, upper_bound = map(float, range_str.split('-'))
            if lower_bound < marks <= upper_bound:
                return grade
        return 'No Grade'

    def on_click(self, event):
        closest_handle = event.widget.find_closest(event.x, event.y)
        if closest_handle:
            self.drag_data = {'handle': closest_handle[0], 'start_x': event.x}

    def on_drag(self, event):
        handle = self.drag_data.get('handle')
        if handle is not None:
            try:
                handle_index = self.handles.index(handle)
                print(f"Handle index: {handle_index}")
                x_diff = event.x - self.drag_data['start_x']
                new_x = self.handle_positions[handle_index] + x_diff

                # Constrain movement to slider track range (50 to 750)
                if handle_index == 0:  # First handle
                    new_x = max(50, min(new_x, self.handle_positions[1]))
                elif handle_index == 8:  # Last handle
                    new_x = max(self.handle_positions[7], min(new_x, 1050))
                else:  # All other handles
                    new_x = max(self.handle_positions[handle_index - 1], min(new_x, self.handle_positions[handle_index + 1]))

                # Move the handle
                self.slider_canvas.move(handle, new_x - self.handle_positions[handle_index], 0)
                self.handle_positions[handle_index] = new_x

                # Update the position where dragging started
                self.drag_data['start_x'] = event.x
                
                # Update entry widget with the current value
                value = round((new_x - 50) / (self.slider_width - 100) * self.max_value, 2)
                self.value_entries[handle_index].delete(0, tk.END)
                self.value_entries[handle_index].insert(0, str(value))

                # Calculate and display mark difference
                self.calculate_mark_difference(value)

                # Update slider labels
                label_text = f"{self.hand_name[handle_index]}: {value}"
                self.slider_canvas.coords(self.value_labels[handle_index], new_x, 40)
                self.slider_canvas.itemconfig(self.value_labels[handle_index], text=label_text)

                # Schedule the histogram update
                if self.update_after_id:
                    self.root.after_cancel(self.update_after_id)
                self.update_after_id = self.root.after(0, self.update_histogram)

            except ValueError as e:
                print(f"Error: {e}. Handle ID {handle} not found in self.handles")
        else:
            print("No handle being dragged")

    def on_release(self, event):
        self.is_dragging = False
        if self.update_after_id:
            self.root.after_cancel(self.update_after_id)
        self.update_histogram()

    def calculate_mark_difference(self, current_value):
        if not self.sorted_marks:
            return

        # Find the two closest marks
        lower_mark = next((mark for mark in self.sorted_marks if mark <= current_value), None)
        upper_mark = next((mark for mark in reversed(self.sorted_marks) if mark > current_value), None)

        if lower_mark is not None and upper_mark is not None:
            difference = upper_mark - lower_mark
            self.diff_label.config(text=f"Difference : {difference:.2f}")
        else:
            self.diff_label.config(text="")

    def update_handle_position(self, index):
        
        try:
            # Get the value from the entry and convert it to a float
            value = float(self.value_entries[index].get())

            # Convert value to x position
            new_x = 50 + (value / self.max_value) * (self.slider_width - 100)

            # Keep the new_x within bounds
            if index == 0:  # First handle
                if new_x > self.handle_positions[1]:
                    new_x = self.handle_positions[1]
            elif index == len(self.handles) - 1:  # Last handle
                if new_x < self.handle_positions[-2]:
                    new_x = self.handle_positions[-2]
            else:  # Other handles
                if new_x < self.handle_positions[index - 1]:
                    new_x = self.handle_positions[index - 1]
                elif new_x > self.handle_positions[index + 1]:
                    new_x = self.handle_positions[index + 1]

            # Move the handle to the new position
            x_diff = new_x - self.handle_positions[index]
            self.slider_canvas.move(self.handles[index], x_diff, 0)
            self.handle_positions[index] = new_x

             # Calculate and display mark difference
            self.calculate_mark_difference(value)


            # Update histogram and table
            self.update_histogram()
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid number.")

    def update_histogram(self):
        print("accessed update histogram")
        if self.data is None:
            return  # Return if data is not loaded
        
        print(f"data is not none")
        # Clear previous plot
        self.ax.clear()

        # Calculate bin edges based on handle positions
        bin_edges = sorted([(x - 50) / (self.slider_width - 100) * self.max_value for x in self.handle_positions] + [self.max_value])
       
        # Compute histogram values
        counts, _ = np.histogram(self.data, bins=bin_edges)

        # Define colors for each bin
        colors = ['#ff6666','#ff9966','#ffc966','#ccff66','#99ff99','#66ffcc','#66cccc','#3399ff','#6633ff']  # Red to indigo gradient

        # Plot each bin separately
        for i in range(len(bin_edges) - 1):
            self.ax.bar(bin_edges[i], counts[i], width=bin_edges[i + 1] - bin_edges[i], 
                        edgecolor='black', color=colors[i], align='edge')

        self.ax.set_title("Histogram of Marks",)
        self.ax.set_xlabel("Marks")
        self.ax.set_ylabel("No. of Students")

        # Get sorted "Marks" in descending order
        sorted_marks = sorted(self.data, reverse=True)

        self.figure.tight_layout(pad=0.5)

            # Plot red dots for each mark on the slider
        for mark in sorted_marks:
            x_position = 50 + (mark / self.max_value) * (self.slider_width - 100)
            self.slider_canvas.create_oval(x_position - 2, 15, x_position + 2, 17, fill='red', outline='red')

        self.average_value = self.data.mean()
        self.ax.axvline(self.average_value, color='blue', linestyle='--', label=f'Average: {self.average_value:.2f}')
        self.ax.legend()

        # Update slider labels with current values
        for i, handle in enumerate(self.handles):
            value = round((self.handle_positions[i] - 50) / (self.slider_width - 100) * self.max_value, 2)
            label_text = f"{self.hand_name[i]}: {value}"
            self.slider_canvas.coords(self.value_labels[i], self.handle_positions[i], 40)
            self.slider_canvas.itemconfig(self.value_labels[i], text=label_text)

        # Update table data
        self.update_table(bin_edges, counts)

        # Cancel any scheduled updates
        if self.update_after_id:
            self.root.after_cancel(self.update_after_id)
            self.update_after_id = None

        # Draw the plot on the Tkinter canvas
        self.plot_canvas.draw()
        self.plot_canvas.flush_events()

    def update_table(self, bin_edges, counts):
        ranges = [f"{bin_edges[i]:.2f}-{bin_edges[i+1]:.2f}" for i in range(9)]
        total_students = np.sum(counts)
        percentage_students = [(count / total_students) * 100 for count in counts]

        # Calculate the %age Consecutive Grades
        percentage_consecutive = [0] * len(percentage_students)
        consecutive_display = [""] * len(percentage_students)

        # Calculate the combined percentages and determine where to display them
        for i in range(9):
            if i == 0 or i == 1:
                percentage_consecutive[i] = percentage_students[0] + percentage_students[1]
                if i == 0:
                    consecutive_display[i] = f"{percentage_consecutive[i]:.2f}%"
            elif i == 3 or i == 4:
                percentage_consecutive[i] = percentage_students[3] + percentage_students[4]
                if i == 3:
                    consecutive_display[i] = f"{percentage_consecutive[i]:.2f}%"
            elif i == 5 or i == 6:
                percentage_consecutive[i] = percentage_students[5] + percentage_students[6]
                if i == 5:
                    consecutive_display[i] = f"{percentage_consecutive[i]:.2f}%"
            elif i == 7 or i == 8:
                percentage_consecutive[i] = percentage_students[7] + percentage_students[8]
                if i == 7:
                    consecutive_display[i] = f"{percentage_consecutive[i]:.2f}%"
            else:
                percentage_consecutive[i] = percentage_students[i]
                consecutive_display[i] = f"{percentage_consecutive[i]:.2f}%"

        grades = ['F  ', 'E  ', 'D  ', 'C- ', 'C  ', 'B- ', 'B  ', 'A- ', 'A  ']

        table_data = []
        for i in range(9):
            table_data.append([grades[i], ranges[i], counts[i], f"{percentage_students[i]:.2f}%", consecutive_display[i]])
         
             # Check if the grade is already present in the criteria dictionary
            if grades[i] in self.criteria:
                # Update the existing range
                self.criteria[grades[i]] = ranges[i]
            else:
                # Add a new entry if the grade doesn't exist
                self.criteria[grades[i]] = ranges[i]
        
        print(f"criteria {self.criteria}")

        # Prepare to merge cells
        merge_rows = {}
        for i in range(8):
            if consecutive_display[i] == consecutive_display[i + 1]:
                if consecutive_display[i] not in merge_rows:
                    merge_rows[consecutive_display[i]] = [i]
                merge_rows[consecutive_display[i]].append(i + 1)

        # Adjust figure size
        self.table_figure.set_size_inches(7, 6, forward=True)

        self.table_ax.clear()
        self.table_ax.axis('tight')
        self.table_ax.axis('off')

        # Create the table
        table = self.table_ax.table(cellText=table_data, colLabels=['Grade', 'Range of \nGrade', 'No. of \nStudents', '%age of \nStudents', '%age Consecutive\n Grades'], loc='center')
        table.auto_set_column_width(col=[0, 1, 2, 3, 4])
        table.auto_set_font_size(False)
        table.set_fontsize(12)
        table.scale(2, 2.65)

        for (row, col), cell in table.get_celld().items():
            cell.set_text_props(weight='bold')
       
        # Draw the table on a new canvas
        if hasattr(self, 'table_canvas'):
            self.table_canvas.get_tk_widget().destroy()
        self.table_canvas = FigureCanvasTkAgg(self.table_figure, master=self.table_frame)
        self.table_canvas.get_tk_widget().pack(fill=tk.X)
        self.table_canvas.draw()

if __name__ == "__main__":
    root = tk.Tk()
    app = MultiHandleSliderApp(root)
    root.mainloop()
