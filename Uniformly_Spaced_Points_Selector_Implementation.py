# -*- coding: utf-8 -*-
"""
Uniformly Spaced Points Selector app.
Copyright (C) 2022  Peter Chmurčiak

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see https://www.gnu.org/licenses/.
"""
import tkinter as tk  # Running the app
from tkinter import messagebox, filedialog  # Interacting with user, displaying messages
import ctypes  # To specify what "group" the app belongs to - to show a custom icon on the taskbar

from Uniformly_Spaced_Points_Selector_Abstract_GUI import Uniformly_Spaced_Points_Selector_GUI
from Subroutine_Select_The_Points import Select_The_Points
from Subroutine_Get_Rows_And_Columns_Of_An_Excel_File import Get_Rows_And_Columns_Of_An_Excel_File
from Subroutine_Validate_Coordinate_Values import Validate_Coordinate_Values


class Uniformly_Spaced_Points_Selector(Uniformly_Spaced_Points_Selector_GUI):
    def __init__(self, master):
        super().__init__(master)
        # Create a reference to the master (root window)
        self.master = master

    # Modify the state of execute and input fields
    def set_execution_button_state(self, state):
        self.input_menu_label.config(state=state)
        self.input_menu.config(state=state)
        self.execute_button.config(state=state)

    # Reset the help window position
    def reset_info_window_position(self):
        x_coordinate_of_main_window = self.master.winfo_x()
        y_coordinate_of_main_window = self.master.winfo_y()
        width_of_main_window = self.master.winfo_width()
        self.info_window.geometry(
            f"550x508+{x_coordinate_of_main_window + width_of_main_window + 10}+{y_coordinate_of_main_window}"
        )
        self.info_window.deiconify()

    # Define info button actions
    def info_button_event(self):
        try:
            self.reset_info_window_position()
        except (AttributeError, tk.TclError):
            self.info_window = tk.Toplevel(self)
            self.info_window.iconbitmap(self.path_to_help_menu_icon)
            self.info_window.title("User Manual")
            self.reset_info_window_position()
            self.info_window.resizable(False, False)
            # self.info_window.focus_set() # done also by deiconify

            self.info_frame = tk.Frame(self.info_window)
            self.info_text = tk.Text(self.info_frame, wrap=tk.WORD)

            # Define normal and bold text style through tags
            self.info_text.tag_configure("normal", font=("Segoe UI", 10))
            self.info_text.tag_configure(
                "bold", font=("Segoe UI", 10, "bold", "underline")
            )

            # Write the help menu text line by line
            self.info_text.insert(tk.END, " ", "normal")
            self.info_text.insert(tk.END, "Purpose:\t", "bold")
            self.info_text.insert(
                tk.END,
                "Find a subset of geographical points in such a way, as to cover the largest area"
                + "\n\tpossible. The points could be therefore thought of as a representative sample"
                + "\n\tof the area.",
                "normal",
            )

            self.info_text.insert(tk.END, "\n\n ", "normal")
            self.info_text.insert(tk.END, "Input:\t", "bold")
            self.info_text.insert(
                tk.END,
                "A prerequisite for the correct functioning of the program is the input in the"
                + "\n\tformat of an Excel file with header in row 1, containing exactly 3 columns (A,B,C)"
                + "\n\tand at least 4 rows including the header. The text of the column headers does"
                + "\n\tnot matter, but it is necessary to maintain the order of the columns, namely:"
                + "\n\tColumn A - Auxiliary name/description"
                + "\n\tColumn B - Latitude [°]"
                + "\n\tColumn C - Longitude [°]",
                "normal",
            )

            self.info_text.insert(tk.END, "\n\n ", "normal")
            self.info_text.insert(tk.END, "Order:\t", "bold")
            self.info_text.insert(
                tk.END,
                '1. Select the Excel file using "Load Excel File"'
                + '\n\t2. Specify number of sought points in the field "Count"'
                + '\n\t3. Begin the selection process with "Select Points"',
                "normal",
            )

            self.info_text.insert(tk.END, "\n\n ", "normal")
            self.info_text.insert(tk.END, "Output:\t", "bold")
            self.info_text.insert(
                tk.END,
                "The output of the program is a modified Excel file containing marked points"
                + '\n\twith extension "_MARKED" and HTML file with extension "_MAP" displaying'
                + "\n\tpoints graphically. Both outputs are stored in the input file folder.",
                "normal",
            )

            self.info_text.insert(tk.END, "\n\n ", "normal")
            self.info_text.insert(tk.END, "Note:\t", "bold")
            self.info_text.insert(
                tk.END,
                "The current algorithm for choosing points may not provide good results when"
                + "\n\tpicking large number of points relative to the original set size - points will begin"
                + "\n\tto cluster. Additionally, with increasing size of the original set or increasing"
                + "\n\tnumber of selected points, the time and computing power demands rise. For"
                + '\n\tthis reason, the "Count" input field is atrificially costrained.',
                "normal",
            )

            self.info_text.insert(tk.END, "\n\n ", "normal")
            self.info_text.insert(tk.END, "Update:", "bold")
            self.info_text.insert(tk.END, "\t11.02.2022", "normal")

            self.info_text.insert(tk.END, "\n\n ", "normal")
            self.info_text.insert(tk.END, "Author:", "bold")
            self.info_text.insert(tk.END, "\tPeter Chmurčiak", "normal")

            self.info_text.pack(fill="both", expand=True, padx=5, pady=5)
            self.info_frame.pack(fill="both", expand=True)

    # Define exit button actions
    def exit_button_event(self):
        self.master.destroy()

    # Define load button actions
    def load_button_event(self):
        # Disable the execute button by default
        self.set_execution_button_state("disabled")
        # Assume invalid file by default
        self.is_chosen_file_valid = False
        # Prompt user to choose a excel file
        self.path_to_excel_file = filedialog.askopenfilename(
            filetypes=(("Excel files (*.xlsx)", ("*.xls", "*.xlsx")),)
        )
        # If file was chosen
        if self.path_to_excel_file:
            # Extract number of rows/columns from the file
            (
                chosen_file_number_of_rows,
                chosen_file_number_of_columns,
            ) = Get_Rows_And_Columns_Of_An_Excel_File(self.path_to_excel_file)
            # Validate the file based on the rows/columns and show error messages based on specific situations
            if chosen_file_number_of_columns == 3 and chosen_file_number_of_rows > 3:
                if Validate_Coordinate_Values(self.path_to_excel_file):
                    self.is_chosen_file_valid = True
                    # Edit the menu bar options based on the number of points within the file
                    # Remove all the original options
                    self.input_menu["menu"].delete(0, "end")
                    # Define and add one by one the new options
                    maximal_allowed_value = min(
                        self.maximal_acceptable_value, chosen_file_number_of_rows - 1
                    )
                    new_choices = [
                        str(number) for number in range(2, maximal_allowed_value + 1)
                    ]
                    for choice in new_choices:
                        self.input_menu["menu"].add_command(
                            label=choice,
                            command=tk._setit(self.input_menu_control_variable, choice),
                        )
                    # In case the new range of options is smaller than the previous one, edit the selected option to be within range
                    if self.input_menu_control_variable.get() not in new_choices:
                        self.input_menu_control_variable.set(new_choices[-1])
                    # Enable the execution button
                    self.set_execution_button_state("normal")
                    messagebox.showinfo(
                        "Info",
                        f"Loading operation was succesful. Loaded file contains {chosen_file_number_of_rows-1} points.",
                    )
                else:
                    messagebox.showerror(
                        "Error",
                        "Invalid data format. The values in 2nd and 3rd column do not appear to be GPS coordinates. Please make sure that the content and order of the columns is valid and repeat the operation.",
                    )
            elif chosen_file_number_of_columns == 3:
                messagebox.showerror(
                    "Error",
                    f"Invalid data format. {chosen_file_number_of_rows} rows were found. Application expects at least 4 rows. Please make sure that the content and order of the columns is valid and repeat the operation.",
                )
            elif chosen_file_number_of_rows > 3:
                messagebox.showerror(
                    "Error",
                    f"Invalid data format. {chosen_file_number_of_columns} columns were found instead of expected 3. Please make sure that the content and order of the columns is valid and repeat the operation.",
                )
            else:
                messagebox.showerror(
                    "Error",
                    "Invalid data format. Please make sure that the content and order of the columns is valid and repeat the operation.",
                )
        # If file selection window was closed
        else:
            messagebox.showerror("Error", "No file was selected.")

    # Define execute button actions
    def execute_button_event(self):
        # If path and file it leads to is valid
        if self.path_to_excel_file and self.is_chosen_file_valid:
            # Get input from input field
            number_of_points_to_select = int(self.input_menu_control_variable.get())
            try:
                save_file_name = Select_The_Points(
                    self.path_to_excel_file, number_of_points_to_select
                )
            except PermissionError:
                messagebox.showerror(
                    "Error",
                    'Access to the selected file was denied. Make sure that the corresponding Excel file marked with "MARKED" is closed and repeat the operation.',
                )
            else:
                self.path_to_excel_file = None
                self.is_chosen_file_valid = False
                self.set_execution_button_state("disabled")
                messagebox.showinfo(
                    "Info",
                    f'Points ({number_of_points_to_select}) were sucessfully selected. Output was stored in the file "{save_file_name}" within the original folder.',
                )


if __name__ == "__main__":
    # Specifying "group" the app will belong to - to be able to show the icon on the taskbar (arbitrary string)
    arbitrary_app_id = "mycompany.myproduct.subproduct.version"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(arbitrary_app_id)
    root = tk.Tk()
    Uniformly_Spaced_Points_Selector(root).pack(side="top", fill="both", expand=True)
    root.mainloop()
