# -*- coding: utf-8 -*-
"""
Abstract class facilitating the GUI for the Uniformly Spaced Points Selector app.
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
from abc import ABC, abstractmethod  # Creating abstract class
import tkinter as tk  # GUI creation tools
import os  # Working with paths


class Uniformly_Spaced_Points_Selector_GUI(tk.Frame, ABC):
    def __init__(self, master):
        super().__init__(master)

        # Create a reference to the master (root window)
        self.master = master

        # Getting the paths to app icons
        main_menu_icon_name = "geo_map_marker_icon.ico"
        help_menu_icon_name = "help_menu_icon.ico"

        path_to_current_folder = os.path.dirname(os.path.abspath(__file__))

        self.path_to_main_menu_icon = os.path.join(
            path_to_current_folder, main_menu_icon_name
        )
        self.path_to_help_menu_icon = os.path.join(
            path_to_current_folder, help_menu_icon_name
        )

        # Changing the icon amd title of the root window and making it non-resizable
        self.master.iconbitmap(self.path_to_main_menu_icon)
        self.master.title("USP Selector")
        self.master.resizable(False, False)

        # Defining additional object parameters/properties used primarily in validating user actions/input ("global" variables)
        self.path_to_excel_file = None
        self.is_chosen_file_valid = False

        # Defining button font and size
        button_font = ("Segoe UI", 18, "bold")
        button_width = 16

        # Creating buttons
        self.info_button = tk.Button(
            self,
            text="User Manual",
            font=button_font,
            width=button_width,
            command=self.info_button_event,
        )

        self.load_button = tk.Button(
            self,
            text="Load Excel File",
            font=button_font,
            width=button_width,
            command=self.load_button_event,
        )

        self.execute_button = tk.Button(
            self,
            text="Select Points",
            font=button_font,
            command=self.execute_button_event,
        )

        self.quit_button = tk.Button(
            self,
            text="Exit Application",
            font=button_font,
            width=button_width,
            command=self.exit_button_event,
        )

        # Defining input field font and size
        input_menu_font = ("Segoe UI", 11, "bold")
        input_menu_width = 2

        # Creating input menu for user
        # Define StringVar control variable with initial value
        self.input_menu_control_variable = tk.StringVar(self, value="2")
        # Define possible options for the menu
        self.maximal_acceptable_value = 15
        input_menu_options = [
            str(number) for number in range(2, self.maximal_acceptable_value + 1)
        ]
        # Create menu widget and configure it
        self.input_menu = tk.OptionMenu(
            self, self.input_menu_control_variable, *input_menu_options
        )
        self.input_menu.config(font=input_menu_font, width=input_menu_width)
        # Change the font of all options, including the not selected ones (so that when the menu is opened, all of them have the same font)
        input_menu_options_handle = self.nametowidget(self.input_menu.menuname)
        # Set the dropdown menu's font
        input_menu_options_handle.config(font=input_menu_font)

        # Add input menu label
        self.input_menu_label = tk.Label(self, text="Count", font=input_menu_font)

        # Set initial state of execute button to disabled
        self.set_execution_button_state("disabled")

        # Define gaps in between the panel elements
        x_padding = 5
        y_padding = 5

        # Placing and positioning the widgets on the window using grid
        self.info_button.grid(row=0, columnspan=2, padx=x_padding, pady=(y_padding, 0))
        self.load_button.grid(row=1, columnspan=2, padx=x_padding, pady=(y_padding, 0))
        self.execute_button.grid(
            row=2,
            column=1,
            columnspan=1,
            rowspan=2,
            sticky="EW",
            padx=(0, x_padding),
            pady=(y_padding, 0),
        )
        self.input_menu_label.grid(row=2, column=0, sticky="S", padx=x_padding - 2)
        self.input_menu.grid(
            row=3, column=0, columnspan=1, sticky="NSEW", padx=x_padding - 2
        )
        self.quit_button.grid(row=4, columnspan=2, padx=x_padding, pady=y_padding)

    # Modify the state of execute and input fields
    @abstractmethod
    def set_execution_button_state(self, state):
        pass

    # Reset the help window position
    @abstractmethod
    def reset_info_window_position(self):
        pass

    # Define info button actions
    @abstractmethod
    def info_button_event(self):
        pass

    # Define exit button actions
    @abstractmethod
    def exit_button_event(self):
        pass

    # Define load button actions
    @abstractmethod
    def load_button_event(self):
        pass

    # Define execute button actions
    @abstractmethod
    def execute_button_event(self):
        pass
