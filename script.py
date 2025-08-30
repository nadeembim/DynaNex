# DynaNex V1.0 - Professional Parameter Export/Import Tool for Dynamo
# "Power your Connections" - Connects Revit with Microsoft Excel

import clr
import sys
import os
import time
import json

# Add references for Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitNodes')
clr.AddReference('RevitServices')

# Add references for Windows Forms and Excel
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('Microsoft.Office.Interop.Excel')

# Import required modules
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

import System
import System.Windows.Forms as WinForms
import System.Drawing as Drawing
from System.Collections.Generic import List
import Microsoft.Office.Interop.Excel as Excel
from System.Windows.Forms import DialogResult

class DynaNextForm(WinForms.Form):
    def __init__(self):
        self.doc = DocumentManager.Instance.CurrentDBDocument
        self.uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
        self.all_parameters = []
        self.filtered_parameters = []
        self.all_categories = []
        self.filtered_categories = []
        self.is_processing = False
        self.current_scope = "ActiveView"  # "ActiveView" or "WholeModel"
        
        # Global selection tracking (persistent across filters)
        self.selected_categories = set()  # Track all selected categories
        self.selected_parameters = set()  # Track all selected parameters
        
        # Theme tracking
        self.current_theme = "dark"  # "dark" or "light"
        
        # Initialize settings dictionary with defaults
        self.settings = {
            'confirm_whole_model': True,
            'default_scope_active': True,
            'auto_select_common_cats': False,
            'use_main_categories': True,
            'include_empty_params': True,
            'auto_open_excel': True,
            'export_only_selected': True,
            'override_existing': True,
            'skip_readonly': True
        }
        
        # Get settings file path in user's Documents folder
        self.settings_file_path = os.path.join(
            os.path.expanduser("~"), 
            "Documents", 
            "DynaNex_Settings.json"
        )
        
        # Initialize UI components that might be referenced during form creation
        self.categories_checklist = None
        self.parameters_checklist = None
        self.category_count_label = None
        self.parameter_count_label = None
        self.info_label = None
        self.progress_bar = None
        self.status_label = None
        self.search_categories_textbox = None
        self.search_parameters_textbox = None
        self.settings_status_label = None
        
        # Initialize form
        self.InitializeForm()
        
        # Try to load saved settings on startup
        self.LoadSettingsFromFile()
        
        # Load data after UI is fully initialized
        try:
            self.LoadDataFromCurrentScope()
        except Exception as e:
            # If initial load fails, just set empty data
            self.all_parameters = []
            self.filtered_parameters = []
            self.all_categories = []
            self.filtered_categories = []
        
    def InitializeForm(self):
        # Form properties - Dark themed form
        self.Text = "DynaNex V1.0 - Power your Connections"
        self.Size = Drawing.Size(1400, 950)  # Wider for dual pane
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.Sizable
        self.MinimumSize = Drawing.Size(1200, 850)
        self.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        
        # Main panel with dark theme
        main_panel = WinForms.TableLayoutPanel()
        main_panel.Dock = WinForms.DockStyle.Fill
        main_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        main_panel.RowCount = 7
        main_panel.ColumnCount = 1
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 80))  # Title
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 50))  # Scope buttons
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 45))  # Search
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 35))  # Info
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 60))   # Categories + Parameters (split)
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 55))  # Selection buttons
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 100)) # Action buttons
        
       # Title panel with DynaNex branding AND Save/Load buttons - Dark Theme
        title_panel = WinForms.TableLayoutPanel()
        title_panel.Dock = WinForms.DockStyle.Fill
        title_panel.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Slightly lighter dark
        title_panel.ColumnCount = 3  # Left (empty), Center (title), Right (save/load buttons)
        title_panel.RowCount = 2
        title_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 25))  # Left space
        title_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50))  # Title center
        title_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 25))  # Right buttons
        title_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 60))
        title_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 40))

# Title section (center)
        title_center_panel = WinForms.TableLayoutPanel()
        title_center_panel.Dock = WinForms.DockStyle.Fill
        title_center_panel.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        title_center_panel.RowCount = 2
        title_center_panel.ColumnCount = 1
        title_center_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 60))
        title_center_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 40))

        title_label = WinForms.Label()
        title_label.Text = "DynaNex"
        title_label.Font = Drawing.Font("Segoe UI", 20, Drawing.FontStyle.Bold)  # Larger, modern font
        title_label.Dock = WinForms.DockStyle.Fill
        title_label.TextAlign = Drawing.ContentAlignment.BottomCenter
        title_label.ForeColor = Drawing.Color.FromArgb(0, 150, 255)  # Brighter blue for dark theme
        title_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background

        tagline_label = WinForms.Label()
        tagline_label.Text = "Power your Connections"
        tagline_label.Font = Drawing.Font("Segoe UI", 11, Drawing.FontStyle.Italic)  # Slightly larger
        tagline_label.Dock = WinForms.DockStyle.Fill
        tagline_label.TextAlign = Drawing.ContentAlignment.TopCenter
        tagline_label.ForeColor = Drawing.Color.FromArgb(180, 180, 180)  # Light gray for dark theme
        tagline_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background

        title_center_panel.Controls.Add(title_label, 0, 0)
        title_center_panel.Controls.Add(tagline_label, 0, 1)

        # Save/Load buttons section (top-right) - REDESIGNED
        saveload_panel = WinForms.TableLayoutPanel()
        saveload_panel.Dock = WinForms.DockStyle.Fill
        saveload_panel.RowCount = 3
        saveload_panel.ColumnCount = 2
        saveload_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 40))  # Buttons row
        saveload_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 25))  # Status row
        saveload_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))  # Fill remaining
        saveload_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50))
        saveload_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50))
        saveload_panel.Padding = WinForms.Padding(10, 5, 10, 5)
        saveload_panel.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background

        # Save Settings button
        save_settings_btn = WinForms.Button()
        save_settings_btn.Text = "Save As"
        save_settings_btn.Size = Drawing.Size(85, 35)
        save_settings_btn.Anchor = WinForms.AnchorStyles.None
        save_settings_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green
        save_settings_btn.ForeColor = Drawing.Color.White
        save_settings_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        save_settings_btn.FlatStyle = WinForms.FlatStyle.Flat
        save_settings_btn.FlatAppearance.BorderSize = 1
        save_settings_btn.FlatAppearance.BorderColor = Drawing.Color.White
        save_settings_btn.Click += self.SaveSettingsToFile

        # Load Settings button
        load_settings_btn = WinForms.Button()
        load_settings_btn.Text = "Load As"
        load_settings_btn.Size = Drawing.Size(85, 35)
        load_settings_btn.Anchor = WinForms.AnchorStyles.None
        load_settings_btn.BackColor = Drawing.Color.FromArgb(59, 130, 246)  # Modern blue
        load_settings_btn.ForeColor = Drawing.Color.White
        load_settings_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        load_settings_btn.FlatStyle = WinForms.FlatStyle.Flat
        load_settings_btn.FlatAppearance.BorderSize = 1
        load_settings_btn.FlatAppearance.BorderColor = Drawing.Color.White
        load_settings_btn.Click += self.LoadSettingsFromFile

        # Settings status label
        self.settings_status_label = WinForms.Label()
        self.settings_status_label.Text = "Settings: Default"
        self.settings_status_label.Font = Drawing.Font("Segoe UI", 8)
        self.settings_status_label.ForeColor = Drawing.Color.FromArgb(156, 163, 175)  # Light gray for dark theme
        self.settings_status_label.TextAlign = Drawing.ContentAlignment.MiddleCenter
        self.settings_status_label.Dock = WinForms.DockStyle.Fill
        self.settings_status_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background

        saveload_panel.Controls.Add(save_settings_btn, 0, 0)
        saveload_panel.Controls.Add(load_settings_btn, 1, 0)
        saveload_panel.Controls.Add(self.settings_status_label, 0, 1)
        saveload_panel.SetColumnSpan(self.settings_status_label, 2)

        # About Me and Theme Toggle section (top-left) - Updated Layout
        aboutme_panel = WinForms.TableLayoutPanel()
        aboutme_panel.Dock = WinForms.DockStyle.Fill
        aboutme_panel.RowCount = 2
        aboutme_panel.ColumnCount = 1
        aboutme_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 40))  # Button row
        aboutme_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))  # Text and spacing
        aboutme_panel.Padding = WinForms.Padding(10, 5, 10, 5)
        aboutme_panel.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background

        # Horizontal panel for button and theme toggles
        button_row_panel = WinForms.TableLayoutPanel()
        button_row_panel.Dock = WinForms.DockStyle.Fill
        button_row_panel.ColumnCount = 4
        button_row_panel.RowCount = 1
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 85))  # About Me button
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 35))  # Larger spacer to push toggles further right
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 25))  # Dark theme dot
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 25))  # Light theme dot

        # About Me button
        aboutme_btn = WinForms.Button()
        aboutme_btn.Text = "About Me"
        aboutme_btn.Size = Drawing.Size(85, 35)
        aboutme_btn.Anchor = WinForms.AnchorStyles.None
        aboutme_btn.BackColor = Drawing.Color.FromArgb(20, 184, 166)  # Modern teal
        aboutme_btn.ForeColor = Drawing.Color.White
        aboutme_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        aboutme_btn.FlatStyle = WinForms.FlatStyle.Flat
        aboutme_btn.FlatAppearance.BorderSize = 1
        aboutme_btn.FlatAppearance.BorderColor = Drawing.Color.White
        aboutme_btn.Click += self.OpenLinkedInProfile

        # Dark theme button (dot style)
        self.dark_theme_btn = WinForms.Button()
        self.dark_theme_btn.Text = "●"
        self.dark_theme_btn.Size = Drawing.Size(25, 25)
        self.dark_theme_btn.Anchor = WinForms.AnchorStyles.None
        self.dark_theme_btn.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark blue
        self.dark_theme_btn.ForeColor = Drawing.Color.White
        self.dark_theme_btn.Font = Drawing.Font("Segoe UI", 12, Drawing.FontStyle.Bold)
        self.dark_theme_btn.FlatStyle = WinForms.FlatStyle.Flat
        self.dark_theme_btn.FlatAppearance.BorderSize = 2
        self.dark_theme_btn.FlatAppearance.BorderColor = Drawing.Color.White  # Active border
        self.dark_theme_btn.Click += self.SwitchToDarkTheme

        # Light theme button (dot style)
        self.light_theme_btn = WinForms.Button()
        self.light_theme_btn.Text = "●"
        self.light_theme_btn.Size = Drawing.Size(25, 25)
        self.light_theme_btn.Anchor = WinForms.AnchorStyles.None
        self.light_theme_btn.BackColor = Drawing.Color.White
        self.light_theme_btn.ForeColor = Drawing.Color.Black
        self.light_theme_btn.Font = Drawing.Font("Segoe UI", 12, Drawing.FontStyle.Bold)
        self.light_theme_btn.FlatStyle = WinForms.FlatStyle.Flat
        self.light_theme_btn.FlatAppearance.BorderSize = 1
        self.light_theme_btn.FlatAppearance.BorderColor = Drawing.Color.Gray  # Inactive border
        self.light_theme_btn.Click += self.SwitchToLightTheme

        # Theme label panel for text under toggles
        theme_label_panel = WinForms.TableLayoutPanel()
        theme_label_panel.Dock = WinForms.DockStyle.Fill
        theme_label_panel.ColumnCount = 4
        theme_label_panel.RowCount = 1
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 85))  # Empty space under button
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 35))  # Larger spacer to match button row
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 50))  # Text spans both toggle columns
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))  # Fill remaining

        # Switch theme label
        theme_label = WinForms.Label()
        theme_label.Text = "Switch theme"
        theme_label.Font = Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Regular)  # Larger font, regular weight
        theme_label.ForeColor = Drawing.Color.FromArgb(156, 163, 175)  # Light gray
        theme_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Match panel background
        theme_label.TextAlign = Drawing.ContentAlignment.TopCenter
        theme_label.Dock = WinForms.DockStyle.Fill

        # Add controls to panels
        button_row_panel.Controls.Add(aboutme_btn, 0, 0)
        # Skip column 1 (spacer)
        button_row_panel.Controls.Add(self.dark_theme_btn, 2, 0)
        button_row_panel.Controls.Add(self.light_theme_btn, 3, 0)

        theme_label_panel.Controls.Add(theme_label, 2, 0)

        # Add panels to main about panel
        aboutme_panel.Controls.Add(button_row_panel, 0, 0)
        aboutme_panel.Controls.Add(theme_label_panel, 0, 1)

        # Add to title panel
        title_panel.Controls.Add(aboutme_panel, 0, 0)  # About Me button
        title_panel.SetRowSpan(aboutme_panel, 2)
        title_panel.Controls.Add(title_center_panel, 1, 0)
        title_panel.SetRowSpan(title_center_panel, 2)  # Span both rows
        title_panel.Controls.Add(saveload_panel, 2, 0)
        title_panel.SetRowSpan(saveload_panel, 2)  # Span both rows







        
        # Scope selection panel (Active View / Whole Model)
        scope_panel = WinForms.TableLayoutPanel()
        scope_panel.Dock = WinForms.DockStyle.Fill
        scope_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        scope_panel.ColumnCount = 3
        scope_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 40))
        scope_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 30))
        scope_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 30))
        scope_panel.Padding = WinForms.Padding(10, 5, 10, 5)
        
        scope_label = WinForms.Label()
        scope_label.Text = "📊 Elements Scope:"
        scope_label.TextAlign = Drawing.ContentAlignment.MiddleLeft
        scope_label.Dock = WinForms.DockStyle.Fill
        scope_label.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        scope_label.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text for dark theme
        
        self.active_view_btn = WinForms.Button()
        self.active_view_btn.Text = "🔍 Active View Elements"
        self.active_view_btn.Dock = WinForms.DockStyle.Fill
        self.active_view_btn.Click += self.SwitchToActiveView
        self.active_view_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green (selected)
        self.active_view_btn.ForeColor = Drawing.Color.White
        self.active_view_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        self.active_view_btn.FlatStyle = WinForms.FlatStyle.Flat
        self.active_view_btn.FlatAppearance.BorderSize = 1
        self.active_view_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        self.whole_model_btn = WinForms.Button()
        self.whole_model_btn.Text = "🏢 Whole Model Elements"
        self.whole_model_btn.Dock = WinForms.DockStyle.Fill
        self.whole_model_btn.Click += self.SwitchToWholeModel
        self.whole_model_btn.BackColor = Drawing.Color.FromArgb(71, 85, 105)  # Modern gray (unselected)
        self.whole_model_btn.ForeColor = Drawing.Color.White
        self.whole_model_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        self.whole_model_btn.FlatStyle = WinForms.FlatStyle.Flat
        self.whole_model_btn.FlatAppearance.BorderSize = 1
        self.whole_model_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        scope_panel.Controls.Add(scope_label, 0, 0)
        scope_panel.Controls.Add(self.active_view_btn, 1, 0)
        scope_panel.Controls.Add(self.whole_model_btn, 2, 0)
        
        # Search panel - Enhanced for dual search - Dark Theme
        search_panel = WinForms.TableLayoutPanel()
        search_panel.Dock = WinForms.DockStyle.Fill
        search_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        search_panel.ColumnCount = 5
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 100))
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50))  # Categories search
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 60))  # Parameters search
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50))
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 60))
        
        search_label = WinForms.Label()
        search_label.Text = "🔍 Filter:"
        search_label.TextAlign = Drawing.ContentAlignment.MiddleLeft
        search_label.Dock = WinForms.DockStyle.Fill
        search_label.Font = Drawing.Font("Segoe UI", 9)
        search_label.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text for dark theme
        
        # Initialize search textboxes properly
        self.search_categories_textbox = WinForms.TextBox()
        self.search_categories_textbox.Dock = WinForms.DockStyle.Fill
        self.search_categories_textbox.Font = Drawing.Font("Segoe UI", 10)
        self.search_categories_textbox.Text = "Search categories..."
        self.search_categories_textbox.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        self.search_categories_textbox.BackColor = Drawing.Color.FromArgb(71, 85, 105)  # Dark input background
        
        self.search_parameters_textbox = WinForms.TextBox()
        self.search_parameters_textbox.Dock = WinForms.DockStyle.Fill
        self.search_parameters_textbox.Font = Drawing.Font("Segoe UI", 10)
        self.search_parameters_textbox.Text = "Search parameters..."
        self.search_parameters_textbox.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        self.search_parameters_textbox.BackColor = Drawing.Color.FromArgb(71, 85, 105)  # Dark input background
        
        # Add event handlers after textboxes are created
        self.search_categories_textbox.TextChanged += self.OnCategorySearchTextChanged
        self.search_categories_textbox.Enter += self.OnCategorySearchEnter
        self.search_categories_textbox.Leave += self.OnCategorySearchLeave
        
        self.search_parameters_textbox.TextChanged += self.OnParameterSearchTextChanged
        self.search_parameters_textbox.Enter += self.OnParameterSearchEnter
        self.search_parameters_textbox.Leave += self.OnParameterSearchLeave
        
        clear_cat_btn = WinForms.Button()
        clear_cat_btn.Text = "Clear"
        clear_cat_btn.Size = Drawing.Size(50, 25)  # Match textbox height approximately
        clear_cat_btn.Anchor = WinForms.AnchorStyles.None
        clear_cat_btn.Click += self.ClearCategorySearch
        clear_cat_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)  # Modern blue-gray
        clear_cat_btn.ForeColor = Drawing.Color.White
        clear_cat_btn.Font = Drawing.Font("Segoe UI", 8)
        clear_cat_btn.FlatStyle = WinForms.FlatStyle.Flat
        clear_cat_btn.FlatAppearance.BorderSize = 1
        clear_cat_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        clear_param_btn = WinForms.Button()
        clear_param_btn.Text = "Clear"
        clear_param_btn.Size = Drawing.Size(50, 25)  # Match textbox height approximately
        clear_param_btn.Anchor = WinForms.AnchorStyles.None
        clear_param_btn.Click += self.ClearParameterSearch
        clear_param_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)  # Modern blue-gray
        clear_param_btn.ForeColor = Drawing.Color.White
        clear_param_btn.Font = Drawing.Font("Segoe UI", 8)
        clear_param_btn.FlatStyle = WinForms.FlatStyle.Flat
        clear_param_btn.FlatAppearance.BorderSize = 1
        clear_param_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        search_panel.Controls.Add(search_label, 0, 0)
        search_panel.Controls.Add(self.search_categories_textbox, 1, 0)
        search_panel.Controls.Add(self.search_parameters_textbox, 2, 0)
        search_panel.Controls.Add(clear_cat_btn, 2, 0)
        search_panel.Controls.Add(clear_param_btn, 4, 0)
        
        # Info label
        self.info_label = WinForms.Label()
        self.info_label.Text = "Loading categories and parameters from active view..."
        self.info_label.Dock = WinForms.DockStyle.Fill
        self.info_label.TextAlign = Drawing.ContentAlignment.MiddleLeft
        self.info_label.Font = Drawing.Font("Segoe UI", 11, Drawing.FontStyle.Bold)
        self.info_label.ForeColor = Drawing.Color.FromArgb(59, 130, 246)  # Modern blue for dark theme
        self.info_label.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        
        # Categories and Parameters panel - Split view
        content_panel = WinForms.TableLayoutPanel()
        content_panel.Dock = WinForms.DockStyle.Fill
        content_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        content_panel.ColumnCount = 2
        content_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 45))  # Categories
        content_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 55))  # Parameters
        
        content_panel.Padding = WinForms.Padding(5)
        
        # Initialize checklists first (needed for event handlers)
        self.categories_checklist = WinForms.CheckedListBox()
        self.categories_checklist.Dock = WinForms.DockStyle.Fill
        self.categories_checklist.CheckOnClick = True
        self.categories_checklist.Font = Drawing.Font("Segoe UI", 10.5)
        self.categories_checklist.BackColor = Drawing.Color.FromArgb(71, 85, 105)  # Dark list background
        self.categories_checklist.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.categories_checklist.ItemCheck += self.OnCategoryChecked
        
        self.parameters_checklist = WinForms.CheckedListBox()
        self.parameters_checklist.Dock = WinForms.DockStyle.Fill
        self.parameters_checklist.CheckOnClick = True
        self.parameters_checklist.Font = Drawing.Font("Segoe UI", 10.5)
        self.parameters_checklist.BackColor = Drawing.Color.FromArgb(71, 85, 105)  # Dark list background
        self.parameters_checklist.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.parameters_checklist.ItemCheck += self.OnParameterChecked
        
        # Categories section
        categories_group = WinForms.GroupBox()
        categories_group.Text = "📁 Categories"
        categories_group.Dock = WinForms.DockStyle.Fill
        categories_group.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        categories_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        categories_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Slightly lighter dark
        
        categories_layout = WinForms.TableLayoutPanel()
        categories_layout.Dock = WinForms.DockStyle.Fill
        categories_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        categories_layout.RowCount = 2
        categories_layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))
        categories_layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 25))
        
        self.category_count_label = WinForms.Label()
        self.category_count_label.Text = "Selected: 0 categories"
        self.category_count_label.TextAlign = Drawing.ContentAlignment.MiddleCenter
        self.category_count_label.Dock = WinForms.DockStyle.Fill
        self.category_count_label.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        self.category_count_label.ForeColor = Drawing.Color.FromArgb(239, 68, 68)  # Modern red
        self.category_count_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        categories_layout.Controls.Add(self.categories_checklist, 0, 0)
        categories_layout.Controls.Add(self.category_count_label, 0, 1)
        categories_group.Controls.Add(categories_layout)
        
        # Parameters section
        parameters_group = WinForms.GroupBox()
        parameters_group.Text = "🔧 Parameters"
        parameters_group.Dock = WinForms.DockStyle.Fill
        parameters_group.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        parameters_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        parameters_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Slightly lighter dark
        
        parameters_layout = WinForms.TableLayoutPanel()
        parameters_layout.Dock = WinForms.DockStyle.Fill
        parameters_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        parameters_layout.RowCount = 2
        parameters_layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))
        parameters_layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 25))
        
        self.parameter_count_label = WinForms.Label()
        self.parameter_count_label.Text = "Selected: 0 parameters"
        self.parameter_count_label.TextAlign = Drawing.ContentAlignment.MiddleCenter
        self.parameter_count_label.Dock = WinForms.DockStyle.Fill
        self.parameter_count_label.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        self.parameter_count_label.ForeColor = Drawing.Color.FromArgb(59, 130, 246)  # Modern blue
        self.parameter_count_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        parameters_layout.Controls.Add(self.parameters_checklist, 0, 0)
        parameters_layout.Controls.Add(self.parameter_count_label, 0, 1)
        parameters_group.Controls.Add(parameters_layout)
        
        content_panel.Controls.Add(categories_group, 0, 0)
        content_panel.Controls.Add(parameters_group, 1, 0)
        
        # Selection buttons panel - Updated for dual control
        selection_panel = WinForms.TableLayoutPanel()
        selection_panel.Dock = WinForms.DockStyle.Fill
        selection_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        selection_panel.ColumnCount = 6
        selection_panel.Padding = WinForms.Padding(5)
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        
        # Category selection buttons
        select_all_cats_btn = WinForms.Button()
        select_all_cats_btn.Text = "✓ All Categories"
        select_all_cats_btn.Size = Drawing.Size(120, 35)
        select_all_cats_btn.Anchor = WinForms.AnchorStyles.None
        select_all_cats_btn.Click += self.SelectAllCategories
        select_all_cats_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green
        select_all_cats_btn.ForeColor = Drawing.Color.White
        select_all_cats_btn.Font = Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Bold)
        select_all_cats_btn.FlatStyle = WinForms.FlatStyle.Flat
        select_all_cats_btn.FlatAppearance.BorderSize = 1
        select_all_cats_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        deselect_all_cats_btn = WinForms.Button()
        deselect_all_cats_btn.Text = "✗ Clear Categories"
        deselect_all_cats_btn.Size = Drawing.Size(120, 35)
        deselect_all_cats_btn.Anchor = WinForms.AnchorStyles.None
        deselect_all_cats_btn.Click += self.DeselectAllCategories
        deselect_all_cats_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)  # Modern blue-gray
        deselect_all_cats_btn.ForeColor = Drawing.Color.White
        deselect_all_cats_btn.Font = Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Bold)
        deselect_all_cats_btn.FlatStyle = WinForms.FlatStyle.Flat
        deselect_all_cats_btn.FlatAppearance.BorderSize = 1
        deselect_all_cats_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        select_filtered_cats_btn = WinForms.Button()
        select_filtered_cats_btn.Text = "🎯 Filter Categories"
        select_filtered_cats_btn.Size = Drawing.Size(120, 35)
        select_filtered_cats_btn.Anchor = WinForms.AnchorStyles.None
        select_filtered_cats_btn.Click += self.SelectFilteredCategories
        select_filtered_cats_btn.BackColor = Drawing.Color.FromArgb(59, 130, 246)  # Modern blue
        select_filtered_cats_btn.ForeColor = Drawing.Color.White
        select_filtered_cats_btn.Font = Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Bold)
        select_filtered_cats_btn.FlatStyle = WinForms.FlatStyle.Flat
        select_filtered_cats_btn.FlatAppearance.BorderSize = 1
        select_filtered_cats_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        # Parameter selection buttons
        select_all_params_btn = WinForms.Button()
        select_all_params_btn.Text = "✓ All Parameters"
        select_all_params_btn.Size = Drawing.Size(120, 35)
        select_all_params_btn.Anchor = WinForms.AnchorStyles.None
        select_all_params_btn.Click += self.SelectAllParameters
        select_all_params_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green
        select_all_params_btn.ForeColor = Drawing.Color.White
        select_all_params_btn.Font = Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Bold)
        select_all_params_btn.FlatStyle = WinForms.FlatStyle.Flat
        select_all_params_btn.FlatAppearance.BorderSize = 1
        select_all_params_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        deselect_all_params_btn = WinForms.Button()
        deselect_all_params_btn.Text = "✗ Clear Parameters"
        deselect_all_params_btn.Size = Drawing.Size(120, 35)
        deselect_all_params_btn.Anchor = WinForms.AnchorStyles.None
        deselect_all_params_btn.Click += self.DeselectAllParameters
        deselect_all_params_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)  # Modern blue-gray
        deselect_all_params_btn.ForeColor = Drawing.Color.White
        deselect_all_params_btn.Font = Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Bold)
        deselect_all_params_btn.FlatStyle = WinForms.FlatStyle.Flat
        deselect_all_params_btn.FlatAppearance.BorderSize = 1
        deselect_all_params_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        select_filtered_params_btn = WinForms.Button()
        select_filtered_params_btn.Text = "🎯 Filter Parameters"
        select_filtered_params_btn.Size = Drawing.Size(120, 35)
        select_filtered_params_btn.Anchor = WinForms.AnchorStyles.None
        select_filtered_params_btn.Click += self.SelectFilteredParameters
        select_filtered_params_btn.BackColor = Drawing.Color.FromArgb(59, 130, 246)  # Modern blue
        select_filtered_params_btn.ForeColor = Drawing.Color.White
        select_filtered_params_btn.Font = Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Bold)
        select_filtered_params_btn.FlatStyle = WinForms.FlatStyle.Flat
        select_filtered_params_btn.FlatAppearance.BorderSize = 1
        select_filtered_params_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        selection_panel.Controls.Add(select_all_cats_btn, 0, 0)
        selection_panel.Controls.Add(deselect_all_cats_btn, 1, 0)
        selection_panel.Controls.Add(select_filtered_cats_btn, 2, 0)
        selection_panel.Controls.Add(select_all_params_btn, 3, 0)
        selection_panel.Controls.Add(deselect_all_params_btn, 4, 0)
        selection_panel.Controls.Add(select_filtered_params_btn, 5, 0)
        
        # Action buttons panel - Updated layout
        action_panel = WinForms.TableLayoutPanel()
        action_panel.Dock = WinForms.DockStyle.Fill
        action_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        action_panel.RowCount = 3
        action_panel.ColumnCount = 3
        action_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 35))  # First row
        action_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 35))  # Second row
        action_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 25))  # Progress bar
        action_panel.Padding = WinForms.Padding(10, 5, 10, 5)
        
        # Standard button size
        button_size = Drawing.Size(200, 30)
        
        # Export button (main function)
        export_btn = WinForms.Button()
        export_btn.Text = "📊 Export to Excel"
        export_btn.Size = button_size
        export_btn.Anchor = WinForms.AnchorStyles.None
        export_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green
        export_btn.ForeColor = Drawing.Color.White
        export_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        export_btn.FlatStyle = WinForms.FlatStyle.Flat
        export_btn.FlatAppearance.BorderSize = 1
        export_btn.FlatAppearance.BorderColor = Drawing.Color.White
        export_btn.Click += self.ExportByCategory
        
        # Import button
        import_btn = WinForms.Button()
        import_btn.Text = "📥 Import from Excel"
        import_btn.Size = button_size
        import_btn.Anchor = WinForms.AnchorStyles.None
        import_btn.BackColor = Drawing.Color.FromArgb(245, 158, 11)  # Modern amber
        import_btn.ForeColor = Drawing.Color.White
        import_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        import_btn.FlatStyle = WinForms.FlatStyle.Flat
        import_btn.FlatAppearance.BorderSize = 1
        import_btn.FlatAppearance.BorderColor = Drawing.Color.White
        import_btn.Click += self.ImportFromExcel
        
        # Refresh button
        refresh_btn = WinForms.Button()
        refresh_btn.Text = "🔄 Refresh Scope"
        refresh_btn.Size = button_size
        refresh_btn.Anchor = WinForms.AnchorStyles.None
        refresh_btn.BackColor = Drawing.Color.FromArgb(20, 184, 166)  # Modern teal
        refresh_btn.ForeColor = Drawing.Color.White
        refresh_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        refresh_btn.FlatStyle = WinForms.FlatStyle.Flat
        refresh_btn.FlatAppearance.BorderSize = 1
        refresh_btn.FlatAppearance.BorderColor = Drawing.Color.White
        refresh_btn.Click += self.RefreshScope
        
        # Settings button
        settings_btn = WinForms.Button()
        settings_btn.Text = "⚙️ Settings"
        settings_btn.Size = button_size
        settings_btn.Anchor = WinForms.AnchorStyles.None
        settings_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)  # Modern blue-gray
        settings_btn.ForeColor = Drawing.Color.White
        settings_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        settings_btn.FlatStyle = WinForms.FlatStyle.Flat
        settings_btn.FlatAppearance.BorderSize = 1
        settings_btn.FlatAppearance.BorderColor = Drawing.Color.White
        settings_btn.Click += self.ShowSettings
        
        # Help button
        help_btn = WinForms.Button()
        help_btn.Text = "❓ Help"
        help_btn.Size = button_size
        help_btn.Anchor = WinForms.AnchorStyles.None
        help_btn.BackColor = Drawing.Color.FromArgb(20, 184, 166)  # Modern teal
        help_btn.ForeColor = Drawing.Color.White
        help_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        help_btn.FlatStyle = WinForms.FlatStyle.Flat
        help_btn.FlatAppearance.BorderSize = 1
        help_btn.FlatAppearance.BorderColor = Drawing.Color.White
        help_btn.Click += self.ShowHelp
        
        # Close button
        close_btn = WinForms.Button()
        close_btn.Text = "❌ Close"
        close_btn.Size = button_size
        close_btn.Anchor = WinForms.AnchorStyles.None
        close_btn.BackColor = Drawing.Color.FromArgb(239, 68, 68)  # Modern red
        close_btn.ForeColor = Drawing.Color.White
        close_btn.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        close_btn.FlatStyle = WinForms.FlatStyle.Flat
        close_btn.FlatAppearance.BorderSize = 1
        close_btn.FlatAppearance.BorderColor = Drawing.Color.White
        close_btn.Click += self.CloseForm
        
        # Progress bar
        self.progress_bar = WinForms.ProgressBar()
        self.progress_bar.Dock = WinForms.DockStyle.Fill
        self.progress_bar.Visible = False
        
        # Status label
        self.status_label = WinForms.Label()
        self.status_label.Text = "Ready"
        self.status_label.Dock = WinForms.DockStyle.Fill
        self.status_label.TextAlign = Drawing.ContentAlignment.MiddleLeft
        self.status_label.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)  # Increased size and bold for better visibility
        self.status_label.ForeColor = Drawing.Color.FromArgb(156, 163, 175)  # Light gray for dark theme
        self.status_label.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        
        action_panel.Controls.Add(export_btn, 0, 0)
        action_panel.Controls.Add(import_btn, 1, 0)
        action_panel.Controls.Add(refresh_btn, 2, 0)
        action_panel.Controls.Add(settings_btn, 0, 1)
        action_panel.Controls.Add(help_btn, 1, 1)
        action_panel.Controls.Add(close_btn, 2, 1)
        action_panel.Controls.Add(self.progress_bar, 0, 2)
        action_panel.SetColumnSpan(self.progress_bar, 2)
        action_panel.Controls.Add(self.status_label, 2, 2)
        
        # Add controls to main panel
        main_panel.Controls.Add(title_panel, 0, 0)
        main_panel.Controls.Add(scope_panel, 0, 1)
        main_panel.Controls.Add(search_panel, 0, 2)
        main_panel.Controls.Add(self.info_label, 0, 3)
        main_panel.Controls.Add(content_panel, 0, 4)  # Split categories/parameters
        main_panel.Controls.Add(selection_panel, 0, 5)
        main_panel.Controls.Add(action_panel, 0, 6)
        
        self.Controls.Add(main_panel)
    
    def SwitchToActiveView(self, sender, e):
        """Switch to Active View Elements mode"""
        if self.is_processing:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
            
        self.current_scope = "ActiveView"
        self.active_view_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green (selected)
        self.whole_model_btn.BackColor = Drawing.Color.FromArgb(71, 85, 105)  # Modern gray (unselected)
        self.LoadDataFromCurrentScope()
    
    def SwitchToWholeModel(self, sender, e):
        """Switch to Whole Model Elements mode"""
        if self.is_processing:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
        
        # Confirm switch to whole model (can be slow)
        result = WinForms.MessageBox.Show(
            "⚠️ Switching to Whole Model Elements will analyze ALL elements in the model.\n\n" +
            "This may take some time for large models. Continue?",
            "DynaNex - Confirm Scope Change",
            WinForms.MessageBoxButtons.YesNo,
            WinForms.MessageBoxIcon.Question
        )
        
        if result == DialogResult.Yes:
            self.current_scope = "WholeModel"
            self.active_view_btn.BackColor = Drawing.Color.FromArgb(71, 85, 105)  # Modern gray (unselected)
            self.whole_model_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green (selected)
            self.LoadDataFromCurrentScope()
    
    def LoadDataFromCurrentScope(self):
        """Load categories and parameters from current scope - ENHANCED"""
        try:
            # Check if UI components are initialized
            if not self.categories_checklist or not self.parameters_checklist:
                return
                
            if self.current_scope == "ActiveView":
                self.ShowProgress(True, "Analyzing active view elements...")
                scope_text = "active view"
            else:
                self.ShowProgress(True, "Analyzing whole model elements...")
                scope_text = "whole model"
            
            self.all_parameters = []
            self.all_categories = []
            param_names = set()
            category_names = set()
            
            # Get elements based on current scope - IMPROVED TO CAPTURE ALL ELEMENTS
            if self.current_scope == "ActiveView":
                current_view = self.uidoc.ActiveView
                collector = FilteredElementCollector(self.doc, current_view.Id)
                # Get both model elements AND element types to ensure we catch everything
                model_elements = collector.WhereElementIsNotElementType().ToElements()
                
                # Also get element types that might be in the view
                type_collector = FilteredElementCollector(self.doc, current_view.Id)
                type_elements = type_collector.WhereElementIsElementType().ToElements()
                
                # Combine both lists
                elements = list(model_elements) + list(type_elements)
            else:  # WholeModel
                collector = FilteredElementCollector(self.doc)
                model_elements = collector.WhereElementIsNotElementType().ToElements()
                
                type_collector = FilteredElementCollector(self.doc)
                type_elements = type_collector.WhereElementIsElementType().ToElements()
                
                elements = list(model_elements) + list(type_elements)
            
            # Store total element count for display
            self.total_elements_count = len(elements)
            
            # For category detection, use ALL elements (don't sample for categories)
            # but sample for parameters to maintain performance
            self.ShowProgress(True, "Extracting categories from all elements...")
            
            # Extract categories from ALL elements to ensure we don't miss any
            for i, element in enumerate(elements):
                if element and element.Category:
                    # Get both main category and subcategory
                    category = element.Category
                    if category and category.Name:
                        category_names.add(category.Name.strip())
                        
                        # Also add parent category if it exists
                        if category.Parent:
                            category_names.add(category.Parent.Name.strip())
                
                # Update progress every 100 elements for categories
                if i % 100 == 0:
                    progress = int((i * 50) / len(elements))  # 50% for category extraction
                    self.ShowProgress(True, "Analyzing categories... " + str(progress) + "%")
                    WinForms.Application.DoEvents()
            
            # Sample elements for parameter extraction (for performance)
            total_elements = len(elements)
            if self.current_scope == "ActiveView":
                sample_size = min(150, total_elements)
            else:
                sample_size = min(300, total_elements)
            
            step = max(1, total_elements // sample_size)
            sample_elements = []
            for i in range(0, total_elements, step):
                sample_elements.append(elements[i])
                if len(sample_elements) >= sample_size:
                    break
            
            self.ShowProgress(True, "Extracting parameters from " + str(len(sample_elements)) + " sample elements...")
            
            # Extract parameters from sample elements
            for i, element in enumerate(sample_elements):
                if element and element.Parameters:
                    for param in element.Parameters:
                        if param and param.Definition and param.Definition.Name:
                            param_name = param.Definition.Name
                            if param_name and len(param_name.strip()) > 0:
                                param_names.add(param_name.strip())
                
                # Update progress
                if i % 10 == 0:
                    progress = 50 + int((i * 50) / len(sample_elements))  # 50-100% for parameters
                    self.ShowProgress(True, "Analyzing parameters... " + str(progress) + "%")
                    WinForms.Application.DoEvents()
            
            # Convert to sorted lists
            self.all_categories = sorted(list(category_names))
            self.all_parameters = sorted(list(param_names))
            self.filtered_categories = self.all_categories[:]
            self.filtered_parameters = self.all_parameters[:]
            
            # Update UI
            self.UpdateCategoriesList()
            self.UpdateParametersList()
            self.UpdateInfoLabel()
            self.ShowProgress(False)
            
        except Exception as e:
            self.ShowProgress(False)
            self.info_label.Text = "Error loading data: " + str(e)
            WinForms.MessageBox.Show("Error loading data: " + str(e), "DynaNex - Error")
    
    # Placeholder text handlers for search boxes
    def OnCategorySearchEnter(self, sender, e):
        """Handle category search box focus enter"""
        try:
            if sender and hasattr(sender, 'Text'):
                if sender.Text == "Search categories...":
                    sender.Text = ""
                    sender.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text for dark theme
        except:
            pass
    
    def OnCategorySearchLeave(self, sender, e):
        """Handle category search box focus leave"""
        try:
            if sender and hasattr(sender, 'Text'):
                if sender.Text == "":
                    sender.Text = "Search categories..."
                    sender.ForeColor = Drawing.Color.FromArgb(156, 163, 175)  # Modern gray for dark theme
        except:
            pass
    
    def OnParameterSearchEnter(self, sender, e):
        """Handle parameter search box focus enter"""
        try:
            if sender and hasattr(sender, 'Text'):
                if sender.Text == "Search parameters...":
                    sender.Text = ""
                    sender.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text for dark theme
        except:
            pass
    
    def OnParameterSearchLeave(self, sender, e):
        """Handle parameter search box focus leave"""
        try:
            if sender and hasattr(sender, 'Text'):
                if sender.Text == "":
                    sender.Text = "Search parameters..."
                    sender.ForeColor = Drawing.Color.FromArgb(156, 163, 175)  # Modern gray for dark theme
        except:
            pass
    
    # Category-related methods
    def UpdateCategoriesList(self):
        """Update the categories checklist with filtered results while preserving global selections"""
        try:
            if self.categories_checklist and hasattr(self.categories_checklist, 'Items'):
                # Update global selection state from current UI
                for i in range(self.categories_checklist.Items.Count):
                    item_text = str(self.categories_checklist.Items[i])
                    if self.categories_checklist.GetItemChecked(i):
                        self.selected_categories.add(item_text)
                    elif item_text in self.selected_categories:
                        self.selected_categories.remove(item_text)
                
                # Clear and repopulate
                self.categories_checklist.Items.Clear()
                for category in self.filtered_categories:
                    index = self.categories_checklist.Items.Add(category)
                    # Restore selection from global state
                    if category in self.selected_categories:
                        self.categories_checklist.SetItemChecked(index, True)
        except:
            pass
    
    def OnCategorySearchTextChanged(self, sender, e):
        """Handle category search text changes"""
        try:
            if not self.search_categories_textbox:
                return
            search_text = self.search_categories_textbox.Text.lower()
            # Ignore placeholder text
            if search_text == "search categories..." or not search_text:
                self.filtered_categories = self.all_categories[:]
            else:
                self.filtered_categories = []
                for category in self.all_categories:
                    if search_text in category.lower():
                        self.filtered_categories.append(category)
            
            self.UpdateCategoriesList()
            self.UpdateInfoLabel()
        except:
            pass
    
    def ClearCategorySearch(self, sender, e):
        """Clear category search text"""
        try:
            if self.search_categories_textbox:
                self.search_categories_textbox.Text = "Search categories..."
                self.search_categories_textbox.ForeColor = Drawing.Color.Gray
        except:
            pass
    
    def OnCategoryChecked(self, sender, e):
        """Update category selection count and global state when categories are checked/unchecked"""
        try:
            # Update global selection state immediately
            if e.Index >= 0 and e.Index < self.categories_checklist.Items.Count:
                item_text = str(self.categories_checklist.Items[e.Index])
                if e.NewValue == System.Windows.Forms.CheckState.Checked:
                    self.selected_categories.add(item_text)
                else:
                    self.selected_categories.discard(item_text)
        except:
            pass
        
        self.BeginInvoke(System.Action(self.UpdateCategoryCount))
    
    def UpdateCategoryCount(self):
        """Update the category selection count label"""
        count = 0
        for i in range(self.categories_checklist.Items.Count):
            if self.categories_checklist.GetItemChecked(i):
                count += 1
        self.category_count_label.Text = "Selected: " + str(count) + " categories"
    
    def SelectAllCategories(self, sender, e):
        """Select all categories"""
        for i in range(self.categories_checklist.Items.Count):
            self.categories_checklist.SetItemChecked(i, True)
        self.UpdateCategoryCount()
    
    def DeselectAllCategories(self, sender, e):
        """Deselect all categories"""
        for i in range(self.categories_checklist.Items.Count):
            self.categories_checklist.SetItemChecked(i, False)
        self.UpdateCategoryCount()
    
    def SelectFilteredCategories(self, sender, e):
        """Select all currently filtered/visible categories"""
        for i in range(self.categories_checklist.Items.Count):
            self.categories_checklist.SetItemChecked(i, True)
        self.UpdateCategoryCount()
    
    def GetSelectedCategories(self):
        """Get list of selected categories"""
        selected = []
        for i in range(self.categories_checklist.Items.Count):
            if self.categories_checklist.GetItemChecked(i):
                selected.append(str(self.categories_checklist.Items[i]))
        return selected
    
    # Parameter-related methods (updated names for consistency)
    def OnParameterSearchTextChanged(self, sender, e):
        """Handle parameter search text changes"""
        try:
            if not self.search_parameters_textbox:
                return
            search_text = self.search_parameters_textbox.Text.lower()
            # Ignore placeholder text
            if search_text == "search parameters..." or not search_text:
                self.filtered_parameters = self.all_parameters[:]
            else:
                self.filtered_parameters = []
                for param in self.all_parameters:
                    if search_text in param.lower():
                        self.filtered_parameters.append(param)
            
            self.UpdateParametersList()
            self.UpdateInfoLabel()
        except:
            pass
    
    def ClearParameterSearch(self, sender, e):
        """Clear parameter search text"""
        try:
            if self.search_parameters_textbox:
                self.search_parameters_textbox.Text = "Search parameters..."
                self.search_parameters_textbox.ForeColor = Drawing.Color.FromArgb(156, 163, 175)  # Modern gray for dark theme
                # Force refresh of parameters list
                self.filtered_parameters = self.all_parameters[:]
                self.UpdateParametersList()
                self.UpdateInfoLabel()
        except:
            pass
    
    def SelectAllParameters(self, sender, e):
        """Select all parameters"""
        for i in range(self.parameters_checklist.Items.Count):
            self.parameters_checklist.SetItemChecked(i, True)
        self.UpdateParameterCount()
    
    def DeselectAllParameters(self, sender, e):
        """Deselect all parameters"""
        for i in range(self.parameters_checklist.Items.Count):
            self.parameters_checklist.SetItemChecked(i, False)
        self.UpdateParameterCount()
    
    def SelectFilteredParameters(self, sender, e):
        """Select all currently filtered/visible parameters"""
        for i in range(self.parameters_checklist.Items.Count):
            self.parameters_checklist.SetItemChecked(i, True)
        self.UpdateParameterCount()
    
    def GetMainCategory(self, element):
        """Get main Revit category, filter out subcategories"""
        if not element or not element.Category:
            return None
            
        category = element.Category
        parent = category.Parent
        
        # If it has a parent, use the parent (main category)
        if parent:
            # Special cases for MEP where subcategories might be useful
            mep_categories = [
                "Mechanical Equipment", "Electrical Equipment", "Plumbing Fixtures", 
                "Air Terminals", "Sprinklers", "Electrical Fixtures", "Communication Devices"
            ]
            
            # For MEP, check if we should keep the subcategory
            if parent.Name in mep_categories:
                return category  # Keep subcategory for MEP
            else:
                return parent   # Use main category for others
        else:
            return category     # Already main category
    
    def ExportByCategory(self, sender, e):
        """Export elements by selected categories to Excel - ENHANCED WITH CATEGORY FILTERING"""
        if self.is_processing:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
            
        selected_categories = self.GetSelectedCategories()
        selected_params = self.GetSelectedParameters()
        
        if len(selected_categories) == 0:
            WinForms.MessageBox.Show("Please select at least one category to export.", "DynaNex - Warning")
            return
            
        if len(selected_params) == 0:
            WinForms.MessageBox.Show("Please select at least one parameter to export.", "DynaNex - Warning")
            return
        
        try:
            self.is_processing = True
            self.ShowProgress(True, "Preparing export...")
            
            # Get elements based on current scope
            if self.current_scope == "ActiveView":
                current_view = self.uidoc.ActiveView
                collector = FilteredElementCollector(self.doc, current_view.Id)
                all_elements = collector.WhereElementIsNotElementType().ToElements()
                scope_name = current_view.Name
            else:  # WholeModel
                collector = FilteredElementCollector(self.doc)
                all_elements = collector.WhereElementIsNotElementType().ToElements()
                scope_name = "WholeModel"
            
            if len(all_elements) == 0:
                WinForms.MessageBox.Show("No elements found in current scope.", "DynaNex - Warning")
                return
            
            # Filter elements by selected categories
            self.ShowProgress(True, "Filtering elements by selected categories...")
            category_elements = {}
            
            for element in all_elements:
                main_category = self.GetMainCategory(element)
                if main_category and main_category.Name in selected_categories:
                    category_name = main_category.Name
                    if category_name not in category_elements:
                        category_elements[category_name] = []
                    category_elements[category_name].append(element)
                
                WinForms.Application.DoEvents()
            
            if len(category_elements) == 0:
                WinForms.MessageBox.Show("No elements found in selected categories.", "DynaNex - Warning")
                return
            
            # Save dialog
            save_dialog = WinForms.SaveFileDialog()
            save_dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            save_dialog.FileName = "DynaNex_Export_" + scope_name.replace(" ", "_") + "_" + System.DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
            save_dialog.Title = "Save DynaNex Export"
            
            if save_dialog.ShowDialog() != DialogResult.OK:
                self.is_processing = False
                self.ShowProgress(False)
                return
            
            # Create Excel with selected category sheets
            self.CreateCategoryBasedExcel(category_elements, selected_params, save_dialog.FileName)
            
        except Exception as e:
            WinForms.MessageBox.Show("Export failed: " + str(e), "DynaNex - Error")
        finally:
            self.is_processing = False
            self.ShowProgress(False)
    
    def CreateCategoryBasedExcel(self, category_elements, parameters, file_path):
        """Create Excel file with separate sheets for each category - OPTIMIZED"""
        excel_app = None
        try:
            self.ShowProgress(True, "Creating Excel application...")
            
            # Create Excel application
            excel_app = Excel.ApplicationClass()
            excel_app.Visible = False
            excel_app.ScreenUpdating = False
            excel_app.DisplayAlerts = False
            
            workbook = excel_app.Workbooks.Add()
            
            # Remove default sheets except first one
            while workbook.Worksheets.Count > 1:
                workbook.Worksheets[workbook.Worksheets.Count].Delete()
            
            first_sheet = True
            total_categories = len(category_elements)
            current_category = 0
            
            # Create sheet for each category
            for category_name, elements in category_elements.items():
                current_category += 1
                self.ShowProgress(True, "Processing " + category_name + " (" + str(current_category) + "/" + str(total_categories) + ")...")
                
                # Create or use worksheet
                if first_sheet:
                    worksheet = workbook.ActiveSheet
                    first_sheet = False
                else:
                    worksheet = workbook.Worksheets.Add()
                
                # Set sheet name (Excel-safe)
                safe_name = self.MakeExcelSafeName(category_name)
                worksheet.Name = safe_name[:31]  # Excel sheet name limit
                
                # Create headers
                headers = ["Element ID", "Category", "Family", "Type", "Level", "Phase"] + parameters
                
                # Set headers with formatting
                for col, header in enumerate(headers):
                    cell = worksheet.Cells[1, col + 1]
                    cell.Value2 = header
                    cell.Font.Bold = True
                    cell.Interior.Color = 15917529  # Light blue
                    cell.Borders.Weight = 1
                
                # Process elements in batches
                row = 2
                batch_size = 25  # Smaller batches for better performance
                processed = 0
                
                for i in range(0, len(elements), batch_size):
                    batch = elements[i:i + batch_size]
                    
                    for element in batch:
                        try:
                            # Basic element info
                            worksheet.Cells[row, 1].Value2 = element.Id.IntegerValue
                            worksheet.Cells[row, 2].Value2 = category_name
                            
                            # Family and type info
                            family_name, type_name = self.GetFamilyTypeInfo(element)
                            worksheet.Cells[row, 3].Value2 = family_name
                            worksheet.Cells[row, 4].Value2 = type_name
                            
                            # Level and phase info
                            level_name = self.GetElementLevel(element)
                            phase_name = self.GetElementPhase(element)
                            worksheet.Cells[row, 5].Value2 = level_name
                            worksheet.Cells[row, 6].Value2 = phase_name
                            
                            # Parameter values
                            for col, param_name in enumerate(parameters):
                                value = self.GetParameterValue(element, param_name)
                                worksheet.Cells[row, 7 + col].Value2 = value
                            
                            row += 1
                            processed += 1
                            
                        except:
                            continue
                    
                    # Update progress and keep UI responsive
                    progress = int((processed * 100) / len(elements))
                    self.ShowProgress(True, category_name + ": " + str(progress) + "% (" + str(processed) + "/" + str(len(elements)) + ")")
                    WinForms.Application.DoEvents()
                
                # Format the sheet
                if row > 2:
                    # Auto-fit columns
                    worksheet.Columns.AutoFit()
                    
                    # Add borders
                    data_range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row - 1, len(headers)]]
                    data_range.Borders.Weight = 1
                    
                    # Add filters
                    header_range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, len(headers)]]
                    header_range.AutoFilter()
                    
                    # Freeze header row
                    worksheet.Range["A2"].Select()
                    excel_app.ActiveWindow.FreezePanes = True
            
            # Save and show Excel
            self.ShowProgress(True, "Saving Excel file...")
            workbook.SaveAs(file_path)
            
            # Make Excel visible
            excel_app.Visible = True
            excel_app.ScreenUpdating = True
            excel_app.DisplayAlerts = True
            
            # Success message
            total_elements = sum(len(elements) for elements in category_elements.values())
            scope_text = "Active View" if self.current_scope == "ActiveView" else "Whole Model"
            WinForms.MessageBox.Show(
                "✅ Export completed successfully!\n\n" +
                "📊 Scope: " + scope_text + "\n" +
                "📂 Selected Categories: " + str(len(category_elements)) + "\n" +
                "📋 Elements: " + str(total_elements) + "\n" +
                "🔧 Parameters: " + str(len(parameters)) + "\n\n" +
                "📁 File: " + file_path,
                "DynaNex - Export Complete"
            )
            
        except Exception as e:
            try:
                if excel_app:
                    excel_app.Visible = True
                    excel_app.ScreenUpdating = True
                    excel_app.DisplayAlerts = True
            except:
                pass
            raise e
    
    def ImportFromExcel(self, sender, e):
        """Import parameters from Excel - ENHANCED with category support"""
        if self.is_processing:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
            
        try:
            # File dialog
            open_dialog = WinForms.OpenFileDialog()
            open_dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
            open_dialog.Title = "Select DynaNex Excel File to Import"
            
            if open_dialog.ShowDialog() != DialogResult.OK:
                return
            
            self.is_processing = True
            self.ShowProgress(True, "Opening Excel file...")
            
            # Open Excel file
            excel_app = Excel.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            workbook = excel_app.Workbooks.Open(open_dialog.FileName)
            
            # Confirm import
            message = "⚠️ IMPORTANT: This will overwrite parameter values in your Revit model!\n\n"
            message += "📊 Excel file: " + open_dialog.FileName + "\n"
            message += "📋 Sheets found: " + str(workbook.Worksheets.Count) + "\n\n"
            message += "🛡️ Make sure you have saved/backed up your Revit model!\n\n"
            message += "Continue with import?"
            
            result = WinForms.MessageBox.Show(message, "DynaNex - Confirm Import", 
                                            WinForms.MessageBoxButtons.YesNo, 
                                            WinForms.MessageBoxIcon.Warning)
            
            if result != DialogResult.Yes:
                workbook.Close(False)
                excel_app.Quit()
                self.is_processing = False
                self.ShowProgress(False)
                return
            
            # Start transaction - CRITICAL FOR REVIT UPDATES
            self.ShowProgress(True, "Starting Revit transaction...")
            TransactionManager.Instance.EnsureInTransaction(self.doc)
            
            try:
                total_updated = 0
                total_errors = 0
                
                # Process each worksheet
                for sheet_index in range(1, workbook.Worksheets.Count + 1):
                    worksheet = workbook.Worksheets[sheet_index]
                    sheet_name = worksheet.Name
                    
                    self.ShowProgress(True, "Importing from: " + sheet_name + " (" + str(sheet_index) + "/" + str(workbook.Worksheets.Count) + ")")
                    
                    updated, errors = self.ImportFromWorksheetFixed(worksheet)
                    total_updated += updated
                    total_errors += errors
                    
                    WinForms.Application.DoEvents()
                
                # COMMIT TRANSACTION - THIS IS CRITICAL!
                TransactionManager.Instance.TransactionTaskDone()
                
                # Close Excel
                workbook.Close(False)
                excel_app.Quit()
                
                # Success message
                message = "✅ Import completed successfully!\n\n"
                message += "📊 Elements updated: " + str(total_updated) + "\n"
                if total_errors > 0:
                    message += "⚠️ Elements with errors: " + str(total_errors) + "\n"
                message += "\n🔄 Parameters have been updated in your Revit model!"
                
                WinForms.MessageBox.Show(message, "DynaNex - Import Complete")
                
            except Exception as transaction_error:
                # ROLLBACK TRANSACTION ON ERROR
                try:
                    TransactionManager.Instance.ForceCloseTransaction()
                except:
                    pass
                
                workbook.Close(False)
                excel_app.Quit()
                raise transaction_error
                
        except Exception as e:
            WinForms.MessageBox.Show("Import failed: " + str(e), "DynaNex - Import Error")
        finally:
            self.is_processing = False
            self.ShowProgress(False)
    
    def ImportFromWorksheetFixed(self, worksheet):
        """FIXED IMPORT - Works with category filtering support"""
        updated_count = 0
        error_count = 0
        
        try:
            # Get worksheet dimensions
            used_range = worksheet.UsedRange
            if not used_range:
                return updated_count, error_count
                
            last_row = used_range.Rows.Count
            last_col = used_range.Columns.Count
            
            if last_row < 2:  # Need at least header + 1 data row
                return updated_count, error_count
            
            # Read all headers at once
            headers = []
            for col in range(1, last_col + 1):
                cell_value = worksheet.Cells[1, col].Value2
                headers.append(str(cell_value) if cell_value else "")
            
            # Validate required columns
            if "Element ID" not in headers:
                return updated_count, error_count
            
            element_id_col = headers.index("Element ID") + 1
            
            # Find parameter columns (exclude standard info columns)
            standard_cols = ["Element ID", "Category", "Family", "Type", "Level", "Phase"]
            param_columns = {}
            
            for i, header in enumerate(headers):
                if header and header not in standard_cols:
                    param_columns[header] = i + 1
            
            if len(param_columns) == 0:
                return updated_count, error_count
            
            # Process each data row
            for row in range(2, last_row + 1):
                try:
                    # Get Element ID
                    element_id_cell = worksheet.Cells[row, element_id_col].Value2
                    if not element_id_cell:
                        continue
                    
                    # Convert to Revit Element ID
                    try:
                        element_id_int = int(float(element_id_cell))  # Handle both int and float
                        element_id = ElementId(element_id_int)
                        element = self.doc.GetElement(element_id)
                        
                        if not element:
                            error_count += 1
                            continue
                            
                    except:
                        error_count += 1
                        continue
                    
                    # Update each parameter for this element
                    element_was_updated = False
                    
                    for param_name, col_index in param_columns.items():
                        try:
                            # Get the cell value
                            cell_value = worksheet.Cells[row, col_index].Value2
                            
                            # Skip empty cells
                            if cell_value is None:
                                continue
                                
                            # Convert cell value to string for processing
                            cell_str = str(cell_value).strip()
                            if cell_str == "":
                                continue
                            
                            # Find the parameter in the element
                            param = element.LookupParameter(param_name)
                            if not param:
                                continue
                                
                            # Skip read-only parameters
                            if param.IsReadOnly:
                                continue
                            
                            # Set parameter value based on its storage type
                            success = self.SetParameterValue(param, cell_value)
                            if success:
                                element_was_updated = True
                                
                        except Exception as param_error:
                            # Continue with other parameters even if one fails
                            continue
                    
                    if element_was_updated:
                        updated_count += 1
                    
                    # Update progress every 10 elements
                    if row % 10 == 2:  # Every 10 elements starting from row 2
                        progress_pct = int(((row - 1) * 100) / (last_row - 1))
                        self.ShowProgress(True, "Updating elements... " + str(progress_pct) + "%")
                        WinForms.Application.DoEvents()
                
                except Exception as row_error:
                    error_count += 1
                    continue
                    
        except Exception as sheet_error:
            # Return what we managed to process
            pass
        
        return updated_count, error_count
    
    def SetParameterValue(self, param, value):
        """Set parameter value with proper type conversion"""
        try:
            if param.StorageType == StorageType.String:
                # String parameters
                param.Set(str(value))
                return True
                
            elif param.StorageType == StorageType.Integer:
                # Integer parameters
                if param.Definition.ParameterType == ParameterType.YesNo:
                    # Yes/No (Boolean) parameters
                    value_str = str(value).lower()
                    if value_str in ['yes', 'true', '1', 'y', 'on']:
                        param.Set(1)
                    else:
                        param.Set(0)
                    return True
                else:
                    # Regular integer parameters
                    int_value = int(float(value))  # Convert via float to handle "1.0" strings
                    param.Set(int_value)
                    return True
                    
            elif param.StorageType == StorageType.Double:
                # Double (decimal) parameters
                double_value = float(value)
                param.Set(double_value)
                return True
                
            elif param.StorageType == StorageType.ElementId:
                # Element reference parameters
                if str(value).isdigit():
                    ref_id = ElementId(int(float(value)))
                    ref_element = self.doc.GetElement(ref_id)
                    if ref_element:
                        param.Set(ref_id)
                        return True
                        
        except Exception as set_error:
            return False
        
        return False
    
    # Helper methods for better organization
    def GetFamilyTypeInfo(self, element):
        """Get family and type information from element"""
        family_name = ""
        type_name = ""
        
        try:
            if hasattr(element, 'Symbol') and element.Symbol is not None:
                family_name = element.Symbol.Family.Name
                type_name = element.Symbol.Name
            elif hasattr(element, 'GetTypeId'):
                type_element = self.doc.GetElement(element.GetTypeId())
                if type_element:
                    type_name = type_element.Name
                    if hasattr(type_element, 'Family'):
                        family_name = type_element.Family.Name
            elif hasattr(element, 'Name') and element.Name:
                type_name = element.Name
        except:
            pass
        
        return family_name, type_name
    
    def GetElementLevel(self, element):
        """Get level information from element"""
        try:
            level_param = element.get_Parameter(BuiltInParameter.LEVEL_PARAM)
            if level_param and level_param.AsElementId() != ElementId.InvalidElementId:
                level_element = self.doc.GetElement(level_param.AsElementId())
                if level_element:
                    return level_element.Name
        except:
            pass
        return ""
    
    def GetElementPhase(self, element):
        """Get phase information from element"""
        try:
            phase_param = element.get_Parameter(BuiltInParameter.PHASE_CREATED)
            if phase_param and phase_param.AsElementId() != ElementId.InvalidElementId:
                phase_element = self.doc.GetElement(phase_param.AsElementId())
                if phase_element:
                    return phase_element.Name
        except:
            pass
        return ""
    
    def GetParameterValue(self, element, param_name):
        """Get parameter value from element with proper formatting"""
        try:
            param = element.LookupParameter(param_name)
            if param is None or not param.HasValue:
                return ""
            
            if param.StorageType == StorageType.String:
                return param.AsString() or ""
            elif param.StorageType == StorageType.Integer:
                if param.Definition.ParameterType == ParameterType.YesNo:
                    return "Yes" if param.AsInteger() == 1 else "No"
                else:
                    return param.AsInteger()
            elif param.StorageType == StorageType.Double:
                # Use display value for length/area parameters when available
                display_value = param.AsValueString()
                if display_value and display_value.strip():
                    return display_value
                else:
                    return param.AsDouble()
            elif param.StorageType == StorageType.ElementId:
                element_id = param.AsElementId()
                if element_id != ElementId.InvalidElementId:
                    ref_element = self.doc.GetElement(element_id)
                    return ref_element.Name if ref_element else str(element_id.IntegerValue)
            
        except:
            pass
        return ""
    
    def MakeExcelSafeName(self, name):
        """Make string safe for Excel sheet name"""
        # Remove invalid characters
        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        safe_name = name
        for char in invalid_chars:
            safe_name = safe_name.replace(char, '_')
        return safe_name.strip()
    
    def RefreshScope(self, sender, e):
        """Refresh categories and parameters from current scope"""
        if self.is_processing:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
        
        self.LoadDataFromCurrentScope()
        scope_text = "Active View" if self.current_scope == "ActiveView" else "Whole Model"
        WinForms.MessageBox.Show("✅ " + scope_text + " refreshed successfully!", "DynaNex - Refresh Complete")
    
    def OpenLinkedInProfile(self, sender, e):
        """Open developer's LinkedIn profile in browser"""
        try:
            linkedin_url = "https://my.linkedin.com/in/nadeemahmed-nbim"
            
            # Show a message with developer info
            message = "🧑‍💻 DynaNex Developer - Nadeem Ahmed\n\n" + \
                     "Creating Advance Plugins with Dynamo Revit\n" + \
                     "Empowering the AEC Industry\n\n" + \
                     "🔗 Opening LinkedIn Profile...\n" + \
                     linkedin_url + "\n\n" + \
                     "Connect with me for custom tools & solutions!"
            
            WinForms.MessageBox.Show(
                message,
                "About DynaNex Developer",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Information
            )
            
            # Open LinkedIn profile in default browser
            import subprocess
            subprocess.Popen(['start', linkedin_url], shell=True)
            
        except Exception as ex:
            # Fallback - show profile info and URL for manual copy
            fallback_message = "🧑‍💻 DynaNex Developer - Nadeem Ahmed\n\n" + \
                              "Professional Revit & Dynamo Developer\n" + \
                              "LinkedIn Profile:\n" + \
                              "https://my.linkedin.com/in/nadeemahmed-nbim\n\n" + \
                              "Please copy the link above to visit my profile."
            
            WinForms.MessageBox.Show(
                fallback_message,
                "About DynaNex Developer",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Information
            )
    
    def SwitchToDarkTheme(self, sender, e):
        """Switch to dark theme"""
        if self.current_theme != "dark":
            self.current_theme = "dark"
            self.ApplyTheme()
            self.UpdateThemeButtonStates()
    
    def SwitchToLightTheme(self, sender, e):
        """Switch to light theme"""
        if self.current_theme != "light":
            self.current_theme = "light"
            self.ApplyTheme()
            self.UpdateThemeButtonStates()
    
    def UpdateThemeButtonStates(self):
        """Update theme button visual states"""
        try:
            if self.current_theme == "dark":
                # Dark theme active
                self.dark_theme_btn.FlatAppearance.BorderSize = 2
                self.dark_theme_btn.FlatAppearance.BorderColor = Drawing.Color.White
                self.light_theme_btn.FlatAppearance.BorderSize = 1
                self.light_theme_btn.FlatAppearance.BorderColor = Drawing.Color.Gray
            else:
                # Light theme active
                self.dark_theme_btn.FlatAppearance.BorderSize = 1
                self.dark_theme_btn.FlatAppearance.BorderColor = Drawing.Color.Gray
                self.light_theme_btn.FlatAppearance.BorderSize = 2
                self.light_theme_btn.FlatAppearance.BorderColor = Drawing.Color.Black
        except:
            pass
    
    def ApplyTheme(self):
        """Apply the selected theme to all UI elements"""
        try:
            if self.current_theme == "dark":
                self.ApplyDarkTheme()
            else:
                self.ApplyLightTheme()
        except Exception as ex:
            WinForms.MessageBox.Show("Error applying theme: " + str(ex), "Theme Error")
    
    def ApplyDarkTheme(self):
        """Apply dark blue theme colors"""
        try:
            # Define dark theme colors
            colors = {
                'main_bg': Drawing.Color.FromArgb(30, 41, 59),      # Dark blue
                'panel_bg': Drawing.Color.FromArgb(51, 65, 85),     # Medium blue  
                'input_bg': Drawing.Color.FromArgb(71, 85, 105),    # Light blue
                'text_light': Drawing.Color.FromArgb(220, 220, 220), # Light text
                'text_medium': Drawing.Color.FromArgb(156, 163, 175) # Medium text
            }
            
            # Apply main form background first
            self.BackColor = colors['main_bg']
            
            # Update all controls recursively
            self.UpdateControlColors(self, colors, True)
            
            # Force refresh the display
            self.Refresh()
            self.Invalidate()
        except:
            pass
        
    def ApplyLightTheme(self):
        """Apply light theme colors"""
        try:
            # Define light theme colors
            colors = {
                'main_bg': Drawing.Color.FromArgb(248, 250, 252),   # Very light blue-gray
                'panel_bg': Drawing.Color.FromArgb(241, 245, 249),  # Light panel
                'input_bg': Drawing.Color.White,                     # White inputs
                'text_light': Drawing.Color.FromArgb(51, 65, 85),   # Dark text
                'text_medium': Drawing.Color.FromArgb(107, 114, 128) # Medium dark text
            }
            
            # Apply main form background first
            self.BackColor = colors['main_bg']
            
            # Update all controls recursively
            self.UpdateControlColors(self, colors, False)
            
            # Force refresh the display
            self.Refresh()
            self.Invalidate()
        except:
            pass
    
    def UpdateControlColors(self, control, colors, is_dark):
        """Recursively update all control colors for uniform appearance"""
        try:
            # Skip theme toggle buttons - they have special colors
            if hasattr(control, 'Name') and ('theme_btn' in str(control.Name) or 
                                           control == self.dark_theme_btn or 
                                           control == self.light_theme_btn):
                return
                
            # Update current control based on its type
            if hasattr(control, 'BackColor'):
                if isinstance(control, WinForms.Form):
                    # Main form
                    control.BackColor = colors['main_bg']
                    
                elif isinstance(control, WinForms.TableLayoutPanel):
                    # All panels use consistent background
                    if control.Parent == self or not hasattr(control, 'Parent'):
                        # Top-level panels
                        control.BackColor = colors['main_bg']
                    else:
                        # Nested panels use panel background
                        control.BackColor = colors['panel_bg']
                        
                elif isinstance(control, WinForms.TextBox) or isinstance(control, WinForms.CheckedListBox):
                    # Input controls
                    control.BackColor = colors['input_bg']
                    control.ForeColor = colors['text_light']
                    
                elif isinstance(control, WinForms.Label):
                    # Labels - check if they should be transparent
                    if str(control.BackColor) != 'Color [Transparent]':
                        control.BackColor = colors['panel_bg']
                    control.ForeColor = colors['text_light']
                    
                elif isinstance(control, WinForms.GroupBox):
                    # Group boxes
                    control.BackColor = colors['panel_bg']
                    control.ForeColor = colors['text_light']
                    
                elif isinstance(control, WinForms.CheckBox):
                    # Checkboxes
                    control.BackColor = colors['panel_bg']
                    control.ForeColor = colors['text_light']
                    
                elif isinstance(control, WinForms.Panel):
                    # Generic panels
                    control.BackColor = colors['panel_bg']
                    
                else:
                    # Any other control with BackColor - use panel background
                    if not isinstance(control, WinForms.Button):  # Don't change button colors
                        control.BackColor = colors['panel_bg']
                        if hasattr(control, 'ForeColor'):
                            control.ForeColor = colors['text_light']
            
            # Recursively update child controls
            if hasattr(control, 'Controls'):
                for child_control in control.Controls:
                    self.UpdateControlColors(child_control, colors, is_dark)
        except Exception as ex:
            # Silent fail to prevent crashes during theme switching
            pass
    
    def ShowHelp(self, sender, e):
        """Show help dialog"""
        help_text = """⚡ DynaNex V1.0 - Help
"Power your Connections"

📋 WHAT IS DYNANEX:
Professional parameter export/import tool for Revit. Export element parameters to Excel by category, edit them, and import back to update your Revit model. Connects Revit with Microsoft Excel seamlessly.

🚀 HOW TO USE:

🎯 ELEMENT SCOPE SELECTION:
   • 🔍 Active View Elements: Only elements visible in current view
   • 🏢 Whole Model Elements: All elements in the entire model
   • Switch between scopes using the buttons at the top

📂 NEW: CATEGORY SELECTION:
   • Select specific categories to export (Walls, Floors, etc.)
   • Only selected categories will be included in export
   • Each category gets its own Excel sheet
   • Filter categories using the search box

1️⃣ EXPORT PROCESS:
   • Choose your element scope (Active View or Whole Model)
   • Select categories you want to export
   • Select parameters you want to work with
   • Click 'Export to Excel' 
   • Excel file created with separate sheets for selected categories
   • Each element has unique Element ID for tracking

2️⃣ EDIT IN EXCEL:
   • Open the exported Excel file
   • Each selected category has its own sheet
   • Modify parameter values as needed
   • ⚠️ DO NOT change Element ID, Category, Family, Type columns
   • Save the Excel file when done

3️⃣ IMPORT PROCESS:
   • Click 'Import from Excel'
   • Select your modified Excel file
   • All parameter values will be updated in Revit
   • Overrides existing values safely

📊 CATEGORY HANDLING:
• Select only categories you need for focused exports
• Exports MAIN Revit categories (not subcategories)
• Example: "Railings" category (not "Top Rail" subcategory)  
• Exception: MEP subcategories are preserved when useful
• Each selected category gets its own Excel sheet

🔍 DUAL FILTERING:
• Category Filter: Search and select specific categories
• Parameter Filter: Search and select specific parameters
• Selection buttons for quick category/parameter selection
• Export only what you need for better performance

⚠️ IMPORTANT NOTES:
• Always backup your Revit model before importing!
• Only modify parameter value columns in Excel
• Element ID column links Excel rows to Revit elements
• Make sure Excel file is closed before importing
• Tool uses Revit transactions for safe updates

💡 PERFORMANCE TIPS:
• Use 'Refresh Scope' to update category/parameter lists
• Search boxes help find categories and parameters quickly
• Active View mode is faster than Whole Model
• Select only needed categories for faster exports
• Whole Model mode samples elements for performance

🏢 DynaNex - "Power your Connections"
Advanced Revit-Excel Parameter Management"""
        
        help_form = WinForms.Form()
        help_form.Text = "DynaNex - Help"
        help_form.Size = Drawing.Size(700, 650)
        help_form.StartPosition = WinForms.FormStartPosition.CenterParent
        help_form.MaximizeBox = False
        help_form.MinimizeBox = False
        
        help_textbox = WinForms.TextBox()
        help_textbox.Multiline = True
        help_textbox.ReadOnly = True
        help_textbox.ScrollBars = WinForms.ScrollBars.Vertical
        help_textbox.Dock = WinForms.DockStyle.Fill
        help_textbox.Font = Drawing.Font("Segoe UI", 9)
        help_textbox.Text = help_text
        
        help_form.Controls.Add(help_textbox)
        help_form.ShowDialog()
    
    # Existing methods (updated for dual control)
    def UpdateParametersList(self):
        """Update the parameters checklist with filtered results while preserving global selections"""
        try:
            if self.parameters_checklist and hasattr(self.parameters_checklist, 'Items'):
                # Update global selection state from current UI
                for i in range(self.parameters_checklist.Items.Count):
                    item_text = str(self.parameters_checklist.Items[i])
                    if self.parameters_checklist.GetItemChecked(i):
                        self.selected_parameters.add(item_text)
                    elif item_text in self.selected_parameters:
                        self.selected_parameters.remove(item_text)
                
                # Clear and repopulate
                self.parameters_checklist.Items.Clear()
                for param in self.filtered_parameters:
                    index = self.parameters_checklist.Items.Add(param)
                    # Restore selection from global state
                    if param in self.selected_parameters:
                        self.parameters_checklist.SetItemChecked(index, True)
        except:
            pass
    
    def UpdateInfoLabel(self):
        """Update the information label with counts"""
        total_cats = len(self.all_categories)
        filtered_cats = len(self.filtered_categories)
        total_params = len(self.all_parameters)
        filtered_params = len(self.filtered_parameters)
        scope_text = "active view" if self.current_scope == "ActiveView" else "whole model"
        
        # Add element count information
        element_count_text = ""
        if hasattr(self, 'total_elements_count'):
            element_count_text = " | 📊 " + str(self.total_elements_count) + " elements in " + scope_text
        
        info_text = "Categories: " + str(filtered_cats)
        if filtered_cats < total_cats:
            info_text += " of " + str(total_cats) + " (filtered)"
        
        info_text += " | Parameters: " + str(filtered_params)
        if filtered_params < total_params:
            info_text += " of " + str(total_params) + " (filtered)"
        
        info_text += " (from " + scope_text + ")" + element_count_text
        
        self.info_label.Text = info_text
    
    def OnParameterChecked(self, sender, e):
        """Update parameter selection count and global state when parameters are checked/unchecked"""
        try:
            # Update global selection state immediately
            if e.Index >= 0 and e.Index < self.parameters_checklist.Items.Count:
                item_text = str(self.parameters_checklist.Items[e.Index])
                if e.NewValue == System.Windows.Forms.CheckState.Checked:
                    self.selected_parameters.add(item_text)
                else:
                    self.selected_parameters.discard(item_text)
        except:
            pass
        
        self.BeginInvoke(System.Action(self.UpdateParameterCount))
    
    def UpdateParameterCount(self):
        """Update the parameter selection count label"""
        count = 0
        for i in range(self.parameters_checklist.Items.Count):
            if self.parameters_checklist.GetItemChecked(i):
                count += 1
        self.parameter_count_label.Text = "Selected: " + str(count) + " parameters"
    
    def GetSelectedParameters(self):
        """Get list of selected parameters"""
        selected = []
        for i in range(self.parameters_checklist.Items.Count):
            if self.parameters_checklist.GetItemChecked(i):
                selected.append(str(self.parameters_checklist.Items[i]))
        return selected
    
    def ShowProgress(self, show=True, message="Processing..."):
        """Show/hide progress bar"""
        try:
            if self.progress_bar:
                self.progress_bar.Visible = show
                if show:
                    self.progress_bar.Style = WinForms.ProgressBarStyle.Marquee
            if self.status_label:
                if show:
                    self.status_label.Text = message
                else:
                    self.status_label.Text = "Ready"
        except:
            # Silently handle any UI update errors during initialization
            pass
    
    def ShowSettings(self, sender, e):
        """Show DynaNex settings dialog"""
        settings_form = DynaNextSettingsForm(self)  # Pass self to settings form
        result = settings_form.ShowDialog()
        
        # If settings were changed, show confirmation
        if result == DialogResult.OK and settings_form.settings_changed:
            WinForms.MessageBox.Show(
                "Settings have been saved successfully!",
                "DynaNex - Settings Saved",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Information
            )
    
    def CloseForm(self, sender, e):
        """Close the form"""
        if self.is_processing:
            result = WinForms.MessageBox.Show(
                "Operation in progress. Are you sure you want to close?", 
                "DynaNex - Confirm Close", 
                WinForms.MessageBoxButtons.YesNo
            )
            if result != DialogResult.Yes:
                return
        
        self.Close()
# NEW METHODS: Save and Load Settings
    def SaveSettingsToFile(self, sender, e):
        """Save current settings to user-selected local file"""
        try:
            # Show Save File Dialog
            save_dialog = WinForms.SaveFileDialog()
            save_dialog.Filter = "JSON Settings Files (*.json)|*.json"
            save_dialog.FileName = "DynaNex_Settings_" + System.DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".json"
            save_dialog.Title = "Save DynaNex Settings As"
            save_dialog.DefaultExt = "json"
            save_dialog.AddExtension = True
            
            if save_dialog.ShowDialog() != DialogResult.OK:
                return  # User cancelled
            
            # Get current UI state - selected categories and parameters
            selected_categories = self.GetSelectedCategories()
            selected_parameters = self.GetSelectedParameters()
            
            # Create complete settings dictionary
            settings_data = {
                'app_settings': self.settings.copy(),
                'ui_state': {
                    'current_scope': self.current_scope,
                    'selected_categories': selected_categories,
                    'selected_parameters': selected_parameters,
                    'category_search': self.search_categories_textbox.Text if self.search_categories_textbox.Text != "Search categories..." else "",
                    'parameter_search': self.search_parameters_textbox.Text if self.search_parameters_textbox.Text != "Search parameters..." else ""
                },
                'version': '1.0',
                'timestamp': System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            }
            
            # Write to selected JSON file
            with open(save_dialog.FileName, 'w') as f:
                json.dump(settings_data, f, indent=4)
            
            # Update status
            self.settings_status_label.Text = "Settings: Saved"
            self.settings_status_label.ForeColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green
            
            WinForms.MessageBox.Show(
                "Settings saved successfully!\n\n" +
                "Location: " + save_dialog.FileName + "\n" +
                "Includes: App settings + current selections\n" +
                "Timestamp: " + settings_data['timestamp'],
                "DynaNex - Settings Saved",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Information
            )
            
        except Exception as e:
            WinForms.MessageBox.Show(
                "Failed to save settings:\n\n" + str(e),
                "DynaNex - Save Error",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Error
            )
    
    def LoadSettingsFromFile(self, sender=None, e=None):
        """Load settings from user-selected local file"""
        try:
            # If called during startup (sender is None), try to load default settings first
            if sender is None:
                if os.path.exists(self.settings_file_path):
                    try:
                        with open(self.settings_file_path, 'r') as f:
                            settings_data = json.load(f)
                        # Load only app settings during startup, not UI state
                        if 'app_settings' in settings_data:
                            self.settings.update(settings_data['app_settings'])
                        if self.settings_status_label:
                            self.settings_status_label.Text = "Settings: Default"
                            self.settings_status_label.ForeColor = Drawing.Color.FromArgb(156, 163, 175)  # Modern gray
                        return
                    except:
                        pass  # Silently fail and use defaults
                
                # No default settings found, use built-in defaults
                if self.settings_status_label:
                    self.settings_status_label.Text = "Settings: Default"
                    self.settings_status_label.ForeColor = Drawing.Color.FromArgb(108, 117, 125)  # Gray
                return
            
            # Manual load - show Open File Dialog
            open_dialog = WinForms.OpenFileDialog()
            open_dialog.Filter = "JSON Settings Files (*.json)|*.json"
            open_dialog.Title = "Load DynaNex Settings From"
            open_dialog.DefaultExt = "json"
            
            if open_dialog.ShowDialog() != DialogResult.OK:
                return  # User cancelled
            
            # Read from selected JSON file
            with open(open_dialog.FileName, 'r') as f:
                settings_data = json.load(f)
            
            # Load app settings
            if 'app_settings' in settings_data:
                self.settings.update(settings_data['app_settings'])
            
            # Load UI state if available
            if 'ui_state' in settings_data:
                ui_state = settings_data['ui_state']
                
                # Restore scope
                if 'current_scope' in ui_state:
                    if ui_state['current_scope'] == "WholeModel":
                        self.current_scope = "WholeModel"
                        if self.active_view_btn and self.whole_model_btn:
                            self.active_view_btn.BackColor = Drawing.Color.FromArgb(71, 85, 105)  # Modern gray
                            self.whole_model_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green
                    else:
                        self.current_scope = "ActiveView"
                        if self.active_view_btn and self.whole_model_btn:
                            self.active_view_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green
                            self.whole_model_btn.BackColor = Drawing.Color.FromArgb(71, 85, 105)  # Modern gray
                
                # Restore search filters
                if 'category_search' in ui_state and ui_state['category_search']:
                    if self.search_categories_textbox:
                        self.search_categories_textbox.Text = ui_state['category_search']
                        self.search_categories_textbox.ForeColor = Drawing.Color.Black
                
                if 'parameter_search' in ui_state and ui_state['parameter_search']:
                    if self.search_parameters_textbox:
                        self.search_parameters_textbox.Text = ui_state['parameter_search']
                        self.search_parameters_textbox.ForeColor = Drawing.Color.Black
                
                # Load data from current scope first
                self.LoadDataFromCurrentScope()
                
                # Then restore selections
                if 'selected_categories' in ui_state:
                    self.RestoreSelectedCategories(ui_state['selected_categories'])
                
                if 'selected_parameters' in ui_state:
                    self.RestoreSelectedParameters(ui_state['selected_parameters'])
            
            # Update status
            timestamp = settings_data.get('timestamp', 'Unknown')
            if self.settings_status_label:
                self.settings_status_label.Text = "Settings: Loaded"
                self.settings_status_label.ForeColor = Drawing.Color.FromArgb(59, 130, 246)  # Modern blue
            
            WinForms.MessageBox.Show(
                "Settings loaded successfully!\n\n" +
                "From: " + open_dialog.FileName + "\n" +
                "Saved: " + timestamp + "\n" +
                "Applied: App settings + UI selections",
                "DynaNex - Settings Loaded",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Information
            )
            
        except Exception as e:
            WinForms.MessageBox.Show(
                "Failed to load settings:\n\n" + str(e),
                "DynaNex - Load Error",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Error
            )
    
    def RestoreSelectedCategories(self, selected_categories):
        """Restore category selections from saved settings"""
        try:
            if not selected_categories or not self.categories_checklist:
                return
            
            for i in range(self.categories_checklist.Items.Count):
                item_text = str(self.categories_checklist.Items[i])
                if item_text in selected_categories:
                    self.categories_checklist.SetItemChecked(i, True)
            
            self.UpdateCategoryCount()
            
        except Exception as e:
            pass  # Silently handle restoration errors
    
    def RestoreSelectedParameters(self, selected_parameters):
        """Restore parameter selections from saved settings"""
        try:
            if not selected_parameters or not self.parameters_checklist:
                return
            
            for i in range(self.parameters_checklist.Items.Count):
                item_text = str(self.parameters_checklist.Items[i])
                if item_text in selected_parameters:
                    self.parameters_checklist.SetItemChecked(i, True)
            
            self.UpdateParameterCount()
            
        except Exception as e:
            pass  # Silently handle restoration errors

class DynaNextSettingsForm(WinForms.Form):
    """DynaNex settings dialog"""
    
    def __init__(self, parent_form=None):
        self.parent_form = parent_form
        self.settings_changed = False
        self.InitializeForm()
        self.LoadCurrentSettings()
        
    def LoadCurrentSettings(self):
        """Load current settings into the form"""
        # Load from parent form if available, otherwise use defaults
        if self.parent_form and hasattr(self.parent_form, 'settings'):
            settings = self.parent_form.settings
            self.confirm_whole_model.Checked = settings.get('confirm_whole_model', True)
            self.default_scope_active.Checked = settings.get('default_scope_active', True)
            self.auto_select_common_cats.Checked = settings.get('auto_select_common_cats', False)
            self.use_main_categories.Checked = settings.get('use_main_categories', True)
            self.include_empty_params.Checked = settings.get('include_empty_params', True)
            self.auto_open_excel.Checked = settings.get('auto_open_excel', True)
            self.export_only_selected.Checked = settings.get('export_only_selected', True)
            self.override_existing.Checked = settings.get('override_existing', True)
            self.skip_readonly.Checked = settings.get('skip_readonly', True)
        
    def SaveSettings(self):
        """Save the current settings back to parent form"""
        if self.parent_form:
            # Initialize settings dictionary if it doesn't exist
            if not hasattr(self.parent_form, 'settings'):
                self.parent_form.settings = {}
            
            # Save all settings
            self.parent_form.settings['confirm_whole_model'] = self.confirm_whole_model.Checked
            self.parent_form.settings['default_scope_active'] = self.default_scope_active.Checked
            self.parent_form.settings['auto_select_common_cats'] = self.auto_select_common_cats.Checked
            self.parent_form.settings['use_main_categories'] = self.use_main_categories.Checked
            self.parent_form.settings['include_empty_params'] = self.include_empty_params.Checked
            self.parent_form.settings['auto_open_excel'] = self.auto_open_excel.Checked
            self.parent_form.settings['export_only_selected'] = self.export_only_selected.Checked
            self.parent_form.settings['override_existing'] = self.override_existing.Checked
            self.parent_form.settings['skip_readonly'] = self.skip_readonly.Checked
            
            self.settings_changed = True
        
    def InitializeForm(self):
        self.Text = "DynaNex Settings"
        self.Size = Drawing.Size(550, 550)  # Increased height to ensure buttons are visible
        self.StartPosition = WinForms.FormStartPosition.CenterParent
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark theme background
        
        # Main panel
        main_panel = WinForms.TableLayoutPanel()
        main_panel.Dock = WinForms.DockStyle.Fill
        main_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        main_panel.RowCount = 6  # Fixed: Changed from 7 to 6
        main_panel.ColumnCount = 1
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 50))   # Title - row 0
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 100))  # Scope settings - row 1
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 100))  # Category settings - row 2
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 120))  # Export settings - row 3
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 100))  # Import settings - row 4
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 60))   # Buttons - row 5 (increased height)
        
        # Title
        title_label = WinForms.Label()
        title_label.Text = "⚡ DynaNex Settings"
        title_label.Font = Drawing.Font("Segoe UI", 14, Drawing.FontStyle.Bold)
        title_label.Dock = WinForms.DockStyle.Fill
        title_label.TextAlign = Drawing.ContentAlignment.MiddleCenter
        title_label.ForeColor = Drawing.Color.FromArgb(59, 130, 246)  # Modern blue
        title_label.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        
        # Scope settings group
        scope_group = WinForms.GroupBox()
        scope_group.Text = "Element Scope Options"
        scope_group.Dock = WinForms.DockStyle.Fill
        scope_group.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        scope_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        scope_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Slightly lighter dark
        
        scope_layout = WinForms.TableLayoutPanel()
        scope_layout.Dock = WinForms.DockStyle.Fill
        scope_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        scope_layout.RowCount = 2
        scope_layout.ColumnCount = 1
        
        self.confirm_whole_model = WinForms.CheckBox()
        self.confirm_whole_model.Text = "Always confirm before switching to Whole Model"
        self.confirm_whole_model.Checked = True
        self.confirm_whole_model.Dock = WinForms.DockStyle.Fill
        self.confirm_whole_model.Font = Drawing.Font("Segoe UI", 9)
        self.confirm_whole_model.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.confirm_whole_model.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        self.default_scope_active = WinForms.CheckBox()
        self.default_scope_active.Text = "Start with Active View scope by default"
        self.default_scope_active.Checked = True
        self.default_scope_active.Dock = WinForms.DockStyle.Fill
        self.default_scope_active.Font = Drawing.Font("Segoe UI", 9)
        self.default_scope_active.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.default_scope_active.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        scope_layout.Controls.Add(self.confirm_whole_model, 0, 0)
        scope_layout.Controls.Add(self.default_scope_active, 0, 1)
        scope_group.Controls.Add(scope_layout)
        
        # Category settings group (NEW)
        category_group = WinForms.GroupBox()
        category_group.Text = "Category Selection Options"
        category_group.Dock = WinForms.DockStyle.Fill
        category_group.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        category_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        category_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Slightly lighter dark
        
        category_layout = WinForms.TableLayoutPanel()
        category_layout.Dock = WinForms.DockStyle.Fill
        category_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        category_layout.RowCount = 2
        category_layout.ColumnCount = 1
        
        self.auto_select_common_cats = WinForms.CheckBox()
        self.auto_select_common_cats.Text = "Auto-select common categories (Walls, Floors, etc.)"
        self.auto_select_common_cats.Checked = False
        self.auto_select_common_cats.Dock = WinForms.DockStyle.Fill
        self.auto_select_common_cats.Font = Drawing.Font("Segoe UI", 9)
        self.auto_select_common_cats.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.auto_select_common_cats.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        self.use_main_categories = WinForms.CheckBox()
        self.use_main_categories.Text = "Group by main categories only (recommended)"
        self.use_main_categories.Checked = True
        self.use_main_categories.Dock = WinForms.DockStyle.Fill
        self.use_main_categories.Font = Drawing.Font("Segoe UI", 9)
        self.use_main_categories.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.use_main_categories.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        category_layout.Controls.Add(self.auto_select_common_cats, 0, 0)
        category_layout.Controls.Add(self.use_main_categories, 0, 1)
        category_group.Controls.Add(category_layout)
        
        # Export settings group
        export_group = WinForms.GroupBox()
        export_group.Text = "Export Options"
        export_group.Dock = WinForms.DockStyle.Fill
        export_group.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        export_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        export_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Slightly lighter dark
        
        export_layout = WinForms.TableLayoutPanel()
        export_layout.Dock = WinForms.DockStyle.Fill
        export_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        export_layout.RowCount = 3
        export_layout.ColumnCount = 1
        
        self.include_empty_params = WinForms.CheckBox()
        self.include_empty_params.Text = "Include empty parameter values"
        self.include_empty_params.Checked = True
        self.include_empty_params.Dock = WinForms.DockStyle.Fill
        self.include_empty_params.Font = Drawing.Font("Segoe UI", 9)
        self.include_empty_params.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.include_empty_params.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        self.auto_open_excel = WinForms.CheckBox()
        self.auto_open_excel.Text = "Automatically open Excel after export"
        self.auto_open_excel.Checked = True
        self.auto_open_excel.Dock = WinForms.DockStyle.Fill
        self.auto_open_excel.Font = Drawing.Font("Segoe UI", 9)
        self.auto_open_excel.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.auto_open_excel.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        self.export_only_selected = WinForms.CheckBox()
        self.export_only_selected.Text = "Export only selected categories (recommended)"
        self.export_only_selected.Checked = True
        self.export_only_selected.Dock = WinForms.DockStyle.Fill
        self.export_only_selected.Font = Drawing.Font("Segoe UI", 9)
        self.export_only_selected.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.export_only_selected.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        export_layout.Controls.Add(self.include_empty_params, 0, 0)
        export_layout.Controls.Add(self.auto_open_excel, 0, 1)
        export_layout.Controls.Add(self.export_only_selected, 0, 2)
        export_group.Controls.Add(export_layout)
        
        # Import settings group
        import_group = WinForms.GroupBox()
        import_group.Text = "Import Options"
        import_group.Dock = WinForms.DockStyle.Fill
        import_group.Font = Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold)
        import_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        import_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Slightly lighter dark
        
        import_layout = WinForms.TableLayoutPanel()
        import_layout.Dock = WinForms.DockStyle.Fill
        import_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        import_layout.RowCount = 3
        import_layout.ColumnCount = 1
        
        self.override_existing = WinForms.CheckBox()
        self.override_existing.Text = "Override existing parameter values (standard mode)"
        self.override_existing.Checked = True
        self.override_existing.Dock = WinForms.DockStyle.Fill
        self.override_existing.Font = Drawing.Font("Segoe UI", 9)
        self.override_existing.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.override_existing.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        self.skip_readonly = WinForms.CheckBox()
        self.skip_readonly.Text = "Skip read-only parameters during import"
        self.skip_readonly.Checked = True
        self.skip_readonly.Dock = WinForms.DockStyle.Fill
        self.skip_readonly.Font = Drawing.Font("Segoe UI", 9)
        self.skip_readonly.ForeColor = Drawing.Color.FromArgb(220, 220, 220)  # Light text
        self.skip_readonly.BackColor = Drawing.Color.FromArgb(51, 65, 85)  # Dark background
        
        import_layout.Controls.Add(self.override_existing, 0, 0)
        import_layout.Controls.Add(self.skip_readonly, 0, 1)
        import_group.Controls.Add(import_layout)
        
        # Buttons panel - Fixed configuration
        button_panel = WinForms.TableLayoutPanel()
        button_panel.Dock = WinForms.DockStyle.Fill
        button_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)  # Dark background
        button_panel.Height = 40  # Explicit height
        button_panel.ColumnCount = 3
        button_panel.RowCount = 1
        button_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 33.33))
        button_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 33.33))
        button_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 33.34))
        button_panel.Padding = WinForms.Padding(10)
        
        reset_btn = WinForms.Button()
        reset_btn.Text = "Reset Defaults"
        reset_btn.Size = Drawing.Size(100, 25)
        reset_btn.Anchor = WinForms.AnchorStyles.None
        reset_btn.Click += self.ResetDefaults
        reset_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)  # Modern blue-gray
        reset_btn.ForeColor = Drawing.Color.White
        reset_btn.Font = Drawing.Font("Segoe UI", 9)
        reset_btn.FlatStyle = WinForms.FlatStyle.Flat
        reset_btn.FlatAppearance.BorderSize = 1
        reset_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        ok_btn = WinForms.Button()
        ok_btn.Text = "OK"
        ok_btn.Size = Drawing.Size(100, 25)
        ok_btn.Anchor = WinForms.AnchorStyles.None
        ok_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)  # Modern green
        ok_btn.ForeColor = Drawing.Color.White
        ok_btn.Font = Drawing.Font("Segoe UI", 9)
        ok_btn.FlatStyle = WinForms.FlatStyle.Flat
        ok_btn.FlatAppearance.BorderSize = 1
        ok_btn.FlatAppearance.BorderColor = Drawing.Color.White
        ok_btn.Click += self.OnOKClicked
        
        cancel_btn = WinForms.Button()
        cancel_btn.Text = "Cancel"
        cancel_btn.Size = Drawing.Size(100, 25)
        cancel_btn.Anchor = WinForms.AnchorStyles.None
        cancel_btn.DialogResult = DialogResult.Cancel
        cancel_btn.BackColor = Drawing.Color.FromArgb(239, 68, 68)  # Modern red
        cancel_btn.ForeColor = Drawing.Color.White
        cancel_btn.Font = Drawing.Font("Segoe UI", 9)
        cancel_btn.FlatStyle = WinForms.FlatStyle.Flat
        cancel_btn.FlatAppearance.BorderSize = 1
        cancel_btn.FlatAppearance.BorderColor = Drawing.Color.White
        
        button_panel.Controls.Add(reset_btn, 0, 0)
        button_panel.Controls.Add(ok_btn, 1, 0)
        button_panel.Controls.Add(cancel_btn, 2, 0)
        
        # Add to main panel
        main_panel.Controls.Add(title_label, 0, 0)
        main_panel.Controls.Add(scope_group, 0, 1)
        main_panel.Controls.Add(category_group, 0, 2)
        main_panel.Controls.Add(export_group, 0, 3)
        main_panel.Controls.Add(import_group, 0, 4)
        main_panel.Controls.Add(button_panel, 0, 5)  # Fixed: Changed from row 6 to row 5
        
        self.Controls.Add(main_panel)
    
    def OnOKClicked(self, sender, e):
        """Handle OK button click"""
        self.SaveSettings()
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def ResetDefaults(self, sender, e):
        """Reset all settings to defaults"""
        self.confirm_whole_model.Checked = True
        self.default_scope_active.Checked = True
        self.auto_select_common_cats.Checked = False
        self.use_main_categories.Checked = True
        self.include_empty_params.Checked = True
        self.auto_open_excel.Checked = True
        self.export_only_selected.Checked = True
        self.override_existing.Checked = True
        self.skip_readonly.Checked = True


# Main execution function
def ShowDynaNex():
    """Show the DynaNex tool"""
    try:
        # Check if we're in the right context
        if not DocumentManager.Instance.CurrentDBDocument:
            WinForms.MessageBox.Show(
                "No active Revit document found.\n\nPlease open a Revit model first.", 
                "DynaNex - No Document", 
                WinForms.MessageBoxButtons.OK, 
                WinForms.MessageBoxIcon.Warning
            )
            return
        
        # Check if current view is valid
        current_view = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument.ActiveView
        if current_view is None:
            WinForms.MessageBox.Show(
                "No active view found.\n\nPlease make sure you have a view open in Revit.", 
                "DynaNex - No View", 
                WinForms.MessageBoxButtons.OK, 
                WinForms.MessageBoxIcon.Warning
            )
            return
        
        # Show welcome message for first-time users
        WinForms.MessageBox.Show(
            "DynaNex V1.0 - \"Power your Connections\"\n\n" +
            "Advanced parameter export/import tool\n" +
            "Connects Revit with Microsoft Excel\n" +
            "NEW: Category selection for focused exports\n" +
            "Export by selected categories to Excel\n" +
            "Import to update Revit parameters\n" +
            "Active View / Whole Model scope selection\n\n" +
            "Click 'Help' button for detailed instructions.",
            "DynaNex V1.0 - Power your Connections",
            WinForms.MessageBoxButtons.OK,
            WinForms.MessageBoxIcon.Information
        )
        
        # Create and show form
        form = DynaNextForm()
        WinForms.Application.EnableVisualStyles()
        WinForms.Application.Run(form)
        
    except Exception as e:
        error_msg = "Failed to start DynaNex:\n\n" + str(e)
        WinForms.MessageBox.Show(error_msg, "DynaNex - Startup Error", 
                               WinForms.MessageBoxButtons.OK, 
                               WinForms.MessageBoxIcon.Error)

# Run DynaNex
ShowDynaNex()