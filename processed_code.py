import clr
import sys
import os
import time
import json
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitNodes')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('Microsoft.Office.Interop.Excel')
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
class App1(WinForms.Form):
    def __init__(self):
        self.doc = DocumentManager.Instance.CurrentDBDocument
        self.uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
        self.data1 = []
        self.data2 = []
        self.data3 = []
        self.data4 = []
        self.flag1 = False
        self.mode = "ActiveView"
        self.sel1 = set()
        self.sel2 = set()
        self.theme = "dark"
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
        self.path1 = os.path.join(
            os.path.expanduser("~"),
            "Documents",
            "DynaNex_Settings.json"
        )
        self.list1 = None
        self.list2 = None
        self.label1 = None
        self.label2 = None
        self.label4 = None
        self.bar1 = None
        self.label5 = None
        self.input1 = None
        self.input2 = None
        self.label3 = None
        self.setup1()
        self.load1()
        try:
            self.process1()
        except Exception as e:
            self.data1 = []
            self.data2 = []
            self.data3 = []
            self.data4 = []
    def setup1(self):
        self.Text = "DynaNex V1.0 - Power your Connections"
        self.Size = Drawing.Size(1400, 980)
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.Sizable
        self.MinimumSize = Drawing.Size(1200, 900)
        self.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        main_panel = WinForms.TableLayoutPanel()
        main_panel.Dock = WinForms.DockStyle.Fill
        main_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        main_panel.RowCount = 7
        main_panel.ColumnCount = 1
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 80))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 50))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 45))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 35))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 45))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 55))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 140))
        title_panel = WinForms.TableLayoutPanel()
        title_panel.Dock = WinForms.DockStyle.Fill
        title_panel.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        title_panel.ColumnCount = 3
        title_panel.RowCount = 2
        title_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 25))
        title_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50))
        title_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 25))
        title_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 60))
        title_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 40))
        title_center_panel = WinForms.TableLayoutPanel()
        title_center_panel.Dock = WinForms.DockStyle.Fill
        title_center_panel.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        title_center_panel.RowCount = 2
        title_center_panel.ColumnCount = 1
        title_center_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 60))
        title_center_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 40))
        title_label = WinForms.Label()
        title_label.Text = "DynaNex"
        title_label.Font = Drawing.Font("Segoe UI", 20, Drawing.FontStyle.Bold)
        title_label.Dock = WinForms.DockStyle.Fill
        title_label.TextAlign = Drawing.ContentAlignment.BottomCenter
        title_label.ForeColor = Drawing.Color.FromArgb(0, 150, 255)
        title_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        tagline_label = WinForms.Label()
        tagline_label.Text = "Power your Connections"
        tagline_label.Font = Drawing.Font("Segoe UI", 11, Drawing.FontStyle.Italic)
        tagline_label.Dock = WinForms.DockStyle.Fill
        tagline_label.TextAlign = Drawing.ContentAlignment.TopCenter
        tagline_label.ForeColor = Drawing.Color.FromArgb(180, 180, 180)
        tagline_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        title_center_panel.Controls.Add(title_label, 0, 0)
        title_center_panel.Controls.Add(tagline_label, 0, 1)
        saveload_panel = WinForms.TableLayoutPanel()
        saveload_panel.Dock = WinForms.DockStyle.Fill
        saveload_panel.RowCount = 3
        saveload_panel.ColumnCount = 2
        saveload_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 40))
        saveload_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 25))
        saveload_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))
        saveload_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50))
        saveload_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50))
        saveload_panel.Padding = WinForms.Padding(10, 5, 10, 5)
        saveload_panel.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        save_settings_btn = WinForms.Button()
        save_settings_btn.Text = "Save"
        save_settings_btn.Size = Drawing.Size(85, 25)
        save_settings_btn.Anchor = WinForms.AnchorStyles.None
        save_settings_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)
        save_settings_btn.ForeColor = Drawing.Color.White
        save_settings_btn.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        save_settings_btn.FlatStyle = WinForms.FlatStyle.Flat
        save_settings_btn.FlatAppearance.BorderSize = 1
        save_settings_btn.FlatAppearance.BorderColor = Drawing.Color.White
        save_settings_btn.Click += self.save1
        load_settings_btn = WinForms.Button()
        load_settings_btn.Text = "Load"
        load_settings_btn.Size = Drawing.Size(85, 25)
        load_settings_btn.Anchor = WinForms.AnchorStyles.None
        load_settings_btn.BackColor = Drawing.Color.FromArgb(59, 130, 246)
        load_settings_btn.ForeColor = Drawing.Color.White
        load_settings_btn.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        load_settings_btn.FlatStyle = WinForms.FlatStyle.Flat
        load_settings_btn.FlatAppearance.BorderSize = 1
        load_settings_btn.FlatAppearance.BorderColor = Drawing.Color.White
        load_settings_btn.Click += self.load1
        self.label3 = WinForms.Label()
        self.label3.Text = "User Settings: Default"
        self.label3.Font = Drawing.Font("Segoe UI", 8)
        self.label3.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        self.label3.TextAlign = Drawing.ContentAlignment.MiddleCenter
        self.label3.Dock = WinForms.DockStyle.Fill
        self.label3.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        saveload_panel.Controls.Add(save_settings_btn, 0, 0)
        saveload_panel.Controls.Add(load_settings_btn, 1, 0)
        saveload_panel.Controls.Add(self.label3, 0, 1)
        saveload_panel.SetColumnSpan(self.label3, 2)
        aboutme_panel = WinForms.TableLayoutPanel()
        aboutme_panel.Dock = WinForms.DockStyle.Fill
        aboutme_panel.RowCount = 2
        aboutme_panel.ColumnCount = 1
        aboutme_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 40))
        aboutme_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 25))
        aboutme_panel.Padding = WinForms.Padding(10, 5, 10, 5)
        aboutme_panel.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        button_row_panel = WinForms.TableLayoutPanel()
        button_row_panel.Dock = WinForms.DockStyle.Fill
        button_row_panel.ColumnCount = 6
        button_row_panel.RowCount = 1
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 85))
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 80))
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 40))
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 10))
        button_row_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 40))
        aboutme_btn = WinForms.Button()
        aboutme_btn.Text = "About Me"
        aboutme_btn.Size = Drawing.Size(85, 25)
        aboutme_btn.Anchor = WinForms.AnchorStyles.None
        aboutme_btn.BackColor = Drawing.Color.FromArgb(20, 184, 166)
        aboutme_btn.ForeColor = Drawing.Color.White
        aboutme_btn.Font = Drawing.Font("Arial", 10, Drawing.FontStyle.Bold)
        aboutme_btn.FlatStyle = WinForms.FlatStyle.Flat
        aboutme_btn.FlatAppearance.BorderSize = 1
        aboutme_btn.FlatAppearance.BorderColor = Drawing.Color.White
        aboutme_btn.Click += self.open1
        switch_mode_label = WinForms.Label()
        switch_mode_label.Text = "Switch Mode"
        switch_mode_label.Font = Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Regular)
        switch_mode_label.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        switch_mode_label.TextAlign = Drawing.ContentAlignment.MiddleCenter
        switch_mode_label.Dock = WinForms.DockStyle.Fill
        switch_mode_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        self.btn3 = WinForms.Button()
        self.btn3.Text = "●"
        self.btn3.Size = Drawing.Size(28, 28)
        self.btn3.Anchor = WinForms.AnchorStyles.None
        self.btn3.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        self.btn3.ForeColor = Drawing.Color.White
        self.btn3.Font = Drawing.Font("Arial", 12, Drawing.FontStyle.Bold)
        self.btn3.FlatStyle = WinForms.FlatStyle.Flat
        self.btn3.FlatAppearance.BorderSize = 2
        self.btn3.FlatAppearance.BorderColor = Drawing.Color.White
        self.btn3.Click += self.theme1
        self.btn4 = WinForms.Button()
        self.btn4.Text = "●"
        self.btn4.Size = Drawing.Size(28, 28)
        self.btn4.Anchor = WinForms.AnchorStyles.None
        self.btn4.BackColor = Drawing.Color.White
        self.btn4.ForeColor = Drawing.Color.Black
        self.btn4.Font = Drawing.Font("Arial", 12, Drawing.FontStyle.Bold)
        self.btn4.FlatStyle = WinForms.FlatStyle.Flat
        self.btn4.FlatAppearance.BorderSize = 1
        self.btn4.FlatAppearance.BorderColor = Drawing.Color.Gray
        self.btn4.Click += self.theme2
        theme_label_panel = WinForms.TableLayoutPanel()
        theme_label_panel.Dock = WinForms.DockStyle.Fill
        theme_label_panel.ColumnCount = 6
        theme_label_panel.RowCount = 1
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 85))
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 80))
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 40))
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 10))
        theme_label_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 40))
        dark_label = WinForms.Label()
        dark_label.Text = "Dark"
        dark_label.Font = Drawing.Font("Segoe UI", 8.5, Drawing.FontStyle.Regular)
        dark_label.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        dark_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        dark_label.TextAlign = Drawing.ContentAlignment.TopCenter
        dark_label.Dock = WinForms.DockStyle.Fill
        light_label = WinForms.Label()
        light_label.Text = "Light"
        light_label.Font = Drawing.Font("Segoe UI", 8.5, Drawing.FontStyle.Regular)
        light_label.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        light_label.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        light_label.TextAlign = Drawing.ContentAlignment.TopCenter
        light_label.Dock = WinForms.DockStyle.Fill
        button_row_panel.Controls.Add(aboutme_btn, 0, 0)
        button_row_panel.Controls.Add(switch_mode_label, 2, 0)
        button_row_panel.Controls.Add(self.btn3, 3, 0)
        button_row_panel.Controls.Add(self.btn4, 5, 0)
        theme_label_panel.Controls.Add(dark_label, 3, 0)
        theme_label_panel.Controls.Add(light_label, 5, 0)
        aboutme_panel.Controls.Add(button_row_panel, 0, 0)
        aboutme_panel.Controls.Add(theme_label_panel, 0, 1)
        title_panel.Controls.Add(aboutme_panel, 0, 0)
        title_panel.SetRowSpan(aboutme_panel, 2)
        title_panel.Controls.Add(title_center_panel, 1, 0)
        title_panel.SetRowSpan(title_center_panel, 2)
        title_panel.Controls.Add(saveload_panel, 2, 0)
        title_panel.SetRowSpan(saveload_panel, 2)
        scope_panel = WinForms.TableLayoutPanel()
        scope_panel.Dock = WinForms.DockStyle.Fill
        scope_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)
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
        scope_label.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.btn1 = WinForms.Button()
        self.btn1.Text = "🔍 Active View Elements"
        self.btn1.Dock = WinForms.DockStyle.Fill
        self.btn1.Click += self.mode1
        self.btn1.BackColor = Drawing.Color.FromArgb(34, 197, 94)
        self.btn1.ForeColor = Drawing.Color.White
        self.btn1.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        self.btn1.FlatStyle = WinForms.FlatStyle.Flat
        self.btn1.FlatAppearance.BorderSize = 1
        self.btn1.FlatAppearance.BorderColor = Drawing.Color.White
        self.btn2 = WinForms.Button()
        self.btn2.Text = "🏢 Whole Model Elements"
        self.btn2.Dock = WinForms.DockStyle.Fill
        self.btn2.Click += self.mode2
        self.btn2.BackColor = Drawing.Color.FromArgb(71, 85, 105)
        self.btn2.ForeColor = Drawing.Color.White
        self.btn2.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        self.btn2.FlatStyle = WinForms.FlatStyle.Flat
        self.btn2.FlatAppearance.BorderSize = 1
        self.btn2.FlatAppearance.BorderColor = Drawing.Color.White
        scope_panel.Controls.Add(scope_label, 0, 0)
        scope_panel.Controls.Add(self.btn1, 1, 0)
        scope_panel.Controls.Add(self.btn2, 2, 0)
        search_panel = WinForms.TableLayoutPanel()
        search_panel.Dock = WinForms.DockStyle.Fill
        search_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        search_panel.ColumnCount = 5
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 100))
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 45))
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 70))
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 45))
        search_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 70))
        search_label = WinForms.Label()
        search_label.Text = "🔍 Filter:"
        search_label.TextAlign = Drawing.ContentAlignment.MiddleLeft
        search_label.Dock = WinForms.DockStyle.Fill
        search_label.Font = Drawing.Font("Segoe UI", 9)
        search_label.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.input1 = WinForms.TextBox()
        self.input1.Dock = WinForms.DockStyle.Fill
        self.input1.Font = Drawing.Font("Segoe UI", 10)
        self.input1.Text = "Search categories..."
        self.input1.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        self.input1.BackColor = Drawing.Color.FromArgb(71, 85, 105)
        self.input2 = WinForms.TextBox()
        self.input2.Dock = WinForms.DockStyle.Fill
        self.input2.Font = Drawing.Font("Segoe UI", 10)
        self.input2.Text = "Search parameters..."
        self.input2.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        self.input2.BackColor = Drawing.Color.FromArgb(71, 85, 105)
        self.input1.TextChanged += self.search1
        self.input1.Enter += self.enter1
        self.input1.Leave += self.leave1
        self.input2.TextChanged += self.search2
        self.input2.Enter += self.enter2
        self.input2.Leave += self.leave2
        clear_cat_btn = WinForms.Button()
        clear_cat_btn.Text = "Clear"
        clear_cat_btn.Size = Drawing.Size(65, 23)
        clear_cat_btn.Anchor = WinForms.AnchorStyles.Top
        clear_cat_btn.Click += self.clear1
        clear_cat_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)
        clear_cat_btn.ForeColor = Drawing.Color.White
        clear_cat_btn.Font = Drawing.Font("Arial", 7, Drawing.FontStyle.Bold)
        clear_cat_btn.FlatStyle = WinForms.FlatStyle.Flat
        clear_cat_btn.FlatAppearance.BorderSize = 1
        clear_cat_btn.FlatAppearance.BorderColor = Drawing.Color.White
        clear_param_btn = WinForms.Button()
        clear_param_btn.Text = "Clear"
        clear_param_btn.Size = Drawing.Size(65, 23)
        clear_param_btn.Anchor = WinForms.AnchorStyles.Top
        clear_param_btn.Click += self.clear2
        clear_param_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)
        clear_param_btn.ForeColor = Drawing.Color.White
        clear_param_btn.Font = Drawing.Font("Arial", 7, Drawing.FontStyle.Bold)
        clear_param_btn.FlatStyle = WinForms.FlatStyle.Flat
        clear_param_btn.FlatAppearance.BorderSize = 1
        clear_param_btn.FlatAppearance.BorderColor = Drawing.Color.White
        search_panel.Controls.Add(search_label, 0, 0)
        search_panel.Controls.Add(self.input1, 1, 0)
        search_panel.Controls.Add(clear_cat_btn, 2, 0)
        search_panel.Controls.Add(self.input2, 3, 0)
        search_panel.Controls.Add(clear_param_btn, 4, 0)
        self.label4 = WinForms.Label()
        self.label4.Text = "Loading categories and parameters from active view..."
        self.label4.Dock = WinForms.DockStyle.Fill
        self.label4.TextAlign = Drawing.ContentAlignment.MiddleLeft
        self.label4.Font = Drawing.Font("Segoe UI", 11, Drawing.FontStyle.Bold)
        self.label4.ForeColor = Drawing.Color.FromArgb(59, 130, 246)
        self.label4.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        content_panel = WinForms.TableLayoutPanel()
        content_panel.Dock = WinForms.DockStyle.Fill
        content_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        content_panel.ColumnCount = 2
        content_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 45))
        content_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 55))
        content_panel.Padding = WinForms.Padding(5)
        self.list1 = WinForms.CheckedListBox()
        self.list1.Dock = WinForms.DockStyle.Fill
        self.list1.CheckOnClick = True
        self.list1.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        self.list1.ItemHeight = 18
        self.list1.BackColor = Drawing.Color.FromArgb(71, 85, 105)
        self.list1.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.list1.ItemCheck += self.check1
        self.list2 = WinForms.CheckedListBox()
        self.list2.Dock = WinForms.DockStyle.Fill
        self.list2.CheckOnClick = True
        self.list2.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        self.list2.ItemHeight = 18
        self.list2.BackColor = Drawing.Color.FromArgb(71, 85, 105)
        self.list2.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.list2.ItemCheck += self.check2
        categories_group = WinForms.GroupBox()
        categories_group.Text = "📁 Categories"
        categories_group.Dock = WinForms.DockStyle.Fill
        categories_group.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        categories_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        categories_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        categories_layout = WinForms.TableLayoutPanel()
        categories_layout.Dock = WinForms.DockStyle.Fill
        categories_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        categories_layout.RowCount = 2
        categories_layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))
        categories_layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 25))
        self.label1 = WinForms.Label()
        self.label1.Text = "Selected: 0 categories"
        self.label1.TextAlign = Drawing.ContentAlignment.MiddleCenter
        self.label1.Dock = WinForms.DockStyle.Fill
        self.label1.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        self.label1.ForeColor = Drawing.Color.FromArgb(239, 68, 68)
        self.label1.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        categories_layout.Controls.Add(self.list1, 0, 0)
        categories_layout.Controls.Add(self.label1, 0, 1)
        categories_group.Controls.Add(categories_layout)
        parameters_group = WinForms.GroupBox()
        parameters_group.Text = "🔧 Parameters"
        parameters_group.Dock = WinForms.DockStyle.Fill
        parameters_group.Font = Drawing.Font("Arial", 10, Drawing.FontStyle.Bold)
        parameters_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        parameters_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        parameters_layout = WinForms.TableLayoutPanel()
        parameters_layout.Dock = WinForms.DockStyle.Fill
        parameters_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        parameters_layout.RowCount = 2
        parameters_layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))
        parameters_layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 25))
        self.label2 = WinForms.Label()
        self.label2.Text = "Selected: 0 parameters"
        self.label2.TextAlign = Drawing.ContentAlignment.MiddleCenter
        self.label2.Dock = WinForms.DockStyle.Fill
        self.label2.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        self.label2.ForeColor = Drawing.Color.FromArgb(59, 130, 246)
        self.label2.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        parameters_layout.Controls.Add(self.list2, 0, 0)
        parameters_layout.Controls.Add(self.label2, 0, 1)
        parameters_group.Controls.Add(parameters_layout)
        content_panel.Controls.Add(categories_group, 0, 0)
        content_panel.Controls.Add(parameters_group, 1, 0)
        selection_panel = WinForms.TableLayoutPanel()
        selection_panel.Dock = WinForms.DockStyle.Fill
        selection_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        selection_panel.ColumnCount = 6
        selection_panel.Padding = WinForms.Padding(5)
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        selection_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 16.66))
        select_all_cats_btn = WinForms.Button()
        select_all_cats_btn.Text = "✓ All Categories"
        select_all_cats_btn.Size = Drawing.Size(140, 35)
        select_all_cats_btn.Anchor = WinForms.AnchorStyles.None
        select_all_cats_btn.Click += self.select1
        select_all_cats_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)
        select_all_cats_btn.ForeColor = Drawing.Color.White
        select_all_cats_btn.Font = Drawing.Font("Arial", 10.5, Drawing.FontStyle.Bold)
        select_all_cats_btn.FlatStyle = WinForms.FlatStyle.Flat
        select_all_cats_btn.FlatAppearance.BorderSize = 1
        select_all_cats_btn.FlatAppearance.BorderColor = Drawing.Color.White
        deselect_all_cats_btn = WinForms.Button()
        deselect_all_cats_btn.Text = "✗ Clear"
        deselect_all_cats_btn.Size = Drawing.Size(100, 35)
        deselect_all_cats_btn.Anchor = WinForms.AnchorStyles.None
        deselect_all_cats_btn.Click += self.deselect1
        deselect_all_cats_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)
        deselect_all_cats_btn.ForeColor = Drawing.Color.White
        deselect_all_cats_btn.Font = Drawing.Font("Arial", 10.5, Drawing.FontStyle.Bold)
        deselect_all_cats_btn.FlatStyle = WinForms.FlatStyle.Flat
        deselect_all_cats_btn.FlatAppearance.BorderSize = 1
        deselect_all_cats_btn.FlatAppearance.BorderColor = Drawing.Color.White
        select_filtered_cats_btn = WinForms.Button()
        select_filtered_cats_btn.Text = "🎯 Filter"
        select_filtered_cats_btn.Size = Drawing.Size(100, 35)
        select_filtered_cats_btn.Anchor = WinForms.AnchorStyles.None
        select_filtered_cats_btn.Click += self.filter1
        select_filtered_cats_btn.BackColor = Drawing.Color.FromArgb(59, 130, 246)
        select_filtered_cats_btn.ForeColor = Drawing.Color.White
        select_filtered_cats_btn.Font = Drawing.Font("Arial", 10.5, Drawing.FontStyle.Bold)
        select_filtered_cats_btn.FlatStyle = WinForms.FlatStyle.Flat
        select_filtered_cats_btn.FlatAppearance.BorderSize = 1
        select_filtered_cats_btn.FlatAppearance.BorderColor = Drawing.Color.White
        select_all_params_btn = WinForms.Button()
        select_all_params_btn.Text = "✓ All Parameters"
        select_all_params_btn.Size = Drawing.Size(140, 35)
        select_all_params_btn.Anchor = WinForms.AnchorStyles.None
        select_all_params_btn.Click += self.select2
        select_all_params_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)
        select_all_params_btn.ForeColor = Drawing.Color.White
        select_all_params_btn.Font = Drawing.Font("Arial", 10.5, Drawing.FontStyle.Bold)
        select_all_params_btn.FlatStyle = WinForms.FlatStyle.Flat
        select_all_params_btn.FlatAppearance.BorderSize = 1
        select_all_params_btn.FlatAppearance.BorderColor = Drawing.Color.White
        deselect_all_params_btn = WinForms.Button()
        deselect_all_params_btn.Text = "✗ Clear"
        deselect_all_params_btn.Size = Drawing.Size(100, 35)
        deselect_all_params_btn.Anchor = WinForms.AnchorStyles.None
        deselect_all_params_btn.Click += self.deselect2
        deselect_all_params_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)
        deselect_all_params_btn.ForeColor = Drawing.Color.White
        deselect_all_params_btn.Font = Drawing.Font("Arial", 10.5, Drawing.FontStyle.Bold)
        deselect_all_params_btn.FlatStyle = WinForms.FlatStyle.Flat
        deselect_all_params_btn.FlatAppearance.BorderSize = 1
        deselect_all_params_btn.FlatAppearance.BorderColor = Drawing.Color.White
        select_filtered_params_btn = WinForms.Button()
        select_filtered_params_btn.Text = "🎯 Filter"
        select_filtered_params_btn.Size = Drawing.Size(100, 35)
        select_filtered_params_btn.Anchor = WinForms.AnchorStyles.None
        select_filtered_params_btn.Click += self.filter2
        select_filtered_params_btn.BackColor = Drawing.Color.FromArgb(59, 130, 246)
        select_filtered_params_btn.ForeColor = Drawing.Color.White
        select_filtered_params_btn.Font = Drawing.Font("Arial", 10.5, Drawing.FontStyle.Bold)
        select_filtered_params_btn.FlatStyle = WinForms.FlatStyle.Flat
        select_filtered_params_btn.FlatAppearance.BorderSize = 1
        select_filtered_params_btn.FlatAppearance.BorderColor = Drawing.Color.White
        selection_panel.Controls.Add(select_all_cats_btn, 0, 0)
        selection_panel.Controls.Add(deselect_all_cats_btn, 1, 0)
        selection_panel.Controls.Add(select_filtered_cats_btn, 2, 0)
        selection_panel.Controls.Add(select_all_params_btn, 3, 0)
        selection_panel.Controls.Add(deselect_all_params_btn, 4, 0)
        selection_panel.Controls.Add(select_filtered_params_btn, 5, 0)
        self.action_panel = WinForms.TableLayoutPanel()
        action_panel = self.action_panel
        action_panel.Dock = WinForms.DockStyle.Fill
        action_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        action_panel.RowCount = 3
        action_panel.ColumnCount = 6
        action_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 35))
        action_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 35))
        action_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 50))
        action_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 210))
        action_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 210))
        action_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
        action_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 50))
        action_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 130))
        action_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Absolute, 130))
        action_panel.Padding = WinForms.Padding(10, 5, 10, 15)
        button_size = Drawing.Size(200, 30)
        export_btn = WinForms.Button()
        export_btn.Text = "📊 Export to Excel"
        export_btn.Size = button_size
        export_btn.Anchor = WinForms.AnchorStyles.None
        export_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)
        export_btn.ForeColor = Drawing.Color.White
        export_btn.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        export_btn.FlatStyle = WinForms.FlatStyle.Flat
        export_btn.FlatAppearance.BorderSize = 1
        export_btn.FlatAppearance.BorderColor = Drawing.Color.White
        export_btn.Click += self.export1
        import_btn = WinForms.Button()
        import_btn.Text = "📥 Import from Excel"
        import_btn.Size = button_size
        import_btn.Anchor = WinForms.AnchorStyles.None
        import_btn.BackColor = Drawing.Color.FromArgb(245, 158, 11)
        import_btn.ForeColor = Drawing.Color.White
        import_btn.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        import_btn.FlatStyle = WinForms.FlatStyle.Flat
        import_btn.FlatAppearance.BorderSize = 1
        import_btn.FlatAppearance.BorderColor = Drawing.Color.White
        import_btn.Click += self.import1
        refresh_btn = WinForms.Button()
        refresh_btn.Text = "🔄 Refresh"
        refresh_btn.Size = Drawing.Size(120, 30)
        refresh_btn.Anchor = WinForms.AnchorStyles.Top
        refresh_btn.BackColor = Drawing.Color.FromArgb(20, 184, 166)
        refresh_btn.ForeColor = Drawing.Color.White
        refresh_btn.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        refresh_btn.FlatStyle = WinForms.FlatStyle.Flat
        refresh_btn.FlatAppearance.BorderSize = 1
        refresh_btn.FlatAppearance.BorderColor = Drawing.Color.White
        refresh_btn.Click += self.refresh1
        settings_btn = WinForms.Button()
        settings_btn.Text = "⚙️ Settings"
        settings_btn.Size = button_size
        settings_btn.Anchor = WinForms.AnchorStyles.None
        settings_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)
        settings_btn.ForeColor = Drawing.Color.White
        settings_btn.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        settings_btn.FlatStyle = WinForms.FlatStyle.Flat
        settings_btn.FlatAppearance.BorderSize = 1
        settings_btn.FlatAppearance.BorderColor = Drawing.Color.White
        settings_btn.Click += self.config1
        help_btn = WinForms.Button()
        help_btn.Text = "❓ Help"
        help_btn.Size = button_size
        help_btn.Anchor = WinForms.AnchorStyles.None
        help_btn.BackColor = Drawing.Color.FromArgb(20, 184, 166)
        help_btn.ForeColor = Drawing.Color.White
        help_btn.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        help_btn.FlatStyle = WinForms.FlatStyle.Flat
        help_btn.FlatAppearance.BorderSize = 1
        help_btn.FlatAppearance.BorderColor = Drawing.Color.White
        help_btn.Click += self.help1
        close_btn = WinForms.Button()
        close_btn.Text = "❌ Close"
        close_btn.Size = Drawing.Size(120, 30)
        close_btn.Anchor = WinForms.AnchorStyles.Top
        close_btn.BackColor = Drawing.Color.FromArgb(239, 68, 68)
        close_btn.ForeColor = Drawing.Color.White
        close_btn.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        close_btn.FlatStyle = WinForms.FlatStyle.Flat
        close_btn.FlatAppearance.BorderSize = 1
        close_btn.FlatAppearance.BorderColor = Drawing.Color.White
        close_btn.Click += self.close1
        self.bar1 = WinForms.ProgressBar()
        self.bar1.Dock = WinForms.DockStyle.Fill
        self.bar1.Visible = False
        self.bar1.Style = WinForms.ProgressBarStyle.Continuous
        self.bar1.Minimum = 0
        self.bar1.Maximum = 100
        self.bar1.Value = 0
        self.bar1.ForeColor = Drawing.Color.FromArgb(34, 197, 94)
        self.panel1 = WinForms.Panel()
        self.panel1.Anchor = WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right | WinForms.AnchorStyles.Top
        self.panel1.Location = Drawing.Point(0, 10)
        self.panel1.Height = 38
        self.panel1.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        self.panel1.BorderStyle = WinForms.BorderStyle.FixedSingle
        self.panel1.Visible = False
        self.fill1 = WinForms.Panel()
        self.fill1.Height = 36
        self.fill1.Width = 0
        self.fill1.BackColor = Drawing.Color.FromArgb(34, 197, 94)
        self.fill1.Location = Drawing.Point(1, 1)
        self.fill1.Anchor = WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom
        self.text1 = WinForms.Label()
        self.text1.Dock = WinForms.DockStyle.Fill
        self.text1.TextAlign = Drawing.ContentAlignment.MiddleCenter
        self.text1.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        self.text1.ForeColor = Drawing.Color.White
        self.text1.Font = Drawing.Font("Segoe UI", 12, Drawing.FontStyle.Bold)
        self.text1.Text = "0%"
        self.panel1.Controls.Add(self.fill1)
        self.label5 = WinForms.Label()
        self.label5.Text = "Ready"
        self.label5.Dock = WinForms.DockStyle.Fill
        self.label5.TextAlign = Drawing.ContentAlignment.MiddleLeft
        self.label5.Font = Drawing.Font("Segoe UI", 10, Drawing.FontStyle.Bold)
        self.label5.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        self.label5.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        action_panel.Controls.Add(export_btn, 0, 0)
        action_panel.Controls.Add(import_btn, 1, 0)
        action_panel.Controls.Add(settings_btn, 0, 1)
        action_panel.Controls.Add(help_btn, 1, 1)
        action_panel.Controls.Add(self.label5, 2, 1)
        action_panel.Controls.Add(self.text1, 0, 2)
        action_panel.Controls.Add(self.panel1, 1, 2)
        action_panel.SetColumnSpan(self.panel1, 2)
        action_panel.Controls.Add(refresh_btn, 4, 1)
        action_panel.Controls.Add(close_btn, 5, 1)
        main_panel.Controls.Add(title_panel, 0, 0)
        main_panel.Controls.Add(scope_panel, 0, 1)
        main_panel.Controls.Add(search_panel, 0, 2)
        main_panel.Controls.Add(self.label4, 0, 3)
        main_panel.Controls.Add(content_panel, 0, 4)
        main_panel.Controls.Add(selection_panel, 0, 5)
        main_panel.Controls.Add(action_panel, 0, 6)
        self.Controls.Add(main_panel)
    def mode1(self, sender, e):
        """Switch to Active View Elements mode"""
        if self.flag1:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
        self.mode = "ActiveView"
        self.btn1.BackColor = Drawing.Color.FromArgb(34, 197, 94)
        self.btn2.BackColor = Drawing.Color.FromArgb(71, 85, 105)
        self.process1()
    def mode2(self, sender, e):
        """Switch to Whole Model Elements mode"""
        if self.flag1:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
        result = WinForms.MessageBox.Show(
            "⚠️ Switching to Whole Model Elements will analyze ALL elements in the model.\n\n" +
            "This may take some time for large models. Continue?",
            "DynaNex - Confirm Scope Change",
            WinForms.MessageBoxButtons.YesNo,
            WinForms.MessageBoxIcon.Question
        )
        if result == DialogResult.Yes:
            self.mode = "WholeModel"
            self.btn1.BackColor = Drawing.Color.FromArgb(71, 85, 105)
            self.btn2.BackColor = Drawing.Color.FromArgb(34, 197, 94)
            self.process1()
    def process1(self):
        """Load categories and parameters from current scope - ENHANCED"""
        try:
            if not self.list1 or not self.list2:
                return
            if self.mode == "ActiveView":
                self.progress1(True, "Analyzing active view elements...")
                scope_text = "active view"
            else:
                self.progress1(True, "Analyzing whole model elements...")
                scope_text = "whole model"
            self.data1 = []
            self.data3 = []
            param_names = set()
            category_names = set()
            if self.mode == "ActiveView":
                current_view = self.uidoc.ActiveView
                collector = FilteredElementCollector(self.doc, current_view.Id)
                model_elements = collector.WhereElementIsNotElementType().ToElements()
                type_collector = FilteredElementCollector(self.doc, current_view.Id)
                type_elements = type_collector.WhereElementIsElementType().ToElements()
                elements = list(model_elements) + list(type_elements)
            else:
                collector = FilteredElementCollector(self.doc)
                model_elements = collector.WhereElementIsNotElementType().ToElements()
                type_collector = FilteredElementCollector(self.doc)
                type_elements = type_collector.WhereElementIsElementType().ToElements()
                elements = list(model_elements) + list(type_elements)
            self.count1 = len(elements)
            self.progress1(True, "Extracting categories from all elements...")
            for i, element in enumerate(elements):
                if element and element.Category:
                    category = element.Category
                    if category and category.Name:
                        category_names.add(category.Name.strip())
                        if category.Parent:
                            category_names.add(category.Parent.Name.strip())
                if i % 100 == 0:
                    progress = int((i * 50) / len(elements))
                    self.update1(progress, "Analyzing categories... " + str(progress) + "%")
                    WinForms.Application.DoEvents()
            total_elements = len(elements)
            if self.mode == "ActiveView":
                sample_size = min(150, total_elements)
            else:
                sample_size = min(300, total_elements)
            step = max(1, total_elements // sample_size)
            sample_elements = []
            for i in range(0, total_elements, step):
                sample_elements.append(elements[i])
                if len(sample_elements) >= sample_size:
                    break
            self.progress1(True, "Extracting parameters from " + str(len(sample_elements)) + " sample elements...")
            for i, element in enumerate(sample_elements):
                if element and element.Parameters:
                    for param in element.Parameters:
                        if param and param.Definition and param.Definition.Name:
                            param_name = param.Definition.Name
                            if param_name and len(param_name.strip()) > 0:
                                param_names.add(param_name.strip())
                if i % 10 == 0:
                    progress = 50 + int((i * 50) / len(sample_elements))
                    self.update1(progress, "Analyzing parameters... " + str(progress) + "%")
                    WinForms.Application.DoEvents()
            self.data3 = sorted(list(category_names))
            self.data1 = sorted(list(param_names))
            self.data4 = self.data3[:]
            self.data2 = self.data1[:]
            self.updatelist1()
            self.updatelist2()
            self.updateinfo1()
            self.progress1(False)
        except Exception as e:
            self.progress1(False)
            self.label4.Text = "Error loading data: " + str(e)
            WinForms.MessageBox.Show("Error loading data: " + str(e), "DynaNex - Error")
    def enter1(self, sender, e):
        """Handle category search box focus enter"""
        try:
            if sender and hasattr(sender, 'Text'):
                if sender.Text == "Search categories...":
                    sender.Text = ""
                    sender.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        except:
            pass
    def leave1(self, sender, e):
        """Handle category search box focus leave"""
        try:
            if sender and hasattr(sender, 'Text'):
                if sender.Text == "":
                    sender.Text = "Search categories..."
                    sender.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        except:
            pass
    def enter2(self, sender, e):
        """Handle parameter search box focus enter"""
        try:
            if sender and hasattr(sender, 'Text'):
                if sender.Text == "Search parameters...":
                    sender.Text = ""
                    sender.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        except:
            pass
    def leave2(self, sender, e):
        """Handle parameter search box focus leave"""
        try:
            if sender and hasattr(sender, 'Text'):
                if sender.Text == "":
                    sender.Text = "Search parameters..."
                    sender.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
        except:
            pass
    def updatelist1(self):
        """Update the categories checklist with filtered results while preserving global selections"""
        try:
            if self.list1 and hasattr(self.list1, 'Items'):
                for i in range(self.list1.Items.Count):
                    item_text = str(self.list1.Items[i])
                    if self.list1.GetItemChecked(i):
                        self.sel1.add(item_text)
                    elif item_text in self.sel1:
                        self.sel1.remove(item_text)
                self.list1.Items.Clear()
                for category in self.data4:
                    index = self.list1.Items.Add(category)
                    if category in self.sel1:
                        self.list1.SetItemChecked(index, True)
        except:
            pass
    def search1(self, sender, e):
        """Handle category search text changes"""
        try:
            if not self.input1:
                return
            search_text = self.input1.Text.lower()
            if search_text == "search categories..." or not search_text:
                self.data4 = self.data3[:]
            else:
                self.data4 = []
                for category in self.data3:
                    if search_text in category.lower():
                        self.data4.append(category)
            self.updatelist1()
            self.updateinfo1()
        except:
            pass
    def clear1(self, sender, e):
        """Clear category search text"""
        try:
            if self.input1:
                self.input1.Text = "Search categories..."
                self.input1.ForeColor = Drawing.Color.Gray
        except:
            pass
    def check1(self, sender, e):
        """Update category selection count and global state when categories are checked/unchecked"""
        try:
            if e.Index >= 0 and e.Index < self.list1.Items.Count:
                item_text = str(self.list1.Items[e.Index])
                if e.NewValue == System.Windows.Forms.CheckState.Checked:
                    self.sel1.add(item_text)
                else:
                    self.sel1.discard(item_text)
        except:
            pass
        self.BeginInvoke(System.Action(self.UpdateCategoryCount))
    def UpdateCategoryCount(self):
        """Update the category selection count label"""
        count = 0
        for i in range(self.list1.Items.Count):
            if self.list1.GetItemChecked(i):
                count += 1
        self.label1.Text = "Selected: " + str(count) + " categories"
    def select1(self, sender, e):
        """Select all categories"""
        for i in range(self.list1.Items.Count):
            self.list1.SetItemChecked(i, True)
        self.UpdateCategoryCount()
    def deselect1(self, sender, e):
        """Deselect all categories"""
        for i in range(self.list1.Items.Count):
            self.list1.SetItemChecked(i, False)
        self.UpdateCategoryCount()
    def filter1(self, sender, e):
        """Select all currently filtered/visible categories"""
        for i in range(self.list1.Items.Count):
            self.list1.SetItemChecked(i, True)
        self.UpdateCategoryCount()
    def GetSelectedCategories(self):
        """Get list of selected categories"""
        selected = []
        for i in range(self.list1.Items.Count):
            if self.list1.GetItemChecked(i):
                selected.append(str(self.list1.Items[i]))
        return selected
    def search2(self, sender, e):
        """Handle parameter search text changes"""
        try:
            if not self.input2:
                return
            search_text = self.input2.Text.lower()
            if search_text == "search parameters..." or not search_text:
                self.data2 = self.data1[:]
            else:
                self.data2 = []
                for param in self.data1:
                    if search_text in param.lower():
                        self.data2.append(param)
            self.updatelist2()
            self.updateinfo1()
        except:
            pass
    def clear2(self, sender, e):
        """Clear parameter search text"""
        try:
            if self.input2:
                self.input2.Text = "Search parameters..."
                self.input2.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
                self.data2 = self.data1[:]
                self.updatelist2()
                self.updateinfo1()
        except:
            pass
    def select2(self, sender, e):
        """Select all parameters"""
        for i in range(self.list2.Items.Count):
            self.list2.SetItemChecked(i, True)
        self.UpdateParameterCount()
    def deselect2(self, sender, e):
        """Deselect all parameters"""
        for i in range(self.list2.Items.Count):
            self.list2.SetItemChecked(i, False)
        self.UpdateParameterCount()
    def filter2(self, sender, e):
        """Select all currently filtered/visible parameters"""
        for i in range(self.list2.Items.Count):
            self.list2.SetItemChecked(i, True)
        self.UpdateParameterCount()
    def GetMainCategory(self, element):
        """Get main Revit category, filter out subcategories"""
        if not element or not element.Category:
            return None
        category = element.Category
        parent = category.Parent
        if parent:
            mep_categories = [
                "Mechanical Equipment", "Electrical Equipment", "Plumbing Fixtures",
                "Air Terminals", "Sprinklers", "Electrical Fixtures", "Communication Devices"
            ]
            if parent.Name in mep_categories:
                return category
            else:
                return parent
        else:
            return category
    def export1(self, sender, e):
        """Export elements by selected categories to Excel - ENHANCED WITH CATEGORY FILTERING"""
        if self.flag1:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
        sel1 = self.GetSelectedCategories()
        selected_params = self.GetSelectedParameters()
        if len(sel1) == 0:
            WinForms.MessageBox.Show("Please select at least one category to export.", "DynaNex - Warning")
            return
        if len(selected_params) == 0:
            WinForms.MessageBox.Show("Please select at least one parameter to export.", "DynaNex - Warning")
            return
        try:
            self.flag1 = True
            self.progress1(True, "Preparing export...")
            if self.mode == "ActiveView":
                current_view = self.uidoc.ActiveView
                collector = FilteredElementCollector(self.doc, current_view.Id)
                all_elements = collector.WhereElementIsNotElementType().ToElements()
                scope_name = current_view.Name
            else:
                collector = FilteredElementCollector(self.doc)
                all_elements = collector.WhereElementIsNotElementType().ToElements()
                scope_name = "WholeModel"
            if len(all_elements) == 0:
                WinForms.MessageBox.Show("No elements found in current scope.", "DynaNex - Warning")
                return
            self.progress1(True, "Filtering elements by selected categories...")
            category_elements = {}
            for element in all_elements:
                main_category = self.GetMainCategory(element)
                if main_category and main_category.Name in sel1:
                    category_name = main_category.Name
                    if category_name not in category_elements:
                        category_elements[category_name] = []
                    category_elements[category_name].append(element)
                WinForms.Application.DoEvents()
            if len(category_elements) == 0:
                WinForms.MessageBox.Show("No elements found in selected categories.", "DynaNex - Warning")
                return
            save_dialog = WinForms.SaveFileDialog()
            save_dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            save_dialog.FileName = "DynaNex_Export_" + scope_name.replace(" ", "_") + "_" + System.DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
            save_dialog.Title = "Save DynaNex Export"
            if save_dialog.ShowDialog() != DialogResult.OK:
                self.flag1 = False
                self.progress1(False)
                return
            self.CreateCategoryBasedExcel(category_elements, selected_params, save_dialog.FileName)
        except Exception as e:
            WinForms.MessageBox.Show("Export failed: " + str(e), "DynaNex - Error")
        finally:
            self.flag1 = False
            self.progress1(False)
    def CreateCategoryBasedExcel(self, category_elements, parameters, file_path):
        """Create Excel file with separate sheets for each category - OPTIMIZED"""
        excel_app = None
        try:
            self.progress1(True, "Creating Excel application...")
            excel_app = Excel.ApplicationClass()
            excel_app.Visible = False
            excel_app.ScreenUpdating = False
            excel_app.DisplayAlerts = False
            workbook = excel_app.Workbooks.Add()
            while workbook.Worksheets.Count > 1:
                workbook.Worksheets[workbook.Worksheets.Count].Delete()
            first_sheet = True
            total_categories = len(category_elements)
            current_category = 0
            for category_name, elements in category_elements.items():
                current_category += 1
                if self.panel1:
                    self.panel1.Visible = True
                    self.fill1.Width = 0
                    self.text1.Text = "0%"
                self.label5.Text = "Processing " + category_name + " (" + str(current_category) + "/" + str(total_categories) + ")..."
                if first_sheet:
                    worksheet = workbook.ActiveSheet
                    first_sheet = False
                else:
                    worksheet = workbook.Worksheets.Add()
                safe_name = self.MakeExcelSafeName(category_name)
                worksheet.Name = safe_name[:31]
                headers = ["Element ID", "Category", "Family", "Type", "Level", "Phase"] + parameters
                for col, header in enumerate(headers):
                    cell = worksheet.Cells[1, col + 1]
                    cell.Value2 = header
                    cell.Font.Bold = True
                    cell.Interior.Color = 15917529
                    cell.Borders.Weight = 1
                row = 2
                batch_size = 25
                processed = 0
                for i in range(0, len(elements), batch_size):
                    batch = elements[i:i + batch_size]
                    for element in batch:
                        try:
                            worksheet.Cells[row, 1].Value2 = element.Id.IntegerValue
                            worksheet.Cells[row, 2].Value2 = category_name
                            family_name, type_name = self.GetFamilyTypeInfo(element)
                            worksheet.Cells[row, 3].Value2 = family_name
                            worksheet.Cells[row, 4].Value2 = type_name
                            level_name = self.GetElementLevel(element)
                            phase_name = self.GetElementPhase(element)
                            worksheet.Cells[row, 5].Value2 = level_name
                            worksheet.Cells[row, 6].Value2 = phase_name
                            for col, param_name in enumerate(parameters):
                                value = self.GetParameterValue(element, param_name)
                                worksheet.Cells[row, 7 + col].Value2 = value
                            row += 1
                            processed += 1
                        except:
                            continue
                    progress = int((processed * 100) / len(elements))
                    self.update1(progress, category_name + ": " + str(progress) + "% (" + str(processed) + "/" + str(len(elements)) + ")")
                    WinForms.Application.DoEvents()
                if row > 2:
                    worksheet.Columns.AutoFit()
                    data_range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row - 1, len(headers)]]
                    data_range.Borders.Weight = 1
                    header_range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, len(headers)]]
                    header_range.AutoFilter()
                    worksheet.Range["A2"].Select()
                    excel_app.ActiveWindow.FreezePanes = True
            self.progress1(True, "Saving Excel file...")
            workbook.SaveAs(file_path)
            excel_app.Visible = True
            excel_app.ScreenUpdating = True
            excel_app.DisplayAlerts = True
            total_elements = sum(len(elements) for elements in category_elements.values())
            scope_text = "Active View" if self.mode == "ActiveView" else "Whole Model"
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
    def import1(self, sender, e):
        """Import parameters from Excel - ENHANCED with category support"""
        if self.flag1:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
        try:
            open_dialog = WinForms.OpenFileDialog()
            open_dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
            open_dialog.Title = "Select DynaNex Excel File to Import"
            if open_dialog.ShowDialog() != DialogResult.OK:
                return
            self.flag1 = True
            self.progress1(True, "Opening Excel file...")
            excel_app = Excel.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            workbook = excel_app.Workbooks.Open(open_dialog.FileName)
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
                self.flag1 = False
                self.progress1(False)
                return
            self.progress1(True, "Starting Revit transaction...")
            TransactionManager.Instance.EnsureInTransaction(self.doc)
            try:
                total_updated = 0
                total_errors = 0
                for sheet_index in range(1, workbook.Worksheets.Count + 1):
                    worksheet = workbook.Worksheets[sheet_index]
                    sheet_name = worksheet.Name
                    self.progress1(True, "Importing from: " + sheet_name + " (" + str(sheet_index) + "/" + str(workbook.Worksheets.Count) + ")")
                    updated, errors = self.ImportFromWorksheetFixed(worksheet)
                    total_updated += updated
                    total_errors += errors
                    WinForms.Application.DoEvents()
                TransactionManager.Instance.TransactionTaskDone()
                workbook.Close(False)
                excel_app.Quit()
                message = "✅ Import completed successfully!\n\n"
                message += "📊 Elements updated: " + str(total_updated) + "\n"
                if total_errors > 0:
                    message += "⚠️ Elements with errors: " + str(total_errors) + "\n"
                message += "\n🔄 Parameters have been updated in your Revit model!"
                WinForms.MessageBox.Show(message, "DynaNex - Import Complete")
            except Exception as transaction_error:
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
            self.flag1 = False
            self.progress1(False)
    def ImportFromWorksheetFixed(self, worksheet):
        """FIXED IMPORT - Works with category filtering support"""
        updated_count = 0
        error_count = 0
        try:
            used_range = worksheet.UsedRange
            if not used_range:
                return updated_count, error_count
            last_row = used_range.Rows.Count
            last_col = used_range.Columns.Count
            if last_row < 2:
                return updated_count, error_count
            headers = []
            for col in range(1, last_col + 1):
                cell_value = worksheet.Cells[1, col].Value2
                headers.append(str(cell_value) if cell_value else "")
            if "Element ID" not in headers:
                return updated_count, error_count
            element_id_col = headers.index("Element ID") + 1
            standard_cols = ["Element ID", "Category", "Family", "Type", "Level", "Phase"]
            param_columns = {}
            for i, header in enumerate(headers):
                if header and header not in standard_cols:
                    param_columns[header] = i + 1
            if len(param_columns) == 0:
                return updated_count, error_count
            for row in range(2, last_row + 1):
                try:
                    element_id_cell = worksheet.Cells[row, element_id_col].Value2
                    if not element_id_cell:
                        continue
                    try:
                        element_id_int = int(float(element_id_cell))
                        element_id = ElementId(element_id_int)
                        element = self.doc.GetElement(element_id)
                        if not element:
                            error_count += 1
                            continue
                    except:
                        error_count += 1
                        continue
                    element_was_updated = False
                    for param_name, col_index in param_columns.items():
                        try:
                            cell_value = worksheet.Cells[row, col_index].Value2
                            if cell_value is None:
                                continue
                            cell_str = str(cell_value).strip()
                            if cell_str == "":
                                continue
                            param = element.LookupParameter(param_name)
                            if not param:
                                continue
                            if param.IsReadOnly:
                                continue
                            success = self.SetParameterValue(param, cell_value)
                            if success:
                                element_was_updated = True
                        except Exception as param_error:
                            continue
                    if element_was_updated:
                        updated_count += 1
                    if row % 10 == 2:
                        progress_pct = int(((row - 1) * 100) / (last_row - 1))
                        self.update1(progress_pct, "Updating elements... " + str(progress_pct) + "%")
                        WinForms.Application.DoEvents()
                except Exception as row_error:
                    error_count += 1
                    continue
        except Exception as sheet_error:
            pass
        return updated_count, error_count
    def SetParameterValue(self, param, value):
        """Set parameter value with proper type conversion"""
        try:
            if param.StorageType == StorageType.String:
                param.Set(str(value))
                return True
            elif param.StorageType == StorageType.Integer:
                if param.Definition.ParameterType == ParameterType.YesNo:
                    value_str = str(value).lower()
                    if value_str in ['yes', 'true', '1', 'y', 'on']:
                        param.Set(1)
                    else:
                        param.Set(0)
                    return True
                else:
                    int_value = int(float(value))
                    param.Set(int_value)
                    return True
            elif param.StorageType == StorageType.Double:
                double_value = float(value)
                param.Set(double_value)
                return True
            elif param.StorageType == StorageType.ElementId:
                if str(value).isdigit():
                    ref_id = ElementId(int(float(value)))
                    ref_element = self.doc.GetElement(ref_id)
                    if ref_element:
                        param.Set(ref_id)
                        return True
        except Exception as set_error:
            return False
        return False
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
        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        safe_name = name
        for char in invalid_chars:
            safe_name = safe_name.replace(char, '_')
        return safe_name.strip()
    def refresh1(self, sender, e):
        """Refresh categories and parameters from current scope"""
        if self.flag1:
            WinForms.MessageBox.Show("Please wait for current operation to complete.", "DynaNex - Processing")
            return
        self.process1()
        scope_text = "Active View" if self.mode == "ActiveView" else "Whole Model"
        WinForms.MessageBox.Show("✅ " + scope_text + " refreshed successfully!", "DynaNex - Refresh Complete")
    def open1(self, sender, e):
        """Open developer's LinkedIn profile in browser"""
        try:
            linkedin_url = "https://my.linkedin.com/in/nadeemahmed-nbim"
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
            import subprocess
            subprocess.Popen(['start', linkedin_url], shell=True)
        except Exception as ex:
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
    def theme1(self, sender, e):
        """Switch to dark theme"""
        if self.theme != "dark":
            self.theme = "dark"
            self.ApplyTheme()
            self.UpdateThemeButtonStates()
    def theme2(self, sender, e):
        """Switch to light theme"""
        if self.theme != "light":
            self.theme = "light"
            self.ApplyTheme()
            self.UpdateThemeButtonStates()
    def UpdateThemeButtonStates(self):
        """Update theme button visual states"""
        try:
            if self.theme == "dark":
                self.btn3.FlatAppearance.BorderSize = 2
                self.btn3.FlatAppearance.BorderColor = Drawing.Color.White
                self.btn4.FlatAppearance.BorderSize = 1
                self.btn4.FlatAppearance.BorderColor = Drawing.Color.Gray
            else:
                self.btn3.FlatAppearance.BorderSize = 1
                self.btn3.FlatAppearance.BorderColor = Drawing.Color.Gray
                self.btn4.FlatAppearance.BorderSize = 2
                self.btn4.FlatAppearance.BorderColor = Drawing.Color.Black
        except:
            pass
    def ApplyTheme(self):
        """Apply the selected theme to all UI elements"""
        try:
            if self.theme == "dark":
                self.ApplyDarkTheme()
            else:
                self.ApplyLightTheme()
        except Exception as ex:
            WinForms.MessageBox.Show("Error applying theme: " + str(ex), "Theme Error")
    def ApplyDarkTheme(self):
        """Apply dark blue theme colors with uniform background"""
        try:
            colors = {
                'main_bg': Drawing.Color.FromArgb(30, 41, 59),
                'panel_bg': Drawing.Color.FromArgb(30, 41, 59),
                'input_bg': Drawing.Color.FromArgb(71, 85, 105),
                'text_light': Drawing.Color.FromArgb(220, 220, 220),
                'text_medium': Drawing.Color.FromArgb(156, 163, 175)
            }
            self.BackColor = colors['main_bg']
            self.UpdateControlColors(self, colors, True)
            self.SuspendLayout()
            self.ResumeLayout(True)
            self.Refresh()
            self.Invalidate(True)
        except Exception as ex:
            pass
    def ApplyLightTheme(self):
        """Apply light theme colors with uniform background"""
        try:
            colors = {
                'main_bg': Drawing.Color.FromArgb(248, 250, 252),
                'panel_bg': Drawing.Color.FromArgb(248, 250, 252),
                'input_bg': Drawing.Color.White,
                'text_light': Drawing.Color.FromArgb(51, 65, 85),
                'text_medium': Drawing.Color.FromArgb(107, 114, 128)
            }
            self.BackColor = colors['main_bg']
            self.UpdateControlColors(self, colors, False)
            self.SuspendLayout()
            self.ResumeLayout(True)
            self.Refresh()
            self.Invalidate(True)
        except Exception as ex:
            pass
    def UpdateControlColors(self, control, colors, is_dark):
        """Recursively update all control colors for uniform appearance"""
        try:
            if (control == self.btn3 or control == self.btn4):
                return
            if hasattr(control, 'BackColor'):
                if isinstance(control, WinForms.Form):
                    control.BackColor = colors['main_bg']
                elif isinstance(control, WinForms.TableLayoutPanel):
                    try:
                        control.BackColor = colors['main_bg']
                    except:
                        control.BackColor = colors['panel_bg']
                elif isinstance(control, WinForms.Panel):
                    control.BackColor = colors['panel_bg']
                elif isinstance(control, WinForms.TextBox) or isinstance(control, WinForms.CheckedListBox):
                    control.BackColor = colors['input_bg']
                    if hasattr(control, 'ForeColor'):
                        control.ForeColor = colors['text_light']
                elif isinstance(control, WinForms.Label):
                    try:
                        current_bg = str(control.BackColor)
                        if 'Transparent' not in current_bg:
                            if isinstance(control.Parent, WinForms.TableLayoutPanel):
                                control.BackColor = colors['main_bg']
                            else:
                                control.BackColor = colors['panel_bg']
                    except:
                        control.BackColor = colors['panel_bg']
                    if hasattr(control, 'ForeColor'):
                        control.ForeColor = colors['text_light']
                elif isinstance(control, WinForms.GroupBox):
                    control.BackColor = colors['panel_bg']
                    if hasattr(control, 'ForeColor'):
                        control.ForeColor = colors['text_light']
                elif isinstance(control, WinForms.CheckBox):
                    control.BackColor = colors['panel_bg']
                    if hasattr(control, 'ForeColor'):
                        control.ForeColor = colors['text_light']
                elif isinstance(control, WinForms.Button):
                    if not (hasattr(control, 'Text') and control.Text in ["Save As", "Load As", "About Me"]):
                        pass
                elif control == self.panel1:
                    if is_dark:
                        control.BackColor = Drawing.Color.FromArgb(51, 65, 85)
                    else:
                        control.BackColor = Drawing.Color.FromArgb(229, 231, 235)
                    control.BorderStyle = WinForms.BorderStyle.FixedSingle
                    control.Invalidate()
                elif control == self.fill1:
                    if is_dark:
                        control.BackColor = Drawing.Color.FromArgb(34, 197, 94)
                    else:
                        control.BackColor = Drawing.Color.FromArgb(22, 163, 74)
                    control.BringToFront()
                    control.Invalidate()
                elif control == self.text1:
                    if is_dark:
                        control.BackColor = colors['main_bg']
                        control.ForeColor = Drawing.Color.White
                    else:
                        control.BackColor = colors['main_bg']
                        control.ForeColor = Drawing.Color.FromArgb(51, 65, 85)
                elif control == self.action_panel:
                    control.BackColor = colors['main_bg']
                    control.Invalidate()
                else:
                    try:
                        control.BackColor = colors['panel_bg']
                        if hasattr(control, 'ForeColor'):
                            control.ForeColor = colors['text_light']
                    except:
                        pass
            if hasattr(control, 'Controls'):
                for child_control in control.Controls:
                    self.UpdateControlColors(child_control, colors, is_dark)
        except Exception as ex:
            pass
    def RefreshProgressBarColors(self):
        """Refresh progress bar colors based on current theme"""
        try:
            is_dark = (self.theme == "dark")
            if hasattr(self, 'action_panel'):
                if is_dark:
                    self.action_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)
                else:
                    self.action_panel.BackColor = Drawing.Color.FromArgb(248, 250, 252)
            if hasattr(self, 'panel1'):
                if is_dark:
                    self.panel1.BackColor = Drawing.Color.FromArgb(51, 65, 85)
                else:
                    self.panel1.BackColor = Drawing.Color.FromArgb(229, 231, 235)
            if hasattr(self, 'fill1'):
                if is_dark:
                    self.fill1.BackColor = Drawing.Color.FromArgb(34, 197, 94)
                else:
                    self.fill1.BackColor = Drawing.Color.FromArgb(22, 163, 74)
                self.fill1.BringToFront()
            if hasattr(self, 'text1'):
                if is_dark:
                    self.text1.BackColor = Drawing.Color.FromArgb(30, 41, 59)
                    self.text1.ForeColor = Drawing.Color.White
                else:
                    self.text1.BackColor = Drawing.Color.FromArgb(248, 250, 252)
                    self.text1.ForeColor = Drawing.Color.FromArgb(51, 65, 85)
            if hasattr(self, 'panel1'):
                self.panel1.Invalidate()
            if hasattr(self, 'fill1'):
                self.fill1.Invalidate()
            if hasattr(self, 'action_panel'):
                self.action_panel.Invalidate()
        except Exception as ex:
            pass
    def help1(self, sender, e):
        """Show help dialog"""
        help_text = """⚡ DynaNex V1.0 - User Guide
"Power your Connections"
📋 WHAT IS DYNANEX:
    DynaNex is an advanced parameters export/import tool for Revit.
    It allows you to export element parameters to Excel by category,
    edit them in Excel, and import them back to update your Revit model.
    Seamlessly connects Revit with Microsoft Excel for efficient parameter management.
🚀 STEP-BY-STEP USAGE GUIDE:
🎯 STEP 1: ELEMENT SCOPE SELECTION
    Choose your working scope:
    • 🔍 Active View Elements
      ↳ Only elements visible in your current view
      ↳ Faster processing for focused work
    • 🏢 Whole Model Elements
      ↳ All elements throughout the entire model
      ↳ More comprehensive but slower processing
    Switch between scopes using the buttons at the top of the interface.
📂 STEP 2: CATEGORY SELECTION
    Select the categories you want to work with:
    • Check the categories you need (Walls, Floors, Doors, etc.)
    • Only selected categories will be included in the export
    • Each category gets its own Excel sheet for organization
    • Use the search box to quickly find specific categories
    • Use selection buttons for quick category management
🔧 STEP 3: PARAMETER SELECTION
    Choose which parameters to export:
    • Check the parameters you want to modify in Excel
    • Parameters are filtered based on your selected categories
    • Use the search box to find specific parameters quickly
    • Use selection buttons for quick parameter management
1️⃣ STEP 4: EXPORT TO EXCEL
    Export your data:
    • Verify your scope, categories, and parameters are selected
    • Click the '📊 Export to Excel' button
    • Choose a location to save your Excel file
    • Excel file is created with separate sheets for each category
    • Each element has a unique Element ID for tracking
2️⃣ STEP 5: EDIT IN EXCEL
    Modify your parameters in Excel:
    • Open the exported Excel file
    • Each selected category has its own sheet tab
    • Edit parameter values in the appropriate columns
    •
    ⚠️ CRITICAL: DO NOT modify these columns:
       - Element ID
       - Category
       - Family
       - Type
    • Save the Excel file when finished editing
3️⃣ STEP 6: IMPORT BACK TO REVIT
    Update Revit with your changes:
    • Click the '📥 Import from Excel' button
    • Select your modified Excel file
    • All parameter values will be updated in Revit
    • The tool safely overrides existing values
    • Progress bar shows import status
📊 CATEGORY HANDLING DETAILS:
    • Select only the categories you need for focused exports
    • Exports MAIN Revit categories (not subcategories)
    • Example: "Railings" category (not "Top Rail" subcategory)
    • Exception: MEP subcategories are preserved when useful
    • Each selected category gets its own Excel sheet for clarity
⚠️ IMPORTANT SAFETY NOTES:
    • Always backup your Revit model before importing!
    • Only modify parameter value columns in Excel
    • Element ID column links Excel rows to Revit elements - never change it
    • Make sure Excel file is closed before importing
    • Tool uses Revit transactions for safe, undoable updates
💡 PERFORMANCE & EFFICIENCY TIPS:
    • Use 'Refresh Scope' button to update lists after model changes
    • Search boxes help find categories and parameters quickly
    • Active View mode is significantly faster than Whole Model
    • Select only needed categories for faster exports and smaller files
    • Whole Model mode samples elements intelligently for better performance
🏢 DynaNex - "Power your Connections"
   Advanced Revit-Excel Parameter Management Tool"""
        help_form = WinForms.Form()
        help_form.Text = "DynaNex - Help"
        help_form.Size = Drawing.Size(850, 750)
        help_form.StartPosition = WinForms.FormStartPosition.CenterParent
        help_form.MaximizeBox = True
        help_form.MinimizeBox = False
        help_form.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        help_textbox = WinForms.RichTextBox()
        help_textbox.ReadOnly = True
        help_textbox.ScrollBars = WinForms.RichTextBoxScrollBars.Vertical
        help_textbox.Dock = WinForms.DockStyle.Fill
        help_textbox.Font = Drawing.Font("Segoe UI", 11)
        help_textbox.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        help_textbox.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        help_textbox.WordWrap = True
        help_textbox.DetectUrls = False
        help_textbox.Text = help_text
        help_textbox.Select(0, 0)
        bold_patterns = [
            "⚡ DynaNex V1.0 - User Guide",
            "📋 WHAT IS DYNANEX:",
            "🚀 STEP-BY-STEP USAGE GUIDE:",
            "🎯 STEP 1: ELEMENT SCOPE SELECTION",
            "📂 STEP 2: CATEGORY SELECTION",
            "🔧 STEP 3: PARAMETER SELECTION",
            "1️⃣ STEP 4: EXPORT TO EXCEL",
            "2️⃣ STEP 5: EDIT IN EXCEL",
            "3️⃣ STEP 6: IMPORT BACK TO REVIT",
            "📊 CATEGORY HANDLING DETAILS:",
            "⚠️ IMPORTANT SAFETY NOTES:",
            "💡 PERFORMANCE & EFFICIENCY TIPS:",
            "🏢 DynaNex - \"Power your Connections\""
        ]
        for pattern in bold_patterns:
            start = help_textbox.Text.find(pattern)
            if start >= 0:
                help_textbox.Select(start, len(pattern))
                help_textbox.SelectionFont = Drawing.Font("Segoe UI", 11, Drawing.FontStyle.Bold)
        help_textbox.Select(0, 0)
        help_form.Controls.Add(help_textbox)
        help_form.ShowDialog()
    def updatelist2(self):
        """Update the parameters checklist with filtered results while preserving global selections"""
        try:
            if self.list2 and hasattr(self.list2, 'Items'):
                for i in range(self.list2.Items.Count):
                    item_text = str(self.list2.Items[i])
                    if self.list2.GetItemChecked(i):
                        self.sel2.add(item_text)
                    elif item_text in self.sel2:
                        self.sel2.remove(item_text)
                self.list2.Items.Clear()
                for param in self.data2:
                    index = self.list2.Items.Add(param)
                    if param in self.sel2:
                        self.list2.SetItemChecked(index, True)
        except:
            pass
    def updateinfo1(self):
        """Update the information label with counts"""
        total_cats = len(self.data3)
        filtered_cats = len(self.data4)
        total_params = len(self.data1)
        filtered_params = len(self.data2)
        scope_text = "active view" if self.mode == "ActiveView" else "whole model"
        element_count_text = ""
        if hasattr(self, 'count1'):
            element_count_text = " | 📊 " + str(self.count1) + " elements in " + scope_text
        info_text = "Categories: " + str(filtered_cats)
        if filtered_cats < total_cats:
            info_text += " of " + str(total_cats) + " (filtered)"
        info_text += " | Parameters: " + str(filtered_params)
        if filtered_params < total_params:
            info_text += " of " + str(total_params) + " (filtered)"
        info_text += " (from " + scope_text + ")" + element_count_text
        self.label4.Text = info_text
    def check2(self, sender, e):
        """Update parameter selection count and global state when parameters are checked/unchecked"""
        try:
            if e.Index >= 0 and e.Index < self.list2.Items.Count:
                item_text = str(self.list2.Items[e.Index])
                if e.NewValue == System.Windows.Forms.CheckState.Checked:
                    self.sel2.add(item_text)
                else:
                    self.sel2.discard(item_text)
        except:
            pass
        self.BeginInvoke(System.Action(self.UpdateParameterCount))
    def UpdateParameterCount(self):
        """Update the parameter selection count label"""
        count = 0
        for i in range(self.list2.Items.Count):
            if self.list2.GetItemChecked(i):
                count += 1
        self.label2.Text = "Selected: " + str(count) + " parameters"
    def GetSelectedParameters(self):
        """Get list of selected parameters"""
        selected = []
        for i in range(self.list2.Items.Count):
            if self.list2.GetItemChecked(i):
                selected.append(str(self.list2.Items[i]))
        return selected
    def progress1(self, show=True, message="Processing...", progress_percent=0):
        """Show/hide visual progress bar with percentage fill"""
        try:
            if self.panel1:
                self.panel1.Visible = show
                if show:
                    self.RefreshProgressBarColors()
                    if progress_percent > 0:
                        panel_width = self.panel1.Width - 2
                        fill_width = int((progress_percent / 100.0) * panel_width)
                        self.fill1.Width = max(0, min(fill_width, panel_width))
                        self.text1.Text = str(progress_percent) + "%"
                    else:
                        self.fill1.Width = 0
                        self.text1.Text = "0%"
                else:
                    self.fill1.Width = 0
                    self.text1.Text = "0%"
            if self.label5:
                if show:
                    self.label5.Text = message
                else:
                    self.label5.Text = "Ready"
        except:
            pass
    def update1(self, progress_percent, message="Processing..."):
        """Update progress bar with specific percentage"""
        try:
            if self.panel1:
                self.panel1.Visible = True
                self.RefreshProgressBarColors()
                panel_width = self.panel1.Width - 2
                fill_width = int((progress_percent / 100.0) * panel_width)
                self.fill1.Width = max(0, min(fill_width, panel_width))
                self.text1.Text = str(progress_percent) + "%"
                if self.label5:
                    self.label5.Text = message
                self.panel1.Invalidate()
                WinForms.Application.DoEvents()
        except:
            pass
    def config1(self, sender, e):
        """Show DynaNex settings dialog"""
        settings_form = DynaNextSettingsForm(self)
        result = settings_form.ShowDialog()
        if result == DialogResult.OK and settings_form.settings_changed:
            WinForms.MessageBox.Show(
                "Settings have been saved successfully!",
                "DynaNex - Settings Saved",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Information
            )
    def close1(self, sender, e):
        """Close the form"""
        if self.flag1:
            result = WinForms.MessageBox.Show(
                "Operation in progress. Are you sure you want to close?",
                "DynaNex - Confirm Close",
                WinForms.MessageBoxButtons.YesNo
            )
            if result != DialogResult.Yes:
                return
        self.Close()
    def save1(self, sender, e):
        """Save current settings to user-selected local file"""
        try:
            save_dialog = WinForms.SaveFileDialog()
            save_dialog.Filter = "JSON Settings Files (*.json)|*.json"
            save_dialog.FileName = "DynaNex_Settings_" + System.DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".json"
            save_dialog.Title = "Save DynaNex Settings As"
            save_dialog.DefaultExt = "json"
            save_dialog.AddExtension = True
            if save_dialog.ShowDialog() != DialogResult.OK:
                return
            sel1 = self.GetSelectedCategories()
            sel2 = self.GetSelectedParameters()
            settings_data = {
                'app_settings': self.settings.copy(),
                'ui_state': {
                    'mode': self.mode,
                    'sel1': sel1,
                    'sel2': sel2,
                    'category_search': self.input1.Text if self.input1.Text != "Search categories..." else "",
                    'parameter_search': self.input2.Text if self.input2.Text != "Search parameters..." else ""
                },
                'version': '1.0',
                'timestamp': System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            }
            with open(save_dialog.FileName, 'w') as f:
                json.dump(settings_data, f, indent=4)
            self.label3.Text = "Settings: Saved"
            self.label3.ForeColor = Drawing.Color.FromArgb(34, 197, 94)
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
    def load1(self, sender=None, e=None):
        """Load settings from user-selected local file"""
        try:
            if sender is None:
                if os.path.exists(self.path1):
                    try:
                        with open(self.path1, 'r') as f:
                            settings_data = json.load(f)
                        if 'app_settings' in settings_data:
                            self.settings.update(settings_data['app_settings'])
                        if self.label3:
                            self.label3.Text = "Settings: Default"
                            self.label3.ForeColor = Drawing.Color.FromArgb(156, 163, 175)
                        return
                    except:
                        pass
                if self.label3:
                    self.label3.Text = "Settings: Default"
                    self.label3.ForeColor = Drawing.Color.FromArgb(108, 117, 125)
                return
            open_dialog = WinForms.OpenFileDialog()
            open_dialog.Filter = "JSON Settings Files (*.json)|*.json"
            open_dialog.Title = "Load DynaNex Settings From"
            open_dialog.DefaultExt = "json"
            if open_dialog.ShowDialog() != DialogResult.OK:
                return
            with open(open_dialog.FileName, 'r') as f:
                settings_data = json.load(f)
            if 'app_settings' in settings_data:
                self.settings.update(settings_data['app_settings'])
            if 'ui_state' in settings_data:
                ui_state = settings_data['ui_state']
                if 'mode' in ui_state:
                    if ui_state['mode'] == "WholeModel":
                        self.mode = "WholeModel"
                        if self.btn1 and self.btn2:
                            self.btn1.BackColor = Drawing.Color.FromArgb(71, 85, 105)
                            self.btn2.BackColor = Drawing.Color.FromArgb(34, 197, 94)
                    else:
                        self.mode = "ActiveView"
                        if self.btn1 and self.btn2:
                            self.btn1.BackColor = Drawing.Color.FromArgb(34, 197, 94)
                            self.btn2.BackColor = Drawing.Color.FromArgb(71, 85, 105)
                if 'category_search' in ui_state and ui_state['category_search']:
                    if self.input1:
                        self.input1.Text = ui_state['category_search']
                        self.input1.ForeColor = Drawing.Color.Black
                if 'parameter_search' in ui_state and ui_state['parameter_search']:
                    if self.input2:
                        self.input2.Text = ui_state['parameter_search']
                        self.input2.ForeColor = Drawing.Color.Black
                self.process1()
                if 'sel1' in ui_state:
                    self.RestoreSelectedCategories(ui_state['sel1'])
                if 'sel2' in ui_state:
                    self.RestoreSelectedParameters(ui_state['sel2'])
            timestamp = settings_data.get('timestamp', 'Unknown')
            if self.label3:
                self.label3.Text = "Settings: Loaded"
                self.label3.ForeColor = Drawing.Color.FromArgb(59, 130, 246)
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
    def RestoreSelectedCategories(self, sel1):
        """Restore category selections from saved settings"""
        try:
            if not sel1 or not self.list1:
                return
            for i in range(self.list1.Items.Count):
                item_text = str(self.list1.Items[i])
                if item_text in sel1:
                    self.list1.SetItemChecked(i, True)
            self.UpdateCategoryCount()
        except Exception as e:
            pass
    def RestoreSelectedParameters(self, sel2):
        """Restore parameter selections from saved settings"""
        try:
            if not sel2 or not self.list2:
                return
            for i in range(self.list2.Items.Count):
                item_text = str(self.list2.Items[i])
                if item_text in sel2:
                    self.list2.SetItemChecked(i, True)
            self.UpdateParameterCount()
        except Exception as e:
            pass
class DynaNextSettingsForm(WinForms.Form):
    """DynaNex settings dialog"""
    def __init__(self, parent_form=None):
        self.parent_form = parent_form
        self.settings_changed = False
        self.setup1()
        self.LoadCurrentSettings()
    def LoadCurrentSettings(self):
        """Load current settings into the form"""
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
            if not hasattr(self.parent_form, 'settings'):
                self.parent_form.settings = {}
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
    def setup1(self):
        self.Text = "DynaNex Settings"
        self.Size = Drawing.Size(550, 600)
        self.StartPosition = WinForms.FormStartPosition.CenterParent
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        main_panel = WinForms.TableLayoutPanel()
        main_panel.Dock = WinForms.DockStyle.Fill
        main_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        main_panel.RowCount = 6
        main_panel.ColumnCount = 1
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 50))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 100))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 100))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 120))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 100))
        main_panel.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Absolute, 80))
        title_label = WinForms.Label()
        title_label.Text = "⚡ DynaNex Settings"
        title_label.Font = Drawing.Font("Segoe UI", 14, Drawing.FontStyle.Bold)
        title_label.Dock = WinForms.DockStyle.Fill
        title_label.TextAlign = Drawing.ContentAlignment.MiddleCenter
        title_label.ForeColor = Drawing.Color.FromArgb(59, 130, 246)
        title_label.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        scope_group = WinForms.GroupBox()
        scope_group.Text = "Element Scope Options"
        scope_group.Dock = WinForms.DockStyle.Fill
        scope_group.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        scope_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        scope_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        scope_layout = WinForms.TableLayoutPanel()
        scope_layout.Dock = WinForms.DockStyle.Fill
        scope_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        scope_layout.RowCount = 2
        scope_layout.ColumnCount = 1
        self.confirm_whole_model = WinForms.CheckBox()
        self.confirm_whole_model.Text = "Always confirm before switching to Whole Model"
        self.confirm_whole_model.Checked = True
        self.confirm_whole_model.Dock = WinForms.DockStyle.Fill
        self.confirm_whole_model.Font = Drawing.Font("Segoe UI", 9)
        self.confirm_whole_model.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.confirm_whole_model.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        self.default_scope_active = WinForms.CheckBox()
        self.default_scope_active.Text = "Start with Active View scope by default"
        self.default_scope_active.Checked = True
        self.default_scope_active.Dock = WinForms.DockStyle.Fill
        self.default_scope_active.Font = Drawing.Font("Segoe UI", 9)
        self.default_scope_active.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.default_scope_active.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        scope_layout.Controls.Add(self.confirm_whole_model, 0, 0)
        scope_layout.Controls.Add(self.default_scope_active, 0, 1)
        scope_group.Controls.Add(scope_layout)
        category_group = WinForms.GroupBox()
        category_group.Text = "Category Selection Options"
        category_group.Dock = WinForms.DockStyle.Fill
        category_group.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        category_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        category_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        category_layout = WinForms.TableLayoutPanel()
        category_layout.Dock = WinForms.DockStyle.Fill
        category_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        category_layout.RowCount = 2
        category_layout.ColumnCount = 1
        self.auto_select_common_cats = WinForms.CheckBox()
        self.auto_select_common_cats.Text = "Auto-select common categories (Walls, Floors, etc.)"
        self.auto_select_common_cats.Checked = False
        self.auto_select_common_cats.Dock = WinForms.DockStyle.Fill
        self.auto_select_common_cats.Font = Drawing.Font("Segoe UI", 9)
        self.auto_select_common_cats.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.auto_select_common_cats.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        self.use_main_categories = WinForms.CheckBox()
        self.use_main_categories.Text = "Group by main categories only (recommended)"
        self.use_main_categories.Checked = True
        self.use_main_categories.Dock = WinForms.DockStyle.Fill
        self.use_main_categories.Font = Drawing.Font("Segoe UI", 9)
        self.use_main_categories.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.use_main_categories.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        category_layout.Controls.Add(self.auto_select_common_cats, 0, 0)
        category_layout.Controls.Add(self.use_main_categories, 0, 1)
        category_group.Controls.Add(category_layout)
        export_group = WinForms.GroupBox()
        export_group.Text = "Export Options"
        export_group.Dock = WinForms.DockStyle.Fill
        export_group.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        export_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        export_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        export_layout = WinForms.TableLayoutPanel()
        export_layout.Dock = WinForms.DockStyle.Fill
        export_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        export_layout.RowCount = 3
        export_layout.ColumnCount = 1
        self.include_empty_params = WinForms.CheckBox()
        self.include_empty_params.Text = "Include empty parameter values"
        self.include_empty_params.Checked = True
        self.include_empty_params.Dock = WinForms.DockStyle.Fill
        self.include_empty_params.Font = Drawing.Font("Segoe UI", 9)
        self.include_empty_params.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.include_empty_params.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        self.auto_open_excel = WinForms.CheckBox()
        self.auto_open_excel.Text = "Automatically open Excel after export"
        self.auto_open_excel.Checked = True
        self.auto_open_excel.Dock = WinForms.DockStyle.Fill
        self.auto_open_excel.Font = Drawing.Font("Segoe UI", 9)
        self.auto_open_excel.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.auto_open_excel.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        self.export_only_selected = WinForms.CheckBox()
        self.export_only_selected.Text = "Export only selected categories (recommended)"
        self.export_only_selected.Checked = True
        self.export_only_selected.Dock = WinForms.DockStyle.Fill
        self.export_only_selected.Font = Drawing.Font("Segoe UI", 9)
        self.export_only_selected.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.export_only_selected.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        export_layout.Controls.Add(self.include_empty_params, 0, 0)
        export_layout.Controls.Add(self.auto_open_excel, 0, 1)
        export_layout.Controls.Add(self.export_only_selected, 0, 2)
        export_group.Controls.Add(export_layout)
        import_group = WinForms.GroupBox()
        import_group.Text = "Import Options"
        import_group.Dock = WinForms.DockStyle.Fill
        import_group.Font = Drawing.Font("Arial", 11, Drawing.FontStyle.Bold)
        import_group.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        import_group.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        import_layout = WinForms.TableLayoutPanel()
        import_layout.Dock = WinForms.DockStyle.Fill
        import_layout.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        import_layout.RowCount = 3
        import_layout.ColumnCount = 1
        self.override_existing = WinForms.CheckBox()
        self.override_existing.Text = "Override existing parameter values (standard mode)"
        self.override_existing.Checked = True
        self.override_existing.Dock = WinForms.DockStyle.Fill
        self.override_existing.Font = Drawing.Font("Segoe UI", 9)
        self.override_existing.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.override_existing.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        self.skip_readonly = WinForms.CheckBox()
        self.skip_readonly.Text = "Skip read-only parameters during import"
        self.skip_readonly.Checked = True
        self.skip_readonly.Dock = WinForms.DockStyle.Fill
        self.skip_readonly.Font = Drawing.Font("Segoe UI", 9)
        self.skip_readonly.ForeColor = Drawing.Color.FromArgb(220, 220, 220)
        self.skip_readonly.BackColor = Drawing.Color.FromArgb(51, 65, 85)
        import_layout.Controls.Add(self.override_existing, 0, 0)
        import_layout.Controls.Add(self.skip_readonly, 0, 1)
        import_group.Controls.Add(import_layout)
        button_panel = WinForms.TableLayoutPanel()
        button_panel.Dock = WinForms.DockStyle.Fill
        button_panel.BackColor = Drawing.Color.FromArgb(30, 41, 59)
        button_panel.Height = 40
        button_panel.ColumnCount = 3
        button_panel.RowCount = 1
        button_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 33.33))
        button_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 33.33))
        button_panel.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 33.34))
        button_panel.Padding = WinForms.Padding(15, 20, 15, 15)
        reset_btn = WinForms.Button()
        reset_btn.Text = "Reset Defaults"
        reset_btn.Size = Drawing.Size(100, 25)
        reset_btn.Anchor = WinForms.AnchorStyles.None
        reset_btn.Click += self.ResetDefaults
        reset_btn.BackColor = Drawing.Color.FromArgb(100, 116, 139)
        reset_btn.ForeColor = Drawing.Color.White
        reset_btn.Font = Drawing.Font("Segoe UI", 9)
        reset_btn.FlatStyle = WinForms.FlatStyle.Flat
        reset_btn.FlatAppearance.BorderSize = 1
        reset_btn.FlatAppearance.BorderColor = Drawing.Color.White
        ok_btn = WinForms.Button()
        ok_btn.Text = "OK"
        ok_btn.Size = Drawing.Size(100, 25)
        ok_btn.Anchor = WinForms.AnchorStyles.None
        ok_btn.BackColor = Drawing.Color.FromArgb(34, 197, 94)
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
        cancel_btn.BackColor = Drawing.Color.FromArgb(239, 68, 68)
        cancel_btn.ForeColor = Drawing.Color.White
        cancel_btn.Font = Drawing.Font("Segoe UI", 9)
        cancel_btn.FlatStyle = WinForms.FlatStyle.Flat
        cancel_btn.FlatAppearance.BorderSize = 1
        cancel_btn.FlatAppearance.BorderColor = Drawing.Color.White
        button_panel.Controls.Add(reset_btn, 0, 0)
        button_panel.Controls.Add(ok_btn, 1, 0)
        button_panel.Controls.Add(cancel_btn, 2, 0)
        main_panel.Controls.Add(title_label, 0, 0)
        main_panel.Controls.Add(scope_group, 0, 1)
        main_panel.Controls.Add(category_group, 0, 2)
        main_panel.Controls.Add(export_group, 0, 3)
        main_panel.Controls.Add(import_group, 0, 4)
        main_panel.Controls.Add(button_panel, 0, 5)
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
def ShowDynaNex():
    """Show the DynaNex tool"""
    try:
        if not DocumentManager.Instance.CurrentDBDocument:
            WinForms.MessageBox.Show(
                "No active Revit document found.\n\nPlease open a Revit model first.",
                "DynaNex - No Document",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Warning
            )
            return
        current_view = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument.ActiveView
        if current_view is None:
            WinForms.MessageBox.Show(
                "No active view found.\n\nPlease make sure you have a view open in Revit.",
                "DynaNex - No View",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Warning
            )
            return
        WinForms.MessageBox.Show(
            "DynaNex V1.0 - \"Power your Connections\"\n\n" +
            "Advanced parameter export/import tool\n" +
            "Connects Revit with Microsoft Excel\n" +
            "Click 'Help' button for detailed instructions.",
            "DynaNex V1.0 - Power your Connections",
            WinForms.MessageBoxButtons.OK,
            WinForms.MessageBoxIcon.Information
        )
        form = App1()
        WinForms.Application.EnableVisualStyles()
        WinForms.Application.Run(form)
    except Exception as e:
        error_msg = "Failed to start DynaNex:\n\n" + str(e)
        WinForms.MessageBox.Show(error_msg, "DynaNex - Startup Error",
                               WinForms.MessageBoxButtons.OK,
                               WinForms.MessageBoxIcon.Error)
ShowDynaNex()
