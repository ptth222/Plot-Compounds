import pandas
import os
import re
from matplotlib import use
use('WXAgg')
from matplotlib import pyplot as plt
import wx
import errno
import openpyxl
from collections import OrderedDict


class Plots_GUI(wx.Frame):
    
    
    def __init__(self, parent, title, *args, **kwargs):
        
        super(Plots_GUI, self).__init__(parent, title=title, size = (1000, 500), *args, **kwargs) 
            
        self.InitUI()
        self.Centre()
        self.Show()

        
    def InitUI(self):
 
        self.groups = OrderedDict()
        self.excel_file_okay = False

        ##############
        ## Menu Bar
        ##############
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        
        ## Add items to file menu.
        open_item = fileMenu.Append(wx.ID_OPEN, "&Open", "Open Excel or .csv File")
        self.Bind(wx.EVT_MENU, self.OnOpen, open_item)
        
        fileMenu.AppendSeparator()
        
        quit_item = fileMenu.Append(wx.ID_EXIT, '&Quit', 'Quit Application')
        self.Bind(wx.EVT_MENU, self.OnQuit, quit_item)
        
        
        ## Add file menu to the menu bar.
        menubar.Append(fileMenu, "&File")
        ## Put menu bar in frame.
        self.SetMenuBar(menubar)



        panel = wx.Panel(self)

        vbox = wx.BoxSizer(wx.VERTICAL)
        
        ##############
        ## Type of File Radio Button
        ##############
        self.plot_type_radio_box = wx.RadioBox(panel, label = "Plot Type", choices = ["NMR", "Mass Spec"], pos = (0,0), majorDimension = 1, style = wx.RA_SPECIFY_ROWS)
        self.plot_type_radio_box.SetSelection(0)
        self.current_plot_type = "NMR"
        self.actual_file_type = "NMR"
        self.plot_type_radio_box.Bind(wx.EVT_RADIOBOX, self.onPlotTypeRadioBox)
        
        vbox.Add(self.plot_type_radio_box, flag = wx.ALIGN_LEFT | wx.LEFT, border = 10)
        
        
        ##############
        ## Current Excel File Display
        ##############
        current_excel_file_header = wx.StaticText(panel, label = "Current File:")
        self.current_excel_file_label = wx.StaticText(panel, label = "")
        
        vbox.Add(current_excel_file_header, flag=wx.ALIGN_LEFT | wx.LEFT | wx.TOP, border=10)
        vbox.Add(self.current_excel_file_label, flag=wx.ALIGN_LEFT | wx.LEFT, border=10)
        
        
        ##############
        ## Group and Sample List Display
        ##############
        sizer = wx.FlexGridSizer(2, 2, 0, 10)
        
        st1 = wx.StaticText(panel, label="Groups")
        st2 = wx.StaticText(panel, label="Samples")
        sizer.Add(st1)
        sizer.Add(st2)
        
        self.samples_listbox = wx.ListBox(panel, style = wx.LB_SINGLE | wx.LB_HSCROLL | wx.LB_NEEDED_SB | wx.LB_SORT)
        self.groups_listbox = wx.ListBox(panel, style = wx.LB_SINGLE | wx.LB_HSCROLL | wx.LB_NEEDED_SB)
        self.groups_listbox.Bind(wx.EVT_LISTBOX, self.onListBox)
        sizer.Add(self.groups_listbox, flag = wx.EXPAND)
        sizer.Add(self.samples_listbox, flag = wx.EXPAND)
        
        sizer.AddGrowableCol(0,1)
        sizer.AddGrowableCol(1,1)
        #sizer.AddGrowableRow(0,1)
        sizer.AddGrowableRow(1,1)
        
        vbox.Add(sizer, proportion=1, flag=wx.EXPAND|wx.ALL|wx.ALIGN_CENTER, border=10)


        ##############
        ## Buttons
        ##############        
        add_group_button = wx.Button(panel, label="Add Group")
        add_group_button.Bind(wx.EVT_BUTTON, self.Add_Group)
        
        delete_group_button = wx.Button(panel, label="Delete Group")
        delete_group_button.Bind(wx.EVT_BUTTON, self.Delete_Group)
        
        move_group_to_top_button = wx.Button(panel, label="Move Group To Top")
        move_group_to_top_button.Bind(wx.EVT_BUTTON, self.Move_Group_To_Top)
        
        move_group_to_bottom_button = wx.Button(panel, label="Move Group to Bottom")
        move_group_to_bottom_button.Bind(wx.EVT_BUTTON, self.Move_Group_To_Bottom)
        
        add_sample_button = wx.Button(panel, label="Add Sample")
        add_sample_button.Bind(wx.EVT_BUTTON, self.Add_Sample)
        
        delete_sample_button = wx.Button(panel, label="Delete Sample")
        delete_sample_button.Bind(wx.EVT_BUTTON, self.Delete_Sample)
        
        add_pivot_table_button = wx.Button(panel, label="Add Pivot Table")
        add_pivot_table_button.Bind(wx.EVT_BUTTON, self.Add_Pivot_Table)
        
        create_plots_button = wx.Button(panel, label="Create Plots")
        create_plots_button.Bind(wx.EVT_BUTTON, self.Create_Plots)
        
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        hbox1.Add(add_group_button, flag = wx.ALL, border = 5)
        hbox1.Add(delete_group_button, flag = wx.ALL, border = 5)
        hbox1.Add(move_group_to_top_button, flag = wx.ALL, border = 5)
        hbox1.Add(move_group_to_bottom_button, flag = wx.ALL, border = 5)
        hbox1.Add(add_sample_button, flag = wx.ALL, border = 5)
        hbox1.Add(delete_sample_button, flag = wx.ALL, border = 5)
        hbox1.Add(add_pivot_table_button, flag = wx.ALL, border = 5)
        hbox1.Add(create_plots_button, flag = wx.ALL, border = 5)
        
        vbox.Add(hbox1, flag=wx.ALIGN_CENTER|wx.ALL, border=10)
        
        panel.SetSizer(vbox)
        panel.Fit()

        
    
    
    def onPlotTypeRadioBox(self, event):
        
        radio_box = event.GetEventObject()
        self.current_plot_type = radio_box.GetStringSelection()
    
    
    
    
    
    def OnQuit(self, e):
        self.Close()



        
    def OnOpen(self, event):
        
        current_plot_type = self.current_plot_type
        
        if current_plot_type == "NMR":
            message = "Select NMR Data File (Excel File)"
        elif current_plot_type == "Mass Spec":
            message = "Select Mass Spec File (.csv)"
        else:
            msg_dlg = wx.MessageDialog(None, "Current selected plot type is unknown.", "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            exit(1)
        
        dlg = wx.FileDialog(None, message = message, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)

        if dlg.ShowModal() == wx.ID_OK:
            excel_filepath = dlg.GetPath()
            
            if not re.match(r".*\.xlsx|.*\.xlsm|.*\.xls", excel_filepath) and current_plot_type == "NMR":
                message = "Please select an Excel file."
                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                
            elif not re.match(r".*\.csv", excel_filepath) and current_plot_type == "Mass Spec":
                message = "Please select a csv file."
                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                
            else:
                self.groups = OrderedDict()
                self.groups_listbox.Clear()
                self.samples_listbox.Clear()
                self.new_directory = None
                self.sample_names = None
                self.compounds = None
                self.dss_normalized_df = None
                self.sample_fraction_df = None
                self.standard_nmols_df = None
                self.protein_mass_df = None
                self.pivot_table_df = None
                self.directory_path = None
                
                self.excel_filepath = excel_filepath
                
                if current_plot_type == "NMR":
                    self.read_excel_file()
                elif current_plot_type == "Mass Spec":
                    self.read_csv_file()
                    
                    
                if self.excel_file_okay == True:
                    if current_plot_type == "NMR":
                        try:
                            self.compile_pivot_table()
                        except Exception as e:
                            message = "Error when computing data.\n" + repr(e)
                            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
                            msg_dlg.ShowModal()
                            msg_dlg.Destroy()
                            exit(1)
                    
                    elif current_plot_type == "Mass Spec":
                        try:
                            self.compile_MS_pivot_table()
                        except Exception as e:
                            message = "Error when computing data.\n" + repr(e)
                            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
                            msg_dlg.ShowModal()
                            msg_dlg.Destroy()
                            exit(1)
                        
                        
                    self.set_excel_file_label()
                
        dlg.Destroy()




    
    
    def set_excel_file_label(self):
        self.current_excel_file_label.SetLabel(self.excel_filepath)
    
 
    
    
    
    
    def read_excel_file(self):
        self.excel_file_okay = True
        excel_filename = os.path.split(self.excel_filepath)[1]
        excel_filename = re.split(r"\.xlsx|\.xlsm|\.xls", excel_filename)[0]
        
        directory_path = os.path.split(self.excel_filepath)[0]
        self.new_directory = os.path.join(directory_path, excel_filename)
        
        
        
        workbook = pandas.ExcelFile(self.excel_filepath)
        
        if not "#normalization" in workbook.sheet_names:
            message = "The selected Excel file does not contain a sheet named \"#normalization\". This is a required sheet."
            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            self.excel_file_okay = False
            
        else:
            self.normalization_df = pandas.read_excel(workbook, sheetname = "#normalization")
            required_columns_NDF = ["#sample", "#standard_concentration_mM", "#sample_fraction", "#protein_mass_mg"]
            
            if not set(required_columns_NDF).issubset(self.normalization_df.columns) :
                missing_columns = set(required_columns_NDF) - set(self.normalization_df.columns)
                message = "The #normalization sheet in the Excel file is missing columns for: \n\n" + " ".join(missing_columns)
                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                self.excel_file_okay = False

            
        if not "#assignment" in workbook.sheet_names:
            message = "The selected Excel file does not contain a sheet named \"#assignment\". This is a required sheet."
            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            self.excel_file_okay = False
            
        else:
            self.assignment_df = pandas.read_excel(workbook, sheetname = "#assignment")
            required_columns_ADF = ["Sample", "assignment", "area/protons/sf"]
            
            if not set(required_columns_ADF).issubset(self.assignment_df.columns) :
                missing_columns = set(required_columns_ADF) - set(self.assignment_df.columns)
                message = "The #assignment sheet in the Excel file is missing columns for: \n\n" + " ".join(missing_columns)
                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                self.excel_file_okay = False
                
            else:
                self.assignment_df = self.assignment_df[self.assignment_df["Sample"].notnull()]
                self.compounds = self.assignment_df[self.assignment_df["assignment"].notnull()].loc[:, "assignment"].unique()
                if not "DSS" in self.compounds:
                    message = "DSS is not assigned in the assignment column of the #assignment sheet in the Excel file."
                    msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                    self.excel_file_okay = False
                    
                for compound in self.compounds:
                    if re.match(r"~|#|%|&|\*|\{|\}|\\|:|<|>|\?|/|\+|\||\"", compound):
                            message = "The compound " + compound + " has an invalid character in the name."
                            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                            msg_dlg.ShowModal()
                            msg_dlg.Destroy()
                            self.excel_file_okay = False
            
        self.workbook = workbook
        self.excel_filename = excel_filename
        self.directory_path = directory_path
            
        

    
    
    def compile_pivot_table(self):
                
        self.normalization_df.index = self.normalization_df.loc[:, "#sample"]
        self.normalization_df = self.normalization_df[self.normalization_df["#standard_concentration_mM"].notnull()]
                                  
        sample_names = self.normalization_df[self.normalization_df["#sample"].notnull()].loc[:, "#sample"].unique()
        
        ## Create new dataframe to be like a pivot table.
        pivot_table_df = pandas.DataFrame(index = sample_names, columns =  self.compounds)
        
        ## Fill pivot table dataframe appropriately.
        for sample_name in pivot_table_df.index:
            for compound in pivot_table_df.columns:
                pivot_table_df.loc[sample_name, compound] = self.assignment_df[(self.assignment_df["Sample"] == sample_name) & (self.assignment_df["assignment"] == compound)]["area/protons/sf"].sum()
        
        
        self.original_pivot_table_df = pivot_table_df.copy()
        
        ## Normalize to DSS by dividing by area/protons/sf of DSS.
        pivot_table_df = pivot_table_df.divide(pivot_table_df["DSS"], axis=0)
        self.dss_normalized_df = pivot_table_df.copy()
        self.dss_normalized_df.insert(0, column = "Original DSS", value = self.original_pivot_table_df.loc[:, "DSS"])
        
        ## Divide by standard nmols per compound.
        pivot_table_df = pivot_table_df.multiply(self.normalization_df.loc[sample_names, "#standard_concentration_mM"], axis=0)
        self.standard_nmols_df = pivot_table_df.copy()
        self.standard_nmols_df.insert(0, column = "#standard_concentration_mM", value = self.normalization_df.loc[sample_names, "#standard_concentration_mM"])
                                                                           
        ## Divide by sample fraction.
        pivot_table_df = pivot_table_df.divide(self.normalization_df.loc[sample_names, "#sample_fraction"], axis=0)
        self.sample_fraction_df = pivot_table_df.copy()
        self.sample_fraction_df.insert(0, column = "#sample_fraction", value = self.normalization_df.loc[sample_names, "#sample_fraction"])
                                                                         
        ## Divide by protein mass mg.
        pivot_table_df = pivot_table_df.divide(self.normalization_df.loc[sample_names, "#protein_mass_mg"], axis=0)
        self.protein_mass_df = pivot_table_df.copy()
        self.protein_mass_df.insert(0, column = "#protein_mass_mg", value = self.normalization_df.loc[sample_names, "#protein_mass_mg"])

    
    
        self.sample_names = sample_names
        self.pivot_table_df = pivot_table_df
    
    
 
    
    
    
    def read_csv_file(self):
        
        self.excel_file_okay = True
        excel_filename = os.path.split(self.excel_filepath)[1]
        excel_filename = re.split(r"\.csv", excel_filename)[0]
        
        directory_path = os.path.split(self.excel_filepath)[0]
        self.new_directory = os.path.join(directory_path, excel_filename)
    
        MS_df = pandas.read_csv(self.excel_filepath)
        
        required_columns = ["Compound", "SamplID", "C_isomers", "Amount_ProteinAdj_uMol_g_protein_SequenceBased"]
        if not set(required_columns).issubset(MS_df.columns) :
            missing_columns = set(required_columns) - set(MS_df.columns)
            message = "The selected csv file is missing columns for: \n\n" + " ".join(missing_columns)
            msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            self.excel_file_okay = False
        
        self.MS_df = MS_df
        self.excel_filename = excel_filename
        self.directory_path = directory_path







    def compile_MS_pivot_table(self):    
        
        self.compounds = self.MS_df[self.MS_df["Compound"].notnull()].loc[:, "Compound"].unique()
        sample_names = self.MS_df[self.MS_df["SamplID"].notnull()].loc[:, "SamplID"].unique()
        
        ## Sort C_isomer and N_isomer so they are plotted correctly.
        if "N_isomers" in self.MS_df.columns:
            self.MS_df.sort_values(by=["SamplID", "Compound", "N_isomers", "C_isomers"], inplace=True)
        else:
            self.MS_df.sort_values(by=["SamplID", "Compound", "C_isomers"], inplace=True)
        
        ## Make C_isomers column a string and concatenate a C to the beginning.
        self.MS_df.loc[:, "C_isomers"] = "C" + self.MS_df.loc[:, "C_isomers"].astype(str)
        
        ## If N_isomers column exists, make N_isomers column a string and concatenate a N to the beginning.
        ## Also combine the C_isomers and N_isomers columns into one.
        if "N_isomers" in self.MS_df.columns:
            self.MS_df.loc[:, "N_isomers"] = "N" + self.MS_df.loc[:, "N_isomers"].astype(str)
            self.MS_df["Isomers_String"] = self.MS_df.loc[:, "C_isomers"] + "_" + self.MS_df.loc[:, "N_isomers"]
        else:
            self.MS_df["Isomers_String"] = self.MS_df.loc[:, "C_isomers"]
                
        important_columns_of_MS_df = self.MS_df.loc[:, ["SamplID", "Compound", "Isomers_String", "Amount_ProteinAdj_uMol_g_protein_SequenceBased"]]
        
        important_columns_of_MS_df.set_index(["SamplID", "Compound", "Isomers_String"], inplace=True)
        
        self.pivot_table_df = important_columns_of_MS_df.unstack(level=["Compound", "Isomers_String"])
        
        self.pivot_table_df.fillna(0, inplace=True)
        
        self.sample_names = sample_names
    
    
    
    
    
    def onListBox(self, event):
        """"""
        
        groups_listbox = event.GetEventObject()
        selected_group = groups_listbox.GetStringSelection()
        
        self.samples_listbox.Set(self.groups[selected_group])


 
    
    
    
    def Update_Group_List(self):
        """"""
        
        self.groups_listbox.Set(list(self.groups.keys()))
    
 
    
    
    
    
    def Add_Group(self, event):
        """"""
        
        if self.excel_file_okay == False:
            message = "Please select a valid file."
            msg_dlg = wx.MessageDialog(None, message, "No File Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return
        
        
        initial_groups_length = len(self.groups)
        group_name = ""
        while group_name == "":
            dlg = wx.TextEntryDialog(None, "Enter the group name:", "Group Name")
            if dlg.ShowModal() == wx.ID_OK:
                group_name = dlg.GetValue()
                if group_name == "":
                    message = "Please enter a valid group name."
                    msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
    
                dlg.Destroy()
            else:
                return
                
        
        selections = []
        while len(selections) == 0:
            dlg = wx.MultiChoiceDialog(None, "Select the samples in group " + group_name + ":", "Samples in Group \"" + group_name + "\"", self.sample_names)
        
        
            if dlg.ShowModal() == wx.ID_OK:
                selections = dlg.GetSelections()
                if len(selections) != 0:
                    self.groups[group_name] = [self.sample_names[x] for x in selections]
                    
                else:
                    message = "Please select at least one sample for the group."
                    msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                dlg.Destroy()
                
            else:
                return
                
                
        if len(self.groups) > initial_groups_length:
            self.Update_Group_List()
            
 




    def Delete_Group(self, event):
        """"""
        
        if self.excel_file_okay == False:
            message = "Please select a valid file."
            msg_dlg = wx.MessageDialog(None, message, "No File Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return
        elif len(self.groups) == 0:
            message = "There are no groups to delete."
            msg_dlg = wx.MessageDialog(None, message, "No Groups To Delete", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return


        selected_group = self.groups_listbox.GetStringSelection()
        if selected_group != wx.NOT_FOUND:
            self.groups.pop(selected_group)
            self.Update_Group_List()
            self.samples_listbox.Clear()
        else:
            message = "No group selected."
            msg_dlg = wx.MessageDialog(None, message, "No Group Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()






    def Move_Group_To_Top(self, event):
        """"""
        
        if self.excel_file_okay == False:
            message = "Please select a valid file."
            msg_dlg = wx.MessageDialog(None, message, "No File Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return
        elif len(self.groups) == 0:
            message = "There are no groups to move."
            msg_dlg = wx.MessageDialog(None, message, "No Groups To Move", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return


        selected_group = self.groups_listbox.GetStringSelection()
        if selected_group != wx.NOT_FOUND:
            self.groups.move_to_end(selected_group, last=False)
            self.Update_Group_List()
            self.samples_listbox.Clear()
        else:
            message = "No group selected."
            msg_dlg = wx.MessageDialog(None, message, "No Group Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()





    
    
    def Move_Group_To_Bottom(self, event):
        """"""
        
        if self.excel_file_okay == False:
            message = "Please select a valid file."
            msg_dlg = wx.MessageDialog(None, message, "No File Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return
        elif len(self.groups) == 0:
            message = "There are no groups to move."
            msg_dlg = wx.MessageDialog(None, message, "No Groups To Move", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return


        selected_group = self.groups_listbox.GetStringSelection()
        if selected_group != wx.NOT_FOUND:
            self.groups.move_to_end(selected_group, last=True)
            self.Update_Group_List()
            self.samples_listbox.Clear()
        else:
            message = "No group selected."
            msg_dlg = wx.MessageDialog(None, message, "No Group Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()




    
    
    def Add_Sample(self, event):
        """"""
        
        if self.excel_file_okay == False:
            message = "Please select a valid file."
            msg_dlg = wx.MessageDialog(None, message, "No File Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return
        
        
        selected_group = self.groups_listbox.GetStringSelection()
        if selected_group != "":
            samples_to_display = list(set(self.sample_names) - set(self.groups[selected_group]))
            dlg = wx.MultiChoiceDialog(None, "Select the samples to add to group " + selected_group + ":", "Add Samples To Group \"" + selected_group + "\"", samples_to_display)
        
            if dlg.ShowModal() == wx.ID_OK:
                selections = dlg.GetSelections()
                if len(selections) != 0:
                    self.groups[selected_group] = self.groups[selected_group] + [samples_to_display[x] for x in selections]
                    self.samples_listbox.Set(self.groups[selected_group])
        else:
            message = "No group selected."
            msg_dlg = wx.MessageDialog(None, message, "No Group Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
                



    def Delete_Sample(self, event):
        """"""
        
        if self.excel_file_okay == False:
            message = "Please select a valid file."
            msg_dlg = wx.MessageDialog(None, message, "No File Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return
        
        
        selected_group = self.groups_listbox.GetStringSelection()
        if selected_group != "":
            index_to_delete = self.samples_listbox.GetSelection()
            if index_to_delete != wx.NOT_FOUND:
                if len(self.groups[selected_group]) == 1:
                    message = "Groups must contain at least one sample."
                    msg_dlg = wx.MessageDialog(None, message, "Not Enough Samples", wx.OK | wx.ICON_EXCLAMATION)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                else:
                    del self.groups[selected_group][index_to_delete]
                    self.samples_listbox.Set(self.groups[selected_group])
            else:
                message = "No sample selected."
                msg_dlg = wx.MessageDialog(None, message, "No Sample Selected", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                
        else:
            message = "No group selected."
            msg_dlg = wx.MessageDialog(None, message, "No Group Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()






    def Add_Pivot_Table(self, event):
        """"""
        if self.current_plot_type == "NMR":
            if self.excel_file_okay == False:
                message = "Please select a valid Excel file."
                msg_dlg = wx.MessageDialog(None, message, "No Excel File Selected", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return
            elif len(self.groups) == 0:
                message = "Please create some groups to include their information in the sheet."
                msg_dlg = wx.MessageDialog(None, message, "No Groups", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return
    
    
            sheet_name = ""
            while sheet_name == "":
                dlg = wx.TextEntryDialog(None, "Enter the name of the sheet to add:", "Sheet Name")
                if dlg.ShowModal() == wx.ID_OK:
                    sheet_name = dlg.GetValue()
                    if sheet_name == "" or len(sheet_name) > 31 :
                        message = "Please enter a valid sheet name."
                        msg_dlg = wx.MessageDialog(None, message, "Invalid Sheet Name", wx.OK | wx.ICON_EXCLAMATION)
                        msg_dlg.ShowModal()
                        msg_dlg.Destroy()
                        
                    elif sheet_name in self.workbook.sheet_names:
                        message = "Sheet name is already in workbook. Overwrite the sheet?"
                        msg_dlg = wx.MessageDialog(None, message, "Sheet Name In Use", wx.YES_NO | wx.ICON_QUESTION)
                        if msg_dlg.ShowModal() == wx.ID_NO:
                            sheet_name = ""
                        msg_dlg.Destroy()
        
                    dlg.Destroy()
                else:
                    dlg.Destroy()
                    return
                
            book = openpyxl.load_workbook(self.excel_filepath)
            writer = pandas.ExcelWriter(self.excel_filepath, engine = "openpyxl")
            writer.book = book
            row = 0
            table_title = pandas.DataFrame(index = ["Table1: Sum of area/protons/sf"])
            table_title.to_excel(writer, columns = None, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + 1
            self.original_pivot_table_df.to_excel(writer, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + len(self.original_pivot_table_df) + 3
            
            table_title = pandas.DataFrame(index = ["Table2: dss normalized  = Table1 / (DSS column of Table1)"])
            table_title.to_excel(writer, columns = None, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + 1
            self.dss_normalized_df.to_excel(writer, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + len(self.dss_normalized_df) + 3
            
            table_title = pandas.DataFrame(index = ["Table3: Table2 * standard nmoles"])
            table_title.to_excel(writer, columns = None, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + 1
            self.standard_nmols_df.to_excel(writer, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + len(self.standard_nmols_df) + 3
            
            table_title = pandas.DataFrame(index = ["Table4: Table3 / sample fraction"])
            table_title.to_excel(writer, columns = None, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + 1
            self.sample_fraction_df.to_excel(writer, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + len(self.sample_fraction_df) + 3
            
            table_title = pandas.DataFrame(index = ["Table5: Table4 / protein mass mg"])
            table_title.to_excel(writer, columns = None, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + 1
            self.protein_mass_df.to_excel(writer, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + len(self.protein_mass_df) + 3
            
            group_names = self.groups.keys()
            group_average = pandas.DataFrame(index = group_names, columns = self.compounds)
            group_std = pandas.DataFrame(index = group_names, columns = self.compounds)
            for group_name, samples in self.groups.items():
                ## Compute average and standard deviation of groups for plotting.
                group_average.loc[group_name] = self.pivot_table_df.loc[samples].mean(axis=0)
                group_std.loc[group_name] = self.pivot_table_df.loc[samples].std(axis=0)
                
            
            
            table_title = pandas.DataFrame(index = ["Table6: Group averages of values from Table5"])
            table_title.to_excel(writer, columns = None, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + 1
            group_average.to_excel(writer, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + len(group_average) + 3
            
            table_title = pandas.DataFrame(index = ["Table7: Group standard deviations of values from Table5"])
            table_title.to_excel(writer, columns = None, sheet_name = sheet_name, startrow = row, startcol = 0)
            row = row + 1
            group_std.to_excel(writer, sheet_name = sheet_name, startrow = row, startcol = 0)
            
            
            writer.save()
            
            
            
        elif self.current_plot_type == "Mass Spec":
            if self.excel_file_okay == False:
                message = "Please select a valid csv file."
                msg_dlg = wx.MessageDialog(None, message, "No csv File Selected", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return
            elif len(self.groups) == 0:
                message = "Please create some groups to include their information in the file."
                msg_dlg = wx.MessageDialog(None, message, "No Groups", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return
            
            dlg = wx.FileDialog(
                self, message="Save file as ...", 
                defaultDir=self.directory_path, 
                defaultFile="", style=wx.FD_SAVE
                )
            if dlg.ShowModal() == wx.ID_OK:
                path = dlg.GetPath()
                dlg.Destroy()
            else:
                dlg.Destroy()
                return
            
            
            group_names = self.groups.keys()
            group_average = pandas.DataFrame(index = group_names, columns = self.pivot_table_df.columns)
            group_std = pandas.DataFrame(index = group_names, columns = self.pivot_table_df.columns)
            for group_name, samples in self.groups.items():
                ## Compute average and standard deviation of groups for plotting.
                group_average.loc[group_name] = self.pivot_table_df.loc[samples].mean(axis=0)
                group_std.loc[group_name] = self.pivot_table_df.loc[samples].std(axis=0)
            
            table_title = pandas.DataFrame(index = ["Pivoted Concentration Data"])
            table_title.to_csv(path, header=False)
            with open(path, "a") as open_file:
                self.pivot_table_df.to_csv(open_file)
                
                table_title = pandas.DataFrame(index = ["", "Average concentrations of groups"])
                table_title.to_csv(open_file, header=False)
                group_average.to_csv(open_file)
                
                table_title = pandas.DataFrame(index = ["", "STD of concentrations of groups"])
                table_title.to_csv(open_file, header=False)
                group_std.to_csv(open_file)
                

           
    
    def Create_Plots(self, event):
        """"""
        if self.excel_file_okay == False:
            message = "Please select a valid Excel file."
            msg_dlg = wx.MessageDialog(None, message, "No Excel File Selected", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return
        elif len(self.groups) == 0:
            message = "Please create some groups to plot."
            msg_dlg = wx.MessageDialog(None, message, "No Groups To Plot", wx.OK | wx.ICON_EXCLAMATION)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return
        
        if not os.path.exists(self.new_directory):
            try:
                os.makedirs(self.new_directory)
                
            except OSError as e:
                if e.errno != errno.EEXIST:
                    raise
    

            
        if self.current_plot_type == "NMR":
            
            group_names = self.groups.keys()
            group_average = pandas.DataFrame(index = group_names, columns = self.compounds)
            group_std = pandas.DataFrame(index = group_names, columns = self.compounds)
            for group_name, samples in self.groups.items():
                ## Compute average and standard deviation of groups for plotting.
                group_average.loc[group_name] = self.pivot_table_df.loc[samples].mean(axis=0)
                group_std.loc[group_name] = self.pivot_table_df.loc[samples].std(axis=0)

            x_pos = list(range(len(group_names)))
            for compound in self.pivot_table_df.columns:
                plt.bar(x_pos, group_average.loc[:, compound], yerr=[ [0]*len(group_std), group_std.loc[:, compound]/2], align="center", alpha=0.5, error_kw = {"ecolor":"black", "capsize":5, "capthick":2})
                plt.xticks(x_pos, group_names)
                plt.ylabel("nmoles/mg protein (+/-std)")
                plt.title(compound)
                plt.savefig(os.path.join(self.new_directory, self.excel_filename) + "_" + compound + ".png")
                plt.close()
                
                
                
        elif self.current_plot_type == "Mass Spec":

            message = "Creating Compound Plots, Please Wait."
            dlg = wx.BusyInfo(message)            
            bar_width = 0.35
            
            for compound in self.pivot_table_df.columns.levels[self.pivot_table_df.columns.names.index("Compound")]:
                
                ## Subset to only the compound of interest.
                temp_df = self.pivot_table_df.xs(compound, level = "Compound", axis=1)
                ## If there is no data for a compound then skip it.
                if temp_df.values.sum() == 0:
                    continue
                ## Get the index of the Isomer_String level.
                isomers_index = temp_df.columns.names.index("Isomers_String")
                ## Get the isomers of the compound.
                compound_isomers = temp_df.columns.levels[isomers_index][temp_df.columns.labels[isomers_index]]
                
                width_between_isomers = bar_width * (len(self.groups)+1)
                x_pos = [x * width_between_isomers for x in list(range(len(compound_isomers)))]
                
                plt.figure(figsize = (4 + len(compound_isomers)*2, 4))
                
                
                for group_name, samples in self.groups.items():
                    
                    group_average = temp_df.loc[samples].mean(axis=0)
                    group_std = temp_df.loc[samples].std(axis=0)/2
                
                    plt.bar(x_pos, group_average, width = bar_width, yerr=[ [0]*len(group_std), group_std], align="center", alpha=0.5, error_kw = {"ecolor":"black", "capsize":5, "capthick":2},)
                
                    x_pos = [x+bar_width for x in x_pos]
                
                
                x_pos = [x * width_between_isomers for x in list(range(len(compound_isomers)))]
                ## Very funky math to make the label appear in the middle of the bars.
                x_pos = [x + bar_width*(len(self.groups)/2) - bar_width/2*(len(self.groups)%2 ^ 1) for x in x_pos]
                plt.xticks(x_pos, compound_isomers)
                plt.ylabel("nmoles/mg protein (+/-std)")
                plt.title(compound)
                plt.legend(self.groups.keys(), loc = "upper right")
                plt.savefig(os.path.join(self.new_directory, self.excel_filename) + "_" + compound + ".png")
                plt.close()
                
            del dlg



def main():

    ex = wx.App(False)
    Plots_GUI(None, title = "Plot Compounds")
    ex.MainLoop()    


if __name__ == '__main__':
    main()
