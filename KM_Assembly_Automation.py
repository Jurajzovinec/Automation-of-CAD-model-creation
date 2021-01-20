import os
import tkinter
import sys
import creopyson
import xlrd
import logging
import time
import shutil
import csv
import threading
import re
from difflib import SequenceMatcher
from PIL import ImageTk, Image
from tkinter import messagebox
from tkinter import filedialog

# TODO: comment everything !!!
# TODO: Replace os module with shutil/subproccess !!!
# TODO: Critical buttons could be refactored.

# ---Logging---
try:
    os.remove("assembly_automation.log")
except:
    pass
logging.basicConfig(filename="assembly_automation.log", level=logging.INFO)
logger = logging.getLogger()


class Application(tkinter.Frame):
    """Main application frontend object. Object inherits from tkinter.Frame Class"""
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.geometry("485x600")
        self.master.resizable(0, 0)
        self.master.title('Kraussmaffei Assembly Automation')
        self.rows = 0

        while self.rows < 13:
            master.rowconfigure(self.rows, weight=1)
            master.columnconfigure(self.rows, weight=1)
            self.rows += 1

        # paint of the master
        set_global_paths()
        self.set_background_theme()
        self.create_entry()
        self.create_buttons()
        self.create_machine_type_list()

    def set_background_theme(self):
        """This function puts background theme in to the application."""
        self.background = ImageTk.PhotoImage(Image.open(icons_folder_path + "Graphical_User_Int_Theme.png"))
        self.background_theme = tkinter.Label(self.master, image=self.background).grid(row=0, column=0, rowspan=100, columnspan=100)

    def create_buttons(self):

        global run_button
        run_button = CreateControlButton(parent=self, row_grid=13, column_grid=0, command=automation_process, icon_name="run_automation.png")
        run_button.disable_this_button()
        global compare_zs_63_button
        compare_zs_63_button = CreateControlButton(parent=self, row_grid=13, column_grid=1, command=compare_with_zs63_file_button, icon_name="compare_zs_63.png")
        self.reset_button = CreateControlButton(parent=self, row_grid=13, column_grid=2, command=self.reset_graphical_user_interface, icon_name="reset.png")
        self.feedback_button = CreateControlButton(parent=self, row_grid=14, column_grid=0, command=self.open_feedback_folder, icon_name="feedback_folder.png")
        self.source_folder_button = CreateControlButton(parent=self, row_grid=14, column_grid=1, command=self.open_database_folder, icon_name="source_folder.png")
        self.quit_button = CreateControlButton(parent=self, row_grid=14, column_grid=2, command=self.close_graphical_user_interface, icon_name="quit.png")

    def create_entry(self):

        global order_number_entry
        order_number_entry = tkinter.Entry(self.master)
        order_number_entry.config(width=20, font=('Helvetica', 14), borderwidth=4)
        order_number_entry.insert(0, 'Number')
        order_number_entry.grid(row=11, column=1)

    def create_machine_type_list(self):
        """This method looks into database excel file and determines Machine types of Injection unit machines according to names of sheets."""
        list_gm_types = []

        input_workbook = xlrd.open_workbook(database_path)
        for sheet in input_workbook.sheets():
            list_gm_types.append(sheet.name)

        self.machine_type_drop_down_menu = CreateDropDownMenu(self, list_properties=list_gm_types, row_grid=2, column_grid=0)
        self.confirm_selected_machinetype = CreateControlButton(parent=self, row_grid=2, column_grid=2, command=self.create_cad_models_list, icon_name="blue_check_mark.png")
        self.confirm_selected_machinetype.set_rowspan_equals_2()

    def create_cad_models_list(self):
        """This method lists all mastermodels (CAD names), according to machine type picked by user."""
        # Initial GUI operations
        run_button.enable_this_button()
        self.machine_type_drop_down_menu.disable_this_dropdown_menu()
        self.confirm_selected_machinetype.disable_this_button()
        self.selected_machine_type = self.machine_type_drop_down_menu.what_is_picked_option()

        # Set up of Excel Workbook
        input_workbook = xlrd.open_workbook(database_path)
        self.input_worksheet = input_workbook.sheet_by_name(self.selected_machine_type)

        self.list_master_models = []
        self.positions_master_models = []

        for row_value in range(self.input_worksheet.nrows):
            if self.input_worksheet.cell_value(row_value, 0) != '' and self.input_worksheet.cell_value(row_value, 0) != 'CAD mastermodel name':
                self.list_master_models.append(self.input_worksheet.cell_value(row_value, 0))
                position_master_model = {'CAD_name': self.input_worksheet.cell_value(row_value, 0), 'rows_start': row_value}
                self.positions_master_models.append(position_master_model.copy())

        self.positions_master_models.reverse()

        # Creating range of properties (defined by range of rows in excel.)
        for each_dict in self.positions_master_models:
            try:
                each_dict['rows_finish'] = previous_row_start
            except:
                each_dict['rows_finish'] = each_dict['rows_start'] + 6
            finally:
                previous_row_start = each_dict['rows_start']

        # list is reversed - so we can define upcoming end of the current range.
        self.positions_master_models.reverse()
        self.list_cad_models = CreateDropDownMenu(self, list_properties=self.list_master_models, row_grid=3, column_grid=0)
        self.confirm_selected_cad = CreateControlButton(parent=self, row_grid=3, column_grid=2, command=self.create_master_model_properties, icon_name="blue_check_mark.png")
        self.confirm_selected_cad.set_rowspan_equals_2()

    def create_master_model_properties(self):
        """Following method determines selectable properties of CAD model according to sheet."""
        # Initial GUI operations
        self.list_cad_models.disable_this_dropdown_menu()
        self.confirm_selected_cad.disable_this_button()

        # For to me unknown reason application did not enabled button with run_button.enable_this_button (even though it works with confirmation button)
        # There for instance of this button is re-created.

        self.run_button = CreateControlButton(parent=self, row_grid=13, column_grid=0, command=automation_process, icon_name="run_automation.png")
        self.source_assembly_name = self.list_cad_models.what_is_picked_option()

        # Lists of newly created values - These properties have to be predifined.
        list_clamp_sizes = []
        list_powerpacks = []
        list_primary_plast = []
        list_secondary_plast = []
        global special_sign
        special_sign = None
        list_second_powerpacks = []
        global properties
        properties = [{'property': 'list_clamp_sizes', 'value': list_clamp_sizes},
                      {'property': 'list_powerpacks', 'value': list_powerpacks},
                      {'property': 'list_primary_plast', 'value': list_primary_plast},
                      {'property': 'list_secondary_plast', 'value': list_secondary_plast},
                      {'property': 'list_second_powerpacks', 'value': list_second_powerpacks}]

        selected_CAD = self.list_cad_models.what_is_picked_option()
        # Range is iterated and correct CAD is selected.
        range_and_selected_CAD = next(i for i in self.positions_master_models if i['CAD_name'] == selected_CAD)
        print(self.positions_master_models)
        # Definition of the range
        start_range = range_and_selected_CAD['rows_start']
        end_range = range_and_selected_CAD['rows_finish']
        working_range_CAD_master = range(start_range, end_range)

        # Now we have picked range ! so we can create lists of all CAD model properties
        for row_value in working_range_CAD_master:
            # Creating clamp_units:
            if self.input_worksheet.cell_value(row_value, 1):
                list_clamp_sizes.append(self.input_worksheet.cell_value(row_value, 1))
            # Creating powerpacks:
            if self.input_worksheet.cell_value(row_value, 2):
                list_powerpacks.append(self.input_worksheet.cell_value(row_value, 2))
            # Creating Primary_plast_options:
            if self.input_worksheet.cell_value(row_value, 3):
                list_primary_plast.append(self.input_worksheet.cell_value(row_value, 3))
            # Creating Secondary_plast options:
            if self.input_worksheet.cell_value(row_value, 4):
                list_secondary_plast.append(self.input_worksheet.cell_value(row_value, 4))
            # Setting special sign
            if self.input_worksheet.cell_value(row_value, 5):
                special_sign = self.input_worksheet.cell_value(row_value, 5)
                properties.append({'property': 'special_sign', 'value': special_sign})
            # Creating list of second powerpacks - suited for GXL machines
            if self.input_worksheet.cell_value(row_value, 6):
                list_second_powerpacks.append(self.input_worksheet.cell_value(row_value, 6))

        # Removing non relevant properties is going to happen with following lambda function
        filtering_properties = filter(lambda x: isinstance(x['value'], list) and len(x['value']) > 1 or isinstance(x['value'], str), properties)
        properties = list(filtering_properties)

        self.create_dropdown_menu_for_properties(self, properties)

    def create_dropdown_menu_for_properties(self, parent, list_of_properties):
        """Following method determines properties of CAD model according to sheet."""
        row_grid = 3
        column_grid_label = 2
        column_grid_dropdown_menu = 0
        self.list_of_option_properties = []
        self.parent = parent
        for property in list_of_properties:
            if isinstance(property['value'], list):
                row_grid += 1

                # Creating label - class Create Label is used.
                label_name = property['property']
                text = label_name.replace('_', ' ')
                text = text.replace('list', '').strip()
                text = text.capitalize()
                self.label_name = CreateLabel(parent=self, row_grid=row_grid, column_grid=column_grid_label, text=text)

                # Creating dropdown menu

                drop_down_menu = CreateDropDownMenu(parent=self, row_grid=row_grid, column_grid=column_grid_dropdown_menu, list_properties=property['value'])
                self.list_of_option_properties.append({'drop_down_obj': drop_down_menu, 'property': property['property']})

    # BUTTON FUNCTIONS
    def reset_graphical_user_interface(self):
        """Button reaction function"""
        try:
            if main_thread.is_alive:
                tkinter.messagebox.showinfo(title='Automation in process', message='While the automation is running, it is not possible to restart app. Use quit button to terminate the application.')
            else:
                self.master.destroy()
                main()
        except:
            self.master.destroy()
            main()

    def close_graphical_user_interface(self):
        """Button reaction function"""
        yes_no_quit = tkinter.messagebox.askquestion('Termination of application', 'KM automation assembly application will be terminated.', icon='warning')
        if yes_no_quit:
            exit()

    def open_database_folder(self):
        """Button reaction function"""
        database_folder_path = '.\DatabaseFolder'
        os.startfile(database_folder_path)

    def open_feedback_folder(self):
        """Button reaction function"""
        os.startfile(feedback_folder_path)

    def open_log_file(self):
        """Button reaction function"""
        log_file = 'assembly_automation.log'
        os.startfile(log_file)


class CreateControlButton:

    def __init__(self, parent, row_grid, column_grid, command, icon_name):
        self.parent = parent
        self.command = command
        self.icon_name = icon_name
        self.row_grid = row_grid
        self.column_grid = column_grid
        self.icon_obj = ImageTk.PhotoImage(Image.open(icons_folder_path + self.icon_name))
        self.parent.button_obj = tkinter.Button(image=self.icon_obj, command=self.command)
        self.parent.button_obj.grid(row=self.row_grid, column=self.column_grid)
        self.image_name = self.parent.button_obj.image_names()

    def set_rowspan_equals_2(self):
        self.parent.button_obj.grid(rowspan=2)

    def disable_this_button(self):
        self.parent.button_obj['state'] = 'disable'

    def enable_this_button(self):
        self.parent.button_obj['state'] = 'normal'

    def destroy_this_button(self):
        self.parent.button_obj.destroy()


class CreateDropDownMenu:
    """This class represents list box shown in graphical user interface."""
    def __init__(self, parent, list_properties, row_grid, column_grid):
        self.parent = parent
        self.preselected_option = tkinter.StringVar(parent.master)
        self.list_properties = list_properties
        self.preselected_option.set(list_properties[0])
        self.parent.drop_down_menu = tkinter.OptionMenu(parent.master, self.preselected_option, *list_properties)
        self.parent.drop_down_menu.config(height=1, width=35, font=('Helvetica 9 bold'))
        self.parent.drop_down_menu.grid(row=row_grid, column=column_grid, rowspan=2, columnspan=2)

    def what_is_picked_option(self):
        return self.preselected_option.get()

    def disable_this_dropdown_menu(self):
        self.parent.drop_down_menu.config(state='disabled')

    def enable_this_dropdown_menu(self):
        self.parent.drop_down_menu.configure(state='normal')

    def return_non_picked_values(self):
        filtering_object = filter(lambda x: x != self.what_is_picked_option(), self.list_properties)
        non_picked_values_list = list(filtering_object)
        return non_picked_values_list


class CreateLabel:
    def __init__(self, parent, row_grid, column_grid, text):
        self.parent = parent
        self.parent.button_obj = tkinter.Label(text=text, font=('Helvetica 10 bold'), borderwidth=3, relief="ridge")
        self.parent.button_obj.grid(row=row_grid, column=column_grid, rowspan=2, columnspan=2)


class CreoAPI:
    """Main backend object. Methods of this class are based on creopyson library. These methods are however optimized for Kraussmaffei master model's data quality."""
    # Model names stick to lowercase convention - incorrect CX_0420_ZE77-0XX_GM1811.asm  - correct cx_420_ze77-0xx_gm1811.asm
    # List of functions
    # connect, create bom information (check whether resumed models are innit), check skeleton exists,
    # take care of flow control
    # possible names - use quicker functions - may slow down
    def __init__(self, open_master=True):
        self.creo_client = creopyson.Client()
        self.setup()
        self.configs_manipulation(api_mode=True)
        self.bill_of_material = []
        self.paired_bill_of_material = []
        self.zs_63 = Zs63()
        if open_master:
            self.open_picked_master_model()
        self.default_master_model = self.current_master_model()
        self.try_to_resume_all()

        print("Creo API has been initialized.")

    def setup(self):
        """Initialization of creo_client object. If creoson application is not running, program ask user to run it manually."""
        try:

            self.creo_client.connect()
            logger.info('Creoson is running')

        except ConnectionError:
            pass
            #creoson_folder = '.\creoson'
            #creoson_bat = r"creoson_run.bat"

            #os.chdir(creoson_folder)
            #os.startfile(creoson_bat)
            #self.creo_client.connect()

    def open_picked_master_model(self):
        creopyson.file_open(self.creo_client, file_=app.list_cad_models.what_is_picked_option())

    def current_master_model(self):
        return creopyson.file_get_fileinfo(self.creo_client)['file']

    def current_order_number(self):
        return self.current_master_model()[-10:-4]

    def create_master_model_bill_of_material_with_suppressed(self, levels):
        """This method uses list-features method to create structured Bill of Material"""
        self.clear_bill_of_material()
        current_master_model = self.current_master_model()
        current_order_number = self.current_order_number()

        # At first we list top level groups
        self.add_models_in_opened_group_to_bom(level_of_master_model_tree=1)
        first_level_groups = tuple(self.bill_of_material)

        if levels > 1:
            for each_dict in first_level_groups:
                try:
                    creopyson.file_open(self.creo_client, file_=each_dict['name'])
                except:
                    pass
                else:
                    self.add_models_in_opened_group_to_bom(level_of_master_model_tree=2)
                    list_of_models = creopyson.feature_list(self.creo_client, type_='COMPONENT')
                    if all(current_order_number.lower() in component['name'].lower() for component in list_of_models):
                        self.try_to_resume_all()
        second_level_groups = tuple(self.bill_of_material)

        if levels > 2:
            # And third Level groups
            # resuming second level groups (those are groups with order number)
            for each_dict in second_level_groups:
                if each_dict['level_of_master_model_tree'] == 2 and '.prt' not in each_dict['name'] and current_order_number in each_dict['name']:
                    try:
                        creopyson.file_open(self.creo_client, file_=each_dict['name'])
                    except:
                        pass
                    else:
                        self.add_models_in_opened_group_to_bom(level_of_master_model_tree=3)
            self.check_whether_is_destination_group()

            file_object = open("BOM.txt", "a+")

            for each in self.bill_of_material:
                if "m" in each["parent"][0:-10].lower() and re.findall("[0-9]{7}", each['name']):
                    file_object.write(f"GM  {each['name']}  {each['parent']}.\n")
                if "ze" in each["parent"][0:-4].lower() and re.findall("[0-9]{7}", each['name']):
                    file_object.write(f"ZE# {each['name']}  {each['parent']}\n")
                if "sa" in each["parent"][0:-4].lower() and re.findall("[0-9]{7}", each['name']):
                    file_object.write(f"SA  {each['name']}  {each['parent']}\n")

        creopyson.file_open(self.creo_client, file_=current_master_model)
        self.determine_assembly_group_type()
        self.check_whether_is_destination_group()

    def add_models_in_opened_group_to_bom(self, level_of_master_model_tree=None):

        bom_opened_group = creopyson.feature_list(self.creo_client, no_datum=True, type_='COMPONENT')
        parent = creopyson.file_get_fileinfo(self.creo_client)['file']
        for each_dict in bom_opened_group:
            each_dict.pop('type', None)
            each_dict['parent'] = parent
            each_dict['name'] = each_dict['name'].lower()
            each_dict['level_of_master_model_tree'] = level_of_master_model_tree
        self.bill_of_material.extend(bom_opened_group)
        # remove duplicates ->
        self.bill_of_material = [i for n, i in enumerate(self.bill_of_material) if i not in self.bill_of_material[n + 1:]]
        return bom_opened_group

    def create_master_model_bill_of_material_no_suppressed(self):
        """This method uses bom_get_paths method to create structured Bill of Material"""
        start_time = time.time()
        bom = []
        bom_cleansed = []
        bom_raw = self.bom_recursion(creopyson.bom_get_paths(self.creo_client))
        bom_raw.pop()
        bom_raw.reverse()

        for x in range(len(bom_raw)):
            if x % 2 == 0:
                pair_root_file_dict = {'root': bom_raw[x].replace('seq_path : ', ''), 'model': bom_raw[x + 1].replace('file : ', '')}
                bom_cleansed.append(pair_root_file_dict)

        for each_pair in bom_cleansed:
            list_split_root = each_pair['root'].split('.')
            list_split_root.pop()
            joined_root = '.'.join(list_split_root)
            for each in bom_cleansed:
                if each['root'] == joined_root:
                    pair_parent_child_dict = {'child': each_pair['model'], 'parent': each['model'], 'status': 'ACTIVE'}
                    bom.append(pair_parent_child_dict)
                    break

        message = "Creation of BOM took program " + ("--- %s seconds ---" % (time.time() - start_time))
        logger.info(message)
        return bom

    def change_parameter_in_bill_of_material(self, key, new_value, feat_id):
        """This method changes parameter in bill of material - it is necessary to track all changes."""
        try:
            change_dict = next(item for item in self.bill_of_material if item['feat_id'] == feat_id)
        except StopIteration:
            logger.warning(f'{feat_id} is not in master model.')
        else:
            index_of_change_dict = self.bill_of_material.index(change_dict)
            new_dict = {key: new_value}
            if isinstance(change_dict, dict):
                self.bill_of_material[index_of_change_dict].update(new_dict)
                print(f"BOM - parameter {key} has been changed to {new_value}. Modified item is {change_dict['name']}.")
                logger.info(f"BOM - parameter {key} has been changed to {new_value}. Modified item is {change_dict['name']}.")
            else:
                logger.warning('Changing element of BOM list is not dictionary.')

    def remove_dict_from_bill_of_material(self, key, value):
        """This method removes item (most likely assembly) from Bill of material"""
        try:
            remove_dict = next(item for item in self.bill_of_material if item[key] == value)
        except StopIteration:
            pass
        else:
            self.bill_of_material.remove(remove_dict)

    def check_whether_is_destination_group(self):
        """method check whether created group is group where material number will be placed."""
        order_number = self.current_order_number()
        for each_dict in self.bill_of_material:
            if order_number in each_dict['name'] and '.prt' not in each_dict['name'].lower():
                filter_children_only = filter(lambda x: x['parent'] == each_dict['name'], self.bill_of_material)
                filter_children_only_list = list(filter_children_only)
                if any(order_number in every_dict['name'] and '.asm' in every_dict['name'] for every_dict in filter_children_only_list):
                    each_dict['destination_group'] = 'no'
                else:
                    each_dict['destination_group'] = 'yes'
            else:
                each_dict['destination_group'] = 'no'

    def clear_bill_of_material(self):
        self.bill_of_material = []

    def try_delete_model(self, model_to_delete):
        try:
            creopyson.feature_delete(self.creo_client, clip=True, name=model_to_delete)
        except RuntimeError:
            pass
        else:
            self.remove_dict_from_bill_of_material(key='name', value=model_to_delete)
            print('This group has been removed ' + model_to_delete)
            logger.info('This group has been removed ' + model_to_delete)

    def filter_assemblies(self):
        """ This method removes top master assemblies groups. Here we will remove level 2 assemblies
            by simple comparing their strings (or names) to picked dimensions (by user). Every top assembly
            goes through test, which determines whether dimension is useful - if assembly is useless is joined to list
            dimension_to_remove. Every property is specially tested. """
        # Here we start by creating of list of the top assemblies
        self.create_master_model_bill_of_material_with_suppressed(levels=1)

        if any(dictionary['property'] == 'special_sign' for dictionary in properties):
            special_sign = next(item for item in properties if item['property'] == 'special_sign')['value'].lower()
            print(f'Special sign is {special_sign}.')

        if any(dictionary['property'] == 'list_clamp_sizes' for dictionary in properties):
            list_clamp_sizes = next(item for item in properties if item['property'] == 'list_clamp_sizes')['value']
            picked_dictionary = next(item for item in app.list_of_option_properties if item['property'] == 'list_clamp_sizes')['drop_down_obj']
            picked_option = picked_dictionary.what_is_picked_option()
            print(f'Picked clamp size is {picked_option}.')
            dimensions_to_remove = picked_dictionary.return_non_picked_values()
            print(dimensions_to_remove)
            # remove groups containing any of dimensions to remove
            # necessary to create copy of bill of material, because in this process models will be deleted out of the model
            top_level_list = tuple(self.bill_of_material)
            for each_group in top_level_list:
                if any(x in each_group['name'] for x in dimensions_to_remove):
                    self.try_delete_model(each_group['name'])

        if any(dictionary['property'] == 'list_powerpacks' for dictionary in properties):
            list_powerpacks = next(item for item in properties if item['property'] == 'list_powerpacks')['value']
            picked_dictionary = next(item for item in app.list_of_option_properties if item['property'] == 'list_powerpacks')['drop_down_obj']
            picked_option = picked_dictionary.what_is_picked_option()
            print(f'Picked powerpack is {picked_option}.')
            dimensions_to_remove = picked_dictionary.return_non_picked_values()
            self.determine_whether_group_consist_powerpack(list_of_powerpacks=list_powerpacks, picked_powerpack=picked_option)

            # Filtering out powerpacks is very difficult task since their names are 10, 15, 01 or 02. However there is rule that sign of powerpack is earlier
            # in model name than M, ZE, or C
            # Therefore following strategy of powerpack creation will be applied:
            # Create list of top level assemblies. For each level assembly will be determined whether it is powerpack group and whether it has pair in Master model.
            # Also determine whether it is ZE, C, or M group
            # Than model convention will be applied  - *_10_* will be renamed to *-10_*

        if any(dictionary['property'] == 'list_second_powerpacks' for dictionary in properties):
            list_second_powerpacks = next(item for item in properties if item['property'] == 'list_second_powerpacks')['value']
            picked_dictionary = next(item for item in app.list_of_option_properties if item['property'] == 'list_second_powerpacks')['drop_down_obj']
            picked_option = picked_dictionary.what_is_picked_option()
            print(f'Picked secondary powerpack is {picked_option}.')
            dimensions_to_remove = picked_dictionary.return_non_picked_values()
            self.determine_whether_group_consist_powerpack(list_of_powerpacks=list_second_powerpacks, picked_powerpack=picked_option)

        if any(dictionary['property'] == 'list_primary_plast' for dictionary in properties):
            list_primary_plast = next(item for item in properties if item['property'] == 'list_primary_plast')['value']
            picked_dictionary = next(item for item in app.list_of_option_properties if item['property'] == 'list_primary_plast')['drop_down_obj']
            picked_option = picked_dictionary.what_is_picked_option()
            print(f'Picked primary plast is {picked_option}.')
            dimensions_to_remove = picked_dictionary.return_non_picked_values()
            # remove groups containing any of dimensions to remove
            # important to store value into tuple - self.material is loosing its elements in this process - which causes overlaps while it iterating
            # tuple is immutable - therefore will not change (avoiding overlaps)
            top_level_list = tuple(self.bill_of_material)
            try:
                for each_group in top_level_list:
                    if any(x in each_group['name'] and special_sign not in each_group['name'] for x in dimensions_to_remove):
                        self.try_delete_model(each_group['name'])
            except:
                for each_group in top_level_list:
                    if any(x in each_group['name'] for x in dimensions_to_remove):
                        self.try_delete_model(each_group['name'])

        if any(dictionary['property'] == 'list_secondary_plast' for dictionary in properties):
            list_secondary_plast = next(item for item in properties if item['property'] == 'list_secondary_plast')['value']
            picked_dictionary = next(item for item in app.list_of_option_properties if item['property'] == 'list_secondary_plast')['drop_down_obj']
            picked_option = picked_dictionary.what_is_picked_option()
            print(f'Picked secondary plastification is {picked_option}.')
            dimensions_to_remove = picked_dictionary.return_non_picked_values()
            top_level_list = tuple(self.bill_of_material)
            for each_group in top_level_list:
                if any(x in each_group['name'] and special_sign in each_group['name'] for x in dimensions_to_remove):
                    self.try_delete_model(each_group['name'])

        self.clear_bill_of_material()

    def determine_assembly_group_type(self):
        """This method adds new model property to dictionary. Group types are destination or non-destination. Destination groups are designed for assembling of material numbers."""
        for bom_dict in self.bill_of_material:
            if 'group_type' not in bom_dict:
                bom_dict['group_type'] = 'not defined'
                if 'm' in bom_dict['name'][:-10] and '.asm' in bom_dict['name']:
                    bom_dict['group_type'] = 'm'
                elif 'ze' in bom_dict['name'][:-10] and '.asm' in bom_dict['name']:
                    bom_dict['group_type'] = 'ze'
                elif 'c' in bom_dict['name'][:-10] and '.asm' in bom_dict['name']:
                    try:
                        number_position = (bom_dict['name'].index('c') + 1)
                        int(bom_dict['name'][number_position])
                    except ValueError:
                        pass
                    else:
                        bom_dict['group_type'] = 'c'
                elif 'sa' in bom_dict['name'][:-10] and '.asm' in bom_dict['name']:
                    bom_dict['group_type'] = 'sa'

    def let_proe_rename_pdm_objects(self, boolean=True):
        if boolean:
            creopyson.creo_set_config(client=self.creo_client, name="let_proe_rename_pdm_objects", value="yes")
        else:
            creopyson.creo_set_config(client=self.creo_client, name="let_proe_rename_pdm_objects", value="no")

    def regenerate_read_only_config_control(self, boolean=True):
        if boolean:
            creopyson.creo_set_config(client=self.creo_client, name="regenerate_read_only_objects", value="yes")
        else:
            creopyson.creo_set_config(client=self.creo_client, name="regenerate_read_only_objects", value="no")

    def display_comps_to_assemble(self, boolean=True):
        if boolean:
            creopyson.creo_set_config(client=self.creo_client, name="display_comps_to_assemble", value="yes")
        else:
            creopyson.creo_set_config(client=self.creo_client, name="display_comps_to_assemble", value="no")

    def determine_whether_group_consist_powerpack(self, list_of_powerpacks, picked_powerpack):
        """This method determines whether group has powerpack and if so also determines its value and renames it to correct model convention."""
        # At first we check whether correct parameters are already in bill of material (we demand group_type from determine_assembly_group_type)
        if 'group_type' not in self.bill_of_material[1]:
            self.determine_assembly_group_type()
        powerpacks_to_remove = []
        powerpacks_to_keep = []
        groups_to_remove = []
        powerpacks_master_groups = []

        # names of powerpacks maybe like -10_ or -10-
        for each_powerpack in list_of_powerpacks:
            if picked_powerpack == each_powerpack:
                powerpacks_to_keep.append(f'_{each_powerpack}_')
                powerpacks_to_keep.append(f'-{each_powerpack}_')
            else:
                powerpacks_to_remove.append(f'_{each_powerpack}_')
                powerpacks_to_remove.append(f'-{each_powerpack}_')

        # correct powerpack groups will be determined here:
        for each_dict in self.bill_of_material:
            if each_dict['group_type'] != 'not defined':
                group_type_pos = each_dict['name'].index(each_dict['group_type'])
                if any(each_powerpack in each_dict['name'][0:group_type_pos] for each_powerpack in powerpacks_to_keep):
                    found_powerpack = next(item for item in powerpacks_to_keep if item in each_dict['name'][0:group_type_pos])
                    print(f'Found powerpack which will stay in master {found_powerpack} in {each_dict}.')

        # invalid powerpack groups will be classified here:
        for each_dict in self.bill_of_material:
            if each_dict['group_type'] != 'not defined':
                group_type_pos = each_dict['name'].index(each_dict['group_type'])
                if any(each_powerpack in each_dict['name'][0:group_type_pos] for each_powerpack in powerpacks_to_remove):
                    found_powerpack = next(item for item in powerpacks_to_remove if item in each_dict['name'][0:group_type_pos])
                    print(f'Found powerpack which will be removed from master {found_powerpack} in {each_dict}.')
                    each_dict['powerpack'] = found_powerpack
                    groups_to_remove.append(each_dict)

        # Faith of powerpack group will be chosen here - powerpack groups can go 3 ways
        # 1.) Removed because correct group exists in Wch and exists in master model
        # 2.) Removed because correct group exists in Wch. Correct group is not in master model therefore has to be assembled into it.
        # 3.) Kept and renamed

        for each_dict in groups_to_remove:
            group_type_pos = each_dict['name'].index(each_dict['group_type'])
            found_powerpack = next(item for item in powerpacks_to_remove if item in each_dict['name'][0:group_type_pos])
            each_dict['test_result'] = 'rename_to_correct'
            for each_correct in powerpacks_to_keep:
                try:
                    try_model = each_dict['name'].replace(found_powerpack, each_correct)
                except KeyError:
                    pass
                if self.check_whether_model_exists(erp_material_number=try_model):
                    print(f"Such a model exists {try_model}.")
                    try:
                        next(item for item in self.bill_of_material if item['name'] == try_model)
                    except StopIteration:
                        change = {'test_result': 'remove-assemble correct', 'correct': try_model}
                        each_dict.update(change)
                    else:
                        change = {'test_result': 'remove'}
                        each_dict.update(change)
                    finally:
                        break
            powerpacks_master_groups.append(each_dict)

        print('================Powerpack Handling============================')

        # Performance part
        for each_dict in groups_to_remove:
            if each_dict['test_result'] == 'remove':
                self.try_delete_model(model_to_delete=each_dict['name'])
            if each_dict['test_result'] == 'rename_to_correct':
                pass
                found_powerpack = each_dict['powerpack']
                self.open_model_and_rename_groups(component_name=each_dict['name'], replace_from=found_powerpack, replace_to=powerpacks_to_keep[1])
                print(f"Due to non-existing group {each_dict['name']} has been renamed.")
            if each_dict['test_result'] == 'remove-assemble correct':
                # TODO: Assemble model
                self.try_delete_model(model_to_delete=each_dict['name'])
                self.assemble_model(erp_material_number=each_dict['correct'], cad_parent_model=self.current_master_model())
                print(f"Missing group {each_dict['correct']} has been assembled to master model.")

        # now we determine whether desired powerpack at group exists in master model.

    def open_model_and_rename_groups(self, component_name, replace_from, replace_to):
        """Model is opened and closed by itself. Creo session returns to previous model."""
        get_session_info = creopyson.file_get_fileinfo(self.creo_client)
        current_model = (get_session_info['file'])
        creopyson.file_open(self.creo_client, file_=component_name)

        self.try_to_resume_all()
        self.add_models_in_opened_group_to_bom()

        filter_children_only = filter(lambda x: x['parent'] == component_name, self.bill_of_material)
        filter_children_only_list = list(filter_children_only)

        parent_dict = next(x for x in self.bill_of_material if x['name'] == component_name)
        filter_children_only_list.append(parent_dict)

        for every_component in filter_children_only_list:
            if replace_from in every_component['name']:
                new_name = every_component['name'].replace(replace_from, replace_to)
                try:
                    creopyson.file_rename(self.creo_client, file_=every_component['name'], new_name=new_name, onlysession=True)
                    print(f'this is new model = {new_name}.')
                    logger.info(f'this is new model = {new_name}.')
                except RuntimeError:
                    print(f"RuntimeError raised while renaming from {every_component['name']} to {new_name}.")
                    logger.warning(f"RuntimeError raised while renaming from {every_component['name']} to {new_name}.")
                except NameError:
                    print(f"NameError raised while renaming from {every_component['name']} to {new_name}.")
                    logger.warning(f"NameError raised while renaming from {every_component['name']} to {new_name}.")
                else:
                    self.change_parameter_in_bill_of_material(key='name', new_value=new_name, feat_id=every_component['feat_id'])

        self.determine_assembly_group_type()
        creopyson.file_close_window(self.creo_client)
        creopyson.file_open(self.creo_client, file_=current_model)

    def assemble_model(self, erp_material_number, cad_parent_model):
        """This method assembles material number into injection_machine group"""
        try:

            child_model = self.check_whether_model_exists(erp_material_number)

            if child_model:
                # Now we test whether model already exists in cad_parent_model. If yes we skip this step to avoid cad model duplicity:
                if not self.check_whether_model_name_is_in_assembly(parent=cad_parent_model, model_name_wild_card=erp_material_number):
                    child_csy = self.pick_csy(cad_parent_model, child_model)
                    parent_csy = self.pick_csy(cad_parent_model, cad_parent_model)

                    if parent_csy != 'CSY does not exist' and child_csy != 'CSY does not exist':
                        creopyson.file_assemble(self.creo_client, into_asm=cad_parent_model, file_=child_model, constraints=[{"asmref": parent_csy, "compref": child_csy, "type": "csys"}])
                        logger.info(f"Model {child_model} has been assembled to {cad_parent_model}")
                        print(f"Model {child_model} has been assembled to {cad_parent_model}")

                    elif parent_csy == 'CSY does not exist' and child_csy != 'CSY does not exist':
                        # TODO: Check whether skeleton exists method might be enhanced.
                        skeleton_information = self.check_whether_skeleton_exists(cad_parent_model)
                        if skeleton_information['csy'] != 'CSY does not exist':
                            creopyson.file_assemble(self.creo_client, file_=child_model, into_asm=cad_parent_model, ref_model=skeleton_information['skel_name'],
                                                    constraints=[{"asmref": skeleton_information['csy'], "compref": child_csy, "type": "csys"}])
                            logger.info(f"Model {child_model} has been assembled to {cad_parent_model}")
                            print(f"Model {child_model} has been assembled to {cad_parent_model}")
                        else:
                            creopyson.file_assemble(self.creo_client, into_asm=cad_parent_model, file_=child_model, constraints=[{"type": "fix"}], package_assembly=True)
                            logger.info(f"Model {child_model} has been assembled to {cad_parent_model}")
                            print(f"Model {child_model} has been assembled to {cad_parent_model}")
                    else:
                        creopyson.file_assemble(self.creo_client, into_asm=cad_parent_model, file_=child_model, constraints=[{"type": "fix"}], package_assembly=True)
                        logger.info(f"Model {child_model} has been assembled to {cad_parent_model}")
                        print(f"Model {child_model} has been assembled to {cad_parent_model}")
                else:
                    print(f"Model {erp_material_number} is already in {cad_parent_model}")

        except:
            print(f"Model {child_model} has not been assembled to {cad_parent_model}")
            sys.exc_info()
            logger.exception("message")
            print(sys.exc_info())

    def pick_csy(self, cad_parent_model, checked_model, mx_optimization=False):
        """This method determines coordinate systems for assembling purposes."""
        max_ratio = 0
        list_csys_model_raw = creopyson.feature_list(self.creo_client, file_=checked_model, type_='COORDINATE SYSTEM', no_comp=False)
        list_csys = []

        for every_dict in list_csys_model_raw:
            list_csys.append(every_dict['name'])
        # Test amount of csys in model
        # At first we test whether there is more than one csys. If so Sequence matcher function is used to determine correct CSY.
        if len(list_csys) > 1:
            for every_csy in list_csys:
                measure_similarity = SequenceMatcher(None, cad_parent_model, every_csy)
                every_csy_ratio = measure_similarity.ratio()
                if every_csy_ratio > max_ratio:
                    max_ratio = every_csy_ratio
                    best_match_csy = every_csy
            # There also has to be added option when ratio equals zero for all compared csys
            if max_ratio != 0:
                picked_csy = best_match_csy
            else:
                picked_csy = list_csys[0]
        # Then we test if number of csys is only one. If so this csy is assigned to picked_csy
        if len(list_csys) == 1:
            picked_csy = list_csys[0]
        # At last there is possibility of when there is none csys in model. In this case following string is assigned to picked_Csy
        if len(list_csys) == 0:
            picked_csy = 'CSY does not exist'
        if mx_optimization:
            # TODO: check whether this scales if yes mx_optimization boolean rule change to if mx in CAD name
            try:
                # 23-06-2020 this section will add possible coordinate system for MX machine
                # Requires more development
                if len(list_csys) > 1 and zs_63_injection_unit != "":
                    range_csys = []
                    # These coordinates system may have their own range of numbers. The first condition of these ,,range csys,, is that they contain ,,SP,,
                    for each_csy in list_csys:
                        if "SP" in each_csy.upper():
                            range_csys.append(each_csy)
                    # At this part we can determine whether range_csys has multiple elements. If yes, we create objects of csys where we determine several properties
                    mx_csys_objects = []
                    if len(range_csys) > 1:
                        for each_csy in range_csys:
                            sp_pos = each_csy.index("SP")
                            range_in_csy_string = each_csy[sp_pos + 2:]
                            range_in_csy_list = range_in_csy_string.split("-")
                            boundering_values = []
                            for each_range_value in range_in_csy_list:
                                boundering_values.append(only_numerics(each_range_value))
                        current_mx_csy_object = {
                            "name": each_csy,
                            "floor": min(boundering_values),
                            "ceil": max(boundering_values)
                        }
                        mx_csys_objects.append(current_mx_csy_object)
                        print(f"MX Optimization - csy object is {current_mx_csy_object}.")
                    if len(mx_csys_objects) > 1:
                        for each_csy_object in mx_csys_objects:
                            if zs_63_injection_unit >= each_csy_object["floor"] and zs_63_injection_unit <= each_csy_object["ceil"]:
                                picked_csy = each_csy_object["name"]
                                logger.info("MX variation of picked CSY is " + picked_csy)
            except:
                logger.exception("message")
            return picked_csy

        try:
            csy_obj = next(x for x in list_csys_model_raw if x['name'] == picked_csy)
        except:
            pass
        else:
            if csy_obj['status'] == 'SUPPRESSED':
                creopyson.feature_resume(self.creo_client, name=csy_obj['name'])

        return picked_csy

    def check_whether_model_exists(self, erp_material_number):
        """This method tests whether ERP material number exists in Windchill and if exists it will assign its modelname to model_name variable"""
        erp_material_number = erp_material_number.replace('.prt', '')
        erp_material_number = erp_material_number.replace('.asm', '')
        model_name = ''
        try:
            creopyson.file_open(self.creo_client, file_=erp_material_number + '.prt', display=False)
            model_name = erp_material_number + '.prt'
            print(f'Yes material number exists ! Model name is {model_name}.')
            logger.info(f'Yes material number exists ! Model name is {model_name}.')
        except RuntimeError:
            pass
        try:
            creopyson.file_open(self.creo_client, file_=erp_material_number + '.asm', display=False)
            model_name = erp_material_number + '.asm'
            print(f'Yes material number exists ! Model name is {model_name}.')
            logger.info(f'Yes material number exists ! Model name is {model_name}.')
        except RuntimeError:
            pass
        return model_name.lower()

    def check_whether_model_name_is_in_assembly(self, parent, model_name_wild_card):

        filter_children_only = filter(lambda x: x['parent'] == parent, self.bill_of_material)
        filter_children_only_list = list(filter_children_only)
        try:
            next(x for x in filter_children_only_list if model_name_wild_card in x['name'])
        except StopIteration:
            logger.info(f'Wildcard {model_name_wild_card} is not in {parent} model.')
            return False
        else:
            logger.info(f'Wildcard {model_name_wild_card} is not in {parent} model.')
            return True

    def check_whether_skeleton_exists(self, cad_parent_model):
        """This method determines whether skeleton model exists. If this model exists it looks for its csys and picks the best matching"""
        try:
            current_model = creopyson.file_get_fileinfo(self.creo_client)['file']
        except:
            pass

        creopyson.file_open(self.creo_client, file_=cad_parent_model)
        skeleton_information = {}

        list_with_skeletons = tuple(bom_recursion(creopyson.bom.get_paths(self.creo_client, skeletons=True, top_level=True)))
        list_with_skeletons = [i for n, i in enumerate(list_with_skeletons) if i not in list_with_skeletons[n + 1:]]

        list_without_skeletons = tuple(bom_recursion(creopyson.bom.get_paths(self.creo_client, skeletons=False, top_level=True)))
        list_without_skeletons = [i for n, i in enumerate(list_without_skeletons) if i not in list_without_skeletons[n + 1:]]

        for every_file in list_without_skeletons:
            list_with_skeletons.remove(every_file)

        if len(list_with_skeletons) == 0:
            messagebox.showerror('Automation status', 'Automation aborted! There is not skeleton model in mastermodel.')
            exit()
        elif len(list_with_skeletons) >= 1:
            cad_skeleton_name = list_with_skeletons[0]
            skeleton_information['skel_name'] = cad_skeleton_name.lower()

        if self.current_master_model() == cad_parent_model:
            creopyson.file_open(self.creo_client, file_=cad_skeleton_name)
            if any(csy['name'] == 'K_M2' for csy in creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')):
                skeleton_information['csy'] = 'K_M2'
            else:
                skeleton_information['csy'] = (creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')[0])['name']
        else:
            skeleton_information['csy'] = self.pick_csy(cad_parent_model, cad_skeleton_name)
            print(f"Skeleton exists, skeleton information is {skeleton_information}.")

        creopyson.file_open(self.creo_client, file_=cad_skeleton_name)
        list_of_coordinate_systems = [subject['name'] for subject in creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')]
        skeleton_information['list_of_csy'] = list_of_coordinate_systems

        if current_model:
            creopyson.file_open(self.creo_client, file_=current_model)

        print(
            f"Skeleton information for {cad_parent_model} are: Skeleton - {skeleton_information['skel_name']}, Default csy - {skeleton_information['csy']}, list of found csys {str(list_of_coordinate_systems)}.")
        logger.info(
            f"Skeleton information for {cad_parent_model} are: Skeleton - {skeleton_information['skel_name']}, Default csy - {skeleton_information['csy']}, list of found csys {str(list_of_coordinate_systems)}.")

        return skeleton_information

    def set_model_convention_on_the_fly(self):
        """This method aligns model naming, due to what users can expect higher percentage of successfully added models."""
        try:
            gmxxxx = self.current_order_number()
            list_of_numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 0]
            for every_component in self.bill_of_material:
                for every_number in list_of_numbers:
                    if "_" + str(every_number) + "_" + gmxxxx in every_component:
                        self.let_proe_rename_pdm_objects(boolean=True)
                        new_name = every_component.replace("_" + str(every_number) + "_" + gmxxxx, "-" + str(every_number) + "_" + gmxxxx)
                        try:
                            creopyson.file_rename(self.creo_client, file_=every_component, new_name=new_name, onlysession=True)
                            self.try_remove_from_ws(every_component)
                            logger.info('Repaired convention in model = ' + new_name)
                        except:
                            pass
        except:
            pass

    def try_remove_from_ws(self, filename_to_remove):
        """This method removes unnecessary models from workspace."""
        list_of_file_names = [filename_to_remove]
        try:
            creopyson.windchill_clear_workspace(self.creo_client, filenames=list_of_file_names)
        except:
            logger.warning("Program was unable to remove " + filename_to_remove + " from workspace.")

    def set_default_view(self):
        try:
            list_views = creopyson.view_list(self.creo_client)
            if any(view_name_in_list == "DEFAULT" for view_name_in_list in list_views):
                creopyson.view_activate(self.creo_client, name="DEFAULT")
        except:
            logger.warning("Program failed to set Default view")

    def change_order_number(self):
        """This method creates new copy of master model."""
        new_number = order_number_entry.get()
        order_number = self.current_order_number()

        current_master_model = self.current_master_model()
        new_master_model = current_master_model.replace(order_number, new_number)
        creopyson.file_rename(self.creo_client, file_=self.current_master_model(), new_name=new_master_model, onlysession=True)

        for every_component in self.bill_of_material:
            if order_number in every_component['name']:
                new_name = every_component['name'].replace(order_number, new_number)
                try:
                    creopyson.file_rename(self.creo_client, file_=every_component['name'], new_name=new_name, onlysession=True)
                    print('this is new model = ' + new_name)
                    logger.info('this is new model = ' + new_name)
                except:
                    logger.exception("message")
                else:
                    self.change_parameter_in_bill_of_material(key='name', new_value=new_name, feat_id=every_component['feat_id'])

    def zs_63_pairing(self):
        """This method uses zs63 as an object which is transformed into bill of material. Bill of material consists dictionaries with keyword SAP group, Material Number, CAD group.
        At transform stage values for keywords SAP_group and Material number are assigned. At pairing stage creo cad group name is assigned to CAD_group keyword."""
        # Transform stage
        self.zs_63.transform_zs_63()
        # Pairing stage
        self.paired_bill_of_material = self.zs_63.pair_converted_zs_63_with_cad_master_model(self.bill_of_material)

    def remove_from_bill_of_material(self, remove_from_bill):
        try:
            self.bill_of_material.remove(remove_from_bill)
        except:
            pass

    def remove_unnecessary_material_numbers(self):

        folder_name = self.current_master_model()[0:3].upper().replace('_', '')
        delete_exclude_file = delete_exclude_path + folder_name + '\\DeleteExclude.csv'
        order_number = self.current_order_number()

        with open(delete_exclude_file, 'r',  errors='replace') as csv_file:
            csv_reader = csv.reader(csv_file)
            delete_exclude_list = []
            for line in csv_reader:
                if line[0] != "":
                    delete_exclude_list.append(line[0])
        tuple_bom = tuple(self.bill_of_material)
        for each_dict in tuple_bom:
            if order_number not in each_dict['name']:
                # Group will be skipped if the parent is in DeleteExclude.csv
                if any(ele.lower() in each_dict['parent'] for ele in delete_exclude_list):
                    print(f"Group {each_dict['parent']} is ignored. Wildcard is in delete exclude file")
                    logger.info(f"Group {each_dict['parent']} is ignored. Wildcard is in delete exclude file")
                    creopyson.file_open(self.creo_client, file_=each_dict['parent'])
                    self.try_to_resume_all()
                elif any(element['ERP_number'] in each_dict['name'] for element in self.paired_bill_of_material):
                    print(f"Component {each_dict['name']} in {each_dict['parent']} is ignored because this material number occurs in zs63.")
                    logger.info(f"Component {each_dict['name']} in {each_dict['parent']} is ignored because this material number occurs in zs63.")
                elif (each_dict['name'][0:-4]).isnumeric() or (each_dict['name'][0:6]).isnumeric() and not any(element['ERP_number'] in each_dict['name'] for element in self.paired_bill_of_material) :
                    # Here will be deleted models that are not in zs 63 file.
                    creopyson.file_open(self.creo_client, file_=each_dict['parent'])
                    self.try_delete_model(model_to_delete=each_dict['name'])
                    self.remove_from_bill_of_material(remove_from_bill=each_dict)
                    print(f"Group {each_dict['name']} is being removed  from {each_dict['parent']}.")
                    logger.info(f"Group {each_dict['name']} is being removed  from {each_dict['parent']}.")
                else:
                    # Here we will avoid files which may have wrong naming. Their data quality is so low that we would rather avoid to delete them from master model.
                    creopyson.file_open(self.creo_client, file_=each_dict['parent'])
                    if each_dict['status'] == 'ACTIVE':
                        self.try_suppress_file(id=each_dict['feat_id'])
        # Resuming file which are found in ZS 63
        tuple_bom = tuple(self.bill_of_material)
        for each_dict in tuple_bom:
            if any(ele['ERP_number'] in each_dict['name'] for ele in self.paired_bill_of_material):
                print(f"Component {each_dict['name']} in {each_dict['parent']} is ignored because this material number occurs in zs63.")
                logger.info(f"Component {each_dict['name']} in {each_dict['parent']} is ignored because this material number occurs in zs63.")
                if each_dict['status'] == 'SUPPRESSED':
                    creopyson.file_open(self.creo_client, file_=each_dict['parent'])
                    self.try_to_resume_all()
                    creopyson.file_close_window(self.creo_client)

    def create_coordinate_system(self, name_of_csy, constraint_to='DEFAULT'):

        creopyson.file_regenerate(self.creo_client)
        list_csys_model = creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM', no_comp=False)
        if not any(x['name'].lower() == name_of_csy.lower() for x in list_csys_model):

            create_csy_script_1 = "~ PopupOver `main_dlg_cur` `sstbar_popup` 1 `Sst_bar.n_sels_show_bin`;\
                                    ~ Open `main_dlg_cur` `sstbar_popup`;\
                                    ~ Close `main_dlg_cur` `sstbar_popup`;\
                                    ~ Activate `main_dlg_cur` `buffer_clean`;\
                                    ~ Command `ProCmdMdlTreeSearch` ;\
                                    ~ Input `selspecdlg0` `SelOptionRadio` `Feature`;\
                                    ~ Update `selspecdlg0` `SelOptionRadio` `Feature`;\
                                    ~ Input `selspecdlg0` `LookByOptionMenu` `Feature`;\
                                    ~ Update `selspecdlg0` `LookByOptionMenu` `Feature`;\
                                    ~ Select `selspecdlg0` `RuleTab` 1 `Misc`;\
                                    ~ Update `selspecdlg0` `ExtRulesLayout.ExtBasicIDLayout.InputIDPanel` `CSY_ID`;\
                                    ~ Activate `selspecdlg0` `ExtRulesLayout.ExtBasicIDLayout.InputIDPanel`;\
                                    ~ Activate `selspecdlg0` `EvaluateBtn`;\
                                    ~ Select `selspecdlg0` `ResultList` -1;\
                                    ~ Activate `selspecdlg0` `ApplyBtn`;\
                                    ~ Activate `selspecdlg0` `CancelButton`;"

            create_csy_script_2 = "~ Command `ProCmdDatumCsys` ;\
                                    ~ Select `Odui_Dlg_00` `pg_vis_tab` 1 `tab_3`;\
                                    ~ Input `Odui_Dlg_00` `t3.datum_csys_name` `K_SAGROUP`;\
                                    ~ Update `Odui_Dlg_00` `t3.datum_csys_name` `K_SAGROUP`;\
                                    ~ FocusOut `Odui_Dlg_00` `t3.datum_csys_name`;\
                                    ~ Activate `Odui_Dlg_00` `stdbtn_1`;\
                                    ~ PopupOver `main_dlg_cur` `sstbar_popup` 1 `Sst_bar.n_sels_show_bin`;\
                                    ~ Open `main_dlg_cur` `sstbar_popup`;\
                                    ~ Close `main_dlg_cur` `sstbar_popup`;\
                                    ~ Activate `main_dlg_cur` `buffer_clean`;"

            list_csys_model = creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')

            if any(coordinate_system['name'].upper() == constraint_to.upper() for coordinate_system in list_csys_model):
                constraint_to_id = next(item for item in list_csys_model if item['name'].lower() == constraint_to.lower())['feat_id']
                logger.info(f"Chosen coordinate system is {constraint_to} with id {str(constraint_to_id)}.")
            elif constraint_to == 'DEFAULT':
                constraint_to_id = list_csys_model[0]['feat_id']
                logger.info(f"Chosen coordinate system is {constraint_to} with id {str(constraint_to_id)}.")
            else:
                constraint_to_id = list_csys_model[0]['feat_id']
                logger.info(f"Chosen coordinate system is {constraint_to} with id {str(constraint_to_id)}.")

            create_csy_script_1 = create_csy_script_1.replace('CSY_ID', str(constraint_to_id))
            create_csy_script_2 = create_csy_script_2.replace('K_SAGROUP', name_of_csy)

            creopyson.interface_mapkey(self.creo_client, create_csy_script_1)
            time.sleep(3)
            creopyson.interface_mapkey(self.creo_client, create_csy_script_2)

            continue_after_mapkey = False

            wait_time_before_run_time = 0
            while not continue_after_mapkey or wait_time_before_run_time > 20:
                time.sleep(0.2)
                wait_time_before_run_time += 0.2
                if any(x['name'].lower() == name_of_csy.lower() for x in creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')):
                    continue_after_mapkey = True

            list_of_coordinate_systems = [subject['name'] for subject in creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')]

            return list_of_coordinate_systems

    def create_sa_groups(self):
        """Creating of empty special assemblies. From CAD perspective - group is assembled to master model and renamed"""
        creopyson.file_open(self.creo_client, file_=self.default_master_model)
        sa_list = []
        # filtering out only SAP group name
        for each_dict in self.paired_bill_of_material:
            if 'SA' in each_dict['SAP_group_name']:
                sa_list.append(each_dict['SAP_group_name'].replace('.', ''))
        # Removing of duplicates.
        sa_list = [i for n, i in enumerate(sa_list) if i not in sa_list[n + 1:]]

        current_master_model = self.current_master_model()
        order_number = self.current_order_number()
        skeleton_information = self.check_whether_skeleton_exists(current_master_model)
        machine_type = current_master_model[0:2].replace('_', '')

        creopyson.file_open(self.creo_client, file_=skeleton_information['skel_name'])

        for each_sap_group in sa_list:
            coordinate_system_name = f'K_{each_sap_group}'.strip().upper()
            list_of_coordinate_systems = self.create_coordinate_system(constraint_to=skeleton_information['csy'], name_of_csy=coordinate_system_name)

        creopyson.file_open(self.creo_client, file_=current_master_model)

        for each_sap_group in sa_list:
            cad_group = machine_type + '_' + each_sap_group + '_' + order_number + '.asm'
            coordinate_system_name = f'K_{each_sap_group}'.strip().upper()
            print(f'CAD group is {cad_group}.')
            try:
                if self.check_whether_model_exists(erp_material_number=cad_group) == '':
                    # Model does not exist. Therefore nwe model has to be created.
                    creopyson.file_open(self.creo_client, file_='MACHINETYPE_SAGROUP_ORDERNUMBER.ASM')
                    creopyson.file_rename(self.creo_client, file_="MACHINETYPE_SAGROUP_ORDERNUMBER.ASM", new_name=cad_group, onlysession=True)
                    creopyson.feature_rename(self.creo_client, new_name=coordinate_system_name, name="K_SAGROUP", file_=cad_group)
                    # Newly created model is saved.
                    default_csy = creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')[0]['name']
                    creopyson.file_open(self.creo_client, file_=current_master_model)
                    creopyson.file_assemble(self.creo_client, file_=cad_group, into_asm=current_master_model, ref_model=skeleton_information['skel_name'],
                                            constraints=[{"asmref": coordinate_system_name, "compref": default_csy, "type": "csys"}])
                    self.add_models_in_opened_group_to_bom()
                    self.check_whether_is_destination_group()
                elif self.check_whether_model_exists(erp_material_number=cad_group) != '' and all(bill_group['name'].lower() != cad_group.lower() for bill_group in self.bill_of_material):
                    creopyson.file_open(self.creo_client, file_=cad_group)
                    default_csy = creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')[0]['name']
                    creopyson.file_open(self.creo_client, file_=current_master_model)
                    creopyson.file_assemble(self.creo_client, file_=cad_group, into_asm=current_master_model, ref_model=skeleton_information['skel_name'],
                                             constraints=[{"asmref": coordinate_system_name, "compref": default_csy, "type": "csys"}])
                    self.add_models_in_opened_group_to_bom()
                    self.check_whether_is_destination_group()
                else:
                    pass
            except:
                creopyson.file_assemble(self.creo_client, into_asm=current_master_model, file_=cad_group, constraints=[{"type": "fix"}], package_assembly=True)

        creopyson.file_open(self.creo_client, file_=current_master_model)
        self.add_models_in_opened_group_to_bom(level_of_master_model_tree=1)

    def assemble_models_to_master_model(self):
        """Assemble model function for suited for assembling process. Loading bar has been commented to make application more robust."""
        current_master_model = self.current_master_model()

        for every_erp_sap_name_cad_group in self.paired_bill_of_material:
            erp_material_number = every_erp_sap_name_cad_group['ERP_number']
            cad_parent_model = every_erp_sap_name_cad_group['CAD_group_name']
            if cad_parent_model != 'Not defined':
                #creopyson.file_open(self.creo_client, file_=cad_parent_model)
                self.assemble_model(erp_material_number, cad_parent_model)

        creopyson.file_open(self.creo_client, file_=current_master_model)

    def check_non_assembled_models(self):
        """ This function is purposed to provide feedback to user about assembling process.
            It collects all non - assembled models and check their existence.
            If model exists, function creates pdf and saves this pdf to PartToConsider folder.
            This method of controlling non-assembled parts eases work of creators."""
        current_master_model = self.current_master_model()
        self.clear_bill_of_material()
        self.create_master_model_bill_of_material_with_suppressed(levels=3)
        self.check_whether_is_destination_group()
        self.paired_bill_of_material = self.zs_63.pair_converted_zs_63_with_cad_master_model(self.bill_of_material)

        newly_created_groups = []

        for each in self.paired_bill_of_material:
            if all(each['ERP_number'] not in x['name'] for x in self.bill_of_material):
                model_name = self.check_whether_model_exists(erp_material_number=each['ERP_number'])
                if model_name:
                    jpeg_name_raw = model_name.replace('.', '_') + "_" + each['SAP_group_name'].replace('.', '_')
                    jpeg_name = ""
                    for each_char in jpeg_name_raw:
                        if each_char.isalnum() or each_char == "_":
                            jpeg_name += each_char
                    creopyson.file_open(self.creo_client, file_=model_name)
                    image_location_dict = creopyson.interface_export_image(self.creo_client, file_type="JPEG", filename=jpeg_name)
                    image_location = image_location_dict["dirname"] + image_location_dict["filename"]
                    final_image_path = os.path.dirname(sys.argv[0]) + '/FeedbackFolder/' + image_location_dict["filename"]
                    print(image_location)
                    new_dict = {'SAP_group_name': each['SAP_group_name'], 'CAD_model_to_assemble': model_name}
                    newly_created_groups.append(new_dict)
                    try:
                        shutil.move(image_location, final_image_path)
                    except:
                        logger.warning("There was some problem to store screenshot of model " + final_image_path)
                        logger.exception("message")

        try:
            # Here will be filtered out only destination groups
            filter_destinations_only = filter(lambda x: x['destination_group'] == 'yes', self.bill_of_material)
            filter_destinations_only_list = list(filter_destinations_only)
            # Merging groups from self.paired_bill_of_material and filter_destinations_only_list
            groups_with_sap_number_and_are_destinated = []
            for each_group in self.paired_bill_of_material:
                if any(each_group['CAD_group_name'] == x['name'] for x in filter_destinations_only_list):
                    groups_with_sap_number_and_are_destinated.append(each_group)
            # Main loop of this section follows
            # Newly created groups consists
            for each in newly_created_groups:
                try:
                    # We will ignore groups named m. (probably foundation plans or clamp unit plan)
                    if 'm.' not in each['SAP_group_name'].lower() and 'ze' not in each['SAP_group_name'].lower():
                        similarities = []
                        for each_bom_group in groups_with_sap_number_and_are_destinated:
                            if each['SAP_group_name'][0:2].lower() == each_bom_group['SAP_group_name'][0:2].lower() and len(each['SAP_group_name']) == len(each_bom_group['SAP_group_name']):
                                measure_similarity_obj = SequenceMatcher(None, each['SAP_group_name'], each_bom_group['SAP_group_name'])
                                measure_similarity = measure_similarity_obj.ratio()
                                dict_to_add = {'similar_name': each_bom_group['CAD_group_name'], 'level_of_similarity': measure_similarity}
                                similarities.append(dict_to_add)

                        max_similarity = max(dict_added['level_of_similarity'] for dict_added in similarities)
                        max_dict_in_sim = next(item for item in similarities if item['level_of_similarity'] == max_similarity)
                        max_dict_in_filtered_destinations = next(item for item in groups_with_sap_number_and_are_destinated if item['CAD_group_name'] == max_dict_in_sim['similar_name'])
                        most_similar_group_in_bom = next(item for item in self.bill_of_material if item['name'] == max_dict_in_filtered_destinations['CAD_group_name'])
                        print(f"Maxed out value for {each['SAP_group_name']} is {max_similarity} and group is {max_dict_in_filtered_destinations['CAD_group_name']}.")
                        origin_name = max_dict_in_filtered_destinations['CAD_group_name']
                        missing_group_sap_format = each['SAP_group_name'].lower().strip()
                        missing_group_sap_format = missing_group_sap_format.replace('.', '_')
                        position_of_missing_guaranteed_sap_name = origin_name.index(missing_group_sap_format[0:2])
                        string_which_will_be_replaced = origin_name[position_of_missing_guaranteed_sap_name:position_of_missing_guaranteed_sap_name + len(missing_group_sap_format)]
                        new_name = origin_name.replace(string_which_will_be_replaced, missing_group_sap_format)
                        # Here is set source sister assembly.
                        new_csy_name = (f'K_{missing_group_sap_format}').upper()
                        print(f"New name will be {new_name}.")
                        creopyson.file_open(self.creo_client, file_=most_similar_group_in_bom['parent'])

                        try:
                            creopyson.file_open(self.creo_client, file_=most_similar_group_in_bom['parent'])
                            skeleton_information = self.check_whether_skeleton_exists(cad_parent_model=most_similar_group_in_bom['parent'])
                            creopyson.file_open(self.creo_client, file_=skeleton_information['skel_name'])
                            self.create_coordinate_system(name_of_csy=new_csy_name, constraint_to='DEFAULT')
                            creopyson.file_open(self.creo_client, file_=most_similar_group_in_bom['parent'])

                            # Here we will check whether model exists already - model probably do not exist
                            if self.check_whether_model_exists(erp_material_number=new_name) == '':

                                # TODO: comment and define !
                                # 'MACHINETYPE_SAGROUP_ORDERNUMBER.ASM' model will be openend and renamed to our needs
                                creopyson.file_open(self.creo_client, file_='MACHINETYPE_SAGROUP_ORDERNUMBER.ASM')
                                creopyson.file_rename(self.creo_client, file_="MACHINETYPE_SAGROUP_ORDERNUMBER.ASM", new_name=new_name, onlysession=True)
                                creopyson.feature_rename(self.creo_client, new_name=new_csy_name, name="K_SAGROUP", file_=new_name)
                                # we will also set default CSY of newly created model
                                default_csy = creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')[0]['name']
                                # here we can assemble model
                                creopyson.file_assemble(self.creo_client, file_=new_name, into_asm=most_similar_group_in_bom['parent'], ref_model=skeleton_information['skel_name'],
                                                        constraints=[{"asmref": new_csy_name, "compref": default_csy, "type": "csys"}])

                                print(f"Self repair part First condition passed - {new_name} has been assembled to {most_similar_group_in_bom['parent']}.")
                                logger.info(f"Self repair part First condition passed  - {new_name} has been assembled to {most_similar_group_in_bom['parent']}.")

                                # if the operation was successful we will open parent group and updated our bill of material by which we will avoid next if condition
                                creopyson.file_open(self.creo_client, file_=most_similar_group_in_bom['parent'])
                                self.add_models_in_opened_group_to_bom()
                                self.check_whether_is_destination_group()

                            # Second condition of self-repair. This condition is correct if new_name exists and model is not in bill of material.
                            if self.check_whether_model_exists(erp_material_number=new_name) != '' and all(bill_group['name'] != new_name.lower() for bill_group in self.bill_of_material):
                                creopyson.file_open(self.creo_client, file_=new_name)
                                # Group may exist so we need to define its coordination system - this possibility is not high
                                default_csy = creopyson.feature_list(self.creo_client, type_='COORDINATE SYSTEM')[0]['name']
                                creopyson.file_assemble(self.creo_client, file_=new_name, into_asm=most_similar_group_in_bom['parent'], ref_model=skeleton_information['skel_name'],
                                                        constraints=[{"asmref": new_csy_name, "compref": default_csy, "type": "csys"}])
                                creopyson.file_open(self.creo_client, file_=most_similar_group_in_bom['parent'])

                                print(f"Self repair part Second Condition passed - {new_name} has been assembled to {most_similar_group_in_bom['parent']}. ")
                                logger.info(f"Self repair part Second Condition passed - {new_name} has been assembled to {most_similar_group_in_bom['parent']}. ")

                                # if the operation was successful we will open parent group and updated our bill of material by which we will avoid next if condition
                                creopyson.file_open(self.creo_client, file_=most_similar_group_in_bom['parent'])
                                self.add_models_in_opened_group_to_bom()
                                self.check_whether_is_destination_group()
                        except:
                            pass
                        else:
                            creopyson.file_open(self.creo_client, file_=most_similar_group_in_bom['parent'])
                            try:
                                new_name_id = next(item for item in self.bill_of_material if item['name'] == new_name)['feat_id']
                                self.try_suppress_file(id=new_name_id)
                            except:
                                pass
                            self.assemble_model(erp_material_number=each['CAD_model_to_assemble'], cad_parent_model=new_name)
                except:
                    sys.exc_info()
                    logger.exception("message")
                    print(sys.exc_info())
        except:
            sys.exc_info()
            logger.exception("message")
            print(sys.exc_info())

        creopyson.file_open(self.creo_client, file_=current_master_model)

    def try_to_resume_all(self):
        """Third rebuilt of this function. Experimenting with simple ResumeAll command to make automation more robust"""
        try:
            resume_all_script = '~ Command `ProCmdResumeAll`;'
            all_resumed = False
            creopyson.interface_mapkey(self.creo_client, script=resume_all_script)

            while not all_resumed:
                listed_components = creopyson.feature_list(self.creo_client, type_='COMPONENT')
                if all(component['status'] == 'ACTIVE' or component['status'] == 'UNREGENERATED' for component in listed_components):
                    all_resumed = True
                else:
                    time.sleep(0.2)
        except:
            pass
        else:
            for each_id in listed_components:
                self.change_parameter_in_bill_of_material(key='status', new_value='ACTIVE', feat_id=each_id['feat_id'])

    def try_suppress_file(self, id):

        print(f"ID of the suppressing model is {str(id)}.")
        try:
            creopyson.feature_suppress(self.creo_client, name=int(id), with_children=True)
        except:
            print(f"There was problem with resuming feature {str(id)} in {creopyson.file_get_fileinfo(self.creo_client)['file']}.")
        else:
            self.change_parameter_in_bill_of_material(key='status', new_value='SUPPRESSED', feat_id=id)

    def session_mapkeys(self, regenerate=False, mc=False, save=False):

        regenerate_script = "~ Command `ProCmdRegenAuto` ;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` `casc10647928`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` `ProCmdRegenAuto(1)`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` `ProCmdRegenCust`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` ``;\
                #CURRENT VALS;#CURRENT VALS;#CURRENT VALS;#CURRENT VALS;#CURRENT VALS;#CURRENT VALS;#CURRENT VALS;\
                ~ Activate `storage_conflicts` `Close_PushButton`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` `` `commands_group`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` `ProCmdMCRegen`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` `ProCmdRegenAuto`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` `casc10647928`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` `ProCmdRegenAuto(1)`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` `ProCmdRegenCust`;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` ``;\
                ~ Trigger `search_panel_list_of_main_dlg_w1` `search_panel_list` ``;"
        save_script = "~ Command `ProCmdModelSave` ;\
                ~ Activate `file_saveas` `OK`;~ Select `storage_conflicts` `CascadeButton1`;\
                ~ Close `storage_conflicts` `CascadeButton1`;\
                ~ Activate `storage_conflicts` `Resol1`;\
                ~ Close `storage_conflicts` `CascadeButton1`;\
                ~ Activate `storage_conflicts` `OK_PushButton`;"
        mc_script = "~ Activate `main_dlg_cur` `page_Annotate_control_btn` 1;\
                ~ Command `lang_all` ;~ Command `ProCmdViewRepaint`;\
                ~ Close `main_dlg_cur` `appl_casc`;~ Command `ProCmdMCModelCHECK` ;;\
                ~ Activate `storage_conflicts` `OK_PushButton`;#TOP LEVEL;#OBERSTE EBENE;\
                ~ Close `main_dlg_cur` `appl_casc`;\
                ~ Command `ProCmdMCModelCHECK` ;#TOP LEVEL;#OBERSTE EBENE;\
                ~ Close `main_dlg_cur` `appl_casc`;\
                ~ Command `ProCmdMCModelCHECK` ;;#TOP LEVEL;#OBERSTE EBENE;\
                ~ Activate `main_dlg_cur` `page_Annotate_control_btn` 1;~ Command `lang_de` ;\
                ~ Activate `main_dlg_cur` `user_custom_page_46721592_control_btn` 1;\
                ~ Activate `main_dlg_cur` `page_Model_control_btn` 1;\
                ~ Command `ProCmdViewRepaint`;"
        try:
            if regenerate:
                creopyson.interface_mapkey(self.creo_client, script=regenerate_script)
            if save:
                creopyson.interface_mapkey(self.creo_client, script=save_script)
            if mc:
                if creopyson.parameter_exists(self.creo_client, name="MC_ERRORS") and next(item for item in creopyson.parameter_list(self.creo_client) if item['name'] == 'MC_ERRORS')['value'] == 0:
                    pass
                else:
                    creopyson.interface_mapkey(self.creo_client, script=mc_script)
        except:
            pass

    def get_config_value(self, name_of_config):
        value = creopyson.creo_get_config(self.creo_client, name=name_of_config)
        print(value)
        return value

    def auto_conflicts_resolution(self, boolean):
        if boolean:
            creopyson.creo_set_config(client=self.creo_client, name="dm_auto_conflict_resolution", value="yes")
        else:
            creopyson.creo_set_config(client=self.creo_client, name="dm_auto_conflict_resolution", value="no")

    def mass_property_calculate(self, boolean):
        if boolean:
            creopyson.creo_set_config(client=self.creo_client, name="mass_property_calculate", value="check_upon_save")
        else:
            creopyson.creo_set_config(client=self.creo_client, name="mass_property_calculate", value="by_request")

    def regen_notebook_w_assem(self, boolean):
        if boolean:
            creopyson.creo_set_config(client=self.creo_client, name="regen_notebook_w_assem", value="no")
        else:
            creopyson.creo_set_config(client=self.creo_client, name="regen_notebook_w_assem", value="yes")

    def add_weld_mp(self, boolean):
        if boolean:
            creopyson.creo_set_config(client=self.creo_client, name="add_weld_mp", value="yes")
        else:
            creopyson.creo_set_config(client=self.creo_client, name="add_weld_mp", value="no")

    def configs_manipulation(self, api_mode):
        """Used configuration option which are suited to enhance performance speed of the application."""
        if api_mode:
            self.regen_notebook_w_assem(boolean=False)
            self.mass_property_calculate(boolean=True)
            self.auto_conflicts_resolution(boolean=True)
            self.regenerate_read_only_config_control(boolean=False)
            self.let_proe_rename_pdm_objects(boolean=True)
            self.display_comps_to_assemble(boolean=False)
            self.add_weld_mp(boolean=False)
        else:
            self.regen_notebook_w_assem(boolean=True)
            self.mass_property_calculate(boolean=False)
            self.auto_conflicts_resolution(boolean=False)
            self.regenerate_read_only_config_control(boolean=True)
            self.let_proe_rename_pdm_objects(boolean=False)
            self.display_comps_to_assemble(boolean=True)
            self.add_weld_mp(boolean=True)


class Zs63notPickedError(Exception):
    """Error class object. This special error has been designed to do not interrupt application if the user does not select zs_63.txt file."""
    def __init__(self, zs63_file, message="Zs63 file has not been selected."):
        self.zs63_file = zs63_file
        self.message = message
        super().__init__(self.message)


class CancelByUserError(Exception):
    """Error class object. This special error has been designed to do not interrupt application if the user does not select zs_63.txt file."""
    def __init__(self, zs63_file, message="User has canceled his operation."):
        self.zs63_file = zs63_file
        self.message = message
        super().__init__(self.message)


class Zs63:
    """This class refers to text file from zs63 SAP transaction. Loading of this function has to be enhanced with SAP scripting."""
    def __init__(self):
        self.final_folder_path = os.path.dirname(sys.argv[0]) + '\\ErpBom\\ZS_63.txt'
        self.all_lists = []
        self.m_groups_list = []
        self.ze_groups_list = []
        self.sa_groups_list = []
        # Methods on innit
        self.get_zs63_file()
        self.transform_zs_63()

    def get_zs63_file(self):
        """This method loads zs63 file to application"""
        remove_files_from_folder(erp_folder_path)
        remove_files_from_folder(feedback_folder_path)
        root.filename = filedialog.askopenfilename(initialdir='\\', title='Choose ZS63 file', filetypes=(('text files', '*.txt'), ('all files', '*.*')))
        current_folder_path = root.filename

        try:
            shutil.copyfile(current_folder_path, self.final_folder_path)
        except FileExistsError:
            os.replace(current_folder_path, self.final_folder_path)
        except FileNotFoundError:
            raise Zs63notPickedError('ZS63 file')

    def transform_zs_63(self):
        """"This method transforms zs63 file to 3 lists (m_groups, ze_groups, sa_groups). Later these groups are merged to all_lists (list type)."""
        m_groups_list = []
        ze_groups_list = []
        sa_groups_list = []
        zs_63 = []
        special_allowed_signs = ['.', '#']

        with open(self.final_folder_path,  errors='replace') as zs_data:

            zs_63_raw = zs_data.readlines()

            # This is encoding part - necessary to implement because of various operating systems - CHINA operating system issues
            for each_line in zs_63_raw:
                each_line.encode("utf-8", "ignore")
                line_text = str(each_line.encode("utf-8", "ignore"))
                line_text = line_text.replace("b'", "")
                line_text = line_text.replace("n'", "")
                zs_63.append(str(line_text))

            # Newly created section where some risky signs will be removed from zs_63.txt - Makes the process more robust
            if type(zs_63) == list:
                for index, each_element in enumerate(zs_63):
                    zs_63[index] = specific_symbols_for_line_in_zs63(each_element)

            # Here we try to set injection unit size for MX
            try:
                global zs_63_injection_unit
                zs_63_injection_unit = ""
                if "mx" in self.current_master_model:
                    following_line = False
                    for line in zs_63:
                        if following_line:
                            list_of_splited_line = line.split("/")
                            zs_63_injection_unit_raw = list_of_splited_line[1]
                            zs_63_injection_unit = zs_63_injection_unit_raw
                            for each_letter in zs_63_injection_unit_raw:
                                if each_letter == "0":
                                    zs_63_injection_unit = zs_63_injection_unit[1:]
                                else:
                                    break
                            print(zs_63_injection_unit)
                            break
                        if "--------------------------------------------------------" in line:
                            following_line = True
                    logger.info('Machine type is MX therefore  program tried to set injection unit according to ZS_63, the value is ' + zs_63_injection_unit)
            except:
                logger.exception("message")

            for line in zs_63:
                first_symbols = line[0:3]
                line = line.replace(' 000  ', '      ')
                if 'GM' or 'PL' in first_symbols:
                    m_groups_list.append(line)
                if 'ZE#' in first_symbols:
                    ze_groups_list.append(line)
                if 'SA' in first_symbols:
                    sa_groups_list.append(line)

            for every_element in m_groups_list:
                split_element = every_element.split()
                pair_group_and_number = []
                for each_split_element in split_element:
                    each_split_element = each_split_element.strip()
                    if 6 < len(each_split_element) < 9 and '7' not in each_split_element[0:1] and each_split_element.isnumeric():
                        pair_group_and_number.append(each_split_element)
                    if 'm' in each_split_element and '.' in each_split_element or 'c' in each_split_element and '.' in each_split_element:
                        pair_group_and_number.append(each_split_element)
                if len(pair_group_and_number) == 2:
                    mat_nr = pair_group_and_number[0]
                    sap_group = pair_group_and_number[1]
                    group_mat_nr_dict = {'SAP_group_name': sap_group, 'ERP_number': mat_nr, 'CAD_group_name': 'Not defined'}
                    self.all_lists.append(group_mat_nr_dict.copy())

            for every_element in ze_groups_list:
                split_element = every_element.split()
                pair_group_and_number = []
                for each_split_element in split_element:
                    each_split_element = each_split_element.strip()
                    if 6 < len(each_split_element) < 9 and '7' not in each_split_element[0:1] and each_split_element.isnumeric():
                        pair_group_and_number.append(each_split_element)
                    if '.' in each_split_element:
                        try:
                            float(each_split_element)
                        except ValueError:
                            pass
                        else:
                            pair_group_and_number.append('ZE' + each_split_element)
                if len(pair_group_and_number) == 2:
                    mat_nr = pair_group_and_number[0]
                    sap_group = pair_group_and_number[1]
                    group_mat_nr_dict = {'SAP_group_name': sap_group, 'ERP_number': mat_nr, 'CAD_group_name': 'Not defined'}
                    self.all_lists.append(group_mat_nr_dict.copy())

            list_of_useful_neukonst = ['7000131', '7009132', '7000134', '7009107', '7009103', '7000182', '7000134', '7000107', '7000146', '7000132', '7000118']

            for every_element in sa_groups_list:
                split_element = every_element.split()
                pair_group_and_number = []
                for each_split_element in split_element:
                    each_split_element = each_split_element.strip()
                    if 4 > len(each_split_element) > 1 and each_split_element.isnumeric():
                        sa_value = int(each_split_element)
                        if 29 < sa_value < 999:
                            if len(each_split_element) == 2:
                                each_split_element = 'SA0' + each_split_element
                            if len(each_split_element) == 3:
                                each_split_element = 'SA' + each_split_element
                            pair_group_and_number.append(each_split_element)
                    if 6 < len(each_split_element) < 9 and '7' not in each_split_element[0:1] and each_split_element.isnumeric():
                        pair_group_and_number.append(each_split_element)
                    # Vladimir Truchly suggested to use this list for better creation of empty SA models.
                    if any(useful_neukonst == each_split_element for useful_neukonst in list_of_useful_neukonst):
                        pair_group_and_number.append(each_split_element)
                if len(pair_group_and_number) == 2:
                    mat_nr = pair_group_and_number[1]
                    sap_group = pair_group_and_number[0]
                    group_mat_nr_dict = {'SAP_group_name': sap_group, 'ERP_number': mat_nr, 'CAD_group_name': 'Not defined'}
                    self.all_lists.append(group_mat_nr_dict.copy())

    def pair_converted_zs_63_with_cad_master_model(self, bill_of_material):
        """This method add mastermodel group name from Creo to CAD_group Keyword. Special approach is chosen if there are multiple M6 (more component master model)"""
        try:
            if special_sign != None:
                m6_list = []
                previous_dict = {}
                # count how many times M6_M01 occures in all lists
                for every_dict in self.all_lists:
                    if 'M6' in every_dict['SAP_group_name']:
                        if not any(x == every_dict['SAP_group_name'] for x in m6_list):
                            m6_list.append(every_dict['SAP_group_name'])
                        else:
                            if every_dict['SAP_group_name'] != previous_dict:
                                every_dict['SAP_group_name'] = special_sign + every_dict['SAP_group_name']
                                previous_dict = every_dict['SAP_group_name']
        except NameError:
            pass

        for every_dict in self.all_lists:

            # Here we are going to generate possible names

            # 1st quality is the name which equals to the name of ['SAP_group_name'] for example M6.E30, ZE07.12 etc.
            name_quality_level_1 = every_dict['SAP_group_name']
            # 2nd quality replaces dod with underscore M6_E30, M22_M20_2 etc.
            name_quality_level_2 = name_quality_level_1.replace('.', '_')
            # 3rd quality takes level 2 and second underscore replaces with dash M22_M20-2
            name_quality_level_2_split = name_quality_level_2.split('_')
            name_quality_level_3 = ""
            if len(name_quality_level_2_split) == 3:
                name_quality_level_3 = name_quality_level_2_split[0] + '_' + name_quality_level_2_split[1] + '-' + name_quality_level_2_split[2]
            # 4th quality is specialized for ZE machines - especially for the naming of pads. In CAD model is no Zero on third position ZE02_50
            name_quality_level_4 = ""
            if name_quality_level_2[2] == "0":
                name_quality_level_4 = name_quality_level_2[0: 2:] + name_quality_level_2[2 + 1::]
            # 5th quality adds X to ending of group name
            name_quality_level_5 = ""
            try:
                if name_quality_level_2[len(name_quality_level_2) - 1] != ('_'):
                    name_quality_level_5 = name_quality_level_2[0:len(name_quality_level_2) - 1] + 'X'
            except IndexError:
                pass
            # 6th quality removes last and adds X to ending of group name
            name_quality_level_6 = ""
            try:
                if name_quality_level_5[len(name_quality_level_5) - 1] != ("_"):
                    name_quality_level_6 = name_quality_level_5[0:len(name_quality_level_5) - 2] + 'X'
            except IndexError:
                pass
            # 7th quality takes string changes dot to dash ZE25.50 -> ZE25-50
            name_quality_level_7 = name_quality_level_1.replace('.', '-')
            # 8th quality uses level 7 and removes 0 from third index
            name_quality_level_8 = ""
            if name_quality_level_7[2] == "0":
                name_quality_level_8 = name_quality_level_7[0: 2:] + name_quality_level_7[2 + 1::]
            # 9th level takes another number from string (if it is not bordering)
            name_quality_level_9 = name_quality_level_5.replace('.', '-')
            # 10th level takes another number from string (if it is not bordering)
            name_quality_level_10 = name_quality_level_6.replace('.', '-')
            # 11th quality takes level 2 and first underscore removes E_M5_M50 -> EM5_M50
            name_quality_level_11 = ""
            if len(name_quality_level_2_split) == 3:
                name_quality_level_11 = name_quality_level_2_split[0] + '' + name_quality_level_2_split[1] + '_' + name_quality_level_2_split[2]
            # At this point we can build list
            all_tested_names = [name_quality_level_1, name_quality_level_2, name_quality_level_3,
                                name_quality_level_4, name_quality_level_5, name_quality_level_6,
                                name_quality_level_7, name_quality_level_8, name_quality_level_9,
                                name_quality_level_10, name_quality_level_11]

            # Deletion of empty elements of the list
            for every_tested_name in all_tested_names:
                if every_tested_name == "":
                    all_tested_names.remove(every_tested_name)

            # Finally the pairing it self. Looped throughout every possible name (for specific SAP name) and every CAD group.
            for every_name in all_tested_names:
                for every_cad_group in bill_of_material:
                    if '_' + every_name.lower() + '_' in every_cad_group['name']:
                        print(f"{name_quality_level_1}  paired with  {every_cad_group['name']}.")
                        logger.info(f"{name_quality_level_1}  paired with  {every_cad_group['name']}.")
                        every_dict['CAD_group_name'] = every_cad_group['name']
                        break

        # Short computation to determine quality of pairing process. Output is logged.
        successful_pairing = 0
        for each_dict in self.all_lists:
            if each_dict["CAD_group_name"] != "Not defined":
                successful_pairing = successful_pairing + 1
        if len(self.all_lists) > 1:
            percentage = successful_pairing / len(self.all_lists) * 100
            logger.info(f"Percentage of defined pairs is {str(percentage)} %.")
            print(f"Percentage of defined pairs is {str(percentage)} %.")
            logger.info('end of pairing')
            self.all_lists = [i for n, i in enumerate(self.all_lists) if i not in self.all_lists[n + 1:]]

        for each in self.all_lists:
            print(each)

        return self.all_lists


def bom_recursion(nest_dict, start=True):
    """With this recursion return value of creopyson is structured to the list of simple dictionaries. Source value is structured as nested dictionary/list."""
    if start:
        global list_of_recursed_bom
        list_of_recursed_bom = []

    for key, value in nest_dict.items():
        if isinstance(value, dict):
            bom_recursion(value, start=False)
        elif isinstance(value, list):
            for each in value:
                if isinstance(each, dict):
                    bom_recursion(each, start=False)
        else:
            if key == 'file':
                list_of_recursed_bom.append(value.lower())

    return list_of_recursed_bom


def automation_process():
    """ This procedure covers whole process of assembling.
        All the steps of this essential function are described below in comments.
        Based on input order number, process of assembling is divided to 2 parts."""

    def main_API_thread():

        # With introduction of threading is necessary to disable critical buttons => run, compare ZS63
        # Button Operations
        run_button = CreateControlButton(parent=app, row_grid=13, column_grid=0, command=automation_process, icon_name="run_automation.png")
        run_button.disable_this_button()
        compare_zs_63_button = CreateControlButton(parent=app, row_grid=13, column_grid=1, command=compare_with_zs63_file_button, icon_name="compare_zs_63.png")
        compare_zs_63_button.disable_this_button()
        app_status_button = CreateControlButton(parent=app, row_grid=11, column_grid=2, icon_name='in_progress_status.png', command=None)
        app.update()

        start = time.time()

        try:
            if len(order_number_entry.get()) == 6:
                session = CreoAPI()
                session.create_master_model_bill_of_material_with_suppressed(levels=3)
                session.zs_63_pairing()
                session.set_default_view()
                session.assemble_models_to_master_model()
            else:
                yes_no_preparation = tkinter.messagebox.askquestion(
                    'Invalid Order Number', 'Order number is not valid. Do you want to continue without preparation of master model?'
                                            ' Currently opened master model will be reference for process of automation. '
                                            'This process will delete invalid material numbers and assemble components from loaded ZS63 text file.',
                    icon='warning')

                if yes_no_preparation == 'no':
                    logger.exception("Selection process has been stopped by user.")
                    raise CancelByUserError('Yes no preparation')

                session = CreoAPI(open_master=False)
                #  function loads ZS_63 - this function is must in all cases
                session.create_master_model_bill_of_material_with_suppressed(levels=3)
                #session.zs_63_pairing()
                #session.remove_unnecessary_material_numbers()
                #session.create_sa_groups()
                session.set_model_convention_on_the_fly()
                session.create_master_model_bill_of_material_with_suppressed(levels=3)
                session.zs_63_pairing()
                session.set_default_view()
                session.assemble_models_to_master_model()

            session.check_non_assembled_models()
            session.configs_manipulation(api_mode=False)
            session.session_mapkeys(regenerate=True, mc=True, save=False)
            tkinter.messagebox.showinfo('Automation status', 'Automation completed ! Wait for the finish of ModelCheck.')
            print('Successful finish of the automation.')
            # Button operations
            app_status_button.destroy_this_button()
            run_button.destroy_this_button()
            compare_zs_63_button.destroy_this_button()
        except Zs63notPickedError:
            pass
        except CancelByUserError:
            pass
        except:
            app_status_button.destroy_this_button()
            run_button.destroy_this_button()
            compare_zs_63_button.destroy_this_button()
            sys.exc_info()
            logger.exception("message")
            print("message")
            tkinter.messagebox.showerror(title='Critical error', message=sys.exc_info())
            try:
                session.configs_manipulation(api_mode=False)
            except:
                pass
            open_log_file()

        end = time.time()
        hours, rem = divmod(end - start, 3600)
        minutes, seconds = divmod(rem, 60)
        elapsed_time = ("{:0>2}:{:0>2}:{:05.2f}".format(int(hours), int(minutes), seconds))
        logger.info(elapsed_time)
        print(f'Elapsed time of automation process is {elapsed_time}.')

    global main_thread
    main_thread = threading.Thread(target=main_API_thread, daemon=True)
    main_thread.start()


def only_numerics(sequence):
    seq_type = type(sequence)
    return seq_type().join(filter(seq_type.isdigit, sequence))


def specific_symbols_for_csy(sequence):

    result = ''
    special_allowed_signs = ['_', '-']
    for index, each_sign in enumerate(sequence):
        if each_sign.isalnum() or each_sign in special_allowed_signs:
            result = result + each_sign
    return result


def specific_symbols_for_line_in_zs63(sequence):

    result = ''
    special_allowed_signs = ['.', ' ']
    for index, each_sign in enumerate(sequence):
        if each_sign.isalnum() or each_sign in special_allowed_signs:
            result = result + each_sign
        elif each_sign == '#' and index == 2:
            result = result + each_sign
    return result


def remove_files_from_folder(folder_name):
    set_global_paths()

    for file_or_folder in os.walk(folder_name):
        for fileX in file_or_folder:
            pass
        for each in fileX:
            os.remove(folder_name + '\\' + each)
            logger.info("removing file " + each)


def set_global_paths():
    """Setting global paths for all key folders."""

    global database_path
    database_path = 'DatabaseFolder\\mastermodels_database.xlsx'
    global icons_folder_path
    icons_folder_path = 'Icons\\'
    global feedback_folder_path
    feedback_folder_path = 'FeedbackFolder\\'
    global erp_folder_path
    erp_folder_path = 'ErpBom\\'
    global delete_exclude_path
    delete_exclude_path = 'DeleteExclude\\'


def open_log_file():
    """Button reaction function"""
    log_file = 'assembly_automation.log'
    os.startfile(log_file)


def compare_with_zs63_file_button():
    """Button reaction function."""

    def compare_thread():

        run_button = CreateControlButton(parent=app, row_grid=13, column_grid=0, command=automation_process, icon_name="run_automation.png")
        run_button.disable_this_button()
        compare_zs_63_button = CreateControlButton(parent=app, row_grid=13, column_grid=1, command=compare_with_zs63_file_button, icon_name="compare_zs_63.png")
        compare_zs_63_button.disable_this_button()
        app_status_button = CreateControlButton(parent=app, row_grid=11, column_grid=2, icon_name='in_progress_status.png', command=None)

        remove_files_from_folder(feedback_folder_path)
        session = CreoAPI(open_master=False)
        session.check_non_assembled_models()
        tkinter.messagebox.showinfo('Comparing is completed', 'View missing material numbers in Feedback folder !')

        app_status_button.destroy_this_button()
        run_button.destroy_this_button()
        compare_zs_63_button.destroy_this_button()

    #global main_thread
    #main_thread = threading.Thread(target=compare_thread, daemon=True)
    #main_thread.start()

    session = CreoAPI(open_master=False)
    session.create_master_model_bill_of_material_with_suppressed(levels=3)


def main():
    """ Main function  - Graphical user interface is initialized by this function.
    Necessary to wrap into function because of its usage in Reset button command."""
    global app
    global root

    root = tkinter.Tk()
    root.lift()
    app = Application(master=root)
    app.mainloop()


if __name__ == "__main__":
    main()
