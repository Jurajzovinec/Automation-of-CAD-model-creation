import os
import tkinter
import sys
import creopyson
import xlrd
import logging
import time
import shutil
import csv
from difflib import SequenceMatcher
from PIL import ImageTk, Image
from tkinter import messagebox
from tkinter import filedialog

# ---Logging---
try:
    os.remove("assembly_automation.log")
except:
    pass
logging.basicConfig(filename="assembly_automation.log", level=logging.INFO)
logger = logging.getLogger()


class Application(tkinter.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.geometry("500x600")
        self.master.resizable(0, 0)
        self.master.title('Kraussmaffei Assembly Automation')
        self.rows = 0

        while self.rows < 13:
            master.rowconfigure(self.rows, weight=1)
            master.columnconfigure(self.rows, weight=1)
            self.rows += 1

        set_global_paths()
        self.set_background_theme()
        self.create_entry()
        self.create_buttons()
        self.create_machine_type_list()

    def set_background_theme(self):

        self.background = ImageTk.PhotoImage(Image.open("IconPictures/Graphical_User_Int_Theme.png"))
        self.background_theme = tkinter.Label(self.master, image=self.background).grid(row=0, column=0, rowspan=100, columnspan=100)

    def create_buttons(self):

        self.run_button = CreateControlButton(parent=self, row_grid=13, column_grid=0, command=self.close_graphical_user_interface, icon_name="run_automation.png")
        self.run_button.disable_this_button()
        self.compare_zs_63_button = CreateControlButton(parent=self, row_grid=13, column_grid=1, command=development_function, icon_name="compare_zs_63.png")
        self.reset_button = CreateControlButton(parent=self, row_grid=13, column_grid=2, command=self.reset_graphical_user_interface, icon_name="reset.png")
        self.feedback_button = CreateControlButton(parent=self, row_grid=14, column_grid=0, command=self.open_feedback_folder, icon_name="feedback_folder.png")
        self.source_folder_button = CreateControlButton(parent=self, row_grid=14, column_grid=1, command=self.open_database_folder, icon_name="source_folder.png")
        self.quit_button = CreateControlButton(parent=self, row_grid=14, column_grid=2, command=self.close_graphical_user_interface, icon_name="quit.png")

    def create_entry(self):

        global order_number_entry
        order_number_entry = tkinter.Entry(self.master)
        order_number_entry.config(width=20, font=('Helvetica', 14), borderwidth=4)
        order_number_entry.insert(0, '--New_Order_Number--')
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
        self.run_button.enable_this_button()
        self.machine_type_drop_down_menu.disable_this_dropdown_menu()
        self.confirm_selected_machinetype.disable_this_button()
        selected_machine_type = self.machine_type_drop_down_menu.what_is_picked_option()

        # Set up of Excel Workbook
        input_workbook = xlrd.open_workbook(database_path)
        self.input_worksheet = input_workbook.sheet_by_name(selected_machine_type)

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

        """Following method determines properties of CAD model according to sheet."""

        # Initial GUI operations
        self.list_cad_models.disable_this_dropdown_menu()
        self.confirm_selected_cad.disable_this_button()

        # For to me unknown reason application did not enabled button with run_button.enable_this_button (even though it works with confirmation button)
        # There for instance of this button is re-created.
        self.run_button = CreateControlButton(parent=self, row_grid=13, column_grid=0, command=self.close_graphical_user_interface, icon_name="run_automation.png")
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

        print(properties)
        # Removing non relevant properties is going to happen with lambda following function
        filtering_properties = filter(lambda x: isinstance(x['value'], list) and len(x['value']) > 1 or isinstance(x['value'], str), properties)
        properties = list(filtering_properties)
        print(list(filtering_properties))
        for each in properties:
            print(each)

        self.create_dropdown_menu_for_properties(self, properties)

    def create_dropdown_menu_for_properties(self, parent, list_of_properties):

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
        self.close_graphical_user_interface()
        main()

    def close_graphical_user_interface(self):
        """Button reaction function"""
        self.master.destroy()

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
        self.icon_obj = ImageTk.PhotoImage(Image.open(icons_folder_path + icon_name))
        self.parent.button_obj = tkinter.Button(image=self.icon_obj, command=command)
        self.parent.button_obj.grid(row=row_grid, column=column_grid)

    def set_rowspan_equals_2(self):
        self.parent.button_obj.grid(rowspan=2)

    def disable_this_button(self):
        self.parent.button_obj.config(state='disabled')

    def enable_this_button(self):
        self.parent.button_obj.config(state='normal')


class CreateDropDownMenu:

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


class FoldersOperations:
    pass


class CreoAPI:

    # Model names stick to lowercase convention - incorrect CX_0420_ZE77-0XX_GM1811.asm  - correct cx_420_ze77-0xx_gm1811.asm
    # List of functions
    # connect, create bom information (check whether resumed models are innit), check skeleton exists,
    # take care of flow control
    # possible names - use quicker functions - may slow down
    def __init__(self):
        self.creo_client = creopyson.Client()
        self.setup()
        self.bill_of_material = []
        self.zs_63 = Zs63(self.current_master_model())
        self.paired_bill_of_material = []
        self.resume_all_groups(order_number_only=True)

        print("Creo API has been initialized.")

    def setup(self):
        try:
            self.creo_client.connect()
            logger.info('Creoson is running')
        except ConnectionError:
            creoson_folder = '.\creoson'
            try:
                os.startfile(creoson_folder)
            except:
                logger.warning("Creoson folder is not found in app folder.")
            finally:
                messagebox.showinfo("Kraussmaffei Assembly Automation", "Creoson is not running. Start Creoson before starting Automation app.")
                logger.critical("Creoson is not running. Start Creoson before starting Automation app.")
                exit()

    def current_master_model(self):
        return creopyson.file_get_fileinfo(self.creo_client)['file']

    def current_order_number(self):
        return self.current_master_model()[-10:-4]

    def create_master_model_bill_of_material_with_suppressed(self, levels):

        """This method uses list-features method to create structured Bill of Material"""

        current_master_model = self.current_master_model()
        current_order_number = self.current_order_number()

        # At first we list top level groups
        self.add_models_in_opened_group_to_bom(level_of_master_model_tree=1)
        first_level_groups = tuple(self.bill_of_material)

        if levels > 1:
            for each_dict in first_level_groups:
                creopyson.file_open(self.creo_client, file_=each_dict['name'])
                self.resume_all_groups(order_number_only=True)
                self.add_models_in_opened_group_to_bom(level_of_master_model_tree=2)
        second_level_groups = tuple(self.bill_of_material)

        if levels > 2:
            # And third Level groups
            # resuming second level groups (those are groups with order number)
            for each_dict in second_level_groups:
                if each_dict['level_of_master_model_tree'] == 2 and '.prt' not in each_dict['name'] and current_order_number in each_dict['name']:
                    creopyson.file_open(self.creo_client, file_=each_dict['name'])
                    self.add_models_in_opened_group_to_bom(level_of_master_model_tree=3)

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
        self.bill_of_material = d = [i for n, i in enumerate(self.bill_of_material) if i not in self.bill_of_material[n + 1:]]
        return bom_opened_group

    def create_master_model_bill_of_material_no_suppressed(self):

        """This method uses """

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

    def bom_recursion(self, nest_dict, list_of_recursion_bom=[]):

        for key, value in nest_dict.items():
            if isinstance(value, dict):
                self.bom_recursion(value)
            elif isinstance(value, list):
                for each in value:
                    if isinstance(each, dict):
                        self.bom_recursion(each)
            else:
                list_of_recursed_bom = list_of_recursion_bom
                if key != 'generic':
                    list_of_recursed_bom.append(("{0} : {1}".format(key, value)))

        return list_of_recursion_bom

    def resume_all_groups(self, order_number_only=True):

        current_model = self.current_master_model()
        if order_number_only:
            order_number_only = self.current_order_number()
        else:
            order_number_only = ""
        id_list = []
        resume_all_mapkey = "~ Command `ProCmdMdlTreeSearch`;\
                                 ~ Activate `selspecdlg0` `SelScopeCheck` 0;\
                                 ~ Input `selspecdlg0` `SelOptionRadio` `Component`;\
                                 ~ Update `selspecdlg0` `SelOptionRadio` `Component`;\
                                 ~ Select `selspecdlg0` `CascadeButton1`;\
                                 ~ Close `selspecdlg0` `CascadeButton1`;\
                                 ~ Activate `selspecdlg0` `Suppressed` 1;\
                                 ~ Select `selspecdlg0` `RuleTab` 1 `Misc`;\
                                 ~ Update `selspecdlg0` `ExtRulesLayout.ExtBasicIDLayout.InputIDPanel` `ID_COMPONENT`;\
                                 ~ Activate `selspecdlg0` `EvaluateBtn`;\
                                 ~ Activate `selspecdlg0` `EvaluateBtn`;\
                                 ~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;\
                                 ~ Activate `selspecdlg0` `CancelButton`;\
                                 ~ Close `selspecdlg0` `selspecdlg0`;\
                                 ~ Command `ProCmdResume@PopupMenuTree`;\
                                ~ Activate `storage_conflicts` `OK_PushButton`;"

        self.allow_conflicts()

        top_assemblies_raw = creopyson.feature_list(self.creo_client, name='*' + order_number_only + '*', no_datum=True, type_='COMPONENT')
        for every_main_group_name in top_assemblies_raw:
            if every_main_group_name["status"] == "SUPPRESSED":
                id_list.append(every_main_group_name["feat_id"])

        all_id_resumed = False

        while not all_id_resumed:
            for each_id in id_list:
                resume_id = resume_all_mapkey.replace("ID_COMPONENT", str(each_id))
                creopyson.interface_mapkey(self.creo_client, resume_id)
                #if self.bill_of_material:
                #    self.change_parameter_in_bill_of_material(key='status', new_value='ACTIVE', feat_id=each_id)
            #time.sleep(0.1)
            components_list = creopyson.feature_list(self.creo_client, name='*' + order_number_only + '*', no_datum=True, type_='COMPONENT')
            testing_suppressed_comp = []
            for each_comp in components_list:
                if each_comp["status"] == "SUPPRESSED":
                    self.change_parameter_in_bill_of_material(key='status', new_value='ACTIVE', feat_id=each_id)
                    testing_suppressed_comp.append(every_main_group_name["feat_id"])

            if testing_suppressed_comp == []:
                logger.info(f'All assemblies are resumed in {current_model}.')
                all_id_resumed = True
            else:
                time.sleep(0.1)

    def allow_conflicts(self):

        """This script is handy when it comes to giving program permissions to solve conflict."""

        old_script = "~ Command `ProCmdDatumCsys` ;~ Activate `storage_conflicts` `OK_PushButton`;~ Close `Odui_Dlg_00` `Odui_Dlg_00`;"
        script = "~ Activate `storage_conflicts` `OK_PushButton`;~ Close `Odui_Dlg_00` `Odui_Dlg_00`;"
        creopyson.interface_mapkey(self.creo_client, script)

    def change_parameter_in_bill_of_material(self, key, new_value, feat_id):

        """This method changes parameter in bill of material - it is necessary to track all  """
        try:
            change_dict = next(item for item in self.bill_of_material if item['feat_id'] == feat_id)
        except StopIteration:
            logger.warning(f'{feat_id} is not in master model.')
        else:
            index_of_change_dict = self.bill_of_material.index(change_dict)
            new_dict = {key: new_value}
            if isinstance(change_dict, dict):
                self.bill_of_material[index_of_change_dict].update(new_dict)
                print(f'BOM - parameter {key} has been changed to {new_value}.')
                logger.info(f'BOM - parameter {key} has been changed to {new_value}.')
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

        top_assemblies = []

        self.allow_conflicts()
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
            print(dimensions_to_remove)
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
            print(dimensions_to_remove)
            self.determine_whether_group_consist_powerpack(list_of_powerpacks=list_second_powerpacks, picked_powerpack=picked_option)


        if any(dictionary['property'] == 'list_primary_plast' for dictionary in properties):
            list_primary_plast = next(item for item in properties if item['property'] == 'list_primary_plast')['value']
            picked_dictionary = next(item for item in app.list_of_option_properties if item['property'] == 'list_primary_plast')['drop_down_obj']
            picked_option = picked_dictionary.what_is_picked_option()
            print(f'Picked primary plast is {picked_option}.')
            dimensions_to_remove = picked_dictionary.return_non_picked_values()
            print(dimensions_to_remove)
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

        """This method adds new model property to dictionary"""
        for bom_dict in self.bill_of_material:
            if 'group_type' not in bom_dict:
                bom_dict['group_type'] = 'not defined'
                if 'm' in bom_dict['name'][:-10] and '.asm' in bom_dict['name']:
                    bom_dict['group_type'] = 'm'
                elif 'ze' in bom_dict['name'][:-10] and '.asm' in bom_dict['name']:
                    bom_dict['group_type'] = 'ze'
                elif 'c' in bom_dict['name'][:-10] and '.asm' in bom_dict['name']:
                    try:
                        number_position = (bom_dict['name'].index('c')+1)
                        int(bom_dict['name'][number_position])
                    except ValueError:
                        print('is not integer')
                    else:
                        bom_dict['group_type'] = 'c'
                elif 'sa' in bom_dict['name'][:-10] and '.asm' in bom_dict['name']:
                    bom_dict['group_type'] = 'sa'

    def rename_config_control(self, boolean):
        if boolean == "yes":
            creopyson.creo_set_config(client=self.creo_client, name="let_proe_rename_pdm_objects", value="yes")
        elif boolean == "no":
            creopyson.creo_set_config(client=self.creo_client, name="let_proe_rename_pdm_objects", value="no")

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
                    print(each_dict)
                if creopyson.file_exists(self.creo_client, file_=try_model):
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
        print('================Powerpack Handling============================')
        print('================Powerpack Handling============================')

        # Performance part
        for each_dict in groups_to_remove:
            if each_dict['test_result'] == 'remove':
                self.try_delete_model(model_to_delete=each_dict['name'])
            if each_dict['test_result'] == 'rename_to_correct':
                print(each_dict)
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

        """Name of this method is very clear. Model is opened and closed by itself. Creo session returns to previous model."""
        # TODO: Refactor this method

        get_session_info = creopyson.file_get_fileinfo(self.creo_client)
        current_model = (get_session_info['file'])
        creopyson.file_open(self.creo_client, file_=component_name)

        self.allow_conflicts()
        self.resume_all_groups(order_number_only=True)
        self.add_models_in_opened_group_to_bom()

        filter_children_only = filter(lambda x: x['parent'] == component_name, self.bill_of_material)
        filter_children_only_list = list(filter_children_only)

        parent_dict = next(x for x in self.bill_of_material if x['name'] == component_name)
        filter_children_only_list.append(parent_dict)

        self.rename_config_control("yes")
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

        child_model = self.check_whether_model_exists(erp_material_number)

        if child_model:
            # Now we test whether model already exists in cad_parent_model. If yes we skip this step to avoid cad model duplicity:
            # TODO: resume which stayed, ADD update to bill_of_material.

            if not self.check_whether_model_name_is_in_assembly(parent=cad_parent_model, model_name_wild_card=erp_material_number):
                child_csy = self.pick_csy(cad_parent_model, child_model)
                parent_csy = self.pick_csy(cad_parent_model, cad_parent_model)

                if parent_csy != 'CSY does not exist' and child_csy != 'CSY does not exist':
                    creopyson.file_assemble(self.creo_client, into_asm=cad_parent_model, file_=child_model, constraints=[{"asmref": parent_csy, "compref": child_csy, "type": "csys"}])

                if parent_csy == 'CSY does not exist' and child_csy != 'CSY does not exist':
                    # TODO: Check whether skeleton exists method might be enhanced.
                    skeleton_information = self.check_whether_skeleton_exists(cad_parent_model)
                    if skeleton_information['csy'] != 'CSY does not exist':
                        creopyson.file_assemble(self.creo_client, file_=child_model, into_asm=cad_parent_model, ref_model=skeleton_information['skel_name'],
                        constraints=[{"asmref": skeleton_information['csy'], "compref": child_csy, "type": "csys"}])
                    else:
                        creopyson.file_assemble(self.creo_client, into_asm=cad_parent_model, file_=child_model, constraints=[{"type": "fix"}], package_assembly=True)
                else:
                    creopyson.file_assemble(self.creo_client, into_asm=cad_parent_model, file_=child_model, constraints=[{"type": "fix"}], package_assembly=True)

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
                        print(current_mx_csy_object)
                        print("Above is csy object")
                    if len(mx_csys_objects) > 1:
                        for each_csy_object in mx_csys_objects:
                            if zs_63_injection_unit >= each_csy_object["floor"] and zs_63_injection_unit <= each_csy_object["ceil"]:
                                picked_csy = each_csy_object["name"]
                                logger.info("MX variation of picked CSY is " + picked_csy)
            except:
                logger.exception("message")
            return picked_csy

        return picked_csy

    def check_whether_model_exists(self, erp_material_number):
        """This method tests whether ERP material number exists in Windchill and if exists it will assign its modelname to model_name variable"""
        model_name = ''
        erp_material_number = erp_material_number.replace('.prt', '')
        erp_material_number = erp_material_number.replace('.asm', '')
        if creopyson.file_exists(self.creo_client, file_=erp_material_number + '.prt'):
            model_name = erp_material_number + '.prt'
            print(f'Yes material number exists ! Model name is {model_name}.')
            logger.info(f'Yes material number exists ! Model name is {model_name}.')
        if creopyson.file_exists(self.creo_client, file_=erp_material_number + '.asm'):
            model_name = erp_material_number + '.asm'
            print(f'Yes material number exists ! Model name is {model_name}.')
            logger.info(f'Yes material number exists ! Model name is {model_name}.')
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

        skeleton_information = {}
        list_with_skeletons = []
        list_without_skeletons = []

        x = creopyson.bom.get_paths(self.creo_client, skeletons=True, top_level=True)
        x = x["children"]
        x = x["children"]
        for every_dict in x:
            list_with_skeletons.append(every_dict["file"])
        y = creopyson.bom.get_paths(self.creo_client, skeletons=False, top_level=True)
        y = y["children"]
        y = y["children"]
        for every_dict in y:
            list_without_skeletons.append(every_dict["file"])
        for every_file in list_without_skeletons:
            list_with_skeletons.remove(every_file)
        if len(list_with_skeletons) == 0:
            messagebox.showerror('Automation status', 'Automation aborted! There is not skeleton model in mastermodel.')
            exit()
        elif len(list_with_skeletons) >= 1:
            cad_skeleton_name = list_with_skeletons[0]
            skeleton_information['skel_name'] = cad_skeleton_name
        skeleton_information['csy'] = self.pick_csy(cad_parent_model, cad_skeleton_name)
        print(skeleton_information)

        return skeleton_information

    def set_model_convention_on_the_fly(self):
        """This method aligns model naming, due to what users can expect higher percentage of successfully added models."""
        # TODO : Could be refactored but low priority
        gmxxxx = self.current_order_number()
        list_of_numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 0]
        for every_component in self.bill_of_material:
            for every_number in list_of_numbers:
                if "_" + str(every_number) + "_" + gmxxxx in every_component:
                    self.rename_config_control("yes")
                    new_name = every_component.replace("_" + str(every_number) + "_" + gmxxxx, "-" + str(every_number) + "_" + gmxxxx)
                    try:
                        creopyson.file_rename(self.creo_client, file_=every_component, new_name=new_name, onlysession=True)
                        self.try_remove_from_ws(every_component)
                        logger.info('Repaired convention in model = ' + new_name)
                    except:
                        pass
                    self.rename_config_control("no")

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

        """This function creates new copy of master model."""

        new_number = order_number_entry.get()
        print(new_number)
        order_number = self.current_order_number()
        self.rename_config_control("yes")
        for every_component in self.bill_of_material:
            if order_number in every_component['name']:
                new_name = every_component['name'].replace(order_number, new_number)
                try:
                    creopyson.file_rename(self.creo_client, file_=every_component['name'], new_name=new_name, onlysession=True)
                    print('this is new model = ' + new_name)
                    logger.info('this is new model = ' + new_name)
                    # creopyson.file_save(creo_client, file_=new_name)
                except:
                    logger.exception("message")
                else:
                    self.change_parameter_in_bill_of_material(key='name', new_value=new_name, feat_id=every_component['feat_id'])
        self.rename_config_control("no")

    def run_with_order_number(self):

        creopyson.file_open(self.creo_client, )

    def zs_63_pairing(self):

        self.zs_63.transform_zs_63()
        self.paired_bill_of_material = self.zs_63.pair_converted_zs_63_with_cad_master_model(self.bill_of_material)

    def get_zs63_file(self):
        pass

    def remove_unnecessary_material_numbers(self):

        folder_name = self.current_master_model()[0:3].upper().replace('_','')
        delete_exclude_file = delete_exclude_path + folder_name + '\\DeleteExclude.csv'
        order_number = self.current_order_number()

        with open(delete_exclude_file, 'r') as csv_file:
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
                    print(f"This group is avoided {each_dict['parent']}, - Is in delete exclude file")
                    logger.info(f"This group is avoided {each_dict['parent']}, - Is in delete exclude file.")
                else:
                    # Model will not be deleted
                    # I any cokolvek z list_groups in each_dict['name']
                    if any(ele['ERP_number'] in each_dict['name'] for ele in self.paired_bill_of_material):
                        print(f"This group is avoided {each_dict['name']} in {each_dict['parent']}. - In zs_63")
                        logger.info(f"This group is avoided {each_dict['name']} in {each_dict['parent']}. - In zs_63")
                    else:
                        creopyson.file_open(self.creo_client, file_=each_dict['parent'])
                        creopyson.feature_delete(self.creo_client, name=each_dict['name'])
                        self.bill_of_material.remove(each_dict)
                        print(f"This group is remove {each_dict['name']} from {each_dict['parent']}. - In zs_63")
                        logger.info(f"This group is remove {each_dict['name']} from {each_dict['parent']}. - In zs_63")


class Zs63:
    """This class refers to text file from zs63 SAP transaction. Loading of this function has to be enhanced with SAP scripting."""
    def __init__(self, current_master_model):
        self.final_folder_path = os.path.dirname(sys.argv[0]) + '\\ErpBom\\ZS_63.txt'
        self.all_lists = []
        self.m_groups_list = []
        self.ze_groups_list = []
        self.sa_groups_list = []
        self.current_master_model = current_master_model
        # Methods on innit
        self.get_zs63_file()
        self.transform_zs_63()

    def get_zs63_file(self):
        """This method loads zs63 file to application"""
        remove_files_from_folder(erp_folder_path)
        remove_files_from_folder(feedback_folder_path)
        sap_source = tkinter.Tk()
        sap_source.filename = filedialog.askopenfilename(initialdir='\\', title='Choose ZS63 file', filetypes=(('text files', '*.txt'), ('all files', '*.*')))
        sap_source.destroy()
        current_folder_path = sap_source.filename

        try:
            shutil.copyfile(current_folder_path, self.final_folder_path)
        except FileExistsError:
            os.replace(current_folder_path, self.final_folder_path)
        except FileNotFoundError:
            pass
            # TODO: define what should happen when user does not select text file.
            #close_graphical_user_interface()

    def transform_zs_63(self):

        """"This method transforms zs63 file to 3 lists (m_groups, ze_groups, sa_groups). Later these groups are merged to all_lists (list type)."""
        m_groups_list = []
        ze_groups_list = []
        sa_groups_list = []
        zs_63 = []

        with open(self.final_folder_path) as zs_data:

            zs_63_raw = zs_data.readlines()

            # This is encoding part - necessary to implement because of various operating systems - CHINA os issues
            for each_line in zs_63_raw:
                each_line.encode("utf-8", "ignore")
                line_text = str(each_line.encode("utf-8", "ignore"))
                line_text = line_text.replace("b'", "")
                zs_63.append(str(line_text))

            # Newly create section where some risky signs will be removed from zs_63.txt
            if type(zs_63) == list:
                for each_element in zs_63:
                    try:
                        if "°C" in each_element:
                            zs_63.remove(each_element)
                            print(each_element + " is removed as banned symbol.")
                    except:
                        logger.info("banned symbol was not possible to remove ")

                for each_element in zs_63:
                    try:
                        if "?" in each_element:
                            zs_63[zs_63.index(each_element)] = each_element.replace("?", " ")
                            print(each_element + " is removed as banned symbol ?")
                    except:
                        pass
                for each_element in zs_63:
                    try:
                        if "#" in each_element and each_element[2] != "#":
                            zs_63[zs_63.index(each_element)] = each_element.replace("#", " ")
                            print(each_element + " is removed as banned symbol #")
                    except:
                        pass

            # Here we try to set injection unit size for MX
            try:
                global zs_63_injection_unit;
                zs_63_injection_unit = ""
                if "mx" in self.current_master_model:
                    following_line = False
                    for line in zs_63:
                        if following_line == True:
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
                    if len(each_split_element) > 6 and len(each_split_element) < 9 and '7' not in each_split_element[0:1] and each_split_element.isnumeric():
                        pair_group_and_number.append(each_split_element)
                    if 'M' in each_split_element and '.' in each_split_element or 'C' in each_split_element and '.' in each_split_element:
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
                    if len(each_split_element) > 6 and len(each_split_element) < 9 and '7' not in each_split_element[0:1] and each_split_element.isnumeric():
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

            for every_element in sa_groups_list:
                split_element = every_element.split()
                pair_group_and_number = []
                for each_split_element in split_element:
                    each_split_element = each_split_element.strip()
                    if len(each_split_element) < 4 and len(each_split_element) > 1 and each_split_element.isnumeric():
                        SAValue = int(each_split_element)
                        if SAValue > 29 and SAValue < 300:
                            if len(each_split_element) == 2:
                                each_split_element = 'SA0' + each_split_element
                            if len(each_split_element) == 3:
                                each_split_element = 'SA' + each_split_element
                            pair_group_and_number.append(each_split_element)
                    if len(each_split_element) > 6 and len(each_split_element) < 9 and '7' not in each_split_element[0:1] and each_split_element.isnumeric():
                        pair_group_and_number.append(each_split_element)
                if len(pair_group_and_number) == 2:
                    mat_nr = pair_group_and_number[1]
                    sap_group = pair_group_and_number[0]
                    group_mat_nr_dict = {'SAP_group_name': sap_group, 'ERP_number': mat_nr, 'CAD_group_name': 'Not defined'}
                    self.all_lists.append(group_mat_nr_dict.copy())

    def pair_converted_zs_63_with_cad_master_model(self, bill_of_material):

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
            print(each_dict)
            if each_dict["CAD_group_name"] != "Not defined":
                successful_pairing = successful_pairing + 1

        percentage = successful_pairing / len(self.all_lists) * 100
        logger.info(f"Percentage of defined pairs is {str(percentage)} %.")
        print(f"Percentage of defined pairs is {str(percentage)} %.")
        logger.info('end of pairing')

        return self.all_lists


def development_function():

    """"Function created for binding of testing functions. Initiated from ZS63 button."""
    start_time = time.time()
    session = CreoAPI()

    creopyson.file_open(session.creo_client, file_=app.list_cad_models.what_is_picked_option())
    session.create_master_model_bill_of_material_with_suppressed(levels=1)
    session.filter_assemblies()
    #session.set_model_convention_on_the_fly()
    #session.change_order_number()

    session.clear_bill_of_material()
    session.create_master_model_bill_of_material_with_suppressed(levels=3)

    session.zs_63_pairing()
    session.remove_unnecessary_material_numbers()

    message = "Creation of BOM took program " + ("--- %s seconds ---" % (time.time() - start_time))
    logger.info(message)
    print(message)


def only_numerics(seq):
    seq_type = type(seq)
    return seq_type().join(filter(seq_type.isdigit, seq))


def remove_files_from_folder(folder_name):

    set_global_paths()

    for f in os.walk(folder_name):
        for fileX in f:
            print(fileX)
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



def main():
    """ Main function  - Graphical user interface is initialized by this function.
    Necessary to wrap into function because of its usage in Reset button command."""
    global app
    root = tkinter.Tk()
    root.lift()
    app = Application(master=root)
    app.mainloop()


if __name__ == "__main__":
    main()
