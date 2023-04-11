# this file is used to save the scope setting for different testing items


class scope_config():

    def __init__(self):

        pass

    def setting_mapping(self, setup_index):
        '''
        scope setting input from here...
        '''

        if setup_index == 'ripple_50374-2':
            # index 0 only for functional test
            # the setting for ripple verification

            '''
            221205
            to improve the find signal, add two dimension dictionary, and change the offset setting to
            nornalization result, x * volt_div, from +3 to -3
            use the 'v_offset_ind' as the new offset normalization index
            '''

            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3 + 0.04,
                          'label_name': 'AVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 3}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.3 - 0.06,
                          'label_name': 'OVSS', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5 - 0.4,
                          'label_name': 'VON', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2.5}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '1', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0.0002, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3.5}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3 + 0.02,
                          'label_name': 'OVDD', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -4.3,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C3', 'trigger_level': '-3.2',
                                'trigger_slope': 'Positive', 'time_scale': '0.0001',
                                'time_offset': '-0.0004', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'SY8386C_ripple':
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'Buck', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 0}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0.045,
                          'label_name': 'VCC', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -2.5}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '1', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '5', 'BW': '20MHz', 'filter': 'None', 'v_offset': -15,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -2.5}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': -0.04,
                          'label_name': 'LDO', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 0}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C1', 'trigger_level': '0',
                                'trigger_slope': 'Positive', 'time_scale': '0.0005', 'time_offset': '0', 'sample_mode': 'RealTime', 'fixed_sample_rate': '5GS/s'}
            # note: ripple measurement follow robert's measurement

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'SY8386C_load_tran':
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'Buck', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 1}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0.045,
                          'label_name': 'VCC', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -2.5}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '1', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '5', 'BW': '20MHz', 'filter': 'None', 'v_offset': -15,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': -0.04,
                          'label_name': 'LDO', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C5', 'trigger_level': '3',
                                'trigger_slope': 'Positive', 'time_scale': '0.0001', 'time_offset': '-0.00025', 'sample_mode': 'RealTime', 'fixed_sample_rate': '5GS/s'}
            # note: ripple measurement follow robert's measurement

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'SY8386C_load_tran_LDO':
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'Buck', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 1}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0.045,
                          'label_name': 'VCC', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -2.5}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '1', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '5', 'BW': '20MHz', 'filter': 'None', 'v_offset': -15,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.1', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': -0.04,
                          'label_name': 'LDO', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C5', 'trigger_level': '0.025',
                                'trigger_slope': 'Positive', 'time_scale': '0.0001', 'time_offset': '-0.00025', 'sample_mode': 'RealTime', 'fixed_sample_rate': '5GS/s'}
            # note: ripple measurement follow robert's measurement

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'SY8386C_line_tran':
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.05', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'Buck', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -0.5}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.1', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0.045,
                          'label_name': 'VCC', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -2.5}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '1', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '5', 'BW': '20MHz', 'filter': 'None', 'v_offset': -15,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1.75}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.1', 'BW': '20MHz', 'filter': 'None', 'v_offset': -0.04,
                          'label_name': 'LDO', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -1.5}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C4', 'trigger_level': '15',
                                'trigger_slope': 'Positive', 'time_scale': '0.00005', 'time_offset': '-0.000125', 'sample_mode': 'RealTime', 'fixed_sample_rate': '5GS/s'}
            # note: ripple measurement follow robert's measurement

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'SY8386C_line_tran_LDO':
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.05', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'Buck', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -0.5}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.1', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0.045,
                          'label_name': 'VCC', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -2.5}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '1', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '5', 'BW': '20MHz', 'filter': 'None', 'v_offset': -15,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1.75}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.1', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.1', 'BW': '20MHz', 'filter': 'None', 'v_offset': -0.04,
                          'label_name': 'LDO', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -1.5}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C4', 'trigger_level': '15',
                                'trigger_slope': 'Positive', 'time_scale': '0.00005', 'time_offset': '-0.000125', 'sample_mode': 'RealTime', 'fixed_sample_rate': '5GS/s'}
            # note: ripple measurement follow robert's measurement

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'BK_pwr_seq_inrush':
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'Buck', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -2}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0.045,
                          'label_name': 'VCC', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 1}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '1', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '5', 'BW': '20MHz', 'filter': 'None', 'v_offset': -15,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2.5}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': 1}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -0.04,
                          'label_name': 'LDO', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C1', 'trigger_level': '0',
                                'trigger_slope': 'Positive', 'time_scale': '0.0005', 'time_offset': '0', 'sample_mode': 'RealTime', 'fixed_sample_rate': '5GS/s'}
            # note: ripple measurement follow robert's measurement

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == '50374_line_tran':

            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3 + 0.04,
                          'label_name': 'AVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 0}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.3 - 0.06,
                          'label_name': 'OVSS', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5 - 0.4,
                          'label_name': 'VON', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 2.5}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3.5}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3 + 0.02,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -4.3,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C4', 'trigger_level': '3.3',
                                'trigger_slope': 'Positive', 'time_scale': '0.00005',
                                'time_offset': '-0.000125', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == '50374_load_tran':

            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3 + 0.04,
                          'label_name': 'AVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 0}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.3 - 0.06,
                          'label_name': 'OVSS', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5 - 0.4,
                          'label_name': 'VON', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -1}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3.5}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3 + 0.02,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -4.3,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C5', 'trigger_level': '0.005',
                                'trigger_slope': 'Positive', 'time_scale': '0.00005',
                                'time_offset': '-0.000125', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'RT4733_ripple':

            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '3bits', 'v_offset': -3.3 + 0.04,
                          'label_name': 'AVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 3}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '3bits', 'v_offset': 3.3 - 0.06,
                          'label_name': 'OVSS', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5 - 0.4,
                          'label_name': 'VON', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2.5}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '1', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.05', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0.0002, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3.5}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '3bits', 'v_offset': -3.3 + 0.02,
                          'label_name': 'OVDD', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -4.3,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C3', 'trigger_level': '-3.2',
                                'trigger_slope': 'Positive', 'time_scale': '0.0001',
                                'time_offset': '-0.0004', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'RT4733_line_tran':

            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '3bits', 'v_offset': -3.3 + 0.04,
                          'label_name': 'AVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 3}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '3bits', 'v_offset': 3.3 - 0.06,
                          'label_name': 'OVSS', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5 - 0.4,
                          'label_name': 'VON', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2.5}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 2.5}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.05', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0.0002, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3.5}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '3bits', 'v_offset': -3.3 + 0.02,
                          'label_name': 'OVDD', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -4.3,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C4', 'trigger_level': '3.15',
                                'trigger_slope': 'Positive', 'time_scale': '0.00005',
                                'time_offset': '-0.000125', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'RT4733_load_tran':

            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '3bits', 'v_offset': -3.3 + 0.04,
                          'label_name': 'AVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 3}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '3bits', 'v_offset': 3.3 - 0.06,
                          'label_name': 'OVSS', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5 - 0.4,
                          'label_name': 'VON', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2.5}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 2.5}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.03', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0.0002, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3.5}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '3bits', 'v_offset': -3.3 + 0.02,
                          'label_name': 'OVDD', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -4.3,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C5', 'trigger_level': '0.005',
                                'trigger_slope': 'Positive', 'time_scale': '0.00005',
                                'time_offset': '-0.000125', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'ripple_50374-2AC':
            # index 0 only for functional test
            # the setting for ripple verification

            '''
            221205
            to improve the find signal, add two dimension dictionary, and change the offset setting to
            nornalization result, x * volt_div, from +3 to -3
            use the 'v_offset_ind' as the new offset normalization index
            '''

            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3 + 0.04,
                          'label_name': 'AVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 3}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.3 - 0.06,
                          'label_name': 'OVSS', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -1}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5 - 0.4,
                          'label_name': 'VON', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -2.5}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '1', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -3}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3.5}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3 + 0.02,
                          'label_name': 'OVDD', 'label_position': 0.0001, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -4.3,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M', 'v_offset_ind': -3}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C3', 'trigger_level': '-3.2',
                                'trigger_slope': 'Positive', 'time_scale': '0.0001',
                                'time_offset': '-0.0004', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'NT50374_pwr_seq_inrush':
            # default set for inrush checking, C4 use Vin, change in pwr_sub program
            '''
            221205
            to improve the find signal, add two dimension dictionary, and change the offset setting to
            nornalization result, x * volt_div, from +3 to -3
            use the 'v_offset_ind' as the new offset normalization index
            '''

            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 2,
                          'label_name': 'AVDD', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -2,
                          'label_name': 'OVSS', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -1}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -1,
                          'label_name': 'VON', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -0.5}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '3bits', 'v_offset': -6,
                          'label_name': 'EN_pin', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -1.5}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.6,
                          'label_name': 'Iin', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'OVDD', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 0}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'VOP', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c8 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '3bits', 'v_offset': -7,
                          'label_name': 'SW_pin', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C4', 'trigger_level': '1.8',
                                'trigger_slope': 'Positive', 'time_scale': '0.001',
                                'time_offset': '-0.004', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "max", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        elif setup_index == 'NT50970_pwr_seq_inrush':
            # default set for inrush checking, C4 use Vin, change in pwr_sub program
            '''
            221205
            to improve the find signal, add two dimension dictionary, and change the offset setting to
            nornalization result, x * volt_div, from +3 to -3
            use the 'v_offset_ind' as the new offset normalization index
            '''

            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'Buck', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 0}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -2,
                          'label_name': 'VCC', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -12.5,
                          'label_name': 'Vin', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2.5}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '3bits', 'v_offset': 4,
                          'label_name': 'EN2_pin', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 3.5}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -1.5,
                          'label_name': 'Iin', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': -3}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -4,
                          'label_name': 'LDO', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -1}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -7.5,
                          'label_name': 'PG', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': -2}
            self.ch_c8 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': '3bits', 'v_offset': 3,
                          'label_name': 'EN1_pin', 'label_position': 0.001, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 3}

            # add the two dimension index for the find signal reference
            self.ch_index = {'C1': self.ch_c1, 'C2': self.ch_c2, 'C3': self.ch_c3, 'C4': self.ch_c4,
                             'C5': self.ch_c5, 'C6': self.ch_c6, 'C7': self.ch_c7, 'C8': self.ch_c8}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C4', 'trigger_level': '1.8',
                                'trigger_slope': 'Positive', 'time_scale': '0.0002',
                                'time_offset': '-0.0008', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "max", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

            pass

        else:
            # if index wrong, back to 374 settings
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3,
                          'label_name': 'CH1', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.3,
                          'label_name': 'CH2', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5,
                          'label_name': 'CH3', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '1', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50', 'v_offset_ind': 1}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M', 'v_offset_ind': 1}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C3', 'trigger_level': '-3.2',
                                'trigger_slope': 'Positive', 'time_scale': '0.0001', 'time_offset': '-0.0004', 'sample_mode': 'RealTime', 'fixed_sample_rate': '1.25GS/s'}

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

        pass


if __name__ == '__main__':
    #  the testing code for this file object
    import parameter_load_obj as par
    excel_t = par.excel_parameter('obj_main')
    default_path = 'C:\\wave_form_raw\\'

    import Scope_LE6100A as sco

    # scope = Scope_LE6100A('GPIB: 5', 3, sim_scope, excel_t)
    scope = sco.Scope_LE6100A(excel0=excel_t)

    test_index = 1

    if test_index == 0:
        # used for checking the scope initialization setting
        scope.open_inst()
        # use the index correction or not
        scope.nor_v_off = 1
        scope.scope_initial('SY8386C_line_tran')
        pass

    elif test_index == 1:
        # used for checking the scope initialization setting
        scope.open_inst()
        excel_t.wave_path = default_path
        excel_t.wave_condition = 'scope_set_temp_capture'
        scope.printScreenToPC()

        pass
