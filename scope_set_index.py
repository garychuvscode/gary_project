# this file is used to save the scope setting for different testing items


class scope_config():

    def __init__(self):

        pass

    def setting_mapping(self, setup_index):

        if setup_index == 'ripple_50374-2':
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3,
                          'label_name': 'AVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.3,
                          'label_name': 'OVSS', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5,
                          'label_name': 'VON', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '1', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50'}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}

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

        elif setup_index == 'SY8386C_ripple':
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'Buck', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M'}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0.045,
                          'label_name': 'LDO', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M'}
            self.ch_c3 = {'ch_view': 'FALSE', 'volt_dev': '1', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '5', 'BW': '20MHz', 'filter': 'None', 'v_offset': -15,
                          'label_name': 'Vin', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50'}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': 'None', 'v_offset': -0.04,
                          'label_name': 'VCC', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'AC1M'}
            self.ch_c7 = {'ch_view': 'FALSE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': 'None', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': 'None', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C1', 'trigger_level': '0',
                                'trigger_slope': 'Positive', 'time_scale': '0.0005', 'time_offset': '0', 'sample_mode': 'RealTime', 'fixed_sample_rate': '5GS/s'}
            # note: ripple measurement follow robert's measurement

            # setting of measurementsinjfij
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

        else:
            # if index wrong, back to 374 settings
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3,
                          'label_name': 'AVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.3,
                          'label_name': 'OVSS', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 3.5,
                          'label_name': 'VON', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '1', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3,
                          'label_name': 'Vin', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -0.06,
                          'label_name': 'I_load', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC50'}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.02', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.3,
                          'label_name': 'OVDD', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '0.2', 'BW': '20MHz', 'filter': '2bits', 'v_offset': -3.5,
                          'label_name': 'VOP', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c8 = {'ch_view': 'FALSE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}

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
