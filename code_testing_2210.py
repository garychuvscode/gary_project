# this file is mainly for the testing of coding


class test_calass():

    def __init__(self):

        pass

    def scope_ch(self):

        self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': 20, 'filter': 2, 'v_offset': 0,
                      'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
        self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': 20, 'filter': 2, 'v_offset': 0,
                      'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}

        for i in range(1, 8+1):
            if i == 1:
                temp_dict = self.ch_c1

            a = temp_dict['ch_view']
            print(f'app.Acquisition.C{i}.View = {a}')
            print(f'app.Acquisition.C{i}.View = {temp_dict["ch_view"]}')

        pass


t_s = test_calass()
testing_index = 0

if testing_index == 0:
    print('a')
    # from 0-9 => < 10 and start from 0 (like array)
    for i in range(10):
        print(i, end=' ')
        # not change line

    for i in range(1, 1+8):
        print(i)

        t_s.scope_ch()
