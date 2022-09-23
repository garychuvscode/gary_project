import main_obj as g

g.excel_m.open_result_book()
g.sim_mode_all(1)
g.sim_mode_independent(1, 1, 1, 1, 1, 0, 1)
g.open_inst_and_name()
g.iq_test.run_verification()
g.excel_m.end_of_file(0)
g.excel_m.open_result_book()
g.sw_test.run_verification()
g.excel_m.end_of_file(0)
g.excel_m.open_result_book()
g.eff_test.run_verification()
g.excel_m.end_of_file(0)
