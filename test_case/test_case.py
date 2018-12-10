from common.my_unit import MyUnit
from common.common_func import create_bat_for_launch, gen_app, launch_app, record_data, close_window,get_screen_shot, init_excel, content_verify, get_app_output


class SDKTest(MyUnit):

    def test_case(self):
        workbook, worksheet = init_excel()
        for index, app in enumerate(gen_app(), 2):
            create_bat_for_launch(app)
            launch_app()
            record_data(app, index, worksheet)
            get_screen_shot(app)
            close_window()
            get_app_output(app)
            content_verify(app, worksheet, index)
        workbook.close()