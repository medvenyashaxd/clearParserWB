from kivy.lang import Builder
from kivymd.app import MDApp
from parserWB import start


class ParserFeedbacksWBApp(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.screen = Builder.load_file('ParserFeedbacksWB.kv')
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Green"

    def build(self):
        return self.screen

    def get_input_data(self, code_WB, identificator_SW=None):
        filter_code_wb = self.screen.ids.text1.text.replace('\n', ',').split(',')
        filter_identificator_sw = self.screen.ids.text2.text.replace('\n', ',').split(',')

        if filter_identificator_sw[0] == '':

            for code in filter_code_wb:
                start(code)

        else:
            for code, ids in zip(filter_code_wb, filter_identificator_sw):
                start(code, ids)


if __name__ == '__main__':
    ParserFeedbacksWBApp().run()
