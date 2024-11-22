from pyhwpx import Hwp

class HwpWeekTableFont:
    def __init__(self, file_path):
        self.hwp = Hwp()
        self.file_path = file_path
        self.hwp.Open(file_path, arg="versionwarning:false")

    def set_spacing(self, line_spacing):
        """
        줄 간격 설정 함수
        """
        paragraph_shape = self.hwp.XHwpDocuments.Item(0).XHwpParagraphShape
        paragraph_shape.LineSpacing = line_spacing

    def format_cell(self, pos, line_spacing, align_method, font_settings):
        """
        요일 폰트 설젇
        """
        self.hwp.set_pos(*pos)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColPageDown()
        align_method()
        self.set_spacing(line_spacing)
        self.hwp.set_font(**font_settings)

    def process_week_cells(self):
        """
        요일 제외 폰트 설정
        """
        for i in range(7):
            pos = (14 + i * 13, 0, 0)
            font_settings = {
                "Height": 12,
                "FaceName": "나눔명조",
                "Ratio": 80,
                "Spacing": 0,
                "TextColor": self.hwp.rgb_color(0, 0, 0)
            }
            if i == 5:
                font_settings["TextColor"] = self.hwp.rgb_color(0, 0, 255)
            elif i == 6:
                font_settings["TextColor"] = self.hwp.rgb_color(255, 0, 0)

            self.format_cell(pos, line_spacing=100, align_method=self.hwp.ParagraphShapeAlignCenter, font_settings=font_settings)

    def process_event_details(self):
        for i in range(15, 21):
            pos = (i, 0, 0)
            font_settings = {
                "Height": 12,
                "FaceName": "나눔명조",
                "Ratio": 80,
                "Spacing": 0,
                "TextColor": self.hwp.rgb_color(0, 0, 0)
            }
            line_spacing = 100

            if i == 15:
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignCenter, font_settings)
            elif i == 16:
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignLeft, font_settings)
            elif i == 17:
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignDistribute, font_settings)
            elif i == 18:
                self.format_cell(pos, 250, self.hwp.ParagraphShapeAlignRight, font_settings)
            elif i == 19:
                self.format_cell(pos, 110, self.hwp.ParagraphShapeAlignDistribute, font_settings)
            elif i == 20:
                font_settings["Bold"] = 1
                font_settings["Ratio"] = 85
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignDistribute, font_settings)

    def process(self):
        self.process_week_cells()
        self.process_event_details()

"""if __name__ == "__main__":
    template_path = r"C:\Users\thdco\OneDrive\Documents\GitHub\font_set\font_set\template2.hwp"
    week_formatter = HwpWeekTableFont(template_path)
    week_formatter.process()"""