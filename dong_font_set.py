from pyhwpx import Hwp

class HwpDongTableFont:
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
        self.hwp.set_pos(*pos)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColPageDown()
        align_method()
        self.set_spacing(line_spacing)
        self.hwp.set_font(**font_settings)

    def process_table_cells(self):
        """
        요일 폰트 설젇
        """
        for i in range(7):
            pos = (11 + i * 13, 0, 0)
            font_settings = {
                "Bold": 1,
                "Height": 11,
                "FaceName": "굴림",
                "TextColor": self.hwp.rgb_color(0, 0, 0)
            }
            if i == 5:
                font_settings["TextColor"] = self.hwp.rgb_color(0, 0, 255)
            elif i == 6:
                font_settings["TextColor"] = self.hwp.rgb_color(255, 0, 0)

            self.format_cell(pos, line_spacing=120, align_method=self.hwp.ParagraphShapeAlignCenter, font_settings=font_settings)

    def process_event_details(self):
        """
        요일 제외 폰트 설정
        """
        for i in range(12, 18):
            pos = (i, 0, 0)
            font_settings = {
                "Bold": 1 if i == 12 else 0,
                "Height": 11,
                "FaceName": "굴림",
                "TextColor": self.hwp.rgb_color(0, 0, 0)
            }
            line_spacing = 100

            if i == 13:
                font_settings["TextColor"] = self.hwp.rgb_color(0, 0, 255)

            if i == 12:
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignCenter, font_settings)
            elif i == 13:
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignCenter, font_settings)
            elif i == 14:
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignLeft, font_settings)
            elif i == 15:
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignDistribute, font_settings)
            elif i == 16:
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignCenter, font_settings)
            elif i == 17:
                self.format_cell(pos, line_spacing, self.hwp.ParagraphShapeAlignDistribute, font_settings)


    def process(self):
        self.process_table_cells()
        self.process_event_details()

"""if __name__ == "__main__":
    template_path = r"C:\Users\thdco\OneDrive\Documents\GitHub\font_set\font_set\dongtemplate.hwp"
    formatter = HwpDongTableFont(template_path)
    formatter.process()"""