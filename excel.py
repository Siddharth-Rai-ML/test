import datetime

import xlsxwriter

import constants
from logger import log


class Excel:
    def __init__(self, filename):
        self._filename = filename
        self._filepath = './' + self._filename
        self._workbook = xlsxwriter.Workbook(filename)
        self._cell_formatter()

    def _cell_formatter(self):
        self.cell_format1 = self._workbook.add_format({'font_color': 'blue'})
        
        self.cell_format2 = self._workbook.add_format()
        self.cell_format2.set_pattern(1)
        self.cell_format2.set_border(1)
        self.cell_format2.set_bg_color('green')
        self.cell_format2.set_font_size(16)
        self.cell_format2.set_align('center')
        self.cell_format2.set_font_color('white')
        self.cell_format2.set_bold()

        self.cell_format3 = self._workbook.add_format()
        self.cell_format3.set_pattern(1)
        self.cell_format3.set_border(1)
        self.cell_format3.set_align('center')
        self.cell_format3.set_font_color('black')

        self.cell_format4 = self._workbook.add_format()
        self.cell_format4.set_pattern(1)
        self.cell_format4.set_border(1)
        self.cell_format4.set_align('center')
        self.cell_format4.set_font_color('black')

        self.cell_format5 = self._workbook.add_format()
        self.cell_format5.set_pattern(1)
        self.cell_format5.set_border(1)
        self.cell_format5.set_align('center')
        self.cell_format5.set_font_color('black')

        self.cell_format6 = self._workbook.add_format()
        self.cell_format6.set_pattern(1)
        self.cell_format6.set_border(1)
        self.cell_format6.set_align('center')
        self.cell_format6.set_font_color('black')

        self.cell_formats = self._workbook.add_format()
        self.cell_formats.set_align('center')
        self.cell_formats.set_border(1)
        self.cell_formats.set_font_color('black')
        self.cell_formats.set_bold()

    def new_worksheet(self, sheet_name):
        return self._workbook.add_worksheet(sheet_name)
    
    def create_summary_worksheet(self):
        row = 0
        col = 0

        self.summary_worksheet = self.new_worksheet(constants.SUMMARY_SHEET_NAME)
        self.summary_worksheet.set_column('A:A', 30)
        self.summary_worksheet.set_column('B:B', 30)
        self.summary_worksheet.set_column('C:C', 25)
        self.summary_worksheet.set_column('D:D', 25)
        self.summary_worksheet.set_column('E:E', 10)
        self.summary_worksheet.set_column('G:G', 10)
        self.summary_worksheet.set_column('H:H', 10)
        self.summary_worksheet.set_column('I:I', 10)
        self.summary_worksheet.set_column('J:J', 60)
        self.summary_worksheet.set_column('K:K', 30)
        self.summary_worksheet.set_column('L:L', 25)
        self.summary_worksheet.set_column('M:M', 30)
        self.summary_worksheet.set_column('N:N', 25)

        self.summary_worksheet.write(row, col, "Image_Name", self.cell_format2)
        self.summary_worksheet.write(row, col + 1, "Registry", self.cell_format2)
        self.summary_worksheet.write(row, col + 2, "sha_id", self.cell_format2)
        self.summary_worksheet.write(row, col + 3, "Tag", self.cell_format2)
        self.summary_worksheet.write(row, col + 4, "Critical", self.cell_format2)
        self.summary_worksheet.write(row, col + 5, "High", self.cell_format2)
        self.summary_worksheet.write(row, col + 6, "Medium", self.cell_format2)
        self.summary_worksheet.write(row, col + 7, "Low", self.cell_format2)
        self.summary_worksheet.write(row, col + 8, "Total", self.cell_format2)
        self.summary_worksheet.write(row, col + 9, "Maintainer", self.cell_format2)
        self.summary_worksheet.write(row, col + 10, "Application_Name", self.cell_format2)
        self.summary_worksheet.write(row, col + 11, "Cost_Center", self.cell_format2)
        self.summary_worksheet.write(row, col + 12, "Email_Distribution", self.cell_format2)
        self.summary_worksheet.write(row, col + 13, "apm", self.cell_format2)
        self.summary_worksheet.write(row, col + 14, "bit", self.cell_format2)

    def create_sub_sheet(self, sheet_name):
        vrow = 0
        vcol = 0
        sworksheet = self.new_worksheet(sheet_name)
        sworksheet.set_column('A:A', 60)
        sworksheet.set_column('B:B', 30)
        sworksheet.set_column('C:C', 30)
        sworksheet.set_column('D:D', 30)
        sworksheet.set_column('E:E', 30)
        sworksheet.set_column('F:F', 30)
        sworksheet.set_column('G:G', 30)
        sworksheet.set_column('H:H', 30)
        sworksheet.set_column('I:I', 30)

        sworksheet.write(vrow, vcol, "PackageName", self.cell_format2)
        sworksheet.write(vrow, vcol + 1, "Severity", self.cell_format2)
        sworksheet.write(vrow, vcol + 2, "CVEID", self.cell_format2)
        sworksheet.write(vrow, vcol + 3, "Fixstatus", self.cell_format2)
        sworksheet.write(vrow, vcol + 4, "Base_image_distro", self.cell_format2)
        sworksheet.write(vrow, vcol + 5, "PackageVersion", self.cell_format2)
        sworksheet.write(vrow, vcol + 6, "Discovered_date", self.cell_format2)
        sworksheet.write(vrow, vcol + 7, "Published_date", self.cell_format2)
        sworksheet.write(vrow, vcol + 8, "Fix_date", self.cell_format2)
        return sworksheet

    def write_to_excel(self, rows):
        log.debug("creating summary sheet")
        self.create_summary_worksheet()

        row = 0
        col = 0

        i = 0
        x = 0
        log.debug(f'length of rows: {len(rows)}')

        for i, table_row in enumerate(rows):
            # ID, Registry, Sha_id, Tag, Distro, Critical, High, Medium, Low, Total, result, repo, vulnerabilities = table_row
            ID, Registry, Sha_id, Tag, Distro, Critical, High, Medium, Low, Total, maintainer, Application, Cost, Email, APM, BIT, result, repo, vulnerabilities = table_row
            if Critical == 0:
                self.cell_format6.set_bg_color('white')
            elif Critical > 0:
                self.cell_format6.set_bg_color('red')

            if High == 0:
                self.cell_format5.set_bg_color('white')
            elif High > 0:
                self.cell_format5.set_bg_color('red')

            if Medium == 0:
                self.cell_format4.set_bg_color('white')
            elif Medium > 0:
                self.cell_format4.set_bg_color('orange')

            if Low == 0:
                self.cell_format3.set_bg_color('white')
            elif Low > 0:
                self.cell_format3.set_bg_color('yellow')

            log.debug(f'writing row : {row+1}')
            self.summary_worksheet.write(row + 1, col, ID, self.cell_format1)
            self.summary_worksheet.write(row + 1, col + 1, Registry)
            self.summary_worksheet.write(row + 1, col + 2, Sha_id)
            self.summary_worksheet.write(row + 1, col + 3, Tag)
            self.summary_worksheet.write(row + 1, col + 4, Critical, self.cell_format6)
            self.summary_worksheet.write(row + 1, col + 5, High, self.cell_format5)
            self.summary_worksheet.write(row + 1, col + 6, Medium, self.cell_format4)
            self.summary_worksheet.write(row + 1, col + 7, Low, self.cell_format3)
            self.summary_worksheet.write(row + 1, col + 8, Total)
            self.summary_worksheet.write(row + 1, col + 9, maintainer)
            self.summary_worksheet.write(row + 1, col + 10, Application)
            self.summary_worksheet.write(row + 1, col + 11, Cost)
            self.summary_worksheet.write(row + 1, col + 12, Email)
            self.summary_worksheet.write(row + 1, col + 13, APM)
            self.summary_worksheet.write(row + 1, col + 14, BIT)
            if result is None:
                pass
            else:
                vrow = 0
                vcol = 0
                sheet_name = str(ID) + "-" + str(i + 1)
                sheet_name = str(sheet_name.split("/"))
                sheet_name = sheet_name.replace("[", "_")
                sheet_name = sheet_name.replace("/", "_")
                sheet_name = sheet_name.replace(":", "_")
                sheet_name = sheet_name.replace("'", "_")
                spsheet = sheet_name[-25:-1]

                log.debug(f'creating sheet - {spsheet}')
                sworksheet = self.create_sub_sheet(spsheet)

                x = x + 1
                self.summary_worksheet.set_column('A:A', 60)
                self.summary_worksheet.write_url(x, 0, f"internal:'{spsheet}'!A1", string=repo)

                vrow = vrow + 1
                sub_sheet_rows = list()
                for index, vul in enumerate(vulnerabilities):
                    Severity = vul['severity']
                    CveId = vul['cve']
                    Fixstatus = vul['status']
                    PackageName = vul['packageName']
                    PackageVersion = vul['packageVersion']
                    Discovered_date = vul['discovered']
                    Discovered_date = Discovered_date.split('T')[0]
                    Published_date = vul['published']
                    Published_date = datetime.datetime.fromtimestamp(Published_date)
                    Published_date = Published_date.strftime('%Y-%m-%d')
                    Fix_date = vul['fixDate']
                    Fix_date = datetime.datetime.fromtimestamp(Fix_date)
                    Fix_date = Fix_date.strftime('%Y-%m-%d')

                    sub_sheet_rows.append([PackageName, Severity, CveId, Fixstatus, PackageVersion, Discovered_date, Published_date, Fix_date])
                
                def sort_sub_sheet(sub_sheet_rows):
                    sorted_rows = list()
                    critical = [row for row in sub_sheet_rows if row[1] == 'critical']
                    high = [row for row in sub_sheet_rows if row[1] == 'high']
                    important = [row for row in sub_sheet_rows if row[1] == 'important']
                    moderate = [row for row in sub_sheet_rows if row[1] == 'moderate']
                    medium = [row for row in sub_sheet_rows if row[1] == 'medium']
                    low = [row for row in sub_sheet_rows if row[1] == 'low']
                    sorted_rows.extend(critical)
                    sorted_rows.extend(high)
                    sorted_rows.extend(important)
                    sorted_rows.extend(moderate)
                    sorted_rows.extend(medium)
                    sorted_rows.extend(low)
                    return sorted_rows

                def move_blanks_to_bottom(rows, key_position):
                    sorted_rows = list()
                    empty_rows = list()
                    non_empty_rows =list()
                    for row in rows:
                        if row[key_position]:
                            non_empty_rows.append(row)
                        else:
                            empty_rows.append(row)
                    sorted_rows.extend(non_empty_rows)
                    sorted_rows.extend(empty_rows)
                    return sorted_rows

                vrow = 1


                log.debug(f"total sub rows: {len(sub_sheet_rows)}")
                for index, sub_sheet_row in enumerate(move_blanks_to_bottom(sort_sub_sheet(sub_sheet_rows), 3)):
                    PackageName, Severity, CveId, Fixstatus, PackageVersion, Discovered_date, Published_date, Fix_date = sub_sheet_row
                    if str(Severity) == 'low':
                        self.cell_formats.set_bg_color('green')
                    elif str(Severity) == 'medium':
                        self.cell_formats.set_bg_color('yellow')
                    elif str(Severity) == 'moderate':
                        self.cell_formats.set_bg_color('orange')
                    elif str(Severity) == 'important':
                        self.cell_formats.set_bg_color('orange')
                    elif str(Severity) == 'high':
                        self.cell_formats.set_bg_color('red')
                    elif str(Severity) == 'critical':
                        self.cell_formats.set_bg_color('red')
                    

                    print(f'writing sub row : {vrow + index}/{len(sub_sheet_rows)}, severity = {Severity}, {self.cell_formats.bg_color}')
                    sworksheet.write(vrow + index, vcol + 0, PackageName)
                    sworksheet.write(vrow + index, vcol + 1, Severity, self.cell_formats)
                    sworksheet.write(vrow + index, vcol + 2, CveId)
                    sworksheet.write(vrow + index, vcol + 3, Fixstatus)
                    sworksheet.write(vrow + index, vcol + 4, Distro)
                    sworksheet.write(vrow + index, vcol + 5, PackageVersion)
                    sworksheet.write(vrow + index, vcol + 6, Discovered_date)
                    sworksheet.write(vrow + index, vcol + 7, Published_date)
                    sworksheet.write(vrow + index, vcol + 8, Fix_date)
            row = row + 1

        print('+++++++++ finished  +++++++')

    def __exit__(self, exc_type, exc_val, exc_tb):
        print('workbook closed')
        self._workbook.close()
