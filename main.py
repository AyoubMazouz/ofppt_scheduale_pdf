import pyexcel as p
from fpdf import FPDF, XPos, YPos
import math


# Records data structure:
# [
#   Single record:
#   {
#       details:
#       {
#           groupe.
#           total_hours.
#       }
#       record:
#       [
#           day:
#           [
#               sessions:
#                   [
#                       professor.
#                       module.
#                       classroom.
#                   ]
#                   ...More Sessions
#           ]
#           ... More Days
#       ]
#   }
#   ...More Records
# ]
#
#


# Constants:
GIT_REPO = "https://github.com/AyoubMazouz/ofppt_scheduale_pdf.git"
LABELS = ["prof", "mod", "room"]
DAYS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]
ACCENT_COLOR = [37, 49, 109]
LIGHT_COLOR = [238, 238, 238]
DARK_COLOR = [12, 12, 12]


def get_records():
    print("Start extracting records from Excel sheet")

    # Pdf file name.
    records = p.get_array(file_name="s4.xlsx")

    print("Record extracted successfully")
    print("Start processing records")

    processed_records = []
    # Each 3 rows is considered one row.
    for i in range(0, len(records), 3):
        # Make sure only take valid row by checking if column one has "NTIC" str in it.
        if str(records[i][0]).find("NTIC") != -1:
            # Details object is for storing:
            # date, groupe, etc...
            single_record = {"details": {}}
            # Session basically is a period e.g:
            # from 8:30 to 11:00
            session = []
            # Since record is a continuous array of value we need to count session,
            # is_day_complete == 4 => day
            is_complete_day = 0
            day = []
            total_hours = 0

            # Start from index 2 to skip groupe cells at col 1 & labels at col 2
            for j in range(2, len(records[i])):
                tempList = []
                tempList.append(records[i][j])
                tempList.append(records[i + 1][j])
                tempList.append(records[i + 2][j])
                day.append(tempList)
                is_complete_day += 1
                total_hours += 2.5 if records[i][j] != "" else 0

                if is_complete_day == 4:
                    session.append(day)
                    day = []
                    is_complete_day = 0

            groupe_name = records[i][0].replace("NTIC1-", "")
            single_record["details"]["groupe"] = groupe_name
            single_record["details"]["total_hours"] = str(total_hours - 2.5)
            single_record["record"] = session
            processed_records.append(single_record)

            print(f"Finish processing {groupe_name} record")

    return processed_records


class PDF(FPDF):
    def __init__(self, **kwarg):
        super(PDF, self).__init__(**kwarg)
        # Initial params.
        self.set_font("helvetica")
        self.add_page()
        self.set_margin(3)
        # Calculate cell width to use all space available.
        self.cell_w = (self.w / 5) - (self.l_margin / 2)
        self.cell_h = 9

    def render_details(self, details):
        # Details
        details_str = (
            f'Groupe: {details["groupe"]}; Total hours: {details["total_hours"]};'
        )
        self.set_font(size=20)
        self.set_text_color(*DARK_COLOR)
        self.cell(130, 12, details_str)
        # Git
        git_str = "automated with > " + GIT_REPO
        self.set_font(size=13, style="IU")
        self.set_text_color(*ACCENT_COLOR)
        self.cell(0, 12, git_str, link=GIT_REPO)

    def render_table(self, record):
        self.set_font(size=28)
        self.set_fill_color(*ACCENT_COLOR)
        self.set_text_color(*LIGHT_COLOR)
        # fist row label session duration
        self.cell(self.cell_w, self.cell_h * 3, "")
        self.cell(self.cell_w, self.cell_h * 3, "8:30-11:00", border=1, fill=True)
        self.cell(self.cell_w, self.cell_h * 3, "11:00-13:30", border=1, fill=True)
        self.cell(self.cell_w, self.cell_h * 3, "13:30-16:00", border=1, fill=True)
        self.cell(self.cell_w, self.cell_h * 3, "16:00-18:30", border=1, fill=True)
        self.ln()

        for d_index, day in enumerate(record):
            for i in range(0, len(day[0])):
                # Decide which sides of the border to render.
                border_switcher = {0: "LTR", 1: "LR", 2: "LBR"}
                # day
                show_day = DAYS[d_index] if i == 1 else ""

                # First column in each row represent days of the week.
                self.set_font(size=32)
                self.set_fill_color(*ACCENT_COLOR)
                self.set_text_color(*LIGHT_COLOR)
                self.cell(
                    self.cell_w,
                    self.cell_h,
                    show_day,
                    border=border_switcher[i],
                    fill=True,
                )
                # sessions
                if i == 1:
                    self.set_font(size=14)
                else:
                    self.set_font(size=18)
                self.set_fill_color(*LIGHT_COLOR)
                self.set_text_color(*DARK_COLOR)
                self.cell(
                    self.cell_w,
                    self.cell_h,
                    day[0][i],
                    border=border_switcher[i],
                    fill=(len(day[0][i])),
                )
                self.cell(
                    self.cell_w,
                    self.cell_h,
                    day[1][i],
                    border=border_switcher[i],
                    fill=(len(day[1][i])),
                )
                self.cell(
                    self.cell_w,
                    self.cell_h,
                    day[2][i],
                    border=border_switcher[i],
                    fill=(len(day[2][i])),
                )
                self.cell(
                    self.cell_w,
                    self.cell_h,
                    day[3][i],
                    border=border_switcher[i],
                    fill=(len(day[3][i])),
                )
                # new line
                self.ln()


def main():
    records = get_records()

    pdf = PDF(orientation="landscape", format="A4", unit="mm")

    print(f"Start rendering")
    for record in records:
        print(f'Rendering {record["details"]["groupe"]} record')
        pdf.render_table(record["record"])
        pdf.render_details(record["details"])
        pdf.add_page(same=True)

    pdf.output("output.pdf")
    print(f"Done")


if __name__ == "__main__":
    main()
