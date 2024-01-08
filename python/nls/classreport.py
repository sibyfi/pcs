
from datetime import datetime
import csv
import re

class ComputeDate:
    """ This Class computes dates """
    def __init__(self, start_date, end_date):
        """ This method pulls in the start and end dates """
        self.start_date = start_date
        self.end_date   = end_date
    def date_diff(self):
        """ This computes the total number of days for the class """
        d1 = datetime.strptime(self.start_date, "%Y-%m-%d")
        d2 = datetime.strptime(self.end_date, "%Y-%m-%d")
        return abs((d2 - d1).days) + 1

class CheckDatesAlert:
    def __init__(self, start_date):
        """ This method pulls in the start date """
        self.start_date = start_date
    def alert(self):
        """ This method calculates the duration of the course. """ 
        d1 = datetime.strptime(self.start_date, "%Y-%m-%d")
        d2 = datetime.now()
        return (abs(d2 - d1).days)

class BuildCsv:
    """ This class manipulates the csv data """
    def __init__(self, reader):
        """ This method pulls in the data """
        self.reader          = reader
        self.header          = next(self.reader)
        self.course_name     = self.header[1]
        self.course_num      = self.header[0]
        self.start_date      = self.header[5]
        self.end_date        = self.header[6]
        self.location        = self.header[7]
        self.region          = self.header[8]
        self.max_students    = self.header[13]
        self.enrolled        = self.header[12]
        self.rep             = self.header[14]
        self.offering_num    = self.header[4]
        self.enroll          = self.header[11]
        self.cat_dom_name    = self.header[2]
        self.offering_dom    = self.header[3]
        self.disp_for_learn  = self.header[10]
        self.class_type      = self.header[15]
        self.offer_status    = self.header[16]
        self.content_version = self.header[17]
        self.instructor      = self.header[9]

        self.new_header = [self.course_name, self.course_num,
            self.start_date, self.end_date, "Course Duration (Days)", self.location, self.region,
            self.max_students, self.enrolled, "Open Seats", self.rep, self.offering_num,
            self.enroll, self.cat_dom_name, self.offering_dom, self.disp_for_learn, 
            self.class_type, self.offer_status, self.content_version, self.instructor]

    def create_csv_header(self):
        """ This method formats the header """
        return self.new_header

    def len_data(self):
        """ Calculate length of header """
        self.len_data = len(self.new_header)
        return self.len_data

    def reformat_date(self, old_date):
        """
        this method is to take into consideration
        the new PowerBI date format
        """
        padded_date         = ''.join(x.zfill(2) for x in old_date.split('/'))
        new_date            = f"20{padded_date[-2:]}-{padded_date[0:2]}-{padded_date[2:4]}"
        return new_date

    def create_csv_data(self):
        """ This method formats the csv data """
        data = []
        for row in self.reader:
            course_number          = str(row[0])
            course_name            = str(row[1])
            catalog_domain_Name    = str(row[2])
            offering_domain        = str(row[3])
            offering_number        = str(row[4])
            #
            start_date             = self.reformat_date(row[5])
            end_date               = self.reformat_date(row[6])
            # start_date             = row[5]
            # end_date               = row[6]
            ####
            offering_start_date    = self.reformat_date(row[5])
            offering_end_date      = self.reformat_date(row[6])
            compute_duration       = ComputeDate(offering_start_date, offering_end_date)
            course_duration        = compute_duration.date_diff()
            offering_location      = str(row[7])
            offering_region        = str(row[8])
            offering_instructor    = str(row[9])
            display_for_learner    = int(row[10])
            enroll                 = str(row[11])
            currently_enrolled     = int(row[12])
            max_student_count      = int(row[13])
            open_seats             = max_student_count - currently_enrolled
            customer_service_rep   = str(row[14])
            class_type             = str(row[15])
            offering_status        = str(row[16])
            content_version_number = str(row[17])
            """
            if (customer_service_rep == "RUCKER_INTERNAL"):
                data.append([course_name, course_number, offering_start_date, offering_end_date, course_duration,
                    offering_location, offering_region, max_student_count, currently_enrolled, open_seats, customer_service_rep,
                    offering_number, enroll, catalog_domain_Name, offering_domain, display_for_learner,
                    class_type, offering_status, content_version_number])
            """
            data.append([course_name, course_number, offering_start_date, offering_end_date, course_duration,
                offering_location, offering_region, max_student_count, currently_enrolled, open_seats, customer_service_rep,
                offering_number, enroll, catalog_domain_Name, offering_domain, display_for_learner,
                class_type, offering_status, content_version_number, offering_instructor])

        return data

class StripUrl:
    """ This Class reformats the URL """
    def __init__(self, enroll, offering_number, offering_location):
        """ This inits the variables """
        self.enroll            = enroll
        self.offering_number   = offering_number
        self.offering_location = offering_location
    def fixurl(self):
        """ This method fixes the broken URL's and
        puts them in the correct for mat for excel """
        if re.findall("Virtual", self.offering_location):
            url_base = "https://netapp.sabacloud.com/Saba/Web_spf/NA1PRD0047/common/leclassview/virtc-"
            return (str(url_base) + str(self.offering_number))
        else:        
            url_base = "https://netapp.sabacloud.com/Saba/Web_spf/NA1PRD0047/common/leclassview/class-"
            return (str(url_base) + str(self.offering_number))
            # regex = re.compile('{}(.*){}'.format(re.escape('<a href='), re.escape('>Enroll')))
            # new_url = regex.findall(self.enroll)[0]
            # return new_url



