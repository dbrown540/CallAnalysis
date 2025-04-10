from datetime import datetime
from openpyxl import Workbook

call_data = {
    "December 25, 2023": [("status", "Canceled"), ("missed_by", "incoming"), ("status", "Canceled"), ("status", "Canceled"), ("incoming", 5)],
    "December 26, 2023": [("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("incoming", 15/60), ("incoming", 5), ("incoming", 15)],
    "December 27, 2023": [("status", "Canceled"), ("status", "Canceled")],
    "December 28, 2023": [("status", "Canceled"), ("outgoing", 32), ("incoming", 2)],
    "December 29, 2023": [("status", "Canceled"), ("status", "Canceled")],
    "December 30, 2023": [("status", "Canceled")],
    "December 31, 2023": [("status", "Canceled"), ("outgoing", 38), ("missed_by", "incoming"), ("status", "Canceled")],
    "January 1, 2024": [("status", "Canceled"), ("outgoing", 10)],
    "January 2, 2024": [("status", "Canceled"), ("incoming", 12)],
    "January 3, 2024": [("status", "Canceled")],
    "January 4, 2024": [("status", "Canceled"), ("incoming", 8)],
    "January 5, 2024": [("outgoing", 6), ("missed_by", "incoming"), ("status", "Canceled"), ("missed_by", "incoming")],
    "January 7, 2024": [("status", "Canceled"), ("outgoing", 7), ("incoming", 15)],
    "January 8, 2024": [("status", "Canceled"), ("incoming", 15)],
    "January 9, 2024": [("status", "Canceled")],
    "January 10, 2024": [("status", "Canceled"), ("incoming", 7)],
    "January 11, 2024": [("status", "Canceled"), ("missed_by", "incoming"), ("status", "Canceled")],
    "January 12, 2024": [("status", "Canceled"), ("status", "Canceled"), ("outgoing", 5), ("missed_by", "incoming"), ("status", "Canceled"), ("outgoing", 4)],
    "January 13, 2024": [("status", "Canceled"), ("incoming", 5), ("outgoing", 2)],
    "January 15, 2024": [("status", "Canceled")],
    "January 16, 2024": [("outgoing", 2)],
    "January 17, 2024": [("status", "Canceled"), ("incoming", 3)],
    "January 18, 2024": [("status", "Canceled")],
    "January 19, 2024": [("outgoing", 6), ("missed_by", "incoming"), ("status", "Canceled")],
    "January 20, 2024": [("status", "Canceled")],
    "January 21, 2024": [("status", "Canceled"), ("incoming", 7)],
    "January 22, 2024": [("status", "Canceled"), ("incoming", 9), ("missed_by", "incoming"), ("status", "Canceled"), ("status", "Canceled")],
    "January 23, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "January 24, 2024": [("status", "Canceled"), ("incoming", 12)],
    "January 25, 2024": [("status", "Canceled"), ("outgoing", 12)],
    "January 26, 2024": [("status", "Canceled")],
    "January 27, 2024": [("status", "Canceled")],
    "January 28, 2024": [("status", "Canceled"), ("incoming", 4)],
    "January 29, 2024": [("status", "Canceled")],
    "January 30, 2024": [("incoming", 23)],
    "January 31, 2024": [("outgoing", 1)],
    "February 1, 2024": [("status", "Canceled")],
    "February 2, 2024": [("status", "Canceled"), ("incoming", 2)],
    "February 3, 2024": [("status", "Canceled"), ("outgoing", 15)],
    "February 4, 2024": [("status", "Canceled"), ("outgoing", 51/66), ("incoming", 15)],
    "February 5, 2024": [("status", "Canceled"), ("status", "Canceled"), ("incoming", 3)],
    "February 6, 2024": [("status", "Canceled")],
    "February 7, 2024": [("incoming", 9), ("status", "Canceled")],
    "February 8, 2024": [("outgoing", 7/60), ("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled")],
    "February 9, 2024": [("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("outgoing", 5), ("incoming", 50/60)],
    "February 10, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "February 11, 2024": [("status", "Canceled"), ("outgoing", 7)],
    "February 12, 2024": [("status", "Canceled")],
    "February 13, 2024": [("status", "Canceled"), ("incoming", 26), ("outgoing", 11/60)],
    "February 14, 2024": [("outgoing", 8), ("outgoing", 2)],
    "February 15, 2024": [("status", "Canceled")],
    "February 16, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "February 17, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "February 18, 2024": [("status", "Canceled"), ("outgoing", 25/60), ("outgoing", 1), ("missed_by", "incoming")],
    "February 19, 2024": [("status", "Canceled"), ("outgoing", 11), ("incoming", 9)],
    "February 20, 2024": [("status", "Canceled")],
    "February 21, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "February 22, 2024": [("status", "Canceled"), ("outgoing", 51)],
    "February 23, 2024": [("status", "Canceled"), ("status", "Canceled"), ("incoming", 3), ("outgoing", 17)],
    "February 24, 2024": [("status", "Canceled")],
    "February 25, 2024": [("status", "Canceled")],
    "February 26, 2024": [("status", "Canceled"), ("outgoing", 4)],
    "February 27, 2024": [("status", "Canceled")],
    "February 28, 2024": [("status", "Canceled"), ("incoming", 10), ("incoming", 5)],
    "February 29, 2024": [("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("incoming", 20), ("incoming", 1)],
    "March 1, 2024": [("status", "Canceled")],
    "March 2, 2024": [("outgoing", 20)],
    "March 3, 2024": [("status", "Canceled"), ("outgoing", 9), ("incoming", 3)],
    "March 4, 2024": [("status", "Canceled")],
    "March 5, 2024": [("status", "Canceled"), ("incoming", 3)],
    "March 6, 2024": [("status", "Canceled"), ("status", "Canceled"), ("incoming", 10)],
    "March 10, 2024": [("status", "Canceled")],
    "March 11, 2024": [("outgoing", 6)],
    "March 12, 2024": [("status", "Canceled")],
    "March 13, 2024": [("status", "Canceled")],
    "March 14, 2024": [("outgoing", 9), ("outgoing", 19)],
    "March 16, 2024": [("outgoing", 5/60), ("outgoing", 8/60)],
    "March 18, 2024": [("status", "Canceled")],
    "March 19, 2024": [("status", "Canceled")],
    "March 20, 2024": [("status", "Canceled"), ("incoming", 4), ("status", "Canceled")],
    "March 21, 2024": [("status", "Canceled")],
    "March 22, 2024": [("status", "Canceled")],
    "March 23, 2024": [("status", "Canceled")],
    "March 24, 2024": [("status", "Canceled")],
    "March 26, 2024": [("status", "Canceled")],
    "March 27, 2024": [("status", "Canceled"), ("incoming", 6), ("outgoing", 2)],
    "March 28, 2024": [("status", "Canceled"), ("outgoing", 15)],
    "March 29, 2024": [("outgoing", 11), ("status", "Canceled")],
    "March 30, 2024": [("status", "Canceled")],
    "March 31, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "April 1, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "April 2, 2024": [("status", "Canceled"), ("status", "Canceled"), ("incoming", 8)],
    "April 3, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "April 4, 2024": [("status", "Canceled")],
    "April 5, 2024": [("status", "Canceled"), ("status", "Canceled"), ("outgoing", 7)],
    "April 6, 2024": [("status", "Canceled"), ("incoming", 8)],
    "April 7, 2024": [("status", "Canceled")],
    "April 8, 2024": [("status", "Canceled"), ("status", "Canceled"), ("incoming", 31)],
    "April 9, 2024": [("status", "Canceled")],
    "April 10, 2024": [("status", "Canceled"), ("status", "Canceled"), ("incoming", 11)],
    "April 11, 2024": [("status", "Canceled"), ("outgoing", 22)],
    "April 12, 2024": [("status", "Canceled")],
    "April 13, 2024": [("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("incoming", 22)],
    "April 14, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "April 15, 2024": [("status", "Canceled")],
    "April 16, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "April 17, 2024": [("outgoing", 8)],
    "April 18, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "April 19, 2024": [("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("incoming", 3), ("incoming", 7)],
    "April 20, 2024": [("status", "Canceled"), ("outgoing", 16)],
    "April 21, 2024": [("missed_by", "incoming"), ("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("incoming", 10)],
    "April 22, 2024": [("status", "Canceled")],
    "April 23, 2024": [("outgoing", 19), ("status", "Canceled")],
    "April 24, 2024": [("status", "Canceled")],
    "April 25, 2024": [("outgoing", 9)],
    "April 26, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "April 27, 2024": [("outgoing", 1), ("status", "Canceled")],
    "April 28, 2024": [("outgoing", 1), ("outgoing", 18), ("missed_by", "incoming"), ("status", "Canceled")],
    "April 29, 2024": [("status", "Canceled")],
    "April 30, 2024": [("status", "Canceled")],
    "May 1, 2024": [("status", "Canceled"), ("status", "Canceled"), ("incoming", 34)],
    "May 2, 2024": [("status", "Canceled")],
    "May 3, 2024": [("outgoing", 8), ("outgoing", 4)],
    "May 4, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "May 5, 2024": [("status", "Canceled"), ("status", "Canceled"), ("outgoing", 4), ("incoming", 8), ("missed_by", "incoming"), ("missed_by", "incoming"), ("missed_by", "incoming")],
    "May 6, 2024": [("status", "Canceled"), ("missed_by", "incoming"), ("outgoing", 2), ("outgoing", 15)],
    "May 7, 2024": [("status", "Canceled")],
    "May 8, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "May 9, 2024": [("status", "Canceled"), ("status", "Canceled"), ("outgoing", 3)],
    "May 10, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "May 11, 2024": [("missed_by", "incoming"), ("missed_by", "incoming"), ("missed_by", "incoming"), ("missed_by", "incoming"), ("missed_by", "incoming")],
    "May 12, 2024": [("incoming", 17), ("status", "Canceled")],
    "May 14, 2024": [("status", "Canceled")],
    "May 15, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "May 16, 2024": [("status", "Canceled"), ("outgoing", 11)],
    "May 17, 2024": [("status", "Canceled")],
    "May 18, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "May 19, 2024": [("status", "Canceled")],
    "May 20, 2024": [("status", "Canceled"), ("outgoing", 1), ("outgoing", 3/60), ("outgoing", 4)],
    "May 21, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "May 22, 2024": [("status", "Canceled")],
    "May 23, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "May 24, 2024": [("status", "Canceled"), ("status", "Canceled"), ("incoming", 24)],
    "May 25, 2024": [("status", "Canceled")],
    "May 26, 2024": [
        ("status", "Canceled"),
        ("outgoing", 0.33),
        ("incoming", 1),
        ("status", "Canceled"),
        ("status", "Canceled"),
        ("missed_by", "incoming"),
        ("status", "Canceled")
    ],
    "May 27, 2024": [("status", "Canceled")],
    "May 28, 2024": [("outgoing", 24/60), ("outgoing", 4/60), ("outgoing", 6), ("missed_by", "incoming")],
    "May 29, 2024": [("status", "Canceled")],
    "May 30, 2024": [("status", "Canceled"), ("outgoing", 8)],
    "May 31, 2024": [
        ("status", "Canceled"),
        ("status", "Canceled"),
        ("status", "Canceled"),
        ("incoming", 2)    
    ],
    "June 1, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "June 2, 2024": [("status", "Canceled"), ("outgoing", 4), ("outgoing", 15)],
    "June 3, 2024": [("status", "Canceled"), ("outgoing", 5)],
    "June 4, 2024": [("status", "Canceled"), ("status", "Canceled"), ("incoming", 12)],
    "June 5, 2024": [("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("incoming", 17)],
    "June 6, 2024": [("outgoing", 35)],
    "June 7, 2024": [("status", "Canceled")],
    "June 8, 2024": [("status", "Canceled")],
    "June 9, 2024": [("outgoing", 10)],
    "June 10, 2024": [("incoming", 6)],
    "June 12, 2024": [("status", "Canceled")],
    "June 13, 2024": [("status", "Canceled")],
    "June 14, 2024": [("status", "Canceled"), ("incoming", 18)],
    "June 15, 2024": [("incoming", 13)],
    "June 16, 2024": [("status", "Canceled"), ("incoming", 12)],
    "June 17, 2024": [("status", "Canceled")],
    "June 19, 2024": [("status", "Canceled")],
    "June 20, 2024": [("status", "Canceled"), ("incoming", 10)],
    "June 21, 2024": [("status", "Canceled")],
    "June 22, 2024": [("status", "Canceled"), ("incoming", 12)],
    "June 23, 2024": [("status", "Canceled")],
    "June 24, 2024": [("status", "Canceled")],
    "June 25, 2024": [("status", "Canceled"), ("incoming", 26), ("missed_by", "incoming")],
    "June 26, 2024": [("status", "Canceled")],
    "June 27, 2024": [("status", "Canceled")],
    "June 28, 2024": [("status", "Canceled"), ("incoming", 25)],
    "June 29, 2024": [("status", "Canceled"), ("incoming", 38)],
    "June 30, 2024": [("status", "Canceled")],
    "June 1, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "July 2, 2024": [("status", "Canceled"), ("incoming", 19)],
    "July 3, 2024": [("status", "Canceled"), ("outgoing", 22/60), ("outgoing", 55/60), ("incoming", 10)],
    "July 4, 2024": [("status", "Canceled"), ("missed_by", "incoming"), ("status", ("Canceled")), ("missed_by", "incoming"), ("status", "Canceled"), ("status", "Canceled")],
    "July 5, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "July 6, 2024": [("status", "Canceled"), ("outgoing", 10), ("missed_by", "incoming"), ("status", "Canceled"), ("status", "Canceled")],
    "July 7, 2024": [("status", "Canceled")],
    "July 8, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "July 9, 2024": [("status", "Canceled")],
    "July 10, 2024": [("status", "Canceled"), ("incoming", 4)],
    "July 11, 2024": [("status", "Canceled")],
    "July 12, 2024": [("outgoing", 54/60)],
    "July 13, 2024": [("status", "Canceled")],
    "July 14, 2024": [("status", "Canceled")],
    "July 15, 2024": [("status", "Canceled")],
    "July 16, 2024": [("status", "Canceled"), ("incoming", 15)],
    "July 17, 2024": [("status", "Canceled")],
    "July 18, 2024": [("outgoing", 9)],
    "July 19, 2024": [("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("incoming", 3), ("missed_by", "incoming"), ("status", "Canceled")],
    "July 21, 2024": [("status", "Canceled")],
    "July 22, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "July 23, 2024": [("status", "Canceled")],
    "July 24, 2024": [("status", "Canceled"), ("incoming", 40)],
    "July 25, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "July 26, 2024": [("outgoing", 18/60)],
    "July 27, 2024": [("status", "Canceled"), ("incoming", 22), ("missed_by", "incoming"), ("status", "Canceled")],
    "July 28, 2024": [("status", "Canceled"), ("status", "Canceled"), ("outgoing", 19)],
    "July 29, 2024": [("incoming", 21)],
    "July 30, 2024": [("status", "Canceled"), ("incoming", 2)],
    "August 6, 2024": [("incoming", 7)],
    "August 10, 2024": [("status", "Canceled")],
    "August 11, 2024": [("status", "Canceled")],
    "August 12, 2024": [("status", "Canceled"), ("missed_by", "incoming")],
    "August 13, 2024": [("outgoing", 24), ("outgoing", 1), ("missed_by", "incoming")],
    "August 14, 2024": [("status", "Canceled")],
    "August 15, 2024": [("status", "Canceled")],
    "August 16, 2024": [("outgoing", 21)],
    "August 17, 2024": [("status", "Canceled")],
    "August 18, 2024": [("status", "Canceled")],
    "August 19, 2024": [("otugoing", 4)],
    "August 20, 2024": [("status", "Canceled")],
    "August 21, 2024": [("outgoing", 7)],
    "August 22, 2024": [("status", "Canceled")],
    "August 23, 2024": [("status", "Canceled")],
    "August 24, 2024": [("outgoing", 5), ("outgoing", 7), ("incoming", 29/60), ("status", "Canceled"), ("status", "Canceled"), ("missed_by", "incoming"), ("status", "Canceled")],
    "August 25, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "August 26, 2024": [("status", "Canceled")],
    "August 27, 2024": [("incoming", 25)],
    "August 28, 2024": [("status", "Canceled")],
    "August 29, 2024": [("status", "Canceled")],
    "August 30, 2024": [("status", "Canceled")],
    "August 31, 2024": [("status", "Canceled"), ("missed_by", "incoming")],
    "September 1, 2024": [("status", "Canceled")],
    "September 2, 2024": [("status", "Canceled"), ("incoming", 14), ("incoming", 6)],
    "September 3, 2024": [("status", "Canceled")],
    "September 4, 2024": [("status", "Canceled")],
    "September 5, 2024": [("status", "Canceled"), ("incoming", 11)],
    "September 6, 2024": [("status", "Canceled")],
    "September 7, 2024": [("status", "Canceled")],
    "September 8, 2024": [("outgoing", 31)],
    "September 9, 2024": [("status", "Canceled"), ("incoming", 9), ("outgoing", 16)],
    "September 10, 2024": [("status", "Canceled")],
    "September 11, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "September 12, 2024": [("status", "Canceled"), ("incoming", 15)],
    "September 13, 2024": [("missed_by", "incoming"), ("status", "Canceled")],
    "September 14, 2024": [("status", "Canceled")],
    "September 15, 2024": [("status", "Canceled")],
    "September 16, 2024": [("status", "Canceled"), ("incoming", 17)],
    "September 17, 2024": [("status", "Canceled")],
    "September 18, 2024": [("status", "Canceled")],
    "September 19, 2024": [("status", "Canceled")],
    "September 20, 2024": [("status", "Canceled")],
    "September 21, 2024": [("status", "Canceled"), ("incoming", 3)],
    "September 22, 2024": [("status", "Canceled")],
    "September 23, 2024": [("outgoing", 24)],
    "September 24, 2024": [("status", "Canceled")],
    "September 25, 2024": [("status", "Canceled"), ("incoming", 12)],
    "September 26, 2024": [("outgoing", 20)],
    "September 28, 2024": [("status", "Canceled")],
    "September 29, 2024": [("status", "Canceled")],
    "September 30, 2024": [("outgoing", 7)],
    "October 1, 2024": [("status", "Canceled")],
    "October 2, 2024": [("outgoing", 16), ("incoming", 28)],
    "October 3, 2024": [("status", "Canceled")],
    "October 4, 2024": [("status", "Canceled")],
    "October 5, 2024": [("status", "Canceled")],
    "October 6, 2024": [("status", "Canceled")],
    "October 7, 2024": [("status", "Canceled")],
    "October 9, 2024": [("status", "Canceled"), ("incoming", 15)],
    "October 10, 2024": [("status", "Canceled")],
    "October 11, 2024": [("outgoing", 4)],
    "October 12, 2024": [("outgoing", 12), ("outgoing", 9)],
    "October 13, 2024": [("status", "Canceled")],
    "October 14, 2024": [("status", "Canceled")],
    "October 15, 2024": [("incoming", 1), ("outgoing", 12)],
    "October 16, 2024": [("status", "Canceled")],
    "October 17, 2024": [("status", "Canceled")],
    "October 18, 2024": [("status", "Canceled")],
    "October 19, 2024": [("status", "Canceled"), ("outgoing", 7), ("status", "Canceled")],
    "October 20, 2024": [("status", "Canceled")],
    "October 22, 2024": [("status", "Canceled")],
    "October 23, 2024": [("status", "Canceled")],
    "October 24, 2024": [("outgoing", 5)],
    "October 25, 2024": [("status", "Canceled")],
    "October 26, 2024": [("status", "Canceled")],
    "October 27, 2024": [("status", "Canceled")],
    "October 28, 2024": [("status", "Canceled")],
    "October 29, 2024": [("status", "Canceled")],
    "October 30, 2024": [("status", "Canceled"), ("outgoing", 9)],
    "October 31, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "November 1, 2024": [("status", "Canceled")],
    "November 2, 2024": [("status", "Canceled")],
    "November 3, 2024": [("status", "Canceled"), ("incoming", 3)],
    "November 4, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "November 5, 2024": [("status", "Canceled")],
    "November 6, 2024": [("status", "Canceled")],
    "November 7, 2024": [("incoming", 14)],
    "November 8, 2024": [("status", "Canceled"), ("missed_by", "incoming"), ("status", "Canceled")],
    "November 9, 2024": [("status", "Canceled"), ],
    "November 10, 2024": [
        ("status", "Canceled"),
        ("outgoing", 3),
        ("missed_by", "incoming"),
        ("incoming", 5),
        ("status", "Canceled"),
    ],
    "November 11, 2024": [("status", "Canceled")],
    "November 12, 2024": [
        ("status", "Canceled"),
        ("incoming", 0.66),
        ("outgoing", 20),
    ],
    "November 13, 2024": [("status", "Canceled")],
    "November 14, 2024": [("outgoing", 6)],
    "November 15, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "November 16, 2024": [("status", "Canceled")],
    "November 17, 2024": [("outgoing", 12)],
    "November 18, 2024": [("status", "Canceled")],
    "November 20, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "November 22, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "November 23, 2024": [("status", "Canceled")],
    "November 24, 2024": [("status", "Canceled")],
    "November 26, 2024": [("outgoing", 4)],
    "November 27, 2024": [("status", "Canceled"), ("outgoing", 7)],
    "November 28, 2024": [("status", "Canceled")],
    "November 29, 2024": [("status", "Canceled"), ("status", "Canceled")],
    "November 30, 2024": [("status", "Canceled")],
    "December 1, 2024": [("outgoing", 29)],
    "December 2, 2024": [("status", "Canceled")],
    "December 3, 2024": [("status", "Canceled")],
    "December 5, 2024": [("status", "Canceled")],
    "December 6, 2024": [("status", "Canceled")],
    "December 7, 2024": [("status", "Canceled")],
    "December 8, 2024": [("status", "Canceled")],
    "December 9, 2024": [("outgoing", 8)],
    "December 10, 2024": [("status", "Canceled")],
    "December 11, 2024": [("status", "Canceled")],
    "December 12, 2024": [("status", "Canceled")],
    "December 13, 2024": [("status", "Canceled")],
    "December 14, 2024": [("status", "Canceled")],
    "December 15, 2024": [("status", "Canceled")],
    "December 16, 2024": [("status", "Canceled")],
    "December 17, 2024": [("status", "Canceled")],
    "December 18, 2024": [("status", "Canceled")],
    "December 19, 2024": [("status", "Canceled"), ("incoming", 13)],
    "December 20, 2024": [("status", "Canceled")],
    "December 21, 2024": [("status", "Canceled")],
    "December 22, 2024": [("status", "Canceled")],
    "December 23, 2024": [("status", "Canceled")],
    "December 24, 2024": [("status", "Canceled")],
    "December 25, 2024": [("status", "Canceled"), ("status", "Canceled"), ("status", "Canceled"), ("incoming", 1)],
    "December 26, 2024": [("status", "Canceled")],
    "December 27, 2024": [("outgoing", 29)],
    "December 28, 2024": [("missed_by", "incoming"), ("status", "Canceled")],
    "December 29, 2024": [("outgoing", 8)],
    "December 30, 2024": [("status", "Canceled")],
    "December 31, 2024": [("status", "Canceled")],
    "January 1, 2024": [("status", "Canceled")],
    "January 3, 2024": [("outgoing", 45)],
    "January 4, 2024": [("status", "Canceled")],
    "January 5, 2024": [("missed_by", "incoming")],

}

def convert_to_mm_ss(minutes):
    # Convert minutes to total seconds
    total_seconds = round(minutes * 60)
    
    # Calculate minutes and seconds
    minutes = total_seconds // 60
    seconds = total_seconds % 60
    
    # Return in mm:ss format
    return f"{minutes:02}:{seconds:02}"

# Consolidated data analysis function to avoid redundant calculations
def analyze_calls(call_data):
    total_calls = 0
    canceled_calls = 0
    missed_calls = 0
    outgoing_calls = 0
    incoming_calls = 0
    outgoing_duration = 0
    incoming_duration = 0
    outgoing_missed = 0
    incoming_missed = 0

    # Iterate over call data to calculate totals
    for date, calls in call_data.items():
        for call_type, value in calls:
            total_calls += 1
            if call_type == "status":
                if value == "Canceled":
                    canceled_calls += 1
            elif call_type == "outgoing":
                outgoing_calls += 1
                outgoing_duration += value
            elif call_type == "incoming":
                incoming_calls += 1
                incoming_duration += value
            elif call_type == "missed_by":
                if value == "outgoing":
                    missed_calls += 1
                    outgoing_missed += 1
                elif value == "incoming":
                    missed_calls += 1
                    incoming_missed += 1



    # Return summary of analysis
    return {
        "Total Calls": total_calls,
        "Canceled Calls": canceled_calls,
        "Total Missed Calls": missed_calls,
        "Missed Incoming": incoming_missed,
        "Missed Outgoing": outgoing_missed,
        "Outgoing Calls": outgoing_calls,
        "Incoming Calls": incoming_calls,
        "Total Outgoing Duration (minutes)": convert_to_mm_ss(outgoing_duration),
        "Total Incoming Duration (minutes)": convert_to_mm_ss(incoming_duration),
    }

# Call breakdown by date
def call_summary_by_date(call_data):
    summary_by_date = {}

    # Iterate through each date's calls to generate a daily summary
    for date, calls in call_data.items():
        daily_summary = {
            "Total Calls": 0,
            "Canceled Calls": 0,
            "Total Missed Calls": 0,
            "Incoming Missed": 0,
            "Outgoing Missed": 0,
            "Outgoing Calls": 0,
            "Incoming Calls": 0,
            "Outgoing Duration (minutes)": 0,
            "Incoming Duration (minutes)": 0
        }

        for call_type, value in calls:
            daily_summary["Total Calls"] += 1
            if call_type == "status":
                if value == "Canceled":
                    daily_summary["Canceled Calls"] += 1
            elif call_type == "outgoing":
                daily_summary["Outgoing Calls"] += 1
                daily_summary["Outgoing Duration (minutes)"] += value
            elif call_type == "incoming":
                daily_summary["Incoming Calls"] += 1
                daily_summary["Incoming Duration (minutes)"] += value
            elif call_type == "missed_by":
                if value == "incoming":
                    daily_summary["Incoming Missed"] += 1
                elif value == "outgoing":
                    daily_summary["Outgoing Missed"] += 1
                daily_summary["Total Missed Calls"] += 1

        summary_by_date[date] = daily_summary

    # Sort the dates in the correct order using the full month name format
    sorted_summary_by_date = {date: summary_by_date[date] for date in sorted(summary_by_date, key=lambda x: datetime.strptime(x, "%B %d, %Y"))}

    return sorted_summary_by_date

# Duration analysis (streamlined)
def duration_analysis(call_data):
    total_outgoing_duration = 0
    total_incoming_duration = 0
    total_outgoing_calls = 0
    total_incoming_calls = 0

    for date, calls in call_data.items():
        for call_type, value in calls:
            if call_type == "outgoing":
                total_outgoing_calls += 1
                total_outgoing_duration += value
            elif call_type == "incoming":
                total_incoming_calls += 1
                total_incoming_duration += value

    # Calculate averages
    average_outgoing_duration = round(total_outgoing_duration / total_outgoing_calls, 2) if total_outgoing_calls else 0
    average_incoming_duration = round(total_incoming_duration / total_incoming_calls, 2) if total_incoming_calls else 0

    # Convert the durations to mm:ss format
    total_outgoing_duration = convert_to_mm_ss(total_outgoing_duration)
    average_outgoing_duration = convert_to_mm_ss(average_outgoing_duration)
    total_incoming_duration = convert_to_mm_ss(total_incoming_duration)
    average_incoming_duration = convert_to_mm_ss(average_incoming_duration)

    return {
        "Total Outgoing Duration (mm:ss)": total_outgoing_duration,
        "Average Outgoing Duration (mm:ss)": average_outgoing_duration,
        "Total Incoming Duration (mm:ss)": total_incoming_duration,
        "Average Incoming Duration (mm:ss)": average_incoming_duration
    }

# Call frequency analysis
def call_frequency_analysis(call_data):
    total_days = len(call_data)
    total_calls = sum([len(calls) for calls in call_data.values()])

    average_calls_per_day = total_calls / total_days if total_days else 0

    return {
        "Total Days Analyzed (answered or not)": total_days,
        "Total Calls (answered or not)": total_calls,
        "Average Calls per Day (answered or not)": round(average_calls_per_day, 2)
    }

# Call type breakdown
def call_type_breakdown(call_data):
    breakdown = {
        "Canceled": 0,
        "Total Missed": 0,
        "Outgoing": 0,
        "Incoming": 0,
        "Missed Outgoing": 0,
        "Missed Incoming": 0
    }

    for date, calls in call_data.items():
        for call_type, value in calls:
            if call_type == "status":
                if value == "Canceled":
                    breakdown["Canceled"] += 1
            elif call_type == "outgoing":
                breakdown["Outgoing"] += 1
            elif call_type == "incoming":
                breakdown["Incoming"] += 1
            elif call_type == "missed_by":
                if value == "outgoing":
                    breakdown["Missed Outgoing"] += 1
                elif value == "incoming":
                    breakdown["Missed Incoming"] += 1
                breakdown["Total Missed"] += 1

    return breakdown

# Total call duration by status
def total_call_duration_by_status(call_data):
    durations = {
        "Canceled": 0,
        "Total Missed": 0,
        "Outgoing": 0,
        "Incoming": 0,
        "Missed Incoming": 0,
        "Missed Outgoing": 0
    }

    for date, calls in call_data.items():
        for call_type, value in calls:
            if call_type == "status":
                if value == "Canceled":
                    durations["Canceled"] += 1  # We count canceled calls, but no duration added
            elif call_type == "outgoing":
                durations["Outgoing"] += value
            elif call_type == "incoming":
                durations["Incoming"] += value
            if call_type == "missed_by":
                if value == "outgoing":
                    durations["Missed Outgoing"] += 1
                elif value == "incoming":
                    durations["Missed Incoming"] += 1
                durations["Total Missed"] += 1

    durations = {key: convert_to_mm_ss(value) for key, value in durations.items()}
    return durations

# Missed and Canceled call percentage
def missed_and_canceled_call_percentage(call_data):
    # Initialize counters for total and missed+canceled calls
    total_calls = 0
    missed_and_canceled_calls = 0
    
    # Loop through the call data to calculate totals
    for date, calls in call_data.items():
        for call_type, value in calls:
            total_calls += 1
            if call_type == "status" and value == "Canceled":
                missed_and_canceled_calls += 1
            if call_type == "missed_by" and value == "outgoing":
                missed_and_canceled_calls += 1

    # Calculate missed and canceled call percentage if total_calls is greater than 0
    missed_and_canceled_percentage = (missed_and_canceled_calls / total_calls * 100) if total_calls > 0 else 0

    return f"Missed & Canceled Call Percentage (calls that she missed): {round(missed_and_canceled_percentage, 2)}%"  # Rounded to 2 decimal places

print("analyze_calls")
analysis1 = analyze_calls(call_data)
for key, value in analysis1.items():
    print(f"{key}: {value}")

print("\n")

print("call_summary_by_date")
daily_summary = call_summary_by_date(call_data)
for date, summary in daily_summary.items():
    print(f"Date: {date}")
    for key, value in summary.items():
        print(f"  {key}: {value}")
    print("\n")

print("duration_analysis")
duration_results = duration_analysis(call_data)
for key, value in duration_results.items():
    print(f"{key}: {value}")
print("\n")

print("total_call_duration_by_status")
status_duration_results = total_call_duration_by_status(call_data)
for key, value in status_duration_results.items():
    print(f"{key} Duration: {value} minutes")

print("missed_and_canceled_call_percentage")
missed_and_canceled_call_results = missed_and_canceled_call_percentage(call_data)
print(missed_and_canceled_call_results)

def calculate_canceled_or_missed_days(call_data):
    total_days = len(call_data)
    matching_days = 0

    for date, records in call_data.items():
        # Check if all records are either ("status", "Canceled") or ("missed_by", "outgoing")
        if all(record in [("status", "Canceled"), ("missed_by", "outgoing")] for record in records):
            matching_days += 1

    # Calculate the percentage
    percentage = (matching_days / total_days) * 100 if total_days > 0 else 0

    return matching_days, percentage

matching_days, percentage = calculate_canceled_or_missed_days(call_data)
print("\n")
print(f"Days where she did not call: {matching_days}")
print(f"Percentage: {percentage:.2f}%")
print("\n")

def create_call_frequency_spreadsheet(call_data, filename="call_frequency_summary.xlsx"):
    # Create a workbook and a sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Call Frequency Summary"
    
    # Add headers
    ws.append(["Date", "Incoming Call Frequency", "Canceled Call Frequency", "Missed Call Frequency", "Incoming Missed Call Frequency", "Outgoing Missed Call Frequency"])
    
    # Iterate over call data to populate the spreadsheet
    for date, calls in call_data.items():
        incoming_count = 0
        canceled_count = 0
        missed_count = 0
        incoming_missed_count = 0
        outgoing_missed_count = 0
        
        # Count the occurrences of each call type
        for call_type, value in calls:
            if call_type == "status":
                if value == "Canceled":
                    canceled_count += 1
            elif call_type == "incoming":
                incoming_count += 1
            if call_type == "missed_by":
                if value == "incoming":
                    incoming_missed_count += 1
                elif value == "outgoing":
                    outgoing_missed_count += 1
                missed_count += 1
        
        # Append the data to the sheet
        ws.append([date, incoming_count, canceled_count, missed_count, incoming_missed_count, outgoing_missed_count])
    
    # Save the workbook to a file
    wb.save(filename)
    print(f"Spreadsheet saved as {filename}")

# Add the function call here to generate the spreadsheet
create_call_frequency_spreadsheet(call_data)