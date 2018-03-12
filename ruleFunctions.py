import re

class DateError(Exception):
    def __init__(self, date, msg=None):
        if msg is None:
            # Set some default useful error message
            msg = "No dates found in string: %s" % date
        super(DateError, self).__init__(msg)
        self.date = date

class RuleParseError(Exception):
    pass

def insideDates(dates):
    #
    dates = splitDates(dates)
    return 'TODO: count inside dates'
    pass

def ddMmYyyy(dates):
    ''' >>> date('July 6, 2016')
    06/06/2016
    '''
    pass

def monthDayYear(ddMmYyyy):
    months = {1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
    month = months(ddMmYyyy[1])
    day = ddMmYyyy[0]
    num_suff = {1:'st',2:'nd',3:'rd'}
    if day > 3:
        suff = 'th'
    else: suff = num_suff[day]
    return month + str(day) + suff + str(ddMmYyyy[2])

def startDate(dates):
    if splitDates(dates):
        return splitDates(dates)[0]
    else: raise DateError(dates)

def endDate(dates):
    if splitDates(dates):
        return splitDates(dates)[-1]
    else: raise DateError(dates)

def splitDates(dates):
    ''' Helper. Returns tuple (or list of tuples) containing ddMmYyyy start, end dates.
    >>> splitDates('July 6,2018 to July 28, 2018')
    [(6,6,2018),(28,06,2018)]
    >>> str('123')
    567
    '''
    splitted = re.findall('\w+ *\d{1,2} +\d{2,4}|\d{2}/\d{2}/\d{2,4}|\w+ *\d{1,2}|\d{1,2} *\w+', dates)
    print(splitted)

    return splitted

def budgetFromSalary(salary):
    try:
        salary = int(salary)
    except ValueError:
        raise RuleParseError('salary %s cannot be converted to integer' %salary)
        budget = 'Unknown'
    # TODO: estimate production budget according to salary
    payrates = {'actor':0.1}

    budget = str(salary)
    return budget
