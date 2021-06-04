from datetime import timedelta, date
a = date.today()
b = timedelta(days=3)
print(date.weekday(date.today()-timedelta(days=1)))