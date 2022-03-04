import datetime

def hours_between(t1:str, t2:str):
    def to_dt(t):
        return datetime.datetime.strptime(t, '%H:%M')
    times = sorted([to_dt(t1), to_dt(t2)])
    delta = times[1] - times[0]
    delta.seconds / 60 / 60