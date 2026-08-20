from datetime import datetime


def stamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")
