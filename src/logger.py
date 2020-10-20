import logging

log_format = "[%(asctime)s][%(levelname)s][%(name)s][%(filename)s][%(lineno)d] %(message)s"

logging.basicConfig(filename='events.log', filemode='a', level='WARNING', format=log_format)