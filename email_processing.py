import time, warnings
warnings.filterwarnings("ignore")
from config import *
from src.logger import *
from src.mailbox import mailbox
from concurrent.futures import ThreadPoolExecutor
from src.print_functions import header, timed_row
from src.EmailClassifier import import_classifiers_from

def _process_message(message, EC_list):
	try:
		recognized = False
		for email_classifier in EC_list:
			if email_classifier.recognizes(message):
				recognized = True
				if email_classifier.action(message):
					message.move(mailbox.get_folder(email_classifier.move_email_to))
				else:
					message.move(mailbox.get_folder(error_folder))
				break
		if not recognized:
			message.move(mailbox.get_folder(unrecognized_folder))
	except Exception as e:
		logging.error(f'_process_message FAILED!\n- message: {message.subject}\n- exception: {e}')
	timed_row(f'Processed: {message.subject}')


def run_email_processing():
	while True:
		EC_list = import_classifiers_from('email_classifiers')
		with ThreadPoolExecutor(max_workers=4) as executor:
			for message in mailbox.messages():
				executor.submit(_process_message, message, EC_list)
		time.sleep(5)


if __name__ == '__main__':
	header('Carpincho Email Processing')
	logging.warning(f'Started.')
	run_email_processing()
